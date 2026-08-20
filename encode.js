/**
 * COURT_DATA_SHEETから裁判所マスタデータ（2行目以降）を取得する
 * @returns {Array<Array<*>>} 裁判所マスタの行データ配列
 */
const getCourtData = () => {
  const lastRow = COURT_DATA_SHEET.getLastRow();
  if (lastRow < 2) return [];
  return COURT_DATA_SHEET.getRange(2, 1, lastRow - 1, 7).getValues();
};

/**
 * 裁判所種別を審級コード（1桁）に変換する
 * @param {string} level - 裁判所種別（例: "地裁", "高裁"）
 * @returns {string} 審級コード。該当なしは空文字
 */
const encodeLevel = (level) => {
  if (level === "最高裁" || level === "大審院") return "0";
  if (level === "高裁" || level === "高等法院" || level === "控訴院")
    return "1";
  if (level === "地裁") return "2";
  if (level === "家裁") return "3";
  if (level === "地裁支部") return "4";
  if (level === "家裁支部") return "5";
  if (level === "地裁出張所") return "6";
  if (level === "家裁出張所") return "7";
  if (level === "簡裁" || level === "区裁") return "8";
  if (level === "その他") return "9";
  if (level === "実例先例例規") return "10";
  if (level === "外国裁判所") return "11";
  if (level === "複数裁判例") return "12";
  return "";
};

/**
 * 改廃区分を1桁コードに変換する
 * @param {string} mod - 改廃区分（例: "統廃合", "名称変更"）
 * @returns {string} 改廃コード。該当なしは"0"
 */
const encodeMod = (mod) => {
  if (mod === "統廃合") return "1";
  if (mod === "名称変更") return "2";
  if (mod === "名称変更後") return "3";
  return "0";
};

/**
 * 裁判所種別から種別コード（家裁系かどうか）を判定する
 * @param {string} level - 裁判所種別
 * @returns {string} 家裁系なら"1"、それ以外は"0"
 */
const getGenreCode = (level) =>
  level === "家裁" || level === "家裁支部" || level === "家裁出張所"
    ? "1"
    : "0";

/**
 * 裁判所マスタデータから裁判所コードを検索する
 * マスタデータはencodePrecs実行中に変化しないため、呼び出し元で1回だけ取得しキャッシュとして渡す想定
 * @param {string} courtName - 裁判所名
 * @param {Array<Array<*>>} courtData - getCourtDataの返り値
 * @returns {string} 審級別裁判所コード。計算できない場合は空文字
 */
const encodeCourtName = (courtName, courtData) => {
  if (courtData.length < 1) return "";

  const trimmedName = courtName.trim();
  const courtRow = courtData.find(
    (row) => String(row[0]).trim() === trimmedName,
  );
  if (!courtRow) return "";

  const [, , , level, code, legacyFlag, mod] = courtRow;
  const levelCode = encodeLevel(level);
  if (levelCode === "") return "";

  return `${levelCode}${legacyFlag ? 0 : 1}${code}${encodeMod(mod)}${getGenreCode(level)}`;
};

/**
 * 和暦の元号コード定義
 * 値配列の1つ目は8桁コード用の2桁数字、2つ目は事件番号コード用の英字1文字
 */
const ERA_CODE_MAP = {
  明治: ["01", "m"],
  大正: ["02", "t"],
  昭和: ["03", "s"],
  平成: ["04", "h"],
  令和: ["05", "r"],
};

/**
 * 数値を2桁ゼロ埋め文字列に変換する
 * @param {number|string} value - 変換対象の値
 * @returns {string} 2桁ゼロ埋め文字列
 */
const padZero2 = (value) => String(value).padStart(2, "0");

/**
 * 和暦の日付文字列を "元号コード + 年2桁 + 月2桁 + 日2桁" の8桁数字文字列に正規化する
 * 例: "平成16年10月13日" -> "04161013"
 * @param {string} labeledYear - 和暦の日付文字列（例: "平成16年10月13日"）
 * @returns {string} 正規化された8桁の数字文字列。変換できない場合は空文字
 */
const encodeYear = (labeledYear) => {
  const matched = labeledYear.match(
    /^(明治|大正|昭和|平成|令和)(\d+)年(\d+)月(\d+)日$/,
  );
  if (!matched) return "";

  const [, era, year, month, day] = matched;
  const [eraCode] = ERA_CODE_MAP[era];
  return `${eraCode}${padZero2(year)}${padZero2(month)}${padZero2(day)}`;
};

/**
 * SIGN_DATA_SHEETから事件符号マスタデータ（2行目以降）を取得する
 * @returns {Array<Array<*>>} 事件符号マスタの行データ配列
 */
const getSignData = () => {
  const lastRow = SIGN_DATA_SHEET.getLastRow();
  if (lastRow < 2) return [];
  return SIGN_DATA_SHEET.getRange(2, 1, lastRow - 1, 2).getValues();
};

/**
 * 事件番号（例: "平成16(ネ)3324"）を事件番号コードに変換する
 * 半角括弧 "()" で符号が囲まれている形式を前提とする
 * 事件符号マスタデータはencodePrecs実行中に変化しないため、呼び出し元で1回だけ取得しキャッシュとして渡す想定
 * @param {string} caseNumber - 事件番号（例: "平成16(ネ)3324"）
 * @param {Array<Array<*>>} signData - getSignDataの返り値
 * @returns {string} 事件番号コード。計算できない場合は空文字
 */
const encodeCaseNumber = (caseNumber, signData) => {
  const [labeledYear, sign, num] = caseNumber.replace(/[()]/g, "_").split("_");

  const label = labeledYear.substring(0, 2);
  const eraCode = ERA_CODE_MAP[label];
  if (!eraCode) return "";
  const [, eraLetter] = eraCode;

  const year = labeledYear.substring(2);
  const trimmedSign = sign.trim();
  const signRow = signData.find((row) => String(row[0]).trim() === trimmedSign);
  if (!signRow) return "";

  const [, signCode] = signRow;
  return `${eraLetter}${year}${signCode}${num}`;
};

/**
 * MAIN_SHEETのB・D・E列（事件番号・裁判年月日・裁判所名）が埋まっていて、
 * かつF列（判例コード）が未計算の行を対象に、判例コードを計算してF列に書き込む
 *
 * 判例コードの計算に事件名（C列）は使用しないため、対象行の詳細読み込みではC列を含めていない
 * 対象行の特定にはB列（事件番号）とF列（判例コード）のみを読み込み、
 * 書き込みも対象行のみに絞ることで、全行読み込み・全行書き込みによるコストを避けている
 * （数万行規模でも、対象は日々増分される未処理分のみになる想定）
 * @returns {void}
 */
const encodePrecs = () => {
  const lastRow = MAIN_SHEET.getLastRow();
  if (lastRow < 1) return;

  const FILLED_COLUMNS_COUNT = 5;
  const PREC_CODE_COLUMN = FILLED_COLUMNS_COUNT + 1;
  const JUDGE_DATE_COLUMN = 4; // D列
  const COURT_NAME_COLUMN = 5; // E列

  // 対象行の絞り込みにはB列とF列のみを読む（軽量化）
  const caseNumberValues = MAIN_SHEET.getRange(
    1,
    CASE_NUMBER_COLUMN,
    lastRow,
    1,
  ).getValues();
  const precCodeValues = MAIN_SHEET.getRange(
    1,
    PREC_CODE_COLUMN,
    lastRow,
    1,
  ).getValues();

  const targetRowIndices = [];
  for (let i = 0; i < lastRow; i++) {
    const caseNumber = caseNumberValues[i][0];
    const precCode = precCodeValues[i][0];
    if (!caseNumber) continue; // B列未取得（スクレイピング未処理）
    if (precCode !== "") continue; // すでにコード計算済み
    targetRowIndices.push(i);
  }

  if (targetRowIndices.length < 1) return;

  // マスタデータは対象行数に関わらず1回だけ読み込み、以降はキャッシュとして使い回す
  const courtData = getCourtData();
  const signData = getSignData();

  // 対象行のみ、裁判年月日・裁判所名（D・E列）を読み込む。事件名（C列）は計算に不要なため対象外
  const minRow = Math.min(...targetRowIndices) + 1;
  const maxRow = Math.max(...targetRowIndices) + 1;
  const detailValues = MAIN_SHEET.getRange(
    minRow,
    JUDGE_DATE_COLUMN,
    maxRow - minRow + 1,
    2,
  ).getValues();

  const results = []; // { row, code }
  for (const i of targetRowIndices) {
    const caseNumber = caseNumberValues[i][0];
    const [labeledYear, courtName] = detailValues[i - (minRow - 1)];

    const courtCode = encodeCourtName(courtName, courtData);
    if (courtCode === "") continue;

    const yearCode = encodeYear(labeledYear);
    if (yearCode === "") continue;

    const caseNumCode = encodeCaseNumber(caseNumber, signData);
    if (caseNumCode === "") continue;

    results.push({
      row: i + 1,
      code: `${courtCode}-${yearCode}-${caseNumCode}`,
    });
  }

  if (results.length < 1) return;

  // 対象行1件ごとにsetValueすると数万件でAPI呼び出し過多になるため、
  // 連続範囲ごとにまとめてsetValuesする
  let batchStart = 0;
  for (let i = 1; i <= results.length; i++) {
    const isEndOfBatch =
      i === results.length || results[i].row !== results[i - 1].row + 1;
    if (!isEndOfBatch) continue;

    const batch = results.slice(batchStart, i);
    const startRow = batch[0].row;
    const values = batch.map(({ code }) => [code]);
    MAIN_SHEET.getRange(startRow, PREC_CODE_COLUMN, values.length, 1).setValues(
      values,
    );

    batchStart = i;
  }
};
