const getCourtData = () => {
  const lastRow = COURT_DATA_SHEET.getLastRow();
  if (lastRow < 1) return [];
  return COURT_DATA_SHEET.getRange(2, 1, lastRow, 7).getValues();
};

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

const encodeMod = (mod) => {
  if (mod === "統廃合") return "1";
  if (mod === "名称変更") return "2";
  if (mod === "名称変更後") return "3";
  return "0";
};

const getGenreCode = (level) => {
  return level === "家裁" || level === "家裁支部" || level === "家裁出張所"
    ? "1"
    : "0";
};

/**
 * 裁判所名から「審級別裁判所コード」を計算する。
 * 計算できなかった場合は空文字を返す。
 */
const encodeCourtName = (courtName) => {
  const data = getCourtData();
  if (data.length < 1) return "";
  const [courtRow] = data.filter((row) => row[0].trim() === courtName.trim());
  if (!courtRow) return "";
  const [, , , level, code, legacyFlag, mod] = courtRow;
  const levelCode = encodeLevel(level);
  if (levelCode === "") return "";
  return `${levelCode}${legacyFlag ? 0 : 1}${code}${encodeMod(mod)}${getGenreCode(level)}`;
};

/**
 * 和暦の元号コード定義
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
 * @returns {string} 正規化された8桁の数字文字列
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

const getSignData = () => {
  const lastRow = SIGN_DATA_SHEET.getLastRow();
  if (lastRow < 1) return [];
  return SIGN_DATA_SHEET.getRange(2, 1, lastRow, 2).getValues();
};

const encodeCaseNumber = (caseNumber) => {
  const [labeledYear, sign, num] = caseNumber
    .replace(/[\(\)]/g, "_")
    .split("_");
  const label = labeledYear.substring(0, 2);
  const eraCode = ERA_CODE_MAP[label];
  if (!eraCode) return "";
  const [, eraLetter] = eraCode;
  const year = labeledYear.substring(2);
  const [signRow] = getSignData().filter(
    (row) => row[0].trim() === sign.trim(),
  );
  if (!signRow) return "";
  const [, signCode] = signRow;
  return `${eraLetter}${year}${signCode}${num}`;
};

const encodePrecs = () => {
  const lastRow = MAIN_SHEET.getLastRow();
  if (lastRow < 1) return;
  const FILLED_COLMUNS_COUNT = 5;
  const values = MAIN_SHEET.getRange(
    1,
    1,
    lastRow,
    FILLED_COLMUNS_COUNT + 1,
  ).getValues();
  const newCol = values.map((v) => {
    const [, caseNumber, , labeledYear, courtName, precCode] = v;
    if (!caseNumber) return [""];
    if (precCode !== "") return [precCode];
    const courtCode = encodeCourtName(courtName);
    if (courtCode === "") return [""];
    const yearCode = encodeYear(labeledYear);
    if (yearCode === "") return [""];
    const caseNumCode = encodeCaseNumber(caseNumber);
    if (caseNumCode === "") return [""];
    return [`${courtCode}-${yearCode}-${caseNumCode}`];
  });
  MAIN_SHEET.getRange(1, FILLED_COLMUNS_COUNT + 1, lastRow, 1).setValues(
    newCol,
  );
};
