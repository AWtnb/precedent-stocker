/**
 * 指定した判例番号のページHTMLを取得する
 * 404などのエラーもレスポンスとしてそのまま返す
 * @param {number|string} caseNumber - 判例番号（URLのXXXXX部分）
 * @returns {GoogleAppsScript.URL_Fetch.HTTPResponse} レスポンス
 */
const fetchResponse = (caseNumber) => {
  const url = `${BASE_URL}${caseNumber}/detail7/index.html`;
  return UrlFetchApp.fetch(url, { muteHttpExceptions: true });
};

/**
 * HTML内の <meta name="xxx" content="yyy"> からnameをキー、contentを値とするマップを作る
 * ページ本体はJavaScriptで動的に組み立てられるため要素の直接取得はできないが、
 * meta要素はサーバー側で埋め込まれるため確実に取得できる
 * @param {string} html - ページ全体のHTML文字列
 * @returns {Object<string, string>} meta要素のnameをキー、contentを値とするマップ
 */
const extractMetaMap = (html) => {
  const metaMap = {};
  const pattern = /<meta\s+name="([^"]+)"\s+content="([^"]*)"/g;

  let match;
  while ((match = pattern.exec(html))) {
    const [, name, content] = match;
    if (name in metaMap) continue; // 同名metaが複数ある場合は最初の値を優先
    metaMap[name] = content;
  }
  return metaMap;
};

/**
 * meta要素のマップから、必要な3項目がすべて揃っているかを確認する
 * @param {Object<string, string>} metaMap - extractMetaMapの返り値
 * @returns {boolean} 「事件番号」「裁判年月日」「裁判所名」の3項目すべて値が入っていればtrue
 */
const isSufficient = (metaMap) =>
  FIELD_COLUMNS.filter(({ metaName }) => metaName !== "jiken_name").every(
    ({ metaName }) => metaMap[metaName],
  );

/**
 * LOG_SHEETのA1から前回処理した行番号（次回の探索開始行）を取得する
 * 空欄なら1行目（先頭）として扱う
 * @returns {number} 探索を開始する行番号（1始まり）
 */
const getLastRowIndex = () => {
  const value = LOG_SHEET.getRange("A1").getValue();
  if (value === "") return 1;
  return Number(value);
};

/**
 * LOG_SHEETのA1に次回の探索開始行を記録する
 * @param {number} rowIndex - 記録する行番号（1始まり）
 * @returns {void}
 */
const setLastRowIndex = (rowIndex) => {
  LOG_SHEET.getRange("A1").setValue(rowIndex);
};

/**
 * B列（事件番号列）を読み込み、startRow以降で最初に見つかった空欄行番号を返す
 * B列が空欄 = A列しか埋まっていない未処理行、という前提に基づく判定
 * B列1列のみを読み込むことで、全列読み込みに比べてデータ転送量を抑えている
 * @param {number} startRow - 検索を開始する行番号（1始まり）
 * @param {number} lastRow - シートの最終行番号
 * @returns {number|null} 見つかった行番号。見つからなければnull
 */
const findFirstUnscrapedRow = (startRow, lastRow) => {
  if (lastRow < startRow) return null;

  const values = MAIN_SHEET.getRange(
    startRow,
    CASE_NUMBER_COLUMN,
    lastRow - startRow + 1,
    1,
  ).getValues();

  for (let i = 0; i < values.length; i++) {
    if (values[i][0] !== "") continue;
    return startRow + i;
  }
  return null;
};

/**
 * MAIN_SHEETの未処理行を1件見つけて判例情報をスクレイピングし、B〜E列に書き込む
 *
 * 処理の流れ:
 * 1. LOG_SHEETのA1から探索開始行を取得する
 * 2. B列を読み込み、空欄の行（未処理行）を探す
 * 3. 見つかった場合はページにアクセスし、meta要素から4項目すべて取得できればB〜E列に書き込む
 *    （404や項目不足の場合は書き込まず、次回以降に再チェックされる）
 * 4. 処理した行番号をLOG_SHEETのA1に記録する（最終行だった場合は1に戻す）
 * 5. 未処理行が見つからず、かつ探索が1行目から始まっていた場合（全行を1周確認した場合）のみ
 *    完了メールを送信する
 * @returns {void}
 */
const scrape = () => {
  const lastRow = MAIN_SHEET.getLastRow();
  if (lastRow < 1) return;

  const startRowIndex = getLastRowIndex();
  const startRow =
    1 <= startRowIndex && startRowIndex <= lastRow ? startRowIndex : 1;

  const targetRow = findFirstUnscrapedRow(startRow, lastRow);
  if (!targetRow) {
    return;
  }

  const caseNumber = MAIN_SHEET.getRange(targetRow, 1).getValue();
  const response = fetchResponse(caseNumber);

  if (response.getResponseCode() === 200) {
    const metaMap = extractMetaMap(response.getContentText());

    if (isSufficient(metaMap)) {
      for (const { metaName, column } of FIELD_COLUMNS) {
        MAIN_SHEET.getRange(targetRow, column).setValue(metaMap[metaName]);
      }
    }
  }
  // 404、meta要素が見つからない、4項目が揃っていない場合は何も書き込まない
  // → 次回以降も未処理行として再チェック対象になる

  const nextRowIndex = targetRow < lastRow ? targetRow + 1 : 1;
  setLastRowIndex(nextRowIndex);
};

/**
 * F列（判例ID）を取得し、すべて埋まっていたらメールで通知する
 * @returns {void}
 */
const checkSheetFilled = () => {
  const lastRow = MAIN_SHEET.getLastRow();
  if (lastRow < 1) return;

  const precCodeColumn = 6;
  const precIDs = MAIN_SHEET.getRange(
    1,
    precCodeColumn,
    lastRow,
    1,
  ).getValues();

  let bottom = -1;
  for (let i = 0; i < precIDs.length; i++) {
    const [precID] = precIDs[i];
    if (precID !== "") {
      bottom = i + 1;
    }
  }

  const mailSub =
    bottom === lastRow
      ? "最終行まで判例IDの発行が完了しました"
      : `${bottom}行まで判例ID発行が済みました`;
  const mailBody =
    bottom === lastRow
      ? "シートを更新してください："
      : "引き続き作業を進めます：";

  MailApp.sendEmail({
    to: MAIL_TO,
    subject: `[GAS]${mailSub}`,
    body: `${mailBody}\n${SHEET.getUrl()}`,
  });
};
