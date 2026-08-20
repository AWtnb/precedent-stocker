const getProperty = (key) => {
  const value = PropertiesService.getScriptProperties().getProperty(key);
  if (!value) throw new Error(`Property not found: ${key}`);
  return value;
};

const SHEET_ID = getProperty("SHEET_ID");
const SHEET = SpreadsheetApp.openById(SHEET_ID);
const MAIN_SHEET = SHEET.getSheets()[0];
const LOG_SHEET = SHEET.getSheets()[1];
const MAIL_TO = getProperty("MAIL_TO");

const BASE_URL = "https://www.courts.go.jp/hanrei/";
const CASE_NUMBER_COLUMN = 2; // B列

// スプレッドシートの列番号（B, C, D, E）とラベルの対応
const FIELD_COLUMNS = [
  { label: "事件番号", column: 2 },
  { label: "事件名", column: 3 },
  { label: "裁判年月日", column: 4 },
  { label: "裁判所名", column: 5 },
];

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
 * HTML内の module-sub-page-parts-table ブロックをすべて抜き出す
 * 同じクラス名のブロックがページ内に複数存在する場合があるため、iterateで全件取得する
 * @param {string} html - ページ全体のHTML文字列
 * @returns {string[]} 各テーブルブロックのHTML文字列配列
 */
const extractTableBlocks = (html) =>
  Parser.data(html).iterate(
    '<div class="module-sub-page-parts-table">',
    "</div>",
  );

/**
 * 1つのテーブルブロックからdt/ddのラベルと値のマップを抽出する
 * dlの出現順が不定・一部欠落があっても対応できるよう、dtのテキストをキーにする
 * @param {string} tableBlock - module-sub-page-parts-table 1件分のHTML文字列
 * @returns {Object<string, string>} ラベルをキー、値を値とするマップ
 */
const extractFieldMap = (tableBlock) => {
  const dlBlocks = Parser.data(tableBlock).iterate("<dl>", "</dl>");

  const fieldMap = {};
  for (const block of dlBlocks) {
    const label = Parser.data(block).from("<dt>").to("</dt>").build().trim();
    if (!label) continue;

    const rawValue = Parser.data(block).from("<dd>").to("</dd>").build();
    const value = rawValue.replace(/<[^>]*>/g, "").trim();
    fieldMap[label] = value;
  }
  return fieldMap;
};

/**
 * 「事件番号」「事件名」「裁判年月日」「裁判所名」の4項目がすべて揃っているテーブルブロックを探す
 * ページ内に同種のテーブルが複数存在する場合があるため、4項目が揃うブロックのみを対象とする
 * @param {string} html - ページ全体のHTML文字列
 * @returns {Object<string, string>|null} 4項目揃ったfieldMap。見つからなければnull
 */
const findTargetFieldMap = (html) => {
  const tableBlocks = extractTableBlocks(html);

  for (const tableBlock of tableBlocks) {
    const fieldMap = extractFieldMap(tableBlock);
    const hasAllFields = FIELD_COLUMNS.every(({ label }) => fieldMap[label]);
    if (!hasAllFields) continue;
    return fieldMap;
  }
  return null;
};

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
 * B列（事件番号列）に対してTextFinderで完全一致の空白セルを検索し、
 * startRow以降で最初に見つかった行番号を返す
 * B列が空欄 = A列しか埋まっていない未処理行、という前提に基づく判定
 * 全行をgetValuesで読み込まず、検索自体をスプレッドシート側に任せることで高速化している
 * @param {number} startRow - 検索を開始する行番号（1始まり）
 * @param {number} lastRow - シートの最終行番号
 * @returns {number|null} 見つかった行番号。見つからなければnull
 */
const findFirstUnscrapedRow = (startRow, lastRow) => {
  if (lastRow < startRow) return null;

  const searchRange = MAIN_SHEET.getRange(
    startRow,
    CASE_NUMBER_COLUMN,
    lastRow - startRow + 1,
    1,
  );
  const finder = searchRange
    .createTextFinder("^$")
    .matchEntireCell(true)
    .useRegularExpression(true);
  const foundCell = finder.findNext();
  if (!foundCell) return null;

  return foundCell.getRow();
};

/**
 * これ以上スクレイピングする必要がない旨をメール通知する
 * @returns {void}
 */
const notifyAllScraped = () => {
  MailApp.sendEmail({
    to: MAIL_TO,
    subject: "スクレイピング完了",
    body: "MAIN_SHEETの全行のB列以降が埋まりました。これ以上スクレイピングする必要はありません。",
  });
};

/**
 * MAIN_SHEETの未処理行を1件見つけて判例情報をスクレイピングし、B〜E列に書き込む
 *
 * 処理の流れ:
 * 1. LOG_SHEETのA1から探索開始行を取得する
 * 2. TextFinderでB列が空欄の行（未処理行）を探す
 * 3. 見つかった場合はページにアクセスし、4項目すべて取得できればB〜E列に書き込む
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
  const isFullScan = startRow === 1;

  const targetRow = findFirstUnscrapedRow(startRow, lastRow);

  if (!targetRow) {
    // 先頭から探索して見つからなかった場合のみ、全行が埋まっていると確定できる
    if (isFullScan) notifyAllScraped();
    return;
  }

  const caseNumber = MAIN_SHEET.getRange(targetRow, 1).getValue();
  const response = fetchResponse(caseNumber);

  if (response.getResponseCode() === 200) {
    const fieldMap = findTargetFieldMap(response.getContentText());

    if (fieldMap) {
      for (const { label, column } of FIELD_COLUMNS) {
        MAIN_SHEET.getRange(targetRow, column).setValue(fieldMap[label]);
      }
    }
  }
  // 404、対象要素が見つからない、4項目が揃っていない場合は何も書き込まない
  // → 次回以降も未処理行として再チェック対象になる

  const nextRowIndex = targetRow < lastRow ? targetRow + 1 : 1;
  setLastRowIndex(nextRowIndex);
};
