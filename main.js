/*
MAIN_SHEETはヘッダなしで、A列に10000～の数値が入っている。
以下のscrape を数分に1回のペースで定期的に実行し、B列以降を埋めていく。
LOG_SHEETは最初の時点では何もデータが入っていない。

scpateがすること：
log_sheetのA1を見に行く
→空欄ならとりあえず10000を取得したことにする
→MAIN_SHEETのA列がその値の行を見に行く
→その行がA列以外空欄なら、その値を `https://www.courts.go.jp/hanrei/XXXXX/detail7/index.html` の `XXXXX` に入れてアクセスし、特定要素をスクレイピングしてB~E列に保存
→もしA列以外が既に埋まっていたら（その行をスクレイピング済みだったら）行を下に見て行き、A列しか埋まっていない行を探す
→行が見つかったらその行のA列の値を前記URLの `XXXXX` に入れて同じことをする。
→いずれの場合も、スクレイピングが終わったらその行の番号をLOG_SHEETのA1に記録しておく
→もし最終行まで見ても「A列以外が空欄」の行がなく、かつ今回の探索が先頭行から始まっていた場合（＝全行を1周確認した結果）、 `MAIL_TO` にメール通知する
→スクレイピングした行がシートの最終行だった場合、次回以降はまたシートの先頭（10000）から見ていく
→行の番号によってはURLが404になることがあるが、どこかのタイミングでページが作成されることがある。なので、このように再度ループすることでチェックしなおすことが可能になる。
*/

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
 */
const fetchResponse = (caseNumber) => {
  const url = `${BASE_URL}${caseNumber}/detail7/index.html`;
  return UrlFetchApp.fetch(url, { muteHttpExceptions: true });
};

/**
 * HTML内の module-sub-page-parts-table ブロックをすべて抜き出す
 */
const extractTableBlocks = (html) =>
  Parser.data(html).iterate(
    '<div class="module-sub-page-parts-table">',
    "</div>",
  );

/**
 * 1つのテーブルブロックからdt/ddのラベルと値のマップを抽出する
 * dlの順番が不定・一部欠落があっても対応できるよう、dtのテキストをキーにする
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
 * 「事件番号」「事件名」「裁判年月日」「裁判所名」の4項目がすべて揃っているテーブルブロックを探し、
 * 見つかればそのfieldMapを返す。見つからなければnullを返す
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
 * LOG_SHEETのA1から前回処理した行番号を取得する
 * 空欄なら1行目（先頭）として扱う
 */
const getLastRowIndex = () => {
  const value = LOG_SHEET.getRange("A1").getValue();
  if (value === "") return 1;
  return Number(value);
};

/**
 * LOG_SHEETのA1に処理済み行番号を記録する
 */
const setLastRowIndex = (rowIndex) => {
  LOG_SHEET.getRange("A1").setValue(rowIndex);
};

/**
 * 指定行が未スクレイピング（B列以降が空欄）かどうかを判定する
 */
const isUnscraped = (rowValues) => {
  for (let i = 1; i < rowValues.length; i++) {
    if (rowValues[i] !== "") return false;
  }
  return true;
};

/**
 * これ以上スクレイピングする必要がない旨をメール通知する
 */
const notifyAllScraped = () => {
  MailApp.sendEmail({
    to: MAIL_TO,
    subject: "スクレイピング完了",
    body: "MAIN_SHEETの全行のB列以降が埋まりました。これ以上スクレイピングする必要はありません。",
  });
};

const scrape = () => {
  const lastRow = MAIN_SHEET.getLastRow();
  if (lastRow < 1) return;

  const dataRange = MAIN_SHEET.getRange(
    1,
    1,
    lastRow,
    MAIN_SHEET.getLastColumn(),
  );
  const data = dataRange.getValues();

  const startRowIndex = getLastRowIndex();
  const startIndex =
    1 <= startRowIndex && startRowIndex <= lastRow ? startRowIndex - 1 : 0;
  const isFullScan = startIndex === 0;

  let targetIndex = -1;
  for (let i = startIndex; i < data.length; i++) {
    if (!isUnscraped(data[i])) continue;
    targetIndex = i;
    break;
  }

  if (targetIndex === -1) {
    // 先頭から探索して見つからなかった場合のみ、全行が埋まっていると確定できる
    if (isFullScan) notifyAllScraped();
    return;
  }

  const caseNumber = data[targetIndex][0];
  const targetRow = targetIndex + 1;
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
