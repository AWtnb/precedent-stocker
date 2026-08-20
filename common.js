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

// スプレッドシートの列番号（B, C, D, E）と、対応するmeta要素のname属性の対応
const FIELD_COLUMNS = [
  { metaName: "composite_jiken_number", column: 2 }, // 事件番号
  { metaName: "jiken_name", column: 3 }, // 事件名
  { metaName: "judge_date_wareki", column: 4 }, // 裁判年月日
  { metaName: "court_name", column: 5 }, // 裁判所名
];
