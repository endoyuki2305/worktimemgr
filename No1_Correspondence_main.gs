/**
 * スプレッドシート起動時に呼ばれる
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('カスタム');
  menu.addItem('実勤務時間を再計算する', 'reCalcWorkingTime');
  menu.addItem('ダイアログ(開発中)...', 'showModalDialog');
  menu.addToUi();
}

/**
 * 勤怠管理表のデータを入力するダイアログを表示する
 */
function showModalDialog() {
  var html = HtmlService
               .createTemplateFromFile('dialog')
               .evaluate()
               .setSandboxMode(HtmlService.SandboxMode.IFRAME)
               .setWidth(250)
               .setHeight(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * 月間合計シートから勤務時間を取得
 */
function reCalcWorkingTime() {
  var SLACK_TIMESHEETS_ID = '1WjdgdQihmUwSX2Pxge1Lfr3b6Jh8w1s5fiNyDdqX8tw';
  var sheets = SpreadsheetApp.openById(SLACK_TIMESHEETS_ID);
  var monthly = sheets.getSheetByName('月間合計');
  var customerName = monthly.getRange('H1').getValue();
  // 客先別休憩時間を取得
  var restTime = getRestTime(sheets, customerName);
  // 勤務時間を取得
  var workingTime = getWorkingTime(monthly);
  // 勤務時間に休憩時間を適用する
  var appliedTime = applyBreakTime(workingTime, restTime);
  // 「月間合計」シートに休憩時間を適用した勤務時間を出力
  writeCalculatedTime(monthly, appliedTime);
}

/**
 * minutesをHH:mm形式で返す
 * ex) 90 -> 1:30
 * @param {number} minutes
 * @returns {string} HH:mm
 */
function getTimeFormat(minutes) {
  return Math.floor(minutes / 60) + ':' + ('00' + (minutes % 60)).slice(-2);
}

/**
 * HH:mm形式の時間を分で返す
 * ex) 1:30 -> 90
 * @returns {string} 分
 */
function getMinutes(hour) {
  var a = hour.split(':');
  return Number(a[0] * 60) + Number(a[1]);
}

/**
 * 勤務時間から休憩時間を引く
 * @param {Object} working 勤務時間 [[出勤,退勤,勤務時間],[...]]
 * @param {Object} rest 休憩時間
 * @returns 休憩時間適用後の勤務時間
 */
function applyBreakTime(working, rest) {
  var appliedAry = [];
  for (var i = 0; i < working.length; i++) {
    // 出退勤時刻、または休み'-'なら何もしない
    if (!working[i][0] || !working[i][1] || !working[i][2] ||
       working[i][0] === '-' || working[i][1] === '-' || working[i][2] === '-') {
      appliedAry.push(['']);
      continue;
    }
    appliedAry.push([getTimeFormat(subtractBreakTime(working[i], rest))]);
  }
  return appliedAry;
}

/**
 * 勤務時間から休憩時間を引く
 * @param {array} today 1日分の勤務時間 [出勤,退勤,勤務時間]
 * @param {object} rest 休憩時間
 * @returns {number} 1日分の勤務時間から休憩時間を引いた分数
 */
function subtractBreakTime(today, rest) {
  var workStart = Moment.moment(today[0]);
  var workEnd   = Moment.moment(today[1]);
  var workTime  = getMinutes(Moment.moment(today[2]).format("HH:mm"));
  var excludeTime = 0;
  for (k in rest) {
    var restStart = Moment.moment(workStart.format("YYYY-MM-DD") + rest[k]['start'], "YYYY-MM-DD HH:mm");
    var restEnd   = Moment.moment(workStart.format("YYYY-MM-DD") + rest[k]['end'], "YYYY-MM-DD HH:mm");
    if (workStart.isBefore(restStart) && workEnd.isAfter(restEnd)) {
      var restTime = Number(rest[k]['rest']);
      excludeTime += restTime;
    }
  }
  return workTime - excludeTime;
}

/**
 * 客先タイムテーブルの休憩時間を計算した値を月間合計シートに書く
 * @param sheet 月間合計シート
 * @param appliedTime 休憩時間適用後の勤務時間
 */
function writeCalculatedTime(sheet, appliedTime) {
  // 2019年03月31日　以前の出力値削除対応　Start
  // 計算した勤務時間を出力する前に前回の出力値を削除する
  sheet.getRange('E3:E33').clearContent();
  // 2019年03月31日　以前の出力値削除対応　End
  
  // 計算した勤務時間をE列に出力する
  sheet.getRange('E3:E33').setValues(appliedTime);
}

/**
 * 1ヶ月の勤務時間を取得
 * @param sheet 月間合計シート
 * @returns 勤務時間 [[出勤,退勤,勤務時間],[...]]
 */
function getWorkingTime(sheet) {
  return sheet.getRange('B3:D33').getValues();
}

/**
 * 客先ごとの休憩タイムテーブルを取得
 * @param sheets Slack Timesheetsブックへの参照
 * @param customerName 客先会社名==休憩タイムテーブルのシート名
 * @returns 休憩時間 [{start:'開始時刻', end:'終了時刻', rest:'休憩時間(分)'}, {...}]
 */
function getRestTime(sheets, customerName) {
  var HEADER_NUM = 1; // ヘッダは1行
  var START_END_REST_COL = 3; // 休憩開始,休憩終了,休憩時間（分）の3列
  var sheet = sheets.getSheetByName(customerName);
  var range = sheet.getRange(2, 1, sheet.getLastRow() - HEADER_NUM, START_END_REST_COL);
  var restTable = [];
  
  for(var row = 1; row <= range.getNumRows(); row++) {
    var rest = {};
    rest['start'] = range.getCell(row, 1).getValue();
    rest['end'] = range.getCell(row, 2).getValue();
    rest['rest'] = range.getCell(row, 3).getValue();
    restTable.push(rest);
  }
  return restTable;
}
