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
  sheet.getRange('E3:E33').clearcontent();
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

//2019年06月09日　祝日設定トリガー試験 Start
// 休日を設定 (iCal)
function MakeHoliday()
{
    var calendarId = 'ja.japanese#holiday@group.v.calendar.google.com';
    var calendar = CalendarApp.getCalendarById(calendarId);
    var startDate = DateUtils.now();
    var endDate = new Date(startDate.getFullYear() + 1, startDate.getMonth());
    var holidays = _.map(calendar.getEvents(startDate, endDate), function(ev) {
      return DateUtils.format("Y-m-d", ev.getAllDayStartDate());
    });
    settings.set('休日', holidays.join(', '));
    settings.setNote('休日', '日付を,区切りで。来年までは自動設定されているので、以後は適当に更新してください');
}

// 日付関係の関数
DateUtils = loadDateUtils();
function loadDateUtils() {
  var DateUtils = {};

  // 今を返す
  var _now = new Date();
  var now = function(datetime) {
    if(typeof datetime != 'undefined') {
      _now = datetime;
    }
    return _now;
  };
  DateUtils.now = now;

  // テキストから時間を抽出
  DateUtils.parseTime = function(str) {
    str = String(str || "").toLowerCase().replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(s) {
      return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
    });
    var reg = /((\d{1,2})\s*[:時]{1}\s*(\d{1,2})\s*(pm|)|(am|pm|午前|午後)\s*(\d{1,2})(\s*[:時]\s*(\d{1,2})|)|(\d{1,2})(\s*[:時]{1}\s*(\d{1,2})|)(am|pm)|(\d{1,2})\s*時)/;
    var matches = str.match(reg);
    if(matches) {
      var hour, min;

      // 1時20, 2:30, 3:00pm
      if(matches[2] != null) {
        hour = parseInt(matches[2], 10);
        min = parseInt((matches[3] ? matches[3] : '0'), 10);
        if(_.contains(['pm'], matches[4])) {
          hour += 12;
        }
      }

      // 午後1 午後2時30 pm3
      if(matches[5] != null) {
        hour = parseInt(matches[6], 10);
        min = parseInt((matches[8] ? matches[8] : '0'), 10);
        if(_.contains(['pm', '午後'], matches[5])) {
          hour += 12;
        }
      }

      // 1am 2:30pm
      if(matches[9] != null) {
        hour = parseInt(matches[9], 10);
        min = parseInt((matches[11] ? matches[11] : '0'), 10);
        if(_.contains(['pm'], matches[12])) {
          hour += 12;
        }
      }

      // 14時
      if(matches[13] != null) {
        hour = parseInt(matches[13], 10);
        min = 0;
      }

      return [hour, min];
    }
    return null;
  };

  // テキストから日付を抽出
  DateUtils.parseDate = function(str) {
    str = String(str || "").toLowerCase().replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(s) {
      return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
    });

    if(str.match(/(明日|tomorrow)/)) {
      var tomorrow = new Date(now().getFullYear(), now().getMonth(), now().getDate()+1);
      return [tomorrow.getFullYear(), tomorrow.getMonth()+1, tomorrow.getDate()]
    }

    if(str.match(/(今日|today)/)) {
      return [now().getFullYear(), now().getMonth()+1, now().getDate()]
    }

    if(str.match(/(昨日|yesterday)/)) {
      var yesterday = new Date(now().getFullYear(), now().getMonth(), now().getDate()-1);
      return [yesterday.getFullYear(), yesterday.getMonth()+1, yesterday.getDate()]
    }

    var reg = /((\d{4})[-\/年]{1}|)(\d{1,2})[-\/月]{1}(\d{1,2})/;
    var matches = str.match(reg);
    if(matches) {
      var year = parseInt(matches[2], 10);
      var month = parseInt(matches[3], 10);
      var day = parseInt(matches[4], 10);
      if(_.isNaN(year) || year < 1970) {
        //
        if((now().getMonth() + 1) >= 11 && month <= 2) {
          year = now().getFullYear() + 1;
        }
        else if((now().getMonth() + 1) <= 2 && month >= 11) {
          year = now().getFullYear() - 1;
        }
        else {
          year = now().getFullYear();
        }
      }

      return [year, month, day];
    }

    return null;
  };

  // 日付と時間の配列から、Dateオブジェクトを生成
  DateUtils.normalizeDateTime = function(date, time) {
    // 時間だけの場合は日付を補完する
    if(date) {
      if(!time) date = null;
    }
    else {
      date = [now().getFullYear(), now().getMonth()+1, now().getDate()];
      if(!time) {
        time = [now().getHours(), now().getMinutes()];
      }
    }

    // 日付を指定したけど、時間を書いてない場合は扱わない
    if(date && time) {
      return(new Date(date[0], date[1]-1, date[2], time[0], time[1], 0));
    }
    else {
      return null;
    }
  };

  // 日時をいれてparseする
  DateUtils.parseDateTime = function(str) {
    var date = DateUtils.parseDate(str);
    var time = DateUtils.parseTime(str);
    if(!date) return null;
    if(time) {
      return(new Date(date[0], date[1]-1, date[2], time[0], time[1], 0));
    }
    else {
      return(new Date(date[0], date[1]-1, date[2], 0, 0, 0));
    }
  };

  // Dateから日付部分だけを取り出す
  DateUtils.toDate = function(date) {
    return(new Date(date.getFullYear(), date.getMonth(), date.getDate(), 0, 0, 0));
  };

  // 曜日を解析
  DateUtils.parseWday = function(str) {
    str = String(str).replace(/曜日/g, '');
    var result = [];
    var wdays = [/(sun|日)/i, /(mon|月)/i, /(tue|火)/i, /(wed|水)/i, /(thu|木)/i, /(fri|金)/i, /(sat|土)/i];
    for(var i=0; i<wdays.length; ++i) {
      if(str.match(wdays[i])) result.push(i);
    }
    return result;
  }

  var replaceChars = {
    Y: function() { return this.getFullYear(); },
    y: function() { return String(this.getFullYear()).substr(-2, 2); },
    m: function() { return ("0"+(this.getMonth()+1)).substr(-2, 2); },
    d: function() { return ("0"+(this.getDate())).substr(-2, 2); },

    H: function() { return ("0"+(this.getHours())).substr(-2, 2); },
    M: function() { return ("0"+(this.getMinutes())).substr(-2, 2); },
    s: function() { return ("0"+(this.getSeconds())).substr(-2, 2); },
  };

  DateUtils.format = function(format, date) {
    var result = '';
    for (var i = 0; i < format.length; i++) {
      var curChar = format.charAt(i);
      if (replaceChars[curChar]) {
        result += replaceChars[curChar].call(date);
      }
      else {
        result += curChar;
      }
    }
    return result;
  };

  return DateUtils;
};

if(typeof exports !== 'undefined') {
  exports.DateUtils = loadDateUtils();
}
// 日付関係の関数
// EventListener = loadEventListener();

loadEventListener = function () {
  var EventListener = function() {
    this._events = {};
  }

  // イベントを捕捉
  EventListener.prototype.on = function(eventName, func) {
    if(this._events[eventName]) {
      this._events[eventName].push(func);
    }
    else {
      this._events[eventName] = [func];
    }
  };

  // イベント発行
  EventListener.prototype.fireEvent = function(eventName) {
    var funcs = this._events[eventName];
    if(funcs) {
      for(var i = 0; i < funcs.length; ++i) {
        funcs[i].apply(this, Array.prototype.slice.call(arguments, 1));
      }
    }
  };

  return EventListener;
};

if(typeof exports !== 'undefined') {
  exports.EventListener = loadEventListener();
}
// KVS
// でも今回は使ってないです

loadGASProperties = function (exports) {
  var GASProperties = function() {
     this.properties = PropertiesService.getScriptProperties();
  };

  GASProperties.prototype.get = function(key) {
    return this.properties.getProperty(key);
  };

  GASProperties.prototype.set = function(key, val) {
    this.properties.setProperty(key, val);
    return val;
  };

  return GASProperties;
};

if(typeof exports !== 'undefined') {
  exports.GASProperties = loadGASProperties();
}
// Google Apps Script専用ユーティリティ

// GASのログ出力をブラウザ互換にする
if(typeof(console) == 'undefined' && typeof(Logger) != 'undefined') {
  console = {};
  console.log = function() {
    Logger.log(Array.prototype.slice.call(arguments).join(', '));
  }
}

// サーバに新しいバージョンが無いかチェックする
checkUpdate = function(responder) {
  if(typeof GASProperties === 'undefined') GASProperties = loadGASProperties();
  var current_version = parseFloat(new GASProperties().get('version')) || 0;

  var response = UrlFetchApp.fetch("https://raw.githubusercontent.com/masuidrive/miyamoto/master/VERSION", {muteHttpExceptions: true});

  if(response.getResponseCode() == 200) {
    var latest_version = parseFloat(response.getContentText());
    if(latest_version > 0 && latest_version > current_version) {
      responder.send("最新のみやもとさんの準備ができました！\nhttps://github.com/masuidrive/miyamoto/blob/master/UPDATE.md を読んでください。");

      var response = UrlFetchApp.fetch("https://raw.githubusercontent.com/masuidrive/miyamoto/master/HISTORY.md", {muteHttpExceptions: true});
      if(response.getResponseCode() == 200) {
        var text = String(response.getContentText()).replace(new RegExp("## "+current_version+"[\\s\\S]*", "m"), '');
        responder.send(text);
      }
    }
  }
};
// KVS

loadGSProperties = function (exports) {
  var GSProperties = function(spreadsheet) {
    // 初期設定
    this.sheet = spreadsheet.getSheetByName('_設定');
    if(!this.sheet) {
      this.sheet = spreadsheet.insertSheet('_設定');
    }
  };

  GSProperties.prototype.get = function(key) {
    if(this.sheet.getLastRow() < 1) return defaultValue;
    var vals = _.find(this.sheet.getRange("A1:B"+this.sheet.getLastRow()).getValues(), function(v) {
      return(v[0] == key);
    });
    if(vals) {
      if(_.isDate(vals[1])) {
        return DateUtils.format("Y-m-d H:M:s", vals[1]);
      }
      else {
        return String(vals[1]);
      }
    }
    else {
      return null;
    }
  };

  GSProperties.prototype.set = function(key, val) {
    if(this.sheet.getLastRow() > 0) {
      var vals = this.sheet.getRange("A1:A"+this.sheet.getLastRow()).getValues();
      for(var i = 0; i < this.sheet.getLastRow(); ++i) {
        if(vals[i][0] == key) {
          this.sheet.getRange("B"+(i+1)).setValue(String(val));
          return val;
        }
      }
    }
    this.sheet.getRange("A"+(this.sheet.getLastRow()+1)+":B"+(this.sheet.getLastRow()+1)).setValues([[key, val]]);
    return val;
  };

  GSProperties.prototype.setNote = function(key, note) {
    if(this.sheet.getLastRow() > 0) {
      var vals = this.sheet.getRange("A1:A"+this.sheet.getLastRow()).getValues();
      for(var i = 0; i < this.sheet.getLastRow(); ++i) {
        if(vals[i][0] == key) {
          this.sheet.getRange("C"+(i+1)).setValue(note);
          return;
        }
      }
    }
    this.sheet.getRange("A"+(this.sheet.getLastRow()+1)+":C"+(this.sheet.getLastRow()+1)).setValues([[key, '', note]]);
    return;
  };

  return GSProperties;
};
//2019年06月09日　祝日設定トリガー試験 End