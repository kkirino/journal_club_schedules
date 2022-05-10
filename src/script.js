function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("サイドバーを開く")
    .addItem("開く", "openSidebar")
    .addToUi();
}

function openSidebar() {
  const htmlOutput = HtmlService.createTemplateFromFile("sidebar").evaluate();
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

const ss = SpreadsheetApp.getActiveSpreadsheet();

function getSettings() {
  const ws = ss.getSheetByName("settings");
  const titles = ws
    .getRange(2, 1, ws.getLastRow() - 1, 1)
    .getValues()
    .map(function (row) {
      return row[0];
    });
  const values = ws
    .getRange(2, 2, ws.getLastRow() - 1, 1)
    .getValues()
    .map(function (row) {
      return row[0];
    });
  const dayOfWeekArray = ["日", "月", "火", "水", "木", "金", "土"];
  const settings = {};
  settings.dayOfWeek = dayOfWeekArray.indexOf(
    values[titles.indexOf("抄読会開催曜日")]
  );
  settings.startHours = values[titles.indexOf("開始時")];
  settings.startMinutes = values[titles.indexOf("開始分")];
  settings.endHours = values[titles.indexOf("終了時")];
  settings.endMinutes = values[titles.indexOf("終了分")];
  settings.meetingUrl = values[titles.indexOf("ミーティングURL")];
  settings.calendarId = values[titles.indexOf("カレンダーID")];
  settings.mailAdress = values[titles.indexOf("送信先メールアドレス")];
  settings.manualPageUrl = values[titles.indexOf("マニュアルページのURL")];
  settings.webAppUrl = values[titles.indexOf("WebアプリのURL")];
  return settings;
}

function getDoctorNamesArray() {
  const ws = ss.getSheetByName("schedule");
  const values = ws.getRange(2, 4, ws.getLastRow() - 1, 2).getValues();
  const array = [];
  for (let i = 0; i < values.length; i++) {
    for (let j = 0; j < values[0].length; j++) {
      if (values[i][j] !== "") {
        array.push(values[i][j]);
      }
    }
  }
  const set = new Set(array);
  const arrayFromSet = Array.from(set);
  return arrayFromSet;
}

function getCandidates() {
  const array = getDoctorNamesArray();
  const output = {};
  array.forEach(function (e) {
    output[e] = null;
  });
  return output;
}

function createSheetForSelect() {
  const businessYear = getBusinessYear();
  const dateStartThisBusinessYear = new Date(businessYear, 3, 1);
  const dateEndThisBusinessYear = new Date(businessYear + 1, 3, 1);
  const now = new Date();
  const yearNow = now.getMonth();
  const monthNow = now.getMonth();
  const dateStartThisThreeMonth = new Date();
  dateStartThisThreeMonth.setFullYear(yearNow - Number(monthNow < 3));
  dateStartThisThreeMonth.setMonth(monthNow + 12 * Number(monthNow < 3) - 3);
  const scheduleInfo = getScheduleInfo();
  const scheduleThisBusinessYear = scheduleInfo.filter(function (schedule) {
    return (
      schedule.dateTime.valueOf() > dateStartThisBusinessYear.valueOf() &&
      schedule.dateTime.valueOf() < dateEndThisBusinessYear.valueOf() &&
      schedule.type === "抄読会"
    );
  });
  const scheduleThisThreeMonth = scheduleInfo.filter(function (schedule) {
    return schedule.dateTime.valueOf() > dateStartThisThreeMonth;
  });
  const doctorNames = getDoctorNamesArray();
  const arrayForWriteSheet = doctorNames.map(function (name) {
    var numThisBusinessYear = 0;
    var numThisThreeMonth = 0;
    var latestDate = new Date(dateStartThisThreeMonth);
    scheduleThisBusinessYear.forEach(function (schedule) {
      if (
        schedule.presenterFirst === name ||
        schedule.presenterSecond === name
      ) {
        numThisBusinessYear += 1;
      }
    });
    scheduleThisThreeMonth.forEach(function (schedule) {
      if (
        schedule.presenterFirst === name ||
        schedule.presenterSecond === name
      ) {
        numThisThreeMonth += 1;
        if (schedule.dateTime.valueOf() > latestDate.valueOf()) {
          latestDate = schedule.dateTime;
        }
      }
    });
    if (latestDate === dateStartThisThreeMonth) {
      latestDate = "";
    }
    return [name, numThisBusinessYear, numThisThreeMonth, latestDate];
  });
  infoForSelect.sort(function (a, b) {
    return b.numPresent - a.numPresent;
  });
  Logger.log(arrayForWriteSheet);
  const tmp = ss.getSheetByName("tmp");
  const ws = tmp.copyTo(ss);
  ws.setName(now.toLocaleDateString() + " " + now.toLocaleTimeString());
  ws.getRange(3, 1, arrayForWriteSheet.length, 4).setValues(arrayForWriteSheet);
  return "担当状況を記載したシートを作成しました";
}

function userClicked(userInfo) {
  if (
    userInfo.date == "" ||
    userInfo.type == "" ||
    (userInfo.type == "抄読会" && userInfo.presenterFirst == "") ||
    (userInfo.type == "抄読会" && userInfo.presenterSecond == "") ||
    (userInfo.type == "ポスグラ" && userInfo.presenterPgc == "") ||
    (userInfo.type == "その他" && userInfo.otherInformation == "")
  ) {
    return "フォームをすべて埋めてください";
  } else {
    const ws = ss.getSheetByName("schedule");
    const isInputDateInEachRow = ws
      .getRange(2, 2, ws.getLastRow() - 1, 1)
      .getValues()
      .map(function (row) {
        return row[0] == userInfo.date;
      });
    const inputDateInSheetIndex = isInputDateInEachRow.indexOf(true);
    if (inputDateInSheetIndex > -1) {
      ws.deleteRow(inputDateInSheetIndex + 2);
    }
    const settings = getSettings();
    const startTime = new Date(userInfo.date);
    startTime.setHours(settings.startHours);
    startTime.setMinutes(settings.startMinutes);
    const endTime = new Date(userInfo.date);
    endTime.setHours(settings.endHours);
    endTime.setMinutes(settings.endMinutes);
    const calendar = CalendarApp.getCalendarById(settings.calendarId);
    let title = "【抄読会】";
    if (userInfo.type === "抄読会") {
      title +=
        userInfo.presenterFirst +
        ", " +
        userInfo.presenterSecond +
        ", " +
        userInfo.otherInformation;
    } else if (userInfo.type === "ポスグラ") {
      title +=
        "ポスグラ, " + userInfo.presenterPgc + ", " + userInfo.otherInformation;
    } else if (userInfo.type === "休会") {
      title += "休会, " + userInfo.otherInformation;
    } else {
      title += userInfo.otherInformation;
    }
    const option = {
      description: settings.meetingUrl,
    };
    const events = calendar.getEvents(startTime, endTime, {
      search: "【抄読会】",
    });
    events.forEach(function (e) {
      e.deleteEvent();
    });
    const holidayCalendar = CalendarApp.getCalendarById(
      "ja.japanese#holiday@group.v.calendar.google.com"
    );
    const isHoliday = holidayCalendar.getEventsForDay(startTime).length > 0;
    if (!(isHoliday && userInfo.type === "休会")) {
      calendar.createEvent(title, startTime, endTime, option);
    }
    ws.appendRow([
      new Date(),
      userInfo.date,
      userInfo.type,
      userInfo.presenterFirst,
      userInfo.presenterSecond,
      userInfo.presenterPgc,
      userInfo.otherInformation,
    ]);
    return "抄読会予定を入力してカレンダーを作成しました";
  }
}

function doGet() {
  return HtmlService.createTemplateFromFile("index").evaluate();
}

function getBusinessYear() {
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth() + 1;
  const businessYear = year - Number(month < 4);
  return businessYear;
}

function getOption() {
  const businessYear = getBusinessYear();
  let tableText = "<option value='0' selected>Future Schedules </option>";
  for (let y = businessYear; y >= 2020; y--) {
    let yearText =
      "<option value='" +
      y +
      "'>Apr. " +
      y +
      " - Mar. " +
      (y + 1) +
      "</option>";
    tableText = tableText + yearText;
  }
  return tableText;
}

function getScheduleInfo() {
  const ws = ss.getSheetByName("schedule");
  const scheduleInSheet = ws
    .getRange(2, 1, ws.getLastRow() - 1, ws.getLastColumn())
    .getValues();
  const scheduleInfo = scheduleInSheet.map(function (row) {
    const dateTime = new Date(row[1]);
    dateTime.setHours(7);
    dateTime.setMinutes(30);
    return {
      date: row[1],
      dateTime: dateTime,
      type: row[2],
      presenterFirst: row[3],
      presenterSecond: row[4],
      presenterPgc: row[5],
      otherInformation: row[6],
    };
  });
  return scheduleInfo;
}

function getTable(yearType) {
  const scheduleInfo = getScheduleInfo();
  function compareDate(a, b) {
    return a.dateTime.valueOf() - b.dateTime.valueOf();
  }
  scheduleInfo.sort(compareDate);
  const now = new Date();
  const yearStart = new Date(yearType, 3, 1, 0, 0, 0, 0);
  Logger.log(yearStart);
  const nextYearStart = new Date(Number(yearType) + 1, 3, 1, 0, 0, 0, 0);
  Logger.log(nextYearStart);
  const filteredSchedule = scheduleInfo.filter(function (schedule) {
    if (yearType == 0) {
      return schedule.dateTime.valueOf() > now.valueOf();
    } else {
      return (
        schedule.dateTime.valueOf() > yearStart.valueOf() &&
        schedule.dateTime.valueOf() < nextYearStart.valueOf()
      );
    }
  });
  let tableText = "";
  filteredSchedule.forEach(function (schedule) {
    let rowText =
      "<td>" + schedule.date + "</td><td>" + schedule.type + "</td>";
    if (schedule.type == "抄読会") {
      rowText +=
        "<td>" +
        schedule.presenterFirst +
        " / " +
        schedule.presenterSecond +
        "</td>";
    } else if (schedule.type == "ポスグラ") {
      rowText += "<td>" + schedule.presenterPgc + "</td>";
    } else {
      rowText = rowText + "<td></td>";
    }
    rowText =
      "<tr>" + rowText + "<td>" + schedule.otherInformation + "</td></tr>";
    tableText += rowText;
  });
  return tableText;
}
