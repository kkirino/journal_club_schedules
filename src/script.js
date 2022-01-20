function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("サイドバーを開く")
    .addItem("開く", "openSidebar")
    .addToUi();
}

function openSidebar() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile("sidebar").setTitle(
    "Input New Schedules"
  );
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

const ss = SpreadsheetApp.getActiveSpreadsheet();
const ws = ss.getSheetByName("schedule");
const numRow = ws.getLastRow() - 1;
const numCol = ws.getLastColumn();

function getJournalClubSettings() {
  const settingsSheet = ss.getSheetByName("settings");
  const settingsTitles = settingsSheet
    .getRange(2, 1, settingsSheet.getLastRow() - 1, 1)
    .getValues()
    .map(function (row) {
      return row[0];
    });
  const settingsValues = settingsSheet
    .getRange(2, 2, settingsSheet.getLastRow() - 1, 1)
    .getValues()
    .map(function (row) {
      return row[0];
    });

  const dayOfWeekArray = ["日", "月", "火", "水", "木", "金", "土"];
  const settings = {};
  settings.dayOfWeek = dayOfWeekArray.indexOf(
    settingsValues[settingsTitles.indexOf("抄読会開催曜日")]
  );
  settings.startHours = settingsValues[settingsTitles.indexOf("開始時")];
  settings.startMinutes = settingsValues[settingsTitles.indexOf("開始分")];
  settings.endHours = settingsValues[settingsTitles.indexOf("終了時")];
  settings.endMinutes = settingsValues[settingsTitles.indexOf("終了分")];
  settings.meetingUrl =
    settingsValues[settingsTitles.indexOf("ミーティングURL")];
  settings.calendarId = settingsValues[settingsTitles.indexOf("カレンダーID")];
  settings.manualUrl =
    settingsValues[settingsTitles.indexOf("操作マニュアルURL")];
  settings.websiteUrl = settingsValues[settingsTitles.indexOf("WebサイトURL")];
  settings.mailAdress =
    settingsValues[settingsTitles.indexOf("送信先メールアドレス")];

  return settings;
}

function getCandidates() {
  const values = ws.getRange(2, 4, numRow, 2).getValues();
  var array = [];
  for (let i = 0; i < values.length; i++) {
    for (let j = 0; j < values[0].length; j++) {
      if (values[i][j] !== "") {
        array.push(values[i][j]);
      }
    }
  }
  const set = new Set(array);
  const arrayFromSet = Array.from(set);

  var output = {};
  arrayFromSet.forEach(function (e) {
    output[e] = null;
  });
  return output;
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
    const isInputDateInEachRow = ws
      .getRange(2, 2, numRow, 1)
      .getValues()
      .map(function (row) {
        return row[0] == userInfo.date;
      });
    const inputDateInSheetIndex = isInputDateInEachRow.indexOf(true);
    if (inputDateInSheetIndex > -1) {
      ws.deleteRow(inputDateInSheetIndex + 2);
    }

    const settings = getJournalClubSettings();

    const startTime = new Date(userInfo.date);
    startTime.setHours(settings.startHours);
    startTime.setMinutes(settings.startMinutes);
    const endTime = new Date(userInfo.date);
    endTime.setHours(settings.endHours);
    endTime.setMinutes(settings.endMinutes);

    const calendar = CalendarApp.getCalendarById(settings.calendarId);
    if (userInfo.type == "抄読会") {
      var title =
        "【抄読会】" +
        userInfo.presenterFirst +
        ", " +
        userInfo.presenterSecond +
        ", " +
        userInfo.otherInformation;
    } else if (userInfo.type == "ポスグラ") {
      var title =
        "【抄読会】ポスグラ, " +
        userInfo.presenterPgc +
        ", " +
        userInfo.otherInformation;
    } else if (userInfo.type == "休会") {
      var title = "【抄読会】休会, " + userInfo.otherInformation;
    } else {
      var title = "【抄読会】" + userInfo.otherInformation;
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

    calendar.createEvent(title, startTime, endTime, option);

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

function sendEmail(message, isChecked) {
  const settings = getJournalClubSettings();
  if (isChecked === false) {
    return "チェックボックスにチェックを入れてください";
  } else {
    MailApp.sendEmail(
      settings.mailAdress,
      "【お知らせ】抄読会の予定【更新】",
      "みなさま\n\n抄読会の予定を更新しましたのでご連絡いたします。\nお手数ですが下記URLより日程と担当をご確認ください。\n" +
        settings.websiteUrl +
        "\n" +
        message +
        "\nよろしくお願いいたします。"
    );

    return "抄読会予定更新のメールを送信しました";
  }
}

function doGet() {
  return HtmlService.createTemplateFromFile("index").evaluate();
}

function getTable(yearType) {
  const scheduleInfo = ws
    .getRange(2, 1, numRow, numCol)
    .getValues()
    .map(function (row) {
      let dateTime = new Date(row[1]);
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

  function compareDate(a, b) {
    return a.dateTime.valueOf() - b.dateTime.valueOf();
  }
  scheduleInfo.sort(compareDate);

  const now = new Date();
  var tableText = "";

  if (yearType == 0) {
    for (let i = 0; i < scheduleInfo.length; i++) {
      if (scheduleInfo[i].dateTime.valueOf() > now.valueOf()) {
        let rowText =
          "<td>" +
          scheduleInfo[i].date +
          "</td><td>" +
          scheduleInfo[i].type +
          "</td>";

        if (scheduleInfo[i].type == "抄読会") {
          rowText +=
            "<td>" +
            scheduleInfo[i].presenterFirst +
            " / " +
            scheduleInfo[i].presenterSecond +
            "</td>";
        } else if (scheduleInfo[i].type == "ポスグラ") {
          rowText += "<td>" + scheduleInfo[i].presenterPgc + "</td>";
        } else {
          rowText = rowText + "<td></td>";
        }

        rowText =
          "<tr>" +
          rowText +
          "<td>" +
          scheduleInfo[i].otherInformation +
          "</td></tr>";
        tableText += rowText;
      }
    }
  } else {
    for (let i = 0; i < scheduleInfo.length; i++) {
      var yearStart = new Date();
      yearStart.setFullYear(yearType);
      yearStart.setMonth(3);
      yearStart.setDate(1);
      yearStart.setHours(0);
      yearStart.setMinutes(0);
      yearStart.setSeconds(0);

      var nextYearStart = new Date(yearStart);
      nextYearStart.setFullYear(Number(yearType) + 1);

      if (
        scheduleInfo[i].dateTime.valueOf() > yearStart.valueOf() &&
        scheduleInfo[i].dateTime.valueOf() < nextYearStart.valueOf()
      ) {
        let rowText =
          "<td>" +
          scheduleInfo[i].date +
          "</td><td>" +
          scheduleInfo[i].type +
          "</td>";

        if (scheduleInfo[i].type == "抄読会") {
          rowText +=
            "<td>" +
            scheduleInfo[i].presenterFirst +
            " / " +
            scheduleInfo[i].presenterSecond +
            "</td>";
        } else if (scheduleInfo[i].type == "ポスグラ") {
          rowText += "<td>" + scheduleInfo[i].presenterPgc + "</td>";
        } else {
          rowText = rowText + "<td></td>";
        }

        rowText =
          "<tr>" +
          rowText +
          "<td>" +
          scheduleInfo[i].otherInformation +
          "</td></tr>";
        tableText += rowText;
      }
    }
  }
  console.log(scheduleInfo);
  return tableText;
}
