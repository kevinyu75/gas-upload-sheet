function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('玩美街專屬功能')
    .addItem('跳至今日', 'jumpToToday')
    .addItem('跳至45天前', 'jumpTo45DaysAgo')
    .addItem('跳回指定天數', 'jumpToCustomDaysAgo')
    .addItem('排程FB互動', '將上架待排轉貼到_FB互動預排表')
    .addSeparator()
    .addItem('推播排程', 'scheduledPush')
    .addSeparator()
    .addItem('結單小幫手', 'listClosedOrders')
    .addSeparator()
    .addItem('文案助手單行', '生成Line推播文案_單行')
    .addItem('文案助手多行', '生成Line推播文案_多行')    
    .addToUi();

  jumpToToday();
}

// 包裝：單行
function 生成Line推播文案_單行() {
  生成Line推播文案_ActiveRow("single");
}

// 包裝：多行
function 生成Line推播文案_多行() {
  生成Line推播文案_ActiveRow("multi");
}

function jumpToToday() {
  jumpToDate(0);
}

function jumpTo45DaysAgo() {
  jumpToDate(45);
}

function jumpToCustomDaysAgo() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('跳回指定天數', '請輸入要跳回的天數:', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    var daysAgo = parseInt(response.getResponseText());
    if (!isNaN(daysAgo)) {
      jumpToDate(daysAgo);
    } else {
      ui.alert('無效的輸入', '請輸入有效的數字。', ui.ButtonSet.OK);
    }
  }
}

function jumpToDate(daysAgo) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var today = new Date();
  var targetDate = new Date(today);
  targetDate.setDate(targetDate.getDate() - daysAgo);

  var dateColumn = 10;
  var dateRange = sheet.getRange(1, dateColumn, sheet.getLastRow(), 1);
  var dates = dateRange.getValues();

  var targetYear = targetDate.getFullYear();
  var targetMonth = targetDate.getMonth();
  var targetDay = targetDate.getDate();

  var matchingRow = null;

  for (var i = 0; i < dates.length; i++) {
    var cellDate = dates[i][0];
    if (cellDate instanceof Date) {
      var cellYear = cellDate.getFullYear();
      var cellMonth = cellDate.getMonth();
      var cellDay = cellDate.getDate();

      if (cellYear === targetYear && cellMonth === targetMonth && cellDay === targetDay) {
        matchingRow = i + 1;
        break;
      }
    }
  }

  if (matchingRow !== null) {
    sheet.setActiveSelection(sheet.getRange(matchingRow, dateColumn));
    sheet.setCurrentCell(sheet.getRange(matchingRow, dateColumn));
  } else {
    SpreadsheetApp.getUi().alert('沒有找到符合的日期: ' + Utilities.formatDate(targetDate, Session.getScriptTimeZone(), "yyyy/MM/dd"));
  }
}

function listClosedOrders() {
  var html = `
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; text-align: center; }
      h3 { margin-bottom: 10px; }
      .btn {
        padding: 10px 20px; margin: 5px;
        font-size: 16px; cursor: pointer; border: none;
        border-radius: 5px;
      }
      .btn-today { background-color: #4CAF50; color: white; }
      .btn-yesterday { background-color: #2196F3; color: white; }
      .btn-submit { background-color: #FF9800; color: white; }

      input {
        padding: 10px; font-size: 16px; margin-top: 10px;
        width: 80%; text-align: center; border-radius: 5px;
        border: 1px solid #ccc;
      }

      #loading {
        display: none; font-size: 16px; color: #555;
        margin-top: 15px;
      }
    </style>

    <h3>結單小幫手</h3>
    <button class="btn btn-today" onclick="selectDate(0)">今日</button>
    <button class="btn btn-yesterday" onclick="selectDate(1)">昨日</button>
    <br>
    <input type="text" id="dateInput" placeholder="請輸入日期 (M/d)" onkeypress="handleKeyPress(event)">
    <br>
    <button class="btn btn-submit" onclick="submitDate()">確認</button>
    <p id="loading">🔄 結單小幫手執行中，請稍候...</p>

    <script>
      let isSubmitting = false;

      function showLoading() {
        document.getElementById("loading").style.display = "block";
      }

      function hideLoading() {
        document.getElementById("loading").style.display = "none";
      }

      function handleKeyPress(event) {
        if (event.keyCode === 13) {
          event.preventDefault();
          submitDate();
        }
      }

      function selectDate(daysAgo) {
        if (isSubmitting) return;
        isSubmitting = true;
        showLoading();

        const today = new Date();
        today.setDate(today.getDate() - daysAgo);
        const formattedDate = (today.getMonth() + 1) + '/' + today.getDate();

        google.script.run
          .withSuccessHandler(() => { isSubmitting = false; hideLoading(); })
          .withFailureHandler(() => { isSubmitting = false; hideLoading(); })
          .processSelectedDate(formattedDate);
      }

      function submitDate() {
        if (isSubmitting) return;
        isSubmitting = true;

        const dateValue = document.getElementById("dateInput").value.trim();
        if (dateValue) {
          showLoading();
          google.script.run
            .withSuccessHandler(() => { isSubmitting = false; hideLoading(); })
            .withFailureHandler(() => { isSubmitting = false; hideLoading(); })
            .processSelectedDate(dateValue);
        } else {
          alert("請輸入有效的日期 (M/d)");
          isSubmitting = false;
          hideLoading();
        }
      }

       // ✅ 新增自動聚焦功能
  window.onload = function() {
    document.getElementById("dateInput").focus();
  }; 
    </script>
  `;

  var htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(400)
    .setHeight(300);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "結單小幫手");
}

function processSelectedDate(selectedDateStr) {
  const user = Session.getActiveUser().getEmail() || "guest";
  const props = PropertiesService.getUserProperties();
  const lockKey = "in_progress_" + user;
  const isInProgress = props.getProperty(lockKey);

  if (isInProgress === "true") return;
  props.setProperty(lockKey, "true");

  try {
    const selectedDate = parseDateFromString(selectedDateStr);
    if (!selectedDate) {
      SpreadsheetApp.getUi().alert("錯誤", "請輸入正確的日期格式 (M/d)", SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    filterAndDisplayOrders(selectedDate);
  } finally {
    props.deleteProperty(lockKey);
  }
}

function filterAndDisplayOrders(selectedDate) {
  var formattedSelectedDate = Utilities.formatDate(selectedDate, Session.getScriptTimeZone(), "M/d");
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var dateColumn = 5;
  var timeColumn = 6; // 結單時段位於第6欄
  var dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  var values = dataRange.getValues();
  var filteredRows = [];

  const seenListingIds = new Set();

  for (var i = 0; i < values.length; i++) {
    var rowDate = values[i][dateColumn - 1];
    var match = false;

    if (rowDate instanceof Date) {
      match = isSameDate(rowDate, selectedDate);
    } else if (typeof rowDate === "string") {
      var parsed = parseDateFromString(rowDate.trim());
      if (parsed) match = isSameDate(parsed, selectedDate);
    }

    if (match) {
      const listingId = values[i][15]; // P 欄 上架編號
      if (!seenListingIds.has(listingId)) {
        seenListingIds.add(listingId);
        filteredRows.push([
          values[i][0],                 // A 欄 行銷
          formattedSelectedDate,        // 顯示統一的結單日期
          values[i][timeColumn - 1],    // F 欄 結單時段
          listingId,                    // P 欄 上架編號
          values[i][16],                // Q 欄 廠商
          truncateText(values[i][18], 40) // S 欄 商品名稱
        ]);
      }
    }
  }

  // 排序（行銷 > 廠商）
  filteredRows.sort((a, b) => {
    if (a[0] > b[0]) return 1;
    if (a[0] < b[0]) return -1;
    if (a[4] > b[4]) return 1;
    if (a[4] < b[4]) return -1;
    return 0;
  });

  if (filteredRows.length > 0) {
    var message = `
      <style>
        body { font-family: Arial, sans-serif; }
        table { width: 100%; border-collapse: collapse; }
        th, td { padding: 10px; border: 1px solid black; text-align: left; font-size: 14px; }
        th { background-color: #8C8C8C; color: white; }
      </style>
      <table>
        <tr>
          <th>行銷</th>
          <th>結單日期</th>
          <th>結單時段</th>
          <th>上架編號</th>
          <th>廠商</th>
          <th style='width: 250px;'>商品名稱</th>
        </tr>`;

    var lastMarketing = "";
    var color1 = "#F9F6F1";
    var color2 = "#E7E7E5";
    var currentColor = color1;

    for (var j = 0; j < filteredRows.length; j++) {
      if (filteredRows[j][0] !== lastMarketing) {
        currentColor = (currentColor === color1) ? color2 : color1;
        lastMarketing = filteredRows[j][0];
      }

      message += `<tr style="background-color: ${currentColor};">`;
      for (var k = 0; k < filteredRows[j].length; k++) {
        message += `<td>${filteredRows[j][k]}</td>`;
      }
      message += `</tr>`;
    }

    message += "</table>";

    var htmlOutput = HtmlService.createHtmlOutput("<h3>結單商品</h3>" + message)
      .setWidth(900)
      .setHeight(500);

    SpreadsheetApp.getUi().showModalDialog(htmlOutput, "結單商品");
  } else {
    SpreadsheetApp.getUi().alert("沒有符合的結單商品。");
  }
}



function isSameDate(date1, date2) {
  return date1.getFullYear() === date2.getFullYear() &&
         date1.getMonth() === date2.getMonth() &&
         date1.getDate() === date2.getDate();
}

function parseDateFromString(dateString) {
  var pattern = /^([1-9]|0[1-9]|1[0-2])\/([1-9]|0[1-9]|[12][0-9]|3[01])$/;
  var match = dateString.match(pattern);
  if (!match) return null;

  var today = new Date();
  var year = today.getFullYear();
  var month = parseInt(match[1], 10) - 1;
  var day = parseInt(match[2], 10);

  return new Date(year, month, day);
}



function truncateText(text, maxLength) {
  if (text.length > maxLength) {
    return text.substring(0, maxLength) + "...";
  }
  return text;
}