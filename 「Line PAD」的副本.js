function doGet2(e) {
  const timezone = "Asia/Taipei";
  const sheetId = "1Uc-96Q1V4x1hI6VUOoKv4MC7AX-0KKbZGhlWEXbcVFM";
  const sheetName = "Line轉貼預排表";
  const historySheetName = "Line轉貼歷史";

  const ss = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheetByName(sheetName);
  const historySheet = ss.getSheetByName(historySheetName);

  const lastRow = sheet.getLastRow();
  const now = new Date();

  const allData = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const dateValues = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  const timeValues = sheet.getRange(2, 3, lastRow - 1, 1).getDisplayValues();

  console.log("=== [doGet] 開始執行 ===");
  console.log(`📌 現在時間: ${Utilities.formatDate(now, timezone, 'yyyy/MM/dd HH:mm:ss')}`);
  console.log(`📌 資料列數: ${allData.length}`);

  let taskList = [];

  for (let i = 0; i < allData.length; i++) {
    const row = allData[i];
    const status = row[9];           // J欄 Status
    const date = dateValues[i][0];   // B欄 日期
    const time = timeValues[i][0];   // C欄 時間（顯示值）
    const productID = row[4]?.toString().trim();   // E欄 商品編號
    const pushMessage = row[6]?.toString().trim(); // G欄 推播訊息

    if (!productID || !pushMessage) continue; // 必要欄位缺資料

    if (!status) {
      let sortKey;
      let shouldSend = false;

      if (!time) {
        sortKey = new Date(0); // 時間空白 → 優先
        shouldSend = true;
      } else if (date) {
        sortKey = new Date(`${Utilities.formatDate(date, timezone, "yyyy/MM/dd")} ${time}:00`);
        if (sortKey.getTime() <= now.getTime()) {
          shouldSend = true;
        }
      } else {
        sortKey = new Date(9999, 0, 1); // 無日期但有時間 → 不處理
      }

      if (shouldSend) {
        taskList.push({
          index: i + 2,
          row: row,
          sortTime: sortKey
        });
      }
    }
  }

  if (taskList.length === 0) {
    console.log("⛔ 沒有符合條件的未執行任務");
    return ContentService.createTextOutput("").setMimeType(ContentService.MimeType.TEXT);
  }

  taskList.sort((a, b) => a.sortTime - b.sortTime); // 空白優先、時間早的優先

  const selected = taskList[0];
  const rowIndex = selected.index;
  const rowData = selected.row;
  const productID = rowData[4];
  const pushMessage = rowData[6];

  // 更新原始表單狀態為「已轉貼」
  sheet.getRange(rowIndex, 10).setValue("已轉貼");

  // 新增完成任務到歷史表單最上面
  const rowWithTimestamp = [...rowData];
  rowWithTimestamp[9] = Utilities.formatDate(now, timezone, "yyyy/MM/dd HH:mm:ss");
  historySheet.insertRows(2, 1);
  historySheet.getRange(2, 1, 1, rowWithTimestamp.length).setValues([rowWithTimestamp]);

  // 刪除原列
  sheet.deleteRow(rowIndex);

  const result = `${productID}|${pushMessage}`;
  console.log(`🎯 傳送任務（第 ${rowIndex} 列）: ${result}`);

  return ContentService.createTextOutput(result).setMimeType(ContentService.MimeType.TEXT);
}
