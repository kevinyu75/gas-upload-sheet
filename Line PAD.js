function doGet(e) {
  const type = (e && e.parameter && e.parameter.type || "Line").toString().toLowerCase();
  if (type === "fb") return _runFB();
  if (type === "line") return _runLine();
  // 未知 type → 回空字串（避免 PAD 當錯誤）
  return ContentService.createTextOutput("").setMimeType(ContentService.MimeType.TEXT);
}

/** 原本的 Line 流程（來源：Line轉貼預排表 → 寫入 Line轉貼歷史） */
function _runLine() {
  const timezone = "Asia/Taipei";
  const sheetId = "1Uc-96Q1V4x1hI6VUOoKv4MC7AX-0KKbZGhlWEXbcVFM";
  const sheetName = "Line轉貼預排表";
  const historySheetName = "Line轉貼歷史";

  const ss = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheetByName(sheetName);
  const historySheet = ss.getSheetByName(historySheetName);

  const lastRow = sheet.getLastRow();
  const now = new Date();

  if (lastRow < 2) {
    return ContentService.createTextOutput("").setMimeType(ContentService.MimeType.TEXT);
  }

  const allData = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const dateValues = sheet.getRange(2, 2, lastRow - 1, 1).getValues();        // B
  const timeValues = sheet.getRange(2, 3, lastRow - 1, 1).getDisplayValues(); // C

  console.log("=== [_runLine] 開始執行 ===");
  console.log(`📌 現在時間: ${Utilities.formatDate(now, timezone, 'yyyy/MM/dd HH:mm:ss')}`);
  console.log(`📌 資料列數: ${allData.length}`);

  let taskList = [];

  for (let i = 0; i < allData.length; i++) {
    const row = allData[i];
    const status = row[9];                         // J 狀態
    const date = dateValues[i][0];                 // B 日期
    const time = timeValues[i][0];                 // C 時間（顯示值）
    const productID = row[4]?.toString().trim();   // E 商品編號
    const pushMessage = row[6]?.toString().trim(); // G 推播訊息

    if (!productID || !pushMessage) continue; // 必要欄位

    if (!status) {
      let sortKey;
      let shouldSend = false;

      if (!time) {
        sortKey = new Date(0); // 時間空白 → 優先
        shouldSend = true;
      } else if (date) {
        sortKey = new Date(`${Utilities.formatDate(date, timezone, "yyyy/MM/dd")} ${time}:00`);
        if (sortKey.getTime() <= now.getTime()) shouldSend = true;
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

  taskList.sort((a, b) => a.sortTime - b.sortTime); // 空白優先、時間早者優先

  const selected = taskList[0];
  const rowIndex = selected.index;
  const rowData = selected.row;
  const productID = rowData[4];
  const pushMessage = rowData[6];

  // 更新狀態為「已轉貼」
  sheet.getRange(rowIndex, 10).setValue("已轉貼");

  // 寫入歷史最上方（J 欄寫入時間戳）
  const rowWithTimestamp = [...rowData];
  rowWithTimestamp[9] = Utilities.formatDate(now, timezone, "yyyy/MM/dd HH:mm:ss");
  historySheet.insertRows(2, 1);
  historySheet.getRange(2, 1, 1, rowWithTimestamp.length).setValues([rowWithTimestamp]);

  // 刪除來源列
  sheet.deleteRow(rowIndex);

  const result = `${productID}|${pushMessage}`;
  console.log(`🎯 傳送任務（第 ${rowIndex} 列）: ${result}`);
  return ContentService.createTextOutput(result).setMimeType(ContentService.MimeType.TEXT);
}

/** FB 流程（來源：FB互動預排表 → 寫入 FB互動歷史；只抓「已過去且最近」的） */
/** FB 流程（來源：FB互動預排表 → 寫入 FB互動歷史）
 *  調整：時間為空 → 視為當天最優先任務
 */
function _runFB() {
  const timezone = "Asia/Taipei";
  const sheetId = "1Uc-96Q1V4x1hI6VUOoKv4MC7AX-0KKbZGhlWEXbcVFM";
  const sheetName = "FB互動預排表";
  const historySheetName = "FB互動歷史";

  const ss = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheetByName(sheetName);
  const historySheet = ss.getSheetByName(historySheetName);

  const lastRow = sheet.getLastRow();
  const now = new Date();

  if (lastRow < 2) {
    return ContentService.createTextOutput("").setMimeType(ContentService.MimeType.TEXT);
  }

  const allData  = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const dateVals = sheet.getRange(2, 2, lastRow - 1, 1).getValues();        // B 日期（實值）
  const timeDisp = sheet.getRange(2, 3, lastRow - 1, 1).getDisplayValues(); // C 時間（顯示值 HH:mm）

  console.log("=== [_runFB] 開始執行 ===");
  console.log(`📌 現在時間: ${Utilities.formatDate(now, timezone, 'yyyy/MM/dd HH:mm:ss')}`);
  console.log(`📌 資料列數: ${allData.length}`);

  // 欄位（A=0）
  const COL_STATUS = 9; // J 狀態（空白=未執行）
  const COL_FBLINK = 6; // G FB Link
  const COL_AIMSG  = 8; // I AI 生成訊息

  const todayStr = Utilities.formatDate(now, timezone, "yyyy/MM/dd");

  /** 三個候選池：
   *  1) emptyToday: 時間空白 + 日期是今天（或日期也空） → 最優先
   *  2) emptyOther: 時間空白 + 日期不是今天             → 次優先
   *  3) candidates: 時間有填，且「時間已過去」者         → 再來
   */
  let emptyToday = [];
  let emptyOther = [];
  let candidates = [];

  for (let i = 0; i < allData.length; i++) {
    const row    = allData[i];
    const status = row[COL_STATUS];
    const fbLink = (row[COL_FBLINK] || "").toString().trim();
    const aiMsg  = (row[COL_AIMSG]  || "").toString().trim();
    const dVal   = dateVals[i][0];         // Date 或 空
    const tStr   = (timeDisp[i][0] || "").toString().trim(); // "" 或 "HH:mm"

    if (status) continue;                // 只抓未執行
    if (!fbLink || !aiMsg) continue;     // 必要欄位

    // —— 時間空白：視為當天最優先（先放到 emptyToday / emptyOther）——
    if (!tStr) {
      // 有填日期就比對是否為今天；若沒填日期，也視為今天（更保守地優先）
      let isToday = false;
      if (dVal) {
        const datePart = Utilities.formatDate(new Date(dVal), timezone, "yyyy/MM/dd");
        isToday = (datePart === todayStr);
      } else {
        isToday = true; // 無日期也當作今天優先處理
      }

      const item = {
        index: i + 2,
        row,
        outText: `${fbLink}|${aiMsg}`
      };

      if (isToday) {
        emptyToday.push(item);
      } else {
        emptyOther.push(item);
      }
      continue;
    }

    // —— 時間有填：維持原邏輯（只取「已過去」者，取最接近現在的一筆）——
    if (!dVal) continue; // 有時間但沒日期 → 略過（避免拼不出有效時間）

    const datePart = Utilities.formatDate(new Date(dVal), timezone, "yyyy/MM/dd");
    const dt = new Date(`${datePart} ${tStr}:00`);
    if (isNaN(dt.getTime())) continue;

    if (dt.getTime() <= now.getTime()) {
      candidates.push({
        index: i + 2,  // 真實列號
        row,
        when: dt,
        outText: `${fbLink}|${aiMsg}`
      });
    }
  }

  // —— 挑選優先順序：emptyToday > emptyOther > candidates（最近過去）——
  let picked = null;

  if (emptyToday.length > 0) {
    picked = emptyToday[0]; // 同一天、時間空白 → 依表格順序先來先服務
    console.log(`✅ 選到「今天時間空白」任務：第 ${picked.index} 列`);
  } else if (emptyOther.length > 0) {
    picked = emptyOther[0]; // 其它日期、時間空白 → 次優先
    console.log(`✅ 選到「非今天時間空白」任務：第 ${picked.index} 列`);
  } else if (candidates.length > 0) {
    // 原本規則：取「已過去中最接近現在」的一筆（when 最大）
    candidates.sort((a, b) => b.when - a.when);
    picked = candidates[0];
    console.log(`✅ 選到「已過去且最近」任務：第 ${picked.index} 列，時間=${picked.when}`);
  }

  if (!picked) {
    console.log("⛔ 沒有符合條件的未執行任務");
    return ContentService.createTextOutput("").setMimeType(ContentService.MimeType.TEXT);
  }

  const rowIndex = picked.index;
  const rowData  = picked.row;

  // 更新來源狀態 → 已轉貼（J 欄）
  sheet.getRange(rowIndex, COL_STATUS + 1).setValue("已轉貼");

  // 寫入歷史最上方（J 欄寫入時間戳）
  const rowWithTimestamp = [...rowData];
  rowWithTimestamp[COL_STATUS] = Utilities.formatDate(now, timezone, "yyyy/MM/dd HH:mm:ss");
  historySheet.insertRows(2, 1);
  historySheet.getRange(2, 1, 1, rowWithTimestamp.length).setValues([rowWithTimestamp]);

  // 刪除來源列
  sheet.deleteRow(rowIndex);

  console.log(`🎯 傳送任務（第 ${rowIndex} 列）: ${picked.outText}`);
  return ContentService.createTextOutput(picked.outText).setMimeType(ContentService.MimeType.TEXT);
}

