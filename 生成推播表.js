function scheduledPush() {
 const sourceSheetId = '1Uc-96Q1V4x1hI6VUOoKv4MC7AX-0KKbZGhlWEXbcVFM';
 const targetSheetId = '1Uc-96Q1V4x1hI6VUOoKv4MC7AX-0KKbZGhlWEXbcVFM';

 const sourceSpreadsheet = SpreadsheetApp.openById(sourceSheetId);
 const targetSpreadsheet = SpreadsheetApp.openById(targetSheetId);

 const sourceSheet = sourceSpreadsheet.getSheetByName('APP推播預排表');
 const targetSheet = targetSpreadsheet.getSheetByName('APP 推播表');

 // 設定時區並取得當前時間
 const timezone = "Asia/Taipei";
 const currentTime = new Date();
 
 // 設定時間格式範圍
 const timeRange = sourceSheet.getRange(2, 3, sourceSheet.getLastRow() - 1, 1);
 // 取得時間的原始顯示格式
 const timeDisplayValues = timeRange.getDisplayValues();

 const lastRow = sourceSheet.getLastRow();
 const dataRange = sourceSheet.getRange(2, 1, lastRow - 1, sourceSheet.getLastColumn());
 const data = dataRange.getValues();

 const filteredData = data.filter((row, index) => {
   if (!row[1] || !timeDisplayValues[index][0]) {
     console.log('跳過：日期或時間為空', row);
     return false;
   }

   try {
     const dateValue = row[1]; // B列的日期
     const timeString = timeDisplayValues[index][0]; // 使用顯示值的時間字串
     console.log('原始時間字串：', timeString);

     // 組合日期時間字串
     const dateString = Utilities.formatDate(dateValue, timezone, 'yyyy/MM/dd');
     const combinedDateTimeString = `${dateString} ${timeString}:00`;
     console.log('組合後的日期時間字串：', combinedDateTimeString);

     // 創建日期對象
     const pushDateTime = new Date(combinedDateTimeString);
     console.log('轉換後的日期時間對象：', pushDateTime);

     const processed = row[9]; // J列是已排程
     const customID = row[4]; // E列是商品編號
     const body = row[7]; // H列是Body

     return pushDateTime > currentTime && 
            processed !== "已排程" && 
            customID?.toString().trim() !== "" && 
            body?.toString().trim() !== "";
   } catch (error) {
     console.error('處理資料時發生錯誤：', error, row);
     return false;
   }
 });

 console.log(`找到 ${filteredData.length} 筆符合條件的資料`);

 const targetLastRow = targetSheet.getLastRow();
 let targetRowIndex = targetLastRow + 1;

 for (let i = 0; i < filteredData.length; i++) {
   const row = filteredData[i];
   
   try {
     const dateValue = row[1];
     const timeString = timeDisplayValues[data.indexOf(row)][0];

     // 組合日期時間字串
     const dateString = Utilities.formatDate(dateValue, timezone, 'yyyy/MM/dd');
     const combinedDateTimeString = `${dateString} ${timeString}:00`;

     const customID = row[4]; // E列是商品編號
     const title = row[6]; // G列是Title
     const body = row[7]; // H列是Body
     const imageUrl = row[8]; // I列是Image URL

     console.log(`處理第 ${i + 1} 筆資料：`);
     console.log(`最終推播時間: ${combinedDateTimeString}`);
     console.log(`商品編號: ${customID}`);

     // 呼叫取得商品DocumentId函數獲取文檔ID
     const documentId = 取得商品DocumentId(customID);
     
     if (!documentId) {
       console.error(`找不到商品 ${customID} 的DocumentId`);
       continue;  // 如果找不到DocumentId，跳過這筆資料
     }

     // 設置目標工作表的值
     targetSheet.getRange(targetRowIndex, 4).setValue(combinedDateTimeString);
     // 使用documentId構建新的URL
     targetSheet.getRange(targetRowIndex, 6).setValue(`https://playbeautyshop.com/product/${documentId}?ch=ap`);
     targetSheet.getRange(targetRowIndex, 1).setValue(title);
     targetSheet.getRange(targetRowIndex, 2).setValue(body);
     targetSheet.getRange(targetRowIndex, 5).setValue(imageUrl);
     targetSheet.getRange(targetRowIndex, 3).setValue('all');
     targetSheet.getRange(targetRowIndex, 7).setValue('high');
     targetSheet.getRange(targetRowIndex, 8).setValue('push');

     // 標記為已處理
     const sourceRowIndex = data.indexOf(row) + 2;
     sourceSheet.getRange(sourceRowIndex, 10).setValue("已排程");

     targetRowIndex++;
   } catch (error) {
     console.error(`處理資料時發生錯誤：${error}`);
     console.error(`問題資料：${JSON.stringify(row)}`);
     continue;
   }
 }

 // 排序目標工作表
 const targetDataRange = targetSheet.getDataRange();
 const targetData = targetDataRange.getValues();
 const header = targetData.shift();

 targetData.sort((a, b) => {
   try {
     return new Date(a[3]) - new Date(b[3]);
   } catch (error) {
     console.error(`排序時發生錯誤：${error}`);
     return 0;
   }
 });
 
 targetData.unshift(header);
 targetSheet.getRange(1, 1, targetData.length, targetData[0].length).setValues(targetData);

 // 執行推播
 sendNotificationFromSheet();
}

// 測試函數
function testTimeFormat() {
 const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('APP推播預排表');
 
 // 取得時間欄位的顯示值
 const timeRange = sourceSheet.getRange(2, 3, 1, 1);
 const timeDisplayValue = timeRange.getDisplayValue();
 console.log('時間欄位顯示值：', timeDisplayValue);
 
 // 取得日期欄位
 const dateRange = sourceSheet.getRange(2, 2, 1, 1);
 const dateValue = dateRange.getValue();
 
 // 組合日期時間字串
 const dateString = Utilities.formatDate(dateValue, "Asia/Taipei", 'yyyy/MM/dd');
 const combinedDateTimeString = `${dateString} ${timeDisplayValue}:00`;
 console.log('組合後的日期時間字串：', combinedDateTimeString);
 
 // 測試轉換
 const finalDate = new Date(combinedDateTimeString);
 console.log('最終日期時間對象：', finalDate);
}