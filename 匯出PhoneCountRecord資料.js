function 匯出PhoneCountRecord資料() {
  const ui = SpreadsheetApp.getUi();
  const customID = ui.prompt("請輸入要查詢的 customID").getResponseText();

  if (!customID) {
    ui.alert("未輸入 customID，動作取消");
    return;
  }

  const email = PropertiesService.getScriptProperties().getProperty('firestore_email');
  const key = PropertiesService.getScriptProperties().getProperty('firestore_key');
  const projectId = PropertiesService.getScriptProperties().getProperty('firestore_projectId');

  if (!email || !key || !projectId) {
    ui.alert("Firestore 連線資訊不完整！");
    return;
  }

  const formattedKey = key.replace(/\\n/g, '\n');
  const firestore = FirestoreApp.getFirestore(email, formattedKey, projectId);

  try {
    const products = firestore.query('products')
      .Where('customID', '==', customID)
      .Execute();

    if (products.length === 0) {
      ui.alert(`找不到 customID 為 "${customID}" 的商品`);
      return;
    }

    const doc = products[0];
    const raw = JSON.parse(JSON.stringify(doc));
    const recordRawArray = raw.fields?.phoneCountRecord?.arrayValue?.values;

    if (!recordRawArray || recordRawArray.length === 0) {
      ui.alert(`customID 為 "${customID}" 的商品，phoneCountRecord 資料不存在或為空`);
      return;
    }

    // ✅ 使用原始電話字串，不加引號，讓 Google Sheets 自動處理格式（首碼 0 會被省略）
    const values = recordRawArray.map(entry => {
      const fields = entry.mapValue.fields;
      const phone = fields.phone?.stringValue || '';
      const count = parseInt(fields.count?.integerValue || '0', 10);
      return [phone, count]; // ← 不加引號
    });

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const startCell = sheet.getActiveCell();
    const startRow = startCell.getRow();
    const startCol = startCell.getColumn();

    // 插入表頭
    sheet.getRange(startRow, startCol, 1, 2).setValues([['電話', '次數']]);

    // 寫入資料
    sheet.getRange(startRow + 1, startCol, values.length, 2).setValues(values);

    ui.alert(`✅ 成功輸出 ${values.length} 筆資料（含表頭），電話開頭 0 會被自動省略。`);

  } catch (error) {
    Logger.log(error);
    ui.alert(`查詢過程中發生錯誤：${error}`);
  }
}
