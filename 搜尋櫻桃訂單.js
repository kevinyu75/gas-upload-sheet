function 查詢櫻桃訂單_REST_Playbeautyshop() {
  const creds = getPlaybeautyshopCredentials();
  const projectId = creds.project_id;
  const clientEmail = creds.client_email;
  const privateKey = creds.private_key;

  const jwt = createJwt_(clientEmail, privateKey);
  const accessToken = getAccessToken_(jwt);

  const url = `https://firestore.googleapis.com/v1/projects/${projectId}/databases/(default)/documents:runQuery`;
  const query = {
    structuredQuery: {
      from: [{ collectionId: "ordersHistory" }],
      where: {
        fieldFilter: {
          field: { fieldPath: "productName" },
          op: "GREATER_THAN_OR_EQUAL",
          value: { stringValue: "櫻" }
        }
      },
      limit: 100
    }
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: `Bearer ${accessToken}`
    },
    payload: JSON.stringify(query)
  };

  const response = UrlFetchApp.fetch(url, options);
  const lines = response.getContentText().split('\n').filter(Boolean);
  const results = [];

  for (const line of lines) {
    // ⚠️ 嚴格過濾掉非有效 JSON
    if (!line.trim().startsWith('{')) continue;

    try {
      const obj = JSON.parse(line);
      const fields = obj.document?.fields;
      if (!fields) continue;

      const productName = fields.productName?.stringValue || '';
      if (!productName.includes("櫻桃")) continue;

      const creationTime = new Date(Number(fields.creationTime?.integerValue || '0'));
      const userCustomNumber = fields.userCustomNumber?.stringValue || '';
      const userID = fields.userID?.stringValue || '';
      const productCustomID = fields.productCustomID?.stringValue || '';

      results.push([
        creationTime,
        userCustomNumber,
        userID,
        productCustomID,
        productName
      ]);

      if (results.length >= 10) break;

    } catch (e) {
      Logger.log(`⚠️ 無法解析某筆資料，跳過：${e.message}`);
    }
  }

  if (results.length === 0) {
    SpreadsheetApp.getUi().alert("❗ 沒有找到包含『櫻桃』的訂單！");
    return;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("櫻桃訂單_REST");
  const headers = ["creationTime", "userCustomNumber", "userID", "productCustomID", "productName"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, results.length, headers.length).setValues(results);
  sheet.getRange(2, 1, results.length, 1).setNumberFormat("yyyy/mm/dd hh:mm:ss");

  SpreadsheetApp.getUi().alert(`✅ 成功匯出 ${results.length} 筆資料`);
}
