/**************************************************************
 * 使用者選取任一格 → Firestore 抓該列商品 → 生成文案 → G 欄
 * 欄位：E=商品編號(customID), I=字數, G=AI生成文案, L=錯誤
 **************************************************************/
function 生成Line推播文案_ActiveRow1() {
  const sh = SpreadsheetApp.getActiveSheet();
  const cell = sh.getActiveCell();
  if (!cell) {
    SpreadsheetApp.getUi().alert("請選取商品所在的列再執行");
    return;
  }

  const row = cell.getRow();
  if (row < 2) {
    SpreadsheetApp.getUi().alert("請選取資料列（第2列以後）");
    return;
  }

  const COL_CODE = 5;   // E 商品編號
  const COL_OUT  = 7;   // G AI 生成文案
  const COL_LEN  = 9;   // I 字數
  const COL_ERR  = 12;  // L 錯誤

  // 清空舊資料
  sh.getRange(row, COL_OUT).setValue('');
  sh.getRange(row, COL_ERR).setValue('');

  try {
    const customID = String(sh.getRange(row, COL_CODE).getValue() || '').trim();
    if (!customID) {
      sh.getRange(row, COL_ERR).setValue('缺少商品編號');
      return;
    }

    const targetRaw = sh.getRange(row, COL_LEN).getValue();
    const approxLen = 解析字數_(targetRaw, 20); // I 空白→20

    // 讀 Firestore
    const { name, desc } = 讀取商品資訊_(customID);

    // 生成 Line 推播文案
    let text = 生成Line推播文案_Gemini(name, desc, approxLen);
    text = 加入智能Emoji_(text, name, desc); // 第1行 emoji
    text = 清理文案_(text);                   // 去掉物流字眼
    text = 隨機變化文案_(text);               // 第二、三行隨機化

    sh.getRange(row, COL_OUT).setValue(text);

  } catch (e) {
    sh.getRange(row, COL_ERR).setValue(String(e).slice(0, 200));
  }
}

/****************************************
 * Firestore：以 customID 找商品資料
 ****************************************/
function 讀取商品資訊_(customID) {
  const props     = PropertiesService.getScriptProperties();
  const email     = props.getProperty('firestore_email');
  const key       = props.getProperty('firestore_key');
  const projectId = props.getProperty('firestore_projectId');
  if (!email || !key || !projectId) throw new Error('缺少 Firestore 連線設定');

  const firestore = FirestoreApp.getFirestore(email, key.replace(/\\n/g, '\n'), projectId);
  let products = [];

  try {
    products = firestore.query('products').Where('customID', '==', String(customID)).Execute();
  } catch (_) {}

  if (!products || products.length === 0) throw new Error('找不到商品：' + customID);

  const doc = products[0];
  const f   = doc.fields || {};

  const readStr = (obj, k) =>
    (obj[k] && obj[k].stringValue !== undefined) ? obj[k].stringValue : (obj[k] ?? '');

  const name = String(readStr(f, 'name') || readStr(f, 'title') || '').trim();
  const desc = String(readStr(f, 'desc') || readStr(f, 'body')  || '').trim();

  if (!name && !desc) throw new Error('商品名稱與內文皆為空');
  return { name, desc };
}

/******************************************************************
 * Gemini：生成適合 Line 社群的推播文案（範例式 prompt + 字數控制）
 ******************************************************************/
function 生成Line推播文案_Gemini(name, desc, approxLen) {
  const props  = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('gemini_api_key');
  if (!apiKey) throw new Error('缺少 gemini_api_key');

  // 使用者指定測試模型：Gemini 3.0 Pro Preview
  // 注意：此模型 ID 可能尚未對外開放，若執行失敗請改回 gemini-1.5-pro 或 gemini-2.0-flash-exp
  const model  = 'gemini-3.0-pro-preview';

  const minLen = Math.max(10, Math.round(approxLen * 0.8));
  const maxLen = Math.round(approxLen * 1.2);

  const prompt = `
你是中文行銷文案助手，專門生成適合 Line 社群推播的文字。

請根據以下商品資訊，生成一則文案，長度大約 ${approxLen} 字（建議落在 ${minLen}-${maxLen} 字）。
- 第一行：品名（不要加 emoji，由程式端處理）
- 第二～三行：主要特色/使用情境，用 ✔️/👉/🔥 等符號
- 最後一行：價格，例如「💟團購價: $150/包」
- 請避免輸出效期、重量、保存方式、出貨週期等物流資訊
- 保留品牌、產地、特色（例：正港台灣豬肉、日本抹茶）
- 每次用詞稍微不同，不要每則文案都用相同句型

範例：
榛果可可醬夾心餅
✔️ 滿滿濃郁巧克力榛果醬
👉 酥脆外層一口爆餡
💟 團購價: $63/包

商品名稱：${name}
商品內文：${desc}
`;

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

  const payload = {
    contents: [{
      parts: [{ text: prompt }]
    }],
    generationConfig: {
      maxOutputTokens: 512,
      temperature: 0.7
    }
  };

  const resp = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    payload: JSON.stringify(payload)
  });

  if (resp.getResponseCode() !== 200) {
    throw new Error('Gemini API error: ' + resp.getContentText());
  }

  const parsed = JSON.parse(resp.getContentText());
  let txt = parsed?.candidates?.[0]?.content?.parts?.[0]?.text || '';
  if (!txt || txt.trim() === '') txt = '[生成失敗，請重試]';

  return txt.trim();
}

/****************************************
 * 後處理：移除物流/規格相關字眼
 ****************************************/
function 清理文案_(txt) {
  return String(txt)
    .replace(/(保存方式|冷藏保存|冷凍保存|效期|有效期限|到貨週期|[0-9]+g|[0-9]+ml|約[0-9]+[包盒入])/gi, '')
    .replace(/\s*\n\s*\n/g, '\n')
    .trim();
}

/****************************************
 * 智能 Emoji 插入（第一行）
 ****************************************/
function 加入智能Emoji_(txt, name, desc) {
  const keywords = (name + ' ' + desc).toLowerCase();
  let candidates = [];

  const map = [
    { re: /(洗碗|清潔|洗滌|dish|detergent)/i, e: ['🧼','🫧','🧽'] },
    { re: /(蝦|shrimp)/i, e: ['🦐','🍤'] },
    { re: /(魚|鮭|鯛|tuna|salmon|fish)/i, e: ['🐟','🍣'] },
    { re: /(牛|牛肉|beef)/i, e: ['🥩','🍖'] },
    { re: /(豬|pork)/i, e: ['🐷','🥓'] },
    { re: /(雞|chicken)/i, e: ['🍗','🐔'] },
    { re: /(蛋|egg)/i, e: ['🥚','🍳'] },
    { re: /(麵|noodle|拉麵)/i, e: ['🍜','🍝'] },
    { re: /(饅頭|包子|饅)/i, e: ['🥟','🍞','🥠'] },
    { re: /(餅|餅乾|cookie)/i, e: ['🍪','🥠'] },
    { re: /(蛋糕|cake)/i, e: ['🍰','🧁'] },
    { re: /(巧克力|choco)/i, e: ['🍫','🥮'] },
    { re: /(咖啡|coffee)/i, e: ['☕','🫘'] },
    { re: /(茶|tea)/i, e: ['🍵','🫖'] },
    { re: /(水果|果汁|apple|banana|orange|berry)/i, e: ['🍎','🍊','🍇','🍓'] },
    { re: /(冰|ice|雪糕|冰淇淋|ice cream)/i, e: ['🍨','🍧','🍦'] }
  ];

  for (let m of map) {
    if (m.re.test(keywords)) {
      candidates = m.e;
      break;
    }
  }

  if (candidates.length === 0) {
    candidates = ['✨','🌟','🎉','🔥','🍀','💎','💡','🎊','🌈','⭐','🥳','😋','👌'];
  }

  const emoji = candidates[Math.floor(Math.random() * candidates.length)];

  let lines = txt.split('\n');
  if (lines.length > 0) {
    lines[0] = lines[0].replace(/^[\u{1F300}-\u{1FAFF}]\s*/u, ''); // 移除 AI 自己加的 emoji
    lines[0] = emoji + ' ' + lines[0];
  }

  return lines.join('\n').trim();
}

/****************************************
 * 後處理：第二、三行隨機符號＋口氣變化
 ****************************************/
function 隨機變化文案_(txt) {
  let lines = txt.split('\n').map(l => l.trim()).filter(l => l);

  if (lines.length >= 3) {
    const emojiPool = ['✔️','👉','🔥','⭐','🍴','💡','😋','🌟','👌','🎯'];
    const suffixPool = ['喔','啦','呢','唷','～','！'];

    // 第二行處理
    if (lines[1]) {
      const e1 = emojiPool[Math.floor(Math.random() * emojiPool.length)];
      const s1 = suffixPool[Math.floor(Math.random() * suffixPool.length)];
      lines[1] = e1 + ' ' + lines[1].replace(/^[^ ]+/, '').trim() + (Math.random() < 0.4 ? s1 : '');
    }

    // 第三行處理
    if (lines[2]) {
      const e2 = emojiPool[Math.floor(Math.random() * emojiPool.length)];
      const s2 = suffixPool[Math.floor(Math.random() * suffixPool.length)];
      lines[2] = e2 + ' ' + lines[2].replace(/^[^ ]+/, '').trim() + (Math.random() < 0.4 ? s2 : '');
    }
  }

  return lines.join('\n');
}

/****************************************
 * 解析 I 欄字數（空白→ fallback）
 ****************************************/
function 解析字數_(val, fallback) {
  if (val == null || String(val).trim() === '') return fallback;
  let s = String(val).trim().replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0));
  s = s.replace(/[,，\s]/g, '').replace(/[^\d]/g, '');
  const n = Number(s);
  return Number.isFinite(n) && n > 0 ? n : fallback;
}
