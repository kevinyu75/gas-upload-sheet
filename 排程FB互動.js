/***************************************************************
 * 單檔完整版：上架待排 → FB互動預排表（修正版）
 * 規則：
 * - 來源：2024新版上架表，僅取 H=待排 且 L=TRUE
 * - 只在候選清單內依 customID 去重（保留最早 K+1h）
 * - Firestore 取 name/desc/facebookPostID/creationTime
 *   ⮕ 取不到或錯誤：回寫 H=「錯誤請檢查」，且不寫入目標
 *   ⮕ facebookPostID 缺失或無法組出連結：回寫 H=「錯誤請檢查」，且不寫入目標
 * - 成功寫入目標（插入第 2 列）：
 *   - F=name（Firestore），G=FB Link（由 facebookPostID 組），C=K+1h
 *   - I 欄 AI 朋友口吻文案（長度預設 20、避免「最近…」開頭、不強制品名）
 *   - H（來源）回寫：若 creationTime >1 個月 →「請檢查文建立時間」，否則「已排」
 * - 最後以 B↓、C↓ 重新排序（最新在最上）
 * 需要 Script Properties：
 *   openai_api_key（必填）、openai_model（選填）、
 *   firestore_email/firestore_key/firestore_projectId（必填）、
 *   fb_group_posts_base（選填，預設 https://www.facebook.com/groups/playbeautystreet/posts/）
 ***************************************************************/

/* ========================= 基礎工具 ========================= */
function parseDateCell_(v) {
  if (v instanceof Date) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  if (typeof v === 'number') {
    const ms = Math.round((v - 25569) * 86400 * 1000);
    const d  = new Date(ms);
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }
  const s = String(v || '').trim();
  if (!s) return null;
  let m = s.match(/^(\d{1,2})[\/\-月](\d{1,2})(?:日)?$/);
  if (m) return new Date(new Date().getFullYear(), Number(m[1]) - 1, Number(m[2]));
  m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  const d2 = new Date(s);
  return isNaN(d2) ? null : new Date(d2.getFullYear(), d2.getMonth(), d2.getDate());
}
function parseTimeCell_(v) {
  if (v instanceof Date) return new Date(1899, 11, 30, v.getHours(), v.getMinutes(), v.getSeconds());
  if (typeof v === 'number') {
    const dayMs = 24 * 3600 * 1000;
    const ms = Math.round((v % 1) * dayMs);
    const base = new Date(1899, 11, 30);
    return new Date(base.getTime() + ms);
  }
  let s = String(v || '').trim();
  if (!s) return null;
  s = s.replace('：', ':'); // 全形冒號→半形
  let m = s.match(/^(上午|下午)\s*(\d{1,2}):(\d{2})$/);
  if (m) {
    let h = Number(m[2]), mi = Number(m[3]);
    if (m[1] === '下午' && h < 12) h += 12;
    if (m[1] === '上午' && h === 12) h = 0;
    return new Date(1899, 11, 30, h, mi, 0);
  }
  m = s.match(/^(AM|PM)\s*(\d{1,2}):(\d{2})$/i);
  if (m) {
    let h = Number(m[2]), mi = Number(m[3]);
    if (/PM/i.test(m[1]) && h < 12) h += 12;
    if (/AM/i.test(m[1]) && h === 12) h = 0;
    return new Date(1899, 11, 30, h, mi, 0);
  }
  m = s.match(/^(\d{1,2}):(\d{2})$/);
  if (m) return new Date(1899, 11, 30, Number(m[1]), Number(m[2]), 0);
  m = s.match(/^(\d{1,2})\s*[時点]\s*(\d{1,2})\s*分?$/);
  if (m) return new Date(1899, 11, 30, Number(m[1]), Number(m[2]), 0);
  return null;
}
function fmtDateYYYYMMDD_(v){ const d=parseDateCell_(v); if(!d) return ''; const y=d.getFullYear(),m=('0'+(d.getMonth()+1)).slice(-2),day=('0'+d.getDate()).slice(-2); return `${y}-${m}-${day}`;}
function fmtTimeHHmm_(v){ const t=parseTimeCell_(v); if(!t) return ''; const h=('0'+t.getHours()).slice(-2),m=('0'+t.getMinutes()).slice(-2); return `${h}:${m}`;}
function combineDateTimeAddHours_(dateObj, timeObj, addHours){
  const d=parseDateCell_(dateObj), t=parseTimeCell_(timeObj);
  if(!d||!t) return null;
  const dt0=new Date(d.getFullYear(),d.getMonth(),d.getDate(),t.getHours(),t.getMinutes(),t.getSeconds());
  return new Date(dt0.getTime()+ (addHours||0)*3600*1000);
}
function clamp_(n,min,max){ return Math.max(min,Math.min(max,n)); }
function 計字_(s){ return Array.from(String(s||'')).length; }
function normalizeOneLine_(s){ return String(s||'').replace(/\s+/g,' ').replace(/[•▪︎·●○◦►▶︎\-\*]\s*/g,'').trim(); }
function 去價格字樣_(s){ return String(s||'').replace(/(\$|NT[.$\s]*|NTD[.\s]*|新台幣|台幣)\s*\d[\d,\.]*/gi,'').replace(/\s{2,}/g,' ').trim(); }
function 偵測開頭禁詞_(s){ const t=String(s||'').trim().replace(/^[“"『「（(【\[\s]+/,''); const m=t.match(/^(最近|近來|近期|這陣子|這幾天|近幾天|我最近|我這陣子|我發現|我找到|最近發現|最近找到|最近入手|剛入手|小編)/); return m?m[1]:''; }
function 萃取品名關鍵詞_(name,desc){
  let s=String(name||'').trim();
  if(!s) s=(String(desc||'').split(/\r?\n/)[0]||'').trim();
  if(!s) return '';
  s=s.replace(/（[^）]*）|\([^)]*\)|\[[^\]]*\]|【[^】]*】/g,'');
  s=s.replace(/\d+(\.\d+)?\s*(ml|mL|ML|g|kg|入|組|盒|瓶|包|片|支|顆|cm|公分|\$|元|NT)\b/gi,'');
  s=s.split(/[，,。:：;；|｜\/]/)[0].replace(/\s+/g,'');
  if (Array.from(s).length<2){
    const han=(String(name||desc||'').match(/[\p{Script=Han}]/gu)||[]).slice(0,6).join('');
    s=han||s;
  }
  return Array.from(s).slice(0,8).join('');
}

/* ========================= Firestore ========================= */
function 讀取商品_info_(customID) {
  const props     = PropertiesService.getScriptProperties();
  const email     = props.getProperty('firestore_email');
  const key       = props.getProperty('firestore_key');
  const projectId = props.getProperty('firestore_projectId');
  if (!email || !key || !projectId) throw new Error('缺少 firestore_email / firestore_key / firestore_projectId');

  const firestore = FirestoreApp.getFirestore(email, key.replace(/\\n/g, '\n'), projectId);

  let products = [];
  try { products = firestore.query('products').Where('customID', '==', String(customID)).Execute(); } catch (e) {
    throw new Error('Firestore query error: ' + e);
  }
  if (!products || products.length === 0) throw new Error('找不到商品：' + customID);

  const f = (products[0].fields || {});
  const readAny = (obj, k) => {
    const v = obj[k];
    if (v == null) return '';
    if (v.stringValue   !== undefined) return v.stringValue;
    if (v.integerValue  !== undefined) return String(v.integerValue);
    if (v.doubleValue   !== undefined) return String(v.doubleValue);
    if (v.timestampValue!== undefined) return v.timestampValue;
    return v;
  };

  const name = String(readAny(f,'name') || readAny(f,'title') || readAny(f,'商品標題') || '').trim();
  const desc = String(readAny(f,'desc') || readAny(f,'body')  || readAny(f,'商品內文') || '').trim();
  const facebookPostID = String(readAny(f,'facebookPostID') || '').trim();

  const ctRaw = String(readAny(f,'creationTime') || readAny(f,'CreationTime') || '').trim();
  let creationMs = 0;
  if (ctRaw) {
    if (/^\d{13}$/.test(ctRaw)) creationMs = Number(ctRaw);
    else if (/^\d{10}$/.test(ctRaw)) creationMs = Number(ctRaw) * 1000;
    else {
      const d = new Date(ctRaw);
      if (!isNaN(d)) creationMs = d.getTime();
    }
  }

  return { name, desc, facebookPostID, creationMs };
}
function 組FBPostLink_(facebookPostID) {
  const props = PropertiesService.getScriptProperties();
  const base  = props.getProperty('fb_group_posts_base') || 'https://www.facebook.com/groups/playbeautystreet/posts/';
  const s = String(facebookPostID || '').trim().replace(/"/g,'');
  if (!s) return '';
  const postId = s.split('_').pop().replace(/[^\d]/g,'');
  return postId ? `${base}${postId}/` : '';
}

/* ========================= OpenAI 生成 ========================= */
function chatCallWithRetry_({ apiKey, model, messages, maxTokens }) {
  const url = 'https://api.openai.com/v1/chat/completions';
  const payload = { model, temperature: 0.7, max_tokens: maxTokens, messages };
  const options = { method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: `Bearer ${apiKey}` }, payload: JSON.stringify(payload) };
  const MAX_RETRY = 3; let wait = 400;
  for (let t=0; t<=MAX_RETRY; t++){
    const resp = UrlFetchApp.fetch(url, options);
    const code = resp.getResponseCode();
    if (code === 200) return (JSON.parse(resp.getContentText())?.choices?.[0]?.message?.content || '').replace(/\r?\n/g,' ').trim();
    if ((code === 429 || code >= 500) && t < MAX_RETRY) { Utilities.sleep(wait); wait = Math.min(wait*2, 2500); continue; }
    throw new Error('OpenAI API error(' + code + '): ' + resp.getContentText());
  }
}
function 生成文案_OpenAI_約長度_朋友口吻(name, desc, tone, targetLen) {
  const props  = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('openai_api_key');
  if (!apiKey) throw new Error('缺少 openai_api_key');
  const model  = props.getProperty('openai_model') || 'gpt-4o-mini';

  const approx = Math.max(10, Number(targetLen) || 20);
  const minLen = Math.max(12, Math.round(approx * 0.85));
  const maxLen = Math.round(approx * 1.15);
  const productKeyword = 萃取品名關鍵詞_(name, desc);

  const system =
`你是中文社群文案助手。輸出單一段落（單段、不加標題），口吻自然像朋友分享，避免推銷語氣。可使用 1～3 個 emoji；如資訊提到價格，請避免在文案中重複價格。`;
  const user =
`請根據以下商品資訊，寫一段約 ${approx} 字的貼文文案（單段落），口吻：${tone}。
可自然帶到品名，但若會讓語句卡或重複，可以不寫品名；如要提及請以自然語境帶過。
【開頭規則】不得以「最近、近來、近期、這陣子、我最近、小編…」等時間詞或第一人稱作為開頭，請以「特點/效果/情境/結果」起手。

品名：${name || productKeyword || ''}
描述：${desc || '（無描述）'}

輸出要求：
- 僅單段文字，不可換行與列點
- 不要出現「本品」「商品」「請提供」「抱歉」等字眼
- 若提到價格、促銷字樣請移除
- 長度目標：約 ${approx} 字，建議落在 ${minLen}～${maxLen} 字
- 直接輸出文案，不要任何解釋`;

  let content = chatCallWithRetry_({ apiKey, model, messages: [
    { role: 'system', content: system },
    { role: 'user',   content: user }
  ], maxTokens: clamp_(Math.round(approx*3), 300, 2048) });

  content = 去價格字樣_(normalizeOneLine_(content));

  let len = 計字_(content);
  if (len < Math.round(approx*0.75)) {
    const feedback =
`你剛產生了約 ${len} 字，偏短。請擴寫到 ${minLen}～${maxLen} 字之間，
補上更具體的感受與情境（香氣/口感/使用時機/誰適合等），
保持單段、不列點、不口號式結尾，且不要以時間詞或第一人稱開頭。直接輸出最終文本。`;
    content = chatCallWithRetry_({ apiKey, model, messages: [
      { role: 'system', content: system },
      { role: 'user',   content: user },
      { role: 'assistant', content },
      { role: 'user',   content: feedback }
    ], maxTokens: clamp_(Math.round(approx*4), 400, 2048) });
    content = 去價格字樣_(normalizeOneLine_(content));
  }

  const bad = 偵測開頭禁詞_(content);
  if (bad) {
    const feedbackOpen =
`你剛以「${bad}」開場，違反開頭規則。請改以「特點/效果/情境/結果」為開頭，其餘內容保留，
單段、不列點、不口號式結尾，長度維持在 ${minLen}～${maxLen} 字。直接輸出最終文本。`;
    content = chatCallWithRetry_({ apiKey, model, messages: [
      { role: 'system', content: system },
      { role: 'user',   content: user },
      { role: 'assistant', content },
      { role: 'user',   content: feedbackOpen }
    ], maxTokens: clamp_(Math.round(approx*3), 300, 1024) });
    content = 去價格字樣_(normalizeOneLine_(content));
  }
  return content;
}

/* ========================= 主流程 ========================= */
function 將上架待排轉貼到_FB互動預排表() {
  const RUN_ID = 'RUN#' + new Date().toISOString().replace('T',' ').replace(/\..+/, '');
  const log = (...a)=>{ const s=`[${RUN_ID}] `+a.map(v=>String(v)).join(' '); Logger.log(s); console.log(s); };

  const SRC_NAME = '2024新版上架表';
  const DST_NAME = 'FB互動預排表';
  const DEFAULT_LEN  = 20;
  const DEFAULT_TONE = '親切、易讀、口語、像朋友一樣的說話、不要露出推銷感';
  const ONE_MONTH_MS = 31 * 24 * 3600 * 1000;

  const ss  = SpreadsheetApp.getActive();
  const shS = ss.getSheetByName(SRC_NAME);
  const shD = ss.getSheetByName(DST_NAME);
  if (!shS || !shD) throw new Error(`找不到工作表：${!shS ? SRC_NAME : DST_NAME}`);
  log('啟動', `來源=${SRC_NAME}`, `目標=${DST_NAME}`);

  const COL_S_PIC      = 1;   // A
  const COL_S_STATUS   = 8;   // H
  const COL_S_FB社團   = 12;  // L
  const COL_S_上架日   = 10;  // J
  const COL_S_時段     = 11;  // K
  const COL_S_customID = 16;  // P
  const COL_S_商品名稱 = 18;  // R
  const SRC_LAST_COL   = 18;

  function 標記同ID其餘列為已排_(customID, exceptRow, idRowsMap) {
    const rows = idRowsMap.get(customID) || [];
    for (const r of rows) {
      if (r === exceptRow) continue;
      try {
        const cur = String(shS.getRange(r, COL_S_STATUS).getValue() || '').trim();
        if (cur === '待排') shS.getRange(r, COL_S_STATUS).setValue('已排');
      } catch (_) {}
    }
  }

  const lastRowS = shS.getLastRow();
  if (lastRowS < 2) { log('來源無資料'); return; }
  const srcVals = shS.getRange(2, 1, lastRowS - 1, SRC_LAST_COL).getValues();

  const lastRowD = shD.getLastRow();
  const existingIdSet = new Set();
  const existingPairSet = new Set();
  if (lastRowD >= 2) {
    const existVals = shD.getRange(2, 1, lastRowD - 1, Math.max(shD.getLastColumn(), 6)).getValues();
    for (const r of existVals) {
      const id  = String(r[4] || '').trim();
      const d   = r[1];
      if (id) {
        existingIdSet.add(id);
        if (d instanceof Date) existingPairSet.add(id + '|' + fmtDateYYYYMMDD_(d));
      }
    }
  }

  const candidates = [];
  const idRowsMap = new Map();
  for (let i=0;i<srcVals.length;i++){
    const rowNum = 2 + i;
    const row = srcVals[i];
    const status = String(row[COL_S_STATUS - 1] || '').trim();
    const fbOK   = row[COL_S_FB社團 - 1] === true;
    if (status !== '待排' || !fbOK) continue;

    const customID = String(row[COL_S_customID - 1] || '').trim();
    if (!customID) continue;

    const dateObj  = parseDateCell_(row[COL_S_上架日 - 1]);
    const timeObj  = parseTimeCell_(row[COL_S_時段 - 1]);
    if (!dateObj) continue;

    const schedDT  = timeObj ? combineDateTimeAddHours_(dateObj, timeObj, 1) : null;
    const dateOnly = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate());

    candidates.push({ rowNum, customID, dateOnly, schedDT, pic: row[COL_S_PIC-1], title: row[COL_S_商品名稱-1] });

    if (!idRowsMap.has(customID)) idRowsMap.set(customID, []);
    idRowsMap.get(customID).push(rowNum);
  }

  const pickMap = new Map();
  for (const c of candidates){
    const e = pickMap.get(c.customID);
    if (!e) { pickMap.set(c.customID, c); continue; }
    const tsC = (c.schedDT ? c.schedDT.getTime() : -Infinity);
    const tsE = (e.schedDT ? e.schedDT.getTime() : -Infinity);
    if (tsC < tsE || (tsC === tsE && c.rowNum < e.rowNum)) pickMap.set(c.customID, c);
  }
  const picks = Array.from(pickMap.values());

  let written = 0, errors = 0;
  for (const p of picks){
    try {
      const dStr = fmtDateYYYYMMDD_(p.dateOnly);
      const pairKey = p.customID + '|' + dStr;

      if (existingIdSet.has(p.customID) || existingPairSet.has(pairKey)) {
        const rng = shS.getRange(p.rowNum, COL_S_STATUS);
        rng.setValue('已排');
        rng.setNote('已在目標表，於 ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'));
        標記同ID其餘列為已排_(p.customID, p.rowNum, idRowsMap);
        continue;
      }

      let info;
      try {
        info = 讀取商品_info_(p.customID);
      } catch (e) {
        shS.getRange(p.rowNum, COL_S_STATUS).setValue('錯誤請檢查');
        標記同ID其餘列為已排_(p.customID, p.rowNum, idRowsMap);
        errors++; continue;
      }

      const fbLink = 組FBPostLink_(info.facebookPostID || '');
      if (!info.facebookPostID || !fbLink) {
        shS.getRange(p.rowNum, COL_S_STATUS).setValue('錯誤請檢查');
        標記同ID其餘列為已排_(p.customID, p.rowNum, idRowsMap);
        errors++; continue;
      }

      const name = info.name || String(p.title || '').trim();
      const desc = info.desc || '';

      let text='';
      try { text = 生成文案_OpenAI_約長度_朋友口吻(name, desc, DEFAULT_TONE, DEFAULT_LEN); }
      catch (e) { text = `生成失敗：${String(e).slice(0,180)}`; }

      const weekday = '週' + '日一二三四五六'[p.dateOnly.getDay()];
      const out = new Array(11).fill('');
      out[0]  = p.pic;
      out[1]  = p.dateOnly;
      out[2]  = p.schedDT || '';
      out[3]  = weekday;
      out[4]  = p.customID;
      out[5]  = name;
      out[6]  = fbLink;
      out[7]  = false;
      out[8]  = text;
      out[9]  = '';
      out[10] = DEFAULT_LEN;

      shD.insertRowBefore(2);
      shD.getRange(2, 1, 1, out.length).setValues([out]);
      written++;

      existingIdSet.add(p.customID);
      existingPairSet.add(pairKey);

      let statusToWrite = '已排';
      if (info.creationMs > 0 && (Date.now() - info.creationMs) > ONE_MONTH_MS) {
        statusToWrite = '請檢查文建立時間';
      }
      shS.getRange(p.rowNum, COL_S_STATUS).setValue(statusToWrite);
      標記同ID其餘列為已排_(p.customID, p.rowNum, idRowsMap);

      Utilities.sleep(120);
    } catch (e) {
      shS.getRange(p.rowNum, COL_S_STATUS).setValue('錯誤請檢查');
      標記同ID其餘列為已排_(p.customID, p.rowNum, idRowsMap);
      errors++;
    }
  }

  const lastRowD2 = shD.getLastRow();
  if (lastRowD2 > 2) {
    const range = shD.getRange(2, 1, lastRowD2 - 1, shD.getLastColumn());
    range.sort([{ column: 2, ascending: false }, { column: 3, ascending: false }]);
  }

  SpreadsheetApp.flush();
  log('完成摘要', `寫入=${written}`, `錯誤=${errors}`);
}
