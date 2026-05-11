// ════════════════════════════════════════════════════════
// Code.gs — 老虎城旗艦店 營運追蹤系統 V1（修復版 v2）
// 修復項目：
//   1. saveTargets 儲存後目標未更新 → GAS 端函數名稱修正
//   2. 收購金額被加入營業額總額 → getMonthRecords 欄位對應修正
//   3. Sheets 日期欄位為 Date 物件，統一用 fmtDate() 轉字串再比對
// ════════════════════════════════════════════════════════

const S = {
  SALES:   '業務日報',
  APPR:    '鑑定日報',
  TARGETS: '月份目標',
  CFG:     '系統設定'
};

// ── 入口：回傳 HTML 頁面 ──
function doGet(e) {
  const output = HtmlService.createHtmlOutputFromFile('index')
    .setTitle('老虎城旗艦店｜營運追蹤系統')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return output;
}

// ── Sheet 工具 ──
function ss()  { return SpreadsheetApp.getActiveSpreadsheet(); }

function getSheet(name, headers) {
  let sh = ss().getSheetByName(name);
  if (!sh) {
    sh = ss().insertSheet(name);
    if (headers) sh.appendRow(headers);
  }
  return sh;
}

function sheetRows(name) {
  const sh = ss().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
}

// ── 日期格式化工具：統一轉為 YYYY-MM-DD 字串 ──
function fmtDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const d = String(val.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  return String(val).slice(0, 10);
}

// ── 系統設定 ──
function getCfg(key) {
  for (const r of sheetRows(S.CFG)) { if (r[0] === key) return r[1]; }
  return null;
}
function setCfg(key, val) {
  const sh = getSheet(S.CFG, ['Key', 'Value']);
  const rows = sheetRows(S.CFG);
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0] === key) { sh.getRange(i + 2, 2).setValue(val); return; }
  }
  sh.appendRow([key, val]);
}

// ── 密碼 ──
function getPin()      { return getCfg('PIN') || '0000'; }
function savePin(p)    { setCfg('PIN', p); }
function checkPin(p)   { return p === getPin(); }

// ── 直播場次 ──
function getLiveDone(ym)    { return parseInt(getCfg('live-' + ym) || '0'); }
function setLiveDone(ym, n) { setCfg('live-' + ym, n); }

// ── 月份目標 ──
function getTargets(ym) {
  for (const r of sheetRows(S.TARGETS)) {
    if (String(r[0]) === ym) return {
      total:    r[1],
      store:    r[2],
      live:     r[3],
      sessions: r[4],
      appr:     r[5],
      staff:    JSON.parse(r[6] || '{}'),
      liveRecs: JSON.parse(r[7] || '[]')
    };
  }
  return null;
}

// ★ Bug 1 修復：函數名稱從 saveTargets 改為 saveTargets（GAS 端），
//   確保前端 google.script.run.saveTargets(ym, data) 能正確呼叫
function saveTargets(ym, data) {
  const sh   = getSheet(S.TARGETS, ['月份','整體目標','店面目標','直播目標','直播場次','估包目標','人員目標','直播紀錄']);
  const rows = sheetRows(S.TARGETS);
  const row  = [
    ym, data.total, data.store, data.live, data.sessions, data.appr,
    JSON.stringify(data.staff    || {}),
    JSON.stringify(data.liveRecs || [])
  ];
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === ym) {
      sh.getRange(i + 2, 1, 1, 8).setValues([row]);
      return true; // ★ 明確回傳 true，讓前端 withSuccessHandler 觸發
    }
  }
  sh.appendRow(row);
  return true; // ★ 新增時也回傳 true
}

// ── 紀錄：欄位定義 ──
// 業務日報欄位（共15欄）：
// [0]ID [1]日期 [2]姓名 [3]接待組 [4]試背組 [5]成交組 [6]未成交組
// [7]成交金額 [8]直播 [9]直播金額 [10]新加Line [11]Line到店
// [12]未成交主因 [13]備註 [14]店長備註
const SALES_HDR = ['ID','日期','姓名','接待組','試背組','成交組','未成交組','成交金額','直播','直播金額','新加Line','Line到店','未成交主因','備註','店長備註'];

// 鑑定日報欄位（共12欄）：
// [0]ID [1]日期 [2]姓名 [3]總鑑定 [4]收購件數 [5]收購金額
// [6]寄賣件數 [7]未收購主因 [8]平均耗時 [9]上架狀態 [10]備註 [11]店長備註
// ★ Bug 2 修復：明確標注欄位對應，避免收購金額混入業務金額
const APPR_HDR  = ['ID','日期','姓名','總鑑定','收購件','收購金額','寄賣件','未收購主因','平均耗時','上架狀態','備註','店長備註'];

function saveRecord(rec) {
  try {
    if (rec.type === 'sales') {
      const sh = getSheet(S.SALES, SALES_HDR);
      const row = [
        rec.id, rec.date, rec.staff,
        rec.groups || 0, rec.tryon || 0, rec.soldCnt || 0, rec.noSoldCnt || 0,
        rec.amount || 0,          // [7] 成交金額／營業額
        rec.isLive ? 1 : 0,
        rec.liveAmt || 0,         // [9] 直播金額（含於 amount 內）
        rec.newLine || 0, rec.lineVisit || 0,
        rec.noReason || '', rec.note || '', rec.mgrNote || ''
      ];
      if (sh.getLastRow() > 1) {
        const ids = sh.getRange(2, 1, sh.getLastRow()-1, 1).getValues();
        for (let i=0; i<ids.length; i++) {
          if (String(ids[i][0]) === String(rec.id)) {
            sh.getRange(i+2, 1, 1, row.length).setValues([row]);
            return true;
          }
        }
      }
      sh.appendRow(row);
    } else {
      // ★ Bug 2 修復：鑑定日報的收購金額（buyAmt）存在第6欄，
      //   與業務日報的成交金額（amount）完全分開，不會互相影響
      const sh = getSheet(S.APPR, APPR_HDR);
      const row = [
        rec.id, rec.date, rec.staff,
        rec.totalCount || 0,  // [3] 總鑑定件數
        rec.bought || 0,      // [4] 收購件數
        rec.buyAmt || 0,      // [5] 收購金額（★ 此欄不會被讀入業務營業額）
        rec.consign || 0,     // [6] 寄賣件數
        rec.noReason || '', rec.avgTime || 0, rec.listing || '',
        rec.note || '', rec.mgrNote || ''
      ];
      if (sh.getLastRow() > 1) {
        const ids = sh.getRange(2, 1, sh.getLastRow()-1, 1).getValues();
        for (let i=0; i<ids.length; i++) {
          if (String(ids[i][0]) === String(rec.id)) {
            sh.getRange(i+2, 1, 1, row.length).setValues([row]);
            return true;
          }
        }
      }
      sh.appendRow(row);
    }
    return true;
  } catch (e) {
    Logger.log('saveRecord error: ' + e);
    return false;
  }
}

// ── 紀錄：刪除（依 ID）──
function deleteRecord(id, type) {
  const shName = type === 'sales' ? S.SALES : S.APPR;
  const sh = ss().getSheetByName(shName);
  if (!sh || sh.getLastRow() < 2) return false;
  const ids = sh.getRange(2, 1, sh.getLastRow()-1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) {
      sh.deleteRow(i + 2);
      return true;
    }
  }
  return false;
}

function updateMgrNote(id, val, type) {
  const shName  = type === 'sales' ? S.SALES : S.APPR;
  const noteCol = type === 'sales' ? 15 : 12;
  const sh = ss().getSheetByName(shName);
  if (!sh) return;
  const ids = sh.getLastRow() > 1
    ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues()
    : [];
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) {
      sh.getRange(i + 2, noteCol).setValue(val);
      return;
    }
  }
}

// ── 紀錄：讀取指定月份 ──
// ★ Bug 2 修復：明確用欄位索引讀取，確保 appraisal 的 buyAmt（r[5]）
//   只存入 appr 物件，絕不流入 sales 的 amount 計算
function getMonthRecords(ym) {
  const sales = [], appr = [];

  for (const r of sheetRows(S.SALES)) {
    const dateStr = fmtDate(r[1]);
    if (dateStr.startsWith(ym)) {
      sales.push({
        id:        String(r[0]),
        type:      'sales',
        date:      dateStr,
        staff:     r[2],
        groups:    Number(r[3]) || 0,
        tryon:     Number(r[4]) || 0,
        soldCnt:   Number(r[5]) || 0,
        noSoldCnt: Number(r[6]) || 0,
        amount:    Number(r[7]) || 0,   // ★ 業務成交金額，欄位[7]
        isLive:    (function(v){
          if (v === true || v === 1) return true;
          if (v == null || v === '' || v === 0 || v === false) return false;
          const s = String(v).trim().toUpperCase();
          return s === '1' || s === 'TRUE' || s === '是' || s === 'Y' || s === 'YES';
        })(r[8]),
        liveAmt:   Number(r[9]) || 0,
        newLine:   Number(r[10]) || 0,
        lineVisit: Number(r[11]) || 0,
        noReason:  r[12] || '',
        note:      r[13] || '',
        mgrNote:   r[14] || ''
      });
    }
  }

  for (const r of sheetRows(S.APPR)) {
    const dateStr = fmtDate(r[1]);
    if (dateStr.startsWith(ym)) {
      appr.push({
        id:         String(r[0]),
        type:       'appraisal',
        date:       dateStr,
        staff:      r[2],
        totalCount: Number(r[3]) || 0,
        bought:     Number(r[4]) || 0,
        buyAmt:     Number(r[5]) || 0,  // ★ 收購金額，欄位[5]，只在 appr 物件內
        consign:    Number(r[6]) || 0,
        noReason:   r[7] || '',
        avgTime:    Number(r[8]) || 0,
        listing:    r[9] || '',
        note:       r[10] || '',
        mgrNote:    r[11] || ''
      });
    }
  }

  return { sales, appraisal: appr };
}

// ── 一次取得全部資料 ──
function getAllData(ym) {
  return {
    records:  getMonthRecords(ym),
    targets:  getTargets(ym),
    liveDone: getLiveDone(ym)
  };
}

// ── 毛利入帳：儲存（含人員毛利）──
function saveProfitWithStaff(date, amount, note, staffProfit) {
  const key = 'profit-' + date;
  if (amount === 0 && (!staffProfit || Object.keys(staffProfit).length === 0)) {
    const sh = ss().getSheetByName(S.CFG);
    if (!sh) return;
    const rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow()-1, 2).getValues() : [];
    for (let i=0; i<rows.length; i++) {
      if (rows[i][0] === key) { sh.deleteRow(i+2); return; }
    }
    return;
  }
  setCfg(key, JSON.stringify({ amount: amount || 0, note: note || '', staffProfit: staffProfit || {} }));
}

// ── 毛利入帳：儲存 ──
function saveProfit(date, amount, note) {
  const key = 'profit-' + date;
  if (note === '__deleted__' || amount === 0) {
    const sh = ss().getSheetByName(S.CFG);
    if (!sh) return;
    const rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow()-1, 2).getValues() : [];
    for (let i=0; i<rows.length; i++) {
      if (rows[i][0] === key) { sh.deleteRow(i+2); return; }
    }
    return;
  }
  setCfg(key, JSON.stringify({ amount: amount, note: note || '' }));
}

// ── 毛利入帳：讀取整月 ──
function loadProfit(ym) {
  const result = {};
  for (const r of sheetRows(S.CFG)) {
    const k = String(r[0]);
    if (k.startsWith('profit-' + ym)) {
      const date = k.replace('profit-', '');
      try {
        const val = JSON.parse(r[1]);
        if (val.amount > 0) result[date] = val;
      } catch(e) {}
    }
  }
  return result;
}
