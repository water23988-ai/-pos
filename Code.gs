// ══════════════════════════════════════════════════════
//  拾逅花事 | 進銷存系統 — Google Apps Script 後端
//  版本：2.4  |  2026-04
//  更新：新增進貨模組（ProcurementBatches / ProcurementItems）
//        成本回寫改為加權平均邏輯
//  [v2.4 新增]
//  - 新增 getHistory action（sales.html 銷售紀錄頁面查詢）
//  - 新增 adjustStock action（inventory.html 庫存 +/- 快速調整）
//  - Products 新增 visible 欄位（商品管理開關：收銀台顯示/隱藏）
//  - updateProduct 支援寫入 visible 欄位
//  - getProducts 回傳 visible 欄位
//  [v2.3 修復]
//  - Bug Fix: deleteProduct 改為軟刪除（status='inactive'），歷史報表成本不斷鏈
//  - 新增 reactivateProduct：可將停用商品重新啟用
//  - Products 工作表新增 status 欄位（active / inactive）
//  [v2.2 修復]
//  - Bug Fix: Studio 交易獨立記錄，不再合併入高雄
//  - Bug Fix: 報廢 action 正確寫入 type='waste'，日結廢棄成本恢復正常
//  - Bug Fix: 進貨批次記錄補上 store 欄位
//  - Bug Fix: 庫存記錄查詢不再混入無點位記錄
//  - 效能優化: getSheet() 加入快取，避免重複開啟試算表
// ══════════════════════════════════════════════════════
//
//  【使用說明】
//  1. 在 Google Sheets 建立一份新試算表
//  2. 點選「延伸功能」>「Apps Script」，將此檔案內容貼上
//  3. 修改下方 SPREADSHEET_ID 為你的試算表 ID
//     （試算表網址中 /d/ 之後、/edit 之前的那段）
//  4. 第一次部署前，先執行 setupSheets() 初始化工作表
//  5. 部署為「網頁應用程式」：
//       執行身分 → 我自己
//       存取權限 → 所有人（匿名）
//  6. 複製部署 URL，貼入前端 API 設定欄位

const SPREADSHEET_ID = '1BQjkFe1pl7OYxvnZLoItVWZjYFnqBZalbwJPLPKMw4w';

// ── 工作表名稱 ──
const SH_PROC = {
  BATCHES : 'ProcurementBatches',
  ITEMS   : 'ProcurementItems',
};

const SH = {
  PRODUCTS    : 'Products',
  TRANSACTIONS: 'Transactions',
  MEMBERS     : 'Members',
  INV_LOG     : 'InventoryLog',
  HANG_ORDERS : 'HangOrders',
  PRICE_LOG   : 'PriceHistory',
};

// ══════════════════════════════════════════════════════
//  初始化（首次部署前執行一次）
// ══════════════════════════════════════════════════════
function setupSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  function ensureSheet(name, headers) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) sh.appendRow(headers);
    return sh;
  }

  ensureSheet(SH.PRODUCTS,     ['id','name','category','price','cost','stockKH','stockTN','stockES','status','visible']);
  ensureSheet(SH.TRANSACTIONS, ['id','date','store','memberId','memberName','subtotal','discount','total','cost','pay','earnedPoints','note','customerType','source','items']);
  ensureSheet(SH.MEMBERS,      ['id','name','phone','birthday','totalPoints','totalSpend','createdAt']);
  ensureSheet(SH.INV_LOG,      ['time','productId','name','store','oldQty','newQty','diff','note','type']);
  ensureSheet(SH.HANG_ORDERS,  ['id','name','store','createdAt','cart','member']);
  ensureSheet(SH.PRICE_LOG,    ['time','name','category','bundlePrice','stems','unitCost','salePrice','note']);

  Logger.log('✅ 所有工作表初始化完成');
}

function setupProcurementSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  function ensureSheet(name, headers) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) sh.appendRow(headers);
    return sh;
  }
  // [v2.2 修復] 補上 store 欄位，讓進貨記錄可以按點位查詢
  ensureSheet(SH_PROC.BATCHES, ['batchId','date','source','store','note','totalItems','totalCost']);
  ensureSheet(SH_PROC.ITEMS,   ['itemId','batchId','flowerName','category','stemsPerBunch','bunchesQty','pricePerBunch','totalStems','costPerStem','suggestedPrice','store','date']);
  Logger.log('✅ 進貨工作表初始化完成');
}

// ══════════════════════════════════════════════════════
//  CORS 設定（讓前端可跨域呼叫）
// ══════════════════════════════════════════════════════
function setCORSHeaders(output) {
  // GAS doGet/doPost 透過 ContentService 不支援自訂 header，
  // 但「以我的身分執行、允許所有人」模式下瀏覽器會通過 CORS。
  return output.setMimeType(ContentService.MimeType.JSON);
}

function jsonOut(data) {
  return setCORSHeaders(
    ContentService.createTextOutput(JSON.stringify(data))
  );
}

// ══════════════════════════════════════════════════════
//  入口點
// ══════════════════════════════════════════════════════
function doGet(e) {
  try {
    const action = e.parameter.action || '';
    switch (action) {
      case 'getProducts':     return jsonOut(getProducts());
      case 'getMembers':      return jsonOut(getMembers());
      case 'getInventoryLog': return jsonOut(getInventoryLog(e.parameter.store));
      case 'getHangOrders':   return jsonOut(getHangOrders(e.parameter.store));
      case 'getSalesReport':        return jsonOut(getSalesReport(e.parameter.store, e.parameter.from, e.parameter.to));
      case 'getDailyReport':        return jsonOut(getDailyReport(e.parameter.store, e.parameter.date));
      case 'getPayBreakdown':       return jsonOut(getPayBreakdown(e.parameter.store, e.parameter.date));
      case 'getWeeklyTrend':        return jsonOut(getWeeklyTrend(e.parameter.store));
      case 'getItemRanking':        return jsonOut(getItemRanking(e.parameter.store, e.parameter.date));
      case 'getMemberStats':        return jsonOut(getMemberStats());
      case 'getTransactionHistory': return jsonOut(getTransactionHistory(e.parameter.store, e.parameter.date));
      case 'getAllStoresReport':     return jsonOut(getAllStoresReport(e.parameter.period));
      case 'getFlowerLibrary':      return jsonOut(getFlowerLibrary());
      case 'getPriceHistory':         return jsonOut(getPriceHistory(e.parameter.name));
      case 'getProcurementBatches':   return jsonOut(getProcurementBatches());
      case 'getProcurementItems':     return jsonOut(getProcurementItems(e.parameter.batchId));
      case 'getFlowerProcHistory':    return jsonOut(getFlowerProcHistory(e.parameter.name));
      // [v2.4] sales.html 銷售紀錄頁面 — 與 getTransactionHistory 相同功能
      case 'getHistory':              return jsonOut(getTransactionHistory(e.parameter.store, e.parameter.date));
      default:                        return jsonOut({ error: 'Unknown GET action: ' + action });
    }
  } catch (err) {
    return jsonOut({ error: err.message });
  }
}

function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action || '';
    const data   = body.data   || {};
    switch (action) {
      // 商品
      case 'addProduct':    return jsonOut(addProduct(data));
      case 'updateProduct': return jsonOut(updateProduct(data));
      case 'deleteProduct':     return jsonOut(deleteProduct(data));
      case 'reactivateProduct': return jsonOut(reactivateProduct(data));
      // 庫存
      case 'updateStock':   return jsonOut(updateStock(data));
      case 'receiveStock':  return jsonOut(receiveStock(data));
      // 交易
      case 'addTransaction':return jsonOut(addTransaction(data));
      // 會員
      case 'addMember':     return jsonOut(addMember(data));
      case 'updateMemberPoints': return jsonOut(updateMemberPoints(data));
      // 掛單
      case 'saveHangOrder':   return jsonOut(saveHangOrder(data));
      case 'deleteHangOrder':      return jsonOut(deleteHangOrder(data));
      case 'logPriceHistory':      return jsonOut(logPriceHistory(data));
      case 'addFlowerToLibrary':   return jsonOut(addFlowerToLibrary(data));
      case 'addProcurement':          return jsonOut(addProcurement(data));
      case 'reportWaste':             return jsonOut(reportWaste(data));
      case 'voidTransaction':         return jsonOut(voidTransaction(data));
      case 'deleteProcurementBatch':  return jsonOut(deleteProcurementBatch(data));
      // [v2.4] inventory.html 庫存管理頁面快速調整 +/-
      case 'adjustStock':             return jsonOut(adjustStock(data));
      default: return jsonOut({ error: 'Unknown POST action: ' + action });
    }
  } catch (err) {
    return jsonOut({ error: err.message });
  }
}

// ══════════════════════════════════════════════════════
//  工具函式
// ══════════════════════════════════════════════════════

// [v2.2 效能優化] 快取試算表物件，避免每次呼叫都重新開啟
let _ss = null;
function getSS() {
  if (!_ss) _ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return _ss;
}
function getSheet(name) {
  return getSS().getSheetByName(name);
}

function sheetToObjects(sh) {
  if (sh.getLastRow() < 2) return [];
  const [headers, ...rows] = sh.getDataRange().getValues();
  return rows.map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      const v = row[i];
      // Google Sheets auto-parses ISO strings → Date objects; convert back to Taipei time
      obj[h] = v instanceof Date
        ? Utilities.formatDate(v, 'Asia/Taipei', "yyyy-MM-dd'T'HH:mm:ss")
        : v;
    });
    return obj;
  });
}

function taipeiNow() {
  return Utilities.formatDate(
    new Date(), 'Asia/Taipei', "yyyy-MM-dd'T'HH:mm:ss"
  );
}

function newId(prefix) {
  return prefix + Date.now().toString().slice(-8) + Math.random().toString(36).slice(2, 5);
}

// ══════════════════════════════════════════════════════
//  商品管理
// ══════════════════════════════════════════════════════
function getProducts() {
  const sh   = getSheet(SH.PRODUCTS);
  const rows = sheetToObjects(sh);
  // [v2.3] 軟刪除：只回傳 status 為 'active' 或空值（舊資料）的商品
  return rows
    .filter(r => (r.status || 'active') !== 'inactive')
    .map(r => ({
      id      : Number(r.id),
      name    : r.name,
      category: r.category || '其他加購',
      price   : Number(r.price)   || 0,
      cost    : Number(r.cost)    || 0,
      stockKH : Number(r.stockKH) || 0,
      stockTN : Number(r.stockTN) || 0,
      stockES : Number(r.stockES) || 0,
      // [v2.4] visible：空值舊資料視為 true（向下相容）
      visible : r.visible === false || r.visible === 'false' ? false : true,
    }));
}

function addProduct(data) {
  const sh      = getSheet(SH.PRODUCTS);
  const existing = sheetToObjects(sh);
  const maxId   = existing.reduce((m, r) => Math.max(m, Number(r.id) || 0), 0);
  const id      = maxId + 1;

  sh.appendRow([
    id,
    data.name     || '',
    data.category || '其他加購',
    Number(data.price)   || 0,
    Number(data.cost)    || 0,
    Number(data.stockKH) || 0,
    Number(data.stockTN) || 0,
    Number(data.stockES) || 0,
    'active',  // [v2.3] status 欄位
    true,      // [v2.4] visible 欄位：新增商品預設顯示在收銀台
  ]);
  return { success: true, id };
}

function updateProduct(data) {
  const sh   = getSheet(SH.PRODUCTS);
  const vals = sh.getDataRange().getValues();
  const [headers] = vals;
  const idCol = headers.indexOf('id');

  for (let i = 1; i < vals.length; i++) {
    if (Number(vals[i][idCol]) === Number(data.id)) {
      const row = i + 1;
      const set = (col, val) => sh.getRange(row, headers.indexOf(col) + 1).setValue(val);
      if (data.name     !== undefined) set('name',     data.name);
      if (data.category !== undefined) set('category', data.category);
      if (data.price    !== undefined) set('price',    Number(data.price));
      if (data.cost     !== undefined) set('cost',     Number(data.cost));
      if (data.stockKH  !== undefined) set('stockKH',  Number(data.stockKH));
      if (data.stockTN  !== undefined) set('stockTN',  Number(data.stockTN));
      if (data.stockES  !== undefined) set('stockES',  Number(data.stockES));
      // [v2.4] visible 欄位：動態新增欄若不存在
      if (data.visible !== undefined) {
        let visCol = headers.indexOf('visible');
        if (visCol === -1) {
          visCol = headers.length;
          sh.getRange(1, visCol + 1).setValue('visible');
        }
        sh.getRange(row, visCol + 1).setValue(data.visible === true || data.visible === 'true');
      }
      return { success: true };
    }
  }
  return { success: false, error: '找不到商品' };
}

function deleteProduct(data) {
  // [v2.3] 軟刪除：不實際刪列，改為設定 status = 'inactive'
  // 這樣歷史報表仍可透過商品 ID 查到名稱與成本，不會出現資料斷鏈
  const sh      = getSheet(SH.PRODUCTS);
  const vals    = sh.getDataRange().getValues();
  const headers = vals[0];
  const idCol   = headers.indexOf('id');

  // 若 status 欄不存在（舊試算表），動態新增
  let statusCol = headers.indexOf('status');
  if (statusCol === -1) {
    statusCol = headers.length;
    sh.getRange(1, statusCol + 1).setValue('status');
  }

  for (let i = 1; i < vals.length; i++) {
    if (Number(vals[i][idCol]) === Number(data.id)) {
      sh.getRange(i + 1, statusCol + 1).setValue('inactive');
      return { success: true };
    }
  }
  return { success: false, error: '找不到商品' };
}

function reactivateProduct(data) {
  // [v2.3] 重新啟用已停用商品
  const sh      = getSheet(SH.PRODUCTS);
  const vals    = sh.getDataRange().getValues();
  const headers = vals[0];
  const idCol   = headers.indexOf('id');
  let   statusCol = headers.indexOf('status');
  if (statusCol === -1) {
    statusCol = headers.length;
    sh.getRange(1, statusCol + 1).setValue('status');
  }
  for (let i = 1; i < vals.length; i++) {
    if (Number(vals[i][idCol]) === Number(data.id)) {
      sh.getRange(i + 1, statusCol + 1).setValue('active');
      return { success: true };
    }
  }
  return { success: false, error: '找不到商品' };
}

// ══════════════════════════════════════════════════════
//  庫存操作
// ══════════════════════════════════════════════════════
function updateStock(data) {
  // data: { id, store, qty, note }
  const sh      = getSheet(SH.PRODUCTS);
  const vals    = sh.getDataRange().getValues();
  const headers = vals[0];
  const idCol   = headers.indexOf('id');
  const storeCol = data.store === '台南FOCUS' ? headers.indexOf('stockTN') : data.store === '誠品生活台南' ? headers.indexOf('stockES') : headers.indexOf('stockKH');

  for (let i = 1; i < vals.length; i++) {
    if (Number(vals[i][idCol]) === Number(data.id)) {
      const oldQty = Number(vals[i][storeCol]) || 0;
      const newQty = Number(data.qty);
      sh.getRange(i + 1, storeCol + 1).setValue(newQty);

      // 寫入盤點記錄
      getSheet(SH.INV_LOG).appendRow([
        taipeiNow(),
        data.id,
        vals[i][headers.indexOf('name')],
        data.store,
        oldQty,
        newQty,
        newQty - oldQty,
        data.note || '盤點調整',
        'adjust',
      ]);
      return { success: true, newStock: newQty };
    }
  }
  return { success: false, error: '找不到商品' };
}

function receiveStock(data) {
  // data: { id, store, qty, unitCost, note }
  const sh      = getSheet(SH.PRODUCTS);
  const vals    = sh.getDataRange().getValues();
  const headers = vals[0];
  const idCol   = headers.indexOf('id');
  const storeCol = data.store === '台南FOCUS' ? headers.indexOf('stockTN') : data.store === '誠品生活台南' ? headers.indexOf('stockES') : headers.indexOf('stockKH');

  for (let i = 1; i < vals.length; i++) {
    if (Number(vals[i][idCol]) === Number(data.id)) {
      const oldQty = Number(vals[i][storeCol]) || 0;
      const newQty = oldQty + Number(data.qty);
      sh.getRange(i + 1, storeCol + 1).setValue(newQty);

      // 更新進價（若有提供）
      if (data.unitCost && Number(data.unitCost) > 0) {
        sh.getRange(i + 1, headers.indexOf('cost') + 1).setValue(Number(data.unitCost));
      }

      // 寫入盤點記錄
      getSheet(SH.INV_LOG).appendRow([
        taipeiNow(),
        data.id,
        vals[i][headers.indexOf('name')],
        data.store,
        oldQty,
        newQty,
        Number(data.qty),
        data.note || '進貨入庫',
        'receive',
      ]);
      return { success: true, newStock: newQty };
    }
  }
  return { success: false, error: '找不到商品' };
}

// [v2.4] inventory.html 庫存管理頁面 — 快速 +/- 調整
// data: { id, store: 'KH'|'TN'|'ES', delta: +1|-1, note? }
function adjustStock(data) {
  const storeMap = {
    'KH': '高雄FOCUS 13',
    'TN': '台南FOCUS',
    'ES': '誠品生活台南',
  };
  // 支援短代碼（KH/TN/ES）或完整店名
  const storeName = storeMap[data.store] || data.store || '高雄FOCUS 13';
  const delta     = Number(data.delta) || 0;
  if (delta === 0) return { success: false, error: 'delta 不可為 0' };

  const sh      = getSheet(SH.PRODUCTS);
  const vals    = sh.getDataRange().getValues();
  const headers = vals[0];
  const idCol   = headers.indexOf('id');
  const colKey  = storeName === '台南FOCUS' ? 'stockTN' : storeName === '誠品生活台南' ? 'stockES' : 'stockKH';
  const storeCol = headers.indexOf(colKey);

  for (let i = 1; i < vals.length; i++) {
    if (Number(vals[i][idCol]) === Number(data.id)) {
      const oldQty = Number(vals[i][storeCol]) || 0;
      const newQty = Math.max(0, oldQty + delta);
      sh.getRange(i + 1, storeCol + 1).setValue(newQty);

      // 寫入庫存日誌
      getSheet(SH.INV_LOG).appendRow([
        taipeiNow(),
        data.id,
        vals[i][headers.indexOf('name')],
        storeName,
        oldQty,
        newQty,
        delta,
        data.note || (delta > 0 ? '手動增加' : '手動減少'),
        'adjust',
      ]);
      return { success: true, newStock: newQty };
    }
  }
  return { success: false, error: '找不到商品' };
}

function getInventoryLog(store) {
  const sh   = getSheet(SH.INV_LOG);
  const rows = sheetToObjects(sh);
  // [v2.2 修復] 移除 !r.store 條件，不再把「無點位」記錄混入點位查詢
  const filtered = store
    ? rows.filter(r => r.store === store)
    : rows;

  return filtered
    .reverse()  // 最新在前
    .slice(0, 200)
    .map(r => ({
      time  : r.time,
      name  : r.name,
      store : r.store,
      oldQty: Number(r.oldQty),
      newQty: Number(r.newQty),
      diff  : Number(r.diff),
      note  : r.note,
      type  : r.type,
    }));
}

// ══════════════════════════════════════════════════════
//  交易（結帳）
// ══════════════════════════════════════════════════════
function addTransaction(data) {
  const sh   = getSheet(SH.TRANSACTIONS);
  const txId = data.id || newId('T');

  // 去重：同一 ID 已存在則直接回傳成功，防止網路重試造成重複寫入
  if (sh.getLastRow() >= 2) {
    const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().flat();
    if (ids.includes(txId)) return { success: true, duplicate: true };
  }

  sh.appendRow([
    txId,
    data.date         || taipeiNow(),
    data.store        || '',
    data.memberId     || '',
    data.memberName   || '',
    Number(data.subtotal)     || 0,
    Number(data.discount)     || 0,
    Number(data.total)        || 0,
    Number(data.cost)         || 0,
    data.pay          || '現金',
    Number(data.earnedPoints) || 0,
    data.note         || '',
    data.customerType || '',
    data.source       || '',
    JSON.stringify(data.items || []),
  ]);

  // 扣減庫存
  if (data.items && data.items.length) {
    data.items.forEach(item => {
      if (item.id && Number(item.id) > 0 && item.qty > 0) {
        updateStockDelta(item.id, data.store, -item.qty, `銷售扣庫：${txId}`);
      }
    });
  }

  // 累積會員點數
  if (data.memberId && Number(data.earnedPoints) > 0) {
    updateMemberPoints({ id: data.memberId, addPoints: data.earnedPoints, addSpend: data.total });
  }

  return { success: true };
}

// 修正版：支援 delta 扣減
function updateStockDelta(id, store, delta, note) {
  const sh      = getSheet(SH.PRODUCTS);
  const vals    = sh.getDataRange().getValues();
  const headers = vals[0];
  const idCol   = headers.indexOf('id');
  const storeCol = store === '台南FOCUS' ? headers.indexOf('stockTN') : store === '誠品生活台南' ? headers.indexOf('stockES') : headers.indexOf('stockKH');

  for (let i = 1; i < vals.length; i++) {
    if (Number(vals[i][idCol]) === Number(id)) {
      const oldQty = Number(vals[i][storeCol]) || 0;
      const newQty = Math.max(0, oldQty + delta);
      sh.getRange(i + 1, storeCol + 1).setValue(newQty);
      getSheet(SH.INV_LOG).appendRow([
        taipeiNow(), id, vals[i][headers.indexOf('name')],
        store, oldQty, newQty, delta, note, 'sale',
      ]);
      return newQty;
    }
  }
}

// ══════════════════════════════════════════════════════
//  銷售報表
// ══════════════════════════════════════════════════════
function getSalesReport(store, from, to) {
  const sh   = getSheet(SH.TRANSACTIONS);
  const rows = sheetToObjects(sh);

  const fromDate = from ? new Date(from) : new Date(0);
  const toDate   = to   ? new Date(to + 'T23:59:59') : new Date();

  const filtered = rows.filter(r => {
    const d = new Date(r.date);
    if (isNaN(d)) return false;
    if (store && r.store !== store) return false;
    if (r.voided === true || String(r.voided).toLowerCase() === 'true') return false; // [v2.4] 排除已退款
    return d >= fromDate && d <= toDate;
  });

  const totalRevenue = filtered.reduce((s, r) => s + (Number(r.total) || 0), 0);
  const totalCost    = filtered.reduce((s, r) => s + (Number(r.cost)  || 0), 0);
  const totalProfit  = totalRevenue - totalCost;
  const txCount      = filtered.length;

  // 付款方式統計
  const payBreakdown = {};
  filtered.forEach(r => {
    const pay = r.pay || '其他';
    payBreakdown[pay] = (payBreakdown[pay] || 0) + (Number(r.total) || 0);
  });

  return {
    totalRevenue,
    totalCost,
    totalProfit,
    txCount,
    payBreakdown,
    transactions: filtered.slice(-100).reverse(),
  };
}

// ══════════════════════════════════════════════════════
//  分析報表
// ══════════════════════════════════════════════════════

function dateRange_(date, tz) {
  // Returns { from, to } for a given 'YYYY-MM-DD' string in Taipei time
  const base = date || Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');
  const from = new Date(base + 'T00:00:00+08:00');
  const to   = new Date(base + 'T23:59:59+08:00');
  return { from, to };
}

function storeFilter_(store) {
  // [v2.2 修復] Studio 獨立記錄，不再合併至高雄
  if (store) return r => r.store === store;
  return () => true;
}

function getDailyReport(store, date) {
  const sh   = getSheet(SH.TRANSACTIONS);
  const rows = sheetToObjects(sh);
  const { from, to } = dateRange_(date);

  const filt    = storeFilter_(store);
  const notVoid = r => !(r.voided === true || String(r.voided).toLowerCase() === 'true'); // [v2.4]
  const todayTxs = rows.filter(r => {
    const d = new Date(r.date);
    return !isNaN(d) && filt(r) && notVoid(r) && d >= from && d <= to;
  });

  // Same weekday last week
  const lwFrom = new Date(from); lwFrom.setDate(lwFrom.getDate() - 7);
  const lwTo   = new Date(to);   lwTo.setDate(lwTo.getDate() - 7);
  const lwTxs  = rows.filter(r => {
    const d = new Date(r.date);
    return !isNaN(d) && filt(r) && notVoid(r) && d >= lwFrom && d <= lwTo;
  });

  const compute = txs => ({
    revenue  : txs.reduce((s,r) => s + (Number(r.total)||0), 0),
    cost     : txs.reduce((s,r) => s + (Number(r.cost)||0),  0),
    profit   : txs.reduce((s,r) => s + (Number(r.total)||0) - (Number(r.cost)||0), 0),
    customers: txs.length,
  });

  const sourceBreakdown = {};
  let newCount = 0, memberCount = 0;
  todayTxs.forEach(r => {
    if (r.source) sourceBreakdown[r.source] = (sourceBreakdown[r.source]||0) + 1;
    if (r.customerType === '新客') newCount++;
    if (r.memberId) memberCount++;
  });

  // Waste cost from inventory log
  const prodMap = {};
  sheetToObjects(getSheet(SH.PRODUCTS)).forEach(r => {
    prodMap[Number(r.id)] = Number(r.cost)||0;
  });
  const totalWasteCost = sheetToObjects(getSheet(SH.INV_LOG))
    .filter(r => r.type === 'waste' && !isNaN(new Date(r.time)) && new Date(r.time) >= from && new Date(r.time) <= to)
    .reduce((s,r) => s + (prodMap[Number(r.productId)]||0) * Math.abs(Number(r.diff)||0), 0);

  return { ...compute(todayTxs), lastWeek: compute(lwTxs), sourceBreakdown, newCount, memberCount, totalWasteCost };
}

function getPayBreakdown(store, date) {
  const sh   = getSheet(SH.TRANSACTIONS);
  const rows = sheetToObjects(sh);
  const { from, to } = dateRange_(date);
  const filt = storeFilter_(store);

  const result = {};
  rows.filter(r => {
    const d = new Date(r.date);
    return !isNaN(d) && filt(r) && d >= from && d <= to;
  }).forEach(r => {
    const pay = r.pay || '其他';
    result[pay] = (result[pay]||0) + (Number(r.total)||0);
  });
  return result;
}

function getWeeklyTrend(store) {
  const sh   = getSheet(SH.TRANSACTIONS);
  const rows = sheetToObjects(sh);
  const filt = storeFilter_(store);
  const now  = new Date();
  const result = [];

  for (let w = 3; w >= 0; w--) {
    const dow = now.getDay() || 7; // 1=Mon
    const weekStart = new Date(now);
    weekStart.setDate(now.getDate() - dow + 1 - w * 7);
    weekStart.setHours(0,0,0,0);
    const weekEnd = new Date(weekStart);
    weekEnd.setDate(weekStart.getDate() + 6);
    weekEnd.setHours(23,59,59,999);

    const txs      = rows.filter(r => { const d = new Date(r.date); const iv = r.voided===true||String(r.voided).toLowerCase()==='true'; return !isNaN(d) && filt(r) && !iv && d >= weekStart && d <= weekEnd; });
    const revenue  = txs.reduce((s,r) => s + (Number(r.total)||0), 0);
    const cost     = txs.reduce((s,r) => s + (Number(r.cost)||0),  0);
    const label    = `${weekStart.getMonth()+1}/${weekStart.getDate()}`;
    result.push({ label, revenue, cost, profit: revenue - cost, customers: txs.length });
  }
  return result;
}

function getItemRanking(store, date) {
  const sh   = getSheet(SH.TRANSACTIONS);
  const rows = sheetToObjects(sh);
  const filt = storeFilter_(store);

  // Use current month (month-to-date)
  const base = date ? new Date(date + 'T00:00:00+08:00') : new Date();
  const from = new Date(base.getFullYear(), base.getMonth(), 1);
  const to   = new Date(base.getFullYear(), base.getMonth() + 1, 0, 23, 59, 59);

  const prodMap = {};
  sheetToObjects(getSheet(SH.PRODUCTS)).forEach(r => {
    prodMap[Number(r.id)] = { cost: Number(r.cost)||0, name: r.name };
  });

  const itemMap = {};
  rows.filter(r => { const d = new Date(r.date); const iv = r.voided===true||String(r.voided).toLowerCase()==='true'; return !isNaN(d) && filt(r) && !iv && d >= from && d <= to; }) // [v2.4]
    .forEach(tx => {
      const items = safeJson(tx.items, []);
      items.forEach(item => {
        const key = String(item.id || item.name);
        if (!itemMap[key]) {
          itemMap[key] = { name: item.name || (prodMap[Number(item.id)] ? prodMap[Number(item.id)].name : key), qty: 0, sales: 0, cost: 0, waste: 0 };
        }
        const qty  = Number(item.qty)   || 0;
        const price= Number(item.price) || 0;
        const uc   = prodMap[Number(item.id)] ? prodMap[Number(item.id)].cost : 0;
        itemMap[key].qty   += qty;
        itemMap[key].sales += price * qty;
        itemMap[key].cost  += uc * qty;
      });
    });

  return Object.values(itemMap).map(it => {
    const profit    = it.sales - it.cost;
    const margin    = it.sales > 0 ? Math.round(profit / it.sales * 100) : 0;
    const wasteRate = (it.qty + it.waste) > 0 ? Math.round(it.waste / (it.qty + it.waste) * 100) : 0;
    return { name: it.name, qty: it.qty, sales: Math.round(it.sales), cost: Math.round(it.cost), profit: Math.round(profit), margin, waste: it.waste, wasteRate };
  }).sort((a,b) => b.sales - a.sales);
}

function getMemberStats() {
  const rows = sheetToObjects(getSheet(SH.MEMBERS));
  const now  = new Date();
  const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);

  const tiers = { '銅卡': 0, '銀卡': 0, '金卡': 0, 'VIP': 0 };
  rows.forEach(r => {
    const spend = Number(r.totalSpend)  || 0;
    const pts   = Number(r.totalPoints) || 0;
    if (spend >= 50000 || pts >= 500)      tiers['VIP']++;
    else if (spend >= 20000 || pts >= 200) tiers['金卡']++;
    else if (spend >= 5000  || pts >= 50)  tiers['銀卡']++;
    else                                   tiers['銅卡']++;
  });

  return {
    total: rows.length,
    newThisMonth: rows.filter(r => r.createdAt && new Date(r.createdAt) >= monthStart).length,
    tiers,
  };
}

function getTransactionHistory(store, date) {
  const sh   = getSheet(SH.TRANSACTIONS);
  const rows = sheetToObjects(sh);
  const filt = storeFilter_(store);

  let filtered = rows.filter(r => filt(r));
  if (date) {
    const { from, to } = dateRange_(date);
    filtered = filtered.filter(r => { const d = new Date(r.date); return !isNaN(d) && d >= from && d <= to; });
  }

  return filtered.reverse().slice(0, 200).map(r => ({
    id        : r.id,
    date      : r.date,
    store     : r.store,
    memberName: r.memberName,
    total     : Number(r.total) || 0,
    cost      : Number(r.cost)  || 0,
    pay       : r.pay,
    note      : r.note,
    items     : safeJson(r.items, []),
    voided    : r.voided === true || String(r.voided).toLowerCase() === 'true', // [v2.4]
  }));
}

function getAllStoresReport(period) {
  // period: 'thisWeek' | 'lastWeek' | 'thisMonth' | 'lastMonth'
  const sh   = getSheet(SH.TRANSACTIONS);
  const rows = sheetToObjects(sh);
  const now  = new Date();
  let from, to;

  if (period === 'lastMonth') {
    const m = now.getMonth() === 0 ? 11 : now.getMonth() - 1;
    const y = now.getMonth() === 0 ? now.getFullYear() - 1 : now.getFullYear();
    from = new Date(y, m, 1);
    to   = new Date(y, m + 1, 0, 23, 59, 59);
  } else if (period === 'thisMonth') {
    from = new Date(now.getFullYear(), now.getMonth(), 1);
    to   = new Date(now.getFullYear(), now.getMonth() + 1, 0, 23, 59, 59);
  } else if (period === 'lastWeek') {
    const dow = now.getDay() || 7;
    from = new Date(now); from.setDate(now.getDate() - dow - 6); from.setHours(0,0,0,0);
    to   = new Date(from); to.setDate(from.getDate() + 6); to.setHours(23,59,59,999);
  } else { // thisWeek
    const dow = now.getDay() || 7;
    from = new Date(now); from.setDate(now.getDate() - dow + 1); from.setHours(0,0,0,0);
    to   = new Date(now); to.setHours(23,59,59,999);
  }

  const filtered = rows.filter(r => {
    const d = new Date(r.date);
    const isVoided = r.voided === true || String(r.voided).toLowerCase() === 'true';
    return !isNaN(d) && d >= from && d <= to && !isVoided; // [v2.4] 排除已退款
  });

  // [v2.2 修復] Studio 獨立報表，不再與高雄合併
  const storeList = [
    { label: '高雄FOCUS 13', keys: ['高雄FOCUS 13'] },
    { label: '台南FOCUS',    keys: ['台南FOCUS'] },
    { label: '誠品生活台南', keys: ['誠品生活台南'] },
    { label: 'Studio',       keys: ['Studio'] },
  ];

  const result = storeList.map(s => {
    const txs      = filtered.filter(r => s.keys.includes(r.store));
    const revenue  = txs.reduce((sum,r) => sum + (Number(r.total)||0), 0);
    const cost     = txs.reduce((sum,r) => sum + (Number(r.cost)||0),  0);
    const profit   = revenue - cost;
    const margin   = revenue > 0 ? Math.round(profit / revenue * 100) : 0;
    return { store: s.label, revenue, cost, profit, margin, customers: txs.length };
  });

  const totRev = result.reduce((s,r) => s + r.revenue, 0);
  const totCost= result.reduce((s,r) => s + r.cost,    0);
  const totPro = totRev - totCost;
  result.push({
    store: '合計', revenue: totRev, cost: totCost, profit: totPro,
    margin: totRev > 0 ? Math.round(totPro / totRev * 100) : 0,
    customers: result.reduce((s,r) => s + r.customers, 0),
  });

  return { stores: result, from: from.toISOString(), to: to.toISOString() };
}

// ══════════════════════════════════════════════════════
//  會員管理
// ══════════════════════════════════════════════════════
function getMembers() {
  const sh   = getSheet(SH.MEMBERS);
  const rows = sheetToObjects(sh);
  return rows.map(r => ({
    id         : r.id,
    name       : r.name,
    phone      : r.phone,
    birthday   : r.birthday,
    totalPoints: Number(r.totalPoints) || 0,
    totalSpend : Number(r.totalSpend)  || 0,
    createdAt  : r.createdAt,
  }));
}

function addMember(data) {
  const sh = getSheet(SH.MEMBERS);
  const id = newId('M');
  sh.appendRow([
    id,
    data.name     || '',
    data.phone    || '',
    data.birthday || '',
    0,  // totalPoints
    0,  // totalSpend
    taipeiNow(),
  ]);
  return { success: true, id, name: data.name, phone: data.phone, totalPoints: 0, totalSpend: 0 };
}

function updateMemberPoints(data) {
  // data: { id, addPoints, addSpend }  OR  { id, setPoints }
  const sh      = getSheet(SH.MEMBERS);
  const vals    = sh.getDataRange().getValues();
  const headers = vals[0];
  const idCol   = headers.indexOf('id');

  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][idCol]) === String(data.id)) {
      const row   = i + 1;
      const ptCol = headers.indexOf('totalPoints') + 1;
      const spCol = headers.indexOf('totalSpend')  + 1;
      if (data.setPoints !== undefined) {
        sh.getRange(row, ptCol).setValue(Number(data.setPoints));
      } else {
        const cur = Number(vals[i][headers.indexOf('totalPoints')]) || 0;
        sh.getRange(row, ptCol).setValue(cur + (Number(data.addPoints) || 0));
        const curSp = Number(vals[i][headers.indexOf('totalSpend')]) || 0;
        sh.getRange(row, spCol).setValue(curSp + (Number(data.addSpend) || 0));
      }
      return { success: true };
    }
  }
  return { success: false, error: '找不到會員' };
}

// ══════════════════════════════════════════════════════
//  掛單管理
// ══════════════════════════════════════════════════════
function getHangOrders(store) {
  const sh   = getSheet(SH.HANG_ORDERS);
  const rows = sheetToObjects(sh);
  const filtered = store ? rows.filter(r => !r.store || r.store === store) : rows;
  return filtered.map(r => ({
    id       : r.id,
    name     : r.name,
    store    : r.store,
    createdAt: r.createdAt,
    cart     : safeJson(r.cart,   []),
    member   : safeJson(r.member, null),
  }));
}

function saveHangOrder(data) {
  // data: { name, store, createdAt, cart, member }
  const sh = getSheet(SH.HANG_ORDERS);
  const id = newId('H');
  sh.appendRow([
    id,
    data.name      || '掛單',
    data.store     || '',
    data.createdAt || taipeiNow(),
    JSON.stringify(data.cart   || []),
    JSON.stringify(data.member || null),
  ]);
  return { success: true, id };
}

function deleteHangOrder(data) {
  // data: { id }
  const sh   = getSheet(SH.HANG_ORDERS);
  const vals = sh.getDataRange().getValues();
  const idCol = vals[0].indexOf('id');
  for (let i = vals.length - 1; i >= 1; i--) {
    if (String(vals[i][idCol]) === String(data.id)) {
      sh.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: '找不到掛單' };
}

// ══════════════════════════════════════════════════════
//  花材資料庫
// ══════════════════════════════════════════════════════

function addFlowerToLibrary(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName('花材資料庫');
  if (!sh) return { success: false, error: '找不到花材資料庫工作表' };

  const vals    = sh.getDataRange().getValues();
  const headers = vals[0]; // 第 1 列：類別名稱

  // 找到對應類別的欄位
  let colIdx = -1;
  for (let i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim() === String(data.category).trim()) { colIdx = i; break; }
  }

  if (colIdx === -1) {
    // 類別不存在 → 新增一欄
    colIdx = headers.length;
    sh.getRange(1, colIdx + 1).setValue(data.category);
    // 重新讀取 vals
    sh.getRange(1, colIdx + 1).setValue(data.category);
  }

  // 找到該欄第一個空白列（從第 2 列開始）
  let rowIdx = 1;
  while (rowIdx < vals.length && String(vals[rowIdx][colIdx] || '').trim() !== '') {
    rowIdx++;
  }

  sh.getRange(rowIdx + 1, colIdx + 1).setValue(data.name);
  return { success: true };
}

/**
 * 讀取「花材資料庫」工作表
 * 格式：第1列為花材大類標題，往下每格為該類的品種名稱
 * 回傳：[{ name, category }, ...]
 */
function getFlowerLibrary() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName('花材資料庫');
  if (!sh) return [];

  const vals = sh.getDataRange().getValues();
  if (vals.length === 0) return [];

  const headers = vals[0]; // 第 1 列：類別名稱
  const result  = [];

  for (let col = 0; col < headers.length; col++) {
    const category = String(headers[col] || '').trim();
    if (!category) continue;
    for (let row = 1; row < vals.length; row++) {
      const name = String(vals[row][col] || '').trim();
      if (name) result.push({ name, category });
    }
  }

  return result;
}

// ══════════════════════════════════════════════════════
//  花材價格歷史
// ══════════════════════════════════════════════════════
function logPriceHistory(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(SH.PRICE_LOG);
  if (!sh) {
    sh = ss.insertSheet(SH.PRICE_LOG);
    sh.appendRow(['time','name','category','bundlePrice','stems','unitCost','salePrice','note']);
  }
  sh.appendRow([
    taipeiNow(),
    data.name         || '',
    data.category     || '',
    Number(data.bundlePrice) || 0,
    Number(data.stems)       || 0,
    Number(data.unitCost)    || 0,
    Number(data.salePrice)   || 0,
    data.note         || '本週進花',
  ]);
  return { success: true };
}

function getPriceHistory(name) {
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SH.PRICE_LOG);
  if (!sh || sh.getLastRow() < 2) return [];
  const rows = sheetToObjects(sh);
  const filtered = name ? rows.filter(r => r.name === name) : rows;
  return filtered.reverse().slice(0, 300).map(r => ({
    time       : r.time,
    name       : r.name,
    category   : r.category,
    bundlePrice: Number(r.bundlePrice) || 0,
    stems      : Number(r.stems)       || 0,
    unitCost   : Number(r.unitCost)    || 0,
    salePrice  : Number(r.salePrice)   || 0,
    note       : r.note,
  }));
}

// ══════════════════════════════════════════════════════
//  一次性資料遷移工具（在 Apps Script 手動執行）
// ══════════════════════════════════════════════════════

/**
 * 從舊試算表複製所有資料到新試算表（一次性執行）
 * 舊試算表 ID: 17_tJXE_uXjLTwIzwxRIxjY0bvLFr0Cmmt241EbT29_w
 * 新試算表 ID: SPREADSHEET_ID（已在上方設定）
 *
 * 複製內容：花材資料庫 / Products / Members / Transactions / InventoryLog
 * 執行前請先在 Apps Script 手動點選此函式並執行
 */
function migrateFromOldSpreadsheet() {
  const OLD_ID = '17_tJXE_uXjLTwIzwxRIxjY0bvLFr0Cmmt241EbT29_w';
  const newSS  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const oldSS  = SpreadsheetApp.openById(OLD_ID);

  function copySheet(sheetName, clearFirst) {
    const oldSh = oldSS.getSheetByName(sheetName);
    if (!oldSh) { Logger.log('⚠️ 舊試算表找不到：' + sheetName); return 0; }
    const data = oldSh.getDataRange().getValues();
    if (data.length === 0) { Logger.log('⚠️ 空工作表：' + sheetName); return 0; }

    let newSh = newSS.getSheetByName(sheetName);
    if (!newSh) newSh = newSS.insertSheet(sheetName);

    if (clearFirst) {
      newSh.clearContents();
      newSh.getRange(1, 1, data.length, data[0].length).setValues(data);
      Logger.log('✅ ' + sheetName + ' 整份複製完成：' + data.length + ' 列');
      return data.length;
    } else {
      // 僅附加資料列（跳過標題行，避免重複）
      const newHeaders = newSh.getLastRow() > 0
        ? newSh.getRange(1, 1, 1, newSh.getLastColumn()).getValues()[0].map(String)
        : [];
      const oldHeaders = data[0].map(String);
      const rows = data.slice(1);
      if (rows.length === 0) { Logger.log('⚠️ ' + sheetName + ' 無資料列'); return 0; }

      if (newSh.getLastRow() === 0) {
        // 新表完全空白 → 整份貼入
        newSh.getRange(1, 1, data.length, data[0].length).setValues(data);
      } else {
        // 已有資料 → 只附加（假設欄位順序相同）
        newSh.getRange(newSh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
      }
      Logger.log('✅ ' + sheetName + ' 附加完成：' + rows.length + ' 筆');
      return rows.length;
    }
  }

  // 花材資料庫：整份覆蓋（格式特殊，欄位是類別名稱）
  copySheet('花材資料庫', true);

  // 各資料工作表：附加（不清空，保留新表已有資料）
  copySheet('Products',     false);
  copySheet('Members',      false);
  copySheet('Transactions', false);
  copySheet('InventoryLog', false);

  Logger.log('🎉 遷移完成！請重新整理前端頁面確認資料正確。');
}

/**
 * 從舊試算表重新建立 Products 工作表（清除後重建，正確欄位對應）
 * 舊格式：前幾列為摘要，某列為標題行（含「品項名稱」），其後為資料
 * 新格式：id | name | category | price | cost | stockKH | stockTN | stockES
 *
 * 執行此函式即可修復「商品名稱/價格不見」問題
 */
function cleanRebuildProducts() {
  const OLD_ID = '17_tJXE_uXjLTwIzwxRIxjY0bvLFr0Cmmt241EbT29_w';
  const ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
  const oldSS  = SpreadsheetApp.openById(OLD_ID);

  const oldSh = oldSS.getSheetByName('Products');
  if (!oldSh) { Logger.log('❌ 舊試算表找不到 Products'); return; }

  const allVals = oldSh.getDataRange().getValues();

  // ── 找標題行（含「品項名稱」的那列）──
  let headerRow = -1;
  for (let i = 0; i < allVals.length; i++) {
    if (allVals[i].some(c => String(c).includes('品項名稱'))) { headerRow = i; break; }
  }
  if (headerRow === -1) {
    Logger.log('⚠️ 找不到含「品項名稱」的標題行，嘗試從第3列開始');
    headerRow = 2; // 預設第3列（0-indexed=2）為標題
  }

  // ── 欄位對應（依舊格式）──
  // A=編號 B=品項名稱 C=顏色/品種 D=廠商 E=進價
  // I=草支售價 J=高雄分配 L=高雄實際售價 O=台南分配 T=類別
  const products = [];
  let autoId = 1;

  for (let i = headerRow + 1; i < allVals.length; i++) {
    const r     = allVals[i];
    const nameB = String(r[1] || '').trim();
    const nameC = String(r[2] || '').trim();
    if (!nameB && !nameC) continue;

    const name    = (nameB && nameC) ? `${nameB} ${nameC}` : (nameB || nameC);
    const id      = Number(r[0]) || autoId++;
    const cost    = Number(r[4])  || 0;
    const priceI  = Number(r[8])  || 0;
    const priceL  = Number(r[11]) || 0;
    const price   = Math.round(priceL || priceI || (cost > 0 ? cost * 4 : 0));
    const stockKH = Math.max(0, Math.round(Number(r[9])  || 0));
    const stockTN = Math.max(0, Math.round(Number(r[14]) || 0));
    const cat     = String(r[19] || '').trim() || '其他加購';

    products.push([id, name, cat, price, cost, stockKH, stockTN, 0]);
  }

  // ── 備份現有 Products ──
  const sh = ss.getSheetByName('Products');
  if (sh && sh.getLastRow() > 0) {
    let bak = ss.getSheetByName('Products_BAK');
    if (!bak) bak = ss.insertSheet('Products_BAK');
    else bak.clearContents();
    const cur = sh.getDataRange().getValues();
    bak.getRange(1, 1, cur.length, cur[0].length).setValues(cur);
    Logger.log('✅ 現有資料已備份至 Products_BAK');
  }

  // ── 清空並寫入新格式 ──
  const newSh = sh || ss.insertSheet('Products');
  newSh.clearContents();
  newSh.getRange(1, 1, 1, 8).setValues([['id','name','category','price','cost','stockKH','stockTN','stockES']]);
  if (products.length > 0) {
    newSh.getRange(2, 1, products.length, 8).setValues(products);
  }

  const msg = `✅ 重建完成：共 ${products.length} 筆商品（從舊試算表讀取）。舊資料備份在 Products_BAK。`;
  Logger.log(msg);
}

/**
 * 將舊格式 Products 工作表遷移為新系統格式
 * 舊格式：第1-2列為摘要，第3列為標題（編號/品項名稱/顏色/廠商/進價…）
 * 新格式：第1列標題 = id|name|category|price|cost|stockKH|stockTN|stockES
 *
 * 執行前會自動備份至 Products_BAK
 */
function migrateProductsToNewFormat() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName('Products');
  if (!sh) { Logger.log('❌ 找不到 Products 工作表'); return; }

  const allVals = sh.getDataRange().getValues();

  // ── 備份 ──
  let bak = ss.getSheetByName('Products_BAK');
  if (!bak) bak = ss.insertSheet('Products_BAK');
  else bak.clearContents();
  bak.getRange(1, 1, allVals.length, allVals[0].length).setValues(allVals);
  Logger.log('✅ 備份至 Products_BAK 完成');

  // ── 找到標題行（含「品項名稱」）或預設第3列 ──
  let headerRow = 2; // 0-indexed，預設第3列
  for (let i = 0; i < Math.min(5, allVals.length); i++) {
    if (allVals[i].some(c => String(c).includes('品項名稱'))) { headerRow = i; break; }
  }

  // ── 欄位對應（0-indexed）──
  const products = [];
  let autoId = 1;

  for (let i = headerRow + 1; i < allVals.length; i++) {
    const r     = allVals[i];
    const nameB = String(r[1] || '').trim();
    const nameC = String(r[2] || '').trim();
    if (!nameB && !nameC) continue;

    const name    = (nameB && nameC) ? `${nameB} ${nameC}` : (nameB || nameC);
    const id      = Number(r[0]) || autoId++;
    const cost    = Number(r[4])  || 0;
    const priceI  = Number(r[8])  || 0;
    const priceL  = Number(r[11]) || 0;
    const price   = Math.round(priceL || priceI || (cost > 0 ? cost * 4 : 0));
    const stockKH = Math.max(0, Math.round(Number(r[9])  || 0));
    const stockTN = Math.max(0, Math.round(Number(r[14]) || 0));
    const cat     = String(r[19] || '').trim() || '其他加購';

    products.push([id, name, cat, price, cost, stockKH, stockTN, 0]);
  }

  // ── 清空並寫入新格式 ──
  sh.clearContents();
  sh.getRange(1, 1, 1, 8).setValues([['id','name','category','price','cost','stockKH','stockTN','stockES']]);
  if (products.length > 0) {
    sh.getRange(2, 1, products.length, 8).setValues(products);
  }

  const msg = `✅ 遷移完成：共 ${products.length} 筆商品。備份在 Products_BAK。`;
  Logger.log(msg);
}

/**
 * 在 Transactions 工作表補上 items 欄位（若尚未存在）
 */
function fixTransactionsSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName('Transactions');
  if (!sh || sh.getLastRow() < 1) { Logger.log('❌ 找不到 Transactions'); return; }

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (headers.includes('items')) { Logger.log('✅ items 欄位已存在，無需處理'); return; }

  const nextCol = sh.getLastColumn() + 1;
  sh.getRange(1, nextCol).setValue('items');
  Logger.log(`✅ 已在第 ${nextCol} 欄新增 items 欄位`);
}

// ══════════════════════════════════════════════════════
//  進貨模組
// ══════════════════════════════════════════════════════
function getProcurementBatches() {
  const sh = getSheet(SH_PROC.BATCHES);
  if (!sh || sh.getLastRow() < 2) return [];
  // [v2.2 修復] 回傳 store 欄位
  return sheetToObjects(sh).reverse().slice(0, 100).map(r => ({
    batchId   : r.batchId,
    date      : r.date,
    source    : r.source,
    store     : r.store || '',
    note      : r.note,
    totalItems: Number(r.totalItems) || 0,
    totalCost : Number(r.totalCost)  || 0,
  }));
}

function getProcurementItems(batchId) {
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SH_PROC.ITEMS);
  if (!sh || sh.getLastRow() < 2) return [];
  const rows     = sheetToObjects(sh);
  const filtered = batchId ? rows.filter(r => r.batchId === batchId) : rows;
  return filtered.reverse().slice(0, 500).map(r => ({
    itemId        : r.itemId,
    batchId       : r.batchId,
    flowerName    : r.flowerName,
    category      : r.category,
    stemsPerBunch : Number(r.stemsPerBunch)  || 0,
    bunchesQty    : Number(r.bunchesQty)     || 0,
    pricePerBunch : Number(r.pricePerBunch)  || 0,
    totalStems    : Number(r.totalStems)     || 0,
    costPerStem   : Number(r.costPerStem)    || 0,
    suggestedPrice: Number(r.suggestedPrice) || 0,
    store         : r.store,
    date          : r.date,
  }));
}

function addProcurement(data) {
  const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
  const batchSh = ss.getSheetByName(SH_PROC.BATCHES);
  const itemSh  = ss.getSheetByName(SH_PROC.ITEMS);
  if (!batchSh || !itemSh) return { success: false, error: '進貨工作表不存在，請先執行 setupProcurementSheets()' };

  const date    = data.date || Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');
  const batchId = 'B' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd') +
                  Math.random().toString(36).slice(2, 5).toUpperCase();
  const items     = data.items || [];
  let   totalCost = 0;
  const itemRows  = [];

  items.forEach((item, idx) => {
    const stemsPerBunch  = Number(item.stemsPerBunch)  || 0;
    const bunchesQty     = Number(item.bunchesQty)     || 0;
    const pricePerBunch  = Number(item.pricePerBunch)  || 0;
    const totalStems     = stemsPerBunch * bunchesQty;
    const costPerStem    = stemsPerBunch > 0 ? Math.round(pricePerBunch / stemsPerBunch * 100) / 100 : 0;
    const suggestedPrice = Math.round(costPerStem * 3.8);
    totalCost += pricePerBunch * bunchesQty;
    itemRows.push([
      batchId + '-' + String(idx + 1).padStart(2, '0'),
      batchId, item.flowerName || '', item.category || '主花',
      stemsPerBunch, bunchesQty, pricePerBunch,
      totalStems, costPerStem, suggestedPrice, item.store || data.store || '', date,
    ]);
    // [v2.4] 商品不存在時自動建立，再更新成本與庫存
    ensureProductExists(item.flowerName, item.category, costPerStem, suggestedPrice);
    updateProductCostWeighted({ flowerName: item.flowerName, store: item.store || data.store, newStems: totalStems, newCost: costPerStem });

    // [v2.3] 進貨時同步記錄價格歷史（修復：以前只有「本週進花」流程才會寫入）
    const phSh = getSheet(SH.PRICE_LOG);
    if (phSh) {
      phSh.appendRow([
        taipeiNow(),
        item.flowerName || '',
        item.category   || '',
        pricePerBunch,
        stemsPerBunch,
        costPerStem,
        Math.round(costPerStem * 4),
        `進貨（${data.source || ''}）`,
      ]);
    }
  });

  // [v2.2 修復] 加入 store 欄位寫入
  batchSh.appendRow([batchId, date, data.source || '', data.store || '', data.note || '', items.length, Math.round(totalCost)]);
  if (itemRows.length > 0) {
    itemSh.getRange(itemSh.getLastRow() + 1, 1, itemRows.length, itemRows[0].length).setValues(itemRows);
  }
  return { success: true, batchId, totalItems: items.length, totalCost: Math.round(totalCost) };
}

// [v2.4] 進貨時若商品不存在，自動建立
function ensureProductExists(name, category, cost, price) {
  if (!name) return;
  const sh      = getSheet(SH.PRODUCTS);
  const vals    = sh.getDataRange().getValues();
  const headers = vals[0];
  const nameCol = headers.indexOf('name');
  // 檢查是否已存在
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][nameCol] || '').trim() === String(name).trim()) return;
  }
  // 不存在 → 自動建立
  const maxId = vals.slice(1).reduce((m, r) => Math.max(m, Number(r[headers.indexOf('id')]) || 0), 0);
  sh.appendRow([maxId + 1, name, category || '主花', price || 0, cost || 0, 0, 0, 0, 'active']);
  Logger.log(`✅ 自動建立商品：${name}`);
}

function updateProductCostWeighted({ flowerName, store, newStems, newCost }) {
  const sh      = getSheet(SH.PRODUCTS);
  const vals    = sh.getDataRange().getValues();
  const headers = vals[0];
  const nameCol    = headers.indexOf('name');
  const costCol    = headers.indexOf('cost');
  const stockKHCol = headers.indexOf('stockKH');
  const stockTNCol = headers.indexOf('stockTN');
  const stockESCol = headers.indexOf('stockES');
  const storeCol   = store === '台南FOCUS' ? stockTNCol : store === '誠品生活台南' ? stockESCol : stockKHCol;

  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][nameCol] || '').trim() !== String(flowerName || '').trim()) continue;
    const row          = i + 1;
    const oldCost      = Number(vals[i][costCol])  || 0;
    const oldStock     = Number(vals[i][storeCol]) || 0;
    const weightedCost = (oldStock + newStems) > 0
      ? Math.round(((oldCost * oldStock) + (newCost * newStems)) / (oldStock + newStems) * 100) / 100
      : newCost;
    sh.getRange(row, costCol  + 1).setValue(weightedCost);
    sh.getRange(row, storeCol + 1).setValue(oldStock + newStems);
    getSheet(SH.INV_LOG).appendRow([
      taipeiNow(), vals[i][headers.indexOf('id')], flowerName,
      store, oldStock, oldStock + newStems, newStems,
      `進貨入庫（加權成本 ${weightedCost}）`, 'receive',
    ]);
    return;
  }
  Logger.log(`⚠️ 進貨警告：找不到「${flowerName}」，成本未回寫`);
}

// ══════════════════════════════════════════════════════
//  花材進貨歷史查詢 [v2.3 新增]
// ══════════════════════════════════════════════════════
function getFlowerProcHistory(name) {
  const itemSh  = getSheet(SH_PROC.ITEMS);
  const batchSh = getSheet(SH_PROC.BATCHES);
  if (!itemSh || itemSh.getLastRow() < 2) return [];

  const items   = sheetToObjects(itemSh);
  const batches = (batchSh && batchSh.getLastRow() >= 2) ? sheetToObjects(batchSh) : [];
  const batchMap = {};
  batches.forEach(b => { batchMap[b.batchId] = b; });

  const filtered = name
    ? items.filter(r => String(r.flowerName || '').trim() === String(name).trim())
    : items;

  return filtered
    .map(r => ({
      date         : r.date || (batchMap[r.batchId] ? batchMap[r.batchId].date : ''),
      source       : batchMap[r.batchId] ? (batchMap[r.batchId].source || '') : '',
      store        : r.store || '',
      stemsPerBunch: Number(r.stemsPerBunch)  || 0,
      bunchesQty   : Number(r.bunchesQty)     || 0,
      pricePerBunch: Number(r.pricePerBunch)  || 0,
      totalStems   : Number(r.totalStems)     || 0,
      costPerStem  : Number(r.costPerStem)    || 0,
      suggestedPrice: Number(r.suggestedPrice) || 0,
    }))
    .sort((a, b) => new Date(b.date) - new Date(a.date))
    .slice(0, 8);
}

// ══════════════════════════════════════════════════════
//  報廢記錄 [v2.2 新增] — 正確寫入 type='waste'，日結廢棄成本才會正確計算
// ══════════════════════════════════════════════════════
function reportWaste(data) {
  // data: { id, store, qty, note }
  const sh      = getSheet(SH.PRODUCTS);
  const vals    = sh.getDataRange().getValues();
  const headers = vals[0];
  const idCol   = headers.indexOf('id');
  const storeCol = data.store === '台南FOCUS' ? headers.indexOf('stockTN')
                 : data.store === '誠品生活台南' ? headers.indexOf('stockES')
                 : headers.indexOf('stockKH');

  for (let i = 1; i < vals.length; i++) {
    if (Number(vals[i][idCol]) === Number(data.id)) {
      const oldQty = Number(vals[i][storeCol]) || 0;
      const wasted = Number(data.qty) || 0;
      const newQty = Math.max(0, oldQty - wasted);
      sh.getRange(i + 1, storeCol + 1).setValue(newQty);

      // type = 'waste'，getDailyReport 才能正確計算廢棄成本
      getSheet(SH.INV_LOG).appendRow([
        taipeiNow(),
        data.id,
        vals[i][headers.indexOf('name')],
        data.store,
        oldQty,
        newQty,
        -wasted,
        data.note || '報廢',
        'waste',
      ]);
      return { success: true, newStock: newQty };
    }
  }
  return { success: false, error: '找不到商品' };
}

// ══════════════════════════════════════════════════════
//  退款 [v2.4 新增]
// ══════════════════════════════════════════════════════
function voidTransaction(data) {
  // data: { txId, reason }
  const sh      = getSheet(SH.TRANSACTIONS);
  const vals    = sh.getDataRange().getValues();
  const headers = vals[0];
  const idCol   = headers.indexOf('id');

  // voided 欄若不存在則動態新增
  let voidedCol = headers.indexOf('voided');
  if (voidedCol === -1) {
    voidedCol = headers.length;
    sh.getRange(1, voidedCol + 1).setValue('voided');
  }

  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][idCol]) !== String(data.txId)) continue;

    // 已退款則拒絕
    const alreadyVoided = vals[i][voidedCol];
    if (alreadyVoided === true || String(alreadyVoided).toLowerCase() === 'true') {
      return { success: false, error: '此交易已退款過了' };
    }

    // 標記退款
    sh.getRange(i + 1, voidedCol + 1).setValue(true);

    // 還原庫存
    const store = vals[i][headers.indexOf('store')];
    const items = safeJson(String(vals[i][headers.indexOf('items')] || ''), []);
    items.forEach(item => {
      if (item.id && Number(item.id) > 0 && Number(item.qty) > 0) {
        updateStockDelta(item.id, store, Number(item.qty), `退款還原：${data.txId}`);
      }
    });

    // 扣回會員點數與消費金額
    const memberId  = String(vals[i][headers.indexOf('memberId')] || '');
    const earnedPts = Number(vals[i][headers.indexOf('earnedPoints')]) || 0;
    const txTotal   = Number(vals[i][headers.indexOf('total')]) || 0;
    if (memberId && earnedPts > 0) {
      updateMemberPoints({ id: memberId, addPoints: -earnedPts, addSpend: -txTotal });
    }

    return { success: true };
  }
  return { success: false, error: '找不到交易紀錄' };
}

// ══════════════════════════════════════════════════════
//  刪除進貨批次 [v2.4 新增]
// ══════════════════════════════════════════════════════
function deleteProcurementBatch(data) {
  // data: { batchId }
  const batchSh = getSheet(SH_PROC.BATCHES);
  const itemSh  = getSheet(SH_PROC.ITEMS);
  if (!batchSh || !itemSh) return { success: false, error: '進貨工作表不存在' };

  // 刪 batch 列
  const bVals  = batchSh.getDataRange().getValues();
  const bIdCol = bVals[0].indexOf('batchId');
  let deleted  = false;
  for (let i = bVals.length - 1; i >= 1; i--) {
    if (String(bVals[i][bIdCol]) === String(data.batchId)) {
      batchSh.deleteRow(i + 1);
      deleted = true;
      break;
    }
  }

  // 刪對應 items（逆序，避免 row index 錯位）
  if (itemSh.getLastRow() >= 2) {
    const iVals = itemSh.getDataRange().getValues();
    const ibCol = iVals[0].indexOf('batchId');
    for (let i = iVals.length - 1; i >= 1; i--) {
      if (String(iVals[i][ibCol]) === String(data.batchId)) {
        itemSh.deleteRow(i + 1);
      }
    }
  }

  return deleted
    ? { success: true }
    : { success: false, error: '找不到此進貨批次' };
}

// ══════════════════════════════════════════════════════
//  小工具
// ══════════════════════════════════════════════════════
function safeJson(str, fallback) {
  try   { return JSON.parse(str); }
  catch { return fallback; }
}
fix: v2.2 修復 Studio 分帳和廢棄成本
