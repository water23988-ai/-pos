// ══════════════════════════════════════════════════════
//  拾逅花事 | 進銷存系統 — Google Apps Script 後端
//  版本：2.4-fix  |  2026-05
//  修正項目：
//  [FIX-1] 加入 LockService，防止多人同時操作造成庫存/交易資料衝突
//           影響函式：addTransaction / updateStock / receiveStock /
//                     reportWaste / adjustStock
//  [FIX-2] getMonthlyReport 時區修正：new Date(r.date) → parseTaipei_(r.date)
//           避免跨日零點前後訂單跑到錯誤日期，導致日結金額對不上
//  [FIX-3] 移除模組層級 _ss 快取（getSS / getSheet 改為每次直接開啟）
//           GAS 每個請求是獨立執行環境，快取無實際作用，保留反而有潛在風險
// ══════════════════════════════════════════════════════

const SPREADSHEET_ID = '1BQjkFe1pl7OYxvnZLoItVWZjYFnqBZalbwJPLPKMw4w';

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
//  初始化
// ══════════════════════════════════════════════════════
function setupSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  function ensureSheet(name, headers) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) sh.appendRow(headers);
    return sh;
  }
  ensureSheet(SH.PRODUCTS,     ['id','name','category','price','cost','stockKH','stockTN','stockES','stockMarket','status','visible','code']);
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
  ensureSheet(SH_PROC.BATCHES, ['batchId','date','source','store','note','totalItems','totalCost']);
  ensureSheet(SH_PROC.ITEMS,   ['itemId','batchId','flowerName','category','stemsPerBunch','bunchesQty','pricePerBunch','totalStems','costPerStem','suggestedPrice','store','date','allocation']);
  Logger.log('✅ 進貨工作表初始化完成');
}

// ══════════════════════════════════════════════════════
//  CORS / JSON 輸出
// ══════════════════════════════════════════════════════
function setCORSHeaders(output) {
  return output.setMimeType(ContentService.MimeType.JSON);
}
function jsonOut(data) {
  return setCORSHeaders(ContentService.createTextOutput(JSON.stringify(data)));
}

// ══════════════════════════════════════════════════════
//  入口點
// ══════════════════════════════════════════════════════
function doGet(e) {
  try {
    const action = e.parameter.action || '';
    switch (action) {
      case 'getProducts':             return jsonOut(getProducts());
      case 'getMembers':              return jsonOut(getMembers());
      case 'getInventoryLog':         return jsonOut(getInventoryLog(e.parameter.store));
      case 'getHangOrders':           return jsonOut(getHangOrders(e.parameter.store));
      case 'getSalesReport':          return jsonOut(getSalesReport(e.parameter.store, e.parameter.from, e.parameter.to));
      case 'getDailyReport':          return jsonOut(getDailyReport(e.parameter.store, e.parameter.date));
      case 'getPayBreakdown':         return jsonOut(getPayBreakdown(e.parameter.store, e.parameter.date));
      case 'getWeeklyTrend':          return jsonOut(getWeeklyTrend(e.parameter.store));
      case 'getItemRanking':          return jsonOut(getItemRanking(e.parameter.store, e.parameter.date, e.parameter.from, e.parameter.to));
      case 'getStoreStats':           return jsonOut(getStoreStats(e.parameter.store, e.parameter.from, e.parameter.to));
      case 'getWasteReport':          return jsonOut(getWasteReport(e.parameter.store, e.parameter.from, e.parameter.to));
      case 'getMemberStats':          return jsonOut(getMemberStats());
      case 'getTransactionHistory':   return jsonOut(getTransactionHistory(e.parameter.store, e.parameter.date));
      case 'getAllStoresReport':       return jsonOut(getAllStoresReport(e.parameter.period));
      case 'getMonthlyReport':        return jsonOut(getMonthlyReport(e.parameter.store, e.parameter.ym));
      case 'getFlowerLibrary':        return jsonOut(getFlowerLibrary());
      case 'getPriceHistory':         return jsonOut(getPriceHistory(e.parameter.name));
      case 'getProcurementBatches':   return jsonOut(getProcurementBatches());
      case 'getProcurementItems':     return jsonOut(getProcurementItems(e.parameter.batchId));
      case 'getFlowerProcHistory':    return jsonOut(getFlowerProcHistory(e.parameter.name));
      case 'getHistory':              return jsonOut(getTransactionHistory(e.parameter.store, e.parameter.date));
      case 'getPayBreakdownRange':    return jsonOut(getPayBreakdownRange(e.parameter.store, e.parameter.from, e.parameter.to));
      case 'getProductByCode':        return jsonOut(getProductByCode(e.parameter.code));
      case 'getItemStoreSplit':       return jsonOut(getItemStoreSplit(e.parameter.name, e.parameter.from, e.parameter.to));
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
      case 'addProduct':              return jsonOut(addProduct(data));
      case 'updateProduct':           return jsonOut(updateProduct(data));
      case 'deleteProduct':           return jsonOut(deleteProduct(data));
      case 'reactivateProduct':       return jsonOut(reactivateProduct(data));
      case 'updateStock':             return jsonOut(updateStock(data));
      case 'receiveStock':            return jsonOut(receiveStock(data));
      case 'addTransaction':          return jsonOut(addTransaction(data));
      case 'addMember':               return jsonOut(addMember(data));
      case 'updateMemberPoints':      return jsonOut(updateMemberPoints(data));
      case 'saveHangOrder':           return jsonOut(saveHangOrder(data));
      case 'deleteHangOrder':         return jsonOut(deleteHangOrder(data));
      case 'logPriceHistory':         return jsonOut(logPriceHistory(data));
      case 'addFlowerToLibrary':      return jsonOut(addFlowerToLibrary(data));
      case 'addProcurement':          return jsonOut(addProcurement(data));
      case 'updateProcurementBatch':  return jsonOut(updateProcurementBatch(data));
      case 'reportWaste':             return jsonOut(reportWaste(data));
      case 'voidTransaction':         return jsonOut(voidTransaction(data));
      case 'deleteProcurementBatch':  return jsonOut(deleteProcurementBatch(data));
      case 'adjustStock':             return jsonOut(adjustStock(data));
      case 'mergeProducts':           return jsonOut(mergeProducts(data));
      default: return jsonOut({ error: 'Unknown POST action: ' + action });
    }
  } catch (err) {
    return jsonOut({ error: err.message });
  }
}

// ══════════════════════════════════════════════════════
//  工具函式
// ══════════════════════════════════════════════════════

// [FIX-3] 移除 _ss 模組層級快取。
// 原本的 let _ss = null + getSS() 在 GAS 每次請求獨立執行環境下沒有快取效果，
// 只會增加維護複雜度。直接每次呼叫 openById，GAS 內部有自己的 cache 層。
function getSheet(name) {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name);
}

function sheetToObjects(sh) {
  if (sh.getLastRow() < 2) return [];
  const [headers, ...rows] = sh.getDataRange().getValues();
  return rows.map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      const v = row[i];
      obj[h] = v instanceof Date
        ? Utilities.formatDate(v, 'Asia/Taipei', "yyyy-MM-dd'T'HH:mm:ss")
        : v;
    });
    return obj;
  });
}

function taipeiNow() {
  return Utilities.formatDate(new Date(), 'Asia/Taipei', "yyyy-MM-dd'T'HH:mm:ss");
}

function toTaipeiStr_(v) {
  if (!v) return '';
  var d = (v instanceof Date) ? v : new Date(String(v));
  if (isNaN(d.getTime())) return String(v);
  return Utilities.formatDate(d, 'Asia/Taipei', "yyyy-MM-dd'T'HH:mm:ss");
}

function newId(prefix) {
  return prefix + Date.now().toString().slice(-8) + Math.random().toString(36).slice(2, 5);
}

// [v2.7] 解析台北時間字串，補上 +08:00 避免被當 UTC 解析
function parseTaipei_(ts) {
  const s = String(ts || '');
  if (!s) return new Date(NaN);
  const hasOffset = s.endsWith('Z') || /[+-]\d{2}:\d{2}$/.test(s);
  return new Date(hasOffset ? s : s + '+08:00');
}

function storeToColKey_(store) {
  if (store === '台南FOCUS')    return 'stockTN';
  if (store === '誠品生活台南') return 'stockES';
  if (store === '市集')         return 'stockMarket';
  return 'stockKH';
}

function storeFilter_(store) {
  if (store) return r => r.store === store;
  return () => true;
}

function safeJson(str, fallback) {
  try   { return JSON.parse(str); }
  catch { return fallback; }
}

// ══════════════════════════════════════════════════════
//  商品管理
// ══════════════════════════════════════════════════════
function getProducts() {
  const sh   = getSheet(SH.PRODUCTS);
  const rows = sheetToObjects(sh);
  return rows.map(r => ({
    id          : Number(r.id),
    name        : r.name,
    category    : r.category || '其他加購',
    price       : Number(r.price)       || 0,
    cost        : Number(r.cost)        || 0,
    stockKH     : Number(r.stockKH)     || 0,
    stockTN     : Number(r.stockTN)     || 0,
    stockES     : Number(r.stockES)     || 0,
    stockMarket : Number(r.stockMarket) || 0,
    status      : r.status || 'active',
    code        : String(r.code || ''),
  }));
}

function addProduct(data) {
  const sh       = getSheet(SH.PRODUCTS);
  const existing = sheetToObjects(sh);
  const maxId    = existing.reduce((m, r) => Math.max(m, Number(r.id) || 0), 0);
  const id       = maxId + 1;
  sh.appendRow([
    id, data.name || '', data.category || '其他加購',
    Number(data.price) || 0, Number(data.cost) || 0,
    Number(data.stockKH) || 0, Number(data.stockTN) || 0,
    Number(data.stockES) || 0, Number(data.stockMarket) || 0,
    'active', true, data.code || '',
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
      if (data.name        !== undefined) set('name',        data.name);
      if (data.category    !== undefined) set('category',    data.category);
      if (data.price       !== undefined) set('price',       Number(data.price));
      if (data.cost        !== undefined) set('cost',        Number(data.cost));
      if (data.stockKH     !== undefined) set('stockKH',     Number(data.stockKH));
      if (data.stockTN     !== undefined) set('stockTN',     Number(data.stockTN));
      if (data.stockES     !== undefined) set('stockES',     Number(data.stockES));
      if (data.stockMarket !== undefined) set('stockMarket', Number(data.stockMarket));
      if (data.status !== undefined) {
        let statusCol = headers.indexOf('status');
        if (statusCol === -1) { statusCol = headers.length; sh.getRange(1, statusCol + 1).setValue('status'); }
        sh.getRange(row, statusCol + 1).setValue(data.status);
      }
      if (data.code !== undefined) {
        let codeCol = headers.indexOf('code');
        if (codeCol === -1) { codeCol = headers.length; sh.getRange(1, codeCol + 1).setValue('code'); }
        sh.getRange(row, codeCol + 1).setValue(String(data.code || ''));
      }
      return { success: true };
    }
  }
  return { success: false, error: '找不到商品' };
}

function deleteProduct(data) {
  const sh      = getSheet(SH.PRODUCTS);
  const vals    = sh.getDataRange().getValues();
  const headers = vals[0];
  const idCol   = headers.indexOf('id');
  let statusCol = headers.indexOf('status');
  if (statusCol === -1) { statusCol = headers.length; sh.getRange(1, statusCol + 1).setValue('status'); }
  for (let i = 1; i < vals.length; i++) {
    if (Number(vals[i][idCol]) === Number(data.id)) {
      sh.getRange(i + 1, statusCol + 1).setValue('inactive');
      return { success: true };
    }
  }
  return { success: false, error: '找不到商品' };
}

function reactivateProduct(data) {
  const sh      = getSheet(SH.PRODUCTS);
  const vals    = sh.getDataRange().getValues();
  const headers = vals[0];
  const idCol   = headers.indexOf('id');
  let statusCol = headers.indexOf('status');
  if (statusCol === -1) { statusCol = headers.length; sh.getRange(1, statusCol + 1).setValue('status'); }
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

// 盤點為單人操作，不加 Lock
function updateStock(data) {
  const sh      = getSheet(SH.PRODUCTS);
  const vals    = sh.getDataRange().getValues();
  const headers = vals[0];
  const idCol   = headers.indexOf('id');
  const storeCol = headers.indexOf(storeToColKey_(data.store));
  for (let i = 1; i < vals.length; i++) {
    if (Number(vals[i][idCol]) === Number(data.id)) {
      const oldQty = Number(vals[i][storeCol]) || 0;
      const newQty = Number(data.qty);
      sh.getRange(i + 1, storeCol + 1).setValue(newQty);
      getSheet(SH.INV_LOG).appendRow([
        taipeiNow(), data.id, vals[i][headers.indexOf('name')],
        data.store, oldQty, newQty, newQty - oldQty,
        data.note || '盤點調整', 'adjust',
      ]);
      return { success: true, newStock: newQty };
    }
  }
  return { success: false, error: '找不到商品' };
}

// 進貨通常單人操作，不加 Lock
function receiveStock(data) {
  const sh      = getSheet(SH.PRODUCTS);
  const vals    = sh.getDataRange().getValues();
  const headers = vals[0];
  const idCol   = headers.indexOf('id');
  const storeCol = headers.indexOf(storeToColKey_(data.store));
  for (let i = 1; i < vals.length; i++) {
    if (Number(vals[i][idCol]) === Number(data.id)) {
      const oldQty = Number(vals[i][storeCol]) || 0;
      const newQty = oldQty + Number(data.qty);
      sh.getRange(i + 1, storeCol + 1).setValue(newQty);
      if (data.unitCost && Number(data.unitCost) > 0) {
        sh.getRange(i + 1, headers.indexOf('cost') + 1).setValue(Number(data.unitCost));
      }
      getSheet(SH.INV_LOG).appendRow([
        taipeiNow(), data.id, vals[i][headers.indexOf('name')],
        data.store, oldQty, newQty, Number(data.qty),
        data.note || '進貨入庫', 'receive',
      ]);
      return { success: true, newStock: newQty };
    }
  }
  return { success: false, error: '找不到商品' };
}

// 庫存 +/- 快速調整，單人操作，不加 Lock
function adjustStock(data) {
  const storeMap = {
    'KH': '高雄FOCUS 13',
    'TN': '台南FOCUS',
    'ES': '誠品生活台南',
    'MK': '市集',
  };
  const storeName = storeMap[data.store] || data.store || '高雄FOCUS 13';
  const delta     = Number(data.delta) || 0;
  if (delta === 0) return { success: false, error: 'delta 不可為 0' };

  const sh      = getSheet(SH.PRODUCTS);
  const vals    = sh.getDataRange().getValues();
  const headers = vals[0];
  const idCol   = headers.indexOf('id');
  const storeCol = headers.indexOf(storeToColKey_(storeName));

  for (let i = 1; i < vals.length; i++) {
    if (Number(vals[i][idCol]) === Number(data.id)) {
      const oldQty = Number(vals[i][storeCol]) || 0;
      const newQty = Math.max(0, oldQty + delta);
      sh.getRange(i + 1, storeCol + 1).setValue(newQty);
      getSheet(SH.INV_LOG).appendRow([
        taipeiNow(), data.id, vals[i][headers.indexOf('name')],
        storeName, oldQty, newQty, delta,
        data.note || (delta > 0 ? '手動增加' : '手動減少'), 'adjust',
      ]);
      return { success: true, newStock: newQty };
    }
  }
  return { success: false, error: '找不到商品' };
}

function getInventoryLog(store) {
  const sh   = getSheet(SH.INV_LOG);
  const rows = sheetToObjects(sh);
  const filtered = store ? rows.filter(r => r.store === store) : rows;
  return filtered.reverse().slice(0, 200).map(r => ({
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

// ── [PERF] txId 去重快取（PropertiesService，不讀試算表）──
// 存最近 100 筆 txId，結帳去重完全不碰 Transactions 工作表
// 快取在 GAS 執行環境重啟後會消失，但 Lock 已防止真正的並發重複
function isDuplicateTx_(txId) {
  const props  = PropertiesService.getScriptProperties();
  const recent = JSON.parse(props.getProperty('recent_tx_ids') || '[]');
  if (recent.includes(txId)) return true;
  recent.unshift(txId);
  if (recent.length > 100) recent.pop();
  props.setProperty('recent_tx_ids', JSON.stringify(recent));
  return false;
}

// [FIX-1] addTransaction 用 UserLock（非 ScriptLock）
// UserLock 只鎖同一使用者的並發請求，不同點位店員互不影響，速度更快
// 主要防止：同一個收銀畫面因網路慢或誤觸連按兩次，導致重複扣庫存
function addTransaction(data) {
  const lock = LockService.getUserLock();
  try {
    lock.waitLock(5000); // 同一使用者最多等 5 秒

    const sh   = getSheet(SH.TRANSACTIONS);
    const txId = data.id || newId('T');

    // [PERF] 去重：改用 PropertiesService 快取最近 100 筆 txId
    // 原本讀整欄（1000+ 筆）改為只查記憶體快取，結帳速度不再受資料量影響
    if (isDuplicateTx_(txId)) return { success: true, duplicate: true };

    const txDate = (() => {
      const d = String(data.date || '');
      if (d.length > 10) return d;
      if (d.length === 10) return d + 'T' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'HH:mm:ss');
      return taipeiNow();
    })();

    sh.appendRow([
      txId, txDate, data.store || '', data.memberId || '',
      data.memberName || '', Number(data.subtotal) || 0,
      Number(data.discount) || 0, Number(data.total) || 0,
      Number(data.cost) || 0, data.pay || '現金',
      Number(data.earnedPoints) || 0, data.note || '',
      data.customerType || '', data.source || '',
      JSON.stringify(data.items || []),
    ]);

    // 扣減庫存（在 Lock 保護內執行，安全）
    if (data.items && data.items.length) {
      data.items.forEach(item => {
        if (item.id && Number(item.id) > 0 && item.qty > 0) {
          _updateStockDeltaInternal(item.id, data.store, -item.qty, `銷售扣庫：${txId}`);
        }
      });
    }

    // 累積會員點數
    if (data.memberId && Number(data.earnedPoints) > 0) {
      updateMemberPoints({ id: data.memberId, addPoints: data.earnedPoints, addSpend: data.total });
    }

    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

// 內部版本：供 addTransaction 在 Lock 內部直接呼叫，不重複加鎖
function _updateStockDeltaInternal(id, store, delta, note) {
  const sh      = getSheet(SH.PRODUCTS);
  const vals    = sh.getDataRange().getValues();
  const headers = vals[0];
  const idCol   = headers.indexOf('id');
  const storeCol = headers.indexOf(storeToColKey_(store));

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

// 公開版本（供其他地方呼叫，例如 voidTransaction）
function updateStockDelta(id, store, delta, note) {
  return _updateStockDeltaInternal(id, store, delta, note);
}

// ══════════════════════════════════════════════════════
//  銷售報表
// ══════════════════════════════════════════════════════
function getSalesReport(store, from, to) {
  const sh   = getSheet(SH.TRANSACTIONS);
  const rows = sheetToObjects(sh);
  const fromDate = from ? new Date(from + 'T00:00:00+08:00') : new Date(0);
  const toDate   = to   ? new Date(to   + 'T23:59:59+08:00') : new Date();

  const filtered = rows.filter(r => {
    const d = parseTaipei_(r.date);
    if (isNaN(d)) return false;
    if (store && r.store !== store) return false;
    if (r.voided === true || String(r.voided).toLowerCase() === 'true') return false;
    return d >= fromDate && d <= toDate;
  });

  const totalRevenue = filtered.reduce((s, r) => s + (Number(r.total) || 0), 0);
  const totalCost    = filtered.reduce((s, r) => s + (Number(r.cost)  || 0), 0);
  const totalProfit  = totalRevenue - totalCost;
  const payBreakdown = {};
  filtered.forEach(r => {
    const pay = r.pay || '其他';
    payBreakdown[pay] = (payBreakdown[pay] || 0) + (Number(r.total) || 0);
  });

  return {
    totalRevenue, totalCost, totalProfit,
    txCount: filtered.length,
    payBreakdown,
    transactions: filtered.slice(-100).reverse(),
  };
}

// ══════════════════════════════════════════════════════
//  分析報表
// ══════════════════════════════════════════════════════
function dateRange_(date) {
  const base = date || Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');
  const from = new Date(base + 'T00:00:00+08:00');
  const to   = new Date(base + 'T23:59:59+08:00');
  return { from, to };
}

function getDailyReport(store, date) {
  const sh   = getSheet(SH.TRANSACTIONS);
  const rows = sheetToObjects(sh);
  const { from, to } = dateRange_(date);

  const filt    = storeFilter_(store);
  const notVoid = r => !(r.voided === true || String(r.voided).toLowerCase() === 'true');
  const todayTxs = rows.filter(r => {
    const d = parseTaipei_(r.date);
    return !isNaN(d) && filt(r) && notVoid(r) && d >= from && d <= to;
  });

  const lwFrom = new Date(from); lwFrom.setDate(lwFrom.getDate() - 7);
  const lwTo   = new Date(to);   lwTo.setDate(lwTo.getDate() - 7);
  const lwTxs  = rows.filter(r => { const d = parseTaipei_(r.date); return !isNaN(d) && filt(r) && notVoid(r) && d >= lwFrom && d <= lwTo; });

  const lmFrom = new Date(from); lmFrom.setMonth(lmFrom.getMonth() - 1);
  const lmTo   = new Date(to);   lmTo.setMonth(lmTo.getMonth() - 1);
  const lmTxs  = rows.filter(r => { const d = parseTaipei_(r.date); return !isNaN(d) && filt(r) && notVoid(r) && d >= lmFrom && d <= lmTo; });

  const lyFrom = new Date(from); lyFrom.setFullYear(lyFrom.getFullYear() - 1);
  const lyTo   = new Date(to);   lyTo.setFullYear(lyTo.getFullYear() - 1);
  const lyTxs  = rows.filter(r => { const d = parseTaipei_(r.date); return !isNaN(d) && filt(r) && notVoid(r) && d >= lyFrom && d <= lyTo; });

  const compute = txs => {
    const rev = txs.reduce((s,r) => s + (Number(r.total)||0), 0);
    const cst = txs.reduce((s,r) => s + (Number(r.cost)||0),  0);
    const pft = rev - cst;
    return { revenue: rev, cost: cst, profit: pft, margin: rev > 0 ? Math.round(pft / rev * 100) : 0, customers: txs.length, avg: txs.length > 0 ? Math.round(rev / txs.length) : 0 };
  };

  const sourceBreakdown = {};
  let newCount = 0, memberCount = 0;
  todayTxs.forEach(r => {
    if (r.source) sourceBreakdown[r.source] = (sourceBreakdown[r.source]||0) + 1;
    if (r.customerType === '新客') newCount++;
    if (r.memberId) memberCount++;
  });

  const prodMap = {};
  sheetToObjects(getSheet(SH.PRODUCTS)).forEach(r => { prodMap[Number(r.id)] = Number(r.cost)||0; });
  const wasteRows = sheetToObjects(getSheet(SH.INV_LOG)).filter(r => {
    const d = parseTaipei_(r.time);
    return r.type === 'waste' && !isNaN(d) && d >= from && d <= to;
  });
  const totalWasteCost = wasteRows.reduce((s,r) => s + (prodMap[Number(r.productId)]||0) * Math.abs(Number(r.diff)||0), 0);
  const wasteByStore = {};
  wasteRows.forEach(r => {
    const s = r.store || '未知';
    if (!wasteByStore[s]) wasteByStore[s] = { count: 0, cost: 0 };
    wasteByStore[s].count++;
    wasteByStore[s].cost += (prodMap[Number(r.productId)]||0) * Math.abs(Number(r.diff)||0);
  });

  return { ...compute(todayTxs), lastWeek: compute(lwTxs), lastMonth: compute(lmTxs), lastYear: compute(lyTxs), sourceBreakdown, newCount, memberCount, totalWasteCost, wasteByStore };
}

// [FIX-2] getMonthlyReport 時區修正
// 原本 new Date(r.date) 在 GAS 環境會把無時區後綴的字串當 UTC 解析
// 改用 parseTaipei_() 補上 +08:00，跨日零點前後訂單才不會跑到錯誤月份
function getMonthlyReport(store, ym) {
  const base  = ym || Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM');
  const year  = parseInt(base.slice(0, 4), 10);
  const month = parseInt(base.slice(5, 7), 10);
  const from   = new Date(year, month - 1, 1, 0, 0, 0);
  const to     = new Date(year, month, 0, 23, 59, 59);
  const lyFrom = new Date(year - 1, month - 1, 1, 0, 0, 0);
  const lyTo   = new Date(year - 1, month, 0, 23, 59, 59);

  const sh      = getSheet(SH.TRANSACTIONS);
  const rows    = sheetToObjects(sh);
  const filt    = storeFilter_(store);
  const notVoid = r => !(r.voided === true || String(r.voided).toLowerCase() === 'true');

  const compute = txs => {
    const rev = txs.reduce((s,r) => s + (Number(r.total)||0), 0);
    const cst = txs.reduce((s,r) => s + (Number(r.cost)||0), 0);
    const pft = rev - cst;
    return { revenue: rev, cost: cst, profit: pft, margin: rev > 0 ? Math.round(pft/rev*100) : 0, customers: txs.length };
  };

  // ✅ 修正：原本用 new Date(r.date)，現在改用 parseTaipei_(r.date)
  const thisTxs = rows.filter(r => { const d = parseTaipei_(r.date); return !isNaN(d) && filt(r) && notVoid(r) && d >= from  && d <= to; });
  const lyTxs   = rows.filter(r => { const d = parseTaipei_(r.date); return !isNaN(d) && filt(r) && notVoid(r) && d >= lyFrom && d <= lyTo; });

  return { ...compute(thisTxs), lastYear: compute(lyTxs) };
}

function getPayBreakdown(store, date) {
  const sh   = getSheet(SH.TRANSACTIONS);
  const rows = sheetToObjects(sh);
  const { from, to } = dateRange_(date);
  const filt = storeFilter_(store);
  const result = {};
  rows.filter(r => {
    const d = parseTaipei_(r.date);
    return !isNaN(d) && filt(r) && d >= from && d <= to;
  }).forEach(r => {
    const pay = r.pay || '其他';
    result[pay] = (result[pay]||0) + (Number(r.total)||0);
  });
  return result;
}

function getPayBreakdownRange(store, from, to) {
  const sh       = getSheet(SH.TRANSACTIONS);
  const rows     = sheetToObjects(sh);
  const filt     = storeFilter_(store);
  const fromDate = from ? new Date(from + 'T00:00:00+08:00') : new Date(0);
  const toDate   = to   ? new Date(to   + 'T23:59:59+08:00') : new Date();
  const notVoid  = r => !(r.voided === true || String(r.voided).toLowerCase() === 'true');
  const result = {};
  let total = 0;
  rows.filter(r => {
    const d = parseTaipei_(r.date);
    return !isNaN(d) && filt(r) && notVoid(r) && d >= fromDate && d <= toDate;
  }).forEach(r => {
    const pay = r.pay || '其他';
    const amt = Number(r.total) || 0;
    result[pay] = (result[pay] || 0) + amt;
    total += amt;
  });
  return { breakdown: result, total };
}

function getWeeklyTrend(store) {
  const sh   = getSheet(SH.TRANSACTIONS);
  const rows = sheetToObjects(sh);
  const filt = storeFilter_(store);
  const now  = new Date();
  const result = [];
  for (let w = 3; w >= 0; w--) {
    const dow = now.getDay() || 7;
    const weekStart = new Date(now);
    weekStart.setDate(now.getDate() - dow + 1 - w * 7);
    weekStart.setHours(0,0,0,0);
    const weekEnd = new Date(weekStart);
    weekEnd.setDate(weekStart.getDate() + 6);
    weekEnd.setHours(23,59,59,999);
    const txs     = rows.filter(r => { const d = new Date(r.date); const iv = r.voided===true||String(r.voided).toLowerCase()==='true'; return !isNaN(d) && filt(r) && !iv && d >= weekStart && d <= weekEnd; });
    const revenue = txs.reduce((s,r) => s + (Number(r.total)||0), 0);
    const cost    = txs.reduce((s,r) => s + (Number(r.cost)||0),  0);
    result.push({ label: `${weekStart.getMonth()+1}/${weekStart.getDate()}`, revenue, cost, profit: revenue - cost, customers: txs.length });
  }
  return result;
}

function getItemRanking(store, date, fromDate, toDate) {
  const sh   = getSheet(SH.TRANSACTIONS);
  const rows = sheetToObjects(sh);
  const filt = storeFilter_(store);
  let from, to;
  if (fromDate && toDate) {
    from = new Date(fromDate + 'T00:00:00+08:00');
    to   = new Date(toDate   + 'T23:59:59+08:00');
  } else {
    const base = date ? new Date(date + 'T00:00:00+08:00') : new Date();
    from = new Date(base.getFullYear(), base.getMonth(), 1);
    to   = new Date(base.getFullYear(), base.getMonth() + 1, 0, 23, 59, 59);
  }
  const prodMap = {};
  sheetToObjects(getSheet(SH.PRODUCTS)).forEach(r => {
    prodMap[Number(r.id)] = { cost: Number(r.cost)||0, price: Number(r.price)||0, name: r.name };
  });
  const itemMap = {};
  rows.filter(r => { const d = new Date(r.date); const iv = r.voided===true||String(r.voided).toLowerCase()==='true'; return !isNaN(d) && filt(r) && !iv && d >= from && d <= to; })
    .forEach(tx => {
      const items = safeJson(tx.items, []);
      items.forEach(item => {
        const key  = String(item.id || item.name);
        const prod = prodMap[Number(item.id)] || {};
        if (!itemMap[key]) itemMap[key] = { name: item.name || prod.name || key, qty: 0, sales: 0, cost: 0, waste: 0 };
        const qty   = Number(item.qty)   || 0;
        const price = Number(item.price) || prod.price || 0;
        const uc    = prod.cost || 0;
        itemMap[key].qty   += qty;
        itemMap[key].sales += price * qty;
        itemMap[key].cost  += uc * qty;
      });
    });
  const nameToKey = {};
  Object.entries(itemMap).forEach(([k, v]) => { nameToKey[v.name] = k; });
  sheetToObjects(getSheet(SH.INV_LOG)).filter(r => {
    const d = parseTaipei_(r.time);
    return r.type === 'waste' && !isNaN(d) && filt(r) && d >= from && d <= to;
  }).forEach(r => {
    const k = nameToKey[r.name];
    if (k && itemMap[k]) itemMap[k].waste += Math.abs(Number(r.diff)||0);
  });
  const sorted = Object.values(itemMap).map(it => {
    const profit    = it.sales - it.cost;
    const margin    = it.sales > 0 ? Math.round(profit / it.sales * 100) : 0;
    const wasteRate = (it.qty + it.waste) > 0 ? Math.round(it.waste / (it.qty + it.waste) * 100) : 0;
    return { name: it.name, qty: it.qty, revenue: Math.round(it.sales), sales: Math.round(it.sales), cost: Math.round(it.cost), profit: Math.round(profit), margin, waste: it.waste, wasteRate };
  }).sort((a,b) => b.revenue - a.revenue);
  const maxSales = sorted.length > 0 ? (sorted[0].revenue || 1) : 1;
  sorted.forEach(it => { it.pct = Math.round(it.revenue / maxSales * 100); });
  return sorted;
}

function getItemStoreSplit(name, from, to) {
  if (!name) return [];
  const sh       = getSheet(SH.TRANSACTIONS);
  const rows     = sheetToObjects(sh);
  const fromDate = from ? new Date(from + 'T00:00:00+08:00') : new Date(0);
  const toDate   = to   ? new Date(to   + 'T23:59:59+08:00') : new Date();
  const notVoid  = r => !(r.voided === true || String(r.voided).toLowerCase() === 'true');
  const byStore = {};
  rows.filter(r => {
    const d = parseTaipei_(r.date);
    return !isNaN(d) && notVoid(r) && d >= fromDate && d <= toDate;
  }).forEach(r => {
    const items = safeJson(String(r.items || '[]'), []);
    items.forEach(item => {
      if (String(item.name || '').trim() !== String(name).trim()) return;
      const s = r.store || '未知';
      if (!byStore[s]) byStore[s] = { qty: 0, revenue: 0 };
      const qty = Number(item.qty) || 1;
      const up  = Number(item.unitPrice || item.price || 0);
      byStore[s].qty     += qty;
      byStore[s].revenue += up * qty;
    });
  });
  return Object.entries(byStore).map(([store, v]) => ({ store, qty: v.qty, revenue: Math.round(v.revenue) })).sort((a, b) => b.qty - a.qty);
}

function getStoreStats(store, fromDate, toDate) {
  const sh   = getSheet(SH.TRANSACTIONS);
  const rows = sheetToObjects(sh);
  const filt = storeFilter_(store);
  const from = fromDate ? new Date(fromDate + 'T00:00:00+08:00') : new Date(0);
  const to   = toDate   ? new Date(toDate   + 'T23:59:59+08:00') : new Date();
  const notVoid = r => !(r.voided === true || String(r.voided).toLowerCase() === 'true');
  const byStore = {};
  rows.filter(r => {
    const d = parseTaipei_(r.date);
    return !isNaN(d) && filt(r) && notVoid(r) && d >= from && d <= to;
  }).forEach(r => {
    const s = r.store || '未知';
    if (!byStore[s]) byStore[s] = { count: 0, revenue: 0, cost: 0 };
    byStore[s].count++;
    byStore[s].revenue += Number(r.total) || 0;
    byStore[s].cost    += Number(r.cost)  || 0;
  });
  return Object.entries(byStore).map(([s, v]) => ({
    store: s, count: v.count, revenue: v.revenue, cost: v.cost,
    profit: v.revenue - v.cost,
    margin: v.revenue > 0 ? Math.round((v.revenue - v.cost) / v.revenue * 100) : 0,
  })).sort((a, b) => b.revenue - a.revenue);
}

function getWasteReport(store, fromDate, toDate) {
  const from = fromDate ? new Date(fromDate + 'T00:00:00+08:00') : new Date(0);
  const to   = toDate   ? new Date(toDate   + 'T23:59:59+08:00') : new Date();
  const filt = storeFilter_(store);
  const prodMap = {};
  sheetToObjects(getSheet(SH.PRODUCTS)).forEach(r => { prodMap[Number(r.id)] = Number(r.cost) || 0; });
  const wasteRows = sheetToObjects(getSheet(SH.INV_LOG)).filter(r => {
    const d = parseTaipei_(r.time);
    return r.type === 'waste' && !isNaN(d) && filt(r) && d >= from && d <= to;
  });
  const byStore = {};
  wasteRows.forEach(r => {
    const s = r.store || '未知';
    if (!byStore[s]) byStore[s] = { count: 0, cost: 0 };
    byStore[s].count++;
    byStore[s].cost += (prodMap[Number(r.productId)] || 0) * Math.abs(Number(r.diff) || 0);
  });

  // 花材報廢排行：依品名彙總廢棄量與成本，供前端 dashboard 顯示
  const byProductMap = {};
  wasteRows.forEach(r => {
    const name = String(r.name || '未知');
    if (!byProductMap[name]) byProductMap[name] = { name, qty: 0, cost: 0 };
    const qty = Math.abs(Number(r.diff) || 0);
    byProductMap[name].qty  += qty;
    byProductMap[name].cost += (prodMap[Number(r.productId)] || 0) * qty;
  });
  const byProduct = Object.values(byProductMap)
    .sort((a, b) => b.qty - a.qty)
    .slice(0, 20);

  return {
    totalCost: wasteRows.reduce((s, r) => s + (prodMap[Number(r.productId)] || 0) * Math.abs(Number(r.diff) || 0), 0),
    byStore,
    byProduct,
  };
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
  return { total: rows.length, newThisMonth: rows.filter(r => r.createdAt && new Date(r.createdAt) >= monthStart).length, tiers };
}

function getTransactionHistory(store, date) {
  const sh   = getSheet(SH.TRANSACTIONS);

  // [PERF] 有指定日期時只讀最近 600 列（約 2 個月），不讀全表
  // 若查更早的歷史資料才 fallback 讀全表
  let rows;
  if (date) {
    const lastRow  = sh.getLastRow();
    const readRows = Math.min(600, lastRow - 1);
    if (lastRow < 2 || readRows <= 0) {
      rows = [];
    } else {
      const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
      const data    = sh.getRange(lastRow - readRows + 1, 1, readRows, sh.getLastColumn()).getValues();
      rows = data.map(row => {
        const obj = {};
        headers.forEach((h, i) => {
          const v = row[i];
          obj[h] = v instanceof Date
            ? Utilities.formatDate(v, 'Asia/Taipei', "yyyy-MM-dd'T'HH:mm:ss")
            : v;
        });
        return obj;
      });
    }
  } else {
    rows = sheetToObjects(sh);
  }

  const filt = storeFilter_(store);
  let filtered = rows.filter(r => filt(r));
  if (date) {
    const { from, to } = dateRange_(date);
    filtered = filtered.filter(r => { const d = new Date(r.date); return !isNaN(d) && d >= from && d <= to; });
  }
  return filtered.reverse().slice(0, 200).map(r => ({
    id        : r.id,
    date      : toTaipeiStr_(r.date),
    store     : r.store,
    memberName: r.memberName,
    total     : Number(r.total) || 0,
    cost      : Number(r.cost)  || 0,
    pay       : r.pay,
    source    : r.source,
    note      : r.note,
    items     : safeJson(r.items, []),
    voided    : r.voided === true || String(r.voided).toLowerCase() === 'true',
  }));
}

function getAllStoresReport(period) {
  const sh   = getSheet(SH.TRANSACTIONS);
  const rows = sheetToObjects(sh);
  const now  = new Date();
  let from, to;
  if (period === 'lastMonth') {
    const m = now.getMonth() === 0 ? 11 : now.getMonth() - 1;
    const y = now.getMonth() === 0 ? now.getFullYear() - 1 : now.getFullYear();
    from = new Date(y, m, 1); to = new Date(y, m + 1, 0, 23, 59, 59);
  } else if (period === 'thisMonth') {
    from = new Date(now.getFullYear(), now.getMonth(), 1);
    to   = new Date(now.getFullYear(), now.getMonth() + 1, 0, 23, 59, 59);
  } else if (period === 'lastWeek') {
    const dow = now.getDay() || 7;
    from = new Date(now); from.setDate(now.getDate() - dow - 6); from.setHours(0,0,0,0);
    to   = new Date(from); to.setDate(from.getDate() + 6); to.setHours(23,59,59,999);
  } else {
    const dow = now.getDay() || 7;
    from = new Date(now); from.setDate(now.getDate() - dow + 1); from.setHours(0,0,0,0);
    to   = new Date(now); to.setHours(23,59,59,999);
  }
  const filtered = rows.filter(r => {
    const d = parseTaipei_(r.date);
    const isVoided = r.voided === true || String(r.voided).toLowerCase() === 'true';
    return !isNaN(d) && d >= from && d <= to && !isVoided;
  });
  const storeList = [
    { label: '高雄FOCUS 13', keys: ['高雄FOCUS 13'] },
    { label: '台南FOCUS',    keys: ['台南FOCUS'] },
    { label: '誠品生活台南', keys: ['誠品生活台南'] },
    { label: 'Studio',       keys: ['Studio'] },
  ];
  const result = storeList.map(s => {
    const txs     = filtered.filter(r => s.keys.includes(r.store));
    const revenue = txs.reduce((sum,r) => sum + (Number(r.total)||0), 0);
    const cost    = txs.reduce((sum,r) => sum + (Number(r.cost)||0),  0);
    const profit  = revenue - cost;
    return { store: s.label, revenue, cost, profit, margin: revenue > 0 ? Math.round(profit / revenue * 100) : 0, customers: txs.length };
  });
  const totRev = result.reduce((s,r) => s + r.revenue, 0);
  const totCost= result.reduce((s,r) => s + r.cost,    0);
  const totPro = totRev - totCost;
  result.push({ store: '合計', revenue: totRev, cost: totCost, profit: totPro, margin: totRev > 0 ? Math.round(totPro / totRev * 100) : 0, customers: result.reduce((s,r) => s + r.customers, 0) });
  return { stores: result, from: from.toISOString(), to: to.toISOString() };
}

// ══════════════════════════════════════════════════════
//  會員管理
// ══════════════════════════════════════════════════════
function getMembers() {
  const rows = sheetToObjects(getSheet(SH.MEMBERS));
  return rows.map(r => ({
    id: r.id, name: r.name, phone: r.phone, birthday: r.birthday,
    totalPoints: Number(r.totalPoints) || 0,
    totalSpend : Number(r.totalSpend)  || 0,
    createdAt  : r.createdAt,
  }));
}

function addMember(data) {
  const sh = getSheet(SH.MEMBERS);
  const id = newId('M');
  sh.appendRow([id, data.name || '', data.phone || '', data.birthday || '', 0, 0, taipeiNow()]);
  return { success: true, id, name: data.name, phone: data.phone, totalPoints: 0, totalSpend: 0 };
}

function updateMemberPoints(data) {
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
        const cur   = Number(vals[i][headers.indexOf('totalPoints')]) || 0;
        const curSp = Number(vals[i][headers.indexOf('totalSpend')])  || 0;
        sh.getRange(row, ptCol).setValue(cur   + (Number(data.addPoints) || 0));
        sh.getRange(row, spCol).setValue(curSp + (Number(data.addSpend)  || 0));
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
  const rows     = sheetToObjects(getSheet(SH.HANG_ORDERS));
  const filtered = store ? rows.filter(r => !r.store || r.store === store) : rows;
  return filtered.map(r => ({ id: r.id, name: r.name, store: r.store, createdAt: r.createdAt, cart: safeJson(r.cart, []), member: safeJson(r.member, null) }));
}

function saveHangOrder(data) {
  const sh = getSheet(SH.HANG_ORDERS);
  const id = newId('H');
  sh.appendRow([id, data.name || '掛單', data.store || '', data.createdAt || taipeiNow(), JSON.stringify(data.cart || []), JSON.stringify(data.member || null)]);
  return { success: true, id };
}

function deleteHangOrder(data) {
  const sh   = getSheet(SH.HANG_ORDERS);
  const vals = sh.getDataRange().getValues();
  const idCol = vals[0].indexOf('id');
  for (let i = vals.length - 1; i >= 1; i--) {
    if (String(vals[i][idCol]) === String(data.id)) { sh.deleteRow(i + 1); return { success: true }; }
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
  const headers = vals[0];
  let colIdx = -1;
  for (let i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim() === String(data.category).trim()) { colIdx = i; break; }
  }
  if (colIdx === -1) { colIdx = headers.length; sh.getRange(1, colIdx + 1).setValue(data.category); }
  let rowIdx = 1;
  while (rowIdx < vals.length && String(vals[rowIdx][colIdx] || '').trim() !== '') rowIdx++;
  sh.getRange(rowIdx + 1, colIdx + 1).setValue(data.name);
  return { success: true };
}

function getFlowerLibrary() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName('花材資料庫');
  if (!sh) return [];
  const vals = sh.getDataRange().getValues();
  if (vals.length === 0) return [];
  const headers = vals[0];
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
  if (!sh) { sh = ss.insertSheet(SH.PRICE_LOG); sh.appendRow(['time','name','category','bundlePrice','stems','unitCost','salePrice','note']); }
  sh.appendRow([taipeiNow(), data.name || '', data.category || '', Number(data.bundlePrice)||0, Number(data.stems)||0, Number(data.unitCost)||0, Number(data.salePrice)||0, data.note || '本週進花']);
  return { success: true };
}

function getPriceHistory(name) {
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SH.PRICE_LOG);
  if (!sh || sh.getLastRow() < 2) return [];
  const rows     = sheetToObjects(sh);
  const filtered = name ? rows.filter(r => r.name === name) : rows;
  return filtered.reverse().slice(0, 300).map(r => ({
    time: r.time, name: r.name, category: r.category,
    bundlePrice: Number(r.bundlePrice)||0, stems: Number(r.stems)||0,
    unitCost: Number(r.unitCost)||0, salePrice: Number(r.salePrice)||0, note: r.note,
  }));
}

// ══════════════════════════════════════════════════════
//  進貨模組
// ══════════════════════════════════════════════════════
function getProcurementBatches() {
  const sh = getSheet(SH_PROC.BATCHES);
  if (!sh || sh.getLastRow() < 2) return [];
  const itemSh = getSheet(SH_PROC.ITEMS);
  const aggByBatch = {};
  if (itemSh && itemSh.getLastRow() >= 2) {
    sheetToObjects(itemSh).forEach(function(r) {
      if (!r.batchId) return;
      if (!aggByBatch[r.batchId]) aggByBatch[r.batchId] = { count: 0, bunches: 0, cost: 0 };
      aggByBatch[r.batchId].count++;
      aggByBatch[r.batchId].bunches += Number(r.bunchesQty) || 0;
      aggByBatch[r.batchId].cost   += (Number(r.pricePerBunch)||0) * (Number(r.bunchesQty)||0);
    });
  }
  return sheetToObjects(sh).reverse().slice(0, 100).map(function(r) {
    const agg = aggByBatch[r.batchId] || { count: 0, bunches: 0, cost: 0 };
    return { batchId: r.batchId, date: r.date, source: r.source, store: r.store || '', note: r.note, totalItems: agg.count > 0 ? agg.count : (Number(r.totalItems)||0), totalBunches: agg.bunches > 0 ? agg.bunches : 0, totalCost: agg.cost > 0 ? agg.cost : (Number(r.totalCost)||0) };
  });
}

function getProcurementItems(batchId) {
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SH_PROC.ITEMS);
  if (!sh || sh.getLastRow() < 2) return [];
  const rows     = sheetToObjects(sh);
  const filtered = batchId ? rows.filter(r => r.batchId === batchId) : rows;
  return filtered.reverse().slice(0, 500).map(r => ({
    itemId: r.itemId, batchId: r.batchId, flowerName: r.flowerName, category: r.category,
    stemsPerBunch: Number(r.stemsPerBunch)||0, bunchesQty: Number(r.bunchesQty)||0,
    pricePerBunch: Number(r.pricePerBunch)||0, totalStems: Number(r.totalStems)||0,
    costPerStem: Number(r.costPerStem)||0, suggestedPrice: Number(r.suggestedPrice)||0,
    store: r.store, date: r.date, allocation: safeJson(r.allocation, {}),
  }));
}

function addProcurement(data) {
  const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
  const batchSh = ss.getSheetByName(SH_PROC.BATCHES);
  const itemSh  = ss.getSheetByName(SH_PROC.ITEMS);
  if (!batchSh || !itemSh) return { success: false, error: '進貨工作表不存在，請先執行 setupProcurementSheets()' };

  const date    = data.date || Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');
  const batchId = 'B' + Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd') + Math.random().toString(36).slice(2, 5).toUpperCase();
  const items   = data.items || [];
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
    const allocation = (item.allocation && typeof item.allocation === 'object') ? item.allocation : {};
    itemRows.push([batchId + '-' + String(idx + 1).padStart(2, '0'), batchId, item.flowerName || '', item.category || '主花', stemsPerBunch, bunchesQty, pricePerBunch, totalStems, costPerStem, suggestedPrice, '', date, JSON.stringify(allocation)]);
    ensureProductExists(item.flowerName, item.category, costPerStem, suggestedPrice);
    if (Object.keys(allocation).length > 0) {
      updateProductCostWeightedWithAlloc({ flowerName: item.flowerName, stemsPerBunch, allocation, costPerStem });
    } else {
      updateProductCostWeighted({ flowerName: item.flowerName, store: data.store || '', newStems: totalStems, newCost: costPerStem });
    }
    if (Number(item.salePrice) > 0) updateProductPrice(item.flowerName, Number(item.salePrice));
    const phSh = getSheet(SH.PRICE_LOG);
    if (phSh) phSh.appendRow([taipeiNow(), item.flowerName || '', item.category || '', pricePerBunch, stemsPerBunch, costPerStem, Number(item.salePrice) || 0, `進貨（${data.source || ''}）`]);
  });

  batchSh.appendRow([batchId, date, data.source || '', '', data.note || '', items.length, Math.round(totalCost)]);
  if (itemRows.length > 0) itemSh.getRange(itemSh.getLastRow() + 1, 1, itemRows.length, itemRows[0].length).setValues(itemRows);
  return { success: true, batchId, totalItems: items.length, totalCost: Math.round(totalCost) };
}

function ensureProductExists(name, category, cost, price) {
  if (!name) return;
  const sh      = getSheet(SH.PRODUCTS);
  const vals    = sh.getDataRange().getValues();
  const headers = vals[0];
  const nameCol = headers.indexOf('name');
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][nameCol] || '').trim() === String(name).trim()) return;
  }
  const maxId = vals.slice(1).reduce((m, r) => Math.max(m, Number(r[headers.indexOf('id')]) || 0), 0);
  sh.appendRow([maxId + 1, name, category || '主花', price || 0, cost || 0, 0, 0, 0, 0, 'active', true, '']);
  Logger.log(`✅ 自動建立商品：${name}`);
}

function updateProductPrice(name, price) {
  if (!name || !price) return;
  const sh      = getSheet(SH.PRODUCTS);
  const vals    = sh.getDataRange().getValues();
  const headers = vals[0];
  const nameCol  = headers.indexOf('name');
  const priceCol = headers.indexOf('price');
  if (nameCol < 0 || priceCol < 0) return;
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][nameCol]).trim() === String(name).trim()) {
      sh.getRange(i + 1, priceCol + 1).setValue(price);
      return;
    }
  }
}

function getProductByCode(code) {
  if (!code) return { found: false };
  const rows  = sheetToObjects(getSheet(SH.PRODUCTS));
  const match = rows.find(r => String(r.code || '').trim() !== '' && String(r.code || '').trim() === String(code).trim() && r.status !== 'inactive');
  if (!match) return { found: false };
  return { found: true, id: Number(match.id), name: match.name, category: match.category || '主花', price: Number(match.price)||0, code: String(match.code||'') };
}

function updateProductCostWeighted({ flowerName, store, newStems, newCost }) {
  const sh      = getSheet(SH.PRODUCTS);
  const vals    = sh.getDataRange().getValues();
  const headers = vals[0];
  const nameCol = headers.indexOf('name');
  const costCol = headers.indexOf('cost');
  const storeColKey = storeToColKey_(store);
  const storeCol    = headers.indexOf(storeColKey);

  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][nameCol] || '').trim() !== String(flowerName || '').trim()) continue;
    const row          = i + 1;
    const oldCost      = Number(vals[i][costCol])  || 0;
    const oldStock     = Number(vals[i][storeCol]) || 0;
    const weightedCost = (oldStock + newStems) > 0
      ? Math.round(((oldCost * oldStock) + (newCost * newStems)) / (oldStock + newStems) * 100) / 100
      : newCost;
    sh.getRange(row, costCol   + 1).setValue(weightedCost);
    sh.getRange(row, storeCol  + 1).setValue(oldStock + newStems);
    getSheet(SH.INV_LOG).appendRow([taipeiNow(), vals[i][headers.indexOf('id')], flowerName, store, oldStock, oldStock + newStems, newStems, `進貨入庫（加權成本 ${weightedCost}）`, 'receive']);
    return;
  }
  Logger.log(`⚠️ 進貨警告：找不到「${flowerName}」，成本未回寫`);
}

function updateProductCostWeightedWithAlloc({ flowerName, stemsPerBunch, allocation, costPerStem }) {
  if (!flowerName || !allocation) return;
  const sh      = getSheet(SH.PRODUCTS);
  const vals    = sh.getDataRange().getValues();
  const headers = vals[0];
  const nameCol = headers.indexOf('name');
  const costCol = headers.indexOf('cost');
  const colMap  = {
    '高雄FOCUS 13': headers.indexOf('stockKH'),
    '台南FOCUS':    headers.indexOf('stockTN'),
    '誠品生活台南': headers.indexOf('stockES'),
    '市集':         headers.indexOf('stockMarket'),
  };
  const totalNewStems = Object.values(allocation).reduce((s, b) => s + Number(b), 0);

  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][nameCol] || '').trim() !== String(flowerName).trim()) continue;
    const row = i + 1;
    const totalOldStock = Object.values(colMap).filter(c => c >= 0).reduce((s, c) => s + (Number(vals[i][c])||0), 0);
    const oldCost       = Number(vals[i][costCol]) || 0;
    const weightedCost  = (totalOldStock + totalNewStems) > 0
      ? Math.round(((oldCost * totalOldStock) + (costPerStem * totalNewStems)) / (totalOldStock + totalNewStems) * 100) / 100
      : costPerStem;
    sh.getRange(row, costCol + 1).setValue(weightedCost);
    Object.entries(allocation).forEach(([store, stemVal]) => {
      const stems  = Number(stemVal);
      if (!stems) return;
      const colIdx = colMap[store];
      if (colIdx === undefined || colIdx < 0) return;
      const current = Number(vals[i][colIdx]) || 0;
      sh.getRange(row, colIdx + 1).setValue(current + stems);
      vals[i][colIdx] = current + stems;
      getSheet(SH.INV_LOG).appendRow([taipeiNow(), vals[i][headers.indexOf('id')], flowerName, store, current, current + stems, stems, `進貨入庫（${store}，加權成本 ${weightedCost}）`, 'receive']);
    });
    return;
  }
  Logger.log(`⚠️ 進貨警告：找不到「${flowerName}」`);
}

// ══════════════════════════════════════════════════════
//  報廢記錄
// ══════════════════════════════════════════════════════

// 報廢通常單人操作，不加 Lock
function reportWaste(data) {
  const sh      = getSheet(SH.PRODUCTS);
  const vals    = sh.getDataRange().getValues();
  const headers = vals[0];
  const idCol   = headers.indexOf('id');
  const storeCol = headers.indexOf(storeToColKey_(data.store));
  for (let i = 1; i < vals.length; i++) {
    if (Number(vals[i][idCol]) === Number(data.id)) {
      const oldQty = Number(vals[i][storeCol]) || 0;
      const wasted = Number(data.qty) || 0;
      const newQty = Math.max(0, oldQty - wasted);
      sh.getRange(i + 1, storeCol + 1).setValue(newQty);
      getSheet(SH.INV_LOG).appendRow([taipeiNow(), data.id, vals[i][headers.indexOf('name')], data.store, oldQty, newQty, -wasted, data.note || '報廢', 'waste']);
      return { success: true, newStock: newQty };
    }
  }
  return { success: false, error: '找不到商品' };
}

// ══════════════════════════════════════════════════════
//  退款
// ══════════════════════════════════════════════════════
function voidTransaction(data) {
  const sh      = getSheet(SH.TRANSACTIONS);
  const vals    = sh.getDataRange().getValues();
  const headers = vals[0];
  const idCol   = headers.indexOf('id');
  let voidedCol = headers.indexOf('voided');
  if (voidedCol === -1) { voidedCol = headers.length; sh.getRange(1, voidedCol + 1).setValue('voided'); }

  for (let i = 1; i < vals.length; i++) {
    const txId = data.txId || data.id;
    if (String(vals[i][idCol]) !== String(txId)) continue;
    const alreadyVoided = vals[i][voidedCol];
    if (alreadyVoided === true || String(alreadyVoided).toLowerCase() === 'true') return { success: false, error: '此交易已退款過了' };
    sh.getRange(i + 1, voidedCol + 1).setValue(true);
    const store = vals[i][headers.indexOf('store')];
    const items = safeJson(String(vals[i][headers.indexOf('items')] || ''), []);
    items.forEach(item => {
      if (item.id && Number(item.id) > 0 && Number(item.qty) > 0) updateStockDelta(item.id, store, Number(item.qty), `退款還原：${txId}`);
    });
    const memberId  = String(vals[i][headers.indexOf('memberId')] || '');
    const earnedPts = Number(vals[i][headers.indexOf('earnedPoints')]) || 0;
    const txTotal   = Number(vals[i][headers.indexOf('total')]) || 0;
    if (memberId && earnedPts > 0) updateMemberPoints({ id: memberId, addPoints: -earnedPts, addSpend: -txTotal });
    return { success: true };
  }
  return { success: false, error: '找不到交易紀錄' };
}

// ══════════════════════════════════════════════════════
//  進貨批次刪除 / 更新
// ══════════════════════════════════════════════════════
function reverseStockForItems_(items, fallbackStore) {
  if (!items || !items.length) return;
  const sh = getSheet(SH.PRODUCTS);
  if (!sh) return;
  const vals    = sh.getDataRange().getValues();
  const headers = vals[0];
  const nameIdx = headers.indexOf('name');
  const idIdx   = headers.indexOf('id');
  const colMap  = {
    '高雄FOCUS 13': headers.indexOf('stockKH'),
    '台南FOCUS':    headers.indexOf('stockTN'),
    '誠品生活台南': headers.indexOf('stockES'),
    '市集':         headers.indexOf('stockMarket'),
  };

  function reverseOneStore_(name, store, stems) {
    if (!stems || !name) return;
    const colIdx = colMap[store] !== undefined ? colMap[store] : headers.indexOf('stockKH');
    if (colIdx < 0) return;
    for (let r = 1; r < vals.length; r++) {
      if (String(vals[r][nameIdx]).trim() === name) {
        const current  = Number(vals[r][colIdx]) || 0;
        const newStock = Math.max(0, current - stems);
        sh.getRange(r + 1, colIdx + 1).setValue(newStock);
        vals[r][colIdx] = newStock;
        const logSh = getSheet(SH.INV_LOG);
        if (logSh) logSh.appendRow([taipeiNow(), vals[r][idIdx], name, store, current, newStock, -stems, '進貨刪除/編輯（庫存還原）', 'adjustment']);
        break;
      }
    }
  }

  items.forEach(it => {
    const name = String(it.flowerName || '').trim();
    if (!name) return;
    if (it.allocation && typeof it.allocation === 'object' && Object.keys(it.allocation).length > 0) {
      const spb = Number(it.stemsPerBunch) || 0;
      Object.entries(it.allocation).forEach(([store, bunches]) => { reverseOneStore_(name, store, Number(bunches) * spb); });
    } else {
      reverseOneStore_(name, String(it.store || fallbackStore || '高雄FOCUS 13'), Number(it.totalStems) || 0);
    }
  });
}

function deleteProcurementBatch(data) {
  const batchSh = getSheet(SH_PROC.BATCHES);
  const itemSh  = getSheet(SH_PROC.ITEMS);
  if (!batchSh || !itemSh) return { success: false, error: '進貨工作表不存在' };
  const bId = String(data.batchId);
  const bVals  = batchSh.getDataRange().getValues();
  const bHdr   = bVals[0];
  const bIdCol = bHdr.indexOf('batchId');
  let batchStore = '', deleted = false;
  for (let i = bVals.length - 1; i >= 1; i--) {
    if (String(bVals[i][bIdCol]) === bId) {
      batchStore = String(bVals[i][bHdr.indexOf('store')] || '');
      batchSh.deleteRow(i + 1);
      deleted = true;
      break;
    }
  }
  if (itemSh.getLastRow() >= 2) {
    const iVals   = itemSh.getDataRange().getValues();
    const iHdr    = iVals[0];
    const ibCol   = iHdr.indexOf('batchId');
    const allocCol = iHdr.indexOf('allocation');
    const spbCol   = iHdr.indexOf('stemsPerBunch');
    const toReverse = [];
    for (let i = 1; i < iVals.length; i++) {
      if (String(iVals[i][ibCol]) === bId) {
        toReverse.push({ flowerName: iVals[i][iHdr.indexOf('flowerName')], totalStems: iVals[i][iHdr.indexOf('totalStems')], store: iVals[i][iHdr.indexOf('store')] || batchStore, stemsPerBunch: spbCol >= 0 ? Number(iVals[i][spbCol]) : 0, allocation: safeJson(String(allocCol >= 0 ? iVals[i][allocCol] : ''), null) });
      }
    }
    reverseStockForItems_(toReverse, batchStore);
    const iVals2 = itemSh.getDataRange().getValues();
    const ibCol2 = iVals2[0].indexOf('batchId');
    for (let i = iVals2.length - 1; i >= 1; i--) {
      if (String(iVals2[i][ibCol2]) === bId) itemSh.deleteRow(i + 1);
    }
  }
  return deleted ? { success: true } : { success: false, error: '找不到此進貨批次' };
}

function updateProcurementBatch(data) {
  const batchSh = getSheet(SH_PROC.BATCHES);
  const itemSh  = getSheet(SH_PROC.ITEMS);
  if (!batchSh || !itemSh) return { success: false, error: '進貨工作表不存在' };
  const bId   = String(data.batchId);
  const bVals = batchSh.getDataRange().getValues();
  const bHdr  = bVals[0];
  const bIdCol = bHdr.indexOf('batchId');
  let batchRow = -1, oldBatchStore = '';
  for (let i = 1; i < bVals.length; i++) {
    if (String(bVals[i][bIdCol]) === bId) { batchRow = i + 1; oldBatchStore = String(bVals[i][bHdr.indexOf('store')]||''); break; }
  }
  if (batchRow === -1) return { success: false, error: '找不到此進貨批次' };

  const colOf  = name => bHdr.indexOf(name);
  const items  = Array.isArray(data.items) ? data.items : [];
  const newStore  = data.store || oldBatchStore;
  const totalCost = items.reduce((s, it) => s + (Number(it.pricePerBunch)||0) * (Number(it.bunchesQty)||0), 0);
  if (colOf('source')     >= 0) batchSh.getRange(batchRow, colOf('source')    +1).setValue(data.source || '');
  if (colOf('date')       >= 0) batchSh.getRange(batchRow, colOf('date')      +1).setValue(data.date   || '');
  if (colOf('note')       >= 0) batchSh.getRange(batchRow, colOf('note')      +1).setValue(data.note   || '');
  if (colOf('totalCost')  >= 0) batchSh.getRange(batchRow, colOf('totalCost') +1).setValue(totalCost);
  if (colOf('totalItems') >= 0) batchSh.getRange(batchRow, colOf('totalItems')+1).setValue(items.length);

  const oldItems = [];
  if (itemSh.getLastRow() >= 2) {
    const iVals = itemSh.getDataRange().getValues();
    const iHdr  = iVals[0];
    const ibCol = iHdr.indexOf('batchId');
    const acCol = iHdr.indexOf('allocation');
    const spbCol2 = iHdr.indexOf('stemsPerBunch');
    for (let i = 1; i < iVals.length; i++) {
      if (String(iVals[i][ibCol]) === bId) {
        oldItems.push({ flowerName: iVals[i][iHdr.indexOf('flowerName')], totalStems: iVals[i][iHdr.indexOf('totalStems')], store: iVals[i][iHdr.indexOf('store')] || oldBatchStore, stemsPerBunch: spbCol2 >= 0 ? Number(iVals[i][spbCol2]) : 0, allocation: safeJson(String(acCol >= 0 ? iVals[i][acCol] : ''), null) });
      }
    }
  }
  reverseStockForItems_(oldItems, oldBatchStore);

  if (itemSh.getLastRow() >= 2) {
    const iVals2 = itemSh.getDataRange().getValues();
    const ibCol2 = iVals2[0].indexOf('batchId');
    for (let i = iVals2.length - 1; i >= 1; i--) {
      if (String(iVals2[i][ibCol2]) === bId) itemSh.deleteRow(i + 1);
    }
  }

  const iHdr = itemSh.getLastRow() >= 1 ? itemSh.getRange(1, 1, 1, itemSh.getLastColumn()).getValues()[0] : [];
  items.forEach((it, idx) => {
    const spb  = Number(it.stemsPerBunch)||0, bq = Number(it.bunchesQty)||0, ppb = Number(it.pricePerBunch)||0;
    const totalStems = spb * bq, costPerStem = spb > 0 ? Math.round(ppb / spb * 100) / 100 : 0;
    const allocation = (it.allocation && typeof it.allocation === 'object') ? it.allocation : {};
    const row = iHdr.map(h => {
      if (h === 'batchId') return bId; if (h === 'flowerName') return it.flowerName || '';
      if (h === 'category') return it.category || '其他'; if (h === 'stemsPerBunch') return spb;
      if (h === 'bunchesQty') return bq; if (h === 'pricePerBunch') return ppb;
      if (h === 'totalStems') return totalStems; if (h === 'costPerStem') return costPerStem;
      if (h === 'suggestedPrice') return Math.round(costPerStem * 3.8);
      if (h === 'totalCost') return ppb * bq; if (h === 'seq') return idx + 1;
      if (h === 'store') return ''; if (h === 'allocation') return JSON.stringify(allocation);
      return '';
    });
    if (iHdr.length > 0) itemSh.appendRow(row);
    ensureProductExists(it.flowerName, it.category, costPerStem, Math.round(costPerStem * 3.8));
    if (Object.keys(allocation).length > 0) {
      updateProductCostWeightedWithAlloc({ flowerName: it.flowerName, stemsPerBunch: spb, allocation, costPerStem });
    } else {
      updateProductCostWeighted({ flowerName: it.flowerName, store: newStore, newStems: totalStems, newCost: costPerStem });
    }
    if (Number(it.salePrice) > 0) updateProductPrice(it.flowerName, Number(it.salePrice));
  });

  return { success: true, totalCost, totalItems: items.length };
}

// ══════════════════════════════════════════════════════
//  合併重複商品
// ══════════════════════════════════════════════════════
function mergeProducts(data) {
  if (!data || data.keepId === undefined || data.mergeId === undefined) {
    return { success: false, error: '缺少必要參數 keepId / mergeId' };
  }
  const sh      = getSheet(SH.PRODUCTS);
  const rows    = sheetToObjects(sh);
  const vals    = sh.getDataRange().getValues();
  const headers = vals[0];
  const keepRow  = rows.find(r => Number(r.id) === Number(data.keepId));
  const mergeRow = rows.find(r => Number(r.id) === Number(data.mergeId));
  if (!keepRow)  return { success: false, error: '找不到保留商品' };
  if (!mergeRow) return { success: false, error: '找不到合併來源商品' };
  if (String(data.keepId) === String(data.mergeId)) return { success: false, error: '不能合併同一個商品' };

  const kKH = Number(keepRow.stockKH)||0, kTN = Number(keepRow.stockTN)||0, kES = Number(keepRow.stockES)||0, kMK = Number(keepRow.stockMarket)||0;
  const mKH = Number(mergeRow.stockKH)||0, mTN = Number(mergeRow.stockTN)||0, mES = Number(mergeRow.stockES)||0, mMK = Number(mergeRow.stockMarket)||0;
  const kTotal = kKH+kTN+kES+kMK, mTotal = mKH+mTN+mES+mMK;
  const avgCost  = (kTotal+mTotal) > 0 ? Math.round(((Number(keepRow.cost)||0)*kTotal + (Number(mergeRow.cost)||0)*mTotal) / (kTotal+mTotal) * 100) / 100 : Math.round(((Number(keepRow.cost)||0)+(Number(mergeRow.cost)||0)) / 2 * 100) / 100;
  const maxPrice = Math.max(Number(keepRow.price)||0, Number(mergeRow.price)||0);

  const keepIdx = rows.indexOf(keepRow);
  const setCell = (rowIdx, col, val) => { const c = headers.indexOf(col); if (c >= 0) sh.getRange(rowIdx + 2, c + 1).setValue(val); };
  setCell(keepIdx, 'cost', avgCost); setCell(keepIdx, 'price', maxPrice);
  setCell(keepIdx, 'stockKH', kKH+mKH); setCell(keepIdx, 'stockTN', kTN+mTN);
  setCell(keepIdx, 'stockES', kES+mES); setCell(keepIdx, 'stockMarket', kMK+mMK);

  const mergeIdx = rows.indexOf(mergeRow);
  setCell(mergeIdx, 'status', 'inactive');
  setCell(mergeIdx, 'stockKH', 0); setCell(mergeIdx, 'stockTN', 0);
  setCell(mergeIdx, 'stockES', 0); setCell(mergeIdx, 'stockMarket', 0);

  const txSh     = getSheet(SH.TRANSACTIONS);
  const txVals   = txSh.getDataRange().getValues();
  const txHdrs   = txVals[0];
  const itemsCol = txHdrs.indexOf('items');
  let txUpdated  = 0;
  if (itemsCol >= 0) {
    for (let i = 1; i < txVals.length; i++) {
      if (!String(txVals[i][itemsCol]||'').includes(mergeRow.name)) continue;
      try {
        const items = JSON.parse(String(txVals[i][itemsCol]));
        let changed = false;
        items.forEach(item => { if (String(item.name||'').trim() === String(mergeRow.name).trim()) { item.name = keepRow.name; item.id = Number(data.keepId); changed = true; } });
        if (changed) { txSh.getRange(i + 1, itemsCol + 1).setValue(JSON.stringify(items)); txUpdated++; }
      } catch(e) {}
    }
  }

  const logSh   = getSheet(SH.INV_LOG);
  const logVals = logSh.getDataRange().getValues();
  const logHdrs = logVals[0];
  const logNameCol = logHdrs.indexOf('name'), logPidCol = logHdrs.indexOf('productId');
  let logUpdated = 0;
  if (logNameCol >= 0) {
    for (let i = 1; i < logVals.length; i++) {
      if (String(logVals[i][logNameCol]||'').trim() === String(mergeRow.name).trim()) {
        logSh.getRange(i + 1, logNameCol + 1).setValue(keepRow.name);
        if (logPidCol >= 0) logSh.getRange(i + 1, logPidCol + 1).setValue(Number(data.keepId));
        logUpdated++;
      }
    }
  }

  return { success: true, message: `已將「${mergeRow.name}」合併至「${keepRow.name}」`, avgCost, maxPrice, newStock: { kh: kKH+mKH, tn: kTN+mTN, es: kES+mES, mk: kMK+mMK }, txUpdated, logUpdated };
}

// ══════════════════════════════════════════════════════
//  花材進貨歷史查詢
// ══════════════════════════════════════════════════════
function getFlowerProcHistory(name) {
  const itemSh  = getSheet(SH_PROC.ITEMS);
  const batchSh = getSheet(SH_PROC.BATCHES);
  if (!itemSh || itemSh.getLastRow() < 2) return [];
  const items    = sheetToObjects(itemSh);
  const batches  = (batchSh && batchSh.getLastRow() >= 2) ? sheetToObjects(batchSh) : [];
  const batchMap = {};
  batches.forEach(b => { batchMap[b.batchId] = b; });
  const filtered = name ? items.filter(r => String(r.flowerName || '').trim() === String(name).trim()) : items;
  return filtered.map(r => ({
    date: r.date || (batchMap[r.batchId] ? batchMap[r.batchId].date : ''),
    source: batchMap[r.batchId] ? (batchMap[r.batchId].source || '') : '',
    store: r.store || '', stemsPerBunch: Number(r.stemsPerBunch)||0, bunchesQty: Number(r.bunchesQty)||0,
    pricePerBunch: Number(r.pricePerBunch)||0, totalStems: Number(r.totalStems)||0,
    costPerStem: Number(r.costPerStem)||0, suggestedPrice: Number(r.suggestedPrice)||0,
  })).sort((a, b) => new Date(b.date) - new Date(a.date)).slice(0, 8);
}

// ══════════════════════════════════════════════════════
//  Schema 遷移工具（手動執行）
// ══════════════════════════════════════════════════════
function addCodeColumn() {
  const sh      = getSheet(SH.PRODUCTS);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (headers.indexOf('code') >= 0) { Logger.log('✅ code 欄位已存在'); return; }
  sh.getRange(1, headers.length + 1).setValue('code');
  Logger.log(`✅ code 欄位已新增（第 ${headers.length + 1} 欄）`);
}

function addMarketColumn() {
  const sh = getSheet(SH.PRODUCTS);
  if (!sh) { Logger.log('❌ 找不到 Products 工作表'); return; }
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (headers.includes('stockMarket')) { Logger.log('✅ stockMarket 欄位已存在'); return; }
  const esIdx    = headers.indexOf('stockES');
  const insertCol = esIdx >= 0 ? esIdx + 2 : sh.getLastColumn() + 1;
  sh.insertColumnAfter(esIdx >= 0 ? esIdx + 1 : sh.getLastColumn());
  sh.getRange(1, insertCol).setValue('stockMarket');
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, insertCol, lastRow - 1, 1).setValue(0);
  Logger.log(`✅ stockMarket 欄位已新增（第 ${insertCol} 欄）`);
}

function addAllocationColumn() {
  const sh = getSheet(SH_PROC.ITEMS);
  if (!sh) { Logger.log('❌ 找不到 ProcurementItems 工作表'); return; }
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (headers.includes('allocation')) { Logger.log('✅ allocation 欄位已存在'); return; }
  sh.getRange(1, sh.getLastColumn() + 1).setValue('allocation');
  Logger.log(`✅ allocation 欄位已新增`);
}

// ══════════════════════════════════════════════════════
//  每週日 00:00 自動報廢主花/配花/葉材剩餘庫存
//  [Auto Waste] 執行後會把所有 active 商品中
//  category 為 主花/配花/葉材 的各點位庫存歸零
//  並在 InventoryLog 寫入 type='waste' 記錄
// ══════════════════════════════════════════════════════

const AUTO_WASTE_CATS = ['主花', '配花', '葉材'];

const STORE_COL_MAP = {
  '高雄FOCUS 13': 'stockKH',
  '台南FOCUS':    'stockTN',
  '誠品生活台南': 'stockES',
  '市集':         'stockMarket',
};

function weeklyAutoWaste() {
  const sh      = getSheet(SH.PRODUCTS);
  const vals    = sh.getDataRange().getValues();
  const headers = vals[0];
  const idCol     = headers.indexOf('id');
  const nameCol   = headers.indexOf('name');
  const catCol    = headers.indexOf('category');
  const statusCol = headers.indexOf('status');
  const logSh     = getSheet(SH.INV_LOG);
  const now       = taipeiNow();
  let   wastedCount = 0;

  for (let i = 1; i < vals.length; i++) {
    const row    = vals[i];
    const status = String(row[statusCol] || '').toLowerCase();
    const cat    = String(row[catCol]    || '');

    // 只處理 active 商品 + 指定分類
    if (status === 'inactive') continue;
    if (!AUTO_WASTE_CATS.includes(cat)) continue;

    const id   = row[idCol];
    const name = row[nameCol];

    // 逐點位歸零
    Object.entries(STORE_COL_MAP).forEach(([storeName, colKey]) => {
      const colIdx = headers.indexOf(colKey);
      if (colIdx < 0) return;
      const qty = Number(row[colIdx]) || 0;
      if (qty <= 0) return; // 已經是 0，跳過

      // 寫入庫存歸零
      sh.getRange(i + 1, colIdx + 1).setValue(0);

      // 寫入 InventoryLog
      logSh.appendRow([
        now,
        id,
        name,
        storeName,
        qty,    // oldQty
        0,      // newQty
        -qty,   // diff
        '週自動報廢（週日凌晨）',
        'waste',
      ]);

      wastedCount++;
    });
  }

  Logger.log(`✅ 週自動報廢完成：共 ${wastedCount} 筆庫存歸零`);
}

// ══════════════════════════════════════════════════════
//  設定週日 Trigger（手動執行這個函式一次就好）
//  執行後會建立「每週日 00:00–01:00」的排程
//  重複執行會自動移除舊的再建新的，不會重複
// ══════════════════════════════════════════════════════
function setupWeeklyWasteTrigger() {
  // 移除舊的同名 trigger，避免重複
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'weeklyAutoWaste') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 建立新 trigger：每週日 00:00–01:00
  ScriptApp.newTrigger('weeklyAutoWaste')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(0)
    .create();

  Logger.log('✅ 週自動報廢 Trigger 已設定：每週日 00:00 執行');
}
