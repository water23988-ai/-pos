// ══════════════════════════════════════════════════════
//  拾逅花事 | 進銷存系統 — Google Apps Script 後端
//  版本：2.0  |  2026-04
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

const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE'; // ← 填入你的試算表 ID

// ── 工作表名稱 ──
const SH = {
  PRODUCTS    : 'Products',
  TRANSACTIONS: 'Transactions',
  MEMBERS     : 'Members',
  INV_LOG     : 'InventoryLog',
  HANG_ORDERS : 'HangOrders',
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

  ensureSheet(SH.PRODUCTS,     ['id','name','category','price','cost','stockKH','stockTN','stockES']);
  ensureSheet(SH.TRANSACTIONS, ['id','date','store','memberId','memberName','subtotal','discount','total','cost','pay','earnedPoints','note','customerType','source','items']);
  ensureSheet(SH.MEMBERS,      ['id','name','phone','birthday','totalPoints','totalSpend','createdAt']);
  ensureSheet(SH.INV_LOG,      ['time','productId','name','store','oldQty','newQty','diff','note','type']);
  ensureSheet(SH.HANG_ORDERS,  ['id','name','store','createdAt','cart','member']);

  Logger.log('✅ 所有工作表初始化完成');
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
      default:                      return jsonOut({ error: 'Unknown GET action: ' + action });
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
      case 'deleteProduct': return jsonOut(deleteProduct(data));
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
      case 'deleteHangOrder': return jsonOut(deleteHangOrder(data));
      default: return jsonOut({ error: 'Unknown POST action: ' + action });
    }
  } catch (err) {
    return jsonOut({ error: err.message });
  }
}

// ══════════════════════════════════════════════════════
//  工具函式
// ══════════════════════════════════════════════════════
function getSheet(name) {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name);
}

function sheetToObjects(sh) {
  if (sh.getLastRow() < 2) return [];
  const [headers, ...rows] = sh.getDataRange().getValues();
  return rows.map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
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
  return rows.map(r => ({
    id      : Number(r.id),
    name    : r.name,
    category: r.category || '其他加購',
    price   : Number(r.price)   || 0,
    cost    : Number(r.cost)    || 0,
    stockKH : Number(r.stockKH) || 0,
    stockTN : Number(r.stockTN) || 0,
    stockES : Number(r.stockES) || 0,
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
      return { success: true };
    }
  }
  return { success: false, error: '找不到商品' };
}

function deleteProduct(data) {
  const sh   = getSheet(SH.PRODUCTS);
  const vals = sh.getDataRange().getValues();
  const idCol = vals[0].indexOf('id');

  for (let i = vals.length - 1; i >= 1; i--) {
    if (Number(vals[i][idCol]) === Number(data.id)) {
      sh.deleteRow(i + 1);
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

function getInventoryLog(store) {
  const sh   = getSheet(SH.INV_LOG);
  const rows = sheetToObjects(sh);
  const filtered = store
    ? rows.filter(r => !r.store || r.store === store)
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
  // data: { id, date, store, memberId, memberName, subtotal, discount, total, cost, pay, earnedPoints, note, customerType, source, items }
  const sh = getSheet(SH.TRANSACTIONS);
  sh.appendRow([
    data.id           || newId('T'),
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
  // 高雄 includes Studio orders stored under the same store name
  if (store === '高雄FOCUS 13') return r => r.store === '高雄FOCUS 13' || r.store === 'Studio';
  if (store) return r => r.store === store;
  return () => true;
}

function getDailyReport(store, date) {
  const sh   = getSheet(SH.TRANSACTIONS);
  const rows = sheetToObjects(sh);
  const { from, to } = dateRange_(date);

  const filt = storeFilter_(store);
  const todayTxs = rows.filter(r => {
    const d = new Date(r.date);
    return !isNaN(d) && filt(r) && d >= from && d <= to;
  });

  // Same weekday last week
  const lwFrom = new Date(from); lwFrom.setDate(lwFrom.getDate() - 7);
  const lwTo   = new Date(to);   lwTo.setDate(lwTo.getDate() - 7);
  const lwTxs  = rows.filter(r => {
    const d = new Date(r.date);
    return !isNaN(d) && filt(r) && d >= lwFrom && d <= lwTo;
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

    const txs      = rows.filter(r => { const d = new Date(r.date); return !isNaN(d) && filt(r) && d >= weekStart && d <= weekEnd; });
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
  rows.filter(r => { const d = new Date(r.date); return !isNaN(d) && filt(r) && d >= from && d <= to; })
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

  const filtered = rows.filter(r => { const d = new Date(r.date); return !isNaN(d) && d >= from && d <= to; });

  const storeList = [
    { label: '高雄FOCUS 13', keys: ['高雄FOCUS 13', 'Studio'] },
    { label: '台南FOCUS',    keys: ['台南FOCUS'] },
    { label: '誠品生活台南', keys: ['誠品生活台南'] },
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
//  小工具
// ══════════════════════════════════════════════════════
function safeJson(str, fallback) {
  try   { return JSON.parse(str); }
  catch { return fallback; }
}
