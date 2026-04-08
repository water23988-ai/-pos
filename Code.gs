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
      case 'getSalesReport':  return jsonOut(getSalesReport(e.parameter.store, e.parameter.from, e.parameter.to));
      default:                return jsonOut({ error: 'Unknown GET action: ' + action });
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
        updateStock({
          id   : item.id,
          store: data.store,
          qty  : null, // 將在下方計算
          note : `銷售扣庫：${data.id}`,
          _delta: -item.qty, // 使用 delta 模式
        });
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
