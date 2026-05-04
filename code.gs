const SHEET_ID = "1thL_dReEZGlpyW0VJztV1V8oJG7eiRQ1oIOjSDINWc8";

const SHEETS = {
  STOCK: "Stock",
  USERS: "Users",
  HISTORY: "History",
  ORDERS: "Orders"
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("SWA Meat and Fish Market")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSS() {
  return SpreadsheetApp.openById(SHEET_ID);
}

// --- AUTH ---
function checkLogin(u, p) {
  const data = getSS().getSheetByName(SHEETS.USERS).getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const username = String(data[i][0] || "");
    const password = String(data[i][1] || "");
    const role = String(data[i][2] || "").toLowerCase();
    const status = String(data[i][3] || "").toLowerCase();

    if (username === u && password === p) {
      if (status === "active") {
        return { success: true, username, role };
      }
      return { success: false, blocked: true, message: "You have been blocked by the admin." };
    }
  }

  return { success: false, message: "Invalid login" };
}

function getUserSession(username) {
  const data = getSS().getSheetByName(SHEETS.USERS).getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const rowUser = String(data[i][0] || "");
    const role = String(data[i][2] || "").toLowerCase();
    const status = String(data[i][3] || "").toLowerCase();

    if (rowUser === username) {
      return { found: true, username: rowUser, role, status, active: status === "active" };
    }
  }

  return { found: false, active: false };
}

// --- STOCK ---
function getDashboardData() {
  const stock = getSS().getSheetByName(SHEETS.STOCK).getDataRange().getValues();
  const tz = Session.getScriptTimeZone();

  return {
    items: stock.slice(1).map(r => ({
      itemId: String(r[0]),
      itemName: String(r[1]),
      category: String(r[2]),
      qty: Number(r[3]),
      lowStockThreshold: Number(r[10] || 5), // K column
      updatedAt: r[6] instanceof Date
        ? Utilities.formatDate(r[6], tz, 'dd MMM yyyy, HH:mm')
        : String(r[6] || ""),
      note: String(r[7] || ""),
      updatedBy: String(r[8] || "")
    })).filter(i => i.itemId)
  };
}

function updateStockAction(itemId, changeQty, notes, username) {
  const ss = getSS();
  const sheet = ss.getSheetByName(SHEETS.STOCK);
  const data = sheet.getDataRange().getValues();
  const targetId = String(itemId).trim();

  for (let i = 1; i < data.length; i++) {
    const sheetId = String(data[i][0]).trim();

    if (sheetId === targetId) {
      const prev = Number(data[i][3]);
      const newQty = prev + Number(changeQty);
      const name = data[i][1];

      sheet.getRange(i + 1, 4).setValue(newQty);
      sheet.getRange(i + 1, 7).setValue(new Date());
      sheet.getRange(i + 1, 8).setValue(notes);
      sheet.getRange(i + 1, 9).setValue(username);

      const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
      const now = new Date();

      ss.getSheetByName(SHEETS.HISTORY).appendRow([
        now,
        days[now.getDay()],
        name,
        prev,
        newQty,
        changeQty > 0 ? `+${changeQty}` : changeQty,
        username,
        notes
      ]);

      return { success: true };
    }
  }

  throw new Error("ID not found: " + targetId);
}

// --- HISTORY ---
function getHistoryData() {
  const sheet = getSS().getSheetByName(SHEETS.HISTORY);
  if (!sheet || sheet.getLastRow() <= 1) return { history: [] };

  return {
    history: sheet.getDataRange().getValues().slice(1).map(r => ({
      time: r[0] instanceof Date ? r[0].toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }) : "00:00",
      day: String(r[1] || ""),
      item: String(r[2] || ""),
      change: String(r[5] || "0"),
      user: String(r[6] || ""),
      notes: String(r[7] || "")
    })).reverse()
  };
}

// --- USERS ---
function getUsersData() {
  const sheet = getSS().getSheetByName(SHEETS.USERS);

  return {
    users: sheet.getDataRange().getValues().slice(1).map(r => ({
      username: String(r[0] || ""),
      password: String(r[1] || ""),
      role: String(r[2] || ""),
      status: String(r[3] || "")
    })).filter(u => u.username)
  };
}

function saveUser(oldUsername, userObj) {
  const sheet = getSS().getSheetByName(SHEETS.USERS);
  const data = sheet.getDataRange().getValues();

  if (oldUsername) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === oldUsername) {
        sheet.getRange(i + 1, 1, 1, 4).setValues([[
          userObj.username,
          userObj.password,
          userObj.role,
          userObj.status
        ]]);
        return { success: true };
      }
    }
  } else {
    sheet.appendRow([userObj.username, userObj.password, userObj.role, userObj.status, new Date()]);
    return { success: true };
  }
}

function getOrdersSheet_() {
  const ss = getSS();
  let sheet = ss.getSheetByName(SHEETS.ORDERS);
  if (!sheet) sheet = ss.insertSheet(SHEETS.ORDERS);

  const headers = [
    "Order ID",
    "Created At",
    "Item ID",
    "Item Name",
    "Prev Qty",
    "Ordered Qty",
    "Received Qty",
    "Expected Total",
    "Supplier",
    "Status",
    "Confirmed By",
    "Confirmed At",
    "Created By",
    "Notes"
  ];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    return sheet;
  }

  const firstRow = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length)).getValues()[0];

  if (String(firstRow[12] || "") !== "Created By") {
    if (String(firstRow[12] || "") === "Notes") {
      sheet.insertColumnBefore(13);
    } else if (sheet.getLastColumn() < headers.length) {
      while (sheet.getLastColumn() < headers.length) sheet.insertColumnAfter(sheet.getLastColumn());
    }
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  return sheet;
}

function getOrdersData() {
  const sheet = getOrdersSheet_();
  if (sheet.getLastRow() <= 1) return { orders: [], suppliers: [] };

  const tz = Session.getScriptTimeZone();
  const rows = sheet.getDataRange().getValues().slice(1);

  const orders = rows.map(r => ({
    orderId: String(r[0] || ""),
    createdAt: r[1] instanceof Date ? Utilities.formatDate(r[1], tz, 'dd MMM yyyy, HH:mm') : String(r[1] || ""),
    itemId: String(r[2] || ""),
    itemName: String(r[3] || ""),
    prevQty: Number(r[4] || 0),
    orderQty: Number(r[5] || 0),
    receivedQty: r[6] === "" ? "" : Number(r[6] || 0),
    expectedTotal: Number(r[7] || 0),
    supplier: String(r[8] || ""),
    status: String(r[9] || "pending"),
    confirmedBy: String(r[10] || ""),
    confirmedAt: r[11] instanceof Date ? Utilities.formatDate(r[11], tz, 'dd MMM yyyy, HH:mm') : String(r[11] || ""),
    createdBy: String(r[12] || ""),
    notes: String(r[13] || "")
  })).reverse();

  const suppliers = [...new Set(
    orders.map(o => String(o.supplier || "").trim()).filter(Boolean)
  )].sort();

  return { orders, suppliers };
}

function createUniversalOrder(payload, username) {
  const ss = getSS();
  const stockSheet = ss.getSheetByName(SHEETS.STOCK);
  const ordersSheet = getOrdersSheet_();
  const stockData = stockSheet.getDataRange().getValues();

  const supplier = String(payload && payload.supplier ? payload.supplier : "").trim();
  const items = Array.isArray(payload && payload.items) ? payload.items : [];
  const createdBy = String(username || "").trim();

  if (!supplier) throw new Error("Supplier is required");

  const selected = items
    .map(it => ({
      itemId: String(it.itemId || "").trim(),
      itemName: String(it.itemName || "").trim(),
      orderQty: Number(it.orderQty || 0),
      note: String(it.note || "").trim()
    }))
    .filter(it => it.itemId && it.orderQty > 0);

  if (!selected.length) throw new Error("Add at least one item with order quantity");

  const orderId = "ORD-" + Utilities.getUuid().slice(0, 8).toUpperCase();
  const now = new Date();

  selected.forEach(line => {
    let foundRow = -1;
    let stockName = line.itemName;
    let prevQty = 0;

    for (let i = 1; i < stockData.length; i++) {
      if (String(stockData[i][0]).trim() === line.itemId) {
        foundRow = i + 1;
        stockName = String(stockData[i][1] || line.itemName);
        prevQty = Number(stockData[i][3] || 0);
        break;
      }
    }

    if (foundRow === -1) throw new Error(`Stock item not found: ${line.itemId}`);

    const expectedTotal = prevQty + line.orderQty;

    ordersSheet.appendRow([
      orderId,
      now,
      line.itemId,
      stockName,
      prevQty,
      line.orderQty,
      "",
      expectedTotal,
      supplier,
      "pending",
      "",
      "",
      createdBy,
      line.note
    ]);
  });

  return { success: true, orderId, created: selected.length };
}

function confirmOrderLine(orderId, itemId, receivedQty, confirmedBy) {
  const ss = getSS();
  const ordersSheet = getOrdersSheet_();
  const stockSheet = ss.getSheetByName(SHEETS.STOCK);
  const historySheet = ss.getSheetByName(SHEETS.HISTORY);
  const ordersData = ordersSheet.getDataRange().getValues();

  const targetOrderId = String(orderId || "").trim();
  const targetItemId = String(itemId || "").trim();

  let rowIndex = -1;
  for (let i = 1; i < ordersData.length; i++) {
    if (String(ordersData[i][0]).trim() === targetOrderId && String(ordersData[i][2]).trim() === targetItemId) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) throw new Error("Order line not found");

  const row = ordersData[rowIndex - 1];

  if (String(row[9] || "").toLowerCase() !== "pending") {
    throw new Error("This order line has already been processed");
  }

  const itemName = String(row[3] || "");
  const prevQty = Number(row[4] || 0);
  const orderQty = Number(row[5] || 0);
  const supplier = String(row[8] || "");
  const notes = String(row[13] || "");
  const rawReceived = String(receivedQty || "").trim();
  const actualReceived = rawReceived === "" ? orderQty : Number(rawReceived);

  if (!Number.isFinite(actualReceived) || actualReceived < 0) {
    throw new Error("Received quantity must be a valid number");
  }

  const status = rawReceived === "" ? "confirmed" : "manual";
  const newStockQty = prevQty + actualReceived;
  const now = new Date();

  const stockData = stockSheet.getDataRange().getValues();
  let stockRowIndex = -1;

  for (let i = 1; i < stockData.length; i++) {
    if (String(stockData[i][0]).trim() === targetItemId) {
      stockRowIndex = i + 1;
      break;
    }
  }

  if (stockRowIndex === -1) throw new Error("Stock item not found");

  stockSheet.getRange(stockRowIndex, 4).setValue(newStockQty);
  stockSheet.getRange(stockRowIndex, 7).setValue(now);
  stockSheet.getRange(stockRowIndex, 8).setValue(notes);
  stockSheet.getRange(stockRowIndex, 9).setValue(confirmedBy);

  ordersSheet.getRange(rowIndex, 7).setValue(actualReceived);
  ordersSheet.getRange(rowIndex, 10).setValue(status);
  ordersSheet.getRange(rowIndex, 11).setValue(confirmedBy);
  ordersSheet.getRange(rowIndex, 12).setValue(now);

  const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

  historySheet.appendRow([
    now,
    days[now.getDay()],
    itemName,
    prevQty,
    newStockQty,
    `+${actualReceived}`,
    confirmedBy,
    `Order ${status}${supplier ? ` from ${supplier}` : ""}${notes ? ` | ${notes}` : ""}`
  ]);

  return { success: true, status, newStockQty };
}

function saveOrderLineNote(orderId, itemId, notes) {
  const sheet = getOrdersSheet_();
  const data = sheet.getDataRange().getValues();
  const targetOrderId = String(orderId || "").trim();
  const targetItemId = String(itemId || "").trim();
  const noteText = String(notes || "").trim();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === targetOrderId && String(data[i][2]).trim() === targetItemId) {
      sheet.getRange(i + 1, 14).setValue(noteText);
      return { success: true };
    }
  }

  throw new Error("Order line not found");
}

function bulkUpdateStockActions(updates, username) {
  if (!Array.isArray(updates) || !updates.length) {
    throw new Error("No stock updates provided");
  }

  const results = [];

  updates.forEach(u => {
    const itemId = String(u.itemId || "").trim();
    const changeQty = Number(u.changeQty || 0);
    const notes = String(u.notes || "").trim();

    if (!itemId) return;
    if (!Number.isFinite(changeQty) || (changeQty === 0 && !notes)) return;

    results.push(updateStockAction(itemId, changeQty, notes, username));
  });

  return { success: true, updated: results.length };
}

function confirmOrderLineWithNotes(orderId, itemId, receivedQty, confirmedBy, notesOverride) {
  const ss = getSS();
  const ordersSheet = getOrdersSheet_();
  const stockSheet = ss.getSheetByName(SHEETS.STOCK);
  const historySheet = ss.getSheetByName(SHEETS.HISTORY);
  const ordersData = ordersSheet.getDataRange().getValues();

  const targetOrderId = String(orderId || "").trim();
  const targetItemId = String(itemId || "").trim();

  let rowIndex = -1;
  for (let i = 1; i < ordersData.length; i++) {
    if (
      String(ordersData[i][0] || "").trim() === targetOrderId &&
      String(ordersData[i][2] || "").trim() === targetItemId
    ) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) throw new Error("Order line not found");

  const row = ordersData[rowIndex - 1];

  if (String(row[9] || "").toLowerCase() !== "pending") {
    throw new Error("This order line has already been processed");
  }

  const itemName = String(row[3] || "");
  const prevQty = Number(row[4] || 0);
  const orderQty = Number(row[5] || 0);
  const supplier = String(row[8] || "");
  const notes = String(notesOverride ?? row[13] ?? "").trim();
  const rawReceived = String(receivedQty ?? "").trim();
  const actualReceived = rawReceived === "" ? orderQty : Number(rawReceived);

  if (!Number.isFinite(actualReceived) || actualReceived < 0) {
    throw new Error("Received quantity must be a valid number");
  }

  const status = rawReceived === "" ? "confirmed" : "manual";
  const newStockQty = prevQty + actualReceived;
  const now = new Date();

  const stockData = stockSheet.getDataRange().getValues();
  let stockRowIndex = -1;

  for (let i = 1; i < stockData.length; i++) {
    if (String(stockData[i][0] || "").trim() === targetItemId) {
      stockRowIndex = i + 1;
      break;
    }
  }

  if (stockRowIndex === -1) throw new Error("Stock item not found");

  stockSheet.getRange(stockRowIndex, 4).setValue(newStockQty);
  stockSheet.getRange(stockRowIndex, 7).setValue(now);
  stockSheet.getRange(stockRowIndex, 8).setValue(notes);
  stockSheet.getRange(stockRowIndex, 9).setValue(confirmedBy);

  ordersSheet.getRange(rowIndex, 7).setValue(actualReceived);
  ordersSheet.getRange(rowIndex, 10).setValue(status);
  ordersSheet.getRange(rowIndex, 11).setValue(confirmedBy);
  ordersSheet.getRange(rowIndex, 12).setValue(now);
  ordersSheet.getRange(rowIndex, 14).setValue(notes);

  const days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

  historySheet.appendRow([
    now,
    days[now.getDay()],
    itemName,
    prevQty,
    newStockQty,
    `+${actualReceived}`,
    confirmedBy,
    `Order ${status}${supplier ? ` from ${supplier}` : ""}${notes ? ` | ${notes}` : ""}`
  ]);

  return { success: true, status, newStockQty };
}

function bulkConfirmOrderLines(lines, confirmedBy) {
  if (!Array.isArray(lines) || !lines.length) {
    throw new Error("No order lines provided");
  }

  const results = [];

  lines.forEach(line => {
    const orderId = String(line.orderId || "").trim();
    const itemId = String(line.itemId || "").trim();
    const receivedQty = String(line.receivedQty ?? "").trim();
    const note = String(line.note || "").trim();

    if (!orderId || !itemId) return;

    results.push(confirmOrderLineWithNotes(orderId, itemId, receivedQty, confirmedBy, note));
  });

  return { success: true, processed: results.length };
}

function setupCustomerOrderSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let customers = ss.getSheetByName('Customers');
  if (!customers) customers = ss.insertSheet('Customers');
  customers.clear();
  customers.appendRow([
    'customerId',
    'name',
    'location',
    'phone',
    'email',
    'createdAt',
    'updatedAt'
  ]);

  let orders = ss.getSheetByName('CustomerOrders');
  if (!orders) orders = ss.insertSheet('CustomerOrders');
  orders.clear();
  orders.appendRow([
    'orderId',
    'customerId',
    'customerName',
    'customerPhone',
    'customerEmail',
    'customerLocation',
    'createdAt',
    'createdBy',
    'status',
    'itemsJSON',
    'total',
    'acceptedJSON',
    'completedJSON',
    'timelineJSON',
    'notes',
    'updatedAt'
  ]);
}

function createCustomer(data) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Customers');
  const id = 'CUST-' + Date.now();
  sh.appendRow([id, data.name, data.location, data.phone, data.email, new Date()]);
  return { customerId: id };
}

function createCustomerOrder(payload, currentUser) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orders = ss.getSheetByName('CustomerOrders');
  if (!orders) throw new Error('CustomerOrders sheet not found');

  const customer = upsertCustomer_(payload);
  const now = new Date();
  const orderId = 'ORD-' + Utilities.getUuid().slice(0, 8).toUpperCase();
  const items = Array.isArray(payload.items) ? payload.items : [];

  const total = items.reduce((sum, it) => {
    const qty = Number(it.qty || 0);
    const price = Number(it.price || 0);
    return sum + (qty * price);
  }, 0);

  const createdByObj = { name: String(currentUser || '').trim(), time: now.toISOString() };
  const timeline = [{
    action: 'created',
    by: String(currentUser || '').trim(),
    time: now.toISOString(),
    note: 'Order created'
  }];

  orders.appendRow([
    orderId,
    customer.customerId,
    String(payload.customerName || '').trim(),
    String(payload.customerPhone || '').trim(),
    String(payload.customerEmail || '').trim(),
    String(payload.customerLocation || '').trim(),
    now,
    String(currentUser || '').trim(),
    String(payload.status || 'new').trim().toLowerCase(),
    JSON.stringify(items),
    total,
    JSON.stringify(createdByObj),
    '',
    JSON.stringify(timeline),
    String(payload.notes || '').trim(),
    now
  ]);

  sendCustomerOrderEmails_({
    orderId,
    action: 'created',
    payload,
    currentUser,
    total,
    items
  });

  return { success: true, orderId, customerId: customer.customerId };
}

function safeJsonParse_(value, fallback) {
  try {
    if (value === null || value === undefined || value === '') return fallback;
    if (typeof value === 'object') return value;
    return JSON.parse(String(value));
  } catch (e) {
    return fallback;
  }
}

function addItemsToCustomerOrder(orderId, newItems, currentUser) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('CustomerOrders');
  if (!sheet) throw new Error('CustomerOrders sheet not found');

  const data = sheet.getDataRange().getValues();
  const target = String(orderId || '').trim();
  const now = new Date();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').trim() === target) {
      const items = safeJsonParse_(data[i][9], []);
      const timeline = safeJsonParse_(data[i][13], []);
      const extra = Array.isArray(newItems) ? newItems : [];

      items.push(...extra);

      timeline.push({
        action: 'items_added',
        by: String(currentUser || '').trim(),
        time: now.toISOString(),
        note: `Added ${extra.length} item(s)`
      });

      const oldTotal = Number(data[i][10] || 0);
      const addedTotal = extra.reduce((sum, it) => sum + (Number(it.qty || 0) * Number(it.price || 0)), 0);

      sheet.getRange(i + 1, 10).setValue(JSON.stringify(items));
      sheet.getRange(i + 1, 11).setValue(oldTotal + addedTotal);
      sheet.getRange(i + 1, 14).setValue(JSON.stringify(timeline));
      sheet.getRange(i + 1, 16).setValue(now);

      sendCustomerOrderEmails_({
        orderId,
        action: 'updated',
        payload: {
          customerName: String(data[i][2] || ''),
          customerEmail: String(data[i][4] || ''),
          items: extra,
          total: oldTotal + addedTotal
        },
        currentUser,
        total: oldTotal + addedTotal,
        items: extra
      });

      return { success: true };
    }
  }

  throw new Error('Order not found');
}

function upsertCustomer_(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Customers');
  if (!sheet) throw new Error('Customers sheet not found');

  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1);

  const name = String(payload.customerName || '').trim();
  const phone = String(payload.customerPhone || '').trim();
  const email = String(payload.customerEmail || '').trim().toLowerCase();
  const location = String(payload.customerLocation || '').trim();
  const now = new Date();

  let foundRow = -1;

  for (let i = 0; i < rows.length; i++) {
    const rowEmail = String(rows[i][4] || '').trim().toLowerCase();
    const rowPhone = String(rows[i][3] || '').trim();

    if ((email && rowEmail === email) || (phone && rowPhone === phone)) {
      foundRow = i + 2;
      break;
    }
  }

  if (foundRow > 0) {
    const customerId = String(sheet.getRange(foundRow, 1).getValue() || '');
    sheet.getRange(foundRow, 2, 1, 6).setValues([[
      name,
      location,
      phone,
      email,
      sheet.getRange(foundRow, 6).getValue() || now,
      now
    ]]);
    return { customerId, created: false };
  }

  const customerId = 'CUST-' + Utilities.getUuid().slice(0, 8).toUpperCase();
  sheet.appendRow([customerId, name, location, phone, email, now, now]);
  return { customerId, created: true };
}

function updateCustomerOrderStatus(orderId, newStatus, currentUser) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('CustomerOrders');
  if (!sheet) throw new Error('CustomerOrders sheet not found');

  const data = sheet.getDataRange().getValues();
  const target = String(orderId || '').trim();
  const status = String(newStatus || '').trim().toLowerCase();
  const now = new Date();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0] || '').trim() === target) {
      const row = data[i];
      const customerName = String(row[2] || '');
      const customerEmail = String(row[4] || '');
      const itemsJSON = String(row[9] || '[]');
      const total = Number(row[10] || 0);
      const timeline = safeJsonParse_(row[13], []);

      timeline.push({
        action: status,
        by: String(currentUser || '').trim(),
        time: now.toISOString(),
        note: `Status changed to ${status}`
      });

      sheet.getRange(i + 1, 9).setValue(status);

      sheet.getRange(i + 1, 12).setValue(
        status === 'accepted'
          ? JSON.stringify({ name: String(currentUser || '').trim(), time: now.toISOString() })
          : row[11]
      );

      sheet.getRange(i + 1, 13).setValue(
        ['completed', 'delivered', 'collected'].includes(status)
          ? JSON.stringify({ name: String(currentUser || '').trim(), time: now.toISOString() })
          : row[12]
      );

      sheet.getRange(i + 1, 14).setValue(JSON.stringify(timeline));
      sheet.getRange(i + 1, 16).setValue(now);

      sendCustomerOrderEmails_({
        orderId,
        action: status,
        payload: {
          customerName,
          customerEmail,
          items: safeJsonParse_(itemsJSON, []),
          total
        },
        currentUser,
        total,
        items: safeJsonParse_(itemsJSON, [])
      });

      return { success: true };
    }
  }

  throw new Error('Order not found');
}

function getTodayCustomerOrders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('CustomerOrders');
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const tz = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const rows = sheet.getDataRange().getValues().slice(1);

  return rows
    .filter(r => {
      const d = r[6];
      if (!(d instanceof Date)) return false;
      return Utilities.formatDate(d, tz, 'yyyy-MM-dd') === today;
    })
    .map(row => ({
      orderId: String(row[0] || ''),
      customerId: String(row[1] || ''),
      customerName: String(row[2] || ''),
      customerPhone: String(row[3] || ''),
      customerEmail: String(row[4] || ''),
      customerLocation: String(row[5] || ''),
      createdAt: row[6] instanceof Date ? row[6].toISOString() : String(row[6] || ''),
      createdBy: String(row[7] || ''),
      status: String(row[8] || 'new'),
      itemsJSON: String(row[9] || '[]'),
      total: Number(row[10] || 0),
      acceptedJSON: String(row[11] || ''),
      completedJSON: String(row[12] || ''),
      timelineJSON: String(row[13] || '[]'),
      notes: String(row[14] || ''),
      updatedAt: row[15] instanceof Date ? row[15].toISOString() : String(row[15] || '')
    }));
}

function getCustomerOrdersByDate(dateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('CustomerOrders');
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const tz = Session.getScriptTimeZone();
  const rows = sheet.getDataRange().getValues().slice(1);

  return rows
    .filter(r => {
      const d = r[6];
      if (!(d instanceof Date)) return false;
      return Utilities.formatDate(d, tz, 'yyyy-MM-dd') === String(dateStr || '').trim();
    })
    .map(row => ({
      orderId: String(row[0] || ''),
      customerId: String(row[1] || ''),
      customerName: String(row[2] || ''),
      customerPhone: String(row[3] || ''),
      customerEmail: String(row[4] || ''),
      customerLocation: String(row[5] || ''),
      createdAt: row[6] instanceof Date ? row[6].toISOString() : String(row[6] || ''),
      createdBy: String(row[7] || ''),
      status: String(row[8] || 'new'),
      itemsJSON: String(row[9] || '[]'),
      total: Number(row[10] || 0),
      acceptedJSON: String(row[11] || ''),
      completedJSON: String(row[12] || ''),
      timelineJSON: String(row[13] || '[]'),
      notes: String(row[14] || ''),
      updatedAt: row[15] instanceof Date ? row[15].toISOString() : String(row[15] || '')
    }));
}

function getSuperAdminEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Users');
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const headers = data.shift();

  const emailIndex = headers.findIndex(h => String(h).toLowerCase() === 'email');
  const roleIndex = headers.findIndex(h => String(h).toLowerCase() === 'role');
  const statusIndex = headers.findIndex(h => String(h).toLowerCase() === 'status');

  if (emailIndex === -1 || roleIndex === -1) return [];

  return data
    .filter(row =>
      String(row[roleIndex]).toLowerCase() === 'superadmin' &&
      (statusIndex === -1 || String(row[statusIndex]).toLowerCase() === 'active')
    )
    .map(row => row[emailIndex])
    .filter(email => email && String(email).includes('@'));
}

function sendOrderEmails(orderId, payload, user, action) {
  const subject = `Order Update: ${orderId}`;
  const adminEmails = getAllUserEmails(); // your existing users system
  const superAdminEmails = getSuperAdminEmails();
  let message = `Order ID: ${orderId}\nAction: ${action}\nBy: ${user}`;

  // ---------- CUSTOMER EMAIL ----------
  if (action === 'created') {
    sendCustomerEmail(payload.customerId, "We are working on your order.");
  }
  if (action === 'completed') {
    sendCustomerEmail(payload.customerId, "Your order is ready to collect.");
  }

  // ---------- ADMINS ----------
  if (action === 'created' || action === 'updated') {
    adminEmails.forEach(e => MailApp.sendEmail(e, subject, message));
  }

  // ---------- SUPER ADMIN ----------
  superAdminEmails.forEach(e => MailApp.sendEmail(e, subject, message));
}

function getAllUserEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Users'); // make sure sheet name matches
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const headers = data.shift();

  const emailIndex = headers.findIndex(h => String(h).toLowerCase() === 'email');
  const statusIndex = headers.findIndex(h => String(h).toLowerCase() === 'status');

  if (emailIndex === -1) return [];

  return data
    .filter(row => {
      if (statusIndex === -1) return true;
      return String(row[statusIndex]).toLowerCase() === 'active';
    })
    .map(row => row[emailIndex])
    .filter(email => email && String(email).includes('@'));
}

function sendCustomerOrderEmails_(customerId, text) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Customers');
  const data = sh.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == customerId) {
      const email = data[i][4];
      if (email) {
        MailApp.sendEmail(email, "Order Update", text);
      }
      return;
    }
  }
}

function testEmail() {
  MailApp.sendEmail("your@email.com", "Test", "Working");
}
