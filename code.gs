/**
 * ============================================================
 *  StudyBazaar Backend — Google Apps Script
 * ============================================================
 *  DEPLOY:
 *   1. sheets.google.com → open your study_bazaar sheet
 *   2. Extensions → Apps Script → delete existing code → paste this whole file
 *   3. Project Settings (gear icon) → Script Properties → Add these 3:
 *        ADMIN_PASSWORD        → choose a strong password (replaces old "owner123")
 *        RAZORPAY_KEY_ID       → from Razorpay Dashboard → Settings → API Keys
 *        RAZORPAY_KEY_SECRET   → from Razorpay Dashboard → Settings → API Keys
 *   4. Deploy → New deployment → Web app → Execute as: Me → Who has access: Anyone
 *   5. Copy the /exec URL → paste into StudyBazaar's Owner Panel → Sheet Setup
 *
 *  Your sheet tabs are created automatically on first use — you do NOT need
 *  to manually create Users/Listings/Assignments/Payments/Reports tabs or
 *  headers anymore. Just make sure the spreadsheet itself exists.
 * ============================================================
 */

const SS = SpreadsheetApp.getActiveSpreadsheet();
const PROPS = PropertiesService.getScriptProperties();

const SHEETS = {
  Users: ['mobile', 'name', 'email', 'dob', 'address', 'city', 'institution', 'aadhar', 'pwHash', 'plan', 'planExpiry', 'payStatus', 'joinDate', 'lastLogin', 'profilePic'],
  Listings: ['id', 'timestamp', 'name', 'mobile', 'type', 'title', 'price', 'city', 'address', 'photo', 'status', 'stream'],
  Assignments: ['id', 'timestamp', 'name', 'mobile', 'role', 'title', 'price', 'city', 'address', 'details', 'status'],
  Payments: ['id', 'timestamp', 'name', 'mobile', 'email', 'plan', 'amount', 'method', 'ref', 'status', 'approvedAt'],
  Reports: ['id', 'timestamp', 'type', 'reported', 'amount', 'desc', 'email']
};
const ADMIN_ONLY = ['Users', 'Payments', 'Reports']; // reading these requires admin password
const PLAN_DAYS = { 'Monthly': 30, 'Annual': 365, 'New Book Seller': 30, 'Lifetime': -1 };
const PLAN_PRICES = { 'Monthly': 10, 'Annual': 100, 'New Book Seller': 50, 'Lifetime': 500 }; // server is the only source of truth for amounts

// ---------------------------------------------------------------- entry points
function doGet(e) {
  return out({ status: 'StudyBazaar API active ✅' });
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    switch (d.action) {
      case 'read': return handleRead(d);
      case 'stats': return handleStats(d);
      case 'append': return handleAppend(d);
      case 'delete': return handleDelete(d);
      case 'register': return handleRegister(d);
      case 'login': return handleLogin(d);
      case 'verifyAdmin': return handleVerifyAdmin(d);
      case 'updateProfile': return handleUpdateProfile(d);
      case 'createOrder': return handleCreateOrder(d);
      case 'verifyPayment': return handleVerifyPayment(d);
      default: return out({ error: 'Unknown action: ' + d.action });
    }
  } catch (err) {
    return out({ error: err.toString() });
  }
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------- sheet helpers
function sheet(name) {
  let sh = SS.getSheetByName(name);
  if (!sh) sh = SS.insertSheet(name);
  if (sh.getLastRow() === 0) sh.appendRow(SHEETS[name]);
  return sh;
}
function colIndex(name, key) { return SHEETS[name].indexOf(key); }

function isAdmin(pw) {
  const real = PROPS.getProperty('38lFWOXaIfAon91n6K58');
  return !!real && !!pw && pw === real;
}
function sha256(str) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(str), Utilities.Charset.UTF_8);
  return raw.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}
function hmacSha256Hex(str, key) {
  const raw = Utilities.computeHmacSha256Signature(str, key);
  return raw.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}
function verifyUserCreds(mobile, password) {
  const sh = sheet('Users');
  const values = sh.getDataRange().getValues();
  const mCol = colIndex('Users', 'mobile'), pCol = colIndex('Users', 'pwHash');
  const hash = sha256(password);
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][mCol]) === String(mobile)) return String(values[i][pCol]) === hash;
  }
  return false;
}

// ---------------------------------------------------------------- READ (public sheets open, private need admin)
function handleRead(d) {
  const name = d.sheet;
  if (!SHEETS[name]) return out({ error: 'Unknown sheet' });
  if (ADMIN_ONLY.indexOf(name) > -1 && !isAdmin(d.adminPw)) return out({ error: 'Unauthorized' });
  const sh = sheet(name);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const rows = values.slice(1).map(r => {
    const o = {}; headers.forEach((h, i) => o[h] = r[i]); return o;
  }).filter(r => r.status !== 'deleted');
  if (name === 'Users') rows.forEach(r => delete r.pwHash); // never leak hashes, even to admin
  return out({ rows });
}

// ---------------------------------------------------------------- PUBLIC STATS (counts only — no PII, no auth needed)
function handleStats(d) {
  const users = sheet('Users').getLastRow() - 1;
  const listings = sheet('Listings').getLastRow() - 1;
  return out({ userCount: Math.max(0, users), listingCount: Math.max(0, listings) });
}

// ---------------------------------------------------------------- APPEND (Listings / Assignments / Reports)
function handleAppend(d) {
  const name = d.sheet;
  if (!SHEETS[name]) return out({ error: 'Unknown sheet' });
  const sh = sheet(name);
  const headers = SHEETS[name];
  const rowData = Object.assign({ id: Utilities.getUuid(), timestamp: new Date().toLocaleString('en-IN'), status: 'Active' }, d.row || {});
  const row = headers.map(h => rowData[h] !== undefined ? rowData[h] : '');
  sh.appendRow(row);
  return out({ ok: true, id: rowData.id });
}

// ---------------------------------------------------------------- DELETE (owner-of-item or admin only)
function handleDelete(d) {
  const name = d.sheet;
  if (!SHEETS[name]) return out({ error: 'Unknown sheet' });
  const sh = sheet(name);
  const values = sh.getDataRange().getValues();
  const idCol = colIndex(name, 'id'), mobileCol = colIndex(name, 'mobile');
  let rowIdx = -1, ownerMobile = null;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idCol]) === String(d.id)) { rowIdx = i; ownerMobile = String(values[i][mobileCol]); break; }
  }
  if (rowIdx === -1) return out({ error: 'Not found' });
  const admin = isAdmin(d.adminPw);
  const owner = !admin && d.mobile && d.pw && String(d.mobile) === ownerMobile && verifyUserCreds(d.mobile, d.pw);
  if (!admin && !owner) return out({ error: 'Not authorized to delete this item' });
  sh.deleteRow(rowIdx + 1);
  return out({ ok: true });
}

// ---------------------------------------------------------------- REGISTER
function handleRegister(d) {
  const sh = sheet('Users');
  const values = sh.getDataRange().getValues();
  const mCol = colIndex('Users', 'mobile'), eCol = colIndex('Users', 'email');
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][mCol]) === String(d.mobile)) return out({ error: 'Mobile already registered' });
    if (String(values[i][eCol]).toLowerCase() === String(d.email).toLowerCase()) return out({ error: 'Email already registered' });
  }
  const now = new Date().toLocaleString('en-IN');
  const rowData = Object.assign({}, d, {
    pwHash: sha256(d.password), plan: 'Free', planExpiry: '', payStatus: '', joinDate: now, lastLogin: now
  });
  const headers = SHEETS.Users;
  sh.appendRow(headers.map(h => rowData[h] !== undefined ? rowData[h] : ''));
  const user = {}; headers.forEach(h => user[h] = rowData[h] !== undefined ? rowData[h] : '');
  delete user.pwHash; delete user.aadhar;
  return out({ ok: true, user });
}

// ---------------------------------------------------------------- LOGIN
function handleLogin(d) {
  const sh = sheet('Users');
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const mCol = headers.indexOf('mobile'), eCol = headers.indexOf('email'), pCol = headers.indexOf('pwHash');
  const hash = sha256(d.password);
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const idMatch = String(row[mCol]) === String(d.id) || String(row[eCol]).toLowerCase() === String(d.id).toLowerCase();
    if (idMatch && String(row[pCol]) === hash) {
      const user = {}; headers.forEach((h, idx) => user[h] = row[idx]);
      delete user.pwHash; delete user.aadhar;
      sh.getRange(i + 1, headers.indexOf('lastLogin') + 1).setValue(new Date().toLocaleString('en-IN'));
      return out({ ok: true, user });
    }
  }
  return out({ error: 'Invalid mobile/email or password' });
}

// ---------------------------------------------------------------- UPDATE OWN PROFILE (name/address/city/institution/pic)
function handleUpdateProfile(d) {
  if (!verifyUserCreds(d.mobile, d.pw)) return out({ error: 'Incorrect password — could not save changes' });
  const sh = sheet('Users');
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const mCol = headers.indexOf('mobile');
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][mCol]) === String(d.mobile)) {
      ['name', 'address', 'city', 'institution', 'profilePic'].forEach(field => {
        if (d[field] !== undefined && d[field] !== '') {
          const col = headers.indexOf(field);
          if (col > -1) sh.getRange(i + 1, col + 1).setValue(d[field]);
        }
      });
      return out({ ok: true });
    }
  }
  return out({ error: 'User not found' });
}

// ---------------------------------------------------------------- ADMIN AUTH
function handleVerifyAdmin(d) {
  return out({ ok: isAdmin(d.password) });
}

function setUserPlan(mobile, plan) {
  const sh = sheet('Users');
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const mCol = headers.indexOf('mobile'), planCol = headers.indexOf('plan'), expCol = headers.indexOf('planExpiry'), payCol = headers.indexOf('payStatus');
  const days = PLAN_DAYS[plan] || 0;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][mCol]) === String(mobile)) {
      sh.getRange(i + 1, planCol + 1).setValue(plan);
      sh.getRange(i + 1, payCol + 1).setValue('approved');
      let expiry = 'lifetime';
      if (days !== -1) { const dt = new Date(); dt.setDate(dt.getDate() + days); expiry = dt.toISOString(); }
      sh.getRange(i + 1, expCol + 1).setValue(expiry);
      return;
    }
  }
}

// ---------------------------------------------------------------- RAZORPAY — auto checkout
function handleCreateOrder(d) {
  const price = PLAN_PRICES[d.plan];
  if (!price) return out({ error: 'Unknown plan' });
  const keyId = PROPS.getProperty('rzp_test_TBNito97ebkFSc');
  const keySecret = PROPS.getProperty('o7EjX7LODw7vTobdXAY1Z5xc');
  if (!keyId || !keySecret) return out({ error: 'Razorpay keys not set in Script Properties yet' });
  const res = UrlFetchApp.fetch('https://api.razorpay.com/v1/orders', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Basic ' + Utilities.base64Encode(keyId + ':' + keySecret) },
    payload: JSON.stringify({ amount: price * 100, currency: 'INR', receipt: 'sb_' + Date.now() }),
    muteHttpExceptions: true
  });
  const data = JSON.parse(res.getContentText());
  if (data.error) return out({ error: data.error.description || 'Order creation failed' });
  return out({ ok: true, orderId: data.id, keyId: keyId, amount: price });
}

function handleVerifyPayment(d) {
  const price = PLAN_PRICES[d.plan];
  if (!price) return out({ error: 'Unknown plan' });
  const keySecret = PROPS.getProperty('RAZORPAY_KEY_SECRET');
  const expected = hmacSha256Hex(d.orderId + '|' + d.paymentId, keySecret);
  if (expected !== d.signature) return out({ error: 'Signature mismatch — payment could not be verified' });
  const sh = sheet('Payments');
  const id = Utilities.getUuid();
  const now = new Date().toLocaleString('en-IN');
  sh.appendRow([id, now, d.name, d.mobile, d.email, d.plan, price, 'Razorpay', d.paymentId, 'approved', now]);
  setUserPlan(d.mobile, d.plan);
  return out({ ok: true });
}
