const express    = require('express');
const multer     = require('multer');
const XLSX       = require('xlsx');
const archiver   = require('archiver');
const bcrypt     = require('bcryptjs');
const jwt        = require('jsonwebtoken');
const cookieParser = require('cookie-parser');
const { v4: uuidv4 } = require('uuid');
const fs   = require('fs');
const path = require('path');

const app = express();
const PORT         = process.env.PORT         || 3000;
const JWT_SECRET   = process.env.JWT_SECRET   || 'AMG_SUPER_SECRET_KEY_2024';
const RAZORPAY_URL = process.env.RAZORPAY_URL || 'https://razorpay.me/@amgassociates';
const ADMIN_USER   = process.env.ADMIN_USER   || 'admin';
const ADMIN_PASS   = process.env.ADMIN_PASS   || 'Admin@AMG2024';

const DB_PATH  = path.join(__dirname, 'db.json');
const INV_DIR  = path.join(__dirname, 'invoices');
const UPL_DIR  = path.join(__dirname, 'uploads');

// Ensure dirs exist (Render ephemeral filesystem)
[UPL_DIR].forEach(d => { if (!fs.existsSync(d)) fs.mkdirSync(d, { recursive: true }); });

// ── DB ─────────────────────────────────────────────────────────────────
function readDB() {
  try { return JSON.parse(fs.readFileSync(DB_PATH, 'utf8')); }
  catch(e) { return getDefaultDB(); }
}
function writeDB(d) {
  try { fs.writeFileSync(DB_PATH, JSON.stringify(d, null, 2)); } catch(e) {}
}
function getDefaultDB() {
  return {
    settings: {
      storeName: 'AMG Associates',
      storeTagline: 'Professional GST Invoice Templates',
      paymentLink: RAZORPAY_URL,
      defaultPrice: 35,
      currency: 'INR',
      adminUser: ADMIN_USER,
      adminPassHash: bcrypt.hashSync(ADMIN_PASS, 10)
    },
    invoices: [
      {id:'GST_SIMPLE_1',  name:'GST Simple Invoice No.1',  type:'Simple Invoice',    icon:'🧾', file:'GST_SIMPLE_Invoice_Formate_No_1.xlsx',         price:35, enabled:true},
      {id:'GST_SIMPLE_2',  name:'GST Simple Invoice No.2',  type:'Simple Invoice',    icon:'🧾', file:'GST_SIMPLE_Invoice_Formate_No_2.xlsx',         price:35, enabled:true},
      {id:'GST_SIMPLE_3',  name:'GST Simple Invoice No.3',  type:'Simple Invoice',    icon:'🧾', file:'GST_SIMPLE_Invoice_Formate_No_3.xlsx',         price:35, enabled:true},
      {id:'GST_DELIVERY_1',name:'Delivery Challan No.1',    type:'Delivery Challan',  icon:'🚚', file:'GST_DELIVERY_CHALAN_FORMAT_1.xlsx',            price:35, enabled:true},
      {id:'GST_DELIVERY_2',name:'Delivery Challan No.2',    type:'Delivery Challan',  icon:'🚚', file:'GST_DELIVERY_CHALAN_FORMAT_2.xlsx',            price:35, enabled:true},
      {id:'GST_QUOTATION', name:'GST Quotation Format',     type:'Quotation',         icon:'📋', file:'GST_QUATATION_FORMATE.xlsx',                   price:35, enabled:true},
      {id:'GST_PERFORMA',  name:'Proforma Invoice',         type:'Proforma Invoice',  icon:'📄', file:'GST_PERFORMA_Invoice_Format.xlsx',             price:35, enabled:true},
      {id:'BILL_SHIP_1',   name:'Bill-to Ship-to No.1',     type:'Bill to Ship to',   icon:'📦', file:'Invoice_Format_No__01_bill_to_ship_to.xlsx',   price:35, enabled:true},
      {id:'BILL_SHIP_2',   name:'Bill-to Ship-to No.2',     type:'Bill to Ship to',   icon:'📦', file:'Invoice_Format_No__02_bill_to_ship_to.xlsx',   price:35, enabled:true},
      {id:'BILL_SHIP_3',   name:'Bill-to Ship-to No.3',     type:'Bill to Ship to',   icon:'📦', file:'Invoice_Format_No__03_bill_to_ship_to.xlsx',   price:35, enabled:true},
      {id:'BILL_SHIP_4',   name:'Bill-to Ship-to No.4',     type:'Bill to Ship to',   icon:'📦', file:'Invoice_Format_No__04_bill_to_ship_to.xlsx',   price:35, enabled:true},
      {id:'BILL_SHIP_5',   name:'Bill-to Ship-to No.5',     type:'Bill to Ship to',   icon:'📦', file:'Invoice_Format_No__05_bill_to_ship_to.xlsx',   price:35, enabled:true},
      {id:'BILL_SHIP_6',   name:'Bill-to Ship-to No.6',     type:'Bill to Ship to',   icon:'📦', file:'Invoice_Format_No__06_bill_to_ship_to.xlsx',   price:35, enabled:true}
    ],
    orders: []
  };
}

// ── MIDDLEWARE ─────────────────────────────────────────────────────────
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(cookieParser());
app.use(express.static(path.join(__dirname, 'public')));

// ── MULTER ─────────────────────────────────────────────────────────────
const uploadLogo = multer({
  storage: multer.diskStorage({
    destination: (req, file, cb) => cb(null, UPL_DIR),
    filename:    (req, file, cb) => cb(null, 'logo_' + Date.now() + path.extname(file.originalname))
  }),
  limits: { fileSize: 5 * 1024 * 1024 }
});
const uploadInv = multer({
  storage: multer.diskStorage({
    destination: (req, file, cb) => cb(null, INV_DIR),
    filename:    (req, file, cb) => cb(null, file.originalname)
  }),
  limits: { fileSize: 20 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    if (file.originalname.match(/\.(xlsx|xls)$/i)) cb(null, true);
    else cb(new Error('Only .xlsx/.xls files'));
  }
});

// ── AUTH ───────────────────────────────────────────────────────────────
function requireAdmin(req, res, next) {
  const token = req.cookies.amg_admin_token ||
                (req.headers['authorization'] || '').replace('Bearer ', '');
  if (!token) return res.status(401).json({ error: 'Not authenticated' });
  try { req.admin = jwt.verify(token, JWT_SECRET); next(); }
  catch(e) { res.status(401).json({ error: 'Session expired' }); }
}

// ══════════════════════════════════════════════════════════════════════
//  PUBLIC API
// ══════════════════════════════════════════════════════════════════════

app.get('/api/store', (req, res) => {
  const db = readDB();
  res.json({
    storeName:    db.settings.storeName,
    storeTagline: db.settings.storeTagline,
    currency:     db.settings.currency,
    invoices:     db.invoices.filter(i => i.enabled)
  });
});

// Create order → returns orderId + Razorpay link with amount
app.post('/api/order/create', uploadLogo.single('logo'), (req, res) => {
  try {
    const db = readDB();
    const { companyName, address, phone, email, gstin, pan, selectedIds } = req.body;

    if (!companyName) return res.status(400).json({ error: 'Company name is required' });
    if (!selectedIds) return res.status(400).json({ error: 'No templates selected' });

    let ids;
    try { ids = JSON.parse(selectedIds); }
    catch(e) { return res.status(400).json({ error: 'Invalid selection' }); }

    const chosenInvoices = db.invoices.filter(i => ids.includes(i.id) && i.enabled);
    if (!chosenInvoices.length) return res.status(400).json({ error: 'No valid invoices selected' });

    const total   = chosenInvoices.reduce((s, i) => s + i.price, 0);
    const orderId = uuidv4();

    db.orders.push({
      id: orderId,
      status: 'pending',
      createdAt: new Date().toISOString(),
      company: { companyName, address: address||'', phone: phone||'', email: email||'', gstin: gstin||'', pan: pan||'' },
      logoPath: req.file ? req.file.filename : null,
      selectedIds: ids,
      total,
      paymentRef: null
    });
    writeDB(db);

    // Razorpay link — amount in INR, not editable by user
    const baseLink = (db.settings.paymentLink || RAZORPAY_URL).trim();
    const payLink  = `${baseLink}?amount=${total}`;

    res.json({ orderId, total, payLink });
  } catch(e) {
    console.error('order/create error:', e);
    res.status(500).json({ error: e.message });
  }
});

// Confirm payment → issue download token
app.post('/api/order/confirm', (req, res) => {
  try {
    const { orderId, paymentRef } = req.body;
    if (!orderId) return res.status(400).json({ error: 'orderId required' });
    const db    = readDB();
    const order = db.orders.find(o => o.id === orderId);
    if (!order) return res.status(404).json({ error: 'Order not found' });
    order.status     = 'paid';
    order.paymentRef = paymentRef || 'MANUAL';
    order.paidAt     = new Date().toISOString();
    writeDB(db);
    const token = jwt.sign({ orderId, type: 'download' }, JWT_SECRET, { expiresIn: '24h' });
    res.json({ downloadToken: token });
  } catch(e) {
    res.status(500).json({ error: e.message });
  }
});

// Download ZIP of all purchased invoices
app.get('/api/download/:token', (req, res) => {
  let payload;
  try { payload = jwt.verify(req.params.token, JWT_SECRET); }
  catch(e) { return res.status(401).send('Download link expired or invalid.'); }

  const db    = readDB();
  const order = db.orders.find(o => o.id === payload.orderId);
  if (!order || order.status !== 'paid')
    return res.status(403).send('Payment not confirmed for this order.');

  const cd       = order.company;
  const invoices = db.invoices.filter(i => order.selectedIds.includes(i.id));
  const safeName = cd.companyName.replace(/[^a-zA-Z0-9]/g, '_');

  res.setHeader('Content-Type', 'application/zip');
  res.setHeader('Content-Disposition', `attachment; filename="${safeName}_GST_Invoices.zip"`);

  const archive = archiver('zip', { zlib: { level: 6 } });
  archive.on('error', err => { console.error('archive error', err); });
  archive.pipe(res);

  for (const inv of invoices) {
    const filePath = path.join(INV_DIR, inv.file);
    if (!fs.existsSync(filePath)) { console.warn('Missing file:', inv.file); continue; }
    try {
      const buf = fs.readFileSync(filePath);
      const wb  = XLSX.read(buf, { type: 'buffer' });
      customizeWorkbook(wb, inv.file, cd);
      const out = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
      archive.append(Buffer.from(out), { name: `${safeName}_${inv.file}` });
    } catch(e) { console.error('Error processing', inv.file, e.message); }
  }
  archive.finalize();
});

// Download single invoice
app.get('/api/download-single/:token/:invoiceId', (req, res) => {
  let payload;
  try { payload = jwt.verify(req.params.token, JWT_SECRET); }
  catch(e) { return res.status(401).send('Link expired.'); }

  const db    = readDB();
  const order = db.orders.find(o => o.id === payload.orderId);
  if (!order || order.status !== 'paid') return res.status(403).send('Payment not confirmed.');
  if (!order.selectedIds.includes(req.params.invoiceId)) return res.status(403).send('Not in your order.');

  const inv = db.invoices.find(i => i.id === req.params.invoiceId);
  if (!inv) return res.status(404).send('Invoice not found.');

  const filePath = path.join(INV_DIR, inv.file);
  if (!fs.existsSync(filePath)) return res.status(404).send('File missing on server.');

  try {
    const cd  = order.company;
    const buf = fs.readFileSync(filePath);
    const wb  = XLSX.read(buf, { type: 'buffer' });
    customizeWorkbook(wb, inv.file, cd);
    const out      = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    const safeName = cd.companyName.replace(/[^a-zA-Z0-9]/g, '_');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${safeName}_${inv.file}"`);
    res.send(Buffer.from(out));
  } catch(e) {
    res.status(500).send('Error generating file: ' + e.message);
  }
});

// ══════════════════════════════════════════════════════════════════════
//  ADMIN API
// ══════════════════════════════════════════════════════════════════════

app.post('/api/admin/login', (req, res) => {
  const { username, password } = req.body;
  const db = readDB();
  if (username !== (db.settings.adminUser || ADMIN_USER))
    return res.status(401).json({ error: 'Invalid credentials' });
  const validPass = db.settings.adminPassHash
    ? bcrypt.compareSync(password, db.settings.adminPassHash)
    : password === ADMIN_PASS;
  if (!validPass) return res.status(401).json({ error: 'Invalid credentials' });
  const token = jwt.sign({ user: username, role: 'admin' }, JWT_SECRET, { expiresIn: '8h' });
  res.cookie('amg_admin_token', token, { httpOnly: true, maxAge: 8 * 3600 * 1000 });
  res.json({ token });
});

app.post('/api/admin/logout', (req, res) => {
  res.clearCookie('amg_admin_token');
  res.json({ message: 'Logged out' });
});

app.get('/api/admin/dashboard', requireAdmin, (req, res) => {
  const db   = readDB();
  const paid = db.orders.filter(o => o.status === 'paid');
  res.json({
    totalOrders:  db.orders.length,
    paidOrders:   paid.length,
    revenue:      paid.reduce((s, o) => s + o.total, 0),
    recentOrders: paid.slice(-5).reverse().map(o => ({
      id:      o.id.slice(0, 8),
      company: o.company.companyName,
      total:   o.total,
      paidAt:  o.paidAt,
      items:   o.selectedIds.length
    }))
  });
});

app.get('/api/admin/settings', requireAdmin, (req, res) => {
  const db = readDB();
  const { adminPassHash, ...safe } = db.settings;
  res.json(safe);
});

app.put('/api/admin/settings', requireAdmin, (req, res) => {
  const db = readDB();
  const { storeName, storeTagline, paymentLink, defaultPrice, currency, newPassword } = req.body;
  if (storeName)    db.settings.storeName    = storeName;
  if (storeTagline) db.settings.storeTagline = storeTagline;
  if (paymentLink)  db.settings.paymentLink  = paymentLink;
  if (defaultPrice) db.settings.defaultPrice = parseInt(defaultPrice);
  if (currency)     db.settings.currency     = currency;
  if (newPassword && newPassword.length >= 6)
    db.settings.adminPassHash = bcrypt.hashSync(newPassword, 10);
  writeDB(db);
  res.json({ message: 'Settings updated' });
});

app.get('/api/admin/invoices', requireAdmin, (req, res) => {
  res.json(readDB().invoices);
});

app.post('/api/admin/invoices', requireAdmin, uploadInv.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  const db  = readDB();
  const inv = {
    id:      'INV_' + Date.now(),
    name:    req.body.name || req.file.originalname,
    type:    req.body.type || 'Custom',
    icon:    req.body.icon || '📄',
    file:    req.file.originalname,
    price:   parseInt(req.body.price) || db.settings.defaultPrice,
    enabled: true
  };
  db.invoices.push(inv);
  writeDB(db);
  res.json({ message: 'Invoice added', invoice: inv });
});

app.put('/api/admin/invoices/:id', requireAdmin, (req, res) => {
  const db  = readDB();
  const inv = db.invoices.find(i => i.id === req.params.id);
  if (!inv) return res.status(404).json({ error: 'Not found' });
  const { name, type, icon, price, enabled } = req.body;
  if (name    !== undefined) inv.name    = name;
  if (type    !== undefined) inv.type    = type;
  if (icon    !== undefined) inv.icon    = icon;
  if (price   !== undefined) inv.price   = parseInt(price);
  if (enabled !== undefined) inv.enabled = enabled === true || enabled === 'true';
  writeDB(db);
  res.json({ message: 'Updated', invoice: inv });
});

app.delete('/api/admin/invoices/:id', requireAdmin, (req, res) => {
  const db  = readDB();
  const idx = db.invoices.findIndex(i => i.id === req.params.id);
  if (idx === -1) return res.status(404).json({ error: 'Not found' });
  db.invoices.splice(idx, 1);
  writeDB(db);
  res.json({ message: 'Deleted' });
});

app.get('/api/admin/orders', requireAdmin, (req, res) => {
  res.json(readDB().orders.slice().reverse());
});

app.post('/api/admin/orders/:id/confirm', requireAdmin, (req, res) => {
  const db    = readDB();
  const order = db.orders.find(o => o.id === req.params.id);
  if (!order) return res.status(404).json({ error: 'Not found' });
  order.status     = 'paid';
  order.paidAt     = new Date().toISOString();
  order.paymentRef = req.body.paymentRef || 'ADMIN_MANUAL';
  writeDB(db);
  const token = jwt.sign({ orderId: order.id, type: 'download' }, JWT_SECRET, { expiresIn: '24h' });
  res.json({ message: 'Confirmed', downloadToken: token });
});

// ── STATIC ROUTES ──────────────────────────────────────────────────────
app.get('/admin',   (req, res) => res.sendFile(path.join(__dirname, 'admin', 'index.html')));
app.get('/admin/*', (req, res) => res.sendFile(path.join(__dirname, 'admin', 'index.html')));
app.get('/download',(req, res) => res.sendFile(path.join(__dirname, 'public', 'download.html')));
app.get('/',        (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));

// ══════════════════════════════════════════════════════════════════════
//  XLSX CUSTOMIZATION — exact cell per file, verified against real data
// ══════════════════════════════════════════════════════════════════════
function sc(ws, addr, val) {
  if (!ws[addr]) ws[addr] = { t: 's' };
  ws[addr].v = String(val || '');
  ws[addr].w = String(val || '');
  delete ws[addr].f; // remove formula so value sticks
}

function buildBlock(cd) {
  return [
    cd.companyName,
    cd.address,
    cd.phone   ? 'Tel: '     + cd.phone  : '',
    cd.email   ? 'Email: '   + cd.email  : '',
    cd.gstin   ? 'GSTIN: '   + cd.gstin  : '',
    cd.pan     ? 'PAN No.: ' + cd.pan    : ''
  ].filter(Boolean).join('\n');
}

function customizeWorkbook(wb, filename, cd) {
  const ws    = wb.Sheets[wb.SheetNames[0]];
  const block = buildBlock(cd);

  switch (filename) {
    case 'GST_SIMPLE_Invoice_Formate_No_1.xlsx':
      sc(ws, 'B2', block); break;

    case 'GST_SIMPLE_Invoice_Formate_No_2.xlsx':
      sc(ws, 'B1', block); break;

    case 'GST_SIMPLE_Invoice_Formate_No_3.xlsx':
      sc(ws, 'A1', block); break;

    case 'GST_DELIVERY_CHALAN_FORMAT_1.xlsx':
    case 'GST_DELIVERY_CHALAN_FORMAT_2.xlsx':
      sc(ws, 'A3', 'Company Name: ' + cd.companyName);
      sc(ws, 'A4', 'Address: '      + cd.address);
      sc(ws, 'A6', 'Phone: '        + cd.phone);
      sc(ws, 'A7', 'Email: '        + cd.email);
      sc(ws, 'A8', 'GSTIN: '        + cd.gstin);
      break;

    case 'GST_QUATATION_FORMATE.xlsx':
      sc(ws, 'A1', block);
      if (cd.gstin) sc(ws, 'B10', cd.gstin);
      break;

    case 'GST_PERFORMA_Invoice_Format.xlsx':
      sc(ws, 'B2', block); break;

    case 'Invoice_Format_No__01_bill_to_ship_to.xlsx':
      sc(ws, 'B2', cd.companyName);
      sc(ws, 'B3', cd.address);
      sc(ws, 'B5', cd.gstin ? 'GSTIN: ' + cd.gstin : '');
      break;

    case 'Invoice_Format_No__02_bill_to_ship_to.xlsx':
      // Seller info in rows 1-5 area (before bill-to section)
      sc(ws, 'A1', cd.companyName);
      sc(ws, 'A2', cd.address);
      sc(ws, 'A3', cd.phone  ? 'Tel: '    + cd.phone  : '');
      sc(ws, 'A4', cd.email  ? 'Email: '  + cd.email  : '');
      sc(ws, 'A5', cd.gstin  ? 'GSTIN: '  + cd.gstin  : '');
      break;

    case 'Invoice_Format_No__03_bill_to_ship_to.xlsx':
      sc(ws, 'B6',  cd.companyName);
      sc(ws, 'B7',  cd.address);
      sc(ws, 'B8',  cd.phone);
      sc(ws, 'B9',  cd.email);
      sc(ws, 'B10', cd.gstin ? 'GSTIN: ' + cd.gstin : '');
      break;

    case 'Invoice_Format_No__04_bill_to_ship_to.xlsx':
      sc(ws, 'A2', cd.companyName);
      sc(ws, 'A3', cd.address);
      break;

    case 'Invoice_Format_No__05_bill_to_ship_to.xlsx':
      // Company header above invoice no. row
      sc(ws, 'A1', cd.companyName);
      sc(ws, 'A2', cd.address);
      sc(ws, 'A3', cd.phone ? 'Tel: ' + cd.phone : '');
      sc(ws, 'A4', cd.email ? 'Email: ' + cd.email : '');
      if (cd.gstin) sc(ws, 'D11', cd.gstin);
      if (cd.pan)   sc(ws, 'D12', cd.pan);
      break;

    case 'Invoice_Format_No__06_bill_to_ship_to.xlsx':
      sc(ws, 'A4', cd.companyName);
      sc(ws, 'A5', cd.address);
      if (cd.gstin) sc(ws, 'A7', 'GSTIN: ' + cd.gstin);
      break;
  }
}

app.listen(PORT, () => {
  console.log(`\n✅ AMG Invoice Store → http://localhost:${PORT}`);
  console.log(`   Admin: http://localhost:${PORT}/admin`);
  console.log(`   Login: ${ADMIN_USER} / ${ADMIN_PASS}\n`);
});
