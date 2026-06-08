const express      = require('express');
const multer       = require('multer');
const XLSX         = require('xlsx');
const archiver     = require('archiver');
const bcrypt       = require('bcryptjs');
const jwt          = require('jsonwebtoken');
const cookieParser = require('cookie-parser');
const { v4: uuidv4 } = require('uuid');
const fs   = require('fs');
const path = require('path');

const app = express();
const PORT       = process.env.PORT       || 3000;
const JWT_SECRET = process.env.JWT_SECRET || 'AMG_SECRET_KEY_2024';
const DB_PATH    = path.join(__dirname, 'db.json');
const INV_DIR    = path.join(__dirname, 'invoices');
const UPL_DIR    = path.join(__dirname, 'uploads');

if (!fs.existsSync(UPL_DIR)) fs.mkdirSync(UPL_DIR, { recursive: true });

// ── DB ─────────────────────────────────────────────────────────────────
function readDB() {
  try { return JSON.parse(fs.readFileSync(DB_PATH, 'utf8')); }
  catch(e) { const d = defaultDB(); writeDB(d); return d; }
}
function writeDB(d) {
  try { fs.writeFileSync(DB_PATH, JSON.stringify(d, null, 2)); } catch(e) {}
}
function defaultDB() {
  return {
    settings: {
      storeName:    'AMG Associates',
      storeTagline: 'Professional GST Invoice Templates',
      comboPrice:   299,
      adminUser:    'admin',
      adminPassHash:'$2a$10$VMXN5zjOV5biq5lOwmKphe3KYRblIxcKLTljyUt.b8Qsvyb8V3d0O'
    },
    invoices: [
      {id:'GST_SIMPLE_1',  name:'GST Simple Invoice No.1',  type:'Simple Invoice',   icon:'🧾', file:'GST_SIMPLE_Invoice_Formate_No_1.xlsx',        enabled:true},
      {id:'GST_SIMPLE_2',  name:'GST Simple Invoice No.2',  type:'Simple Invoice',   icon:'🧾', file:'GST_SIMPLE_Invoice_Formate_No_2.xlsx',        enabled:true},
      {id:'GST_SIMPLE_3',  name:'GST Simple Invoice No.3',  type:'Simple Invoice',   icon:'🧾', file:'GST_SIMPLE_Invoice_Formate_No_3.xlsx',        enabled:true},
      {id:'GST_DELIVERY_1',name:'Delivery Challan No.1',    type:'Delivery Challan', icon:'🚚', file:'GST_DELIVERY_CHALAN_FORMAT_1.xlsx',           enabled:true},
      {id:'GST_DELIVERY_2',name:'Delivery Challan No.2',    type:'Delivery Challan', icon:'🚚', file:'GST_DELIVERY_CHALAN_FORMAT_2.xlsx',           enabled:true},
      {id:'GST_QUOTATION', name:'GST Quotation Format',     type:'Quotation',        icon:'📋', file:'GST_QUATATION_FORMATE.xlsx',                  enabled:true},
      {id:'GST_PERFORMA',  name:'Proforma Invoice',         type:'Proforma Invoice', icon:'📄', file:'GST_PERFORMA_Invoice_Format.xlsx',            enabled:true},
      {id:'BILL_SHIP_1',   name:'Bill-to Ship-to No.1',     type:'Bill to Ship to',  icon:'📦', file:'Invoice_Format_No__01_bill_to_ship_to.xlsx',  enabled:true},
      {id:'BILL_SHIP_2',   name:'Bill-to Ship-to No.2',     type:'Bill to Ship to',  icon:'📦', file:'Invoice_Format_No__02_bill_to_ship_to.xlsx',  enabled:true},
      {id:'BILL_SHIP_3',   name:'Bill-to Ship-to No.3',     type:'Bill to Ship to',  icon:'📦', file:'Invoice_Format_No__03_bill_to_ship_to.xlsx',  enabled:true},
      {id:'BILL_SHIP_4',   name:'Bill-to Ship-to No.4',     type:'Bill to Ship to',  icon:'📦', file:'Invoice_Format_No__04_bill_to_ship_to.xlsx',  enabled:true},
      {id:'BILL_SHIP_5',   name:'Bill-to Ship-to No.5',     type:'Bill to Ship to',  icon:'📦', file:'Invoice_Format_No__05_bill_to_ship_to.xlsx',  enabled:true},
      {id:'BILL_SHIP_6',   name:'Bill-to Ship-to No.6',     type:'Bill to Ship to',  icon:'📦', file:'Invoice_Format_No__06_bill_to_ship_to.xlsx',  enabled:true}
    ],
    orders: []
  };
}

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(cookieParser());
app.use(express.static(path.join(__dirname, 'public')));

const uploadLogo = multer({
  storage: multer.diskStorage({
    destination: (req, f, cb) => cb(null, UPL_DIR),
    filename:    (req, f, cb) => cb(null, 'logo_' + Date.now() + path.extname(f.originalname))
  }),
  limits: { fileSize: 5*1024*1024 }
});
const uploadInv = multer({
  storage: multer.diskStorage({
    destination: (req, f, cb) => cb(null, INV_DIR),
    filename:    (req, f, cb) => cb(null, f.originalname)
  }),
  limits: { fileSize: 20*1024*1024 },
  fileFilter: (req, f, cb) =>
    f.originalname.match(/\.(xlsx|xls)$/i) ? cb(null,true) : cb(new Error('Only .xlsx files'))
});

function auth(req, res, next) {
  const token = req.cookies.amg_admin ||
    (req.headers.authorization||'').replace('Bearer ','');
  if (!token) return res.status(401).json({ error: 'Login required' });
  try { req.admin = jwt.verify(token, JWT_SECRET); next(); }
  catch(e) { res.status(401).json({ error: 'Session expired' }); }
}

// ═══════════════════════════════════════════════════════════════
//  PUBLIC ROUTES
// ═══════════════════════════════════════════════════════════════

app.get('/api/store', (req, res) => {
  const db = readDB();
  res.json({
    storeName:    db.settings.storeName,
    storeTagline: db.settings.storeTagline,
    comboPrice:   db.settings.comboPrice,
    invoices:     db.invoices.filter(i => i.enabled)
  });
});

// Create order → return orderId + Razorpay link (amount in PAISE)
app.post('/api/order/create', uploadLogo.single('logo'), (req, res) => {
  try {
    const db = readDB();
    const { companyName, address, phone, email, gstin, pan } = req.body;
    if (!companyName) return res.status(400).json({ error: 'Company name required' });

    const price   = db.settings.comboPrice || 299;
    const orderId = uuidv4();
    const now     = new Date().toISOString();

    db.orders.push({
      id:        orderId,
      status:    'pending',
      createdAt: now,
      company:   { companyName, address:address||'', phone:phone||'',
                   email:email||'', gstin:gstin||'', pan:pan||'' },
      logoPath:  req.file ? req.file.filename : null,
      total:     price,
      downloads: [],
      paidAt:    null
    });
    writeDB(db);

    // Razorpay.me amount must be in PAISE (₹299 = 29900 paise)
    const amountPaise = price * 100;
    const payLink = `https://razorpay.me/@amgassociates?amount=${amountPaise}`;

    res.json({ orderId, total: price, payLink });
  } catch(e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});

// Confirm payment → issue 24h download token
app.post('/api/order/confirm', (req, res) => {
  try {
    const { orderId } = req.body;
    const db    = readDB();
    const order = db.orders.find(o => o.id === orderId);
    if (!order) return res.status(404).json({ error: 'Order not found' });
    order.status = 'paid';
    order.paidAt = new Date().toISOString();
    writeDB(db);
    const token = jwt.sign({ orderId, type:'download' }, JWT_SECRET, { expiresIn:'24h' });
    res.json({ downloadToken: token });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// Download ALL as ZIP
app.get('/api/download/:token', (req, res) => {
  let payload;
  try { payload = jwt.verify(req.params.token, JWT_SECRET); }
  catch(e) { return res.status(401).send('Download link expired.'); }

  const db    = readDB();
  const order = db.orders.find(o => o.id === payload.orderId);
  if (!order || order.status !== 'paid') return res.status(403).send('Payment not confirmed.');

  const cd       = order.company;
  const safeName = cd.companyName.replace(/[^a-zA-Z0-9]/g,'_');

  // Log download event
  order.downloads = order.downloads || [];
  order.downloads.push({ type:'ZIP_ALL', at: new Date().toISOString() });
  writeDB(db);

  res.setHeader('Content-Type','application/zip');
  res.setHeader('Content-Disposition',`attachment; filename="${safeName}_GST_Invoices.zip"`);

  const archive = archiver('zip', { zlib:{ level:6 } });
  archive.on('error', e => console.error(e));
  archive.pipe(res);

  for (const inv of db.invoices.filter(i => i.enabled)) {
    const fp = path.join(INV_DIR, inv.file);
    if (!fs.existsSync(fp)) continue;
    try {
      const wb  = XLSX.read(fs.readFileSync(fp), { type:'buffer' });
      customizeWB(wb, inv.file, cd);
      archive.append(Buffer.from(XLSX.write(wb,{type:'buffer',bookType:'xlsx'})),
                     { name:`${safeName}_${inv.name}.xlsx` });
    } catch(e) { console.error('ZIP error', inv.file, e.message); }
  }
  archive.finalize();
});

// Download SINGLE invoice
app.get('/api/download-single/:token/:invoiceId', (req, res) => {
  let payload;
  try { payload = jwt.verify(req.params.token, JWT_SECRET); }
  catch(e) { return res.status(401).send('Link expired.'); }

  const db    = readDB();
  const order = db.orders.find(o => o.id === payload.orderId);
  if (!order || order.status !== 'paid') return res.status(403).send('Payment not confirmed.');

  const inv = db.invoices.find(i => i.id === req.params.invoiceId);
  if (!inv) return res.status(404).send('Invoice not found.');

  const fp = path.join(INV_DIR, inv.file);
  if (!fs.existsSync(fp)) return res.status(404).send('File missing.');

  // Log download event
  order.downloads = order.downloads || [];
  order.downloads.push({ type:'SINGLE', invoiceName: inv.name, at: new Date().toISOString() });
  writeDB(db);

  const cd  = order.company;
  const wb  = XLSX.read(fs.readFileSync(fp), { type:'buffer' });
  customizeWB(wb, inv.file, cd);
  const out = XLSX.write(wb, { type:'buffer', bookType:'xlsx' });
  const sn  = cd.companyName.replace(/[^a-zA-Z0-9]/g,'_');

  res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition',`attachment; filename="${sn}_${inv.name}.xlsx"`);
  res.send(Buffer.from(out));
});

// ═══════════════════════════════════════════════════════════════
//  ADMIN ROUTES
// ═══════════════════════════════════════════════════════════════

app.post('/api/admin/login', (req, res) => {
  const { username, password } = req.body;
  const db = readDB();
  if (username !== db.settings.adminUser)
    return res.status(401).json({ error: 'Invalid credentials' });
  if (!bcrypt.compareSync(password, db.settings.adminPassHash))
    return res.status(401).json({ error: 'Invalid credentials' });
  const token = jwt.sign({ user: username }, JWT_SECRET, { expiresIn:'8h' });
  res.cookie('amg_admin', token, { httpOnly:true, maxAge:8*3600*1000 });
  res.json({ token });
});

app.post('/api/admin/logout', (req, res) => {
  res.clearCookie('amg_admin');
  res.json({ ok:true });
});

app.get('/api/admin/settings', auth, (req, res) => {
  const { adminPassHash, ...s } = readDB().settings;
  res.json(s);
});

app.put('/api/admin/settings', auth, (req, res) => {
  const db = readDB();
  const { storeName, storeTagline, comboPrice, newPassword } = req.body;
  if (storeName)    db.settings.storeName    = storeName;
  if (storeTagline) db.settings.storeTagline = storeTagline;
  if (comboPrice)   db.settings.comboPrice   = parseInt(comboPrice);
  if (newPassword && newPassword.length >= 6)
    db.settings.adminPassHash = bcrypt.hashSync(newPassword, 10);
  writeDB(db);
  res.json({ ok:true });
});

app.get('/api/admin/invoices', auth, (req, res) => res.json(readDB().invoices));

app.post('/api/admin/invoices', auth, uploadInv.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error:'No file uploaded' });
  const db  = readDB();
  const inv = {
    id:      'INV_' + Date.now(),
    name:    req.body.name || req.file.originalname.replace(/\.xlsx?$/i,''),
    type:    req.body.type || 'Custom',
    icon:    req.body.icon || '📄',
    file:    req.file.originalname,
    enabled: true
  };
  db.invoices.push(inv);
  writeDB(db);
  res.json({ ok:true, invoice:inv });
});

app.put('/api/admin/invoices/:id', auth, (req, res) => {
  const db  = readDB();
  const inv = db.invoices.find(i => i.id === req.params.id);
  if (!inv) return res.status(404).json({ error:'Not found' });
  const { name, type, icon, enabled } = req.body;
  if (name    !== undefined) inv.name    = name;
  if (type    !== undefined) inv.type    = type;
  if (icon    !== undefined) inv.icon    = icon;
  if (enabled !== undefined) inv.enabled = enabled === true || enabled === 'true';
  writeDB(db);
  res.json({ ok:true, invoice:inv });
});

app.delete('/api/admin/invoices/:id', auth, (req, res) => {
  const db  = readDB();
  const idx = db.invoices.findIndex(i => i.id === req.params.id);
  if (idx === -1) return res.status(404).json({ error:'Not found' });
  db.invoices.splice(idx,1);
  writeDB(db);
  res.json({ ok:true });
});

// All orders with download history
app.get('/api/admin/orders', auth, (req, res) => {
  res.json(readDB().orders.slice().reverse());
});

// Dashboard stats
app.get('/api/admin/dashboard', auth, (req, res) => {
  const db   = readDB();
  const paid = db.orders.filter(o => o.status === 'paid');
  const pend = db.orders.filter(o => o.status === 'pending');
  res.json({
    totalOrders:   db.orders.length,
    paidOrders:    paid.length,
    pendingOrders: pend.length,
    revenue:       paid.reduce((s,o) => s + (o.total||0), 0),
    totalDownloads:paid.reduce((s,o) => s + (o.downloads||[]).length, 0),
    recentPaid:    paid.slice(-10).reverse().map(o => ({
      id:        o.id.slice(0,8),
      company:   o.company.companyName,
      phone:     o.company.phone,
      email:     o.company.email,
      gstin:     o.company.gstin,
      total:     o.total,
      paidAt:    o.paidAt,
      downloads: o.downloads || []
    }))
  });
});

// Manually confirm an order (admin)
app.post('/api/admin/orders/:id/confirm', auth, (req, res) => {
  const db    = readDB();
  const order = db.orders.find(o => o.id === req.params.id);
  if (!order) return res.status(404).json({ error:'Not found' });
  order.status = 'paid';
  order.paidAt = new Date().toISOString();
  writeDB(db);
  const token = jwt.sign({ orderId:order.id, type:'download' }, JWT_SECRET, { expiresIn:'24h' });
  res.json({ ok:true, downloadToken: token });
});

// Static routes
app.get('/admin',   (req,res) => res.sendFile(path.join(__dirname,'admin','index.html')));
app.get('/admin/*', (req,res) => res.sendFile(path.join(__dirname,'admin','index.html')));
app.get('/download',(req,res) => res.sendFile(path.join(__dirname,'public','download.html')));
app.get('/',        (req,res) => res.sendFile(path.join(__dirname,'public','index.html')));

// ═══════════════════════════════════════════════════════════════
//  XLSX CUSTOMIZATION — exact cells verified for all 13 files
// ═══════════════════════════════════════════════════════════════
function sc(ws, addr, val) {
  if (!ws[addr]) ws[addr] = { t:'s' };
  ws[addr].v = String(val||'');
  ws[addr].w = String(val||'');
  delete ws[addr].f;
}
function blk(cd) {
  return [cd.companyName, cd.address,
    cd.phone  ? 'Tel: '+cd.phone   : '',
    cd.email  ? 'Email: '+cd.email : '',
    cd.gstin  ? 'GSTIN: '+cd.gstin : '',
    cd.pan    ? 'PAN: '+cd.pan     : ''
  ].filter(Boolean).join('\n');
}

function customizeWB(wb, filename, cd) {
  const sheets = {
    'GST_SIMPLE_Invoice_Formate_No_1.xlsx':       '1',
    'GST_SIMPLE_Invoice_Formate_No_2.xlsx':       '3',
    'GST_SIMPLE_Invoice_Formate_No_3.xlsx':       'Sheet1',
    'GST_DELIVERY_CHALAN_FORMAT_1.xlsx':          'Sheet2',
    'GST_DELIVERY_CHALAN_FORMAT_2.xlsx':          'Sheet2',
    'GST_QUATATION_FORMATE.xlsx':                 '1',
    'GST_PERFORMA_Invoice_Format.xlsx':           '1',
    'Invoice_Format_No__01_bill_to_ship_to.xlsx': 'Sheet1',
    'Invoice_Format_No__02_bill_to_ship_to.xlsx': 'Sales Invoice',
    'Invoice_Format_No__03_bill_to_ship_to.xlsx': 'Invoice',
    'Invoice_Format_No__04_bill_to_ship_to.xlsx': 'Sheet1',
    'Invoice_Format_No__05_bill_to_ship_to.xlsx': 'sheet',
    'Invoice_Format_No__06_bill_to_ship_to.xlsx': 'Sheet1',
  };
  const sn = sheets[filename];
  if (!sn || !wb.Sheets[sn]) return;
  const ws = wb.Sheets[sn];

  switch (filename) {
    case 'GST_SIMPLE_Invoice_Formate_No_1.xlsx': sc(ws,'B2',blk(cd)); break;
    case 'GST_SIMPLE_Invoice_Formate_No_2.xlsx': sc(ws,'B1',blk(cd)); break;
    case 'GST_SIMPLE_Invoice_Formate_No_3.xlsx': sc(ws,'A1',blk(cd)); break;
    case 'GST_QUATATION_FORMATE.xlsx':           sc(ws,'A1',blk(cd)); break;
    case 'GST_PERFORMA_Invoice_Format.xlsx':     sc(ws,'B2',blk(cd)); break;

    case 'GST_DELIVERY_CHALAN_FORMAT_1.xlsx':
    case 'GST_DELIVERY_CHALAN_FORMAT_2.xlsx':
      sc(ws,'A3','Company Name: '+cd.companyName);
      sc(ws,'A4','Address: '+cd.address);
      sc(ws,'A6','Phone No.: '+cd.phone);
      sc(ws,'A7','Email: '+cd.email);
      sc(ws,'A8','GSTIN: '+cd.gstin);
      break;

    case 'Invoice_Format_No__01_bill_to_ship_to.xlsx':
      sc(ws,'B2',cd.companyName); sc(ws,'B3',cd.address);
      if(cd.gstin) sc(ws,'B5','GSTIN: '+cd.gstin); break;

    case 'Invoice_Format_No__02_bill_to_ship_to.xlsx':
      sc(ws,'A1',cd.companyName); sc(ws,'A2',cd.address);
      sc(ws,'A3',cd.phone?'Tel: '+cd.phone:'');
      sc(ws,'A4',cd.email?'Email: '+cd.email:'');
      sc(ws,'A5',cd.gstin?'GSTIN: '+cd.gstin:'');
      sc(ws,'A6',cd.pan?'PAN: '+cd.pan:''); break;

    case 'Invoice_Format_No__03_bill_to_ship_to.xlsx':
      sc(ws,'B6',cd.companyName); sc(ws,'B7',cd.address);
      sc(ws,'B8',cd.phone); sc(ws,'B9',cd.email);
      sc(ws,'B10',cd.gstin?'GSTIN: '+cd.gstin:''); break;

    case 'Invoice_Format_No__04_bill_to_ship_to.xlsx':
      sc(ws,'A2',cd.companyName); sc(ws,'A3',cd.address); break;

    case 'Invoice_Format_No__05_bill_to_ship_to.xlsx':
      sc(ws,'A1',cd.companyName); sc(ws,'A2',cd.address);
      sc(ws,'A3',cd.phone?'Tel: '+cd.phone:'');
      sc(ws,'A4',cd.email?'Email: '+cd.email:'');
      sc(ws,'A5',cd.gstin?'GSTIN: '+cd.gstin:'');
      sc(ws,'A6',cd.pan?'PAN: '+cd.pan:'');
      if(cd.gstin) sc(ws,'D11',cd.gstin);
      if(cd.pan)   sc(ws,'D12',cd.pan); break;

    case 'Invoice_Format_No__06_bill_to_ship_to.xlsx':
      sc(ws,'A4',cd.companyName); sc(ws,'A5',cd.address);
      if(cd.gstin) sc(ws,'A7','GSTIN: '+cd.gstin); break;
  }
}

app.listen(PORT, () => {
  console.log(`\n✅ AMG Invoice Store → http://localhost:${PORT}`);
  console.log(`   Admin: http://localhost:${PORT}/admin  (admin / Admin@AMG2024)\n`);
});
