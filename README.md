# AMG Associates – GST Invoice Template Store

A complete web application for selling customized GST invoice templates.

---

## 🗂 Project Structure

```
amg_project/
├── server.js          ← Main Node.js server (all API routes)
├── db.json            ← Database (invoices, orders, settings)
├── package.json
├── public/
│   ├── index.html     ← Customer storefront (3-step flow)
│   └── download.html  ← Post-payment download page
├── admin/
│   └── index.html     ← Admin panel (login-protected)
├── invoices/          ← All 13 XLSX template files
└── uploads/           ← Customer logo uploads
```

---

## 🚀 Quick Start (Local)

### Requirements
- Node.js 16+ (download from nodejs.org)

### Steps
```bash
# 1. Enter project folder
cd amg_project

# 2. Install dependencies (only needed once)
npm install

# 3. Start the server
node server.js
```

Open your browser:
- **Store:** http://localhost:3000
- **Admin:** http://localhost:3000/admin

### Default Admin Credentials
```
Username: admin
Password: Admin@AMG2024
```
**Change the password immediately** from Admin → Settings after first login.

---

## 🌐 Deploy FREE Online (Render.com)

1. Create account at **render.com** (free)
2. Click **New → Web Service**
3. Connect your GitHub repo (or upload as zip)
4. Settings:
   - **Build Command:** `npm install`
   - **Start Command:** `node server.js`
   - **Environment:** Node
5. Click **Deploy**

Your live URLs will be:
- Store: `https://your-app.onrender.com`
- Admin: `https://your-app.onrender.com/admin`

**Note:** On Render's free tier, the server sleeps after 15 minutes of inactivity. Upgrade to Starter ($7/mo) for always-on.

---

## 🌐 Deploy FREE Online (Railway.app)

1. Create account at **railway.app**
2. Click **New Project → Deploy from GitHub**
3. Railway auto-detects Node.js
4. Done — your app is live!

---

## ⚙️ Razorpay Setup

1. Log in to your Razorpay Dashboard
2. Go to **Payment Links → Payment Pages**
3. Create a Payment Page (or use your existing link)
4. Copy the link: `https://razorpay.me/@yourhandle`
5. In Admin → Settings, paste this link
6. Razorpay will receive the `amount` as a query parameter — make sure your payment page is set to **accept custom amounts** or is configured for the right amount

---

## 🔐 Admin Panel Features

- **Dashboard:** Revenue, order count, recent orders
- **Invoices:** Add / Edit / Delete / Show / Hide templates
- **Orders:** View all orders, manually confirm payments
- **Settings:** Change store name, tagline, Razorpay link, price, password

---

## 📄 How the Flow Works

1. **Customer fills company details** (name, address, GSTIN, PAN, logo)
2. **Customer selects templates** (price shown per template, running total)
3. **Order created on server** → customer redirected to Razorpay
4. **After payment**, customer clicks "Payment Done" → server marks order paid
5. **Download page** → server customizes each XLSX with company details → ZIP download

---

## 🔧 Customization

### Change admin password
Admin → Settings → Change Admin Password

### Add new invoice template
Admin → Invoices → Add Template → Upload .xlsx file

### Change pricing
Admin → Invoices → Edit (per invoice) OR Admin → Settings → Default Price

---

## 📞 Support
Built for AMG Associates. Contact your developer for server issues.
