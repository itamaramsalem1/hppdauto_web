# HPPD Automator Web App

This app lets you upload labor templates and actual staffing reports, runs a Python script to compare them, and gives you a downloadable Excel report.

## 🛠 How to Deploy (Render.com)

1. Sign in at [https://render.com](https://render.com)
2. Click **“New Web Service”**
3. Connect your GitHub repo
4. Set the following:
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `gunicorn app:app`
   - **Environment:** Python 3.10+
5. Deploy!

## 🌐 Link to Custom Domain (e.g. hppdauto.me)

1. Go to your Render app → **Settings → Custom Domains**
2. Add: `hppdauto.me`
3. Render will give you a **CNAME or A record**
4. Go to your Namecheap DNS settings:
   - Add a **CNAME Record**
   - Host: `@`
   - Value: your Render URL (e.g., `hppdauto.onrender.com`)
   - TTL: automatic
5. Wait ~15–60 mins to propagate

## ✅ Usage

1. Upload your labor template Excel files
2. Upload actual report files
3. Select a date
4. Download the results Excel

## 🗂 Folder Structure

