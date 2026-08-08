# TU Basketball Alumni — แบบฟอร์มสั่งจองชุด

ฟอร์มสั่งจองชุด Thammasat Basketball Alumni (`index.html`, ไฟล์เดียวจบ ไม่มี build step)
พร้อม backend เก็บออเดอร์ลง Google Sheet ผ่าน Apps Script (`apps-script/Code.gs`)

## โครงสร้าง

```
index.html            หน้าเว็บฟอร์มทั้งหมด (host บน GitHub Pages)
apps-script/Code.gs    โค้ด backend (วางใน Google Apps Script ของชีต)
```

## ขั้นตอนที่ 1 — ตั้งค่า Apps Script (เก็บออเดอร์ลงชีต)

1. เปิด Google Sheet **"2026 Pre-Order Uniform TU Bas"**
2. เมนู **Extensions > Apps Script**
3. ลบโค้ดเดิมในไฟล์ `Code.gs` ทั้งหมด แล้ววางโค้ดจาก `apps-script/Code.gs` ในโปรเจกต์นี้แทน
4. กด **Deploy > New deployment**
   - Select type: **Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
5. กด **Deploy** → อนุญาตสิทธิ์ (จะมีหน้าเตือนว่าแอปยังไม่ยืนยันตัวตน กด Advanced > ไปต่อ ได้เลย เพราะเป็นสคริปต์ของคุณเอง)
6. คัดลอก **Web app URL** ที่ได้ (หน้าตาประมาณ `https://script.google.com/macros/s/xxxxx/exec`)
7. เปิด `index.html` หาบรรทัด:
   ```js
   var APPS_SCRIPT_URL = 'PASTE_YOUR_WEB_APP_URL_HERE';
   ```
   แล้ววาง URL ที่คัดลอกมาแทน `PASTE_YOUR_WEB_APP_URL_HERE`

**หมายเหตุ**: ทุกครั้งที่แก้โค้ดใน `Code.gs` ต้องกด **Deploy > Manage deployments > แก้ไข (ไอคอนดินสอ) > Version: New version > Deploy** ใหม่ ไม่งั้นระบบจะยังใช้โค้ดเวอร์ชันเก่าอยู่

ทดสอบ: เปิด `index.html` ในเบราว์เซอร์ (เปิดจากไฟล์ในเครื่องก็ทดสอบได้เลย) ลองกรอกออเดอร์ปลอมทั้งหมดจนถึงหน้าชำระเงิน แนบรูปอะไรก็ได้เป็นสลิป แล้วกด "ยืนยันคำสั่งซื้อ" — เช็คว่ามีแถวใหม่ขึ้นในชีต (tab "Orders") และมีไฟล์ในโฟลเดอร์ Drive ชื่อ "TU Basketball Preorder Slips"

## ขั้นตอนที่ 2 — ขึ้น GitHub Pages (เพื่อได้ลิงก์สำหรับโปสเตอร์/QR)

Repo: **https://github.com/hoopmania/TUA-2026-Jersey** (สร้างไว้แล้ว)

โฟลเดอร์นี้ init เป็น git repo, ตั้ง remote ไปที่ repo ข้างบน, และ commit ไฟล์ทั้งหมดไว้ให้แล้ว เหลือแค่ push:

```bash
git push -u origin main
```

ถ้าเป็นครั้งแรกที่ push จากเครื่องนี้ Git Credential Manager จะเปิดเบราว์เซอร์ให้ล็อกอิน GitHub — ล็อกอินแล้วค่อยลองรันคำสั่งข้างบนอีกครั้งถ้าจำเป็น

จากนั้นเปิด **Settings > Pages** ของ repo → Source เลือก **Deploy from a branch** → Branch: **main**, folder: **/ (root)** → Save

รอสัก 1-2 นาที ลิงก์ฟอร์มจะอยู่ที่ `https://hoopmania.github.io/TUA-2026-Jersey/` — เอาลิงก์นี้ไปทำ QR code และใส่ในโปสเตอร์ได้เลย

## ข้อควรระวัง

- ไฟล์สลิปที่แนบจะถูกแปลงเป็น base64 ส่งผ่าน POST ไปให้ Apps Script — ถ้าไฟล์ใหญ่มาก (เช่นรูปถ่าย 10MB+) อาจช้าหรือเกินขีดจำกัดของ Apps Script (~50MB ต่อ request) แนะนำเตือนผู้สั่งซื้อให้แนบรูปสลิป ไม่ใช่รูปถ่ายความละเอียดสูงเกินจำเป็น
- Web App ที่ตั้ง Access เป็น "Anyone" หมายความว่าใครก็ยิง request เข้ามาที่ URL นี้ได้ (ไม่ใช่แค่จากฟอร์มนี้) — สำหรับงานสั่งจองขนาดนี้ถือว่าเพียงพอ แต่ไม่ควรเก็บข้อมูลอ่อนไหวอื่นเพิ่มในนี้
- ราคา/ไซส์ ในฟอร์มยังเป็นค่าที่ hardcode ไว้ใน `index.html` (ตัวแปร `PRICE` และ `SIZE_CHART`) ถ้าราคาชุดเปลี่ยนต้องแก้ในไฟล์นี้แล้ว commit ใหม่
