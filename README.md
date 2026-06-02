# Telegram Checker Dashboard

ระบบนี้ใช้ไฟล์หลักจาก clean folder เท่านั้น:

- `server.js`
- `auth-addon.js`
- `public/`
- `data/`
- `package.json`
- `package-lock.json`
- `.env`
- `.env.example`

## วิธีรัน

```powershell
npm install
npm run check
npm start
```

เปิดเว็บที่ `http://127.0.0.1:3000`

## Health Check

หลัง `npm start` แล้ว เปิดอีก terminal:

```powershell
npm run health
```

หรือเรียกตรง:

```powershell
curl http://127.0.0.1:3000/api/health
```

ควรได้ผลลัพธ์ประมาณ:

```json
{"ok":true}
```

## ทดสอบเร็วด้วย sample 5 เบอร์

สร้างไฟล์ CSV ขนาดเล็ก เช่น:

```csv
name,phone
Sample 1,0812345678
Sample 2,0912345678
Sample 3,+66812345678
Bad Phone,abc123
Bad Prefix,021234567
```

จากหน้าเว็บ:

1. ล็อกอิน dashboard
2. เพิ่มบัญชี Telegram และกดส่ง OTP
3. ยืนยัน OTP หรือ 2FA
4. กด `ใช้บัญชีนี้`
5. อัปโหลด sample CSV
6. กด `เริ่มตรวจ`
7. กด `ทำต่อจากค้าง`
8. ตรวจว่า processed เพิ่มขึ้น และผลมี `YES`, `NO`, `RETRY`, หรือ `INVALID`

## Export

ดาวน์โหลดผลได้จากหน้าเว็บหรือ endpoint:

- `/download/all.csv`
- `/download/yes.csv`
- `/download/no.csv`
- `/download/retry.csv`
- `/download/invalid.csv`
- `/download/all.json`

ไฟล์ export legacy เดิมยังใช้ได้ เช่น `telegram_matches_all.csv`, `yes_only.csv`, `no_only.csv`, `retry_only.csv`
