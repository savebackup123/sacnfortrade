# 80/20 SET Stock Screener Bot 📈

บอทคัดกรองหุ้นตลาด SET (ไม่รวม mai) ตามกลยุทธ์ **80/20 Trade Setup** โดยคัดเลือกหุ้นที่มี Momentum แข็งแกร่ง (เคยมี RSI > 80 ในรอบ 75 วันทำการล่าสุด) และปัจจุบันราคาดึงกลับมาพักฐานในแนวโน้มขาขึ้น (EMA10 > Close และ EMA10 > EMA50) พร้อมคัดกรองปัจจัยพื้นฐาน (Dividend Yield > 5%, P/E < 12)

ระบบจะจัดส่งรายงานสรุปและไฟล์ผลลัพธ์ Excel (.xlsx) ไปยัง Telegram ของคุณโดยอัตโนมัติทุกวันทำการ

---

## 🚀 วิธีการติดตั้งและตั้งค่าบน GitHub (สำหรับทำงานอัตโนมัติบน Cloud)

เพื่อให้บอททำงานอัตโนมัติในทุกๆ วันเวลา 18:00 น. (เวลาไทย) บน GitHub Actions ให้ทำตามขั้นตอนดังนี้:

### 1. ตั้งค่า Telegram Bot และ Chat ID
1. สร้างบอทของคุณบน Telegram โดยทักไปที่แชทกับ [@BotFather](https://t.me/BotFather) ส่งข้อความ `/newbot` จากนั้นทำตามขั้นตอนเพื่อรับ **API Token**
2. ค้นหา **Chat ID** ของคุณ (หรือ ID ของกรุ๊ป/แชนแนลที่ต้องการส่ง) โดยนำบอทเข้าไปในแชทนั้นแล้วลองส่งข้อความทดสอบ หรือทักแชท [@userinfobot](https://t.me/userinfobot) เพื่อรับ ID ของคุณเอง

### 2. นำรหัสผ่านไปใส่ใน GitHub Secrets
เมื่อสร้าง Repository นี้บน GitHub แล้ว ให้ไปที่แท็บ:
1. **Settings** -> **Secrets and variables** -> **Actions**
2. กด **New repository secret** และเพิ่มตัวแปรดังนี้:
   * **Name**: `TELEGRAM_BOT_TOKEN`
   * **Secret**: (โทเค็นของบอทที่ได้จาก BotFather)
3. กด **New repository secret** เพิ่มตัวแปรอีกครั้ง:
   * **Name**: `TELEGRAM_CHAT_ID`
   * **Secret**: (Chat ID หรือ ID กลุ่มที่ได้รับ)

---

## 🛠️ วิธีการรันบนเครื่องคอมพิวเตอร์ของคุณเอง (Local Run)

หากต้องการรันเพื่อดึงข้อมูลบนเครื่องของคุณเอง:

1. ติดตั้งไลบรารีที่จำเป็น:
   ```bash
   pip install yfinance pandas pandas-ta openpyxl lxml html5lib requests
   ```
2. สร้างไฟล์ชื่อ `.env` ไว้ในโฟลเดอร์เดียวกับโปรแกรม แล้วใส่ค่าดังนี้:
   ```env
   TELEGRAM_BOT_TOKEN=8828506982:AAEOejNPRjZXlMlHjfXQ9cq_9cqj3VNuzu0
   TELEGRAM_CHAT_ID=(ใส่ Chat ID ของคุณที่นี่)
   ```
3. สั่งรันโปรแกรม:
   ```bash
   python screener_bot.py
   ```
4. ระบบจะส่งข้อมูลไปที่ Telegram และบันทึกไฟล์ Excel สรุปผลไว้ในโฟลเดอร์ปัจจุบัน
