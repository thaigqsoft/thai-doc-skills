# Thai Document Skills for Claude Code / Agent Skills

สอง Agent Skill สำหรับสร้างเอกสารภาษาไทยให้สวยระดับเอกสารราชการ — แก้ปัญหาภาษาไทย **ตัดบรรทัดผิดที่** และ **ฟอนต์ไทยเพี้ยน/ไม่ฝัง** ที่เครื่องมือทั่วไปทำได้ไม่ดี

| Skill | สร้างไฟล์ | จุดเด่น |
|-------|-----------|---------|
| [`thai-docx`](./thai-docx) | Word (`.docx`) | ตั้งฟอนต์ฝั่ง Complex Script ของไทยครบ, ไทยเขียนเต็มบรรทัด, สารบัญ/ตาราง/หัว-ท้ายกระดาษ/เลขหน้า |
| [`thai-pdf`](./thai-pdf)  | PDF (`.pdf`) | ฝังฟอนต์ลงไฟล์อัตโนมัติ, ตัดคำไทยถูกต้อง, preset/paragraph model ชุดเดียวกับ thai-docx |

ทั้งสอง skill ใช้ **ฟอนต์แห่งชาติ** (13 ตระกูล TH Sarabun, TH Krub ฯลฯ) ที่มาพร้อม repo — ดูสัญญาอนุญาตที่ `fonts/LICENSE-NATIONAL-FONTS.txt`

## ติดตั้ง (Install)

คัดลอกโฟลเดอร์ skill ไปไว้ในไดเรกทอรี skills ของ Claude Code:

```bash
git clone https://github.com/thaigqsoft/thai-doc-skills.git
# เลือกติดตั้งอันที่ต้องการ (หรือทั้งคู่)
cp -r thai-doc-skills/thai-docx ~/.claude/skills/
cp -r thai-doc-skills/thai-pdf  ~/.claude/skills/

# ติดตั้ง dependency ของแต่ละ skill
pip install -r ~/.claude/skills/thai-docx/requirements.txt
pip install -r ~/.claude/skills/thai-pdf/requirements.txt
```

จากนั้นเปิด Claude Code ใหม่ — skill จะถูกค้นพบอัตโนมัติ แล้วสั่งได้เลย เช่น "สร้างรายงานภาษาไทยเป็น Word" หรือ "ทำ PDF หนังสือราชการ"

## ใช้กับ agent อื่น

สคริปต์หลักอยู่ที่ `scripts/thai_docx.py` และ `scripts/thai_pdf.py` รันตรงๆ ด้วย Python ก็ได้ ไม่จำเป็นต้องใช้ผ่าน Claude Code — ดูรายละเอียดใน `SKILL.md` ของแต่ละโฟลเดอร์

## License

โค้ด: ดู `LICENSE` — ฟอนต์แห่งชาติ: ดู `thai-*/fonts/LICENSE-NATIONAL-FONTS.txt`

---
สร้าง/ดูแลโดย [@thaigqsoft](https://github.com/thaigqsoft) · ยินดีรับ PR และ issue 🙌
