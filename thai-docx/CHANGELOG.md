# Changelog

## 2.1.0

- **แถมฟอนต์แห่งชาติ 10 ตระกูล** (โฟลเดอร์ `fonts/`) ที่แจกจ่ายต่อได้ตามกฎหมาย พร้อมชุด IT๙
- เพิ่ม `scripts/install_fonts.py` ติดตั้งฟอนต์เข้าเครื่องแบบ per-user (Windows/macOS/Linux)
- เพิ่ม registry ฟอนต์ + `resolve_font()` / `list_bundled_fonts()` — เลือกฟอนต์ด้วย key สั้น เช่น `--font krub`
- เพิ่ม `--list-fonts` ใน CLI
- เอกสารลิขสิทธิ์ฟอนต์ (`fonts/LICENSE-NATIONAL-FONTS.txt`) — เว้นกลุ่มจามรมานที่สิทธิ์แจกจ่ายจำกัด

## 2.0.0

ยกเครื่องครั้งใหญ่ — เน้นการจัดหน้าและความถูกต้องของฟอนต์

### แก้ไขสำคัญ (Bug fixes)
- **แก้ฟอนต์ Complex Script ของไทย** — เดิมตั้งฟอนต์ผ่าน `run.font.name` ซึ่ง python-docx เซ็ตแค่ฝั่ง `w:ascii`/`w:hAnsi` ทำให้ฟอนต์ไทยไม่เปลี่ยนในบางเครื่อง ตอนนี้ตั้งครบ `w:cs`/`w:szCs`/`w:bCs`/`w:iCs` ผ่าน XML

### เพิ่มความสามารถ (Features)
- Preset การจัดหน้า: `saraban` (เอกสารราชการ), `default`, `book`
- หัวกระดาษ/ท้ายกระดาษ และเลขหน้า (รองรับรูปแบบ `หน้า {n}`, `{n} / {total}`)
- สารบัญอัตโนมัติ (TOC field)
- ตาราง พร้อมจัดหัวตาราง/ฟอนต์อัตโนมัติ
- รูปภาพ พร้อมคำบรรยาย
- ชนิด paragraph เพิ่ม: `title`, `subtitle`, `bullet`, `number` (หลายระดับ), `quote`, `caption`, `pagebreak`
- ตัวหนา `**...**` / ตัวเอียง `*...*` แบบ markdown ในเนื้อความ
- การย่อหน้าแรก (first-line indent) ตาม preset
- CLI รองรับ markdown (`#`, `-`, `1.`) และตัวเลือก `--preset --toc --page-number --title --header --footer`
- ทำงานต่อได้แม้ยังไม่ได้ติดตั้ง pythainlp (เตือนแล้วข้ามการตัดคำ)

## 1.0.0
- เวอร์ชันแรก: แทรก Zero-Width Space ระหว่างคำไทยให้ข้อความเต็มบรรทัด, สร้าง heading/body พื้นฐาน
