# Thai DOCX 🇹🇭📄

> สร้างไฟล์ Word (.docx) ภาษาไทยให้ **เขียนเต็มบรรทัด ฟอนต์ถูกต้อง จัดหน้าสวยระดับเอกสารราชการ** — ใช้ฟรี ไม่ต้องแอดไลน์ใคร ไม่ต้องจ่ายเงิน

[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Python 3.8+](https://img.shields.io/badge/Python-3.8%2B-blue.svg)](https://www.python.org/)

เครื่องมือสร้างเอกสาร Word ภาษาไทยด้วย Python ที่แก้ปัญหาคลาสสิก 2 ข้อซึ่งโปรแกรมสร้าง .docx ทั่วไปทำพลาด เปิดเป็นโอเพนซอร์สให้ใช้และแก้ไขได้อิสระ

![เปรียบเทียบก่อน/หลังใช้ thai-docx](docs/comparison.png)

> 📦 **มี 2 เวอร์ชัน**
> - **`main` (เวอร์ชันนี้)** — ฉบับเต็มสำหรับ**ข้าราชการ**: รวมฟอนต์แห่งชาติ 10 ตระกูล + preset เอกสารราชการ (สารบรรณ) + สารบัญ/ตาราง/เลขหน้า
> - **branch [`thai-docx`](../../tree/thai-docx)** — ฉบับเบสิกสำหรับ**คนทั่วไป**: เน้นแก้ข้อความไทยให้เต็มบรรทัด ไฟล์เล็ก ไม่รวมฟอนต์
>
> โหลดเฉพาะเบสิก: `git clone -b thai-docx https://github.com/Netthip/thai-docx-for-a-thai-civil-servant-eager-to-learn.git`

---

## ปัญหาที่แก้ให้

### 1. ภาษาไทยตัดบรรทัดก่อนเวลา → เสียพื้นที่ขอบขวา
ภาษาไทยไม่มีช่องว่างระหว่างคำ Word เลยไม่รู้ว่าจะตัดบรรทัดตรงไหน มักตัดเร็วเกินไปจนเกิดบรรทัดสั้น ๆ
**แก้โดย** แทรก Zero-Width Space (U+200B) ระหว่างคำไทย (ตัดคำด้วย [PyThaiNLP](https://github.com/PyThaiNLP/pythainlp)) — มองไม่เห็น แต่บอก Word ว่าตัดบรรทัดตรงไหนได้บ้าง

### 2. ฟอนต์ไทยไม่เปลี่ยนตามที่สั่ง
`python-docx` ตั้งฟอนต์ผ่าน `run.font.name` ให้แค่ฝั่งอังกฤษ (`w:ascii`/`w:hAnsi`) **ไม่ตั้งฝั่ง Complex Script (`w:cs`) ที่อักษรไทยใช้จริง** — ผลคือบางเครื่องฟอนต์ไทยไม่เปลี่ยนเลย
**แก้โดย** ตั้งฟอนต์ผ่าน XML ให้ครบทุกฝั่ง รวมขนาด (`w:szCs`) และตัวหนา/เอียง (`w:bCs`/`w:iCs`)

---

## ความสามารถ

- ✅ ข้อความไทยเต็มบรรทัด ไม่ตัดก่อนเวลา
- ✅ ฟอนต์ไทยถูกต้องทุกเครื่อง (แก้บั๊ก Complex Script)
- ✅ Preset จัดหน้า: **ราชการ (สารบรรณ)**, รายงานทั่วไป, บทความ/หนังสือ
- ✅ หัวเรื่อง/หัวข้อย่อย 3 ระดับ, ย่อหน้า, bullet/เลขข้อหลายชั้น, อ้างอิง (quote), คำบรรยายภาพ
- ✅ **ตาราง** จัดฟอนต์/หัวตารางให้อัตโนมัติ
- ✅ **รูปภาพ** พร้อมคำบรรยาย
- ✅ **สารบัญอัตโนมัติ** (TOC field ของ Word)
- ✅ **หัวกระดาษ/ท้ายกระดาษ + เลขหน้า** (เช่น "หน้า ๑", "1 / 10")
- ✅ ตัวหนา `**...**` / ตัวเอียง `*...*` แบบ markdown
- ✅ ใช้ได้ทั้งแบบ CLI และเรียกจากโค้ด Python

---

## ติดตั้ง

```bash
git clone https://github.com/Netthip/thai-docx-for-a-thai-civil-servant-eager-to-learn.git
cd thai-docx-for-a-thai-civil-servant-eager-to-learn
pip install -r requirements.txt

# ติดตั้งฟอนต์แห่งชาติเข้าเครื่อง (มีให้ในรีโป ไม่ต้องไปหาดาวน์โหลดที่อื่น)
python scripts/install_fonts.py
```

## 🔤 ฟอนต์แห่งชาติแถมมาในรีโป — ใช้ฟรี ไม่ต้องไปแอดไลน์ใครให้รำคาญ

รีโปนี้รวม **ฟอนต์ราชการไทย 10 ตระกูล** ที่แจกจ่ายต่อได้ตามกฎหมายไว้ในโฟลเดอร์ [`fonts/`](fonts/) แล้ว — TH SarabunPSK, TH Krub, TH KoHo, TH Niramit AS, TH Kodchasal, TH Baijam, TH Chakra Petch, TH Fah kwang, TH K2D July8, TH Mali Grade 6 (พร้อมชุด IT๙ สำหรับเลขไทย)

```bash
python scripts/install_fonts.py            # ติดตั้งทั้ง 10 ตระกูล (ไม่ต้องสิทธิ์ admin)
python scripts/thai_docx.py --list-fonts   # ดูชื่อ key ของฟอนต์ที่ใช้ได้
```

เลือกฟอนต์ตอนสร้างเอกสารด้วย key สั้น ๆ:

```bash
python scripts/thai_docx.py input.md -o out.docx --font krub      # หรือ koho, niramit, ...
```

```python
create_docx(paragraphs, "out.docx", font_name="krub")   # รับทั้ง key และชื่อเต็ม
```

> ลิขสิทธิ์ฟอนต์: ใช้/ทำซ้ำ/ดัดแปลง/แจกจ่ายได้ฟรี ห้ามขายฟอนต์โดยลำพัง — ดู [`fonts/LICENSE-NATIONAL-FONTS.txt`](fonts/LICENSE-NATIONAL-FONTS.txt)
> (กลุ่ม TH Charmonman/Srisakdi/Charm of AU ไม่ได้รวมไว้ เพราะสัญญาแจกจ่ายจำกัดเฉพาะ DIP/SIPA)

---

## ใช้งานเร็ว (CLI)

```bash
# จากไฟล์ markdown/ข้อความ → เอกสารราชการพร้อมสารบัญและเลขหน้า
python scripts/thai_docx.py input.md -o output.docx \
    --preset saraban --toc --page-number footer-center --title "รายงานประจำปี"
```

รูปแบบไฟล์ input (markdown ง่าย ๆ):

```markdown
# บทที่ 1 บทนำ

ย่อหน้าภาษาไทยที่จะ **เต็มบรรทัด** โดยอัตโนมัติ

## 1.1 วัตถุประสงค์

- ข้อแรก
- ข้อสอง

1. ขั้นตอนหนึ่ง
2. ขั้นตอนสอง
```

---

## ใช้งานผ่านโค้ด Python

```python
from thai_docx import create_docx

paragraphs = [
    {"text": "รายงานการทดสอบระบบ", "type": "title"},
    {"text": "บทที่ ๑ บทนำ", "type": "heading1"},
    {"text": "เนื้อหาภาษาไทยที่รองรับ **ตัวหนา** และ *ตัวเอียง*", "type": "body"},
    {"text": "วัตถุประสงค์ข้อแรก", "type": "bullet"},
    {"type": "table", "header": True, "rows": [
        ["รายการ", "งบประมาณ (บาท)"],
        ["ครุภัณฑ์", "1,250,000"],
    ]},
]

create_docx(
    paragraphs, "output.docx",
    preset="saraban",
    title="รายงานประจำปี",
    toc=True,
    header_text="สำนักงบประมาณ",
    page_number="footer-center",
    page_number_format="หน้า {n}",
)
```

ดูตัวอย่างเต็มที่ [`examples/example_report.py`](examples/example_report.py)

---

## Preset การจัดหน้า

| preset | ฟอนต์/ขนาด | ขอบ บน/ล่าง/ซ้าย/ขวา (ซม.) | เว้นบรรทัด | ย่อหน้าแรก |
|---|---|---|---|---|
| `saraban` | TH Sarabun New / 16 | 2.5 / 2.0 / 3.0 / 2.0 | 1.0 | 1.25 ซม. |
| `default` | TH Sarabun New / 14 | 2.54 รอบด้าน | 1.5 | — |
| `book` | TH Sarabun New / 16 | 2.54 / 2.54 / 3.0 / 2.54 | 1.3 | 1.25 ซม. |

ระบุพารามิเตอร์เองเพื่อทับค่า preset ได้เสมอ (`font_name`, `font_size`, `margins`, `line_spacing`, ...)

---

## ชนิดของ paragraph

`title` · `subtitle` · `heading1`–`heading3` · `body` · `bullet` · `number` · `quote` · `caption` · `table` · `image` · `pagebreak`

---

## หมายเหตุ

- **สารบัญ (TOC)**: เป็น field ของ Word เปิดไฟล์แล้วคลิกขวาที่สารบัญ → **Update Field** (หรือกด F9) เพื่อให้เลขหน้าแสดงผล
- การจัดข้อความใช้ **ชิดซ้าย** เสมอ เพราะ Justify จะยืดช่องว่างน่าเกลียดในภาษาไทย
- ค่าขอบกระดาษของ preset `saraban` อิงแนวทางเอกสารราชการ ปรับได้ตามระเบียบ/หน่วยงานของคุณ

---

## License

[MIT](LICENSE) — ใช้ฟรี แก้ไขได้ แจกจ่ายได้ ทั้งงานส่วนตัวและเชิงพาณิชย์ ไม่ต้องขออนุญาต

## ขอบคุณ

- [PyThaiNLP](https://github.com/PyThaiNLP/pythainlp) — การตัดคำภาษาไทย
- [python-docx](https://github.com/python-openxml/python-docx) — การสร้างไฟล์ Word

หากเครื่องมือนี้มีประโยชน์ ฝากกด ⭐ Star และช่วยแชร์ต่อ เพื่อให้คนไทยได้ใช้เครื่องมือดี ๆ ฟรี ๆ กันครับ
