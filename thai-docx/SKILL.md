---
name: thai-docx
description: |
  สร้างไฟล์ Word (.docx) ที่ภาษาไทยเขียนเต็มบรรทัดไม่ตัดบรรทัดก่อนเวลา และจัดหน้าได้สวยงามระดับเอกสารราชการ ใช้ skill นี้ทุกครั้งที่ต้องสร้างเอกสาร Word ที่มีเนื้อหาภาษาไทย ไม่ว่าจะเป็นรายงาน บทความ เอกสารวิชาการ บันทึกข้อความ หนังสือราชการ หรือเอกสารทั่วไป ใช้ได้กับข้อความภาษาไทยล้วนและไทย-อังกฤษผสม รองรับสารบัญอัตโนมัติ ตาราง รูปภาพ หัวกระดาษ-ท้ายกระดาษ และเลขหน้า
  Trigger เมื่อ: ผู้ใช้ต้องการสร้าง Word ที่มีภาษาไทย, "ทำ docx ภาษาไทย", "สร้างเอกสาร Word", "ข้อความไทยไม่เต็มบรรทัด", "ข้อความตัดบรรทัดผิด", "Thai text line break", "สร้างรายงาน", "เขียนเอกสาร", "หนังสือราชการ", "บันทึกข้อความ", "เอกสารสารบรรณ", "ใส่สารบัญ", "เลขหน้าภาษาไทย"
  ใช้ skill นี้แทน docx skill ปกติเมื่อเอกสารมีภาษาไทย เพราะ docx-js (JavaScript) ไม่จัดการ Thai line breaking ได้ดี และ python-docx ปกติก็ตั้งฟอนต์ฝั่ง Complex Script ของไทยไม่ครบ — สคริปต์ใน skill นี้แก้ทั้งสองเรื่องให้แล้ว
---

# Thai DOCX — เอกสาร Word ภาษาไทยที่เต็มบรรทัด ฟอนต์ถูก จัดหน้าสวย

## ปัญหา 2 ข้อที่ skill นี้แก้

1. **ไทยตัดบรรทัดก่อนเวลา** — ภาษาไทยไม่มีช่องว่างระหว่างคำ Word จึงไม่รู้จะตัดบรรทัดตรงไหน เลยตัดเร็วเกินไป เกิดบรรทัดสั้น ๆ เสียพื้นที่ขอบขวา (อังกฤษไม่เป็นเพราะตัดที่ช่องว่างได้)
2. **ฟอนต์ไทยไม่เปลี่ยนตามที่สั่ง** — python-docx ตั้งฟอนต์ผ่าน `run.font.name` ให้แค่ฝั่ง `w:ascii`/`w:hAnsi` (อังกฤษ) **ไม่ตั้งฝั่ง `w:cs` (Complex Script) ที่อักษรไทยใช้จริง** ผลคือบางเครื่องฟอนต์ไทยไม่เปลี่ยน

## วิธีแก้

1. แทรก **Zero-Width Space (U+200B)** ระหว่างคำไทย (ตัดคำด้วย pythainlp) — มองไม่เห็น กว้างศูนย์ แต่บอก Word ว่า "ตัดบรรทัดตรงนี้ได้" จุดตัดเยอะขึ้น → ข้อความเต็มบรรทัด
2. ตั้งฟอนต์ผ่าน XML ให้ครบทั้ง `w:ascii`/`w:hAnsi`/**`w:cs`**/`w:eastAsia` รวมถึงขนาด (`w:szCs`) และตัวหนา/เอียง (`w:bCs`/`w:iCs`) ฝั่ง complex script

## ติดตั้ง

```bash
pip install pythainlp python-docx
```

## ฟอนต์แห่งชาติ (มีให้ในรีโป)

โฟลเดอร์ `fonts/` มีฟอนต์ราชการไทย 10 ตระกูลที่แจกจ่ายต่อได้ตามกฎหมาย (TH SarabunPSK, TH Krub, TH KoHo, TH Niramit AS, TH Kodchasal, TH Baijam, TH Chakra Petch, TH Fah kwang, TH K2D July8, TH Mali Grade 6) พร้อมชุด IT๙

```bash
python scripts/install_fonts.py          # ติดตั้งเข้าเครื่อง (per-user ไม่ต้อง admin)
python scripts/thai_docx.py --list-fonts # ดู key ของฟอนต์
```

เลือกฟอนต์ด้วย key สั้น ๆ ได้เลย — `font_name` รับทั้ง key (`"krub"`, `"koho"`, `"sarabun"`) และชื่อเต็ม (`"TH Krub"`, `"TH Sarabun New"`)
ฟังก์ชัน `resolve_font(name)` และ `list_bundled_fonts()` ช่วยจัดการชื่อให้

## ใช้งานเร็ว (CLI)

```bash
python scripts/thai_docx.py input.md -o output.docx --preset saraban --toc --page-number footer-center
```

- ไฟล์ input รองรับ markdown ง่าย ๆ: `#`, `##`, `###` = หัวข้อ; `-`/`*` = bullet; `1.` = เลขข้อ; `**ตัวหนา**`
- บรรทัดที่ขึ้นต้นด้วยเลขแบบ `1.2.3 ...` จะกลายเป็นหัวข้อให้อัตโนมัติ
- บรรทัดว่างคั่นย่อหน้า

## ใช้งานผ่านโค้ด (ยืดหยุ่นเต็มที่)

```python
from thai_docx import create_docx

paragraphs = [
    {"text": "รายงานประจำปี", "type": "title"},
    {"text": "บทที่ ๑ บทนำ", "type": "heading1"},
    {"text": "เนื้อหาย่อหน้าภาษาไทย รองรับ **ตัวหนา** และ *ตัวเอียง*", "type": "body"},
    {"text": "ข้อแรก", "type": "bullet"},
    {"text": "ข้อย่อย", "type": "bullet", "level": 1},
    {"type": "table", "header": True, "rows": [
        ["รายการ", "จำนวน"],
        ["ก", "100"],
    ]},
    {"type": "image", "path": "chart.png", "width": 12, "caption": "ภาพที่ 1"},
    {"type": "pagebreak"},
]

create_docx(
    paragraphs, "output.docx",
    preset="saraban",                 # "saraban" | "default" | "book"
    title="ชื่อเรื่องบนหัวเอกสาร",
    toc=True,                          # สารบัญอัตโนมัติ
    header_text="สำนักงบประมาณ",
    footer_text=None,
    page_number="footer-center",      # footer-center | footer-right | header-right
    page_number_format="หน้า {n}",    # ใช้ {n} และ {total} ได้
)
```

### Preset การจัดหน้า

| preset | ฟอนต์/ขนาด | ขอบ (บน/ล่าง/ซ้าย/ขวา ซม.) | เว้นบรรทัด | ย่อหน้า |
|---|---|---|---|---|
| `saraban` (ราชการ) | TH Sarabun New / 16 | 2.5 / 2.0 / 3.0 / 2.0 | 1.0 | 1.25 ซม. |
| `default` (รายงานทั่วไป) | TH Sarabun New / 14 | 2.54 รอบด้าน | 1.5 | ไม่มี |
| `book` (บทความ/หนังสือ) | TH Sarabun New / 16 | 2.54 / 2.54 / 3.0 / 2.54 | 1.3 | 1.25 ซม. |

พารามิเตอร์ที่ระบุเอง (`font_name`, `font_size`, `margins`, `line_spacing`, `first_line_indent` ...) จะทับค่าของ preset เสมอ

### ชนิดของ paragraph

`title`, `subtitle`, `heading1`–`heading3`, `body`, `bullet`, `number` (มี `level` 0–2), `quote`, `caption`, `table`, `image`, `pagebreak`

## ฝังฟอนต์ลงไฟล์ (ค่าเริ่มต้น)

`create_docx()` **ฝังฟอนต์แห่งชาติลงไฟล์ .docx ให้อัตโนมัติ** (ผ่าน `scripts/font_embed.py`,
obfuscate ตาม ECMA-376 §17.8.1) — เปิดเครื่องไหนก็เห็นฟอนต์ถูกแม้ไม่ได้ลงฟอนต์
ปิดได้ด้วย `embed_fonts=False` หรือ CLI `--no-embed` (ไฟล์เล็กลงแต่ต้องมีฟอนต์ที่เครื่องปลายทาง)

## สร้าง PDF แทน DOCX (scripts/thai_pdf.py)

ต้องการ **PDF** ภาษาไทยแทน Word ใช้ `thai_pdf.py` (ใช้ weasyprint) — ใช้ paragraph model,
preset และฟอนต์แห่งชาติ "ชุดเดียวกับ" docx แต่ได้ไฟล์ PDF ที่:

- **ตัดบรรทัดไทยถูกต้อง** อัตโนมัติ (Pango/libthai + ZWSP จาก pythainlp)
- **ฝังฟอนต์ลงไฟล์เสมอ** (subset อัตโนมัติ) — เปิดเครื่องไหนก็เห็นฟอนต์ถูก ไม่ต้องลงฟอนต์ก่อน

```bash
pip install weasyprint          # ติดตั้งครั้งเดียว
python scripts/thai_pdf.py input.md -o output.pdf --preset saraban --toc \
       --page-number footer-center --header "สำนักงบประมาณ"
```

ผ่านโค้ด — signature ล้อตาม `create_docx`:

```python
from thai_pdf import create_pdf
create_pdf(
    paragraphs, "output.pdf",
    preset="saraban", title="ชื่อเรื่อง", toc=True,
    header_text="สำนักงบประมาณ", page_number="footer-center",
    page_number_format="หน้า {n}",   # ใช้ {n} และ {total} ได้
)   # คืน (output_path, embedded_bool)
```

รองรับ paragraph ชนิดเดียวกับ docx ครบ พร้อมสารบัญ/หัว-ท้ายกระดาษ/เลขหน้า
ปิดการแทรก ZWSP ได้ด้วย `--no-zwsp` (หรือ `zwsp=False`)

> เลือกชนิดไฟล์: **ไทยใน Word → thai_docx** · **ไทยใน PDF → thai_pdf** — ทั้งคู่ฝังฟอนต์เป็นค่าเริ่มต้น

## กฎสำคัญ

- **ใช้ python-docx (Python) เท่านั้น ห้ามใช้ docx-js (JavaScript)** สำหรับเอกสารไทย
- **ใช้การจัดชิดซ้าย (LEFT)** — JUSTIFIED จะยืดช่องว่างน่าเกลียด, THAI_DISTRIBUTE จะยืดตัวอักษร
- `insert_zwsp()` ถูกเรียกอัตโนมัติใน `create_docx()` แล้ว ไม่ต้องเรียกเอง
- ZWS มองไม่เห็น กว้างศูนย์ ไม่เปลี่ยนหน้าตาข้อความ แค่เปิดให้ตัดบรรทัดได้
- สารบัญ (TOC) เป็น field ของ Word — เปิดไฟล์แล้วคลิกขวา → Update Field (หรือกด F9) เพื่อให้เลขหน้าแสดง

## ฟังก์ชันที่ใช้ได้ (scripts/thai_docx.py)

- `insert_zwsp(text, engine="newmm")` — แทรก ZWS ระหว่างคำไทย
- `set_run_font(run, name, size, bold, italic, color)` — ตั้งฟอนต์ครบฝั่ง cs
- `create_docx(paragraphs, output_path, ...)` — สร้างไฟล์
- `parse_markdown(raw)` — แปลง markdown/ข้อความ → รายการ paragraph
