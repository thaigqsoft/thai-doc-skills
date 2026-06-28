---
name: thai-pdf
description: |
  สร้างไฟล์ PDF ที่ภาษาไทยเขียนเต็มบรรทัดไม่ตัดบรรทัดก่อนเวลา ฝังฟอนต์ลงไฟล์ให้อัตโนมัติ และจัดหน้าได้สวยงามระดับเอกสารราชการ ใช้ skill นี้ทุกครั้งที่ต้องสร้าง PDF ที่มีเนื้อหาภาษาไทย ไม่ว่าจะเป็นรายงาน บทความ เอกสารวิชาการ หนังสือราชการ ใบประกาศ หรือเอกสารทั่วไป ใช้ได้กับข้อความไทยล้วนและไทย-อังกฤษผสม รองรับสารบัญอัตโนมัติ ตาราง รูปภาพ หัวกระดาษ-ท้ายกระดาษ และเลขหน้า
  Trigger เมื่อ: ผู้ใช้ต้องการสร้าง PDF ที่มีภาษาไทย, "ทำ PDF ภาษาไทย", "สร้างไฟล์ PDF", "export เป็น PDF", "ข้อความไทยใน PDF ไม่เต็มบรรทัด", "ฟอนต์ไทยใน PDF เพี้ยน", "Thai PDF", "สร้างรายงาน PDF", "หนังสือราชการ PDF", "ใบประกาศ PDF", "ฝังฟอนต์ใน PDF", "เลขหน้า PDF ภาษาไทย"
  ใช้ skill นี้เมื่อ output ที่ต้องการเป็น "PDF" — ถ้าต้องการ "Word (.docx)" ภาษาไทยให้ใช้ skill thai-docx แทน ทั้งสอง skill ใช้ paragraph model / preset / ฟอนต์แห่งชาติชุดเดียวกัน
---

# Thai PDF — PDF ภาษาไทยเต็มบรรทัด ฟอนต์ฝังในไฟล์ จัดหน้าสวย

## ปัญหา 2 ข้อที่ skill นี้แก้ (เหมือน thai-docx แต่ output เป็น PDF)

1. **ไทยตัดบรรทัดก่อนเวลา** — ไทยไม่มีช่องว่างระหว่างคำ เครื่องมือทั่วไปจึงตัดบรรทัดมั่ว
2. **ฟอนต์ไทยไม่ติดไปกับไฟล์** — เปิดเครื่องที่ไม่มีฟอนต์แล้วหน้าตาเพี้ยน

## วิธีแก้

ใช้ **weasyprint** (HTML/CSS → PDF) ซึ่งจัดการให้อัตโนมัติ:

1. **ตัดบรรทัดไทยถูกต้อง** ผ่าน Pango/libthai + เสริม Zero-Width Space (U+200B) จาก pythainlp
2. **ฝังฟอนต์ลง PDF เสมอ** (subset เฉพาะ glyph ที่ใช้ → ไฟล์ไม่อ้วน) เปิดเครื่องไหนก็เห็นฟอนต์ถูก

## ติดตั้ง

```bash
pip install weasyprint pythainlp
```

> weasyprint ต้องมี Pango/cairo ในระบบ (ลินุกซ์ส่วนใหญ่มีอยู่แล้ว) เพื่อให้ตัดบรรทัดไทยได้ดี

## ฟอนต์แห่งชาติ (มีให้ในรีโป `fonts/`)

10 ตระกูล: TH SarabunPSK, TH Krub, TH KoHo, TH Niramit AS, TH Kodchasal, TH Baijam,
TH Chakra Petch, TH Fah kwang, TH K2D July8, TH Mali Grade 6 — ใช้ฝังลง PDF ได้เลย
ไม่ต้องลงฟอนต์เข้าเครื่องก่อน

```bash
python scripts/thai_pdf.py --list-fonts     # ดู key ของฟอนต์
```

เลือกฟอนต์ด้วย `--font <key>` (เช่น `krub`, `koho`, `sarabun`) หรือชื่อเต็ม

## ใช้งานเร็ว (CLI)

```bash
python scripts/thai_pdf.py input.md -o output.pdf --preset saraban --toc \
       --page-number footer-center --header "สำนักงบประมาณ"
```

- input รองรับ markdown ง่าย ๆ: `#`/`##`/`###` = หัวข้อ; `-`/`*` = bullet; `1.` = เลขข้อ; `**หนา**`
- บรรทัดขึ้นต้นด้วยเลขแบบ `1.2.3 ...` กลายเป็นหัวข้ออัตโนมัติ
- `--no-zwsp` ปิดการแทรก ZWSP (พึ่ง Pango ตัดบรรทัดล้วน)

## ใช้งานผ่านโค้ด

```python
from thai_pdf import create_pdf

paragraphs = [
    {"text": "รายงานประจำปี", "type": "title"},
    {"text": "บทที่ ๑ บทนำ", "type": "heading1"},
    {"text": "เนื้อหาย่อหน้าภาษาไทย รองรับ **ตัวหนา** และ *ตัวเอียง*", "type": "body"},
    {"text": "ข้อแรก", "type": "bullet"},
    {"type": "table", "header": True, "rows": [["รายการ", "จำนวน"], ["ก", "100"]]},
    {"type": "image", "path": "chart.png", "width": 12, "caption": "ภาพที่ 1"},
    {"type": "pagebreak"},
]

create_pdf(
    paragraphs, "output.pdf",
    preset="saraban",                 # "saraban" | "default" | "book"
    title="ชื่อเรื่อง",
    toc=True,                          # สารบัญอัตโนมัติ (มีเลขหน้า + leader dots)
    header_text="สำนักงบประมาณ",
    page_number="footer-center",      # footer-center | footer-right | header-right
    page_number_format="หน้า {n}",    # ใช้ {n} และ {total} ได้
)   # คืน (output_path, embedded_bool)
```

### Preset การจัดหน้า

| preset | ฟอนต์/ขนาด | ขอบ (บน/ล่าง/ซ้าย/ขวา ซม.) | เว้นบรรทัด | ย่อหน้า |
|---|---|---|---|---|
| `saraban` (ราชการ) | TH Sarabun New / 16 | 2.5 / 2.0 / 3.0 / 2.0 | 1.0 | 1.25 ซม. |
| `default` (รายงานทั่วไป) | TH Sarabun New / 14 | 2.54 รอบด้าน | 1.5 | ไม่มี |
| `book` (บทความ/หนังสือ) | TH Sarabun New / 16 | 2.54 / 2.54 / 3.0 / 2.54 | 1.3 | 1.25 ซม. |

พารามิเตอร์ที่ระบุเอง (`font_name`, `font_size`, `margins`, `line_spacing`, `first_line_indent`) ทับค่า preset เสมอ

### ชนิดของ paragraph

`title`, `subtitle`, `heading1`–`heading3`, `body`, `bullet`, `number` (มี `level` 0–2),
`quote`, `caption`, `table`, `image`, `pagebreak`

## กฎสำคัญ

- **output เป็น PDF → ใช้ skill นี้** · ถ้าต้องการ **Word (.docx)** ใช้ skill `thai-docx`
- ฝังฟอนต์เป็นค่าเริ่มต้นเสมอ (ถ้าใช้ฟอนต์ที่มีในรีโป) — ไม่พบไฟล์ฟอนต์จะ fallback ฟอนต์ระบบ + แจ้งเตือน
- ใช้การจัดชิดซ้าย (LEFT) — ไม่ใช้ justified เพื่อเลี่ยงช่องว่างยืดน่าเกลียดในภาษาไทย
- ZWSP มองไม่เห็น กว้างศูนย์ ไม่เปลี่ยนหน้าตาข้อความ แค่เปิดให้ตัดบรรทัดได้

## ฟังก์ชันที่ใช้ได้ (scripts/thai_pdf.py)

- `create_pdf(paragraphs, output_path, ...)` — สร้าง PDF (คืน path + ฝังฟอนต์ไหม)
- `parse_markdown(raw)` — แปลง markdown/ข้อความ → รายการ paragraph
- `insert_zwsp(text)` — แทรก ZWSP ระหว่างคำไทย
- `resolve_font(name)` / `list_bundled_fonts()` — จัดการชื่อฟอนต์
