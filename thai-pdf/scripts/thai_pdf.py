#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
thai_pdf.py — สร้าง PDF ภาษาไทยเต็มบรรทัด ฟอนต์ฝังในไฟล์ จัดหน้าระดับเอกสารราชการ
(skill: thai-pdf — standalone, ไม่พึ่ง skill อื่น)

ใช้ weasyprint (HTML/CSS → PDF) ซึ่งจัดการ 2 เรื่องยากของภาษาไทยให้อัตโนมัติ:
  1) ตัดบรรทัดไทย (ไทยไม่มีช่องว่างระหว่างคำ) — ผ่าน Pango/libthai + เสริม ZWSP จาก pythainlp
  2) ฝังฟอนต์ลง PDF (subset อัตโนมัติ) — เปิดเครื่องไหนก็เห็นฟอนต์ถูกต้อง ไม่ต้องลงฟอนต์ก่อน

รองรับ: หัวเรื่อง/หัวข้อย่อย, ย่อหน้า (**หนา**/*เอียง*), bullet/เลขข้อ, ตาราง, รูปภาพ,
สารบัญอัตโนมัติ, หัว-ท้ายกระดาษ, เลขหน้า, preset เอกสารราชการ

MIT License — ใช้ฟรี แก้ไข แจกจ่ายได้
"""
from __future__ import annotations

import os
import re
import struct
import html as _html
import argparse

try:
    from pythainlp import word_tokenize
    _HAS_PYTHAINLP = True
except Exception:  # pragma: no cover
    _HAS_PYTHAINLP = False

    def word_tokenize(text, engine="newmm"):  # type: ignore
        return [text]


_HERE = os.path.dirname(os.path.abspath(__file__))
FONTS_DIR = os.path.join(_HERE, os.pardir, "fonts", "national")

ZWS = "​"  # Zero-Width Space
_THAI_RUN = re.compile(r"([฀-๿]+)")


# ─────────────────────────────────────────────────────────────────────────────
# ตัดคำ / แทรก ZWS
# ─────────────────────────────────────────────────────────────────────────────
def insert_zwsp(text: str, engine: str = "newmm") -> str:
    """แทรก Zero-Width Space ระหว่างคำไทย — อังกฤษ/ตัวเลข/URL ไม่ถูกแตะ"""
    if not text:
        return text
    out = []
    for part in _THAI_RUN.split(text):
        if part and _THAI_RUN.fullmatch(part):
            out.append(ZWS.join(word_tokenize(part, engine=engine)))
        else:
            out.append(part)
    return "".join(out)


# ─────────────────────────────────────────────────────────────────────────────
# Preset การจัดหน้า (ตรงกับ skill thai-docx)
# ─────────────────────────────────────────────────────────────────────────────
PRESETS = {
    "saraban": dict(font_name="TH Sarabun New", font_size=16, page_size="A4",
                    margins=dict(top=2.5, bottom=2.0, left=3.0, right=2.0),
                    line_spacing=1.0, first_line_indent=1.25, space_after=0),
    "default": dict(font_name="TH Sarabun New", font_size=14, page_size="A4",
                    margins=dict(top=2.54, bottom=2.54, left=2.54, right=2.54),
                    line_spacing=1.5, first_line_indent=0.0, space_after=6),
    "book": dict(font_name="TH Sarabun New", font_size=16, page_size="A4",
                 margins=dict(top=2.54, bottom=2.54, left=3.0, right=2.54),
                 line_spacing=1.3, first_line_indent=1.25, space_after=0),
}

# ─────────────────────────────────────────────────────────────────────────────
# ฟอนต์แห่งชาติในรีโป (key → ชื่อ family จริงที่ฝังในไฟล์ฟอนต์)
# ─────────────────────────────────────────────────────────────────────────────
BUNDLED_FONTS = {
    "sarabun": "TH SarabunPSK", "krub": "TH Krub", "koho": "TH KoHo",
    "niramit": "TH Niramit AS", "kodchasal": "TH Kodchasal", "baijam": "TH Baijam",
    "chakrapetch": "TH Chakra Petch", "fahkwang": "TH Fah kwang",
    "k2d": "TH K2D July8", "mali": "TH Mali Grade 6",
}
_FONT_ALIASES = {
    "th sarabun new": "TH Sarabun New", "sarabun new": "TH Sarabun New",
    "th sarabunpsk": "TH SarabunPSK", "sarabun psk": "TH SarabunPSK",
    "th sarabun psk": "TH SarabunPSK",
}
# ชื่อ family ที่เอกสารอ้าง แต่ไฟล์ฟอนต์ใช้ชื่อ family อื่น
_EMBED_ALIASES = {
    "th sarabun new": "th sarabunpsk", "sarabun new": "th sarabunpsk",
    "sarabun": "th sarabunpsk",
}


def resolve_font(name: str) -> str:
    if not name:
        return name
    low = name.strip().lower()
    if low in BUNDLED_FONTS:
        return BUNDLED_FONTS[low]
    if low in _FONT_ALIASES:
        return _FONT_ALIASES[low]
    return name


# ─────────────────────────────────────────────────────────────────────────────
# อ่านชื่อ family/subfamily จากตาราง 'name' ของไฟล์ TTF → index family→style→path
# ─────────────────────────────────────────────────────────────────────────────
def _read_font_meta(path):
    with open(path, "rb") as f:
        data = f.read()
    if data[:4] not in (b"\x00\x01\x00\x00", b"true", b"OTTO"):
        return None
    num = struct.unpack(">H", data[4:6])[0]
    name_off = None
    off = 12
    for _ in range(num):
        if data[off:off + 4] == b"name":
            name_off = struct.unpack(">I", data[off + 8:off + 12])[0]
        off += 16
    if name_off is None:
        return None
    count, stroff = struct.unpack(">HH", data[name_off + 2:name_off + 6])
    base, strbase = name_off + 6, name_off + stroff
    fam = sub = None
    for i in range(count):
        pid, eid, lid, nid, ln, o = struct.unpack(
            ">HHHHHH", data[base + i * 12:base + i * 12 + 12])
        if nid not in (1, 2):
            continue
        raw = data[strbase + o:strbase + o + ln]
        try:
            s = raw.decode("utf-16-be") if pid in (0, 3) else raw.decode("latin-1")
        except Exception:
            continue
        if nid == 1 and fam is None:
            fam = s
        elif nid == 2 and sub is None:
            sub = s
    return (fam, sub)


_INDEX = None


def _font_index():
    global _INDEX
    if _INDEX is not None:
        return _INDEX
    idx = {}
    if os.path.isdir(FONTS_DIR):
        for fn in sorted(os.listdir(FONTS_DIR)):
            if not fn.lower().endswith((".ttf", ".otf")):
                continue
            meta = _read_font_meta(os.path.join(FONTS_DIR, fn))
            if meta and meta[0]:
                idx.setdefault(meta[0].lower(), {})[(meta[1] or "Regular").lower()] = \
                    os.path.join(FONTS_DIR, fn)
    _INDEX = idx
    return idx


def resolve_font_files(display_name):
    """คืน {subfamily_lower: path} ของ family ที่ขอ (รองรับ alias) หรือ {} ถ้าไม่พบ"""
    key = (display_name or "").strip().lower()
    key = _EMBED_ALIASES.get(key, key)
    return _font_index().get(key, {})


def list_bundled_fonts():
    return dict(BUNDLED_FONTS)


# ─────────────────────────────────────────────────────────────────────────────
# markdown ง่าย ๆ → รายการ paragraph
# ─────────────────────────────────────────────────────────────────────────────
def parse_markdown(raw: str) -> list:
    items, buf = [], []
    lines = raw.replace("\r\n", "\n").split("\n")

    def flush():
        if buf:
            items.append({"text": " ".join(buf).strip(), "type": "body"})
            buf.clear()

    i = 0
    while i < len(lines):
        s = lines[i].rstrip().strip()
        if not s:
            flush(); i += 1; continue
        m = re.match(r"^(#{1,3})\s+(.*)$", s)
        if m:
            flush(); items.append({"text": m.group(2), "type": f"heading{len(m.group(1))}"})
            i += 1; continue
        if re.match(r"^[-*]\s+", s):
            flush(); items.append({"text": re.sub(r"^[-*]\s+", "", s), "type": "bullet"})
            i += 1; continue
        if re.match(r"^\d+[.)]\s+", s):
            flush(); items.append({"text": re.sub(r"^\d+[.)]\s+", "", s), "type": "number"})
            i += 1; continue
        if re.match(r"^\d+(\.\d+)*\s+\S", s) and len(s) < 100:
            flush(); depth = s.split()[0].count(".")
            items.append({"text": s, "type": f"heading{min(depth + 1, 3)}"})
            i += 1; continue
        buf.append(s); i += 1
    flush()
    return items


# ─────────────────────────────────────────────────────────────────────────────
# paragraph → HTML
# ─────────────────────────────────────────────────────────────────────────────
def _inline(text):
    s = _html.escape(text, quote=False)
    s = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", s)
    s = re.sub(r"(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)", r"<em>\1</em>", s)
    return s


def _thai_text(text, zwsp=True):
    if zwsp and text:
        text = insert_zwsp(text)
    return _inline(text)


_STYLE_CSS = {
    "regular": ("normal", "normal"), "bold": ("bold", "normal"),
    "italic": ("normal", "italic"), "bold italic": ("bold", "italic"),
}


def _font_face_css(family_display):
    files = resolve_font_files(family_display)
    if not files:
        return "", False
    rules = []
    for sub, path in files.items():
        weight, style = _STYLE_CSS.get(sub, ("normal", "normal"))
        url = "file://" + os.path.abspath(path).replace(" ", "%20")
        rules.append("@font-face{font-family:'docfont';src:url('%s');"
                     "font-weight:%s;font-style:%s;}" % (url, weight, style))
    return "\n".join(rules), True


def _items_to_html(paragraphs, zwsp=True):
    out, toc, hc = [], [], 0
    i, n = 0, len(paragraphs)
    while i < n:
        it = paragraphs[i]
        t = it.get("type", "body")
        if t in ("bullet", "number"):
            tag = "ul" if t == "bullet" else "ol"
            out.append("<%s class='lst'>" % tag)
            while i < n and paragraphs[i].get("type") == t:
                lvl = paragraphs[i].get("level", 0)
                out.append("<li style='margin-left:%dem'>%s</li>"
                           % (lvl * 2, _thai_text(paragraphs[i].get("text", ""), zwsp)))
                i += 1
            out.append("</%s>" % tag)
            continue
        if t == "title":
            out.append("<h1 class='title'>%s</h1>" % _thai_text(it.get("text", ""), zwsp))
        elif t == "subtitle":
            out.append("<p class='subtitle'>%s</p>" % _thai_text(it.get("text", ""), zwsp))
        elif t in ("heading1", "heading2", "heading3"):
            lvl = int(t[-1]); hc += 1; anc = "h%d" % hc
            toc.append((lvl, it.get("text", ""), anc))
            out.append("<h%d id='%s' class='hd hd%d'>%s</h%d>"
                       % (lvl + 1, anc, lvl, _thai_text(it.get("text", ""), zwsp), lvl + 1))
        elif t == "quote":
            out.append("<blockquote>%s</blockquote>" % _thai_text(it.get("text", ""), zwsp))
        elif t == "caption":
            out.append("<p class='caption'>%s</p>" % _thai_text(it.get("text", ""), zwsp))
        elif t == "pagebreak":
            out.append("<div class='pb'></div>")
        elif t == "table":
            rows, hdr = it.get("rows", []), it.get("header", False)
            out.append("<table><tbody>")
            for r, row in enumerate(rows):
                cell = "th" if (hdr and r == 0) else "td"
                out.append("<tr>" + "".join(
                    "<%s>%s</%s>" % (cell, _thai_text(str(c), zwsp), cell) for c in row) + "</tr>")
            out.append("</tbody></table>")
        elif t == "image":
            path, w = it.get("path", ""), it.get("width")
            style = ("width:%scm;" % w) if w else ""
            url = "file://" + os.path.abspath(path).replace(" ", "%20") if path else ""
            out.append("<figure><img src='%s' style='%s'/>" % (url, style))
            if it.get("caption"):
                out.append("<figcaption>%s</figcaption>" % _thai_text(it["caption"], zwsp))
            out.append("</figure>")
        else:
            out.append("<p class='body'>%s</p>" % _thai_text(it.get("text", ""), zwsp))
        i += 1
    return "\n".join(out), toc


def _toc_html(entries, title="สารบัญ"):
    if not entries:
        return ""
    rows = ["<h2 class='hd hd1'>%s</h2>" % _html.escape(title), "<div class='toc'>"]
    for lvl, text, anc in entries:
        rows.append("<p class='tocline toc%d'><a href='#%s'>%s</a>"
                    "<span class='tocpage'></span></p>" % (lvl, anc, _inline(text)))
    rows.append("</div><div class='pb'></div>")
    return "\n".join(rows)


def _page_css(cfg, header_text, footer_text, page_number, page_number_format):
    m, size = cfg["margins"], cfg.get("page_size", "A4")
    boxes = []

    def _num(fmt):
        c = _html.escape(fmt).replace("{n}", "\" counter(page) \"").replace(
            "{total}", "\" counter(pages) \"")
        return '"%s"' % c

    pos = {"footer-center": "@bottom-center", "footer-right": "@bottom-right",
           "header-right": "@top-right"}
    if page_number in pos:
        boxes.append("%s{content:%s;font-family:'docfont';font-size:%dpt;}"
                     % (pos[page_number], _num(page_number_format), cfg["font_size"] - 2))
    if header_text:
        boxes.append("@top-center{content:\"%s\";font-family:'docfont';font-size:%dpt;color:#444;}"
                     % (_html.escape(header_text), cfg["font_size"] - 2))
    if footer_text:
        boxes.append("@bottom-left{content:\"%s\";font-family:'docfont';font-size:%dpt;color:#444;}"
                     % (_html.escape(footer_text), cfg["font_size"] - 2))
    return ("@page{size:%s;margin:%scm %scm %scm %scm;%s}"
            % (size, m["top"], m["right"], m["bottom"], m["left"], "\n".join(boxes)))


def create_pdf(paragraphs, output_path, *, preset=None, font_name=None, font_size=None,
               margins=None, line_spacing=None, first_line_indent=None,
               header_text=None, footer_text=None, page_number=None,
               page_number_format="{n}", toc=False, title=None, zwsp=True):
    """สร้าง PDF ภาษาไทย (ฝังฟอนต์อัตโนมัติ) — คืน (output_path, embedded_bool)"""
    from weasyprint import HTML

    cfg = dict(PRESETS.get(preset, PRESETS["default"]))
    if font_name is not None: cfg["font_name"] = font_name
    if font_size is not None: cfg["font_size"] = font_size
    if margins is not None: cfg["margins"] = margins
    if line_spacing is not None: cfg["line_spacing"] = line_spacing
    if first_line_indent is not None: cfg["first_line_indent"] = first_line_indent

    fam = resolve_font(cfg["font_name"])
    face_css, embedded = _font_face_css(fam)
    fallback = "'%s', 'TH Sarabun New', sans-serif" % fam

    body_html, toc_entries = _items_to_html(paragraphs, zwsp)
    toc_html = _toc_html(toc_entries) if toc else ""
    title_html = ("<h1 class='title'>%s</h1>" % _thai_text(title, zwsp)) if title else ""

    fs, ls = cfg["font_size"], cfg["line_spacing"]
    indent, sp = cfg.get("first_line_indent", 0), cfg.get("space_after", 0)

    css = f"""
{face_css}
{_page_css(cfg, header_text, footer_text, page_number, page_number_format)}
body {{ font-family:{'docfont' if embedded else fallback}; font-size:{fs}pt; line-height:{ls}; color:#000; }}
p.body {{ margin:0 0 {sp}pt 0; text-indent:{indent}cm; text-align:left; }}
h1.title {{ text-align:center; font-size:{fs+6}pt; font-weight:bold; margin:0 0 4pt 0; }}
p.subtitle {{ text-align:center; font-size:{fs+1}pt; color:#333; margin:0 0 12pt 0; }}
.hd {{ font-weight:bold; margin:10pt 0 4pt 0; }}
.hd1 {{ font-size:{fs+4}pt; }} .hd2 {{ font-size:{fs+2}pt; }} .hd3 {{ font-size:{fs+1}pt; }}
ul.lst, ol.lst {{ margin:0 0 {sp}pt 0; padding-left:1.2cm; }}
li {{ margin:0 0 2pt 0; }}
blockquote {{ margin:6pt 0 6pt 1cm; padding-left:0.5cm; border-left:2pt solid #bbb; color:#333; }}
p.caption {{ font-size:{fs-2}pt; color:#555; text-align:center; margin:2pt 0 8pt 0; }}
table {{ border-collapse:collapse; width:100%; margin:6pt 0; }}
th, td {{ border:0.75pt solid #555; padding:3pt 6pt; text-align:left; vertical-align:top; }}
th {{ background:#eee; font-weight:bold; }}
figure {{ text-align:center; margin:8pt 0; }}
figcaption {{ font-size:{fs-2}pt; color:#555; margin-top:2pt; }}
.pb {{ break-after:page; }}
.toc p.tocline {{ margin:0 0 3pt 0; display:flex; }}
.toc a {{ text-decoration:none; color:#000; }}
.tocline a::after {{ content:leader('.'); }}
.tocpage::before {{ content:target-counter(attr(href), page); }}
.toc2 {{ margin-left:1.2em; }} .toc3 {{ margin-left:2.4em; }}
"""
    doc_html = (f"<!DOCTYPE html><html><head><meta charset='utf-8'><style>{css}</style>"
                f"</head><body>{title_html}{toc_html}{body_html}</body></html>")
    HTML(string=doc_html, base_url=os.getcwd()).write_pdf(output_path)
    return output_path, embedded


# ─────────────────────────────────────────────────────────────────────────────
def main():
    ap = argparse.ArgumentParser(
        description="สร้าง PDF ภาษาไทยเต็มบรรทัด ฝังฟอนต์ จัดหน้าสวย (weasyprint)")
    ap.add_argument("input", nargs="?", help="ไฟล์ข้อความ/markdown")
    ap.add_argument("-o", "--output", default="output.pdf", help="ไฟล์ผลลัพธ์ (.pdf)")
    ap.add_argument("--preset", default="default", choices=list(PRESETS))
    ap.add_argument("--font", default=None, help="ชื่อฟอนต์/key (เช่น krub, sarabun)")
    ap.add_argument("--size", type=int, default=None)
    ap.add_argument("--title", default=None)
    ap.add_argument("--toc", action="store_true")
    ap.add_argument("--page-number", default=None,
                    help="footer-center | footer-right | header-right")
    ap.add_argument("--header", default=None)
    ap.add_argument("--footer", default=None)
    ap.add_argument("--no-zwsp", action="store_true", help="ไม่แทรก ZWSP (พึ่ง Pango ล้วน)")
    ap.add_argument("--list-fonts", action="store_true", help="แสดงฟอนต์ที่มากับรีโปแล้วออก")
    args = ap.parse_args()

    if args.list_fonts:
        print("ฟอนต์ที่มากับรีโป (ใช้เป็น --font <key>):")
        for k, v in list_bundled_fonts().items():
            print(f"  {k:14s} → {v}")
        return
    if not args.input:
        ap.error("ต้องระบุไฟล์ input (หรือใช้ --list-fonts)")

    with open(args.input, "r", encoding="utf-8") as f:
        raw = f.read()
    _, embedded = create_pdf(
        parse_markdown(raw), args.output, preset=args.preset, font_name=args.font,
        font_size=args.size, title=args.title, toc=args.toc,
        page_number=args.page_number, header_text=args.header,
        footer_text=args.footer, zwsp=not args.no_zwsp)
    if not _HAS_PYTHAINLP:
        print("⚠️  ไม่พบ pythainlp — ใช้ Pango ตัดบรรทัดล้วน (pip install pythainlp เพื่อผลดีขึ้น)")
    tag = "ฝังฟอนต์แล้ว" if embedded else "ใช้ฟอนต์ระบบ (ไม่พบไฟล์ฟอนต์ในรีโป)"
    print(f"✅ สร้างไฟล์: {args.output}  [{tag}]")


if __name__ == "__main__":
    main()
