#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
thai_pdf.py — สร้าง PDF ภาษาไทยเต็มบรรทัด ฟอนต์ฝังในไฟล์ จัดหน้าระดับเอกสารราชการ

ใช้ weasyprint (HTML/CSS → PDF) ซึ่งจัดการ 2 เรื่องยากของไทยให้อัตโนมัติ:
  1) ตัดบรรทัดไทย (ไม่มีช่องว่างระหว่างคำ) — ผ่าน Pango/libthai + เสริม ZWSP จาก pythainlp
  2) ฝังฟอนต์ลง PDF (subset อัตโนมัติ) — เปิดเครื่องไหนก็เห็นฟอนต์ถูกต้อง

ใช้ paragraph model / preset / ฟอนต์แห่งชาติ ชุดเดียวกับ thai_docx.py
"""
from __future__ import annotations

import os
import re
import html as _html
import argparse

# reuse config/helpers จาก thai_docx
try:
    from thai_docx import PRESETS, parse_markdown, insert_zwsp, resolve_font
except ImportError:  # pragma: no cover
    from .thai_docx import PRESETS, parse_markdown, insert_zwsp, resolve_font  # type: ignore
try:
    from font_embed import resolve_font_files
except ImportError:  # pragma: no cover
    from .font_embed import resolve_font_files  # type: ignore


# ── markdown inline (**bold**, *italic*) → HTML ──────────────────────────────
def _inline(text):
    s = _html.escape(text, quote=False)
    s = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", s)
    s = re.sub(r"(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)", r"<em>\1</em>", s)
    return s


def _thai_text(text, zwsp=True):
    """แทรก ZWSP ช่วยตัดบรรทัด แล้วแปลง markdown inline"""
    if zwsp and text:
        # ใส่ ZWSP ก่อน escape (insert_zwsp ไม่ยุ่งกับ * markdown)
        text = insert_zwsp(text)
    return _inline(text)


# ── @font-face: ฝังฟอนต์แห่งชาติทั้ง 4 สไตล์ ──────────────────────────────────
_STYLE_CSS = {
    "regular": ("normal", "normal"),
    "bold": ("bold", "normal"),
    "italic": ("normal", "italic"),
    "bold italic": ("bold", "italic"),
}


def _font_face_css(family_display):
    """สร้าง @font-face ของ family ที่ขอ (ถ้ามีไฟล์ในรีโป) คืน (css, ใช้ฟอนต์ฝังไหม)"""
    files = resolve_font_files(family_display)
    if not files:
        return "", False
    rules = []
    for sub, path in files.items():
        weight, style = _STYLE_CSS.get(sub, ("normal", "normal"))
        url = "file://" + os.path.abspath(path).replace(" ", "%20")
        rules.append(
            "@font-face{font-family:'docfont';src:url('%s');"
            "font-weight:%s;font-style:%s;}" % (url, weight, style))
    return "\n".join(rules), True


# ── แปลง paragraph item → HTML ───────────────────────────────────────────────
def _items_to_html(paragraphs, zwsp=True):
    out = []
    toc_entries = []      # (level, text, anchor)
    hcount = 0
    i = 0
    n = len(paragraphs)
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
            lvl = int(t[-1])
            hcount += 1
            anc = "h%d" % hcount
            toc_entries.append((lvl, it.get("text", ""), anc))
            out.append("<h%d id='%s' class='hd hd%d'>%s</h%d>"
                       % (lvl + 1, anc, lvl, _thai_text(it.get("text", ""), zwsp), lvl + 1))
        elif t == "quote":
            out.append("<blockquote>%s</blockquote>" % _thai_text(it.get("text", ""), zwsp))
        elif t == "caption":
            out.append("<p class='caption'>%s</p>" % _thai_text(it.get("text", ""), zwsp))
        elif t == "pagebreak":
            out.append("<div class='pb'></div>")
        elif t == "table":
            rows = it.get("rows", [])
            hdr = it.get("header", False)
            out.append("<table><tbody>")
            for r, row in enumerate(rows):
                cell = "th" if (hdr and r == 0) else "td"
                out.append("<tr>" + "".join(
                    "<%s>%s</%s>" % (cell, _thai_text(str(c), zwsp), cell) for c in row) + "</tr>")
            out.append("</tbody></table>")
        elif t == "image":
            path = it.get("path", "")
            w = it.get("width")
            style = ("width:%scm;" % w) if w else ""
            url = "file://" + os.path.abspath(path).replace(" ", "%20") if path else ""
            out.append("<figure><img src='%s' style='%s'/>" % (url, style))
            if it.get("caption"):
                out.append("<figcaption>%s</figcaption>" % _thai_text(it["caption"], zwsp))
            out.append("</figure>")
        else:  # body
            out.append("<p class='body'>%s</p>" % _thai_text(it.get("text", ""), zwsp))
        i += 1
    return "\n".join(out), toc_entries


def _toc_html(entries, title="สารบัญ"):
    if not entries:
        return ""
    rows = ["<h2 class='hd hd1 toctitle'>%s</h2>" % _html.escape(title), "<div class='toc'>"]
    for lvl, text, anc in entries:
        rows.append(
            "<p class='tocline toc%d'><a href='#%s'>%s</a>"
            "<span class='tocpage'></span></p>" % (lvl, anc, _inline(text)))
    rows.append("</div><div class='pb'></div>")
    return "\n".join(rows)


def _page_css(cfg, header_text, footer_text, page_number, page_number_format):
    m = cfg["margins"]
    size = cfg.get("page_size", "A4")
    # margin boxes
    boxes = []

    def _num_content(fmt):
        c = _html.escape(fmt)
        c = c.replace("{n}", "\" counter(page) \"").replace("{total}", "\" counter(pages) \"")
        return '"%s"' % c

    pos_map = {"footer-center": "@bottom-center", "footer-right": "@bottom-right",
               "header-right": "@top-right"}
    if page_number in pos_map:
        boxes.append("%s{content:%s;font-family:'docfont';font-size:%dpt;}"
                     % (pos_map[page_number], _num_content(page_number_format),
                        cfg["font_size"] - 2))
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
    """สร้าง PDF ภาษาไทย (ฝังฟอนต์อัตโนมัติ) — พารามิเตอร์ล้อตาม create_docx ของ thai_docx"""
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

    fs = cfg["font_size"]
    ls = cfg["line_spacing"]
    indent = cfg.get("first_line_indent", 0)
    sp_after = cfg.get("space_after", 0)

    css = f"""
{face_css}
{_page_css(cfg, header_text, footer_text, page_number, page_number_format)}
body {{ font-family:{'docfont' if embedded else fallback}; font-size:{fs}pt;
        line-height:{ls}; color:#000; }}
p.body {{ margin:0 0 {sp_after}pt 0; text-indent:{indent}cm; text-align:left; }}
h1.title {{ text-align:center; font-size:{fs+6}pt; font-weight:bold; margin:0 0 4pt 0; }}
p.subtitle {{ text-align:center; font-size:{fs+1}pt; color:#333; margin:0 0 12pt 0; }}
.hd {{ font-weight:bold; margin:10pt 0 4pt 0; }}
.hd1 {{ font-size:{fs+4}pt; }} .hd2 {{ font-size:{fs+2}pt; }} .hd3 {{ font-size:{fs+1}pt; }}
ul.lst, ol.lst {{ margin:0 0 {sp_after}pt 0; padding-left:1.2cm; }}
li {{ margin:0 0 2pt 0; }}
blockquote {{ margin:6pt 0 6pt 1cm; padding-left:0.5cm; border-left:2pt solid #bbb; color:#333; }}
p.caption {{ font-size:{fs-2}pt; color:#555; text-align:center; margin:2pt 0 8pt 0; }}
table {{ border-collapse:collapse; width:100%; margin:6pt 0; }}
th, td {{ border:0.75pt solid #555; padding:3pt 6pt; text-align:left; vertical-align:top; }}
th {{ background:#eee; font-weight:bold; }}
figure {{ text-align:center; margin:8pt 0; }}
figcaption {{ font-size:{fs-2}pt; color:#555; margin-top:2pt; }}
.pb {{ break-after:page; }}
.toc p.tocline {{ margin:0 0 3pt 0; }}
.toc a {{ text-decoration:none; color:#000; }}
.tocline {{ display:flex; }}
.tocline a::after {{ content:leader('.'); }}
.tocpage::before {{ content:target-counter(attr(href), page); }}
.toc2 {{ margin-left:1.2em; }} .toc3 {{ margin-left:2.4em; }}
"""
    # ใช้ target-counter ต้องให้ a มี href; ปรับ tocline ให้ span ดึงเลขหน้า
    doc_html = f"""<!DOCTYPE html><html><head><meta charset='utf-8'>
<style>{css}</style></head><body>
{title_html}
{toc_html}
{body_html}
</body></html>"""

    HTML(string=doc_html, base_url=os.getcwd()).write_pdf(output_path)
    return output_path, embedded


# ── CLI ──────────────────────────────────────────────────────────────────────
def main():
    ap = argparse.ArgumentParser(
        description="สร้าง PDF ภาษาไทยเต็มบรรทัด ฝังฟอนต์ จัดหน้าสวย (weasyprint)")
    ap.add_argument("input", nargs="?", help="ไฟล์ข้อความ/markdown")
    ap.add_argument("-o", "--output", default="output.pdf", help="ไฟล์ผลลัพธ์ (.pdf)")
    ap.add_argument("--preset", default="default", choices=list(PRESETS))
    ap.add_argument("--font", default=None)
    ap.add_argument("--size", type=int, default=None)
    ap.add_argument("--title", default=None)
    ap.add_argument("--toc", action="store_true")
    ap.add_argument("--page-number", default=None,
                    help="footer-center | footer-right | header-right")
    ap.add_argument("--header", default=None)
    ap.add_argument("--footer", default=None)
    ap.add_argument("--no-zwsp", action="store_true", help="ไม่แทรก ZWSP (พึ่ง Pango ล้วน)")
    args = ap.parse_args()

    if not args.input:
        ap.error("ต้องระบุไฟล์ input")
    with open(args.input, "r", encoding="utf-8") as f:
        raw = f.read()
    paragraphs = parse_markdown(raw)
    _, embedded = create_pdf(
        paragraphs, args.output, preset=args.preset, font_name=args.font,
        font_size=args.size, title=args.title, toc=args.toc,
        page_number=args.page_number, header_text=args.header,
        footer_text=args.footer, zwsp=not args.no_zwsp)
    tag = "ฝังฟอนต์แล้ว" if embedded else "ใช้ฟอนต์ระบบ (ไม่พบไฟล์ฟอนต์ในรีโป)"
    print(f"✅ สร้างไฟล์: {args.output}  [{tag}]")


if __name__ == "__main__":
    main()
