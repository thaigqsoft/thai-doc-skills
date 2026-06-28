#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
font_embed.py — ฝังฟอนต์ TrueType ลงไฟล์ .docx ตามมาตรฐาน OOXML (ECMA-376 §17.8.1)

ทำไมต้องมีไฟล์นี้:
  python-docx ฝังฟอนต์เองไม่ได้ และ OOXML บังคับว่าฟอนต์ที่ฝังต้องถูก "obfuscate"
  (XOR 32 ไบต์แรกด้วย GUID key) ไฟล์นี้จัดการให้ครบ: obfuscate + แก้ fontTable.xml,
  เพิ่ม relationship, ใส่ไฟล์ฟอนต์เข้า zip, ประกาศ content-type และเปิด
  <w:embedTrueTypeFonts/> ใน settings.xml

ผลลัพธ์: เปิดไฟล์เครื่องไหนก็เห็นฟอนต์ถูกต้อง แม้เครื่องนั้นไม่ได้ลงฟอนต์ไทย
"""
from __future__ import annotations

import os
import re
import struct
import hashlib
import zipfile
import shutil

_HERE = os.path.dirname(os.path.abspath(__file__))
FONTS_DIR = os.path.join(_HERE, os.pardir, "fonts", "national")

_NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_REL_FONT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font"
_CT_OBFUSCATED = "application/vnd.openxmlformats-officedocument.obfuscatedFont"

# ชื่อ family ที่เอกสารอ้างถึง แต่ไฟล์ฟอนต์จริงใช้ชื่อ family อื่น
_EMBED_ALIASES = {
    "th sarabun new": "th sarabunpsk",
    "sarabun new": "th sarabunpsk",
    "sarabun": "th sarabunpsk",
}

# map subfamily → ชนิด embed element ใน fontTable.xml
_STYLE_TAG = {
    "regular": "embedRegular",
    "bold": "embedBold",
    "italic": "embedItalic",
    "bold italic": "embedBoldItalic",
}


# ── อ่านชื่อ family/subfamily จากตาราง 'name' ของไฟล์ TTF/OTF ──────────────────
def read_font_meta(path):
    """คืน (family, subfamily) จาก name table; None ถ้าไม่ใช่ sfnt"""
    with open(path, "rb") as f:
        data = f.read()
    if data[:4] not in (b"\x00\x01\x00\x00", b"true", b"OTTO"):
        return None
    num = struct.unpack(">H", data[4:6])[0]
    name_off = None
    off = 12
    for _ in range(num):
        tag = data[off:off + 4]
        o = struct.unpack(">I", data[off + 8:off + 12])[0]
        if tag == b"name":
            name_off = o
        off += 16
    if name_off is None:
        return None
    count, stroff = struct.unpack(">HH", data[name_off + 2:name_off + 6])
    base = name_off + 6
    strbase = name_off + stroff
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


_INDEX_CACHE = None


def build_font_index(fonts_dir=FONTS_DIR):
    """สแกนโฟลเดอร์ฟอนต์ → {family_lower: {subfamily_lower: path}}"""
    global _INDEX_CACHE
    if _INDEX_CACHE is not None:
        return _INDEX_CACHE
    index = {}
    if os.path.isdir(fonts_dir):
        for fn in sorted(os.listdir(fonts_dir)):
            if not fn.lower().endswith((".ttf", ".otf")):
                continue
            p = os.path.join(fonts_dir, fn)
            meta = read_font_meta(p)
            if not meta or not meta[0]:
                continue
            fam, sub = meta
            sub = (sub or "Regular").lower()
            index.setdefault(fam.lower(), {})[sub] = p
    _INDEX_CACHE = index
    return index


def resolve_font_files(display_name, fonts_dir=FONTS_DIR):
    """คืน {subfamily_lower: path} ของ family ที่ขอ (ตาม alias ด้วย) หรือ {} ถ้าไม่พบ"""
    index = build_font_index(fonts_dir)
    key = (display_name or "").strip().lower()
    key = _EMBED_ALIASES.get(key, key)
    return index.get(key, {})


# ── obfuscation (ECMA-376 §17.8.1) ───────────────────────────────────────────
def _obfuscate(font_bytes, key16):
    """XOR 32 ไบต์แรกด้วย key (key[15-(i%16)]) — involutive, ใช้ deobfuscate ได้ด้วย"""
    out = bytearray(font_bytes)
    for i in range(min(32, len(out))):
        out[i] ^= key16[15 - (i % 16)]
    return bytes(out)


def _guid_string(key16):
    """ฟอร์แมต 16 ไบต์ → '{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}' (ตรงกับ key byte order)"""
    h = key16.hex().upper()
    return "{%s-%s-%s-%s-%s}" % (h[0:8], h[8:12], h[12:16], h[16:20], h[20:32])


def _xml_escape(s):
    return (s.replace("&", "&amp;").replace("<", "&lt;")
             .replace(">", "&gt;").replace('"', "&quot;"))


# ── main entry ───────────────────────────────────────────────────────────────
def embed_fonts(docx_path, font_names, fonts_dir=FONTS_DIR):
    """
    ฝังฟอนต์ที่ระบุลงไฟล์ .docx (เขียนทับไฟล์เดิม)

    docx_path  : path ไฟล์ .docx ที่สร้างไว้แล้ว
    font_names : list ชื่อ family ที่ต้องการฝัง (เช่น ['TH Sarabun New'])
    คืน list ของ family ที่ฝังสำเร็จ
    """
    # รวบรวมไฟล์ฟอนต์ที่จะฝัง (กันชื่อซ้ำ)
    wanted = []
    seen = set()
    for name in font_names:
        if not name or name in seen:
            continue
        seen.add(name)
        files = resolve_font_files(name, fonts_dir)
        if files:
            wanted.append((name, files))

    if not wanted:
        return []

    # สร้าง entry: (display_name, {style_tag: (part_name, fontKey, obf_bytes)})
    embedded = []
    font_parts = {}   # part_name -> bytes
    rels = []         # (rId, target)
    rid = 0
    fidx = 0
    for display_name, files in wanted:
        style_map = {}
        for sub, path in files.items():
            tag = _STYLE_TAG.get(sub)
            if not tag:
                continue
            with open(path, "rb") as f:
                raw = f.read()
            key16 = hashlib.sha256(raw + sub.encode()).digest()[:16]
            obf = _obfuscate(raw, key16)
            fidx += 1
            part = "word/fonts/font%d.odttf" % fidx
            rid += 1
            rId = "rId%d" % rid
            font_parts[part] = obf
            rels.append((rId, "fonts/font%d.odttf" % fidx))
            style_map[tag] = (rId, _guid_string(key16))
        if style_map:
            embedded.append((display_name, style_map))

    if not embedded:
        return []

    # ── ประกอบ fontTable.xml ──
    ft = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
          '<w:fonts xmlns:w="%s" xmlns:r="%s">' % (_NS_W, _NS_R)]
    # ลำดับ child ต้องตาม schema: Regular, Bold, Italic, BoldItalic
    order = ["embedRegular", "embedBold", "embedItalic", "embedBoldItalic"]
    for display_name, style_map in embedded:
        ft.append('<w:font w:name="%s">' % _xml_escape(display_name))
        for tag in order:
            if tag in style_map:
                rId, key = style_map[tag]
                ft.append('<w:%s r:id="%s" w:fontKey="%s" w:subsetted="false"/>'
                          % (tag, rId, key))
        ft.append('</w:font>')
    ft.append('</w:fonts>')
    fonttable_xml = "".join(ft).encode("utf-8")

    # ── fontTable.xml.rels ──
    rl = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
    for rId, target in rels:
        rl.append('<Relationship Id="%s" Type="%s" Target="%s"/>'
                  % (rId, _REL_FONT, target))
    rl.append('</Relationships>')
    fonttable_rels = "".join(rl).encode("utf-8")

    # ── อ่าน zip เดิมทั้งหมด แล้วเขียนใหม่พร้อมแก้ไข ──
    tmp_path = docx_path + ".embed.tmp"
    with zipfile.ZipFile(docx_path, "r") as zin:
        names = zin.namelist()
        contents = {n: zin.read(n) for n in names}

    # 1) [Content_Types].xml : เพิ่ม Default สำหรับ .odttf
    ct = contents["[Content_Types].xml"].decode("utf-8")
    if "Extension=\"odttf\"" not in ct:
        inject = '<Default Extension="odttf" ContentType="%s"/>' % _CT_OBFUSCATED
        ct = ct.replace("<Default", inject + "<Default", 1) \
            if "<Default" in ct else ct.replace(">", ">" + inject, 1)
        contents["[Content_Types].xml"] = ct.encode("utf-8")

    # 2) settings.xml : เปิด embedTrueTypeFonts + saveSubsetFonts (ตามลำดับ schema)
    st = contents["word/settings.xml"].decode("utf-8")
    if "embedTrueTypeFonts" not in st:
        snippet = '<w:embedTrueTypeFonts/><w:saveSubsetFonts w:val="false"/>'
        m = re.search(r"<w:defaultTabStop\b", st)
        if m:
            st = st[:m.start()] + snippet + st[m.start():]
        else:
            st = re.sub(r"</w:settings>", snippet + "</w:settings>", st, count=1)
        contents["word/settings.xml"] = st.encode("utf-8")

    # 3) แทนที่ fontTable.xml + เพิ่ม rels + ไฟล์ฟอนต์
    contents["word/fontTable.xml"] = fonttable_xml
    contents["word/_rels/fontTable.xml.rels"] = fonttable_rels
    for part, b in font_parts.items():
        contents[part] = b

    # ── เขียน zip ใหม่ ──
    with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
        # [Content_Types].xml ควรอยู่ต้น ๆ
        ordered = ["[Content_Types].xml"] + [n for n in contents if n != "[Content_Types].xml"]
        for n in ordered:
            zout.writestr(n, contents[n])

    shutil.move(tmp_path, docx_path)
    return [name for name, _ in embedded]
