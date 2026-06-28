#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
install_fonts.py — ติดตั้งฟอนต์แห่งชาติที่มากับรีโปนี้เข้าเครื่อง
เพื่อให้ Word/LibreOffice แสดงผลฟอนต์ได้ถูกต้องเมื่อสร้างเอกสารด้วย thai_docx

Windows : ติดตั้งแบบ per-user ไม่ต้องสิทธิ์ผู้ดูแล (admin)
macOS   : คัดลอกไป ~/Library/Fonts
Linux   : คัดลอกไป ~/.local/share/fonts แล้วรัน fc-cache

ใช้งาน:
    python scripts/install_fonts.py            # ติดตั้ง 10 ฟอนต์แห่งชาติ
    python scripts/install_fonts.py --it9       # ติดตั้งชุด IT๙ ด้วย (เลขไทยแบบราชการ)
    python scripts/install_fonts.py --list      # แสดงรายชื่อฟอนต์ที่จะติดตั้ง
"""
import os
import sys
import glob
import shutil
import argparse
import platform


def _safe_stdout():
    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def _font_files(include_it9: bool):
    here = os.path.dirname(os.path.abspath(__file__))
    fonts_root = os.path.join(here, "..", "fonts")
    files = glob.glob(os.path.join(fonts_root, "national", "*.ttf"))
    if include_it9:
        files += glob.glob(os.path.join(fonts_root, "it9-variant", "*.ttf"))
    return sorted(files)


def _full_name(path):
    """อ่านชื่อเต็มของฟอนต์ (สำหรับ registry บน Windows)"""
    try:
        from fontTools.ttLib import TTFont
        return TTFont(path)["name"].getDebugName(4) or os.path.splitext(os.path.basename(path))[0]
    except Exception:
        return os.path.splitext(os.path.basename(path))[0]


def install_windows(files):
    import ctypes
    from ctypes import wintypes
    import winreg

    local = os.environ.get("LOCALAPPDATA", "")
    dest_dir = os.path.join(local, "Microsoft", "Windows", "Fonts")
    os.makedirs(dest_dir, exist_ok=True)

    gdi32 = ctypes.WinDLL("gdi32")
    AddFontResourceW = gdi32.AddFontResourceW
    reg_key = r"Software\Microsoft\Windows NT\CurrentVersion\Fonts"

    ok = 0
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, reg_key) as key:
        for f in files:
            base = os.path.basename(f)
            dest = os.path.join(dest_dir, base)
            try:
                shutil.copy2(f, dest)
                AddFontResourceW(dest)  # ให้แอปที่เปิดอยู่เห็นฟอนต์ทันที
                winreg.SetValueEx(key, f"{_full_name(f)} (TrueType)", 0, winreg.REG_SZ, dest)
                ok += 1
            except Exception as e:
                print(f"  ! ข้าม {base}: {e}")

    # แจ้งระบบว่ามีการเปลี่ยนแปลงฟอนต์
    try:
        HWND_BROADCAST = 0xFFFF
        WM_FONTCHANGE = 0x001D
        ctypes.windll.user32.SendMessageTimeoutW(
            HWND_BROADCAST, WM_FONTCHANGE, 0, 0, 0, 1000, ctypes.byref(wintypes.DWORD())
        )
    except Exception:
        pass
    return ok, dest_dir


def install_unix(files):
    home = os.path.expanduser("~")
    if sys.platform == "darwin":
        dest_dir = os.path.join(home, "Library", "Fonts")
    else:
        dest_dir = os.path.join(home, ".local", "share", "fonts")
    os.makedirs(dest_dir, exist_ok=True)
    ok = 0
    for f in files:
        try:
            shutil.copy2(f, os.path.join(dest_dir, os.path.basename(f)))
            ok += 1
        except Exception as e:
            print(f"  ! ข้าม {os.path.basename(f)}: {e}")
    if sys.platform != "darwin":
        os.system("fc-cache -f >/dev/null 2>&1")
    return ok, dest_dir


def main():
    _safe_stdout()
    ap = argparse.ArgumentParser(description="ติดตั้งฟอนต์แห่งชาติเข้าเครื่อง")
    ap.add_argument("--it9", action="store_true", help="ติดตั้งชุด IT๙ ด้วย (ชื่อฟอนต์ซ้ำกับ Sarabun/Niramit ปกติ)")
    ap.add_argument("--list", action="store_true", help="แสดงรายชื่อแล้วออก")
    args = ap.parse_args()

    files = _font_files(args.it9)
    if not files:
        print("✗ ไม่พบไฟล์ฟอนต์ในโฟลเดอร์ fonts/ — ตรวจสอบว่ารันจากในรีโป")
        sys.exit(1)

    if args.list:
        for f in files:
            print(" -", os.path.basename(f))
        print(f"\nรวม {len(files)} ไฟล์")
        return

    if args.it9:
        print("⚠️  ชุด IT๙ ใช้ชื่อฟอนต์ภายในเดียวกับ TH SarabunPSK / TH Niramit AS ปกติ")
        print("    การติดตั้งทับอาจทำให้สองตัวปนกัน แนะนำให้ติดตั้งอย่างใดอย่างหนึ่ง")

    print(f"กำลังติดตั้ง {len(files)} ไฟล์ฟอนต์...")
    if platform.system() == "Windows":
        ok, dest = install_windows(files)
    else:
        ok, dest = install_unix(files)

    print(f"\n✅ ติดตั้งสำเร็จ {ok}/{len(files)} ไฟล์")
    print(f"   ปลายทาง: {dest}")
    print("   เปิด Word/LibreOffice ใหม่เพื่อให้เห็นฟอนต์ (ถ้ายังไม่เห็น ลอง log off / รีสตาร์ท)")


if __name__ == "__main__":
    main()
