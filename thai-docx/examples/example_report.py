# -*- coding: utf-8 -*-
"""ตัวอย่างการใช้งาน thai_docx — เอกสารราชการพร้อมสารบัญ ตาราง หัว-ท้ายกระดาษ เลขหน้า"""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts"))
from thai_docx import create_docx

paragraphs = [
    {"text": "บทที่ ๑ บทนำ", "type": "heading1"},
    {"text": "เอกสารฉบับนี้จัดทำขึ้นเพื่อทดสอบการสร้างไฟล์ Word ภาษาไทยที่เขียน"
             "เต็มบรรทัดโดยไม่ตัดบรรทัดก่อนเวลา ซึ่งเป็นปัญหาที่พบบ่อยเมื่อสร้าง"
             "เอกสารภาษาไทยด้วยโปรแกรมอัตโนมัติ ข้อความนี้ควรไหลเต็มความกว้างของหน้า"
             "กระดาษอย่างเป็นธรรมชาติ และคำที่เป็น **ตัวหนา** ก็ต้องแสดงผลถูกต้อง", "type": "body"},
    {"text": "๑.๑ วัตถุประสงค์", "type": "heading2"},
    {"text": "เพื่อให้การจัดทำเอกสารราชการเป็นไปอย่างมีประสิทธิภาพ", "type": "bullet"},
    {"text": "เพื่อทดสอบฟอนต์ฝั่ง Complex Script ของอักษรไทย", "type": "bullet"},
    {"text": "ตารางสรุปงบประมาณ", "type": "heading2"},
    {"type": "table", "header": True, "rows": [
        ["รายการ", "งบประมาณ (บาท)", "หมายเหตุ"],
        ["ครุภัณฑ์สำนักงาน", "1,250,000", "จัดซื้อใหม่"],
        ["ค่าใช้สอย", "880,500", "ประจำปี"],
    ]},
]

out = os.path.join(os.path.dirname(__file__), "example_report.docx")
create_docx(
    paragraphs, out,
    preset="saraban",
    title="รายงานการทดสอบระบบจัดทำเอกสาร",
    toc=True,
    header_text="สำนักงบประมาณ",
    page_number="footer-center",
    page_number_format="หน้า {n}",
)
print("created:", out)
