"""
Template: User Manual
Folder 06 - Deployment & Training
Document ID: [CODE]-06-UM-v[VER]
"""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from utils.doc_builder import (
    new_document, add_cover_page, add_document_control, add_version_history,
    add_section_heading, add_paragraph, add_table, add_signature_table, get_doc_id,
)

def generate(config: dict, output_path: str) -> str:
    doc_id = get_doc_id(config, "06", "UM")
    doc    = new_document()
    p      = config["project"]
    t      = config["team"]
    ts     = config.get("tech_stack", {})
    reqs   = config.get("requirements", [])

    add_cover_page(doc, config, doc_id, "User Manual", "คู่มือการใช้งาน")
    add_document_control(doc, config, doc_id)
    add_version_history(doc, config)
    doc.add_page_break()

    add_section_heading(doc, "1. บทนำ / Introduction", level=1)
    add_paragraph(doc, f"คู่มือนี้อธิบายวิธีการใช้งานระบบ {p.get('name','')} "
                      "สำหรับผู้ใช้งานทั่วไปและผู้ดูแลระบบ "
                      "ครอบคลุมฟังก์ชันหลักทั้งหมดตามที่กำหนดใน Requirements Document (02-BRD)")

    add_section_heading(doc, "2. System Requirements (ความต้องการของเครื่อง)", level=1)
    add_table(doc, ["รายการ","ความต้องการขั้นต่ำ","แนะนำ"], [
        ["Operating System", "Windows 10 / macOS 12", "Windows 11 / macOS 14"],
        ["Browser",          "Chrome 100+ / Firefox 100+", "Chrome Latest"],
        ["RAM",              "4 GB", "8 GB"],
        ["Network",          "10 Mbps", "50 Mbps+"],
        ["Screen Resolution","1366x768", "1920x1080"],
    ], col_widths=[5,5,7])

    add_section_heading(doc, "3. การเข้าสู่ระบบ / Login", level=1)
    add_paragraph(doc, f"URL: {ts.get('infrastructure','https://[system-url]')}")
    steps_login = [
        "เปิด Browser และไปที่ URL ของระบบ",
        "กรอก Username (เช่น email หรือรหัสพนักงาน)",
        "กรอก Password",
        "คลิก 'เข้าสู่ระบบ' หรือกด Enter",
        "หากลืมรหัสผ่าน ให้คลิก 'ลืมรหัสผ่าน' และทำตามขั้นตอน",
    ]
    for i, step in enumerate(steps_login):
        add_paragraph(doc, f"{i+1}. {step}")

    add_section_heading(doc, "4. ฟังก์ชันหลัก / Main Features", level=1)
    func_reqs = [r for r in reqs if r.get("type","") != "Non-Functional"]
    if func_reqs:
        for req in func_reqs:
            add_section_heading(doc, f"4.x {req.get('title','')}", level=2)
            add_paragraph(doc, req.get("description",""))
            add_paragraph(doc, f"วิธีใช้งาน:")
            add_paragraph(doc, f"1. ไปที่เมนู {req.get('title', '[เมนู]')}")
            add_paragraph(doc, f"2. {req.get('acceptance_criteria', 'กรอกข้อมูลตามที่กำหนดแล้วคลิก Save')}")
            add_paragraph(doc, f"(แนบ screenshot ที่นี่)")
    else:
        add_section_heading(doc, "4.1 [ชื่อฟังก์ชัน]", level=2)
        add_paragraph(doc, "อธิบายขั้นตอนการใช้งาน:")
        for i in range(1, 5):
            add_paragraph(doc, f"{i}. [ขั้นตอนที่ {i}]")
        add_paragraph(doc, "(แนบ screenshot ที่นี่)")

    add_section_heading(doc, "5. Administration Functions (สำหรับ Admin)", level=1)
    add_paragraph(doc, "ฟังก์ชันนี้สำหรับผู้ดูแลระบบเท่านั้น:")
    admin_funcs = [
        ["จัดการผู้ใช้งาน (User Management)", "System Settings > Users > เพิ่ม/แก้ไข/ลบ"],
        ["กำหนดสิทธิ์ (Role Management)",      "System Settings > Roles > กำหนด permissions"],
        ["ดู Audit Log",                        "System Settings > Audit Log"],
        ["Backup & Restore",                    "System Settings > Maintenance"],
    ]
    add_table(doc, ["ฟังก์ชัน","ขั้นตอน/เมนู"], admin_funcs, col_widths=[6,11])

    add_section_heading(doc, "6. Error Messages & Troubleshooting", level=1)
    add_table(doc, ["Error Message","สาเหตุ","วิธีแก้ไข"], [
        ["Session expired — กรุณา Login ใหม่","Session หมดอายุ","Login เข้าระบบใหม่"],
        ["ไม่มีสิทธิ์เข้าถึง (403 Forbidden)","Role ไม่มีสิทธิ์","ติดต่อ Admin เพื่อขอสิทธิ์"],
        ["ไม่พบข้อมูล (404 Not Found)","ข้อมูลถูกลบหรือ URL ผิด","ตรวจสอบ URL หรือ refresh"],
        ["ระบบขัดข้อง (500 Server Error)","ข้อผิดพลาด server","ติดต่อ IT Support พร้อมแจ้ง error code"],
    ], col_widths=[5,5,7])

    add_section_heading(doc, "7. ช่องทางการติดต่อ / Support", level=1)
    add_table(doc, ["ช่องทาง","รายละเอียด"], [
        ["Email",    "it-support@company.com"],
        ["Phone",    "02-xxx-xxxx ต่อ xxxx"],
        ["Helpdesk", "https://helpdesk.company.com"],
        ["Hours",    "จันทร์-ศุกร์ 08:00-17:00"],
    ], col_widths=[5,12])

    add_section_heading(doc, "8. Glossary / คำอธิบายศัพท์", level=1)
    add_table(doc, ["คำ","ความหมาย"], [
        ["Admin","ผู้ดูแลระบบที่มีสิทธิ์เต็ม"],
        ["User","ผู้ใช้งานทั่วไป"],
        ["Role","กลุ่มสิทธิ์การเข้าถึง"],
        ["Session","การเชื่อมต่อที่ยังใช้งานอยู่"],
    ], col_widths=[4,13])

    os.makedirs(output_path, exist_ok=True)
    file_path = os.path.join(output_path, f"{doc_id}_User_Manual.docx")
    doc.save(file_path)
    return file_path
