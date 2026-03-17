"""
Template: Database Design Document (ERD)
Folder 03 - Design Architecture
Document ID: [CODE]-03-DBD-v[VER]
"""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from utils.doc_builder import (
    new_document, add_cover_page, add_document_control, add_version_history,
    add_section_heading, add_paragraph, add_table, add_signature_table, get_doc_id,
)

def generate(config: dict, output_path: str) -> str:
    doc_id = get_doc_id(config, "03", "DBD")
    doc    = new_document()
    p      = config["project"]
    t      = config["team"]
    tables = config.get("database_tables", [])
    ts     = config.get("tech_stack", {})

    add_cover_page(doc, config, doc_id, "Database Design Document", "เอกสารออกแบบฐานข้อมูล (ERD)")
    add_document_control(doc, config, doc_id)
    add_version_history(doc, config)
    doc.add_page_break()

    add_section_heading(doc, "1. บทนำ / Introduction", level=1)
    add_paragraph(doc, f"เอกสารนี้อธิบายการออกแบบฐานข้อมูลของระบบ {p.get('name','')} "
                      f"ครอบคลุม table structure, relationships และ data dictionary "
                      f"Database: {ts.get('database','—')}")

    add_section_heading(doc, "2. Database Standards", level=1)
    std_rows = [
        ["Naming Convention","snake_case สำหรับ table และ column names"],
        ["Primary Key","ทุก table มี id (INT AUTO_INCREMENT หรือ UUID)"],
        ["Timestamps","created_at, updated_at ทุก table"],
        ["Soft Delete","use is_deleted / deleted_at แทนการลบจริง"],
        ["Indexing","Index บน Foreign Keys และ columns ที่ query บ่อย"],
        ["Character Set","UTF-8 (รองรับภาษาไทย)"],
    ]
    add_table(doc, ["Standard","Detail"], std_rows, col_widths=[5,12])

    add_section_heading(doc, "3. Entity Relationship Overview", level=1)
    add_paragraph(doc, "ERD Diagram: (แนบ diagram ที่นี่ หรือดูใน design tool ที่กำหนด)")
    add_paragraph(doc, f"Tool ที่ใช้: {ts.get('erd_tool', 'dbdiagram.io')}")
    add_paragraph(doc, "Table ทั้งหมดในระบบ:")
    if tables:
        tbl_summary = [[tb.get("name",""), tb.get("description",""),
                        ", ".join(tb.get("related_requirements",[]))]
                       for tb in tables]
        add_table(doc, ["Table Name","Description","REQ Reference"], tbl_summary, col_widths=[5,8,5])
    else:
        add_paragraph(doc, "• (ยังไม่ได้กำหนด tables — กรุณาระบุใน config)")

    add_section_heading(doc, "4. Data Dictionary", level=1)
    if tables:
        for tb in tables:
            add_section_heading(doc, f"Table: {tb.get('name','')} — {tb.get('description','')}", level=2)
            col_rows = [[c.get("name",""), c.get("type",""), c.get("description","")]
                        for c in tb.get("columns", [])]
            if col_rows:
                add_table(doc, ["Column Name","Data Type","Description"], col_rows, col_widths=[5,5,8])
    else:
        add_paragraph(doc, "(ตัวอย่าง table structure — กรอกข้อมูลจริงใน config)")
        add_table(doc, ["Column Name","Data Type","Description"],
            [["id","INT PRIMARY KEY","Primary key auto-increment"],
             ["name","VARCHAR(255)","ชื่อ"],
             ["created_at","DATETIME","วันที่สร้าง record"],
             ["updated_at","DATETIME","วันที่แก้ไขล่าสุด"],
             ["is_deleted","BIT DEFAULT 0","Soft delete flag"]],
            col_widths=[5,5,8])

    add_section_heading(doc, "5. Backup & Recovery Strategy", level=1)
    backup_rows = [
        ["Full Backup","Daily (00:00)","Retain 30 days"],
        ["Incremental Backup","Every 6 hours","Retain 7 days"],
        ["Transaction Log","Continuous","Retain 72 hours"],
        ["Recovery Time Objective (RTO)","< 4 hours","—"],
        ["Recovery Point Objective (RPO)","< 1 hour","—"],
    ]
    add_table(doc, ["Backup Type","Schedule","Retention"], backup_rows, col_widths=[5,4,8])

    doc.add_page_break()
    add_signature_table(doc, [
        {"role": "เตรียมโดย", "name": t.get("system_analyst",{}).get("name",""), "title": "System Analyst"},
        {"role": "ตรวจสอบโดย","name": t.get("dba",{}).get("name",""),            "title": "DBA"},
        {"role": "อนุมัติโดย", "name": t.get("project_manager",{}).get("name",""),"title": "Project Manager"},
    ])

    os.makedirs(output_path, exist_ok=True)
    file_path = os.path.join(output_path, f"{doc_id}_Database_Design.docx")
    doc.save(file_path)
    return file_path
