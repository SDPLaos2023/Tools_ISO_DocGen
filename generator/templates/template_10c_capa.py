"""
Template: CAPA (Corrective Action and Preventive Action)
Folder 10 - Regulatory Compliance
Document ID: [CODE]-10-CAPA-v[VER]
Cross-references: Audit Report (10-AR), Risk Register (09-RR)
"""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from utils.doc_builder import (
    new_document, add_cover_page, add_document_control, add_version_history,
    add_section_heading, add_paragraph, add_table, add_signature_table, get_doc_id,
)

def generate(config: dict, output_path: str) -> str:
    doc_id = get_doc_id(config, "10", "CAPA")
    doc    = new_document()
    t      = config["team"]
    p      = config["project"]
    capas  = config.get("capas", [])
    audit  = config.get("audit", {})

    add_cover_page(doc, config, doc_id, "CAPA", "แผนการแก้ไขและป้องกัน (Corrective & Preventive Action)")
    add_document_control(doc, config, doc_id)
    add_version_history(doc, config)
    doc.add_page_break()

    add_section_heading(doc, "1. วัตถุประสงค์ / Purpose", level=1)
    add_paragraph(doc, "CAPA (Corrective Action and Preventive Action) ใช้บันทึกและติดตาม "
                      "การดำเนินการแก้ไขข้อบกพร่องที่พบจากการ Audit หรือจากการวิเคราะห์ความเสี่ยง "
                      "อ้างอิงเอกสาร: Audit Report (10-AR), Risk Register (09-RR)")

    add_section_heading(doc, "2. CAPA Process", level=1)
    add_table(doc, ["ขั้นตอน","รายละเอียด"], [
        ["1. Identify",    "ระบุปัญหาจาก Audit Finding / Risk / Incident"],
        ["2. Root Cause",  "วิเคราะห์สาเหตุที่แท้จริง (5-Why / Fishbone)"],
        ["3. Action Plan", "กำหนดแผนการแก้ไข/ป้องกัน พร้อม owner และ deadline"],
        ["4. Implement",   "ดำเนินการตาม action plan"],
        ["5. Verify",      "ตรวจสอบประสิทธิภาพของการแก้ไข"],
        ["6. Close",       "ปิด CAPA เมื่อยืนยันว่าปัญหาแก้ไขแล้ว"],
    ], col_widths=[3,14])

    add_section_heading(doc, "3. CAPA Register", level=1)
    if capas:
        capa_rows = [[c.get("id",""), c.get("type",""), c.get("related_finding",""),
                      c.get("description",""), c.get("responsible",""),
                      c.get("target_date",""), c.get("status",""), c.get("closed_date","")]
                     for c in capas]
        add_table(doc,
            headers=["CAPA ID","Type","Finding/Risk","Description","Owner","Target Date","Status","Closed Date"],
            rows=capa_rows,
            col_widths=[2.5,2.5,2.5,5,3,3,2.5,3]
        )
    else:
        add_paragraph(doc, "(ยังไม่มี CAPA — จะสร้างหลังจาก Audit พบข้อบกพร่อง)")
        add_table(doc,
            headers=["CAPA ID","Type","Finding/Risk","Description","Owner","Target Date","Status"],
            rows=[["CAPA-001","Corrective Action","FIND-001","(กรอกเมื่อมี findings)","PM","—","Open"]],
            col_widths=[2.5,3,2.5,5.5,3,3,3]
        )

    add_section_heading(doc, "4. CAPA Detail", level=1)
    if capas:
        for c in capas:
            add_section_heading(doc, f"{c.get('id','')} — {c.get('type','')} [{c.get('status','')}]", level=2)
            add_table(doc, ["หัวข้อ","รายละเอียด"], [
                ["Type",                    c.get("type","")],
                ["Related Finding / Risk",  c.get("related_finding","")],
                ["Description",             c.get("description","")],
                ["Root Cause Analysis",     c.get("root_cause","")],
                ["Action Plan",             c.get("action_plan","")],
                ["Responsible Person",      c.get("responsible","")],
                ["Target Date",             c.get("target_date","")],
                ["Effectiveness Review",    c.get("effectiveness_review","")],
                ["Status",                  c.get("status","Open")],
                ["Closed Date",             c.get("closed_date","")],
            ], col_widths=[4,13])
    else:
        add_paragraph(doc, "(รายละเอียด CAPA แต่ละรายการจะปรากฏที่นี่หลังจากมี findings)")

    add_section_heading(doc, "5. Effectiveness Summary", level=1)
    closed_capas = len([c for c in capas if c.get("status","") == "Closed"])
    open_capas   = len([c for c in capas if c.get("status","") != "Closed"])
    add_table(doc, ["Metric","Count"], [
        ["Total CAPAs",   str(len(capas)) if capas else "0"],
        ["Closed",        str(closed_capas)],
        ["Open/In Progress", str(open_capas)],
        ["On-time Closure Rate","—%"],
    ], col_widths=[8,4])

    doc.add_page_break()
    add_signature_table(doc, [
        {"role": "จัดทำโดย / Prepared by",   "name": t.get("project_manager",{}).get("name",""),  "title": "Project Manager"},
        {"role": "ตรวจสอบโดย / Reviewed by", "name": audit.get("auditor",""),                     "title": "Auditor"},
        {"role": "อนุมัติ / Approved by",     "name": "", "title": "Quality Manager / Sponsor"},
    ])

    os.makedirs(output_path, exist_ok=True)
    file_path = os.path.join(output_path, f"{doc_id}_CAPA.docx")
    doc.save(file_path)
    return file_path
