"""
Template: Incident / Support Log
Folder 07 - Support & Maintenance
Document ID: [CODE]-07-ISL-v[VER]
"""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from utils.doc_builder import (
    new_document, add_cover_page, add_document_control, add_version_history,
    add_section_heading, add_paragraph, add_table, add_signature_table, get_doc_id,
)

def generate(config: dict, output_path: str) -> str:
    doc_id    = get_doc_id(config, "07", "ISL")
    doc       = new_document()
    t         = config["team"]
    p         = config["project"]
    incidents = config.get("incidents", [])

    add_cover_page(doc, config, doc_id, "Incident / Support Log", "บันทึกเหตุการณ์และการสนับสนุน")
    add_document_control(doc, config, doc_id)
    add_version_history(doc, config)
    doc.add_page_break()

    add_section_heading(doc, "1. Incident Management Process", level=1)
    add_paragraph(doc, "กระบวนการจัดการ Incident:")
    process_rows = [
        ["1. รายงาน (Report)",    "ผู้ใช้รายงาน incident ผ่าน helpdesk / email / phone"],
        ["2. บันทึก (Log)",       "Support team บันทึกใน log พร้อม ID และ severity"],
        ["3. จำแนก (Classify)",   "ประเมิน severity: Critical/High/Medium/Low"],
        ["4. กำหนดผู้รับผิดชอบ", "Assign ให้ทีมที่เกี่ยวข้อง"],
        ["5. แก้ไข (Resolve)",    "ดำเนินการแก้ไขตาม SLA"],
        ["6. ปิด (Close)",        "ยืนยันกับผู้รายงานก่อน close"],
        ["7. บทเรียน (Lesson)",   "บันทึก root cause เพื่อป้องกันซ้ำ"],
    ]
    add_table(doc, ["ขั้นตอน","รายละเอียด"], process_rows, col_widths=[5,12])

    add_section_heading(doc, "2. SLA Definition", level=1)
    sla_rows = [
        ["Critical", "ระบบล่มทั้งหมด / ข้อมูลสูญหาย", "15 นาที", "4 ชั่วโมง"],
        ["High",     "ฟังก์ชันหลักใช้งานไม่ได้",       "1 ชั่วโมง", "8 ชั่วโมง"],
        ["Medium",   "ฟังก์ชันรองมีปัญหา",             "4 ชั่วโมง", "3 วัน"],
        ["Low",      "ปัญหาเล็กน้อย / คำถาม",         "1 วัน", "5 วัน"],
    ]
    add_table(doc, ["Severity","ตัวอย่าง","Response Time","Resolution Time"], sla_rows, col_widths=[2.5,6,3.5,3.5])

    add_section_heading(doc, "3. Incident Log", level=1)
    if incidents:
        inc_rows = [[
            inc.get("id",""), inc.get("title",""), inc.get("severity",""),
            inc.get("date_reported",""), inc.get("status",""),
            inc.get("assigned_to",""), inc.get("resolved_date",""),
            inc.get("linked_change_request",""),
        ] for inc in incidents]
        add_table(doc,
            headers=["ID","Title","Severity","Date","Status","Assigned To","Resolved","CR Link"],
            rows=inc_rows,
            col_widths=[2,4,2.5,3,3,3,3,2.5]
        )
    else:
        add_paragraph(doc, "(ยังไม่มี incidents บันทึกไว้ — จะบันทึกหลัง Go-Live)")
        add_table(doc,
            headers=["INC ID","Title","Severity","Date","Status","Assigned","Resolved","CR Link"],
            rows=[["INC-001","(Placeholder)","Medium","—","Open","—","—","—"]],
            col_widths=[2,4,2.5,3,3,3,3,2.5]
        )

    add_section_heading(doc, "4. Incident Detail", level=1)
    if incidents:
        for inc in incidents:
            add_section_heading(doc, f"{inc.get('id','')} — {inc.get('title','')}", level=2)
            add_table(doc, ["หัวข้อ","รายละเอียด"], [
                ["Description",    inc.get("description","")],
                ["Root Cause",     inc.get("root_cause","")],
                ["Resolution",     inc.get("resolution","")],
            ], col_widths=[4,13])
    else:
        add_paragraph(doc, "(รายละเอียด incident จะปรากฏที่นี่)")

    add_section_heading(doc, "5. Monthly Summary", level=1)
    add_table(doc, ["เดือน","Total","Critical","High","Medium","Low","Resolved","Open"], [
        ["-", str(len(incidents)), "-", "-", "-", "-", "-", "-"],
    ], col_widths=[3,2,2,2,2,2,2,2])

    doc.add_page_break()
    add_signature_table(doc, [
        {"role": "จัดทำโดย / Prepared by",  "name": t.get("project_manager",{}).get("name",""), "title": "Project Manager"},
        {"role": "ตรวจสอบโดย / Reviewed",  "name": "", "title": "IT Manager / Department Head"},
    ])

    os.makedirs(output_path, exist_ok=True)
    file_path = os.path.join(output_path, f"{doc_id}_Incident_Support_Log.docx")
    doc.save(file_path)
    return file_path
