"""
Template: Change Request Form
Folder 08 - Change Logs & Versioning
Document ID: [CODE]-08-CRF-v[VER]
Requires approval signature
Cross-references: Requirements (02-BRD), Risks (09-RR)
"""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from utils.doc_builder import (
    new_document, add_cover_page, add_document_control, add_version_history,
    add_section_heading, add_paragraph, add_table, add_signature_table, get_doc_id,
)

def generate(config: dict, output_path: str) -> str:
    doc_id  = get_doc_id(config, "08", "CRF")
    doc     = new_document()
    t       = config["team"]
    changes = config.get("change_requests", [])

    add_cover_page(doc, config, doc_id, "Change Request Form", "แบบฟอร์มขอเปลี่ยนแปลง")
    add_document_control(doc, config, doc_id)
    add_version_history(doc, config)
    doc.add_page_break()

    add_section_heading(doc, "1. Change Management Process", level=1)
    add_table(doc, ["ขั้นตอน","รายละเอียด","ผู้รับผิดชอบ"], [
        ["1. Submit CR","กรอกแบบฟอร์ม CR พร้อมเหตุผลและผลกระทบ","Requestor"],
        ["2. Impact Analysis","วิเคราะห์ผลกระทบต่อ Scope/Time/Cost","SA / BA"],
        ["3. Review","ทบทวนใน Change Advisory Board (CAB) Meeting","PM + Team Lead"],
        ["4. Approve/Reject","อนุมัติหรือปฏิเสธพร้อมเหตุผล","Project Manager / Sponsor"],
        ["5. Implement","ดำเนินการเปลี่ยนแปลงตาม approved plan","Development Team"],
        ["6. Verify","ทดสอบและยืนยันว่าการเปลี่ยนแปลงถูกต้อง","QA / Requestor"],
        ["7. Close","ปิด CR และอัปเดตเอกสารที่เกี่ยวข้อง","PM / BA"],
    ], col_widths=[3,9,4])

    if not changes:
        changes = [{
            "id": "CR-001",
            "title": "(Template — กรอก Change Request จริงที่นี่)",
            "description": "คำอธิบายการเปลี่ยนแปลงที่ร้องขอ",
            "requestor": "",
            "request_date": "",
            "priority": "Medium",
            "impact": "ผลกระทบต่อ scope/timeline/budget",
            "affected_documents": [],
            "status": "Pending",
            "approved_by": "",
            "approval_date": "",
            "implementation_date": "",
        }]

    for cr in changes:
        add_section_heading(doc, f"Change Request: {cr.get('id','')} — {cr.get('title','')}", level=1)

        add_section_heading(doc, "ข้อมูลทั่วไป / General Information", level=2)
        add_table(doc, ["หัวข้อ","รายละเอียด"], [
            ["CR ID",                      cr.get("id","")],
            ["ชื่อ / Title",               cr.get("title","")],
            ["ผู้ขอ / Requestor",          cr.get("requestor","")],
            ["วันที่ขอ / Request Date",    cr.get("request_date","")],
            ["ลำดับความสำคัญ / Priority",  cr.get("priority","")],
            ["สถานะ / Status",             cr.get("status","Pending")],
        ], col_widths=[6,11])

        add_section_heading(doc, "รายละเอียดการเปลี่ยนแปลง / Change Description", level=2)
        add_paragraph(doc, cr.get("description",""))

        add_section_heading(doc, "ผลกระทบ / Impact Analysis", level=2)
        add_paragraph(doc, cr.get("impact",""))
        affected = cr.get("affected_documents",[])
        if affected:
            add_paragraph(doc, f"เอกสารที่ได้รับผลกระทบ: {', '.join(affected)}")

        add_section_heading(doc, "การตัดสินใจ / Decision", level=2)
        add_table(doc, ["หัวข้อ","รายละเอียด"], [
            ["ผลการพิจารณา / Decision",    cr.get("status","Pending")],
            ["อนุมัติโดย / Approved by",   cr.get("approved_by","")],
            ["วันที่อนุมัติ / Date",        cr.get("approval_date","")],
            ["วันที่ดำเนินการ / Impl. Date",cr.get("implementation_date","")],
        ], col_widths=[6,11])

        # Individual sign-off for each CR
        doc.add_page_break()
        add_section_heading(doc, f"ลายมือชื่อ CR: {cr.get('id','')}", level=2)
        add_signature_table(doc, [
            {"role": "ผู้ขอ / Requestor",        "name": cr.get("requestor",""),            "title": "Requestor"},
            {"role": "ตรวจสอบ / Reviewed by",    "name": t.get("lead_developer",{}).get("name",""), "title": "Tech Lead"},
            {"role": "อนุมัติ / Approved by",     "name": t.get("project_manager",{}).get("name",""),"title": "Project Manager"},
        ])

    os.makedirs(output_path, exist_ok=True)
    file_path = os.path.join(output_path, f"{doc_id}_Change_Request_Form.docx")
    doc.save(file_path)
    return file_path
