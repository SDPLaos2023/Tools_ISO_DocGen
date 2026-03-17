"""
Template: ISO Checklist
Folder 10 - Regulatory Compliance
Document ID: [CODE]-10-IC-v[VER]
ISO/IEC 29110 compliance checklist
"""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from utils.doc_builder import (
    new_document, add_cover_page, add_document_control, add_version_history,
    add_section_heading, add_paragraph, add_table, add_signature_table, get_doc_id,
)

def generate(config: dict, output_path: str) -> str:
    doc_id = get_doc_id(config, "10", "IC")
    doc    = new_document()
    t      = config["team"]
    p      = config["project"]
    code   = p.get("code", "PROJ")
    ver    = p.get("version","1.0")

    add_cover_page(doc, config, doc_id, "ISO Checklist", "รายการตรวจสอบ ISO/IEC 29110")
    add_document_control(doc, config, doc_id)
    add_version_history(doc, config)
    doc.add_page_break()

    add_section_heading(doc, "1. บทนำ / Introduction", level=1)
    add_paragraph(doc, f"รายการตรวจสอบนี้ใช้ประเมินความสอดคล้องของโครงการ {p.get('name','')} "
                      "กับมาตรฐาน ISO/IEC 29110 (Software Engineering – Lifecycle profiles for VSE) "
                      "ใช้ประกอบการ Audit ในเอกสาร 10-AR")

    add_section_heading(doc, "2. PM Process Checklist", level=1)
    add_paragraph(doc, "Project Management Process")
    pm_items = [
        [f"PM-01","มี Project Plan ที่อนุมัติแล้ว",f"{code}-01-PP-v{ver}","☐ Yes  ☐ No",""],
        [f"PM-02","Project Plan มี schedule, milestones, resources", f"{code}-01-PP-v{ver}","☐ Yes  ☐ No",""],
        [f"PM-03","มีการประชุมและบันทึก Meeting Minutes",f"{code}-01-MM-v{ver}","☐ Yes  ☐ No",""],
        [f"PM-04","มีการติดตาม progress เทียบ plan",f"{code}-01-PP-v{ver}","☐ Yes  ☐ No",""],
        [f"PM-05","มีการบริหารความเสี่ยง (Risk Register)",f"{code}-09-RR-v{ver}","☐ Yes  ☐ No",""],
        [f"PM-06","Stakeholders รับทราบ project plan","—","☐ Yes  ☐ No",""],
    ]
    add_table(doc, ["Item","Requirement","Reference Doc","Compliant","Remarks"],
              pm_items, col_widths=[1.5,6,4,3,4])

    add_section_heading(doc, "3. SI Process Checklist", level=1)
    add_paragraph(doc, "Software Implementation Process")
    si_items = [
        [f"SI-01","มี Requirements Document (BRD/FRS) ที่อนุมัติแล้ว",f"{code}-02-BRD-v{ver}","☐ Yes  ☐ No",""],
        [f"SI-02","มี Requirements Traceability Matrix (RTM)",f"{code}-02-RTM-v{ver}","☐ Yes  ☐ No",""],
        [f"SI-03","ทุก Requirement มี Test Case อ้างอิง",f"{code}-02-RTM-v{ver}","☐ Yes  ☐ No",""],
        [f"SI-04","มี System Design Document",f"{code}-03-SDD-v{ver}","☐ Yes  ☐ No",""],
        [f"SI-05","มี Database Design Document",f"{code}-03-DBD-v{ver}","☐ Yes  ☐ No",""],
        [f"SI-06","มี Coding Standards ที่ทีมปฏิบัติตาม",f"{code}-04-CS-v{ver}","☐ Yes  ☐ No",""],
        [f"SI-07","มีการ Code Review ทุก feature",f"{code}-04-CRR-v{ver}","☐ Yes  ☐ No",""],
        [f"SI-08","มี Test Plan",f"{code}-05-TP-v{ver}","☐ Yes  ☐ No",""],
        [f"SI-09","มี Test Cases พร้อมผลการทดสอบ",f"{code}-05-TCR-v{ver}","☐ Yes  ☐ No",""],
        [f"SI-10","Test Pass Rate ≥ 95%",f"{code}-05-TCR-v{ver}","☐ Yes  ☐ No",""],
        [f"SI-11","มีการบันทึก Defects และ resolution",f"{code}-05-BDL-v{ver}","☐ Yes  ☐ No",""],
        [f"SI-12","มี User Manual",f"{code}-06-UM-v{ver}","☐ Yes  ☐ No",""],
        [f"SI-13","มีการ Training ผู้ใช้งานก่อน Go-Live",f"{code}-06-TR-v{ver}","☐ Yes  ☐ No",""],
    ]
    add_table(doc, ["Item","Requirement","Reference Doc","Compliant","Remarks"],
              si_items, col_widths=[1.5,6,4,3,4])

    add_section_heading(doc, "4. Maintenance & Support Checklist", level=1)
    maint_items = [
        ["MS-01","มี Incident/Support Log",f"{code}-07-ISL-v{ver}","☐ Yes  ☐ No",""],
        ["MS-02","มีกระบวนการ Change Management",f"{code}-08-CRF-v{ver}","☐ Yes  ☐ No",""],
        ["MS-03","มี Version Release Notes",f"{code}-08-VRN-v{ver}","☐ Yes  ☐ No",""],
        ["MS-04","มีการทบทวน Risk Register สม่ำเสมอ",f"{code}-09-RR-v{ver}","☐ Yes  ☐ No",""],
    ]
    add_table(doc, ["Item","Requirement","Reference Doc","Compliant","Remarks"],
              maint_items, col_widths=[1.5,6,4,3,4])

    add_section_heading(doc, "5. Summary", level=1)
    add_table(doc, ["Process","Total Items","Compliant","Non-Compliant","Compliance %"], [
        ["PM Process",    str(len(pm_items)),   "—","—","—%"],
        ["SI Process",    str(len(si_items)),   "—","—","—%"],
        ["Support/Maint", str(len(maint_items)),"—","—","—%"],
        ["TOTAL",         str(len(pm_items)+len(si_items)+len(maint_items)),"—","—","—%"],
    ], col_widths=[4,3,3,3,4])

    doc.add_page_break()
    add_signature_table(doc, [
        {"role": "ตรวจสอบโดย / Auditor",     "name": "", "title": "Internal / External Auditor"},
        {"role": "รับทราบ / PM",             "name": t.get("project_manager",{}).get("name",""), "title": "Project Manager"},
        {"role": "อนุมัติ / Approved by",    "name": "", "title": "Quality Manager / Sponsor"},
    ])

    os.makedirs(output_path, exist_ok=True)
    file_path = os.path.join(output_path, f"{doc_id}_ISO_Checklist.docx")
    doc.save(file_path)
    return file_path
