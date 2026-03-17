"""
Template: Audit Report
Folder 10 - Regulatory Compliance
Document ID: [CODE]-10-AR-v[VER]
References ISO Checklist (10-IC), links to CAPA (10-CAPA)
"""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from utils.doc_builder import (
    new_document, add_cover_page, add_document_control, add_version_history,
    add_section_heading, add_paragraph, add_table, add_signature_table, get_doc_id,
)

def generate(config: dict, output_path: str) -> str:
    doc_id  = get_doc_id(config, "10", "AR")
    doc     = new_document()
    t       = config["team"]
    p       = config["project"]
    audit   = config.get("audit", {})

    add_cover_page(doc, config, doc_id, "Audit Report", "รายงานการตรวจสอบ (Audit)")
    add_document_control(doc, config, doc_id)
    add_version_history(doc, config)
    doc.add_page_break()

    add_section_heading(doc, "1. Audit Information", level=1)
    add_table(doc, ["หัวข้อ","รายละเอียด"], [
        ["Audit Date",      audit.get("audit_date", p.get("document_date",""))],
        ["Auditor",         audit.get("auditor","")],
        ["Audit Scope",     audit.get("audit_scope","ISO/IEC 29110 Software Development Lifecycle")],
        ["Project",         p.get("name","")],
        ["Project Manager", t.get("project_manager",{}).get("name","")],
        ["ISO Standard",    "ISO/IEC 29110 — Software Engineering Lifecycle Profiles for VSE"],
        ["Reference",       f"ISO Checklist: {p.get('code','PROJ')}-10-IC-v{p.get('version','1.0')}"],
    ], col_widths=[5,12])

    add_section_heading(doc, "2. Audit Scope & Objectives", level=1)
    add_paragraph(doc, "วัตถุประสงค์การตรวจสอบ:")
    add_paragraph(doc, "• ตรวจสอบว่ากระบวนการพัฒนาซอฟต์แวร์สอดคล้องกับมาตรฐาน ISO/IEC 29110")
    add_paragraph(doc, "• ตรวจสอบว่าเอกสารครบถ้วนและเป็นปัจจุบัน")
    add_paragraph(doc, "• ระบุข้อบกพร่องและโอกาสปรับปรุง")

    add_section_heading(doc, "3. Documents Reviewed", level=1)
    code = p.get("code","PROJ")
    ver  = p.get("version","1.0")
    doc_list = [
        [f"{code}-01-PP-v{ver}",  "Project Plan",             "01_Project_Management"],
        [f"{code}-01-MM-v{ver}",  "Meeting Minutes",          "01_Project_Management"],
        [f"{code}-02-BRD-v{ver}", "Requirements Document",    "02_Requirements_Analysis"],
        [f"{code}-02-RTM-v{ver}", "RTM",                      "02_Requirements_Analysis"],
        [f"{code}-03-SDD-v{ver}", "System Design",            "03_Design_Architecture"],
        [f"{code}-03-DBD-v{ver}", "Database Design",          "03_Design_Architecture"],
        [f"{code}-04-CS-v{ver}",  "Coding Standards",         "04_Development"],
        [f"{code}-04-CRR-v{ver}", "Code Review Records",      "04_Development"],
        [f"{code}-05-TP-v{ver}",  "Test Plan",               "05_Testing_QA"],
        [f"{code}-05-TCR-v{ver}", "Test Cases & Results",    "05_Testing_QA"],
        [f"{code}-05-BDL-v{ver}", "Bug/Defect Log",          "05_Testing_QA"],
        [f"{code}-06-UM-v{ver}",  "User Manual",             "06_Deployment_Training"],
        [f"{code}-06-TR-v{ver}",  "Training Record",         "06_Deployment_Training"],
        [f"{code}-07-ISL-v{ver}", "Incident Log",            "07_Support_Maintenance"],
        [f"{code}-08-CRF-v{ver}", "Change Request",          "08_Change_Logs_Versioning"],
        [f"{code}-08-VRN-v{ver}", "Version Release Notes",   "08_Change_Logs_Versioning"],
        [f"{code}-09-RR-v{ver}",  "Risk Register",           "09_Risk_Management"],
        [f"{code}-10-IC-v{ver}",  "ISO Checklist",           "10_Regulatory_Compliance"],
    ]
    add_table(doc, ["Document ID","Document Name","Folder"], doc_list, col_widths=[4.5,6,7])

    add_section_heading(doc, "4. Audit Findings", level=1)
    findings = audit.get("findings", [])
    if findings:
        finding_rows = [[f.get("id",""), f.get("clause",""), f.get("finding_type",""),
                         f.get("description",""), f.get("linked_capa","")]
                        for f in findings]
        add_table(doc,
            headers=["Finding ID","ISO Clause","Type","Description","Linked CAPA"],
            rows=finding_rows, col_widths=[2.5,3,3,7,3])
    else:
        add_paragraph(doc, "(No findings at this time — ระบุ findings หลัง audit จริง)")
        add_table(doc,
            headers=["Finding ID","ISO Clause","Type","Description","Linked CAPA"],
            rows=[["FIND-001","—","Observation","(กรอก findings หลัง audit)","—"]],
            col_widths=[2.5,3,3,7,3])

    add_section_heading(doc, "5. Summary & Conclusion", level=1)
    add_table(doc, ["ประเภท","จำนวน"], [
        ["Non-Conformance (NC)",              str(len([f for f in findings if f.get("finding_type","") == "Non-Conformance"]))],
        ["Observation",                       str(len([f for f in findings if f.get("finding_type","") == "Observation"]))],
        ["Opportunity for Improvement (OFI)", str(len([f for f in findings if "Improvement" in f.get("finding_type","")]))],
        ["Total Findings",                    str(len(findings))],
    ], col_widths=[10,4])
    add_paragraph(doc, "ข้อสรุป (Conclusion):")
    add_paragraph(doc, "[กรอก overall conclusion ของการ audit เช่น: "
                      "โครงการมีความสอดคล้องกับ ISO/IEC 29110 ในระดับดี "
                      "พบข้อบกพร่องเล็กน้อยที่ได้กำหนด CAPA แล้ว]")

    add_section_heading(doc, "6. Follow-up Actions", level=1)
    add_paragraph(doc, f"CAPA สำหรับข้อบกพร่องที่พบ อ้างอิงเอกสาร: "
                      f"{code}-10-CAPA-v{ver}")

    doc.add_page_break()
    add_signature_table(doc, [
        {"role": "ผู้ตรวจสอบ / Auditor",    "name": audit.get("auditor",""), "title": "Auditor"},
        {"role": "Project Manager",          "name": t.get("project_manager",{}).get("name",""), "title": "Project Manager"},
        {"role": "อนุมัติ / Quality Manager","name": "", "title": "Quality Manager / Sponsor"},
    ])

    os.makedirs(output_path, exist_ok=True)
    file_path = os.path.join(output_path, f"{doc_id}_Audit_Report.docx")
    doc.save(file_path)
    return file_path
