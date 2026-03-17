"""
Template: Code Review Records
Folder 04 - Development
Document ID: [CODE]-04-CRR-v[VER]
Requires sign-off from reviewer
"""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from utils.doc_builder import (
    new_document, add_cover_page, add_document_control, add_version_history,
    add_section_heading, add_paragraph, add_table, add_signature_table, get_doc_id,
)

def generate(config: dict, output_path: str) -> str:
    doc_id = get_doc_id(config, "04", "CRR")
    doc    = new_document()
    t      = config["team"]
    p      = config["project"]

    add_cover_page(doc, config, doc_id, "Code Review Records", "บันทึกการตรวจสอบโค้ด")
    add_document_control(doc, config, doc_id)
    add_version_history(doc, config)
    doc.add_page_break()

    add_section_heading(doc, "1. วัตถุประสงค์ / Purpose", level=1)
    add_paragraph(doc, "เอกสารนี้บันทึกผลการ Code Review ทุกครั้งเพื่อให้สามารถตรวจสอบย้อนกลับได้ว่า "
                      "โค้ดผ่านการตรวจสอบก่อน merge เข้า main branch อยู่เสมอ")

    add_section_heading(doc, "2. Code Review Checklist", level=1)
    checklist = [
        ["โค้ดตรงตาม Coding Standards (04-CS)", "☐"],
        ["ฟังก์ชันมี docstring / comments ครบ", "☐"],
        ["ไม่มี hardcoded credentials หรือ secrets", "☐"],
        ["Error handling ครบถ้วน", "☐"],
        ["Unit tests ผ่านทั้งหมด", "☐"],
        ["ไม่มี code duplication (DRY)", "☐"],
        ["Logic และ algorithm ถูกต้อง", "☐"],
        ["Security vulnerabilities ไม่มี (SQL injection, XSS, etc.)", "☐"],
        ["Performance ไม่มี N+1 query หรือปัญหาชัดเจน", "☐"],
        ["ผ่าน linter / static analysis", "☐"],
    ]
    add_table(doc, ["รายการตรวจสอบ / Checklist Item","ผ่าน / Pass"],
              checklist, col_widths=[13,2])

    add_section_heading(doc, "3. Code Review Records", level=1)
    add_paragraph(doc, "บันทึกการ review แต่ละครั้ง:")

    # Example/template record
    record_rows = [
        ["CR-REC-001", "Feature: User Authentication",
         t.get("lead_developer",{}).get("name",""), t.get("qa_engineer",{}).get("name",""),
         p.get("document_date",""), "Approved",
         "Pass — minor comment ให้เพิ่ม null check ใน getUserById()"],
        ["CR-REC-002", "Feature: (กรอกชื่อ Feature)","","","","Pending",""],
    ]
    add_table(doc,
        headers=["Review ID","Feature / PR Title","Developer","Reviewer","Date","Status","Comments"],
        rows=record_rows,
        col_widths=[2.5,4,3,3,2.5,2,4]
    )

    add_section_heading(doc, "4. Issues Found & Resolution", level=1)
    issue_rows = [
        ["ISS-001","CR-REC-001","Medium","Null check missing in getUserById()","Fixed in commit abc1234","Closed"],
    ]
    add_table(doc,
        headers=["Issue ID","Review ID","Severity","Description","Resolution","Status"],
        rows=issue_rows,
        col_widths=[2,2,2,5,5,2]
    )

    add_section_heading(doc, "5. Summary", level=1)
    summary_rows = [
        ["Total Reviews Conducted","1"],
        ["Approved","1"],
        ["Rejected / Requires Changes","0"],
        ["Issues Found","1"],
        ["Issues Resolved","1"],
    ]
    add_table(doc, ["Metric","Count"], summary_rows, col_widths=[8,4])

    doc.add_page_break()
    add_signature_table(doc, [
        {"role": "Developer",         "name": t.get("lead_developer",{}).get("name",""), "title": "Lead Developer"},
        {"role": "Reviewer",          "name": t.get("qa_engineer",{}).get("name",""),    "title": "QA Engineer"},
        {"role": "อนุมัติ / Approved","name": t.get("project_manager",{}).get("name",""),"title": "Project Manager"},
    ])

    os.makedirs(output_path, exist_ok=True)
    file_path = os.path.join(output_path, f"{doc_id}_Code_Review_Records.docx")
    doc.save(file_path)
    return file_path
