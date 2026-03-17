"""
Template: Test Plan
Folder 05 - Testing QA
Document ID: [CODE]-05-TP-v[VER]
"""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from utils.doc_builder import (
    new_document, add_cover_page, add_document_control, add_version_history,
    add_section_heading, add_paragraph, add_table, add_signature_table, get_doc_id,
)

def generate(config: dict, output_path: str) -> str:
    doc_id = get_doc_id(config, "05", "TP")
    doc    = new_document()
    p      = config["project"]
    t      = config["team"]
    ts     = config.get("tech_stack", {})
    reqs   = config.get("requirements", [])
    backend = ts.get("backend", "")
    # Determine testing frameworks from tech stack
    if ".NET" in backend or "C#" in backend:
        unit_fw  = "NUnit / MSTest / FluentAssertions"
        api_fw   = "Postman"
        ui_fw    = "Playwright"
        perf_fw  = "k6"
    else:
        unit_fw  = "[e.g. Jest / pytest / NUnit]"
        api_fw   = "[e.g. Postman / RestAssured]"
        ui_fw    = "[e.g. Selenium / Playwright]"
        perf_fw  = "[e.g. JMeter / k6]"

    add_cover_page(doc, config, doc_id, "Test Plan", "แผนการทดสอบ")
    add_document_control(doc, config, doc_id)
    add_version_history(doc, config)
    doc.add_page_break()

    add_section_heading(doc, "1. บทนำ / Introduction", level=1)
    add_paragraph(doc, f"เอกสารนี้อธิบายแผนการทดสอบสำหรับระบบ {p.get('name','')} "
                      "ครอบคลุมกลยุทธ์การทดสอบ ขอบเขต เกณฑ์ผ่าน/ไม่ผ่าน และทรัพยากรที่ใช้ "
                      "อ้างอิงจาก RTM (เอกสาร 02-RTM)")

    add_section_heading(doc, "2. Test Scope", level=1)
    add_section_heading(doc, "2.1 In Scope", level=2)
    add_paragraph(doc, "• ทุก Functional Requirements ที่ระบุใน 02-BRD")
    add_paragraph(doc, "• Non-Functional Requirements: Performance, Security, Usability")
    add_paragraph(doc, "• Integration ระหว่าง components ตาม 03-SDD")
    add_section_heading(doc, "2.2 Out of Scope", level=2)
    add_paragraph(doc, "• Third-party service testing (ทดสอบเฉพาะ integration point)")
    add_paragraph(doc, "• Load testing เกิน 500 concurrent users (ถ้าไม่ระบุใน requirements)")

    add_section_heading(doc, "3. Test Strategy", level=1)
    strategy_rows = [
        ["Unit Testing",       "นักพัฒนา",     "ทุก module/function", "≥ 80% code coverage"],
        ["Integration Testing","QA Engineer",  "API และ component integration","ทุก interface"],
        ["System Testing",     "QA Team",      "End-to-end workflows", "ตาม Requirements"],
        ["UAT",                "End Users",    "Business scenarios","ผู้ใช้ยืนยันผ่าน"],
        ["Performance Testing","QA Engineer",  "Load ≥ 100 users","Response < 3s"],
        ["Security Testing",   "Security Team","OWASP Top 10","ไม่มี Critical/High vuln"],
        ["Regression Testing", "QA Team",      "ทุก release","Automated test pass"],
    ]
    add_table(doc, ["Test Type","ผู้รับผิดชอบ","Coverage","Criteria"],
              strategy_rows, col_widths=[3.5,3.5,5,5])

    add_section_heading(doc, "4. Test Environment", level=1)
    env_rows = [
        ["Development","localhost","Unit testing ระหว่าง development"],
        ["Testing/QA",  "test.server.internal","Integration & System testing"],
        ["Staging",     "staging.server.internal","UAT & Pre-production testing"],
        ["Production",  "prod.server.internal","Smoke test หลัง deployment"],
    ]
    add_table(doc, ["Environment","Server/URL","Purpose"], env_rows, col_widths=[4,5,8])

    add_section_heading(doc, "5. Entry & Exit Criteria", level=1)
    add_section_heading(doc, "5.1 Entry Criteria (เงื่อนไขเริ่มทดสอบ)", level=2)
    add_paragraph(doc, "• Requirements Sign-off ครบ (เอกสาร 02-BRD)")
    add_paragraph(doc, "• Test Cases พร้อมแล้ว (เอกสาร 05-TCR)")
    add_paragraph(doc, "• Test Environment พร้อมใช้งาน")
    add_paragraph(doc, "• Build ผ่าน Unit Tests ≥ 80%")
    add_section_heading(doc, "5.2 Exit Criteria (เงื่อนไขผ่านการทดสอบ)", level=2)
    add_paragraph(doc, "• Test Case ผ่าน ≥ 95%")
    add_paragraph(doc, "• ไม่มี Critical/High severity defects ที่ยังเปิดอยู่")
    add_paragraph(doc, "• Performance ผ่านตาม SLA")
    add_paragraph(doc, "• UAT Sign-off จาก Business owner")

    add_section_heading(doc, "6. Requirements Coverage", level=1)
    req_cov_rows = [[req.get("id",""), req.get("title",""), req.get("priority",""),
                     ", ".join(req.get("linked_test_cases",["—"]))]
                    for req in reqs]
    if not req_cov_rows:
        req_cov_rows = [["REQ-001","(Placeholder)","High","TC-001"]]
    add_table(doc, ["REQ ID","Title","Priority","Test Cases"], req_cov_rows, col_widths=[2.5,7,3,5])

    add_section_heading(doc, "7. Test Schedule", level=1)
    milestones = config.get("milestones", [])
    ms_map = {ms.get("id",""): ms for ms in milestones}
    test_sched = [
        ["Unit Testing",       "ระหว่าง development","MS-03"],
        ["Integration Testing","หลัง MS-03 (Dev Complete)","MS-03"],
        ["System Testing",     "2 สัปดาห์หลัง Integration","MS-04"],
        ["UAT",                "3 สัปดาห์ก่อน Go-Live","MS-04"],
        ["Performance Testing","1 สัปดาห์ก่อน Go-Live","MS-05"],
    ]
    add_table(doc, ["Test Phase","Period","Linked Milestone"], test_sched, col_widths=[4,8,3])

    add_section_heading(doc, "8. Tools", level=1)
    add_table(doc, ["Tool","Purpose"], [
        [unit_fw,                      "Unit & Integration Testing framework"],
        [api_fw,                       "API Testing"],
        [ui_fw,                        "UI Automation Testing"],
        [perf_fw,                      "Performance Testing"],
        ["Excel / TestRail / Zephyr",  "Test Case Management"],
        ["Jira / Azure DevOps",        "Defect Tracking"],
    ], col_widths=[7,10])

    doc.add_page_break()
    add_signature_table(doc, [
        {"role": "เตรียมโดย",  "name": t.get("qa_engineer",{}).get("name",""),      "title": "QA Engineer"},
        {"role": "ตรวจสอบโดย","name": t.get("lead_developer",{}).get("name",""),    "title": "Lead Developer"},
        {"role": "อนุมัติโดย", "name": t.get("project_manager",{}).get("name",""),  "title": "Project Manager"},
    ])

    os.makedirs(output_path, exist_ok=True)
    file_path = os.path.join(output_path, f"{doc_id}_Test_Plan.docx")
    doc.save(file_path)
    return file_path
