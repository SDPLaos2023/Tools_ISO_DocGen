"""
Template: Risk Register
Folder 09 - Risk Management
Document ID: [CODE]-09-RR-v[VER]
Cross-references: Project Plan (01-PP), CAPA (10-CAPA), Change Request (08-CRF)
"""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from utils.doc_builder import (
    new_document, add_cover_page, add_document_control, add_version_history,
    add_section_heading, add_paragraph, add_table, add_signature_table, get_doc_id,
)

def generate(config: dict, output_path: str) -> str:
    doc_id = get_doc_id(config, "09", "RR")
    doc    = new_document()
    t      = config["team"]
    p      = config["project"]
    risks  = config.get("risks", [])

    add_cover_page(doc, config, doc_id, "Risk Register", "ทะเบียนความเสี่ยง")
    add_document_control(doc, config, doc_id)
    add_version_history(doc, config)
    doc.add_page_break()

    add_section_heading(doc, "1. Risk Management Approach", level=1)
    add_paragraph(doc, "Risk Register นี้ใช้ติดตามความเสี่ยงของโครงการตลอดช่วงเวลา "
                      "ทบทวนทุกสัปดาห์ในการประชุม Project Status Meeting "
                      "อ้างอิงกับ ISO/IEC 29110 Risk Management process")
    add_table(doc, ["ระดับความเสี่ยง","เกณฑ์","การตอบสนอง"], [
        ["High (สูง)",    "Probability × Impact ≥ 6","ต้องดำเนินการทันที — มี contingency plan"],
        ["Medium (กลาง)","Probability × Impact 3-5", "Monitor และมี mitigation plan"],
        ["Low (ต่ำ)",     "Probability × Impact ≤ 2","Accept และ monitor"],
    ], col_widths=[3,5,8])

    add_section_heading(doc, "2. Risk Matrix", level=1)
    add_paragraph(doc, "การประเมิน Risk Level = Probability × Impact")
    add_table(doc,
        headers=["","Impact: Low (1)","Impact: Medium (2)","Impact: High (3)"],
        rows=[
            ["Probability: High (3)","Medium (3)","High (6)","High (9)"],
            ["Probability: Medium (2)","Low (2)","Medium (4)","High (6)"],
            ["Probability: Low (1)","Low (1)","Low (2)","Medium (3)"],
        ],
        col_widths=[4,4,4,4]
    )

    add_section_heading(doc, "3. Risk Register", level=1)
    if risks:
        risk_rows = [[
            r.get("id",""), r.get("category",""), r.get("description",""),
            r.get("probability",""), r.get("impact",""), r.get("risk_level",""),
            r.get("mitigation",""), r.get("owner",""), r.get("status",""), r.get("review_date",""),
        ] for r in risks]
        add_table(doc,
            headers=["RISK ID","Category","Description","Prob.","Impact","Level","Mitigation","Owner","Status","Review Date"],
            rows=risk_rows,
            col_widths=[2.2,2.5,4,2,2,2,4,2.5,2,2.5]
        )
    else:
        add_paragraph(doc, "(ยังไม่มีข้อมูล risks — เพิ่มใน config.risks)")
        add_table(doc,
            headers=["RISK ID","Description","Probability","Impact","Level","Mitigation","Owner","Status"],
            rows=[
                ["RISK-001","Resource constraint — key member unavailable","Medium","High","High",
                 "Cross-train team members, document knowledge","PM","Open"],
                ["RISK-002","Requirements change late in project","Medium","Medium","Medium",
                 "Freeze requirements after sign-off, use CR process","BA","Open"],
                ["RISK-003","Integration failure with external systems","Low","High","Medium",
                 "Early integration testing, mock services","Tech Lead","Open"],
            ],
            col_widths=[2.5,4.5,3,2.5,2.5,4,3,2.5]
        )

    add_section_heading(doc, "4. Risk Detail", level=1)
    if risks:
        for r in risks:
            add_section_heading(doc, f"{r.get('id','')} — {r.get('description','')} [{r.get('risk_level','')}]", level=2)
            add_table(doc, ["หัวข้อ","รายละเอียด"], [
                ["Category",         r.get("category","")],
                ["Probability",      r.get("probability","")],
                ["Impact",           r.get("impact","")],
                ["Risk Level",       r.get("risk_level","")],
                ["Mitigation Plan",  r.get("mitigation","")],
                ["Contingency Plan", r.get("contingency","")],
                ["Owner",            r.get("owner","")],
                ["Status",           r.get("status","Open")],
                ["Review Date",      r.get("review_date","")],
                ["Linked CAPA",      r.get("linked_capa","—")],
            ], col_widths=[4,13])

    add_section_heading(doc, "5. Risk Review History", level=1)
    add_table(doc, ["Review Date","Reviewer","Changes","Next Review"], [
        [p.get("document_date",""), t.get("project_manager",{}).get("name",""),
         "Initial risk identification","สัปดาห์ถัดไป"],
    ], col_widths=[4,5,5,4])

    doc.add_page_break()
    add_signature_table(doc, [
        {"role": "จัดทำโดย / Prepared by",   "name": t.get("project_manager",{}).get("name",""),  "title": "Project Manager"},
        {"role": "ตรวจสอบโดย / Reviewed by", "name": t.get("lead_developer",{}).get("name",""),   "title": "Tech Lead"},
        {"role": "อนุมัติโดย / Approved by",  "name": "", "title": "Project Sponsor"},
    ])

    os.makedirs(output_path, exist_ok=True)
    file_path = os.path.join(output_path, f"{doc_id}_Risk_Register.docx")
    doc.save(file_path)
    return file_path
