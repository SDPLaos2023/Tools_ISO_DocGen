"""
Template: Requirements Document (BRD/FRS)
Folder 02 - Requirements Analysis
Document ID pattern: [CODE]-02-BRD-v[VER]
ISO/IEC 29110 Reference: SI Process - Software Requirements
"""
import os
import sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from utils.doc_builder import (
    new_document, add_cover_page, add_document_control,
    add_version_history, add_section_heading, add_paragraph,
    add_table, add_signature_table, get_doc_id,
)


def generate(config: dict, output_path: str) -> str:
    doc_id = get_doc_id(config, "02", "BRD")
    doc    = new_document()
    p      = config["project"]
    t      = config["team"]
    reqs   = config.get("requirements", [])

    add_cover_page(doc, config, doc_id,
                   "Requirements Document", "เอกสารความต้องการระบบ (BRD/FRS)")
    add_document_control(doc, config, doc_id)
    add_version_history(doc, config)
    doc.add_page_break()

    # ── 1. Introduction ───────────────────────────────────────────────────────
    add_section_heading(doc, "1. บทนำ / Introduction", level=1)
    add_paragraph(doc, f"เอกสารฉบับนี้อธิบายความต้องการทางธุรกิจและฟังก์ชันของระบบ {p.get('name','')} "
                      f"เพื่อใช้เป็นข้อตกลงร่วมระหว่างทีมพัฒนาและผู้มีส่วนได้เสีย "
                      f"อ้างอิงตามมาตรฐาน ISO/IEC 29110.")
    add_paragraph(doc, f"This document describes business and functional system requirements for {p.get('name','')}.")

    # ── 2. Scope ─────────────────────────────────────────────────────────────
    add_section_heading(doc, "2. ขอบเขต / Scope", level=1)
    add_paragraph(doc, p.get("scope", "—"))

    # ── 3. Stakeholders ───────────────────────────────────────────────────────
    add_section_heading(doc, "3. ผู้มีส่วนได้เสีย / Stakeholders", level=1)
    stake_rows = [[s.get("name",""), s.get("role",""), s.get("responsibility","")]
                  for s in config.get("stakeholders", [])]
    if not stake_rows:
        stake_rows = [["—","—","—"]]
    add_table(doc, ["ชื่อ","บทบาท","ความรับผิดชอบ"], stake_rows, col_widths=[5,5,8])

    # ── 4. Business Requirements ──────────────────────────────────────────────
    add_section_heading(doc, "4. ความต้องการทางธุรกิจ / Business Requirements", level=1)
    func_reqs = [r for r in reqs if r.get("type","") == "Functional"]
    if not func_reqs:
        func_reqs = reqs  # show all if type not specified
    for req in func_reqs:
        add_section_heading(doc, f"{req.get('id','')} — {req.get('title','')}", level=2)
        add_table(doc,
            headers=["หัวข้อ","รายละเอียด"],
            rows=[
                ["คำอธิบาย / Description",       req.get("description","")],
                ["ลำดับความสำคัญ / Priority",    req.get("priority","Medium")],
                ["ประเภท / Type",                req.get("type","Functional")],
                ["หมวดหมู่ / Category",          req.get("category","")],
                ["แหล่งที่มา / Source",          req.get("source","")],
                ["เกณฑ์ยอมรับ / Acceptance",    req.get("acceptance_criteria","")],
            ],
            col_widths=[6, 10]
        )

    # ── 5. Non-Functional Requirements ────────────────────────────────────────
    add_section_heading(doc, "5. ความต้องการที่ไม่ใช่ฟังก์ชัน / Non-Functional Requirements", level=1)
    nfr_reqs = [r for r in reqs if r.get("type","") == "Non-Functional"]
    if nfr_reqs:
        for req in nfr_reqs:
            add_section_heading(doc, f"{req.get('id','')} — {req.get('title','')}", level=2)
            add_paragraph(doc, req.get("description",""))
    else:
        nfr_defaults = [
            ["Performance",  "ระบบต้องตอบสนองภายใน 3 วินาทีสำหรับ 95% ของ requests"],
            ["Security",     "ข้อมูลต้องเข้ารหัส (Encryption at rest and in transit)"],
            ["Availability", "ระบบต้องมี uptime ≥ 99.5% (ยกเว้นเวลา maintenance)"],
            ["Scalability",  "รองรับผู้ใช้พร้อมกันได้ไม่น้อยกว่า 100 concurrent users"],
        ]
        add_table(doc, ["หัวข้อ / Category","รายละเอียด / Detail"], nfr_defaults, col_widths=[5,12])

    # ── 6. System Interfaces ───────────────────────────────────────────────────
    add_section_heading(doc, "6. Interface กับระบบอื่น / System Interfaces", level=1)
    add_paragraph(doc, "ระบุ interface เชื่อมต่อกับระบบภายนอกหรือ API ที่เกี่ยวข้อง:")
    add_paragraph(doc, "• (กรอกข้อมูล interface ที่เกี่ยวข้อง)")

    # ── 7. Glossary ───────────────────────────────────────────────────────────
    add_section_heading(doc, "7. คำนิยาม / Glossary", level=1)
    add_table(doc, ["คำ / Term","คำอธิบาย / Definition"],
        [["BRD","Business Requirements Document"],
         ["FRS","Functional Requirements Specification"],
         ["REQ","Requirement ID prefix"],
         ["UAT","User Acceptance Testing"]],
        col_widths=[4, 13])

    # Sign-off
    doc.add_page_break()
    add_signature_table(doc, [
        {"role": "เตรียมโดย / Prepared by",     "name": t.get("business_analyst",{}).get("name",""), "title": "Business Analyst"},
        {"role": "ตรวจสอบโดย / Reviewed by",   "name": t.get("system_analyst",{}).get("name",""),    "title": "System Analyst"},
        {"role": "อนุมัติโดย / Approved by",    "name": t.get("project_manager",{}).get("name",""),  "title": "Project Manager"},
    ])

    os.makedirs(output_path, exist_ok=True)
    file_path = os.path.join(output_path, f"{doc_id}_Requirements_Document.docx")
    doc.save(file_path)
    return file_path
