"""
Template: Requirements Traceability Matrix (RTM)
Folder 02 - Requirements Analysis
Document ID pattern: [CODE]-02-RTM-v[VER]
ISO/IEC 29110 Reference: SI Process - Traceability
Cross-references: Requirements → Design Components → Test Cases → Defects
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
    doc_id = get_doc_id(config, "02", "RTM")
    doc    = new_document()
    t      = config["team"]
    reqs   = config.get("requirements", [])
    comps  = config.get("design_components", [])
    tcs    = config.get("test_cases", [])
    defects= config.get("defects", [])

    add_cover_page(doc, config, doc_id,
                   "Requirements Traceability Matrix", "เมทริกซ์การตรวจสอบย้อนกลับความต้องการ")
    add_document_control(doc, config, doc_id)
    add_version_history(doc, config)
    doc.add_page_break()

    # ── Introduction ──────────────────────────────────────────────────────────
    add_section_heading(doc, "1. วัตถุประสงค์ / Purpose", level=1)
    add_paragraph(doc,
        "RTM ใช้ติดตามและตรวจสอบว่าทุก Requirement ได้รับการออกแบบ พัฒนา และทดสอบครบถ้วน "
        "สามารถ trace กลับจาก Test Case หรือ Defect ไปยัง Requirement ต้นทางได้เสมอ")

    # ── 2. Traceability Matrix ────────────────────────────────────────────────
    add_section_heading(doc, "2. Traceability Matrix", level=1)
    add_paragraph(doc, "ตารางแสดงความสัมพันธ์: Requirement → Design Component → Test Cases → Defects")

    # Build lookup maps
    comp_map = {}
    for c in comps:
        for rid in c.get("related_requirements", []):
            comp_map.setdefault(rid, []).append(c.get("id",""))

    tc_map = {}
    for tc in tcs:
        rid = tc.get("related_requirement","")
        tc_map.setdefault(rid, []).append(tc.get("id",""))

    defect_map = {}
    for d in defects:
        rid = d.get("related_requirement","")
        defect_map.setdefault(rid, []).append(d.get("id",""))

    rtm_rows = []
    if reqs:
        for req in reqs:
            rid   = req.get("id","")
            title = req.get("title","")
            prio  = req.get("priority","Medium")
            comps_linked = ", ".join(comp_map.get(rid, ["—"]))
            tcs_linked   = ", ".join(tc_map.get(rid, ["—"]))
            defs_linked  = ", ".join(defect_map.get(rid, ["—"]))
            rtm_rows.append([rid, title, prio, comps_linked, tcs_linked, defs_linked, "✓"])
    else:
        rtm_rows = [["REQ-001","(Placeholder)","High","COMP-001","TC-001","—","✓"]]

    add_table(doc,
        headers=["REQ ID","Title","Priority","Design Comp.","Test Cases","Defects","Covered"],
        rows=rtm_rows,
        col_widths=[2.2, 5, 2.2, 3, 3, 2.5, 1.8]
    )

    # ── 3. Coverage Summary ───────────────────────────────────────────────────
    add_section_heading(doc, "3. Coverage Summary", level=1)
    total_req      = len(reqs) if reqs else 1
    covered_design = len([r for r in rtm_rows if r[3] != "—"])
    covered_test   = len([r for r in rtm_rows if r[4] != "—"])
    summary_rows = [
        ["Total Requirements",         str(total_req)],
        ["Requirements with Design",   f"{covered_design} ({covered_design*100//total_req}%)"],
        ["Requirements with Test Cases",f"{covered_test} ({covered_test*100//total_req}%)"],
        ["Open Defects",               str(len([d for d in defects if d.get("status","") in ["Open","In Progress"]]))],
    ]
    add_table(doc, ["Metric","Value"], summary_rows, col_widths=[8, 4])

    # ── 4. Legends ────────────────────────────────────────────────────────────
    add_section_heading(doc, "4. คำอธิบายสัญลักษณ์ / Legends", level=1)
    add_table(doc, ["สัญลักษณ์","ความหมาย"],
        [["✓","Covered/Traced",""],["—","Not yet assigned"],["REQ-xxx","Requirement ID"],
         ["COMP-xxx","Design Component ID"],["TC-xxx","Test Case ID"],["BUG-xxx","Defect ID"]],
        col_widths=[3,8]
    )

    # Sign-off
    doc.add_page_break()
    add_signature_table(doc, [
        {"role": "เตรียมโดย / Prepared by",     "name": t.get("business_analyst",{}).get("name",""), "title": "Business Analyst"},
        {"role": "ตรวจสอบโดย / Reviewed by",   "name": t.get("qa_engineer",{}).get("name",""),       "title": "QA Engineer"},
        {"role": "อนุมัติโดย / Approved by",    "name": t.get("project_manager",{}).get("name",""),  "title": "Project Manager"},
    ])

    os.makedirs(output_path, exist_ok=True)
    file_path = os.path.join(output_path, f"{doc_id}_RTM.docx")
    doc.save(file_path)
    return file_path
