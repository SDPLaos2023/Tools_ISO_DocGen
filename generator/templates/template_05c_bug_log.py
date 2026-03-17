"""
Template: Bug / Defect Log
Folder 05 - Testing QA
Document ID: [CODE]-05-BDL-v[VER]
Cross-references: Test Cases (05-TCR), Requirements (02-BRD)
"""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from utils.doc_builder import (
    new_document, add_cover_page, add_document_control, add_version_history,
    add_section_heading, add_paragraph, add_table, add_signature_table, get_doc_id,
)

def generate(config: dict, output_path: str) -> str:
    doc_id  = get_doc_id(config, "05", "BDL")
    doc     = new_document()
    t       = config["team"]
    defects = config.get("defects", [])

    add_cover_page(doc, config, doc_id, "Bug / Defect Log", "บันทึกข้อผิดพลาด")
    add_document_control(doc, config, doc_id)
    add_version_history(doc, config)
    doc.add_page_break()

    add_section_heading(doc, "1. Defect Summary", level=1)
    total      = len(defects) if defects else 0
    critical   = len([d for d in defects if d.get("severity","") == "Critical"])
    high       = len([d for d in defects if d.get("severity","") == "High"])
    medium     = len([d for d in defects if d.get("severity","") == "Medium"])
    low        = len([d for d in defects if d.get("severity","") == "Low"])
    open_count = len([d for d in defects if d.get("status","") in ["Open","In Progress"]])
    closed_cnt = len([d for d in defects if d.get("status","") == "Closed"])
    add_table(doc, ["Metric","Count"], [
        ["Total Defects",   str(total)],
        ["Critical",        str(critical)],
        ["High",            str(high)],
        ["Medium",          str(medium)],
        ["Low",             str(low)],
        ["Open",            str(open_count)],
        ["Closed",          str(closed_cnt)],
    ], col_widths=[8,4])

    add_section_heading(doc, "2. Defect List", level=1)
    if defects:
        defect_rows = [[
            d.get("id",""), d.get("title",""), d.get("severity",""),
            d.get("priority",""), d.get("status",""), d.get("related_requirement",""),
            d.get("related_test_case",""), d.get("reported_by",""), d.get("reported_date",""),
            d.get("assigned_to",""), d.get("fixed_date",""),
        ] for d in defects]
        add_table(doc,
            headers=["ID","Title","Severity","Priority","Status","REQ","TC","Reported By","Date","Assigned","Fixed Date"],
            rows=defect_rows,
            col_widths=[2,4,2.5,2.5,3,2,2,3,3,3,3]
        )
    else:
        add_paragraph(doc, "(ยังไม่มี defects บันทึกไว้)")
        add_table(doc,
            headers=["BUG ID","Title","Severity","Status","REQ Ref","TC Ref"],
            rows=[["BUG-001","(Placeholder)","Medium","Open","REQ-001","TC-001"]],
            col_widths=[3,5,3,3,3,3]
        )

    add_section_heading(doc, "3. Defect Detail", level=1)
    if defects:
        for d in defects:
            add_section_heading(doc, f"{d.get('id','')} — {d.get('title','')}", level=2)
            add_table(doc, ["หัวข้อ","รายละเอียด"], [
                ["Description",       d.get("description","")],
                ["Steps to Reproduce",d.get("steps_to_reproduce","")],
                ["Root Cause",        d.get("root_cause","")],
                ["Resolution",        d.get("resolution","")],
            ], col_widths=[4,13])
    else:
        add_paragraph(doc, "(จะแสดงรายละเอียด defect แต่ละรายการที่นี่)")

    add_section_heading(doc, "4. Defect Trend", level=1)
    add_paragraph(doc, "กราฟแนวโน้ม defects (เพิ่มรูป Defect Trend Chart ที่นี่):")
    add_paragraph(doc, "[แนบ chart / หรือดูจาก Jira/Azure DevOps Dashboard]")

    doc.add_page_break()
    add_signature_table(doc, [
        {"role": "รายงานโดย / Reported by","name": t.get("qa_engineer",{}).get("name",""),     "title": "QA Engineer"},
        {"role": "ตรวจสอบโดย / Reviewed","name": t.get("lead_developer",{}).get("name",""),    "title": "Lead Developer"},
        {"role": "รับทราบ / Acknowledged","name": t.get("project_manager",{}).get("name",""),  "title": "Project Manager"},
    ])

    os.makedirs(output_path, exist_ok=True)
    file_path = os.path.join(output_path, f"{doc_id}_Bug_Defect_Log.docx")
    doc.save(file_path)
    return file_path
