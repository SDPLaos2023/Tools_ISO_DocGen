"""
Template: Test Cases + Test Results
Folder 05 - Testing QA
Document ID: [CODE]-05-TCR-v[VER]
Cross-references: RTM (02-RTM) requirement IDs
"""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from utils.doc_builder import (
    new_document, add_cover_page, add_document_control, add_version_history,
    add_section_heading, add_paragraph, add_table, add_signature_table, get_doc_id,
    COLOR_GREEN, COLOR_RED,
)

def generate(config: dict, output_path: str) -> str:
    doc_id = get_doc_id(config, "05", "TCR")
    doc    = new_document()
    t      = config["team"]
    tcs    = config.get("test_cases", [])

    add_cover_page(doc, config, doc_id, "Test Cases & Test Results", "กรณีทดสอบและผลการทดสอบ")
    add_document_control(doc, config, doc_id)
    add_version_history(doc, config)
    doc.add_page_break()

    add_section_heading(doc, "1. Test Cases Summary", level=1)
    total  = len(tcs) if tcs else 1
    passed = len([tc for tc in tcs if tc.get("status","") == "Pass"])
    failed = len([tc for tc in tcs if tc.get("status","") == "Fail"])
    pend   = len([tc for tc in tcs if tc.get("status","") not in ["Pass","Fail"]])
    add_table(doc, ["Metric","Count"], [
        ["Total Test Cases",   str(total)],
        ["Passed",             str(passed)],
        ["Failed",             str(failed)],
        ["Pending",            str(pend)],
        ["Pass Rate",          f"{passed*100//total}%" if total else "0%"],
    ], col_widths=[8,4])

    add_section_heading(doc, "2. Test Cases Detail", level=1)
    if tcs:
        for tc in tcs:
            add_section_heading(doc, f"{tc.get('id','')} — {tc.get('title','')}", level=2)
            steps_text = "\n".join(tc.get("steps",[])) if tc.get("steps") else "—"
            add_table(doc, ["หัวข้อ","รายละเอียด"], [
                ["Test Type",         tc.get("test_type","")],
                ["REQ Reference",     tc.get("related_requirement","")],
                ["Preconditions",     tc.get("preconditions","")],
                ["Test Steps",        steps_text],
                ["Expected Result",   tc.get("expected_result","")],
                ["Actual Result",     tc.get("actual_result","")],
                ["Status",            tc.get("status","Pending")],
                ["Tester",            tc.get("tester","")],
                ["Test Date",         tc.get("test_date","")],
                ["Remarks",           tc.get("remarks","")],
            ], col_widths=[4,13])
    else:
        add_paragraph(doc, "(ยังไม่มี test cases — เพิ่มใน config.test_cases)")
        add_table(doc, ["TC ID","Title","REQ Ref","Type","Status"], [
            ["TC-001","(Placeholder)","REQ-001","Functional","Pending"],
        ], col_widths=[2.5,5,2.5,3.5,3])

    add_section_heading(doc, "3. Test Execution Log", level=1)
    exec_rows = [[tc.get("id",""), tc.get("tester",""), tc.get("test_date",""), tc.get("status","")]
                 for tc in tcs] if tcs else [["TC-001","—","—","Pending"]]
    add_table(doc, ["TC ID","Tester","Date","Status"], exec_rows, col_widths=[3,5,4,5])

    doc.add_page_break()
    add_signature_table(doc, [
        {"role": "ทดสอบโดย / Tested by",    "name": t.get("qa_engineer",{}).get("name",""),     "title": "QA Engineer"},
        {"role": "ตรวจสอบโดย / Reviewed by","name": t.get("lead_developer",{}).get("name",""), "title": "Lead Developer"},
        {"role": "อนุมัติโดย / Approved by", "name": t.get("project_manager",{}).get("name",""),"title": "Project Manager"},
    ])

    os.makedirs(output_path, exist_ok=True)
    file_path = os.path.join(output_path, f"{doc_id}_Test_Cases_Results.docx")
    doc.save(file_path)
    return file_path
