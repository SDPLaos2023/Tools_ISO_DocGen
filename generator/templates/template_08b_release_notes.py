"""
Template: Version Release Notes
Folder 08 - Change Logs & Versioning
Document ID: [CODE]-08-VRN-v[VER]
"""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from utils.doc_builder import (
    new_document, add_cover_page, add_document_control, add_version_history,
    add_section_heading, add_paragraph, add_table, add_signature_table, get_doc_id,
)

def generate(config: dict, output_path: str) -> str:
    doc_id   = get_doc_id(config, "08", "VRN")
    doc      = new_document()
    t        = config["team"]
    p        = config["project"]
    versions = config.get("versions", [])
    deps     = config.get("deployments", [])

    add_cover_page(doc, config, doc_id, "Version Release Notes", "บันทึก Version และการ Release")
    add_document_control(doc, config, doc_id)
    add_version_history(doc, config)
    doc.add_page_break()

    add_section_heading(doc, "1. Version History Summary", level=1)
    if versions:
        ver_rows = [[v.get("version",""), v.get("release_date",""), v.get("release_type",""),
                     v.get("description","")]
                    for v in versions]
    else:
        ver_rows = [["1.0.0", p.get("go_live_date",""), "Major", "Initial release"]]
    add_table(doc, ["Version","Release Date","Type","Description"], ver_rows, col_widths=[3,4,3,8])

    if not versions:
        versions = [{"version":"1.0.0","release_date":p.get("go_live_date",""),
                     "release_type":"Major","description":"Initial release",
                     "changes":["Initial deployment"],"deployed_by":"","environment":"Production"}]

    for v in versions:
        add_section_heading(doc, f"Release v{v.get('version','')} — {v.get('release_date','')}", level=1)
        add_table(doc, ["หัวข้อ","รายละเอียด"], [
            ["Version",          v.get("version","")],
            ["Release Date",     v.get("release_date","")],
            ["Release Type",     v.get("release_type","")],
            ["Environment",      v.get("environment","Production")],
            ["Deployed by",      v.get("deployed_by","")],
            ["Description",      v.get("description","")],
        ], col_widths=[5,12])

        add_section_heading(doc, "Changes in this Release", level=2)
        for change in v.get("changes", []):
            add_paragraph(doc, f"• {change}")

        # Linked deployment
        dep_match = [d for d in deps if d.get("version","") == v.get("version","")]
        if dep_match:
            dep = dep_match[0]
            add_section_heading(doc, "Deployment Details", level=2)
            add_table(doc, ["หัวข้อ","รายละเอียด"], [
                ["Deployment ID",  dep.get("id","")],
                ["Date",           dep.get("date","")],
                ["Type",           dep.get("deployment_type","")],
                ["Status",         dep.get("status","")],
                ["Approved by",    dep.get("approval","")],
                ["Rollback Plan",  dep.get("rollback_plan","")],
            ], col_widths=[5,12])

        add_section_heading(doc, "Known Issues", level=2)
        defects = config.get("defects", [])
        ver_changes = v.get("changes", [])
        ver_bugs = [d for d in defects
                    if d.get("status","") in ("In Progress", "Open")
                    and any(d.get("id","") in chg for chg in ver_changes)]
        if ver_bugs:
            for d in ver_bugs:
                add_paragraph(doc, f"• {d.get('id','')} — {d.get('title','')} [{d.get('severity','')} severity] — {d.get('status','')}")
        else:
            add_paragraph(doc, "• ไม่มี known issues ที่ยังค้างอยู่ใน release นี้")

    doc.add_page_break()
    add_signature_table(doc, [
        {"role": "เตรียมโดย / Prepared by",  "name": t.get("lead_developer",{}).get("name",""),  "title": "Lead Developer"},
        {"role": "ตรวจสอบโดย / Reviewed",   "name": t.get("qa_engineer",{}).get("name",""),      "title": "QA Engineer"},
        {"role": "อนุมัติโดย / Approved by", "name": t.get("project_manager",{}).get("name",""),  "title": "Project Manager"},
    ])

    os.makedirs(output_path, exist_ok=True)
    file_path = os.path.join(output_path, f"{doc_id}_Version_Release_Notes.docx")
    doc.save(file_path)
    return file_path
