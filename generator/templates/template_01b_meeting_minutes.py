"""
Template: Meeting Minutes
Folder 01 - Project Management
Document ID pattern: [CODE]-01-MM-v[VER]
ISO/IEC 29110 Reference: PM Process - Project Meetings
"""
import os
import sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from utils.doc_builder import (
    new_document, add_cover_page, add_document_control,
    add_version_history, add_section_heading, add_paragraph,
    add_table, add_signature_table, get_doc_id,
    COLOR_PRIMARY, COLOR_SECONDARY, COLOR_LIGHT_GRAY
)


def generate(config: dict, output_path: str) -> str:
    doc_id = get_doc_id(config, "01", "MM")
    doc    = new_document()
    t      = config["team"]
    meetings = config.get("meetings", [])

    add_cover_page(doc, config, doc_id, "Meeting Minutes", "บันทึกการประชุม")
    add_document_control(doc, config, doc_id)
    add_version_history(doc, config)
    doc.add_page_break()

    if not meetings:
        meetings = [{
            "id": "MTG-001",
            "title": "Kick-off Meeting",
            "date": config["project"].get("start_date",""),
            "time": "09:00-10:00",
            "location": "ห้องประชุม",
            "chair": t.get("project_manager",{}).get("name",""),
            "attendees": [],
            "agenda": ["แนะนำทีมงาน", "ทบทวน Project Plan", "กำหนดการประชุมครั้งถัดไป"],
            "action_items": [],
            "summary": "ประชุมเปิดตัวโครงการ ทบทวนขอบเขตและแผนโครงการ",
        }]

    for mtg in meetings:
        # ── Header ────────────────────────────────────────────────────────────
        add_section_heading(doc, f"บันทึกการประชุม — {mtg.get('title','')} ({mtg.get('id','')})", level=1)

        add_table(doc,
            headers=["หัวข้อ","รายละเอียด"],
            rows=[
                ["วันที่ / Date",       mtg.get("date","")],
                ["เวลา / Time",         mtg.get("time","")],
                ["สถานที่ / Location",  mtg.get("location","")],
                ["ประธาน / Chair",      mtg.get("chair","")],
                ["Meeting ID",          mtg.get("id","")],
            ],
            col_widths=[6, 10]
        )

        # Attendees
        add_section_heading(doc, "ผู้เข้าร่วมประชุม / Attendees", level=2)
        attendees = mtg.get("attendees", [])
        if attendees:
            att_rows = [[i+1, name, "", ""] for i, name in enumerate(attendees)]
            add_table(doc,
                headers=["#","ชื่อ / Name","ตำแหน่ง / Title","ลายมือชื่อ / Signature"],
                rows=att_rows,
                col_widths=[1.5, 6, 5, 5]
            )
        else:
            add_paragraph(doc, "— (ไม่มีข้อมูลผู้เข้าร่วม / No attendees recorded)")

        # Agenda
        add_section_heading(doc, "วาระการประชุม / Agenda", level=2)
        for i, item in enumerate(mtg.get("agenda", [])):
            add_paragraph(doc, f"{i+1}. {item}")

        # Summary
        add_section_heading(doc, "สรุปการประชุม / Meeting Summary", level=2)
        add_paragraph(doc, mtg.get("summary","—"))

        # Action Items
        add_section_heading(doc, "การดำเนินงานต่อไป / Action Items", level=2)
        actions = mtg.get("action_items", [])
        if actions:
            act_rows = [[a.get("item",""), a.get("owner",""), a.get("due_date",""), a.get("status","Open")]
                        for a in actions]
            add_table(doc,
                headers=["รายการ / Action","ผู้รับผิดชอบ / Owner","กำหนด / Due","สถานะ / Status"],
                rows=act_rows,
                col_widths=[8, 4, 3, 3]
            )
        else:
            add_paragraph(doc, "— ไม่มี action items")

        # Next Meeting
        add_section_heading(doc, "การประชุมครั้งถัดไป / Next Meeting", level=2)
        add_paragraph(doc, "วันที่: _______________  เวลา: _______________  สถานที่: _______________")

        doc.add_paragraph()
        doc.add_paragraph("─" * 60)
        doc.add_paragraph()

    # Sign-off
    doc.add_page_break()
    add_signature_table(doc, [
        {"role": "ประธาน / Chair",                  "name": t.get("project_manager",{}).get("name",""), "title": "Project Manager"},
        {"role": "เลขานุการ / Secretary",           "name": t.get("business_analyst",{}).get("name",""), "title": "Business Analyst"},
        {"role": "รับทราบโดย / Acknowledged by",    "name": "", "title": "Stakeholder"},
    ])

    os.makedirs(output_path, exist_ok=True)
    file_path = os.path.join(output_path, f"{doc_id}_Meeting_Minutes.docx")
    doc.save(file_path)
    return file_path
