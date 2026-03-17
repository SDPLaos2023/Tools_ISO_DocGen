"""
Template: Training Record
Folder 06 - Deployment & Training
Document ID: [CODE]-06-TR-v[VER]
Requires sign-off from attendees (who attended training)
"""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from utils.doc_builder import (
    new_document, add_cover_page, add_document_control, add_version_history,
    add_section_heading, add_paragraph, add_table, add_signature_table, get_doc_id,
)

def generate(config: dict, output_path: str) -> str:
    doc_id   = get_doc_id(config, "06", "TR")
    doc      = new_document()
    t        = config["team"]
    sessions = config.get("training_sessions", [])
    p        = config["project"]

    add_cover_page(doc, config, doc_id, "Training Record", "บันทึกการฝึกอบรม")
    add_document_control(doc, config, doc_id)
    add_version_history(doc, config)
    doc.add_page_break()

    add_section_heading(doc, "1. วัตถุประสงค์ / Purpose", level=1)
    add_paragraph(doc, f"เอกสารนี้บันทึกการฝึกอบรมผู้ใช้งานระบบ {p.get('name','')} "
                      "ก่อน Go-Live เพื่อให้มั่นใจว่าผู้ใช้งานมีความรู้ความสามารถในการใช้ระบบ")

    if not sessions:
        sessions = [{
            "id": "TRN-001",
            "title": f"การฝึกอบรมการใช้งานระบบ {p.get('name','')}",
            "date": p.get("go_live_date",""),
            "duration": "4 hours",
            "location": "ห้องฝึกอบรม / Online",
            "trainer": t.get("lead_developer",{}).get("name",""),
            "topics": [
                "แนะนำระบบและวัตถุประสงค์",
                "การ Login และการตั้งค่าส่วนตัว",
                "ฟังก์ชันหลักของระบบ",
                "การจัดการข้อมูล",
                "ถาม-ตอบ และสาธิตการใช้งาน",
            ],
            "attendees": [],
        }]

    for session in sessions:
        add_section_heading(doc, f"Training Session: {session.get('title','')} ({session.get('id','')})", level=1)
        add_table(doc, ["หัวข้อ","รายละเอียด"], [
            ["Training ID",        session.get("id","")],
            ["วันที่ / Date",      session.get("date","")],
            ["ระยะเวลา / Duration",session.get("duration","")],
            ["สถานที่ / Location", session.get("location","")],
            ["วิทยากร / Trainer",  session.get("trainer","")],
        ], col_widths=[5,12])

        add_section_heading(doc, "หัวข้อที่อบรม / Topics Covered", level=2)
        for i, topic in enumerate(session.get("topics", [])):
            add_paragraph(doc, f"{i+1}. {topic}")

        add_section_heading(doc, "รายชื่อผู้เข้าอบรม / Attendees", level=2)
        attendees = session.get("attendees", [])
        if attendees:
            att_rows = [[i+1, a.get("name",""), a.get("department",""),
                         "✓" if a.get("signed") else "☐"]
                        for i, a in enumerate(attendees)]
            add_table(doc,
                headers=["#","ชื่อ / Name","หน่วยงาน / Dept","ลายมือชื่อ / Signature"],
                rows=att_rows,
                col_widths=[1.5,6,5,5]
            )
        else:
            # Empty attendance sheet for manual filling
            empty_rows = [[str(i+1),"","",""] for i in range(20)]
            add_table(doc,
                headers=["#","ชื่อ / Name","หน่วยงาน / Dept","ลายมือชื่อ / Signature"],
                rows=empty_rows,
                col_widths=[1.5,6,5,5]
            )

        add_section_heading(doc, "ผลการประเมิน / Training Evaluation", level=2)
        add_table(doc, ["หัวข้อประเมิน","ดีมาก (5)","ดี (4)","ปานกลาง (3)","ปรับปรุง (2)","ต้องปรับปรุง (1)"], [
            ["เนื้อหาการอบรม","","","","",""],
            ["วิทยากร","","","","",""],
            ["สื่อและอุปกรณ์","","","","",""],
            ["ระยะเวลา","","","","",""],
        ], col_widths=[5,2,2,2,2,2])
        doc.add_paragraph()

    # Sign-off
    doc.add_page_break()
    add_signature_table(doc, [
        {"role": "วิทยากร / Trainer",         "name": t.get("lead_developer",{}).get("name",""),  "title": "Lead Developer"},
        {"role": "ผู้จัดอบรม / Organizer",    "name": t.get("project_manager",{}).get("name",""), "title": "Project Manager"},
        {"role": "รับทราบ / Acknowledged",    "name": "", "title": "Department Head"},
    ])

    os.makedirs(output_path, exist_ok=True)
    file_path = os.path.join(output_path, f"{doc_id}_Training_Record.docx")
    doc.save(file_path)
    return file_path
