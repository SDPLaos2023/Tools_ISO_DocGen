"""
Template: Coding Standards
Folder 04 - Development
Document ID: [CODE]-04-CS-v[VER]
"""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from utils.doc_builder import (
    new_document, add_cover_page, add_document_control, add_version_history,
    add_section_heading, add_paragraph, add_table, add_signature_table, get_doc_id,
)

def generate(config: dict, output_path: str) -> str:
    doc_id = get_doc_id(config, "04", "CS")
    doc    = new_document()
    p      = config["project"]
    t      = config["team"]
    ts     = config.get("tech_stack", {})

    add_cover_page(doc, config, doc_id, "Coding Standards", "มาตรฐานการเขียนโปรแกรม")
    add_document_control(doc, config, doc_id)
    add_version_history(doc, config)
    doc.add_page_break()

    add_section_heading(doc, "1. วัตถุประสงค์ / Purpose", level=1)
    add_paragraph(doc, f"เอกสารนี้กำหนดมาตรฐานการเขียนโค้ดสำหรับโครงการ {p.get('name','')} "
                      "เพื่อให้โค้ดมีคุณภาพ อ่านเข้าใจง่าย และบำรุงรักษาได้ง่าย "
                      "นักพัฒนาทุกคนต้องปฏิบัติตามแนวทางนี้")

    add_section_heading(doc, "2. General Principles", level=1)
    principles = [
        "Write clean, readable, and self-documenting code",
        "DRY (Don't Repeat Yourself) — หลีกเลี่ยงการเขียนโค้ดซ้ำ",
        "SOLID Principles — Single Responsibility, Open/Closed, Liskov, Interface Segregation, Dependency Inversion",
        "KISS (Keep It Simple, Stupid) — เขียนให้เรียบง่ายที่สุดเท่าที่ทำได้",
        "Code must be peer-reviewed ก่อน merge into main branch",
    ]
    for pr in principles:
        add_paragraph(doc, f"• {pr}")

    add_section_heading(doc, "3. Naming Conventions", level=1)
    add_paragraph(doc, f"Stack หลัก: {ts.get('backend','—')} / {ts.get('frontend','—')}")
    naming_rows = [
        ["Variables",      "camelCase",    "e.g. userName, totalAmount"],
        ["Functions/Methods","camelCase",  "e.g. getUserById(), calculateTotal()"],
        ["Classes",        "PascalCase",   "e.g. UserService, PaymentController"],
        ["Constants",      "UPPER_SNAKE",  "e.g. MAX_RETRY_COUNT, API_BASE_URL"],
        ["Database Tables","snake_case",   "e.g. user_accounts, order_items"],
        ["Files",          "kebab-case",   "e.g. user-service.ts, payment-controller.py"],
        ["API Endpoints",  "kebab-case",   "e.g. /api/v1/user-accounts"],
    ]
    add_table(doc, ["Element","Convention","Example"], naming_rows, col_widths=[4,4,10])

    add_section_heading(doc, "4. Code Structure & Formatting", level=1)
    add_table(doc, ["Rule","Detail"], [
        ["Indentation",         "4 spaces (ห้ามใช้ tab)"],
        ["Line Length",         "Max 120 characters per line"],
        ["Blank Lines",         "2 blank lines ระหว่าง top-level definitions"],
        ["Imports",             "Group: stdlib → third-party → local; alphabetical sort"],
        ["File Encoding",       "UTF-8"],
        ["Line Endings",        "LF (Unix-style)"],
        ["Trailing Whitespace", "ห้ามมี trailing whitespace"],
    ], col_widths=[5,12])

    add_section_heading(doc, "5. Comments & Documentation", level=1)
    add_paragraph(doc, "5.1 Function/Method Comments (JSDoc / Docstring)")
    add_paragraph(doc, """All public functions ต้องมี docstring บอก:
  - @param / Args — parameters และ types
  - @returns / Returns — return value
  - @throws / Raises — exceptions ที่อาจเกิด""")
    add_paragraph(doc, "5.2 Inline Comments")
    add_paragraph(doc, "• ใช้ inline comment เฉพาะตรรกะที่ซับซ้อนหรือ workaround จำเป็น")
    add_paragraph(doc, "• ห้าม comment out โค้ดเก่าทิ้งไว้ — ใช้ Git สำหรับ history")

    add_section_heading(doc, "6. Error Handling", level=1)
    add_table(doc, ["Rule","Detail"], [
        ["Never swallow exceptions",  "ห้าม catch exception แล้วไม่ทำอะไร (empty catch)"],
        ["Log errors properly",       "Log ด้วย structured logging พร้อม stack trace"],
        ["Use custom exceptions",     "สร้าง custom exception class ตาม business domain"],
        ["Return meaningful errors",  "API error response ต้องมี code, message, details"],
    ], col_widths=[6,11])

    add_section_heading(doc, "7. Security Coding Rules", level=1)
    add_table(doc, ["Rule","Detail"], [
        ["No hardcoded credentials","ห้าม hardcode password, API key, token ในโค้ด"],
        ["Input Validation",        "Validate ทุก input ฝั่ง server เสมอ"],
        ["SQL Injection Prevention","ใช้ parameterized queries / ORM เท่านั้น"],
        ["XSS Prevention",          "Escape HTML output, ใช้ CSP header"],
        ["Secrets Management",      "ใช้ environment variables หรือ secrets manager"],
    ], col_widths=[6,11])

    add_section_heading(doc, "8. Version Control Standards", level=1)
    add_paragraph(doc, f"Source Control: {ts.get('source_control','Git')}")
    add_table(doc, ["Rule","Detail"], [
        ["Branch Naming",  "feature/[JIRA-ID]-description, bugfix/[BUG-ID]-description"],
        ["Commit Message", "[TYPE]([scope]): short description — e.g. feat(auth): add JWT login"],
        ["PR/MR Size",     "ไม่เกิน 400 lines changed ต่อ Pull Request"],
        ["Code Review",    "ต้องผ่าน peer review จาก developer อย่างน้อย 1 คนก่อน merge"],
        ["Main Branch",    "ห้าม commit ตรงลง main/master — ต้องผ่าน PR เท่านั้น"],
    ], col_widths=[4,13])

    add_section_heading(doc, "9. Tools & Linting", level=1)
    add_table(doc, ["Tool","Purpose"], [
        ["ESLint / Pylint",     "Code linting — ตรวจ code style อัตโนมัติ"],
        ["Prettier / Black",    "Code formatting — format โค้ดอัตโนมัติ"],
        ["SonarQube / SonarCloud","Code quality & security scanning"],
        ["Git Hooks (pre-commit)","รัน lint + tests อัตโนมัติก่อน commit"],
    ], col_widths=[6,11])

    doc.add_page_break()
    add_signature_table(doc, [
        {"role": "เตรียมโดย",  "name": t.get("lead_developer",{}).get("name",""),  "title": "Lead Developer"},
        {"role": "ตรวจสอบโดย","name": t.get("qa_engineer",{}).get("name",""),      "title": "QA Engineer"},
        {"role": "อนุมัติโดย", "name": t.get("project_manager",{}).get("name",""), "title": "Project Manager"},
    ])

    os.makedirs(output_path, exist_ok=True)
    file_path = os.path.join(output_path, f"{doc_id}_Coding_Standards.docx")
    doc.save(file_path)
    return file_path
