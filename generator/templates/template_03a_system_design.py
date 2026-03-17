"""
Template: System Design Document (SDD)
Folder 03 - Design Architecture
Document ID: [CODE]-03-SDD-v[VER]
"""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from utils.doc_builder import (
    new_document, add_cover_page, add_document_control, add_version_history,
    add_section_heading, add_paragraph, add_table, add_signature_table, get_doc_id,
)

def generate(config: dict, output_path: str) -> str:
    doc_id = get_doc_id(config, "03", "SDD")
    doc    = new_document()
    p      = config["project"]
    t      = config["team"]
    comps  = config.get("design_components", [])
    ts     = config.get("tech_stack", {})

    add_cover_page(doc, config, doc_id, "System Design Document", "เอกสารการออกแบบระบบ")
    add_document_control(doc, config, doc_id)
    add_version_history(doc, config)
    doc.add_page_break()

    add_section_heading(doc, "1. บทนำ / Introduction", level=1)
    add_paragraph(doc, f"เอกสารนี้อธิบายสถาปัตยกรรมและการออกแบบระบบ {p.get('name','')} "
                      "ครอบคลุม architecture overview, component design และ interface design "
                      "อ้างอิงตาม Requirements Document (02-BRD)")

    add_section_heading(doc, "2. System Architecture Overview", level=1)
    add_paragraph(doc, "2.1 Architecture Pattern", bold=True)
    add_paragraph(doc, "ระบบออกแบบตาม C# .NET MVC 3-Tier Architecture — แบ่งเป็น Presentation Layer (Nuxt.js / Razor Views), Business Logic Layer (C# .NET API + Services), และ Data Layer (Microsoft SQL Server + Dapper ORM)")
    layers = [
        ["Presentation Layer", ts.get("frontend","—"), "User Interface, Web/Mobile App"],
        ["Business Logic Layer", ts.get("backend","—"), "API, Business Rules, Services"],
        ["Data Layer", ts.get("database","—"), "Database, Data Access, Caching"],
        ["Infrastructure", ts.get("infrastructure","—"), "Hosting, CI/CD, Monitoring"],
    ]
    add_table(doc, ["Layer","Technology","Description"], layers, col_widths=[4,5,9])

    add_section_heading(doc, "3. Component Design", level=1)
    add_paragraph(doc, "Components ที่อ้างอิงใน RTM (เอกสาร 02-RTM):")
    if comps:
        comp_rows = [[c.get("id",""), c.get("name",""), c.get("type",""), c.get("description",""),
                      ", ".join(c.get("related_requirements",[]))]
                     for c in comps]
        add_table(doc, ["COMP ID","ชื่อ","ประเภท","คำอธิบาย","REQ Reference"],
                  comp_rows, col_widths=[2.5,4,3,5,3])
    else:
        add_table(doc, ["COMP ID","ชื่อ","ประเภท","คำอธิบาย","REQ Reference"],
                  [["COMP-001","(Placeholder)","Module","—","REQ-001"]], col_widths=[2.5,4,3,5,3])

    add_section_heading(doc, "4. Interface Design", level=1)
    add_section_heading(doc, "4.1 API Endpoints (สรุป)", level=2)
    add_table(doc, ["Method","Endpoint","Description","Auth Required"],
        [["GET","/api/v1/[resource]","Retrieve resource list","Yes"],
         ["POST","/api/v1/[resource]","Create new resource","Yes"],
         ["PUT","/api/v1/[resource]/{id}","Update resource","Yes"],
         ["DELETE","/api/v1/[resource]/{id}","Delete resource","Yes (Admin)"]],
        col_widths=[2,5,7,3])

    add_section_heading(doc, "4.2 External System Interfaces", level=2)
    add_paragraph(doc, "ระบุ interface กับระบบภายนอก (ถ้ามี):")
    add_paragraph(doc, "• Mobile App (Flutter) ↔ REST API (C# .NET) — Blacklist Scanner ผ่าน HTTPS/JSON")
    add_paragraph(doc, "• Web Client (Nuxt.js) ↔ REST API (C# .NET) — Dashboard, Reports ผ่าน Ajax/JSON")
    add_paragraph(doc, "• .NET MVC Razor Views ↔ Business Layer Services — Internal method calls")
    add_paragraph(doc, "• Business Layer ↔ SQL Server via Dapper ORM — Parameterized queries")
    add_paragraph(doc, "• CI/CD: GitHub Actions → Deploy to On-premise Windows/Linux Server via Script")

    add_section_heading(doc, "5. Security Design", level=1)
    sec_rows = [
        ["Authentication","JWT Token / OAuth 2.0","ทุก API endpoint"],
        ["Authorization","Role-Based Access Control (RBAC)","ตาม user role"],
        ["Encryption","TLS 1.3 (in transit), AES-256 (at rest)","ข้อมูลสำคัญ"],
        ["Input Validation","Server-side validation","ทุก input form"],
        ["Audit Log","Log ทุก action สำคัญ","User actions, Admin actions"],
    ]
    add_table(doc, ["Security Control","Mechanism","Applies To"], sec_rows, col_widths=[4,7,6])

    add_section_heading(doc, "6. Performance Design", level=1)
    perf_rows = [
        ["Response Time","< 3 seconds for 95% requests"],
        ["Concurrent Users","≥ 100 concurrent users"],
        ["Data Caching","Redis / In-memory cache for frequent queries"],
        ["Database Indexing","Index บน columns ที่ใช้ใน WHERE/JOIN"],
    ]
    add_table(doc, ["หัวข้อ","รายละเอียด"], perf_rows, col_widths=[5,12])

    add_section_heading(doc, "7. Deployment Architecture", level=1)
    add_paragraph(doc, f"Infrastructure: {ts.get('infrastructure','—')}")
    add_paragraph(doc, f"Source Control: {ts.get('source_control','—')}")
    add_paragraph(doc, f"CI/CD Pipeline: {ts.get('ci_cd','—')}")

    doc.add_page_break()
    add_signature_table(doc, [
        {"role": "เตรียมโดย / Prepared by", "name": t.get("system_analyst",{}).get("name",""), "title": "System Analyst"},
        {"role": "ตรวจสอบโดย / Reviewed by","name": t.get("lead_developer",{}).get("name",""), "title": "Lead Developer"},
        {"role": "อนุมัติโดย / Approved by", "name": t.get("project_manager",{}).get("name",""),"title": "Project Manager"},
    ])

    os.makedirs(output_path, exist_ok=True)
    file_path = os.path.join(output_path, f"{doc_id}_System_Design.docx")
    doc.save(file_path)
    return file_path
