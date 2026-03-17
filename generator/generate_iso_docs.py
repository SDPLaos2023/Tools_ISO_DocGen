"""
generate_iso_docs.py — Main ISO Document Generator
=====================================================
Usage:
    python generate_iso_docs.py --config <path_to_config.json>
    python generate_iso_docs.py --config <path_to_config.json> --folder 05
    python generate_iso_docs.py --demo
    python generate_iso_docs.py --config <file> --no-validate   (skip QA)
    python generate_iso_docs.py --validate-only --config <file> (QA only)

Options:
    --config  PATH       Path to project config JSON file
    --folder  NUM        Generate only this folder (01-10). Default: all
    --demo               Run with built-in demo config (no config file needed)
    --list               List all documents that will be generated
    --no-validate        Skip the QA validation step after generation
    --validate-only      Run QA validator on existing docs (no generation)
    --verbose-validate   Show INFO-level issues in the QA report

ISO/IEC 29110 Document Set:
    01 Project Management      : Project Plan, Meeting Minutes
    02 Requirements Analysis   : Requirements Document (BRD/FRS), RTM
    03 Design Architecture     : System Design, Database Design
    04 Development             : Coding Standards, Code Review Records
    05 Testing QA              : Test Plan, Test Cases+Results, Bug Log
    06 Deployment Training     : User Manual, Training Record
    07 Support Maintenance     : Incident/Support Log
    08 Change Logs Versioning  : Change Request Form, Version Release Notes
    09 Risk Management         : Risk Register
    10 Regulatory Compliance   : ISO Checklist, Audit Report, CAPA
"""

import argparse
import json
import os
import sys
import time
from datetime import datetime
from pathlib import Path

# Ensure templates and utils are importable
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)

# ─── Document Registry ────────────────────────────────────────────────────────
# Each entry: (folder_key, subfolder, module_name, doc_label)
DOCUMENT_REGISTRY = [
    ("01", "01_Project_Management",      "templates.template_01a_project_plan",    "Project Plan"),
    ("01", "01_Project_Management",      "templates.template_01b_meeting_minutes",  "Meeting Minutes"),
    ("02", "02_Requirements_Analysis",   "templates.template_02a_requirements",     "Requirements Document (BRD/FRS)"),
    ("02", "02_Requirements_Analysis",   "templates.template_02b_rtm",              "Requirements Traceability Matrix"),
    ("03", "03_Design_Architecture",     "templates.template_03a_system_design",    "System Design Document"),
    ("03", "03_Design_Architecture",     "templates.template_03b_database_design",  "Database Design (ERD)"),
    ("04", "04_Development",             "templates.template_04a_coding_standards", "Coding Standards"),
    ("04", "04_Development",             "templates.template_04b_code_review",      "Code Review Records"),
    ("05", "05_Testing_QA",              "templates.template_05a_test_plan",        "Test Plan"),
    ("05", "05_Testing_QA",              "templates.template_05b_test_cases",       "Test Cases & Results"),
    ("05", "05_Testing_QA",              "templates.template_05c_bug_log",          "Bug/Defect Log"),
    ("06", "06_Deployment_Training",     "templates.template_06a_user_manual",      "User Manual"),
    ("06", "06_Deployment_Training",     "templates.template_06b_training_record",  "Training Record"),
    ("07", "07_Support_Maintenance",     "templates.template_07_incident_log",      "Incident/Support Log"),
    ("08", "08_Change_Logs_Versioning",  "templates.template_08a_change_request",   "Change Request Form"),
    ("08", "08_Change_Logs_Versioning",  "templates.template_08b_release_notes",    "Version Release Notes"),
    ("09", "09_Risk_Management",         "templates.template_09_risk_register",     "Risk Register"),
    ("10", "10_Regulatory_Compliance",   "templates.template_10a_iso_checklist",    "ISO Checklist"),
    ("10", "10_Regulatory_Compliance",   "templates.template_10b_audit_report",     "Audit Report"),
    ("10", "10_Regulatory_Compliance",   "templates.template_10c_capa",             "CAPA"),
]


# ─── Demo Config ──────────────────────────────────────────────────────────────
DEMO_CONFIG = {
    "project": {
        "code": "DEMO001",
        "name": "ระบบจัดการทรัพยากรบุคคล (HRM Demo)",
        "short_name": "HRMDemo",
        "version": "1.0",
        "description": "ระบบบริหารจัดการทรัพยากรบุคคลครบวงจร รองรับการจัดการพนักงาน เงินเดือน และการลา",
        "scope": "ครอบคลุม: ระบบจัดการพนักงาน, ระบบเงินเดือน, ระบบลางาน, รายงาน HR\nไม่ครอบคลุม: ระบบ ERP อื่น, Payroll banking integration",
        "objectives": [
            "ลดเวลาการทำงาน HR ด้วย automation ≥ 50%",
            "มี real-time dashboard สำหรับผู้บริหาร",
            "รองรับ Mobile access",
        ],
        "start_date": "2026-01-01",
        "end_date": "2026-12-31",
        "go_live_date": "2026-11-01",
        "organization": "บริษัท เดโม จำกัด",
        "department": "ฝ่ายเทคโนโลยีสารสนเทศ",
        "document_date": datetime.now().strftime("%Y-%m-%d"),
        "classification": "Internal",
    },
    "team": {
        "project_manager":  {"name": "สมชาย รักงาน",    "title": "Project Manager",    "email": "somchai@demo.com"},
        "lead_developer":   {"name": "วิชัย โค้ดเก่ง",   "title": "Lead Developer",     "email": "wichai@demo.com"},
        "business_analyst": {"name": "สมหญิง วิเคราะห์", "title": "Business Analyst",   "email": "somying@demo.com"},
        "system_analyst":   {"name": "อานนท์ ออกแบบดี",  "title": "System Analyst",     "email": "arnon@demo.com"},
        "qa_engineer":      {"name": "มาลี ทดสอบเก่ง",   "title": "QA Engineer",        "email": "malee@demo.com"},
        "dba":              {"name": "สุรินทร์ ฐานข้อมูล","title": "DBA",                "email": "surin@demo.com"},
        "members": [],
    },
    "stakeholders": [
        {"name": "ผู้อำนวยการ HR", "role": "Project Sponsor", "organization": "บริษัท เดโม จำกัด", "email": "", "responsibility": "อนุมัติงบประมาณ"},
    ],
    "tech_stack": {
        "frontend": "React 18, TypeScript, Ant Design",
        "backend": "Python FastAPI 0.110",
        "database": "PostgreSQL 15",
        "infrastructure": "Azure Cloud",
        "source_control": "Azure DevOps Git",
        "ci_cd": "Azure Pipelines",
        "other": [],
    },
    "milestones": [
        {"id": "MS-01", "name": "Requirements Sign-off",  "target_date": "2026-02-28", "status": "Completed", "owner": "BA"},
        {"id": "MS-02", "name": "Design Complete",         "target_date": "2026-03-31", "status": "Completed", "owner": "SA"},
        {"id": "MS-03", "name": "Development Complete",    "target_date": "2026-08-31", "status": "Planned",   "owner": "Dev Lead"},
        {"id": "MS-04", "name": "UAT Complete",            "target_date": "2026-10-15", "status": "Planned",   "owner": "QA"},
        {"id": "MS-05", "name": "Go-Live",                 "target_date": "2026-11-01", "status": "Planned",   "owner": "PM"},
    ],
    "requirements": [
        {"id": "REQ-001", "title": "จัดการข้อมูลพนักงาน",     "description": "ระบบต้องสามารถเพิ่ม แก้ไข ลบ และค้นหาข้อมูลพนักงานได้", "priority": "High",   "type": "Functional",     "category": "Employee Management", "source": "HR Director",    "acceptance_criteria": "CRUD operations สำเร็จ 100%", "linked_design": ["COMP-001"], "linked_test_cases": ["TC-001"]},
        {"id": "REQ-002", "title": "คำนวณเงินเดือน",           "description": "ระบบต้องคำนวณเงินเดือนพนักงานประจำเดือนได้ถูกต้อง รวม OT, ค่าล่วงเวลา, หัก ณ ที่จ่าย", "priority": "High",   "type": "Functional",     "category": "Payroll",   "source": "HR Director",    "acceptance_criteria": "ความถูกต้อง 100% เทียบ manual calculation", "linked_design": ["COMP-002"], "linked_test_cases": ["TC-002"]},
        {"id": "REQ-003", "title": "ระบบลางาน",               "description": "พนักงานขอลาผ่านระบบ ผู้จัดการอนุมัติ/ปฏิเสธ ระบบแจ้งเตือนอัตโนมัติ",                  "priority": "High",   "type": "Functional",     "category": "Leave",     "source": "HR Manager",     "acceptance_criteria": "Workflow ครบถ้วน ส่ง email notification",    "linked_design": ["COMP-003"], "linked_test_cases": ["TC-003"]},
        {"id": "REQ-004", "title": "Dashboard ผู้บริหาร",      "description": "แสดง KPI HR แบบ real-time: จำนวนพนักงาน, อัตราการลา, ค่าใช้จ่าย",                      "priority": "Medium", "type": "Functional",     "category": "Reporting", "source": "CEO",            "acceptance_criteria": "โหลดหน้า dashboard < 3 วินาที",              "linked_design": ["COMP-004"], "linked_test_cases": ["TC-004"]},
        {"id": "REQ-005", "title": "Performance ≥ 100 users",  "description": "ระบบต้องรองรับผู้ใช้งานพร้อมกัน ≥ 100 คน",                                              "priority": "High",   "type": "Non-Functional", "category": "Performance","source": "IT Director",    "acceptance_criteria": "Response time < 3s ที่ load 100 users",       "linked_design": [],           "linked_test_cases": ["TC-005"]},
    ],
    "design_components": [
        {"id": "COMP-001", "name": "Employee Module",    "description": "CRUD employee master data", "type": "API + UI",  "related_requirements": ["REQ-001"], "technology": "FastAPI + React"},
        {"id": "COMP-002", "name": "Payroll Module",     "description": "Salary calculation engine", "type": "Service",   "related_requirements": ["REQ-002"], "technology": "Python + PostgreSQL"},
        {"id": "COMP-003", "name": "Leave Module",       "description": "Leave request workflow",    "type": "API + UI",  "related_requirements": ["REQ-003"], "technology": "FastAPI + React + Email"},
        {"id": "COMP-004", "name": "Dashboard Module",   "description": "Real-time HR KPI charts",   "type": "UI + API",  "related_requirements": ["REQ-004"], "technology": "React + Chart.js"},
    ],
    "database_tables": [
        {"name": "employees",      "description": "ข้อมูลพนักงานหลัก", "columns": [{"name":"id","type":"SERIAL PRIMARY KEY","description":"PK"},{"name":"emp_code","type":"VARCHAR(20) UNIQUE","description":"รหัสพนักงาน"},{"name":"full_name","type":"VARCHAR(255)","description":"ชื่อ-นามสกุล"},{"name":"department_id","type":"INT FK","description":"FK departments"},{"name":"salary","type":"DECIMAL(12,2)","description":"เงินเดือนปัจจุบัน"},{"name":"hire_date","type":"DATE","description":"วันที่เริ่มงาน"},{"name":"is_active","type":"BOOLEAN DEFAULT TRUE","description":"สถานะ active"},{"name":"created_at","type":"TIMESTAMP","description":"วันที่สร้าง"},{"name":"updated_at","type":"TIMESTAMP","description":"วันที่แก้ไข"}], "related_requirements": ["REQ-001"]},
        {"name": "leave_requests", "description": "คำขอลางาน", "columns": [{"name":"id","type":"SERIAL PRIMARY KEY","description":"PK"},{"name":"employee_id","type":"INT FK","description":"FK employees"},{"name":"leave_type","type":"VARCHAR(50)","description":"ประเภทการลา"},{"name":"start_date","type":"DATE","description":"วันที่เริ่มลา"},{"name":"end_date","type":"DATE","description":"วันที่สิ้นสุด"},{"name":"status","type":"VARCHAR(20)","description":"Pending/Approved/Rejected"},{"name":"approved_by","type":"INT FK","description":"FK employees (approver)"}], "related_requirements": ["REQ-003"]},
    ],
    "test_cases": [
        {"id": "TC-001", "title": "เพิ่มพนักงานใหม่",       "related_requirement": "REQ-001", "test_type": "Functional", "preconditions": "Login เป็น Admin", "steps": ["ไปที่ Employee > Add New","กรอกข้อมูลครบถ้วน","คลิก Save"], "expected_result": "พนักงานถูกบันทึก แสดงใน list", "actual_result": "Pass", "status": "Pass", "tester": "มาลี ทดสอบเก่ง", "test_date": "2026-09-10", "remarks": ""},
        {"id": "TC-002", "title": "คำนวณเงินเดือนถูกต้อง",   "related_requirement": "REQ-002", "test_type": "Functional", "preconditions": "มีข้อมูลพนักงานและ OT hours", "steps": ["ไปที่ Payroll > Calculate","เลือกเดือน","คลิก Calculate All"], "expected_result": "ผลลัพธ์ตรงกับ manual calculation 100%", "actual_result": "Pass", "status": "Pass", "tester": "มาลี ทดสอบเก่ง", "test_date": "2026-09-11", "remarks": ""},
        {"id": "TC-003", "title": "ขอลาและผู้จัดการอนุมัติ",  "related_requirement": "REQ-003", "test_type": "Integration", "preconditions": "Login เป็น Employee", "steps": ["ไปที่ Leave > New Request","กรอกรายละเอียด","Submit"], "expected_result": "ผู้จัดการได้รับ email notification", "actual_result": "Pass", "status": "Pass", "tester": "มาลี ทดสอบเก่ง", "test_date": "2026-09-12", "remarks": ""},
        {"id": "TC-004", "title": "Dashboard โหลดเร็ว",     "related_requirement": "REQ-004", "test_type": "Performance", "preconditions": "มีข้อมูล 500 พนักงาน", "steps": ["Login","ไปที่ Dashboard"], "expected_result": "โหลดภายใน 3 วินาที", "actual_result": "2.1s", "status": "Pass", "tester": "มาลี ทดสอบเก่ง", "test_date": "2026-09-15", "remarks": ""},
        {"id": "TC-005", "title": "Load test 100 users",    "related_requirement": "REQ-005", "test_type": "Performance", "preconditions": "k6 load test script ready", "steps": ["รัน k6 --vus 100 --duration 60s"], "expected_result": "p95 response time < 3s, error rate < 1%", "actual_result": "p95=2.4s, error=0.1%", "status": "Pass", "tester": "มาลี ทดสอบเก่ง", "test_date": "2026-09-20", "remarks": ""},
    ],
    "defects": [
        {"id": "BUG-001", "title": "เงินเดือน OT คำนวณผิดกรณีค่าแรงพิเศษ", "related_test_case": "TC-002", "related_requirement": "REQ-002", "severity": "High", "priority": "High", "status": "Closed", "reported_by": "มาลี ทดสอบเก่ง", "reported_date": "2026-09-11", "assigned_to": "วิชัย โค้ดเก่ง", "fixed_date": "2026-09-13", "description": "OT rate คำนวณผิดเมื่อพนักงานมี allowance พิเศษ", "steps_to_reproduce": "1. สร้างพนักงานมี special allowance 2. คำนวณ OT", "root_cause": "Formula ไม่รวม allowance ใน base pay calculation", "resolution": "แก้ไข formula ใน payroll_calculator.py ใน commit abc123"},
    ],
    "risks": [
        {"id": "RISK-001", "category": "Resource",  "description": "Lead Developer อาจลาออกระหว่างโครงการ",        "probability": "Low",    "impact": "High",   "risk_level": "Medium", "mitigation": "สร้าง knowledge sharing และ documentation ครบถ้วน", "contingency": "มี backup developer พร้อม", "owner": "สมชาย รักงาน", "status": "Open", "review_date": "2026-03-01", "linked_capa": ""},
        {"id": "RISK-002", "category": "Technical",  "description": "Payroll module ซับซ้อนกว่าที่ประเมิน",          "probability": "Medium", "impact": "High",   "risk_level": "High",   "mitigation": "Spike ก่อน 2 สัปดาห์ ประเมิน effort จริง",         "contingency": "ขยายเวลา development 2 สัปดาห์", "owner": "วิชัย โค้ดเก่ง","status": "Mitigated","review_date": "2026-03-01","linked_capa": ""},
        {"id": "RISK-003", "category": "Schedule",   "description": "UAT ล่าช้าเพราะ user ไม่พร้อม",               "probability": "Medium", "impact": "Medium", "risk_level": "Medium", "mitigation": "นัด UAT date ล่วงหน้า 1 เดือน",                      "contingency": "เลื่อน Go-Live 2 สัปดาห์", "owner": "สมชาย รักงาน", "status": "Open", "review_date": "2026-09-01", "linked_capa": ""},
    ],
    "change_requests": [
        {"id": "CR-001", "title": "เพิ่มฟังก์ชัน Mobile App สำหรับ Leave Request", "description": "HR Director ขอให้เพิ่ม Mobile App รองรับ iOS/Android", "requestor": "HR Director", "request_date": "2026-04-15", "priority": "Medium", "impact": "เพิ่ม scope +4 สัปดาห์, งบเพิ่ม 200,000 บาท", "affected_documents": ["REQ-003","03-SDD"], "status": "Approved", "approved_by": "สมชาย รักงาน", "approval_date": "2026-04-20", "implementation_date": "2026-06-01"},
    ],
    "versions": [
        {"version": "1.0.0", "release_date": "2026-11-01", "release_type": "Major", "description": "Initial production release", "changes": ["Employee management module","Payroll calculation module","Leave management module","HR Dashboard","Mobile-responsive UI"], "deployed_by": "วิชัย โค้ดเก่ง", "environment": "Production"},
    ],
    "meetings": [
        {"id": "MTG-001", "title": "Kick-off Meeting", "date": "2026-01-05", "time": "09:00-10:30", "location": "ห้องประชุม A", "chair": "สมชาย รักงาน", "attendees": ["สมชาย รักงาน","วิชัย โค้ดเก่ง","สมหญิง วิเคราะห์","มาลี ทดสอบเก่ง","HR Director"], "agenda": ["แนะนำทีมงาน","ทบทวน Project Charter","ทบทวน Project Plan","กำหนดการประชุม status สัปดาห์ละครั้ง"], "action_items": [{"item": "BA จัดทำ BRD ฉบับร่าง","owner": "สมหญิง วิเคราะห์","due_date": "2026-01-20","status": "Completed"}], "summary": "ทีมงานรับทราบ scope และ timeline ของโครงการ ตกลง hold weekly status meeting ทุกวันจันทร์ 10:00"},
    ],
    "training_sessions": [
        {"id": "TRN-001", "title": "การฝึกอบรมการใช้งานระบบ HRM", "date": "2026-10-20", "duration": "4 hours", "location": "ห้องฝึกอบรม B", "trainer": "วิชัย โค้ดเก่ง", "topics": ["Overview ระบบ HRM","การจัดการข้อมูลพนักงาน","ระบบเงินเดือน","ระบบลางาน","Dashboard และ Reports","ถาม-ตอบ"], "attendees": [{"name": "พนักงาน HR 01","department": "ฝ่าย HR","signed": True},{"name": "พนักงาน HR 02","department": "ฝ่าย HR","signed": True}]},
    ],
    "incidents": [],
    "deployments": [
        {"id": "DEP-001", "version": "1.0.0", "date": "2026-11-01", "environment": "Production", "deployed_by": "วิชัย โค้ดเก่ง", "deployment_type": "Initial", "steps": ["รัน database migration","Deploy backend API","Deploy frontend","Smoke test","อนุมัติ Go-Live"], "rollback_plan": "Restore จาก backup + rollback database migration", "status": "Success", "approval": "สมชาย รักงาน"},
    ],
    "audit": {
        "audit_date": "2026-12-01",
        "auditor": "ผู้ตรวจสอบภายใน (Internal Auditor)",
        "audit_scope": "ISO/IEC 29110 compliance for HRM software development",
        "findings": [],
    },
    "capas": [],
    "output_path": "d:\\Work\\DocISOGen\\projects",
}


# ─── Core Generator ───────────────────────────────────────────────────────────
def load_config(config_path: str) -> dict:
    """Load and validate project config from JSON file."""
    with open(config_path, "r", encoding="utf-8") as f:
        config = json.load(f)

    # Basic validation
    required_keys = ["project", "team"]
    for key in required_keys:
        if key not in config:
            raise ValueError(f"Config missing required key: '{key}'")

    project_required = ["code", "name", "organization"]
    for key in project_required:
        if not config["project"].get(key):
            raise ValueError(f"config.project missing required field: '{key}'")

    return config


def get_output_root(config: dict, config_path: str = None) -> str:
    """Determine the output root folder for this project.

    Priority:
    1. If config_path is given and is inside a projects/<Name>/ folder,
       use that folder as the output root (config lives next to the docs).
    2. Otherwise use config["output_path"] / project_short_name.
    """
    if config_path:
        config_dir = os.path.dirname(os.path.abspath(config_path))
        project_name = os.path.basename(config_dir)
        parent_name = os.path.basename(os.path.dirname(config_dir))
        # Config lives next to the doc folders (legacy layout)
        if parent_name == "projects" or any(
            d.startswith(("01_", "02_", "03_")) for d in os.listdir(config_dir)
            if os.path.isdir(os.path.join(config_dir, d))
        ):
            return config_dir
        # Config lives in configs/<ProjectCode>/ → output to projects/<ProjectCode>/
        if parent_name == "configs":
            workspace_root = os.path.dirname(os.path.dirname(config_dir))
            return os.path.join(workspace_root, "projects", project_name)
    base = config.get("output_path", "d:\\Work\\DocISOGen\\projects")
    project_name = config["project"].get("short_name") or config["project"].get("name", "Project")
    # Sanitize folder name
    safe_name = "".join(c if c.isalnum() or c in " _-" else "_" for c in project_name).strip()
    return os.path.join(base, safe_name)


def generate_document(folder_key: str, subfolder: str, module_name: str,
                      doc_label: str, config: dict, output_root: str) -> dict:
    """Generate a single document. Returns result dict."""
    output_path = os.path.join(output_root, subfolder)
    start = time.time()
    try:
        import importlib
        module = importlib.import_module(module_name)
        file_path = module.generate(config, output_path)
        elapsed = time.time() - start
        return {"status": "OK", "label": doc_label, "path": file_path, "elapsed": elapsed}
    except Exception as e:
        elapsed = time.time() - start
        return {"status": "ERROR", "label": doc_label, "error": str(e), "elapsed": elapsed}


def run_generator(config: dict, folder_filter: str = None,
                  run_validate: bool = True, verbose_validate: bool = False,
                  config_path: str = None) -> list:
    """Generate all (or filtered) documents. Returns list of results."""
    output_root = get_output_root(config, config_path=config_path)
    project = config["project"]

    print("=" * 70)
    print(f"  ISO Document Generator — ISO/IEC 29110")
    print(f"  Project : {project.get('name','')}")
    print(f"  Code    : {project.get('code','')}")
    print(f"  Output  : {output_root}")
    print("=" * 70)

    results = []
    docs_to_run = [
        doc for doc in DOCUMENT_REGISTRY
        if folder_filter is None or doc[0] == folder_filter
    ]

    if not docs_to_run:
        print(f"No documents found for folder filter: {folder_filter}")
        return results

    total = len(docs_to_run)
    for i, (folder_key, subfolder, module_name, doc_label) in enumerate(docs_to_run, 1):
        print(f"  [{i:2d}/{total}] Generating: {doc_label}...", end=" ", flush=True)
        result = generate_document(folder_key, subfolder, module_name, doc_label, config, output_root)
        results.append(result)
        if result["status"] == "OK":
            print(f"✓  ({result['elapsed']:.1f}s)")
        else:
            print(f"✗  ERROR: {result['error']}")

    # Summary
    ok_count    = sum(1 for r in results if r["status"] == "OK")
    err_count   = sum(1 for r in results if r["status"] == "ERROR")
    total_time  = sum(r["elapsed"] for r in results)

    print("=" * 70)
    print(f"  Complete: {ok_count}/{total} documents generated  ({total_time:.1f}s)")
    if err_count:
        print(f"  Errors  : {err_count} document(s) failed — see above")
    print(f"  Output folder: {output_root}")
    print("=" * 70)

    if ok_count > 0:
        print("\n  Documents created:")
        for r in results:
            if r["status"] == "OK":
                rel = os.path.relpath(r["path"], output_root)
                print(f"    ✓ {rel}")
        print()

    # ── QA Validation ────────────────────────────────────────────────────────
    if run_validate and ok_count > 0:
        _run_validation(config, output_root, verbose=verbose_validate)

    return results


def _run_validation(config: dict, output_root: str, verbose: bool = False):
    """Run the QA validator and print the report."""
    try:
        from utils.doc_validator import validate_all, print_report
    except ImportError:
        print("  [QA] doc_validator not found — skipping validation")
        return

    print("  Running QA Validation...")
    print("  " + "-" * 68)
    qa_results = validate_all(config, output_root)
    all_approved = print_report(qa_results, verbose=verbose)

    if not all_approved:
        print("  ⚠  แก้ไข Critical Issues แล้วรันใหม่ หรือรัน --validate-only เพื่อตรวจซ้ำ\n")


def list_documents():
    """Print the full document list."""
    print("\n  ISO Document Set — ISO/IEC 29110")
    print("  " + "─" * 55)
    prev_folder = None
    for folder_key, subfolder, module_name, doc_label in DOCUMENT_REGISTRY:
        if folder_key != prev_folder:
            print(f"\n  Folder {folder_key} — {subfolder.replace('_', ' ')}")
            prev_folder = folder_key
        print(f"    • {doc_label}")
    print(f"\n  Total: {len(DOCUMENT_REGISTRY)} documents")


# ─── CLI Entry Point ──────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(
        description="ISO Document Generator — ISO/IEC 29110",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python generate_iso_docs.py --demo
  python generate_iso_docs.py --config d:\\Work\\DocISOGen\\MYPROJ_config.json
  python generate_iso_docs.py --config myproject.json --folder 05
  python generate_iso_docs.py --config myproject.json --no-validate
  python generate_iso_docs.py --config myproject.json --validate-only
  python generate_iso_docs.py --list
        """
    )
    parser.add_argument("--config",           help="Path to project config JSON file")
    parser.add_argument("--folder",           help="Generate only this folder (01-10)")
    parser.add_argument("--demo",             action="store_true", help="Run with built-in demo config")
    parser.add_argument("--list",             action="store_true", help="List all documents")
    parser.add_argument("--no-validate",      action="store_true", help="Skip QA validation after generation")
    parser.add_argument("--validate-only",    action="store_true", help="Run QA validator only (no generation)")
    parser.add_argument("--verbose-validate", action="store_true", help="Show INFO-level items in QA report")
    args = parser.parse_args()

    if args.list:
        list_documents()
        return

    if args.demo:
        config = DEMO_CONFIG
        print("\n  [DEMO MODE] Using built-in DEMO config (HRM Demo project)")
    elif args.config:
        if not os.path.exists(args.config):
            print(f"Error: Config file not found: {args.config}")
            sys.exit(1)
        config = load_config(args.config)
    else:
        parser.print_help()
        print("\nError: Provide --config <file> or --demo")
        sys.exit(1)

    if args.validate_only:
        # Run QA only on existing docs
        output_root = get_output_root(config, config_path=args.config if not args.demo else None)
        if not os.path.isdir(output_root):
            print(f"Error: Output folder not found: {output_root}")
            sys.exit(1)
        _run_validation(config, output_root, verbose=args.verbose_validate)
        return

    run_generator(
        config,
        folder_filter     = args.folder,
        run_validate      = not args.no_validate,
        verbose_validate  = args.verbose_validate,
        config_path       = args.config if not args.demo else None,
    )


if __name__ == "__main__":
    main()
