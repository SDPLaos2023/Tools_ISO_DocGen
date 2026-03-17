# ISO Document Generator — ระบบสร้างเอกสาร ISO/IEC 29110

ระบบสร้างเอกสาร ISO Software ชุดสมบูรณ์ (20+ ไฟล์ `.docx`) อัตโนมัติจาก config ไฟล์เดียว  
รองรับมาตรฐาน **ISO/IEC 29110** สำหรับ Small Software Organization

---

## โครงสร้างโปรเจกต์

```
DocISOGen/
├── AGENTS.md                  ← คู่มือสำหรับ Copilot Agent (workflow สร้างเอกสาร)
├── Prompt.md                  ← Prompt สำหรับ ISO Audit Team Agent
├── generator/
│   ├── generate_iso_docs.py   ← Main entry point (CLI)
│   ├── config_template.json   ← Schema/ตัวอย่าง config สำหรับแต่ละโปรเจกต์
│   ├── requirements.txt       ← Python dependencies
│   ├── templates/             ← Template แต่ละเอกสาร (21 ไฟล์)
│   └── utils/
│       ├── doc_builder.py     ← Helper สร้าง .docx (cover page, table, header ฯลฯ)
│       └── doc_validator.py   ← QA validator ตรวจความครบถ้วนหลังสร้าง
├── tools/
│   ├── audit_snapshot.py      ← สแกนเอกสารที่มีอยู่ → สร้าง snapshot.json
│   └── doc_patcher.py         ← Patch เอกสาร .docx ตาม patches.json
├── Logo/
│   └── LogoCompany.png        ← โลโก้ที่ใช้ใน cover page
├── configs/
│   └── [ProjectCode]/
│       ├── [ProjectCode]_config.json      ← Config ของโปรเจกต์
│       ├── [ProjectCode]_snapshot.json    ← Audit snapshot
│       ├── [ProjectCode]_patches.json     ← Auto-fix patches
│       └── AUDIT_ACTION_PLAN.md           ← แผนแก้ไขจาก audit
└── projects/
    └── [ProjectCode]/
        ├── 01_Project_Management/
        ├── 02_Requirements_Analysis/
        ├── 03_Design_Architecture/
        ├── 04_Development/
        ├── 05_Testing_QA/
        ├── 06_Deployment_Training/
        ├── 07_Support_Maintenance/
        ├── 08_Change_Logs_Versioning/
        ├── 09_Risk_Management/
        └── 10_Regulatory_Compliance/
```

---

## เอกสารที่สร้างได้ (20 ไฟล์)

| โฟลเดอร์ | เอกสาร | Document ID |
|----------|--------|-------------|
| 01 Project Management | Project Plan, Meeting Minutes | `PP`, `MM` |
| 02 Requirements Analysis | Requirements Doc (BRD/FRS), RTM | `BRD`, `RTM` |
| 03 Design Architecture | System Design, Database Design (ERD) | `SDD`, `DBD` |
| 04 Development | Coding Standards, Code Review Records | `CS`, `CRR` |
| 05 Testing QA | Test Plan, Test Cases & Results, Bug Log | `TP`, `TCR`, `BDL` |
| 06 Deployment Training | User Manual, Training Record | `UM`, `TR` |
| 07 Support Maintenance | Incident / Support Log | `ISL` |
| 08 Change Logs | Change Request Form, Version Release Notes | `CRF`, `VRN` |
| 09 Risk Management | Risk Register | `RR` |
| 10 Regulatory Compliance | ISO Checklist, Audit Report, CAPA | `IC`, `AR`, `CAPA` |

> Document ID format: `[ProjectCode]-[FolderNum]-[DocCode]-v[Version]`  
> ตัวอย่าง: `HRMS-02-RTM-v1.0`

---

## การติดตั้ง

### ความต้องการ
- Python 3.10 ขึ้นไป
- pip packages: `python-docx`, `openpyxl`, `Pillow`

### ติดตั้ง dependencies

```powershell
# สร้าง virtual environment (แนะนำ)
python -m venv .venv
.\.venv\Scripts\Activate.ps1

# ติดตั้ง packages
pip install -r generator/requirements.txt
```

---

## วิธีใช้งาน

### 1. สร้าง Config ไฟล์ใหม่

คัดลอก `generator/config_template.json` แล้วกรอกข้อมูลโปรเจกต์:

```powershell
Copy-Item generator/config_template.json configs/MYPROJ/MYPROJ_config.json
# แก้ไข MYPROJ_config.json ด้วย text editor
```

**ข้อมูลที่ต้องกรอก:**
- `project.code` — รหัสโปรเจกต์ (ใช้ตั้งชื่อโฟลเดอร์)
- `project.name`, `project.organization`, `project.start_date`, `project.end_date`
- `team.project_manager`, `team.lead_developer`
- `tech_stack` — Frontend / Backend / Database
- `requirements` — รายการ requirements อย่างน้อย 3 รายการ

### 2. สร้างเอกสารทั้งหมด

```powershell
python generator/generate_iso_docs.py --config configs/MYPROJ/MYPROJ_config.json
```

### 3. สร้างเฉพาะโฟลเดอร์เดียว

```powershell
python generator/generate_iso_docs.py --config configs/MYPROJ/MYPROJ_config.json --folder 05
```

### 4. ตรวจ QA เอกสารที่มีอยู่

```powershell
python generator/generate_iso_docs.py --validate-only --config configs/MYPROJ/MYPROJ_config.json
```

### 5. ทดสอบด้วย Demo Config

```powershell
python generator/generate_iso_docs.py --demo
```

---

## Audit Workflow (ISO 29110 Audit Team)

สำหรับการตรวจสอบเอกสารที่สร้างแล้ว ใช้ workflow นี้:

### Step 1 — สร้าง Snapshot

```powershell
.\.venv\Scripts\python.exe tools/audit_snapshot.py --project MYPROJ --report
```

สร้างไฟล์ `configs/MYPROJ/MYPROJ_snapshot.json` ที่สรุปเนื้อหาทุกเอกสาร

### Step 2 — วิเคราะห์และสร้าง Patches

วิเคราะห์ snapshot แล้วสร้าง `configs/MYPROJ/MYPROJ_patches.json`  
สำหรับ findings ที่ระบบแก้ได้อัตโนมัติ

### Step 3 — Dry Run (ตรวจก่อน apply)

```powershell
.\.venv\Scripts\python.exe tools/doc_patcher.py --project MYPROJ --patches configs/MYPROJ/MYPROJ_patches.json --dry-run
```

### Step 4 — Apply Patches

```powershell
.\.venv\Scripts\python.exe tools/doc_patcher.py --project MYPROJ --patches configs/MYPROJ/MYPROJ_patches.json
```

---

## การใช้งานกับ GitHub Copilot Agent

โปรเจกต์นี้มี `AGENTS.md` และ `Prompt.md` สำหรับสั่ง Copilot สร้างเอกสารอัตโนมัติ:

- **`AGENTS.md`** — คำสั่งสำหรับ Copilot สร้างเอกสารใหม่ทั้งชุดจากข้อมูลโปรเจกต์
- **`Prompt.md`** — คำสั่งสำหรับ ISO Audit Team Agent ตรวจสอบเอกสารที่มีอยู่

---

## เทคโนโลยีที่ใช้

| เทคโนโลยี | การใช้งาน |
|-----------|-----------|
| Python 3.10+ | Language หลัก |
| python-docx | สร้างไฟล์ .docx |
| openpyxl | รองรับ Excel tables |
| Pillow | แทรกโลโก้ใน cover page |

---

## License

Internal use — SDP Laos 2023
