# ISO Document Generator — Copilot Instructions

## วัตถุประสงค์
ระบบนี้ใช้สร้างเอกสาร ISO Software ชุดสมบูรณ์ (20+ ไฟล์ .docx) จากรายละเอียด project ที่ user ส่งมา
โดย Copilot จะทำหน้าที่ extract ข้อมูล → populate config → สั่ง subagent สร้างเอกสารแต่ละกลุ่ม

---

## โครงสร้าง Workspace
```
d:\Work\DocISOGen\
├── generator\          ← engine สร้างเอกสาร (ไม่ต้องแตะ)
├── tools\              ← audit_snapshot.py, doc_patcher.py
├── Logo\               ← โลโก้องค์กร
├── AGENTS.md           ← คู่มือนี้
├── configs\            ← config, snapshot, patches, AUDIT_ACTION_PLAN ของแต่ละ project
│   ├── [ProjectCode]\
│   │   ├── [ProjectCode]_config.json
│   │   ├── [ProjectCode]_snapshot.json
│   │   ├── [ProjectCode]_patches.json
│   │   └── AUDIT_ACTION_PLAN.md
│   └── ...
└── projects\           ← เฉพาะ folder เอกสาร (01-10) — ไม่มีไฟล์อื่น
    ├── [ProjectCode]\
    │   ├── 01_Project_Management\
    │   └── ...
    └── ...
```

---

## ขั้นตอนการทำงานทุกครั้งที่ user ส่ง project มาใหม่

### Step 1: Extract Project Details
วิเคราะห์ข้อความที่ user ส่งมา แล้ว extract ออกมาเป็น JSON ตาม schema ใน `generator/config_template.json`

**ถ้าข้อมูลไม่ครบ ให้ถามก่อน ห้ามเดาในส่วนที่สำคัญเหล่านี้:**
- ชื่อ Project / Project Code (ใช้ตั้งชื่อโฟลเดอร์และ Document ID)
- ชื่อ Organization / Department
- Project Manager และ Lead Developer (ชื่อ-นามสกุล)
- ช่วงเวลา: วันเริ่ม / วันสิ้นสุด
- Requirements หลักอย่างน้อย 3 รายการ (ถ้าไม่มีให้สร้าง placeholder)
- Technology stack (Frontend, Backend, Database)

**ถ้าขาดข้อมูลรอง เช่น risks หรือ incidents → สร้าง placeholder ว่าง โดยไม่ต้องถาม**

### Step 2: สร้าง project config file
สร้างโฟลเดอร์ config ก่อน แล้วบันทึก JSON config เป็นไฟล์ที่:
```
d:\Work\DocISOGen\configs\[ProjectCode]\[ProjectCode]_config.json
```
> ✅ Generator จะ auto-detect ว่า config อยู่ใน `configs/[ProjectCode]/` แล้ว output ไปยัง `projects/[ProjectCode]/` โดยอัตโนมัติ

### Step 3: สร้างโฟลเดอร์ output
Generator จะสร้างโฟลเดอร์ย่อยอัตโนมัติใน `projects\[ProjectCode]\`:
```
d:\Work\DocISOGen\projects\[ProjectCode]\
├── 01_Project_Management\
├── 02_Requirements_Analysis\
├── 03_Design_Architecture\
├── 04_Development\
├── 05_Testing_QA\
├── 06_Deployment_Training\
├── 07_Support_Maintenance\
├── 08_Change_Logs_Versioning\
├── 09_Risk_Management\
└── 10_Regulatory_Compliance\
```

### Step 4: สร้างเอกสาร — รัน main generator
```powershell
C:/Users/Admin/AppData/Local/Programs/Python/Python312/python.exe d:\Work\DocISOGen\generator\generate_iso_docs.py --config "d:\Work\DocISOGen\configs\[ProjectCode]\[ProjectCode]_config.json"
```

ระบบจะสร้างเอกสารทั้งหมด 20 ไฟล์ในโฟลเดอร์ที่กำหนดในคำสั่งเดียว
> ✅ Generator จะ auto-detect ว่า config อยู่ใน `configs/[ProjectCode]/` แล้ว output ไปยัง `projects/[ProjectCode]/` โดยอัตโนมัติ

หากต้องการสร้างเฉพาะ folder เดียว (ใช้กับ subagent):
```powershell
... generate_iso_docs.py --config "d:\Work\DocISOGen\configs\[ProjectCode]\[ProjectCode]_config.json" --folder 05
```

---

## Document ID Convention
```
[ProjectCode]-[FolderNum]-[DocCode]-v[Version]
ตัวอย่าง: HRMS-02-RTM-v1.0
```

| Folder | DocCode |
|--------|---------|
| 01 | PP (Project Plan), MM (Meeting Minutes) |
| 02 | BRD (Requirements Doc), RTM |
| 03 | SDD (System Design), DBD (DB Design) |
| 04 | CS (Coding Standards), CRR (Code Review) |
| 05 | TP (Test Plan), TCR (Test Cases), BDL (Bug Log) |
| 06 | UM (User Manual), TR (Training Record) |
| 07 | ISL (Incident Log) |
| 08 | CRF (Change Request), VRN (Version Notes) |
| 09 | RR (Risk Register) |
| 10 | IC (ISO Checklist), AR (Audit Report), CAPA |

---

## Cross-Reference Map
เอกสารเหล่านี้ต้องใช้ข้อมูลร่วมกัน:

```
Requirements (REQ-xxx)
    ↓ referenced by
RTM (02-RTM) ←→ Test Cases (05-TCR) ←→ Bug Log (05-BDL)
    ↓                                          ↓
Design Components (03-SDD)          Change Request (08-CRF)
    ↓
Code Review (04-CRR)

Risks (RISK-xxx) ←→ CAPA (10-CAPA) ←→ Change Request (08-CRF)
Risk Register (09-RR) ←→ ISO Checklist (10-IC) ←→ Audit Report (10-AR)
```

---

## Subagent Parallelism Groups
Group A (สร้างได้พร้อมกัน — independent):
- Folder 01 (Project Management)
- Folder 03 (Design Architecture)
- Folder 04 (Development)
- Folder 09 (Risk Management)

Group B (รอ Group A เสร็จ):
- Folder 02 (Requirements) — รอ 01 มีลำดับ requirements
- Folder 06 (Deployment/Training)
- Folder 07 (Support)
- Folder 08 (Change Logs)

Group C (รอ Group B เสร็จ):
- Folder 05 (Testing — รอ RTM จาก 02)
- Folder 10 (Compliance — รอ Risk จาก 09 และ test results จาก 05)

---

## Template Pattern
เอกสารทุกฉบับมีโครงสร้างเดียวกันเสมอ:
1. Cover Page (ชื่อเอกสาร, Document ID, Version, Date, Organization)
2. Document Control (Prepared by, Reviewed by, Approved by)
3. Version History Table
4. Table of Contents (หัวข้อหลัก)
5. เนื้อหาตามมาตรฐาน ISO/IEC 29110
6. Signature / Sign-off Table (สำหรับเอกสารที่กำหนด)

---

## ISO/IEC 29110 Compliance Notes
- ทุกเอกสารต้อง traceable กลับไปหา requirement
- Test Case ต้องมี REQ-ID อ้างอิง
- Change Request ต้องมี approval signature
- Risk Register ต้องมี review date
- Audit Report ต้องอ้างอิง ISO Checklist item

---

## Files ที่สำคัญ
- `generator/config_template.json` — schema ของ project config
- `generator/generate_iso_docs.py` — main entry point
- `generator/utils/doc_builder.py` — common document helpers
- `generator/templates/template_XX_*.py` — เนื้อหาแต่ละเอกสาร
