# AUDIT ACTION PLAN
**Project:** Dvets (Digital Vehicle and Entry Tracking System) | **Date:** 2026-03-05 | **Result:** CONDITIONAL

## 🔴 Major Non-Conformances
| # | Folder | Finding | แนวทางแก้ไข | Due | Owner |
|---|--------|---------|------------|-----|-------|
| 1 | 03_Design_Architecture | SDD — Architecture pattern ยังเป็น placeholder `[e.g. 3-Tier / Microservices / MVC]` | replace_in_paragraph: ระบุ Architecture จริงของ Dvets | 2026-03-12 | Souliya Singvongsa |
| 2 | 03_Design_Architecture | DBD — ERD Diagram ยังไม่แนบ `(แนบ diagram ที่นี่)` | แนบ ERD จริงจาก draw.io หรือ dbdiagram.io | 2026-03-19 | Khamta Phantaboun |
| 3 | 04_Development | CRR — Checklist 10 checkboxes ยังไม่ถูก tick | replace_checkbox: col_header "pass" → ☐ เป็น ☑ | 2026-03-05 | Souliya Singvongsa |
| 4 | 04_Development | CRR — Duplicate ID: ISS-001 ซ้ำกัน | fix_duplicate_id patch | 2026-03-05 | Souliya Singvongsa |
| 5 | 06_Deployment_Training | UM — Placeholders 3 จุด: `[ระบุเมนู]`, `[ระบุขั้นตอน]`, `(แนบ screenshot)` | replace_in_paragraph สำหรับ menu/steps; แนบ screenshot จริง | 2026-03-19 | Siwarit Chantham |
| 6 | 07_Support_Maintenance | ISL — INC-001 สถานะ Open ทั้งหมด ทั้งที่ project go-live แล้ว | update_table_status: เปลี่ยนเป็น Resolved | 2026-03-12 | Houmphaeng Duangpaserth |
| 7 | 09_Risk_Management | RR — RISK-001/002/003 ทั้ง 3 ยังสถานะ Open หลัง go-live | update_table_status: ปิด/update risk status | 2026-03-12 | Chan Neeammart |
| 8 | 10_Regulatory_Compliance | IC (ISO Checklist) — ไฟล์หายทั้งหมด | รัน generator --folder 10 | 2026-03-05 | Chan Neeammart |
| 9 | 10_Regulatory_Compliance | AR (Audit Report) — ไฟล์หายทั้งหมด | รัน generator --folder 10 | 2026-03-05 | Chan Neeammart |
| 10 | 10_Regulatory_Compliance | CAPA — ไฟล์หายทั้งหมด | รัน generator --folder 10 | 2026-03-05 | Chan Neeammart |

## 🟡 Minor Non-Conformances
| # | Folder | Finding | แนวทางแก้ไข | Due | Owner |
|---|--------|---------|------------|-----|-------|
| 1 | 01_Project_Management | MM — Missing Signatures 6 รายการ (วันที่/เวลา/สถานที่) | fill_empty_dates patch + กรอกข้อมูลจริง | 2026-03-12 | Chan Neeammart |
| 2 | 01_Project_Management | MM — Placeholder text `_______________` ยังอยู่ | replace_in_paragraph | 2026-03-12 | Chan Neeammart |
| 3 | 01_Project_Management | PP — ไม่มี REQ IDs — ขาด traceability กับ Requirements | เพิ่ม Requirements reference section ใน Project Plan | 2026-03-19 | Chan Neeammart |
| 4 | 01_Project_Management | MM — พบเพียง MS-03 — Milestone set ไม่ครบ | เพิ่ม MS-01, MS-02 ใน Meeting Minutes | 2026-03-19 | Chan Neeammart |
| 5 | 02_Requirements_Analysis | RTM — ไม่พบ REQ IDs ใน snapshot (อาจเป็น template เปล่า) | ตรวจสอบและ map REQ→COMP→TC ใน RTM | 2026-03-19 | Souliya Singvongsa |
| 6 | 02_Requirements_Analysis | BRD — placeholder_cells 6 จุด (version history ว่าง) | fill_table_cell / fill_empty_dates | 2026-03-12 | Souliya Singvongsa |
| 7 | 03_Design_Architecture | SDD/DBD — ไม่พบ REQ IDs ใน Design docs | เพิ่ม REQ reference ใน Design components | 2026-03-19 | Souliya Singvongsa |
| 8 | 03_Design_Architecture | COMP-001 เดียวทั้ง project vs REQ 15 รายการ | เพิ่ม Component definitions ให้ครบ | 2026-03-19 | Souliya Singvongsa |
| 9 | 04_Development | CRR — ไม่พบ COMP IDs ใน Code Review | เพิ่ม component reference ใน CRR | 2026-03-19 | Souliya Singvongsa |
| 10 | 05_Testing_QA | TCR — empty_date_columns 2 จุด (Test execution dates ว่าง) | fill_empty_dates patch | 2026-03-12 | Wikanda Thongsuk |
| 11 | 05_Testing_QA | TP — ไม่พบ REQ IDs ใน Test Plan | เพิ่ม REQ scope/coverage mapping ใน Test Plan | 2026-03-19 | Wikanda Thongsuk |
| 12 | 06_Deployment_Training | TR/UM — ไม่มี UAT evidence / sign-off | เพิ่ม UAT acceptance record พร้อม sign-off | 2026-03-19 | Chan Neeammart |
| 13 | 07_Support_Maintenance | ISL — Resolved date ว่าง | fill_empty_dates (หลังจาก update status แล้ว) | 2026-03-12 | Houmphaeng Duangpaserth |
| 14 | 08_Change_Logs_Versioning | VRN — REQ-007, REQ-014, REQ-015 ไม่ปรากฏใน Release Notes | ตรวจสอบและเพิ่ม Release coverage | 2026-03-19 | Souliya Singvongsa |
| 15 | 08_Change_Logs_Versioning | CRF — Approval section placeholder_cells 18 จุด | fill_table_cell / ได้รับ approval จริง | 2026-03-19 | Chan Neeammart |
| 16 | 09_Risk_Management | RR — placeholder_cells 6 จุด (review section ว่าง) | fill_table_cell | 2026-03-12 | Chan Neeammart |

## 🟠 Cross-Document Inconsistencies
| # | Folders | ปัญหา | แนวทางแก้ไข | Due |
|---|---------|------|------------|-----|
| 1 | 02 ↔ 05 | RTM ไม่พบ REQ IDs → ไม่สามารถยืนยันว่า TC-001~015 map ครบกับ REQ-001~015 | แก้ RTM ให้มี REQ→TC mapping ชัดเจน จากนั้น verify ใน TCR | 2026-03-19 |
| 2 | 03 ↔ 04 | COMP มีเพียง COMP-001 ทั้ง project — CRR ควร reference Component หลายรายการ | เพิ่ม COMP IDs ใน SDD/DBD แล้ว align กับ CRR | 2026-03-19 |
| 3 | 05 ↔ 06 | ไม่มีหลักฐาน test pass / UAT sign-off ใน folder 06 | เพิ่ม UAT evidence ใน TR พร้อม reference TC pass | 2026-03-19 |
| 4 | 08 ↔ 02 | VRN cover 12/15 REQs — REQ-007, 014, 015 ยังไม่ปรากฏใน Change baseline | ตรวจสอบว่า REQs ดังกล่าว deploy แล้วหรือยัง และอัปเดต VRN | 2026-03-19 |

## 🔵 Observations
| # | Folder | ข้อสังเกต | คำแนะนำ |
|---|--------|----------|---------|
| 1 | ทุก folder | เอกสารทุกไฟล์ (17/17) มี table_issues ประเภท placeholder_cells | พิจารณาใช้ fill_table_cell patch แบบ batch เพื่อ populate ข้อมูล version history |
| 2 | 01_Project_Management | PP ไม่มี Risk section link ไปยัง folder 09 | เพิ่ม Risk Summary section ใน Project Plan ที่อ้างอิง RISK-001/002/003 |

## ✅ Summary
🔴 Major: 10 | 🟡 Minor: 16 | 🟠 Cross-doc: 4 | 🔵 Obs: 2
**Next Review:** 2026-04-05 (+30 days สำหรับ CONDITIONAL)

---

## 🤖 AI-Assisted Remediation Plan

### 🤖 AI แก้ได้ทันที
| # | Finding | ไฟล์ต้นทาง | การดำเนินการ |
|---|---------|------------|-------------|
| 1 | CRR — unticked_checkboxes 10 รายการ | DVETS-04-CRR-v4.2_Code_Review_Records.docx | `replace_checkbox` col_header="pass" |
| 2 | CRR — Duplicate ID ISS-001 | DVETS-04-CRR-v4.2_Code_Review_Records.docx | `fix_duplicate_id` |
| 3 | RR — RISK-001/002/003 all_open_status | DVETS-09-RR-v4.2_Risk_Register.docx | `update_table_status` → Mitigated |
| 4 | ISL — INC-001 all_open_status | DVETS-07-ISL-v4.2_Incident_Support_Log.docx | `update_table_status` → Resolved |
| 5 | ทุกไฟล์ — empty_date_columns (version history table_index 3) | ทุก 17 ไฟล์ | `fill_empty_dates` |
| 6 | TCR — empty_date_columns (table_index 21) | DVETS-05-TCR-v4.2_Test_Cases_Results.docx | `fill_empty_dates` |
| 7 | SDD — Architecture placeholder | DVETS-03-SDD-v4.2_System_Design.docx | `replace_in_paragraph` แทน `[e.g. 3-Tier / Microservices / MVC]` |
| 8 | SDD — Interface placeholder | DVETS-03-SDD-v4.2_System_Design.docx | `replace_in_paragraph` แทน `[ระบุ interface ที่เชื่อมต่อ]` |
| 9 | DBD — Tool placeholder | DVETS-03-DBD-v4.2_Database_Design.docx | `replace_in_paragraph` แทน `[e.g. draw.io / dbdiagram.io / ERwin]` |
| 10 | IC/AR/CAPA — ไม่มีไฟล์ใน folder 10 | (ไม่มีไฟล์) | รัน `generator --folder 10` |

### 🤝 AI ทำได้บางส่วน (ต้องการข้อมูลเพิ่ม)
| # | Finding | ส่วนที่ AI ทำได้ | ข้อมูลที่ต้องการ |
|---|---------|----------------|----------------|
| 1 | CRF — Approval placeholder 18 จุด | fill_table_cell สำหรับ field ที่ AI รู้ (Requestor, Date, Description) | วันที่ approve จริง, ชื่อ approver |
| 2 | UM — `[ระบุเมนู]` `[ระบุขั้นตอน]` | replace_in_paragraph สำหรับชื่อเมนู | รายชื่อ menu จริง และ step-by-step จาก dev |
| 3 | MM — Missing signature dates | fill_empty_dates สำหรับ format | วันที่จริง, เวลา, สถานที่ประชุม |
| 4 | ISL — Resolved date ว่าง | fill_empty_dates | วันที่ resolve จริงของ INC-001 |

### 🙋 AI ทำไม่ได้ (ต้องการการตัดสินใจ/ลายเซ็น)
| # | Finding | เหตุผล | ผู้รับผิดชอบ |
|---|---------|--------|------------|
| 1 | DBD — ERD Diagram ยังไม่แนบ | ต้องการ diagram จริง (draw.io/dbdiagram.io) | Khamta Phantaboun |
| 2 | UM — `(แนบ screenshot ที่นี่)` | ต้องการ screenshot จริงจากระบบที่ deploy แล้ว | Siwarit Chantham |
| 3 | MM — ลายเซ็นผู้เข้าร่วมประชุม | ต้องการลายเซ็นจริง | Chan Neeammart |
| 4 | CRF — Approval Sign-off | ต้องการลายเซ็น Approver จริง | Chan Neeammart |
| 5 | TR — UAT Sign-off | ต้องการการยืนยัน customer acceptance จริง | ผู้บังคับบัญชาด่านตรวจ |
| 6 | PP — Milestone MS-01, MS-02 | ต้องการวันที่จริงของ milestone ที่ผ่านมา | Chan Neeammart |
