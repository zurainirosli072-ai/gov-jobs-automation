---
name: extractjob
description: >
  ALWAYS use this skill when the user types /extractjob or asks to extract, scrape, or collect
  job details from job portal URLs (Workday, JPA, SPA, eRekrut, or any government/public sector
  portal) into an Excel tracker. Also trigger when the user says things like: "add these links
  to the job sheet", "populate Gov Jobs.xlsx", "get job details from this link", "I have new gov
  job links", "update the job tracker", "prepare jobs for AJobThing", or pastes one or more
  job posting URLs alongside any request about job data. Use this skill immediately in all
  those cases — do not wait to be asked twice.
  Phase 2 handles bulk posting to AJobThing once their bulk upload feature is live — trigger on
  "bulk post to AJobThing", "upload jobs to AJobThing", or "post all jobs".
---

# extractjob — Government Job Data Extractor for AJobThing

Two-phase workflow for posting government/public sector jobs on AJobThing:

- **Phase 1 (Live now):** Scrape job data from source portal links → populate `Gov Jobs.xlsx`
- **Phase 2 (Coming soon):** Bulk upload the populated Excel to AJobThing's bulk post feature

---

## Quick Start

When `/extractjob` is triggered or job links are provided:

1. Find `Gov Jobs.xlsx` in the workspace (`C:\Users\AJTADMIN\Desktop\Gov Jobs\`)
2. Read it — identify rows where Column A has a URL but other columns are empty
3. For each unfilled row: open the link in the browser, extract all fields, write them back
4. Save and present the file to the user

If the user provides new links not yet in the file, add them to the next empty rows in Column A first, then extract.

---

## Excel Structure

`Gov Jobs.xlsx` always uses these exact 9 columns:

| Col | Header | What to fill |
|-----|--------|-------------|
| A | Job Link | Source URL — input only, never overwrite |
| B | Company Name | Full legal name (e.g. "KPJ Healthcare Berhad") |
| C | Job Title | Exact title from the portal, as shown |
| D | Job Specialization | Category from the mapping table below |
| E | Job Description | 2–4 sentence summary of key duties |
| F | role location | **Specific** facility or hospital name (e.g. "KPJ Pahang Specialist Hospital") |
| G | Company brand | **Short** brand only (e.g. "KPJ") — Never the full company name |
| H | Language | Languages stated in posting, or "Not specified" |
| I | Minimum education | Minimum qualification (e.g. "SPM or equivalent") |

If the file doesn't exist yet, create it with these exact headers before proceeding.

---

## Phase 1 — Extracting Job Data

### Navigating the portal

Use the browser navigation tool to open each link. Many government portals (especially Workday — recognisable by `myworkdayjobs.com` in the URL) are JavaScript-rendered. After navigating, call `get_page_text` and check whether the job content has loaded. If the page looks empty or generic, wait a moment and try again.

### What to extract and where to find it

**Job Title (Col C)** — The largest heading at the top of the listing. Copy it exactly as written.

**Company Name (Col B)** — Usually in a "Company Overview" section near the bottom, or in the page header/branding. Use the full legal name (e.g. "KPJ Healthcare Berhad").

**Role Location (Col F)** — Look for a "locations" label near the top of the page, just below the job title. This always shows the **specific facility, hospital, or campus name** (e.g. "Kuching Specialist Hospital", "KPJ Pahang Specialist Hospital"). Use exactly what is written there. Never write "Onsite" or leave this generic — if you see "locations: KPJ Pahang Specialist Hospital", write "KPJ Pahang Specialist Hospital".

**Company Brand (Col G)** — This is the **short name only** of the organisation group. The goal is a brief recognisable label, not the full legal name. Examples:
- "KPJ Healthcare Berhad" → `KPJ`
- "Universiti Malaya" → `UM`
- "Universiti Teknologi MARA" → `UiTM`
- "Jabatan Perkhidmatan Awam" → `JPA`
- "Tenaga Nasional Berhad" → `TNB`
- When in doubt: use the first word or the well-known acronym

**Language (Col H)** — Look for language requirements in "Job Requirement" or "Kelayakan" sections. If none are stated, write "Not specified". Common values: "English and Bahasa Malaysia", "Mandarin", "Not specified".

**Minimum Education (Col I)** — Look for "Job Requirement", "Education", or "Kelayakan" sections. Write exactly what's stated, e.g. "SPM or equivalent", "Minimum SPM/STPM", "Diploma".

**Job Description (Col E)** — Write a 2–4 sentence summary of what the person actually does day-to-day. Do not copy-paste the entire job description. Think of it as a quick WhatsApp summary: "This role handles patient billing, processes insurance claims, and liaises with corporate clients."

### Job Specialization mapping (Col D)

Pick the closest category based on the job title:

| If the title contains… | Use this specialization |
|------------------------|------------------------|
| Receptionist, Front Desk, Clerk, Admin, Secretary | Administration / Customer Service |
| Cashier, Finance, Accounts, Billing, Audit | Finance / Billing |
| Nurse, Doctor, Medical Officer, Pharmacist, Dentist | Healthcare / Medical |
| IT, Software, Developer, System, Network | Information Technology |
| HR, Human Resource, Recruitment, Training | Human Resources |
| Engineer, Technician, Mechanic, Electrician | Engineering / Technical |
| Lecturer, Teacher, Tutor, Professor | Education |
| Security, Guard | Security |
| Cleaner, Housekeeping, Janitor | General Worker |
| Driver, Despatch, Logistics, Warehouse | Transportation / Logistics |
| Anything else | General / Others |

### Saving back to Excel

Use `openpyxl` to load and update the file so existing formatting is preserved:

```python
from openpyxl import load_workbook
from openpyxl.styles import Alignment

wb = load_workbook('Gov Jobs.xlsx')
ws = wb.active
ws.cell(row=i, column=2).value = company_name       # B: full legal name
ws.cell(row=i, column=3).value = job_title
ws.cell(row=i, column=4).value = specialization
ws.cell(row=i, column=5).value = job_description
ws.cell(row=i, column=5).alignment = Alignment(wrap_text=True, vertical='top')
ws.cell(row=i, column=6).value = role_location      # F: specific facility name
ws.cell(row=i, column=7).value = company_brand      # G: short brand only (e.g. "KPJ")
ws.cell(row=i, column=8).value = language
ws.cell(row=i, column=9).value = min_education
wb.save('Gov Jobs.xlsx')
```

Recommended column widths: A=70, B=25, C=35, D=28, E=60, F=30, G=15, H=28, I=22
Row height for data rows: 80

After saving, read the file back with pandas and print the new rows to confirm correctness before presenting the file.

---

## Common Issues

| Situation | What to do |
|-----------|-----------|
| Page still loading / blank | Wait 3–5 seconds, call `get_page_text` again |
| Location shows generic "Onsite" | Check the "locations" label near the top of the page — use that specific name |
| Company brand is unclear | Use the most recognisable short form or acronym |
| Language requirement not mentioned | Write "Not specified" |
| Education requirement not found | Write "Not specified" |
| Job description is very long | Summarise to 2–4 sentences only |
| Excel file permission denied | Ask user to close the file in Excel, then retry |
| Portal requires login to view | Let the user know and ask them to share the details |

---

## Phase 2 — Bulk Post to AJobThing *(Activate when feature is live)*

> The bulk post feature is under development at AJobThing. When it launches, use this section.

### Pre-upload checklist
- All rows have Job Title, Company Name, Specialization, Location, Language, Education filled
- Descriptions are concise (2–4 sentences)
- Company brand is short and consistent (e.g. all "KPJ", not mixed with "KPJ Healthcare Berhad")
- No required fields are blank

### Steps when live
1. Log in at `https://www.ajobthing.com`
2. Go to **Post a New Job** → look for **"Bulk Upload"** or **"Import from Excel"**
3. If AJobThing requires their own template, map columns:

| Gov Jobs.xlsx column | AJobThing field |
|---------------------|----------------|
| Job Title | Job Title |
| Company Name | Company / Brand Name |
| Job Specialization | Specialization |
| Job Description | Job Description |
| role location | Primary Role Location |
| Language | Required Language(s) |
| Minimum education | Minimum Education |

4. Upload, review the preview, confirm and publish

**Education dropdown values:** SPM/O-Level → "SPM / O Level" | Diploma → "Diploma" | Degree → "Bachelor's Degree" | Master → "Master's Degree"

**Language checkboxes:** English | Bahasa Malaysia | Mandarin | Tamil — tick all that apply; if "Not specified", tick English as default

---

## Workflow summary

```
/extractjob triggered or job links provided
        ↓
Read Gov Jobs.xlsx → find rows with links but missing data
        ↓
For each link: navigate → wait for load → get_page_text → extract fields
        ↓
Apply Specialization mapping + shorten Company brand + use specific Location name
        ↓
Write back to Excel with openpyxl → verify with pandas → save
        ↓
[Now]  Present updated file to user
[Soon] Upload to AJobThing Bulk Post → review → publish
```
