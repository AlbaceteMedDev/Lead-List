CLAUDE.md — Lead List Processing Pipeline
Project Context
This is a lead enrichment and outreach automation pipeline for Albacete MedDev, a medical device distributor (wound care biologics, post-op incision care products) selling to orthopedic surgeons across the Northeast US. The pipeline takes raw AcuityMD HCP targeting CSV exports and produces an enriched, tiered, outreach-ready Excel workbook.
The owner is Gabe Albacete. The products being sold are ProPacks (post-op incision care kits — pitched to all leads) and Microlyte SAM (advanced antimicrobial wound dressing — pitched ONLY to surgeons in non-LCD MAC jurisdictions). The sales approach is direct-to-physician outreach: in-person lunch & learns or intro calls.
What the Pipeline Must Do
Step 1: Ingest & Merge

Read all CSV files from data/input/
AcuityMD exports have inconsistent column names across reports — the merge must handle this. Common identity columns: HCP NPI, First Name, Last Name, Middle Name, Prefix, Credential, Specialty, Phone Number, Email, Primary Site of Care, Address 1, Address 2, City, State, Postal Code, Medical School, Medical School Graduation Year, HCP URL
Procedure volume columns are report-specific (e.g., one CSV has joint replacement volumes, another has collagen/DME volumes). Outer join preserves all columns.
Deduplicate on HCP NPI. If the same NPI appears in multiple files, merge their procedure volume columns — don't create duplicate rows.
Column names arrive wrapped in extra quotes from AcuityMD CSV exports (e.g., ""Joint Replacement - Procedure Volume""). Strip these.
All columns should be read as strings initially to prevent NPI truncation or phone number formatting issues.

Step 2: Practice Classification
Classify each lead as Private Practice or Hospital-Based based on the Primary Site of Care field.
Use keyword matching against a configurable list in config/hospital_keywords.json. The list must include:

Explicit hospital/health system names (Atlantic Health, Hackensack Meridian, Northwell, NYU Langone, HSS, Mount Sinai, RWJB, Virtua, Geisinger, UPMC, Yale New Haven, Hartford Healthcare, etc.)
Generic patterns: "hospital", "medical center", "health system", "health network", "university of" (as a hospital indicator), "VA ", "veterans", etc.
Known NJ/NY/PA/CT/MD hospital systems relevant to this territory

Important edge case: "University Orthopedics Center" is a PRIVATE PRACTICE despite having "university" in the name. The keyword list should not overmatch on "university" alone — match on "university of" or "university hospital" or "university medical" to avoid this.
Default unmatched leads to Private Practice.
Step 3: Drive-Time Tiering
For Private Practice leads, calculate approximate drive time from NYC Midtown (40.7580, -73.9855) using zip code geocoding.
Use pgeocode with the US postal code database to geocode each lead's Postal Code (first 5 digits only). Calculate haversine distance in miles, then convert to tiers using these road-travel-adjusted thresholds:
Miles (straight-line)Tier0–20Tier 1 (0–30 min)20–45Tier 2 (30–60 min)45–90Tier 3 (60–120 min)90–140Tier 4 (120–180 min)140–350Tier 5 (180+ min drivable)350+Tier 6 (Requires flight)
Hospital-Based leads get tier = Hospital-Based regardless of distance.
If a zip code cannot be geocoded, default to Tier 5.
The origin point should be configurable (default NYC) for future territory expansion.
Step 4: NPPES Phone Verification
Query the CMS NPPES NPI Registry API for every lead:
GET https://npiregistry.cms.hhs.gov/api/?version=2.1&number={NPI}
Extract from each response:

Practice location phone number (from addresses[] where address_purpose == "LOCATION")
Practice fax number
Credential (from basic.credential)
Primary taxonomy description (from taxonomies[] where primary == true)
Enumeration status

Rate limiting: Max ~6 requests/second. Sleep 0.15s between requests. Full run of 632 leads takes ~2 minutes.
Caching: Cache NPPES results locally (JSON or pickle) keyed by NPI. On subsequent runs, skip NPIs already cached unless --force-nppes flag is set. Cache should include a timestamp so results older than 30 days can be refreshed.
Phone status logic:

Original phone matches NPPES phone (normalized to 10 digits) → Verified
No original phone, NPPES has one → Added from NPPES
Original phone differs from NPPES → Updated (NPPES differs) — use NPPES phone as primary, preserve original
No original phone, NPI not in NPPES → Missing

Step 5: Email Enrichment
Three-pass approach:
Pass 1 — Classify existing emails:
Check each lead's email against these categories (in priority order):

Missing — email field is null/empty
Generic Office Email — starts with: info@, office@, admin@, contact@, billing@, reception@, appointments@, scheduling@, frontdesk@, help@, support@, mail@, general@, hello@
Hospital System Email — domain matches known hospital domains (atlantichealth.org, rwjbh.org, hackensackmeridian.org, nyu.edu, nyulangone.org, mountsinai.org, northwell.edu, hss.edu, etc.)
Verified (name + practice domain) — email local part contains the lead's first or last name AND the domain is NOT a free email provider
Personal Email (name match) — domain is a free provider (gmail.com, yahoo.com, hotmail.com, aol.com, outlook.com, icloud.com, etc.) but local part contains the lead's first or last name
Personal Email (no name match) — free provider domain, name not in local part
Practice Email (review recommended) — practice domain but name not found in local part

Pass 2 — Infer missing emails from practice patterns:
For each lead with a Missing email, look at all OTHER leads at the same Primary Site of Care who DO have emails. If there's a pattern:

Detect the email format: first.last@domain, flast@domain, firstlast@domain, f.last@domain, last@domain, firstl@domain
Detect the most common domain at that practice
Generate the missing email using the detected pattern + domain

Critical filter: NEVER infer an email using a free email provider domain (gmail, yahoo, hotmail, etc.) or a hospital system domain that doesn't match the lead's actual practice. Only infer using practice-specific domains.
Pass 3 — Flag for manual sourcing:
Any emails still missing after Pass 2 get status Missing. The output workbook highlights these in red. The NPPES-verified phone number is available for manual follow-up ("call the practice, ask for Dr. X's direct email").
Step 6: MAC Jurisdiction Mapping
Map each lead's state to their Part B MAC contractor and determine Microlyte SAM eligibility.
Source of truth: config/mac_jurisdictions.json
Current mapping (as of April 2026):
Microlyte ELIGIBLE (no active skin substitute LCD):

NGS (Jurisdiction K): NY, CT, MA, ME, NH, VT, RI
Palmetto GBA (Jurisdiction J): VA (EXCEPT Arlington, Fairfax, Alexandria), WV, NC, SC
Noridian (Jurisdiction E/F): AK, WA, OR, ID, MT, WY, ND, SD, NE, IA, UT, AZ, CA, NV, HI, AS, GU, MP
WPS (Jurisdiction 5/8): MI, WI, MN, IN (verify L37228 doesn't restrict post-op)

Microlyte NOT ELIGIBLE (active LCD L35041/L35125):

Novitas (Jurisdiction L): NJ, PA, MD, DE, DC, VA (Arlington, Fairfax, Alexandria only)
CGS (Jurisdiction 15): OH, KY
First Coast (Jurisdiction N): FL, PR, USVI

Virginia special handling: Flag all VA leads for manual county-level review. Northern Virginia (Arlington County, Fairfax County, City of Alexandria) is carved out to Novitas. All other VA counties fall under Palmetto.
Step 7: Personalized Email Generation
Generate one cold outreach email per lead using two template tracks:
Track A — ProPacks Only (for leads where Microlyte Eligible == No):

Subject: Post-op incision care for {practice_name}
Pitch ProPacks only. Four value props: better outcomes, reduced physician time, compliance/audit protection, increased profitability
CTA: in-person lunch & learn or quick intro call

Track B — ProPacks + Microlyte SAM (for leads where Microlyte Eligible == Yes):

Subject: Post-op incision innovation for {practice_name}
Pitch both ProPacks AND Microlyte SAM. Mention Microlyte's 99.99% microbial reduction for 72+ hours, bioresorbable, "peel and place" application, HCPCS A2005
CTA: in-person lunch & learn or quick intro call

Personalization merge fields:

{first_name}, {last_name}, {practice_name}, {city}
{procedure_focus} — auto-detect from volume columns:

If Knee Vol > Hip Vol × 1.5 → "knee replacement"
If Hip Vol > Knee Vol × 1.5 → "hip replacement"
If Shoulder Vol > 200 → "shoulder and joint replacement"
Default → "joint replacement"


{volume_hook} — volume-appropriate language:

Joint Replacement Vol > 500 → "With the volume of {procedure_focus} cases you're managing"
Vol 200–500 → "Given your {procedure_focus} practice"
Vol < 200 or missing → "As an orthopedic surgeon performing {procedure_focus} procedures"



Sender block (all emails):
Gabe Albacete
Owner, Albacete MedDev
gabe@albacetemeddev.com
Step 8: Excel Output
Generate a single .xlsx workbook with these tabs (in order):

Summary — pipeline stats (total leads, per-tier counts, phone/email status breakdowns, Microlyte eligibility counts, template track split)
Tier 1 (0-30 min) — Private Practice leads only
Tier 2 (30-60 min) — Private Practice leads only
Tier 3 (60-120 min) — Private Practice leads only
Tier 4 (120-180 min) — Private Practice leads only
Tier 5 (180+ drivable) — Private Practice leads only
Tier 6 (Requires flight) — Private Practice leads only
Hospital-Based — all Hospital-Based leads

Each tier tab sorted by Joint Replacement - Procedure Volume descending.
Columns per tier tab:
HCP NPI | First Name | Last Name | Credential | Specialty | Email | Email Status | Verified Phone | Phone Status | Primary Site of Care | Practice Type | Address 1 | City | State | Postal Code | Tier | MAC Jurisdiction | Microlyte Eligible | Joint Repl Vol | Knee Vol | Hip Vol | Shoulder Vol | Open Ortho Vol | Medical School | HCP URL | Subject Line | Draft Email
Formatting:

Font: Arial, 10pt headers, 9pt data
Header row: white text on dark navy (#1B4F72), bold, center-aligned, wrap text, frozen
Alternating row fills: white / light blue (#EBF5FB)
Auto-filters on all columns
Color-coded status cells:

Email Status = Verified → green fill (#D5F5E3)
Email Status = Missing → red fill (#FADBD8)
Phone Status = Added from NPPES → yellow fill (#FEF9E7)
Microlyte Eligible = Yes → green fill (#D4EFDF)


Column widths: set per column (Email: 32, Draft Email: 60, Practice Name: 30, etc.)

Tech Stack

Python 3.10+
pandas — data manipulation, merge, dedup
openpyxl — Excel workbook generation with formatting
pgeocode — zip code geocoding for drive-time tiering
requests — NPPES API calls
No external paid APIs required (NPPES is free, pgeocode uses bundled GeoNames data)

Commands
bash# Install dependencies
pip install -r requirements.txt

# Run full pipeline
python run.py

# Run with specific input files
python run.py --files file1.csv file2.csv

# Only process Tiers 1-4
python run.py --tiers 1 2 3 4

# Force NPPES re-query (ignore cache)
python run.py --force-nppes

# Custom origin point (default: NYC Midtown)
python run.py --origin-lat 40.7580 --origin-lon -73.9855

# Run tests
pytest tests/
File Locations

Input CSVs: data/input/
Output workbooks: data/output/
NPPES cache: data/cache/nppes_cache.json
Config files: config/

Important Business Rules — Do Not Change Without Approval

Microlyte is NEVER pitched in LCD states. If the state has an active skin substitute LCD (Novitas, CGS, First Coast), the email must be Track A (ProPacks only). Getting this wrong creates compliance risk.
Virginia requires county-level MAC determination. Arlington, Fairfax, and Alexandria → Novitas (no Microlyte). All other VA counties → Palmetto (Microlyte eligible). Flag all VA leads for manual review.
Email inference must never use free email domains. If the only known emails at a practice are @gmail.com, do NOT infer other doctors' emails as @gmail.com. Only infer using practice-specific or hospital system domains.
Hospital-Based leads always go in their own tab regardless of distance from origin.
The sender is always: Gabe Albacete, Owner, Albacete MedDev, gabe@albacetemeddev.com
NPPES API rate limit: 0.15s between requests. Do not reduce this — CMS will throttle or block.
All data read as strings initially. NPI numbers and phone numbers must not be truncated by pandas type inference.

Testing

test_nppes.py — verify API parsing, phone normalization, caching
test_classify.py — verify practice classification edge cases (e.g., "University Orthopedics Center" = Private)
test_tier.py — verify distance calculations and tier boundaries
test_mac_mapping.py — verify state-to-MAC mapping, Microlyte eligibility, Virginia carve-out

Error Handling

If NPPES API is down or returns errors for specific NPIs, log the error and continue. Don't fail the entire pipeline.
If pgeocode can't geocode a zip code, default to Tier 5 and log a warning.
If a CSV has unexpected column names, log a warning and attempt fuzzy matching on common column patterns.
Never silently drop leads. If a lead can't be fully enriched, include it in the output with status columns reflecting what failed.
