# Lead-List
Albacete MedDev — Lead List Processing Pipeline
Automated lead enrichment, verification, tiering, MAC jurisdiction mapping, and personalized outreach generation for orthopedic surgeon prospecting.
What This Does
Takes raw AcuityMD HCP targeting CSV exports and produces a fully enriched, tiered, outreach-ready Excel workbook. The pipeline:

Merges & deduplicates multiple AcuityMD CSV exports by NPI number
Classifies practices as Private Practice vs Hospital-Based using site-of-care name matching
Tiers leads by drive time from NYC (Midtown) using zip code geocoding:

Tier 1: 0–30 min
Tier 2: 30–60 min
Tier 3: 60–120 min
Tier 4: 120–180 min
Tier 5: 180+ min (drivable)
Tier 6: Requires flight
Hospital-Based (separate tab regardless of distance)


Verifies phone numbers against the CMS NPPES NPI Registry API — confirms, updates, or adds practice phone numbers for every lead
Validates and enriches emails through:

Pattern analysis: detects email format from known colleagues at the same practice (e.g., flast@domain.com) and infers missing emails
Classification: flags as Verified, Personal, Hospital System, Generic Office, Inferred, or Missing
Web search fallback for missing emails at priority practices


Maps MAC jurisdictions to determine Microlyte SAM eligibility by state:

NGS states (NY, CT, MA, ME, NH, VT, RI): Microlyte eligible (no active skin substitute LCD)
Novitas states (NJ, PA, MD, DE, DC): Microlyte not eligible (LCD L35041 active)
WPS states (MI, WI, MN, etc.): Microlyte eligible (verify L37228 doesn't restrict post-op use)
Palmetto states (VA excl. NoVA, WV, NC, SC): Microlyte eligible


Generates personalized cold outreach emails for every lead — two template tracks:

Track A (ProPacks Only): For Novitas/LCD-state surgeons
Track B (ProPacks + Microlyte SAM): For NGS/non-LCD-state surgeons
Emails personalize on: surgeon name, practice name, city, procedure focus (knee vs hip vs shoulder), case volume tier


Outputs a formatted Excel workbook with:

Separate tabs per tier (Tier 1 through Tier 6 + Hospital-Based)
Color-coded status columns (green = verified, red = missing, yellow = added)
Frozen headers, auto-filters, sorted by procedure volume descending
Subject line + full draft email per lead



Repository Structure
lead-list-processing/
├── README.md
├── CLAUDE.md                  # Claude Code project instructions
├── requirements.txt
├── config/
│   ├── mac_jurisdictions.json # State → MAC → LCD mapping
│   ├── hospital_keywords.json # Keywords for practice classification
│   └── email_templates.json   # Outreach email templates (Track A & B)
├── src/
│   ├── __init__.py
│   ├── ingest.py              # CSV ingestion, merge, dedup by NPI
│   ├── classify.py            # Practice classification (Private vs Hospital)
│   ├── tier.py                # Drive-time tiering from NYC via zip geocoding
│   ├── nppes.py               # NPPES NPI Registry API verification
│   ├── email_enrich.py        # Email validation, pattern inference, classification
│   ├── mac_mapping.py         # MAC jurisdiction + Microlyte eligibility
│   ├── outreach.py            # Personalized email generation
│   └── export.py              # Excel workbook generation (openpyxl)
├── data/
│   ├── input/                 # Drop raw AcuityMD CSVs here
│   └── output/                # Enriched workbooks land here
├── tests/
│   ├── test_nppes.py
│   ├── test_classify.py
│   ├── test_tier.py
│   └── test_mac_mapping.py
└── run.py                     # Main pipeline entrypoint
Setup
bashpip install -r requirements.txt
Dependencies
pandas>=2.0
openpyxl>=3.1
pgeocode>=0.4
requests>=2.31
Usage
Basic Run
Drop your AcuityMD CSV exports into data/input/, then:
bashpython run.py
This processes all CSVs in data/input/, runs the full pipeline, and outputs the enriched workbook to data/output/.
With Options
bash# Process specific files
python run.py --files joint_replacement.csv open_ortho.csv

# Only process Tiers 1-4 (skip Tier 5, 6, Hospital-Based)
python run.py --tiers 1 2 3 4

# Set origin city (default: NYC Midtown)
python run.py --origin-lat 40.7580 --origin-lon -73.9855

# Skip NPPES verification (use cached results)
python run.py --skip-nppes

# Skip email generation
python run.py --skip-emails

# Custom output filename
python run.py --output my_lead_list.xlsx
Configuration
config/mac_jurisdictions.json
Maps each US state to its Part B MAC contractor and whether an active skin substitute LCD exists. This determines Microlyte SAM eligibility. Update this file if CMS withdraws or publishes new LCDs.
json{
  "NY": {"mac": "NGS", "jurisdiction": "JK", "lcd_active": false, "microlyte_eligible": true},
  "NJ": {"mac": "Novitas", "jurisdiction": "JL", "lcd_active": true, "lcd_ids": ["L35041", "L35125"], "microlyte_eligible": false},
  ...
}
Critical nuance — Virginia: Most of Virginia falls under Palmetto GBA (Microlyte eligible), but Arlington County, Fairfax County, and City of Alexandria are carved out to Novitas (not eligible). The pipeline flags all VA leads for manual county-level review.
config/hospital_keywords.json
Keywords used to classify a site-of-care as Hospital-Based. Includes health systems, university hospitals, VA hospitals, and named hospital networks. Add new entries as you encounter misclassified practices.
config/email_templates.json
Two template tracks with merge fields:

{{first_name}}, {{last_name}}, {{practice_name}}, {{city}}
{{procedure_focus}} — auto-detected from volume data (knee, hip, shoulder, general joint replacement)
{{volume_hook}} — volume-appropriate language (high-volume vs standard)
Track A: ProPacks only (Novitas states)
Track B: ProPacks + Microlyte SAM (NGS states)

Sender Configuration
Update sender info in config/email_templates.json:
json{
  "sender_name": "Gabe Albacete",
  "sender_title": "Owner, Albacete MedDev",
  "sender_email": "gabe@albacetemeddev.com"
}
Pipeline Details
NPPES Verification (src/nppes.py)
Queries the CMS NPPES NPI Registry API (https://npiregistry.cms.hhs.gov/api/?version=2.1) for every lead by NPI number. Returns:

Verified practice phone number
Provider credentials
Practice address (location vs mailing)
Taxonomy (specialty) confirmation
Enumeration status (active/deactivated)

Rate limited to ~6 requests/second. Full 632-lead run takes ~2 minutes. Results are cached locally to avoid re-querying on subsequent runs.
Phone status values:

Verified — original phone matches NPPES record
Added from NPPES — no phone on file, NPPES phone added
Updated (NPPES differs) — original phone differs from NPPES; NPPES phone used, original preserved
Missing — no phone on file and NPI not found in NPPES

Email Enrichment (src/email_enrich.py)
Three-pass approach:

Classify existing emails: Pattern-match against name (first.last, flast, etc.) and domain type (practice, hospital, personal, generic)
Infer missing emails: For leads at practices where other colleagues have known emails, detect the email pattern and domain, then generate the missing email (e.g., if jsmith@orthonj.com exists, infer mkay@orthonj.com for another OrthoNJ surgeon). Only uses practice/hospital domains — never infers to gmail/yahoo.
Flag for manual sourcing: Remaining missing emails get flagged with practice phone number for manual follow-up

Email status values:

Verified (name + practice domain) — email contains surgeon's name and uses a practice domain
Personal Email (name match) — free email provider but name matches (gmail, yahoo, etc.)
Personal Email (no name match) — free email provider, name doesn't match; may be outdated
Hospital System Email — uses a hospital/health system domain
Inferred (pattern@domain) — generated from same-practice colleague pattern
Practice Email (review recommended) — practice domain but name doesn't match; could be shared inbox
Generic Office Email — info@, office@, admin@ etc.
Missing — no email found; needs manual sourcing

Drive-Time Tiering (src/tier.py)
Uses pgeocode for zip code geocoding and haversine distance calculation from NYC Midtown (40.7580, -73.9855). Converts straight-line miles to approximate drive time using road-travel multipliers:
Distance (miles)TierApprox Drive Time0–2010–30 min20–45230–60 min45–90360–120 min90–1404120–180 min140–3505180+ min (drivable)350+6Requires flight
Hospital-Based leads are placed in a separate tab regardless of distance.
Practice Classification (src/classify.py)
Keyword-based matching against the Primary Site of Care field. Over 150 hospital/health system keywords including named systems (Atlantic Health, Hackensack Meridian, Northwell, NYU Langone, etc.) and generic patterns (hospital, medical center, university of, etc.).
Any unmatched site defaults to Private Practice. Manual review is recommended for edge cases like "University Orthopedics Center" (private despite the name).
Products Referenced in Outreach
ProPacks (pitched to all leads)
Standardized post-operative incision care kits for joint replacement patients. Value props:

Better, more predictable patient outcomes
Reduced physician time and effort on post-op wound management
Full compliance and audit protection
Increased per-case profitability

Microlyte SAM (pitched only in non-LCD states)
Advanced antimicrobial wound matrix (510(k) classified as wound dressing, HCPCS A2005). Key features:

Fully synthetic, bioresorbable
Ionic and metallic silver for 99.99% microbial reduction for 72+ hours
"Peel and place" application — no painful removal
Day 1 use supported in non-LCD MAC jurisdictions (no 30-day conservative care failure requirement)

Only pitch Microlyte to surgeons in states where no active skin substitute LCD exists. The pipeline handles this automatically via MAC jurisdiction mapping.
Updating the Pipeline
Adding new AcuityMD exports
Drop additional CSVs into data/input/. The merge logic handles overlapping NPIs by outer-joining on shared columns and preserving all procedure volume columns unique to each export.
Updating MAC/LCD mappings
Edit config/mac_jurisdictions.json when CMS issues new LCDs or withdraws existing ones. The December 2025 CMS withdrawal of seven newly finalized skin substitute LCDs is already reflected. Monitor CMS MCD for changes.
Adding new hospital keywords
Edit config/hospital_keywords.json. Run python -m tests.test_classify to verify no private practices get misclassified.
Changing the origin city
Pass --origin-lat and --origin-lon to run.py, or update the defaults in src/tier.py.
Output Schema
Each tier tab in the output workbook contains these columns:
ColumnSourceDescriptionHCP NPIAcuityMD10-digit National Provider IdentifierFirst NameAcuityMDLast NameAcuityMDCredentialNPPESMD, DO, etc.SpecialtyAcuityMDe.g., "Orthopaedic Surgery > Adult Reconstructive"EmailAcuityMD + EnrichmentBest available emailEmail StatusPipelineVerified / Inferred / Missing / etc.Verified PhoneNPPESNPPES-verified practice phonePhone StatusPipelineVerified / Added / UpdatedPrimary Site of CareAcuityMDPractice or hospital namePractice TypePipelinePrivate Practice / Hospital-BasedAddress, City, State, ZipAcuityMDTierPipelineDrive-time tier from NYCMAC JurisdictionPipelineNGS / Novitas / Palmetto / etc.Microlyte EligiblePipelineYes / NoJoint Repl VolAcuityMDTotal joint replacement proceduresKnee VolAcuityMDKnee replacement proceduresHip VolAcuityMDHip replacement proceduresShoulder VolAcuityMDShoulder replacement proceduresOpen Ortho VolAcuityMDOpen orthopedic proceduresMedical SchoolAcuityMDHCP URLAcuityMDLink to AcuityMD profileSubject LinePipelinePersonalized email subjectDraft EmailPipelineFull personalized outreach email
