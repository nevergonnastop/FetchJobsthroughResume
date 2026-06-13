# Personal Job Radar

This is a lightweight personal job tracker. It reads company career URLs, fetches jobs, scores them against your resume/profile, and updates one Excel workbook.

## Files

- `companies.xlsx` - company list to edit
- `profile.json` - target roles, locations, scoring filters
- `resume.txt` or `resume.pdf` - your resume input
- `output/job_matches.xlsx` - updated results after every run
- `run_job_scan.ps1` - easiest Windows runner

## First setup

Run this once from PowerShell:

```powershell
.\run_job_scan.ps1
```

The first run creates the templates if they do not exist.

Edit:

- `companies.xlsx`: add company name and career URL
- `profile.json`: adjust preferred skills, locations, max experience, minimum score
- `resume.txt`: paste your resume text, or put a `resume.pdf` in this folder

## Run every time

From `C:\Users\lkmee\Desktop\CodexDData`:

```powershell
.\run_job_scan.ps1
```

Then open:

```text
output\job_matches.xlsx
```

The workbook is rebuilt every run, but your manual `Status` and `Notes` are preserved when the job key/apply link stays the same.

## Excel tabs

- `Top Matches`: jobs above your `minimum_score`
- `All Jobs`: everything fetched
- `Low Match`: below threshold
- `Applied Tracker`: jobs where status is not `Not Applied`
- one tab per company
- `Company Status`: fetch status and errors

## Notes

Best supported ATS pages:

- Ashby
- Greenhouse
- Lever
- SmartRecruiters

Other company pages use a simple generic scrape, so direct ATS URLs work best.
