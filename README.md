# Local application packet builder

Create a tailored, ATS-friendly application packet from a job description and a candidate profile you maintain locally. The tool produces a Word resume, an application review sheet, and an optional cover letter. It does not sign in to a job site, upload files, answer screening questions, or submit an application.

## Requirements

- Python 3.10 or later
- `python-docx`
- LibreOffice for PDF export

Install the Python dependency:

```sh
python -m pip install -r requirements.txt
```

## Setup

Copy the example profile and replace its placeholders with information you have verified:

```sh
cp config/candidate_profile.example.json config/candidate_profile.json
```

The real candidate profile and generated application packets are excluded by `.gitignore`.

## Prepare a packet

Save a job description as a UTF-8 text file, then run:

```sh
python scripts/linkedin_application_workflow.py prepare \
  --company "Example Company" \
  --title "Forward Deployed Engineer" \
  --description /path/to/job-description.txt \
  --listing-url "https://www.linkedin.com/jobs/view/1234567890/" \
  --remote "Remote, United States" \
  --salary "$140K-$180K"
```

The packet is written under `output/applications/`. Use `--cover-letter required` or `--cover-letter none` to override the default, which creates a cover letter when the job description explicitly requests one. Use `--profile` or `--output-root` to select alternate local paths. Use `--force` only when you intend to replace the matching packet folder.

## Review before applying

Read the generated review sheet and inspect the resume and cover letter. Confirm all role details, dates, metrics, screening answers, and contact information. The matching logic selects only facts in the supplied profile, but the output still needs human review. This tool prepares documents only and does not submit applications.

## Privacy

Do not commit candidate profiles, job descriptions, resumes, application packets, or contact details. The repository ignores those local files by default. Keep job descriptions only when you have permission to store them.

## Implementation notes

The script uses keyword and tag overlap to rank profile skills and accomplishments for a role, then writes a one-column Word document. PDF export uses a local LibreOffice installation. No API keys, external services, or network calls are needed to prepare a packet.
