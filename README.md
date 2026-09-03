# YC Startup Outreach Database

Python script that builds a corporate outreach database from publicly available Y Combinator startup information.

## Output

Generates:

* `yc_outreach_database.xlsx`
* `yc_outreach_database.csv`

Columns include company, YC batch, headquarters email, and up to two founders with their publicly available LinkedIn URLs and emails.

## Setup

```bash
git clone <your-repo-url>
cd yc-outreach
python -m venv .venv
source .venv/bin/activate
pip install requests beautifulsoup4 pandas openpyxl lxml tqdm
```

## Run

Test with 20 companies:

```bash
python yc_outreach.py --max-companies 20 --workers 8
```

Run the full dataset:

```bash
python yc_outreach.py --workers 8
```

The script only collects publicly displayed professional contact information and does not guess private email addresses.
