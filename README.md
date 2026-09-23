# Databuild API Script

Automates pulling **approved project** data from your
[dbonline.co.za](https://www.dbonline.co.za) Databuild account into clean
Excel/CSV/JSON files, so your **Microsoft 365 Copilot agents** can reference them
in reports and proposals.

The script logs into your own Databuild account, reads the dashboard's project
data, filters to approved projects, and writes:

- `DataBuild_Approved_Projects.xlsx` — for Excel on Windows
- `DataBuild_Approved_Projects.csv` — UTF-8, opens directly in Excel
- `DataBuild_Approved_Projects.json` — machine-readable
- `last_run.json` — status of the last run

Default output location: `C:\Users\<you>\OneDrive\Databuild\` (set
`DATABUILD_OUTPUT_DIR` in `.env` to change it).

---

## 1. Install Python (once)

1. Download Python from <https://www.python.org/downloads/> (3.10 or newer).
2. During install **tick "Add python.exe to PATH"**, then install.
3. Open **Command Prompt** (Win+R → `cmd`) and verify:
   ```bat
   python --version
   ```

## 2. Get the script onto your PC

If you cloned with Git:

```bat
git clone https://github.com/Davedes83/Databuild-API-Script.git
cd Databuild-API-Script
```

Or just download the ZIP (green **Code** button → **Download ZIP**) and extract.

## 3. Install dependencies (once)

In Command Prompt, inside the folder:

```bat
python -m pip install -r requirements.txt
```

## 4. Create your .env file (secrets stay on your PC)

```bat
copy .env.example .env
```

Then open `.env` in Notepad and fill in:

```ini
DATABUILD_USERNAME=your-dbonline-email
DATABUILD_PASSWORD=your-dbonline-password
# optional:
DATABUILD_OUTPUT_DIR=C:\Users\you\OneDrive\Databuild
```

`.env` is git-ignored and never uploaded. Never paste these into chat or commit
them anywhere.

## 5. First run — probe the endpoints

DataBuild's internal API isn't publicly documented, so run once with `--probe`
to confirm the script can reach every candidate endpoint after login:

```bat
python databuild_pull.py --probe
```

This writes `probe_report.json` to your output folder showing which endpoints
respond. **If login failed**
(check `DATABUILD_USERNAME` / `DATABUILD_PASSWORD`), fix `.env` and rerun.
**If no JSON project list is found**, capture a browser network (HAR/DevTools)
log of your dashboard and we can update the endpoint map.

## 6. Run the pull (on demand)

Double-click **`run_databuild.bat`**, or run:

```bat
python databuild_pull.py
```

Result: `DataBuild_Approved_Projects.xlsx` (and `.csv` / `.json`) in your
output folder, ready to open in Excel on Windows.

### Useful flags

| Flag | What it does |
|---|---|
| `--probe` | Inspect endpoints and save `probe_report.json`, then stop. |
| `--all` | Keep non-approved statuses too (Status column shows them all). |
| `--statuses approved,awarded` | Change which statuses count as "approved". |
| `--output-dir C:\path` | Override the output folder for one run. |
| `--dotenv C:\path\.env` | Use a specific `.env` file. |

---

## Feeding the data into Copilot reports

The pulled workbook lands in your OneDrive, which your Copilot agents already
read. In **Copilot Studio** for each agent you want to use it (e.g. BidBuilder,
Plans/BOQ):

1. Open the agent → **Knowledge** → **Add knowledge** → **OneDrive and
   SharePoint**.
2. Point it at your `OneDrive/Databuild` folder.
3. In work Copilot Chat, ask:
   _"Include approved DataBuild projects in my BidBuilder proposal."_

The agent reads `DataBuild_Approved_Projects.xlsx` as its source. Re-run the
script (double-click `run_databuild.bat`) whenever you want fresh data.

## Not automated?

This repo is deliberately **on-demand only**. If you later want daily refresh,
use Windows Task Scheduler:

```bat
schtasks /create /sc daily /tn "DatabuildPull" /tr "C:\path\run_databuild.bat" /st 07:00
```

---

## Notes & honesty

- The pulled data uses DataBuild's internal dashboard API; endpoint paths can
  change. If a run stops working, run `--probe` and share `probe_report.json`
  to rebuild the map.
- Intended for **your own subscribed Databuild account**. Review their Terms of
  Service; DataBuild also sells an official integration endpoint if you need it
  formally sanctioned.
- The script shows its work; nothing is guessed. Fields it cannot find in a
  payload are left blank rather than invented.