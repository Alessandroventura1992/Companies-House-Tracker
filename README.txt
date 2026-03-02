============================================================
  UK COMPANY INCORPORATION TRACKER
  Setup & Operating Guide
============================================================

WHAT THIS DOES
--------------
This script automatically tracks newly incorporated UK companies
using the Companies House API. It runs every 30 minutes on your
PC and exports results to an Excel spreadsheet, organised by
industry category. Each company includes:

  - Company name and registration number
  - Incorporation date and registered address
  - Directors (with nationality and date of birth)
  - Beneficial owners / persons with significant control
  - Linked companies that share the same director

The Excel file updates automatically every 30 minutes while
the script is running.


------------------------------------------------------------
STEP 1: INSTALL PYTHON
------------------------------------------------------------
1. Open your web browser and go to:
   https://www.python.org/downloads/

2. Click the yellow "Download Python" button (latest version).

3. Run the installer. IMPORTANT: On the first screen, tick the
   box that says "Add Python to PATH" before clicking Install.

4. Click "Install Now" and wait for it to finish.

5. To confirm Python installed correctly:
   - Press Windows key + R
   - Type: cmd
   - Press Enter
   - In the black window, type: python --version
   - Press Enter
   - You should see something like: Python 3.12.0
   If you see an error, repeat Step 1 and make sure you ticked
   "Add Python to PATH".


------------------------------------------------------------
STEP 2: SET UP THE TRACKER FILES
------------------------------------------------------------
1. Create a new folder on your PC where you want the tracker
   to live. For example:
   C:\Users\YourName\Documents\CompanyTracker\

2. Copy all of the following files into that folder:
   - companies_house_tracker.py
   - config.json
   - requirements.txt
   - run_tracker.bat
   - README.txt (this file)

   All five files must be in the same folder.


------------------------------------------------------------
STEP 3: RUN THE TRACKER
------------------------------------------------------------
1. Open the folder you created in Step 2.

2. Double-click "run_tracker.bat"

3. A black console window will open. It will:
   - Automatically install the required software (first run only)
   - Connect to the Companies House API
   - Search for newly incorporated companies
   - Save results to "companies_tracker.xlsx" in the same folder

4. The first run searches the past 7 days of incorporations,
   so it may take 10-30 minutes to complete depending on how
   many companies are found. This is normal.

5. After the first run, it will automatically repeat every
   30 minutes. Leave the window open to keep it running.

DO NOT close the black console window while the tracker is
running. Minimising it is fine.


------------------------------------------------------------
STEP 4: VIEW YOUR RESULTS
------------------------------------------------------------
1. Open "companies_tracker.xlsx" in the same folder.
   (It is created after the first successful run.)

2. The spreadsheet has the following sheets (tabs):

   INDUSTRY SHEETS (one per category):
   - Holding Companies
   - Consultancy
   - IT Companies
   - AI Companies
   - Manufacturing
   - Travel
   - HealthTech & Life Sciences
   - Fintech & Financial Services
   - Energy & Climate
   - Advanced Mfg & Deep Tech
   - Data & Cloud Infra

   SUMMARY SHEETS:
   - All Companies     (every company across all categories)
   - Director Cross-Reference  (directors linked to 2+ companies)

3. Each row is one company. Columns are:
   A: Company Name
   B: Company Number
   C: Incorporation Date
   D: Registered Address
   E: SIC Code(s)
   F: Directors
   G: Beneficial Owners
   H: Linked Companies (same director)

NOTE: A company may appear on more than one industry sheet if
its SIC code falls under multiple categories. This is intentional.

TIP: You can have the Excel file open while the tracker is
running. Close and reopen it every 30 minutes to see new rows.


------------------------------------------------------------
STOPPING THE TRACKER
------------------------------------------------------------
Click the black console window and press Ctrl+C to stop.
The tracker will stop gracefully and save its progress.

To restart it, simply double-click run_tracker.bat again.
It will only process companies it has not seen before —
no duplicates will be added.


------------------------------------------------------------
DAILY / ONGOING USE
------------------------------------------------------------
- Run the tracker during business hours (or leave it running
  all day). Companies House registers new companies throughout
  the day, so longer run times mean more companies captured.

- The Excel file grows over time as new companies are added.
  It is fully rebuilt on each cycle so the data is always
  up to date and sorted by incorporation date (newest first).

- The tracker remembers every company it has already processed.
  Restarting it will not create duplicate rows.


------------------------------------------------------------
CUSTOMISING THE SIC CODES
------------------------------------------------------------
To add or remove SIC codes from any category:

1. Open "config.json" with Notepad (right-click → Open with
   → Notepad).

2. Find the category you want to edit under "sic_categories".

3. Add or remove codes from the list. Each code must be in
   quotes and separated by commas. For example:
   "IT Companies": ["62012", "62020", "62030"]

4. Save the file and restart the tracker.

To change how far back the first run searches (default 7 days):
   Change "initial_lookback_days" to any number of days.

To change the poll interval (default 30 minutes):
   Change "poll_interval_minutes" to any number of minutes.


------------------------------------------------------------
TROUBLESHOOTING
------------------------------------------------------------
PROBLEM: Double-clicking run_tracker.bat opens and closes
         immediately.
FIX:     Python is not installed or not on PATH.
         Repeat Step 1, making sure to tick "Add Python to PATH".

PROBLEM: The console shows "Authentication failed (401)".
FIX:     The API key in config.json is incorrect.
         Open config.json in Notepad and check the "api_key" value.

PROBLEM: The console shows "Rate limited (429)".
FIX:     This is normal. The script pauses for 65 seconds and
         retries automatically. No action needed.

PROBLEM: No new companies appearing after several runs.
FIX:     There may genuinely be no new incorporations in your
         SIC categories today. This is uncommon — check tracker.log
         to confirm the API is returning results.

PROBLEM: The Excel file is not updating.
FIX:     The script only saves a new Excel file when new companies
         are found. If no new companies have been incorporated since
         the last run, the file stays the same.

PROBLEM: The script appears to have stalled (no new console output
         for more than 2 minutes).
FIX:     1. Check tracker.log in the same folder — open it in
            Notepad to see the last action and any errors.
         2. Press Ctrl+C to stop the tracker.
         3. Double-click run_tracker.bat to restart it.
            It will resume from where it left off.

PROBLEM: "ModuleNotFoundError" appears in the console.
FIX:     The dependencies failed to install. Open Command Prompt,
         navigate to the tracker folder, and run:
         pip install -r requirements.txt


------------------------------------------------------------
FILES IN THIS FOLDER
------------------------------------------------------------
companies_house_tracker.py   Main script (do not edit)
config.json                  Settings and SIC codes (editable)
requirements.txt             Software dependencies (do not edit)
run_tracker.bat              Double-click to start the tracker
README.txt                   This guide
tracker.log                  Log file (auto-created on first run)
tracker_state.json           Tracks which companies have been seen
                             (auto-created, do not delete)
data_store.json              All fetched company data
                             (auto-created, do not delete)
companies_tracker.xlsx       Your Excel results
                             (auto-created on first run)

------------------------------------------------------------
SUPPORT
------------------------------------------------------------
If you encounter an issue not listed above, send the contents
of "tracker.log" to your developer for diagnosis.

============================================================
