# SCGPC Fire Resistance Lab Portal

A mobile-first lab data portal based on `SCGPC_Fire_Resistance_SCBA.xlsx`.

## Included
- Team login with hashed passwords
- Admin panel for users and 28-day baseline strengths
- 144 specimen slots: 12 mixes × 4 temperatures × 3 replicates
- Phone-friendly data-entry cards
- Automatic mass-loss %, residual compressive strength (100 mm cube), and residual-strength %
- Live dashboard and results tables
- Automatic line/bar charts in the browser
- One-click Excel export using your uploaded workbook as the template
- SQLite database for simple local/deployment use

## Run locally
1. Install Python 3.11+.
2. In this folder run:
   `python -m venv .venv`
3. Activate it and run:
   `pip install -r requirements.txt`
4. Set a strong admin password (recommended):
   - Windows PowerShell: `$env:ADMIN_USERNAME="admin"; $env:ADMIN_PASSWORD="your-strong-password"`
   - Linux/macOS: `export ADMIN_USERNAME=admin; export ADMIN_PASSWORD='your-strong-password'`
5. Start:
   `python app.py`
6. Open `http://localhost:5000` on the computer.
7. For phone access on the same Wi-Fi, find the computer's local IP and open `http://<PC-IP>:5000` on the phone.

## Deployment
For public deployment, use a host with persistent storage or replace SQLite with PostgreSQL. Set:
- `SECRET_KEY`
- `ADMIN_USERNAME`
- `ADMIN_PASSWORD`

Use HTTPS in production. Do not use the example password in a public deployment.

## Data model
Mixes are M0–M11 with SCBA 0, 5, 10, 15, 20, 25% under 4M and 6M NaOH. Fire temperatures are 200, 400, 600 and 800°C. Each condition has 3 replicates.

The Excel export fills the existing workbook's raw-data sheet and baseline-strength cells. Excel formulas/charts in the template remain available when the exported workbook is opened in Excel.
