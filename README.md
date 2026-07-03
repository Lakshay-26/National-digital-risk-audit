# National Digital Risk Audit 🛡️

A comprehensive Python-driven engine that evaluates the digital security posture of local businesses and automatically generates professional client deliverables, including an interactive web dashboard, detailed Excel database, and presentation-ready PowerPoint deck.

## Features ✨

* **Interactive Web Dashboard:** A self-contained HTML/JS dashboard featuring Light/Dark mode toggles, instant multi-parameter sorting (Risk, Score, Popularity), category-based donut charts, and an embedded Posture Calculator drawer.
* **Automated Pitch Engine:** Generates highly tailored, ready-to-copy email templates for client outreach based on specific security vulnerabilities identified during the audit.
* **Intelligent Data Generation:** The Python engine automatically formats the dataset, handles capitalization, formats `csv` outputs beautifully with `csv.QUOTE_ALL`, and generates premium charts.
* **Full Suite Deliverables:** Running a single script outputs:
  * 📊 **Excel Workbook** (`.xlsx`) with conditional formatting
  * 📝 **CSV Database** (`.csv`) with clean title-case headers
  * 🌐 **Interactive Dashboard** (`.html`)
  * 📑 **PowerPoint Deck** (`.pptx`)
  * 📉 **Matplotlib Charts** (`/charts/` folder)

## Getting Started 🚀

### Prerequisites
Make sure you have Python 3 installed. Install the required dependencies:
```bash
pip install pandas matplotlib openpyxl python-pptx
```

### Usage
To regenerate all deliverables and export fresh data, simply run the generator script:
```bash
python generate_national_risk_audit.py
```

All deliverables will be created and saved directly in your working directory. You can instantly view the dashboard by opening `National_Digital_Risk_Audit_Dashboard.html` in any modern web browser.

## Project Structure 📁
- `generate_national_risk_audit.py` — The core engine containing the dataset, HTML logic, and Excel/PPTX generation code.
- `National_Digital_Risk_Audit_Dashboard.html` — The generated interactive web app.
- `/charts/` — Contains dark-mode optimized charts used by the dashboard.
- `National_Digital_Risk_Audit_Briefing.md` — The markdown executive summary.
