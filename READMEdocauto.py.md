# 📊 KLM Report Automation Tool

This tool automates the generation of **monthly travel reports** (KM Maps) and **KPI tracking** for each employee based on Excel input data. It ensures consistent, formatted outputs in Excel and PDF, along with centralized performance metrics.

---

## ✅ Key Features

- Reads from a master Excel file (`Master.xlsx`) with data by company and employee.
- Uses a template file (`Template_Mapa_KM.xlsx`) to generate filled-in monthly reports.
- Exports:
  - Individual `.xlsx` files per company.
  - Individual PDF reports per employee.
- Aggregates all data into a `KPI.xlsx` with:
  - Monthly summaries.
  - Automatic chart generation.
- Fully automated:
  - No user interaction with Excel required (uses `win32com` automation).

---

## 📁 Suggested Folder Structure

YourProjectRoot/
│
├── Master.xlsx
├── Template_Mapa_KM.xlsx
├── klm.py / klm.exe
│
├── Mapas_Gerados/
│ └── [Month_Year]/
│
├── Mapas_Gerados/PDF/
│ └── [Month_Year]/
│
├── KPIs/
│ └── KPI.xlsx

yaml
Copy
Edit


---

## 🖥️ How to Run

**Using the `.py` script:**
```bash
python klm.py

Using the .exe (if packaged):

Place the .exe and Excel files in the same folder and double-click to run.

Requirements (for .py version)
pip install pandas openpyxl pywin32

🧠 Use Case
For organizations needing to:

Track mileage per employee monthly.

Automate travel reports and PDF generation.

Maintain centralized KPI tracking and reporting.


📅 Example Output for July 2025
Mapas_Gerados/July_2025/ → .xlsx files per company

Mapas_Gerados/PDF/July_2025/ → PDFs per employee

KPIs/KPI.xlsx → Updated with July 2025 data and charts





