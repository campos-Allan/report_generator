# 📊 Excel Auto-Refresh & PDF Export Automation

This project automates opening multiple Excel workbooks (with heavy queries that auto-refresh on open), waits until all queries finish updating, exports specific worksheets to PDF, and closes the files cleanly — requiring no manual intervention.

It also includes an additional script that generates and exports a monthly styled graphic as a PNG image, which triggers a Power Automate flow to post everything daily automatically to Microsoft Teams.

---

## 🚀 Objective

Eliminate manual opening, monitoring, and exporting Excel files connected to external data sources by automating the entire refresh-to-export-to-Teams workflow.

---

## ⚙️ Features

- 🔄 Opens multiple Excel files simultaneously via `win32com.client`
- ⏳ Waits for queries to finish by monitoring Excel’s CPU usage with `psutil`
- 📂 Exports selected worksheets to daily PDFs, timestamped by current date
- 📈 Generates and exports a monthly styled graphic as PNG for reporting
- 🤖 Triggers a Power Automate flow when the graphic PNG file is updated
- 🛑 Closes Excel files cleanly without force killing processes
- ♻️ Retries failed file openings (up to 3 attempts)
- 🖱️ Prevents screen lock and sleep with simulated mouse movement (`ctypes`)
- ❌ Requires no macros or VBA code in the Excel files

---

## 📁 Project Structure
```
├── report_generator.py # Main automation script to refresh Excel files & export PDFs
├── graphic.py # Script to refresh query, generate graphic, export PNG, trigger Teams post
└── README.md # This documentation file
```

---

## 🧰 Requirements

- **Python 3.8+**
- **Windows OS** (due to Excel COM and user input simulation)
- **Excel** installed locally

### Python packages

- `pywin32`
- `psutil`
- `pyautogui`

Install with:

```bash
pip install pywin32 psutil pyautogui
```

---

## 🧠 How It Works

### report_generator.py

- Opens all Excel workbooks defined in the `arq` dictionary.
- Monitors Excel’s CPU usage to detect when Power Query updates finish.
- Exports specified worksheet indexes to PDF files with a date suffix.
- Closes all workbooks gracefully.
- Uses subtle mouse movements to prevent screen lock or sleep during wait times.
- Retries opening any failed files up to 3 times.

### graphic.py

- Opens a specific Excel file with the data query.
- Refreshes queries and waits for completion.
- Copies updated data into a monthly destination workbook.
- Copies a graphic from the workbook and pastes it into a new PowerPoint slide.
- Exports the slide as a PNG image file to a fixed path.
- This PNG overwrite triggers a Power Automate flow.
- The Power Automate flow posts everything automatically to a Teams channel.

---

## ✅ Usage

1. Update file paths and worksheet indexes in `report_generator.py` and `graphic.py`.
2. Run the main script:

`python report_generator.py`
- `report_generator.py` will refresh and export PDFs.

- It will call `graphic.py` to generate and export the monthly graphic PNG.

- The PNG file update triggers your Power Automate flow to post the graphic in Teams.

---

## ⚠️ Notes

- Excel files must be accessible on the local machine.

- Excel must be installed and functional.

- No VBA or macros are needed in your Excel files.

- Ensure Power Automate flow is configured to watch the PNG file path and post to Teams.

- Paths in the scripts should be absolute and user-specific.

---

## 📈 Benefits

- Saves 1 hour of manual Excel refresh and export work daily.

- Reliable and consistent refresh/export without human error.

- Automates visual report posting to Teams with zero manual steps.

- Enables faster decision making with up-to-date data and visuals.

---

## 👨‍💻 Author

Created for internal automation in corporate logistics and reporting teams to streamline Excel-based dashboards.
