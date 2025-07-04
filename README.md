# 📊 Excel Auto-Refresh & PDF Export Automation

This project automates the opening of multiple Excel workbooks (with heavy Power Query queries), waits until all queries have finished updating, exports specific worksheets to PDF, and closes the files cleanly — with no manual intervention required.

## 🚀 Objective

Eliminate the need to manually open, monitor, and close Excel files connected to external data sources by automating the entire refresh-to-export workflow.

## ⚙️ Features

- 🔄 Opens multiple Excel files simultaneously via `win32com.client`
- ⏳ Waits for queries to finish by monitoring Excel's CPU usage with `psutil`
- 📂 Exports selected worksheets to PDF, naming them with the current date
- 🛑 Closes files properly without using `taskkill`
- ♻️ Automatically retries failed file openings (up to 3 attempts)
- 🖱️ Prevents screen lock or monitor sleep by simulating subtle user activity with `ctypes`
- ❌ Does not modify the original files (no macros required)

## 📁 Project Structure

```
├── report_generator.py          # Main automation script
└── README.md                    # This file
```

## 🧰 Requirements

- **Python 3.8+**
- **Windows OS** (due to Excel COM interface and user input simulation)
- **Excel** installed locally

### Python Packages

- `pywin32`
- `psutil`
- `pyautogui`

### Installation

Install required packages via pip:

```bash
pip install pywin32 psutil pyautogui
```

## 🧠 Smart Wait Logic

Excel doesn't reliably expose when Power Queries are finished. After testing `.Refreshing`, `.Saved`, and `.CalculateUntilAsyncQueriesDone()` — all proved unreliable or unstable.

Instead, we use a more robust method: **monitoring the CPU usage of the EXCEL.EXE process**. When usage drops below a certain threshold (e.g., 10%) consistently for a few minutes, the script infers that queries have finished.

## ✅ How to Use

1. Edit the `arq` dictionary with your Excel file paths and worksheet indexes
2. Run the script with Python:
   ```bash
   python report_generator.py
   ```

The script will:
- Open all files
- Wait for query updates to finish
- Export selected worksheets to PDF
- Close all files automatically

PDFs will be saved to a specified output directory with timestamped filenames.

## ⚠️ Notes

- No VBA or macro code is required
- Files must be locally accessible 
- Ensure Excel is installed and functional on the machine

## 📈 Benefits

- **Reduced manual effort** from ~1 hour to ~15 minutes
- **Improved reliability** and consistency of report generation
- **Frees up time** for higher-value analysis by removing repetitive tasks
- **Error-proof**: no forgotten exports, missed files, or outdated data

## 👨‍💻 Author

Developed in a corporate environment to streamline Excel-based reporting processes, especially in logistics and operational dashboards.

## 📄 License

This project is developed for internal corporate use. Please ensure compliance with your organization's policies regarding automation tools.

---

**Note**: This automation tool is designed for Windows environments and requires Excel to be installed locally. For questions or issues, please refer to the project documentation or contact the development team. 
