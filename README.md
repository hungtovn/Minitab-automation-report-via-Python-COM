# 📊 Minitab Automation via Python COM

Automate **Minitab desktop** using **Python and Windows COM interface**.  
This project demonstrates how to control Minitab programmatically — run statistical commands, inject data, perform Gage R&R studies, and export results — all from Python scripts.

---

## 🔧 Features

- Launch and control Minitab via `win32com.client`
- Send Minitab commands (`Let`, `Stats`, `Regress`, `GageRR`, etc.)
- Inject custom data into worksheets
- Rename and organize columns and sheets
- Automate common quality tools like **Gage R&R (Crossed)** studies
- Save and export projects or session logs

---

## 🧪 Example Use Cases

- Quality/Test Engineers automating repetitive analysis
- Lab measurement reports (Gage R&R, Attribute Agreement, ANOVA)
- Integrating Minitab with Python-based test frameworks or MES systems

---

## 📁 Sample Scripts

| File | Description |
|------|-------------|
| `basic_command.py` | Run simple Minitab commands |
| `gage_rr.py` | Automate a full Gage R&R (Crossed) study |
| `inject_data.py` | Write Python lists directly into Minitab columns |

---

## 💡 Requirements

- ✅ Windows OS  
- ✅ Minitab Desktop (v19, v20, v21 or later)  
- ✅ Python 3.x  
- ✅ `pywin32` (`pip install pywin32`)

---

## 🔗 Related Topics

- Python COM Automation
- Statistical Quality Control (SQC)
- Minitab Macros & Command Line Integration

---

## 📬 Feedback / Contributions

Feel free to open issues or submit PRs.  
This project is open to collaboration, especially from test engineers or quality professionals automating Minitab workflows.

