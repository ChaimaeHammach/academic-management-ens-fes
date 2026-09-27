# Academic Management – ENS Fès

Desktop applications that automate the academic deliberation process at the **École Normale Supérieure de Fès** (ENS Fès): from the grade sheets exported by the school's grading system to the final deliberation reports and the preparation of make-up exams.

Developed during a one-month internship (July 2026) at the Student Affairs Office of ENS Fès.

---

## Applications

### PV Generator — `pv_generator/`

Generates the deliberation reports (*procès-verbaux*, PV) of the Licences d'Éducation.

- Reads grade sheets in **French and Arabic**, including multi-level headers
- Merges all module sheets of a semester, using the **Massar code** as the student identifier
- Detects **repeating students** and retrieves their grades from the previous year's report
- Produces semester, annual and degree-level reports
- **Dynamic reports**: every result is an Excel formula linked to the original sheet — if the jury changes a grade during the deliberation, averages and decisions are recalculated automatically

### Rattrapage Manager — `rattrapage_generator/`

Prepares the make-up exam session.

- Detects students who failed each module
- Generates make-up grade sheets in the original format, restricted to the students concerned
- Builds a **conflict matrix** (students shared by each pair of modules)
- Suggests an exam schedule using **greedy graph colouring**, so that no student has two exams in the same time slot

---

## Tech stack

| Purpose | Library |
|---|---|
| Graphical interface | PySide6 |
| Excel reading/writing (with formulas) | openpyxl |
| Data processing | pandas *(PV Generator)* |
| PDF export | reportlab *(PV Generator)* |
| Windows executable | PyInstaller |

---

## Installation

Requires **Python 3.10 or later**.

```bash
cd pv_generator            # or rattrapage_generator
python -m venv venv
venv\Scripts\activate      # Windows
pip install -r requirements.txt
python app.py
```

To build a standalone Windows executable, run `build_exe.bat` in the application folder. The `.exe` is created in `dist/`.

---

## Project structure

```
academic-management-ens-fes/
├── pv_generator/
│   ├── app.py                  # Graphical interface
│   ├── filieres.json           # School name and list of programmes
│   └── modules/
│       ├── reader.py           # Grade sheet reading
│       ├── copie_fiche.py      # Sheet copy with formulas
│       ├── merger.py           # Merge by student identifier
│       ├── calculator.py       # Averages and decisions
│       ├── generateur.py       # Report generation (Excel formulas)
│       ├── inject_redoublants.py
│       └── pv_global_licence.py
└── rattrapage_generator/
    ├── app.py                  # Graphical interface
    └── core.py                 # Reading, make-up sheets, conflicts, schedule
```

---

## Data privacy

This repository contains **no student data**. Grade sheets and generated reports are excluded by `.gitignore`.

---

## Author

**Chaimae Hammach** — Engineering student, ENSA Fès (ILIA)

Internship supervised by M. Ali Ahaitouf, ENS Fès.
