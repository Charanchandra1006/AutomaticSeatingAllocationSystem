# 🪑 Automatic Seating Allocation System

A desktop GUI application built with **Python & Tkinter** that automates exam hall seating arrangements for colleges. It reads student data from Excel files, intelligently distributes students across rooms using an odd/even branch interleaving strategy, and exports the final seating plan as a formatted Excel workbook.

---

## 📋 Table of Contents

- [Features](#features)
- [Tech Stack](#tech-stack)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [How to Run](#how-to-run)
- [Usage Guide](#usage-guide)
- [Excel Input Format](#excel-input-format)
- [Seating Logic](#seating-logic)
- [Exporting the Seating Plan](#exporting-the-seating-plan)
- [Building as a Standalone Executable](#building-as-a-standalone-executable)
- [Project Structure](#project-structure)
- [Known Limitations](#known-limitations)

---

## ✨ Features

- 🏫 **College Details Configuration** — Set college name, exam type (Mid1 / Mid2 / Semester), exam date, and exam time.
- 🚪 **Flexible Room Management** — Add multiple exam halls with custom row × column seat configurations. Delete rooms as needed.
- 📂 **Multi-Branch Excel Upload** — Upload separate Excel files per branch, specifying year and whether the branch belongs to the **Odd** or **Even** semester group.
- 🔀 **Smart Seating Allocation** — Automatically interleaves Odd and Even branch students column-by-column to ensure no two students of the same branch sit adjacent to each other.
- 🔍 **Find Your Room** — Students can search their allocated room and seat by entering their Student PIN.
- 📊 **Live Seating Plan Preview** — Displays a scrollable popup with the complete room-wise seating grid, along with total allotted seats and an invigilator signature field.
- 📤 **Export to Excel** — Generates a polished `.xlsx` workbook with one sheet per room, including college info headers, seating grid, and attendance summary rows with auto-adjusted column widths.
- 🖥️ **Standalone `.exe` Support** — Can be compiled into a single executable using PyInstaller (no Python installation required on the target machine).

---

## 🛠️ Tech Stack

| Technology | Purpose |
|---|---|
| Python 3.x | Core programming language |
| Tkinter + ttk | GUI framework |
| Pandas | Reading student Excel files |
| OpenPyXL | Generating formatted Excel exports |
| PyInstaller | Packaging as a standalone executable |

---

## ✅ Prerequisites

- Python **3.8+**
- pip

Install required packages:

```bash
pip install pandas openpyxl
```

> **Note:** `tkinter` is included with the standard Python installation on Windows. If it is missing, reinstall Python and make sure the *tcl/tk* option is checked during setup.

---

## 🚀 Installation

```bash
# 1. Clone or download the repository
git clone https://github.com/your-username/AutomaticSeatingAllocationSystem.git

# 2. Navigate into the project directory
cd AutomaticSeatingAllocationSystem

# 3. (Optional) Create and activate a virtual environment
python -m venv .venv
.venv\Scripts\activate   # Windows

# 4. Install dependencies
pip install pandas openpyxl
```

---

## ▶️ How to Run

```bash
python seat.py
```

This opens the GUI application.

---

## 📖 Usage Guide

The application is organized into **six tabs**. Work through them in order for a smooth experience.

### Tab 1 — College Details
Fill in:
- **College Name** (e.g., *JNTU Hyderabad*)
- **Exam Type** — select from `Mid1`, `Mid2`, or `Semester`
- **Exam Date** — enter in `DD/MM/YYYY` format
- **Exam Time** — enter the slot (e.g., `10:00 AM - 1:00 PM`)

Click **Submit** to save, then **Next** to continue.

---

### Tab 2 — Room Details
Add each exam hall:
- **Room Name** (e.g., `Room 101`)
- **Rows** — number of rows of seats
- **Columns** — number of columns of seats

Click **Add Room** to save. Repeat for every hall. Use **Delete Room** to remove the last added room. Click **Next** when done.

---

### Tab 3 — Excel Upload
For each branch:
1. Enter the **Branch** name (e.g., `CSE`, `EEE`)
2. Enter the **Year** (e.g., `2`, `3`)
3. Select **Odd** or **Even** (semester group)
4. Click **Upload File** and select the branch's `.xlsx` file

Repeat for all branches. Click **Next** when all files are uploaded.

---

### Tab 4 — Generate Seating
Click **Generate Seating Plan** to run the allocation algorithm. A scrollable popup will appear showing the complete, room-by-room seating arrangement.

---

### Tab 5 — Find Your Room
Students can enter their **Student PIN** to instantly look up their assigned room, row, and column number.

---

### Tab 6 — Export Seating Plan
Click **Export to Excel** to save the full seating plan as an `.xlsx` file. Each room gets its own sheet with:
- College & exam metadata at the top
- Full seating grid (Row 1, Row 2 … with PINs)
- Total allotted count, present, absent rows
- Invigilator signature field

---

## 📄 Excel Input Format

Each uploaded student file must be an `.xlsx` file containing **at minimum** the following column:

| StudentPIN |
|---|
| 220001 |
| 220002 |
| 220003 |
| … |

> The column header must be exactly `StudentPIN` (case-sensitive).

Sample files for different branches (CSE, EEE, CE, ME, EI) are available in the `studentsexcels/` directory.

---

## 🧠 Seating Logic

The allocation algorithm follows this strategy:

1. **Separate** uploaded students into two pools — **Odd** (odd-semester branches) and **Even** (even-semester branches).
2. For each room, **odd-numbered columns** (1st, 3rd, 5th…) are filled with Odd-branch students, and **even-numbered columns** (2nd, 4th, 6th…) are filled with Even-branch students. This ensures no two students from the same branch share adjacent seats.
3. If any students remain after all rooms are filled in the primary pass, a **dynamic adjustment** pass fills remaining `"Empty"` seats column-alternately.
4. If students still cannot be seated, a **warning** is displayed and the user is advised to add more rooms.

---

## 📦 Exporting the Seating Plan

The exported `.xlsx` has:
- One **worksheet per room**
- College metadata (name, exam type, date, time) at the top of each sheet
- A seating grid where each row is labeled `Row 1`, `Row 2`, etc., and each cell contains the student PIN or is blank
- Auto-adjusted column widths to fit content

---

## 🏗️ Building as a Standalone Executable

A `seat.spec` file is included for PyInstaller. To build:

```bash
pip install pyinstaller
pyinstaller seat.spec
```

The compiled executable will appear in the `dist/` folder as `seat.exe`. It bundles the Python runtime, so **no Python installation is needed** on the end user's machine.

> The app icon (`app_icon.ico`) is embedded automatically via the spec file.

---

## 📁 Project Structure

```
AutomaticSeatingAllocationSystem/
│
├── seat.py                  # Main application source code
├── seat.spec                # PyInstaller build spec
├── app_icon.ico             # Application icon
│
├── studentsexcels/          # Sample student data files
│   ├── ce_students.xlsx
│   ├── cse_students.xlsx
│   ├── eee_students.xlsx
│   ├── ei_students.xlsx
│   └── me_students.xlsx
│
├── build/                   # PyInstaller build artifacts (auto-generated)
├── dist/                    # Compiled executable output (auto-generated)
└── .venv/                   # Virtual environment (if used)
```

---

## ⚠️ Known Limitations

- Attendance tracking (Present / Absent counts) is currently a placeholder — it shows `0` in both the UI and the export. Future versions could add an attendance input step.
- The **Delete Room** button removes only the **last** added room.
- Student-to-room lookup relies on the seating plan being generated in the current session; it does not persist across application restarts.
- The application is designed for **Windows** (uses a `.ico` icon and is built with PyInstaller for Windows).

---

## 👤 Author

**Charan Chandra**  
B.Tech Student | Python Developer  

---

## 📜 License

This project is open-source and available for educational purposes.
"# AutomaticSeatingAllocationSystem" 
