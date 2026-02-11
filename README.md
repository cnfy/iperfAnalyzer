[![Download](https://img.shields.io/badge/Download-Latest%20Release-blue?style=for-the-badge)](https://github.com/cnfy/iperfAnalyzer/releases/latest)

# Iperf Analyzer

A desktop utility for parsing **iperf3** JSON results and generating formatted Excel reports.

> JSON in, Excel out.

---

## 📌 Overview

Iperf Analyzer is a lightweight GUI-based tool built with Python that converts **iperf3 JSON output files** into structured and formatted Excel reports.

It supports multi-stream tests and generates per-second throughput tables with connection metadata included.

The tool is designed for performance analysis, reporting, and test result archiving.

---

## ✨ Features

- Parse iperf3 JSON result files
- Convert per-second throughput data into structured Excel format
- Support multiple parallel streams
- Automatically extract and display connection metadata:
  - LocalHost
  - LocalPort
  - RemoteHost
  - RemotePort
- Throughput rounded to integer (bit/s)
- Timestamp converted to **UTC+9 (JST)**
- Clean Excel formatting:
  - Two-level headers
  - Center alignment
  - Adjustable column width
  - Freeze panes enabled
- GUI-based file selection (Tkinter)
- Batch processing of multiple JSON files
- Automatically generates timestamped result folders

---

## 📊 Output Format

Each Excel file contains:

### 1️⃣ Connection Information (Top Rows)

| Times(UTC+9) | Stream_4 | Stream_6 | ... |
|--------------|----------|----------|-----|
| LocalHost    | ...      | ...      |     |
| LocalPort    | ...      | ...      |     |
| RemoteHost   | ...      | ...      |     |
| RemotePort   | ...      | ...      |     |

---

### 2️⃣ Per-Second Throughput Data

| Times(UTC+9)       | Stream_4 | Stream_6 |
|--------------------|----------|----------|
| 2026-01-16 16:35:50 | 941234567 | 938765432 |
| 2026-01-16 16:35:51 | 945678123 | 940112233 |

- Unit: **bit/s**
- Time format: `YYYY-MM-DD HH:MM:SS`
- Timezone: UTC+9 (JST)

---

## 🕒 Time Handling

- Base time extracted from `start.timestamp.timesecs`
- Converted to UTC+9 using timezone-aware `datetime`
- Per-second offsets calculated using `timedelta`
- All timestamps are timezone-safe

---

## 🛠 Requirements

- Python 3.9+
- pandas
- openpyxl

Install dependencies:

```bash
pip install pandas openpyxl
```


## 🚀 How to Use

### Step 1 — Run the Application

```bash
python main.py
```

### Step 2 — Select JSON Files

Click:
选择文件

Select one or multiple iperf3 JSON result files.

### Step 3 — Start Analysis

Click:开始

### Step 4 — Choose Output Directory

The tool will:

- Create a folder named: 
IperfAnalyzer_Result_YYMMDDHHMMSS
- Save all generated Excel files inside it.
- Automatically open the result folder after completion.
## 📁 Project Structure
```
iperfAnalyzer/
│
├── main.py
├── README.md
└── LICENSE (MIT)
```

## 🔧 Build Standalone Executable (Optional)
Using PyInstaller:
```shell
 pyinstaller .\IperfAnalyzer.spec
```
Output will be located in:
```shell
dist/
```
## 📈 Typical Use Cases

- Network throughput benchmarking
- Multi-stream performance comparison
- Long-duration performance logging
- Test result documentation
- Performance regression tracking

## 📄 License
This project is licensed under the MIT License.

## 👤 Author
Developed as a lightweight analysis tool for iperf3 performance testing.
