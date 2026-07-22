# TCG (Test Case Generator)

A localized ETL (Extract, Transform, Load) application built for the USAF to automate the generation of simulation test cases. 

**Note:** *Some variable names and content have been sanitized to avoid sharing classified information. Example data cannot be provided for the same reason.*

## The Challenge
This project was developed within a highly restricted, air-gapped computing environment. The primary engineering challenge was building a robust data processing pipeline with **zero access to external package managers**, necessitating highly optimized, memory-efficient use of standard and pre-approved Python libraries (like `pandas` and `xlwings`) to handle messy, legacy spreadsheet data.

## Technical Architecture
While the frontend is a lightweight Tkinter GUI, the backend operates as a dedicated ETL pipeline:
* **Extract:** `TCG_generate.py` parses unstructured `.xlsx`, `.xls`, and `.scn` files, handling missing data and dynamically locating headers across disparate legacy formats.
* **Transform:** `pandas` and `numpy` are utilized to clean, filter, and reshape the extracted identifiers and simulation modes into standardized DataFrames.
* **Load:** `TCG_format.py` acts as the loading mechanism, utilizing `xlwings` context managers to interact with the COM interface. It rapidly injects the transformed data into a highly formatted, color-coded output template with auto-generated statistics.

## Project Scope & Inputs
The application requires two primary inputs:
1. **Test System List:** An `.xlsx` file containing coded identifiers and nomenclature for the target systems.
2. **Simulation Summary Files:** Unstructured `.xlsx` / `.xls` files summarizing the simulations to be tested.

**Output:** A consolidated Excel workbook containing two sheets per target system: a fully formatted test case outlining available simulation modes with dedicated QA spaces, and a summarized reference sheet.

## Setup & Installation
Since this was built for a restricted environment, dependencies are kept to a minimum. 

**Requirements:**
* Python 3.8+
* `pandas`
* `numpy`
* `xlwings`
* `openpyxl`
* `Pillow` (PIL)

**Run the application:**
```bash
git clone [https://github.com/yourusername/TCG.git](https://github.com/yourusername/TCG.git)
cd TCG
python main.py
