# Forgematica Material-List to Google Sheets Converter

A Python script that reads a Forgematica (or similarly-structured) CSV file containing materials, totals, missing quantities, etc., and produces a Google Sheets-friendly Excel workbook with two sheets: one for total requirements and one adjusted for missing items plus user inventory inputs.

---

## 🚀 Features

- Automatically detects the relevant CSV columns (materials/item name, total required, missing, available) using fuzzy matching.  
- Generates two sheets in the output workbook:
  - **TOTALS_ALL**: shows total quantities needed per material.
  - **MISSING_ONLY**: shows what is missing + allows user to input how much they already have (units or stacks), and computes what remains.  
- All derived calculations (stack rounding, double chest calculation, remainders, etc.) are done via formulas in the spreadsheet (Google Sheets / Excel), not by the Python script.  
- Includes a **REFS** sheet with a lookup table for custom stack sizes and helpful documentation/links.  

---

## 📋 Requirements

- Python 3.10 or newer  
- Packages:  
  - `pandas`  
  - `openpyxl`  

You may install dependencies via:

```bash
pip install -r requirements.txt
````

(If you’re in a system-managed Python environment, using a virtual environment is recommended.)

---

## ⚙️ Usage

```bash
python3 forgematica_to_sheets.py --csv path/to/material_list.csv
```

Optional arguments:

| Flag                   | Description                                                                         |
| ---------------------- | ----------------------------------------------------------------------------------- |
| `--out`                | Path to save the output `.xlsx` file (default: `forgematica_materials_sheets.xlsx`) |
| `--delimiter`          | Override the CSV delimiter (auto-detected if omitted)                               |
| `--name-col`           | Column name in CSV for item/material names                                          |
| `--total-col`          | Column name for the total units required                                            |
| `--missing-col`        | Column name for how many are missing                                                |
| `--available-col`      | Column name for how many are available                                              |
| `--default-stack-size` | Default stack size to use if a material isn’t in the lookup table (default: 64)     |

---

## 🧮 How the Sheets Are Structured

The output Excel/Sheets file contains:

* **TOTALS\_ALL**: shows for each material:

  * Total units required
  * Stack size
  * Number of stacks needed (rounded up)
  * Number of double chests needed (rounded up)
  * Remainders after last double chest and after last stack

* **MISSING\_ONLY**: starting from the “Missing” quantities, lets you fill in:

  * “User units (you have)”
  * “User stacks (you have)”
  * Then it computes effective total units you still need
  * Same derived columns as above (stacks, chests, remainders)

* **REFS**: reference sheet with material → stack size mappings (editable), plus documentation links (about `CEILING`, `MOD`, etc.).

---

## 🛠 Tips & Customization

* You can edit the **Stack Size** per material via the REFS sheet; the script sets a default if no custom value is found.
* If you want, you can convert per-row formulas into `ARRAYFORMULA`s in Google Sheets to reduce repetition.
* If your CSV has unusual headers, use the `--name-col`, `--total-col` etc. flags to manually tell the script which columns to use.

---

## ⚠ Known Issues / Limitations

* Materials with stack sizes other than 64 must be added to the REFS lookup, otherwise they'll default (which may over- or under-estimate).
* Very large CSVs may lead to large spreadsheets; Google Sheets may slow down for many thousands of rows with formulas.
* The script sums duplicate material names, but exact matching is case-sensitive after normalization; very slightly different names may produce separate lines.

---

## 📂 Project Structure

```text
/
├─ forgematica_to_sheets.py        # main script
├─ requirements.txt                # Python dependencies
├─ README.md                       # this file
├─ sample_csv/                     # (optional) example CSV files
│   └─ example_material_list.csv
└─ output/                         # (optional) folder for generated .xlsx outputs
```

---

## ✍ Contributing

Contributions welcome! If you see bugs, want new features (e.g. inventory reconciliation, automated stack size suggestions, better CSV format support), feel free to:

1. Fork the repository
2. Create a feature branch (e.g. `feature/my-addition`)
3. Make your changes; include tests/examples if possible
4. Submit a pull request