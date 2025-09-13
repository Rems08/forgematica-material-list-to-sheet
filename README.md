# forgematica-material-list-to-sheet

# 🧪 How to use

```bash
# Basic usage (auto-detect columns and delimiter)
python forgematica_to_sheets.py --csv path/to/materials.csv

# Specify the output filename and default stack size
python forgematica_to_sheets.py --csv materials.csv --out my_sheet.xlsx --default-stack-size 64

# Override column names if your CSV headers are different
python forgematica_to_sheets.py \
  --csv materials.csv \
  --name-col "Item Name" \
  --total-col "Qty Required" \
  --missing-col "Missing Qty" \
  --available-col "Have"
```

Then upload the generated `.xlsx` to Google Drive and **open with Google Sheets** — all formulas are preserved.

# 📊 Columns & formulas (Sheets-side)

Both **TOTALS_ALL** and **MISSING_ONLY** include:

- **Materials**
- **Total (units)**

  - On **MISSING_ONLY**, the effective total is computed as:
    `=MAX(0, Missing + User units + User stacks × Stack Size)`

- **Stack Size**

  - On **MISSING_ONLY** this is:
    `=IFERROR(VLOOKUP(Materials, REFS!A:B, 2, FALSE), <default>)` (defaults to 64 unless changed)
  - You can tailor stack size per item (e.g., Ender Pearls 16, tools/armor 1)

- **# Stacks (ceil)** = `CEILING(Total/StackSize, 1)` (round up full stacks) — official Sheets function. ([Assistance Google][1])
- **# Double Chests** = `IF(Total=0, 0, CEILING(Stacks/54, 1))` — **54 slots** per double chest per Minecraft Wiki. ([Minecraft Wiki][2])
- **Stacks after last double** = `MOD(Stacks, 54)` — official Sheets MOD. ([Assistance Google][3])
- **Units after last stack** = `MOD(Total, StackSize)` — official Sheets MOD. ([Assistance Google][3])

On **MISSING_ONLY** you also get:

- **User units (editable)**
- **User stacks (editable)**
- **Computed Total (units)** (formula above)

The script adds **data validation** so editable numeric cells don’t go negative.

# 🔧 Under the hood (why it’s generic)

- **Delimiter auto-detection** (`,`, `;`, tab, `|`) with override flag `--delimiter`.
- **Fuzzy header detection** for common words (e.g., `name/item/material`, `total/required/amount`, `missing/needed`, `available/have`). You can override with `--name-col`, `--total-col`, etc.
- **Grouping by material** in Python (just to consolidate rows); **all derived math** (stacks, chests, remainders) is done via **Google Sheets functions**:

  - `CEILING` / `CEILING.MATH` for rounded-up counts in Sheets. ([Assistance Google][1])
  - `MOD` for remainders. ([Assistance Google][3])
  - `VLOOKUP` for per-item stack sizes from the REFS sheet. ([Assistance Google][4])

- **Double chest capacity** fixed at 54 slots (Minecraft Wiki). ([Minecraft Wiki][2])

# 💡 Tips

- Extend the **REFS** sheet with your own `Materials → Stack Size` pairs; the `VLOOKUP` picks them up automatically. (Docs: VLOOKUP. ([Assistance Google][4]))
- Prefer `CEILING`/`CEILING.MATH` for “full stacks” rounding; see Sheets docs for nuances. ([Assistance Google][1])
- `MOD` is ideal for “remaining stacks/units after last chest/stack”. ([Assistance Google][3])

If you’d like, I can also:

- Add an optional **third sheet** that reconciles `Available + User inputs` versus `Totals` to highlight what’s still missing.
- Switch per-row formulas to **`ARRAYFORMULA`** per column if you prefer a single formula at the top (common in large sheets).
