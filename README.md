# Excel Data Cleaning - U.S. Presidents Dataset

## 📌 Project Overview
This project demonstrates a practical, step-by-step data cleaning workflow using **Microsoft Excel** on a real-world dataset of U.S. Presidents (sourced from Kaggle, with intentional errors added for demonstration purposes).

While advanced analysis often happens in SQL, Python, or Tableau, Excel remains one of the fastest and most accessible tools for bulk-cleaning small-to-medium datasets before they're passed downstream. This project walks through identifying common data quality issues and resolving them using built-in Excel functions and features — no macros or scripting required.

---

## 🧾 Common Data Issues Identified
- Inconsistent text casing (all-caps, all-lowercase, title case mixed together)
- Categorical misspellings and inconsistent labels (e.g., `Republicans` vs `Republican`)
- Unwanted leading, trailing, or embedded whitespace
- Currency symbols interfering with numeric data types
- Inconsistent date formats across records
- Duplicate rows
- Unusable or irrelevant columns

---

## 🛠️ Tools & Excel Functions Used
- **Remove Duplicates** (Data tab)
- **Text Functions:** `PROPER()`, `UPPER()`, `LOWER()`, `TRIM()`
- **Filters** (to inspect and correct inconsistent categorical values)
- **Number Formatting** (converting currency-formatted text to clean numeric values)
- **Date Formatting** (standardizing to Short Date format)
- **Paste Special → Values** (to lock in formula results before deleting helper columns)

---

## 🧹 Data Cleaning Process & Methodology

### 1. Removing Duplicates
- Selected the full dataset and used **Data → Remove Duplicates**, keeping all columns checked to catch exact full-row duplicates.
- Result: one duplicate record (Barack Obama) was identified and removed.

### 2. Fixing Text Formatting & Casing
- Added a helper column next to the target field (e.g., `President - Fixed`).
- Applied `=PROPER(cell)` to standardize name casing, filled down the column.
- `UPPER()` and `LOWER()` used similarly where needed for other fields.

### 3. Handling Unusable / Garbage Data
- Inspected columns like `Prior` containing messy, non-parseable values.
- Removed columns with no clear use case to streamline the dataset.

### 4. Cleaning Inconsistent Categorical Data & Typos
- Enabled **Filters** on the header row.
- Reviewed unique values in categorical columns (e.g., `Party`) via the filter dropdown.
- Identified inconsistent labels (e.g., `Republicans` vs `Republican`, or `Whig` with trailing text).
- Filtered and manually corrected each variant to ensure consistency for future grouping and pivot table operations.

### 5. Removing Unseen Spaces
- Added a helper column and applied `=TRIM(cell)` to strip invisible leading/trailing/internal extra spaces.
- Critical for downstream compatibility — trailing spaces are invisible but break database imports (e.g., SQL) and string-matching logic.

### 6. Formatting Currency & Numeric Data
- Selected the currency-formatted column and changed its format from **Currency** to **Number/General**.
- Removed unnecessary decimal places.
- Necessary because special characters like `$` force Excel to treat values as text, blocking sums and numeric operations.

### 7. Standardizing Dates
- Checked for mixed text/number date entries using column filters.
- Set the date column format to **Short Date** for uniformity.

### 8. Replacing Formulas with Static Values
- Selected the cleaned helper column, copied it, and used **Paste Special → Values** to freeze the results as static text/numbers.
- Deleted the now-redundant original column and helper column.

### 9. Final Column Pruning & Review
- Removed unnecessary columns (e.g., `Prior`) and redundant index columns.
- Retained useful audit columns like `Date Created` / `Date Updated`.
- Performed a final visual pass to confirm the dataset was fully standardized and ready for analysis or export.
