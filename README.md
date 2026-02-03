# Unstructured-sales-data-cleaning & transformation-python
Data cleaning and Excel automation using Python (Pandas & OpenPyXL)

## Business Context
Drawing from my professional experience at a **tech startup** and **NITI Aayog**, I observed a recurring bottleneck: critical operational sales data often arrives in highly unstructured formats (merged cells, inconsistent dates-formats, missing values). This forces analysts to spend hours on manual cleanup in Excel & Google Sheets before any analysis can begin.

**My this project automates that ETL (Extract, Transform, Load) process.** It transforms raw, messy vouchers into a clean, analysis-ready format and applies automated formatting (colors, borders) to generate executive-ready reports instantly.

## Key Features
* **Data Ingestion:** Reads raw `.xlsx` operational files using `pandas`.
* **Data Cleaning:**
    * Identifies and handles **Missing Values (NaNs)**.
    * Standardizes column names (snake_case) for consistency.
    * Parses and formats Date/Time columns.
* **Excel Automation (OpenPyXL):**
    * Exports clean data to a new Excel file.
    * **Conditional Formatting:** Applies alternating row colors (called Zebra striping) for readability.
    * Auto-adjusts column widths based on content.

## Data Transformation & Business Insights
Beyond cleaning and formatting, this project demonstrates how structured data can be used to generate **actionable operational insights**.

### Aggregations & Metrics
Using Pandas, the cleaned dataset is grouped and analyzed to calculate:
- **Total Sales:** Aggregated amount per day/month.
- **Volume Analysis:** Operator-wise transaction volume.
- **Regional Performance:** State-wise performance comparison.
- **Transaction Quality:** Count of unique Txn IDs and average transaction value.
  ( I believe this mirrors real-world reporting dashboards used by operations & finance teams.)

### Feature Engineering
Derived columns are created to simulate SQL logic and ETL transformations:
- **Revenue Logic:** Revenue after tax & Discount calculations.
- **Categorization:** Classifying transactions (e.g., Recharge vs. Setup).
- **Time Analysis:** Weekday vs. Weekend transaction trends.

### Data Validation & Quality Checks
To ensure reporting accuracy, the script performs automated audits:
- Flags **duplicate Txn IDs**.
- Identifies **negative or zero-value** transactions.
- Validates date ranges and detects outliers in transaction amounts.

## Business Impact
By automating this workflow-
✔ **Efficiency:** Manual cleanup time reduces from **hours to minutes**.  
✔ **Accuracy:** Reporting becomes **standardized and error-free**.  
✔ **Focus:** Teams can focus on **analysis instead of formatting**.  
✔ **Scalability:** The logic mirrors **SQL CASE statements** and can scale into production ETL systems.

## Tech Stack
* **Python:** Core logic & scripting.
* **Pandas:** Data manipulation, aggregation & cleaning.
* **OpenPyXL:** Advanced Excel formatting (styling, borders, colors).

## Project Structure
| File | Description |
| :--- | :--- |
| `sales_analysis.py` | Main script for ETL and formatting logic. |
| `sample_input.xlsx` | Raw, unstructured sales data (for testing). |
| `styled_sales_voucher.xlsx` | Final output: Cleaned & formatted Excel report. |

