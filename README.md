# Unstructured-sales-data-cleaning & transformation-python
Data cleaning and Excel automation using Python (Pandas & OpenPyXL)

## Business Context
Drawing from my professional experience at a **Tech startup** and **NITI Aayog**, I observed a recurring bottleneck: critical operational sales data often arrives in highly unstructured formats (merged cells, inconsistent dates-formats, missing values). This forces analysts to spend hours on manual cleanup in Excel & Google Sheets before any analysis can begin.

**My this project automates that Extract, Transform, Load (ETL) process.** It transforms raw, messy vouchers into a clean, analysis-ready format and applies automated formatting (colors, borders, graphs) to generate executive-ready reports instantly.

## Key Features
* **Data Ingestion:** Reads raw `.xlsx` operational files using `pandas`
* **Data Cleaning:**
    * Identifies and handles **Missing Values (NaNs)**
    * Standardizes column names (snake_case) for consistency
    * Parses and formats Date/Time columns
    * Check duplicates
* **Excel Automation (OpenPyXL):**
    * Exports clean data to a new Excel file
    * **Conditional Formatting:** Applies alternating row colors (called Zebra striping) for readability
    * Auto-adjusts column widths based on content
    * Make automated graphs to understand the trend

## Data Transformation & Business Insights
Beyond cleaning and formatting, this project demonstrates how structured data can be used to generate **actionable operational insights**.

### Aggregations & Metrics
Using Pandas, the cleaned dataset is grouped and analyzed to calculate:
- **Total Sales:** Aggregated amount per day/month
- **Volume Analysis:** Operator-wise transaction volume
- **Regional Performance:** State-wise performance comparison
- **Transaction Quality:** Count of unique Txn IDs and average transaction value
  ( I believe this mirrors real-world reporting dashboards used by operations & finance teams. And more can definately be done)

### Feature Engineering
Derived columns are created to simulate SQL logic and ETL transformations:
- **Revenue Logic:** Revenue after tax & Discount calculations
- **Categorization:** Classifying transactions (For e.g.: Recharge vs. Setup)
- **Time Analysis:** Weekday vs. Weekend transaction trends

### Data Validation & Quality Checks
To ensure reporting accuracy, the script performs automated audits:
- Flags **duplicate Txn IDs**
- Identifies **negative or zero value** transactions
- Validates date ranges & detects outliers in transaction amounts

## Business Impact
By automating this workflow-
**Efficiency:** Manual cleanup time reduces from **hours to minutes** 
**Accuracy:** Reporting becomes **standardized and error-free**
**Focus:** Teams can focus more on **analysis after quick formatting**  
**Scalability:** The logic mirrors **SQL CASE statements** and can scale into production ETL systems & more

## Tech Stack
* **Python:** Core logic & scripting
* **Pandas:** Data manipulation, aggregation & cleaning
* **OpenPyXL:** Advanced Excel formatting (styling, borders, colors, graphs)

## Project Structure
| File | Description |
| :--- | :--- |
| `sales_analysis.py` | Main Python script that performed data cleaning, KPI calculation, Excel report generation, Inserting graphs etc. |
| `sample_voucher.xlsx` | Synthetic raw, unstructured sales data (for testing) |
| `sales_report.xlsx` | Final output: Cleaned & formatted Excel report with summary and charts |
| `requirements.xlsx` | List of Python libraries required to run my project |

## Data Privacy Note
The original dataset used in this project belongs to a pvt org and I cann't be share it publicly.
To demonstrate the workflow, I used a **synthetic sample dataset** with a similar structure. The script logic remains the same and reflects real-world ETL and reporting processes.

`pip install -r requirements.txt`

`python sales_analysis.py`
