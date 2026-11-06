# Automated ETL & Interactive Dashboard System

A Python and Streamlit application that brings common data analyst workflow steps into one interactive interface. Upload a dataset, inspect its structure, review and clean data-quality issues, analyze the data, build visualizations, explore an interactive dashboard, and export an analyst-style report.

---

## Project Overview

Data analysis projects often involve several repeated steps across data preparation, validation, exploration, visualization, and reporting. This project combines those steps into a guided workflow so users can work with different tabular datasets from a single application.

The application detects a broad dataset domain from column names and uses the available columns to guide KPI generation, analysis, and visualizations. Domain detection and automated observations are heuristic, so results should always be reviewed before being used in a business presentation.

## Project Workflow

![Automated ETL Project Workflow](Assets/Images/Automated ETL Dashboard Workflow.png)

## Features

### 1. Dataset Upload
- Upload CSV, Excel (`.xlsx` / `.xls`), and JSON files.
- Preview the uploaded dataset.
- Retain the raw dataset separately from the cleaned working dataset during the app session.

### 2. Data Overview
- Review row and column counts.
- Inspect column names and data types.
- View dataset previews and statistical summaries.

### 3. Data Cleaning
- Review column-wise cleaning recommendations.
- View suggested primary-key columns.
- Apply supported date and numeric type conversions.
- Handle missing values using the available cleaning actions.
- Review duplicate rows and possible identifier duplicates.
- Inspect potential outliers rather than automatically treating every unusual value as an error.
- Review the cleaning log and download the cleaned dataset.

### 4. Data Quality Validation
- Review column-level quality information.
- Check missing values and duplicate records.
- Inspect possible duplicate identifiers.
- Download the quality report when available.

### 5. Explore & Analyze
- Review numeric summaries.
- Explore category-level summaries.
- Inspect correlations between numeric fields.
- Perform supported custom aggregations.

### 6. Data Visualization
Choose from chart types supported by the application, including:
- Bar and column charts
- Line charts
- Pie and donut charts
- Scatter plots
- Histograms
- Box plots
- Treemaps
- Heatmaps
- Gauge charts

The options available and the usefulness of each chart depend on the columns and data types in the uploaded dataset.

### 7. Interactive Dashboard
- Generate domain-aware KPI cards from available fields.
- Display up to six KPI cards at a time.
- Use up to five dataset filters where suitable columns are available.
- Explore dashboard visualizations and filtered records.
- Download the filtered dataset.

### 8. Findings & Recommendations
- Generate observations from the current dataset.
- Provide recommendations linked to identified patterns or data-quality concerns.
- Display the detected domain.
- Review automated observations before treating them as final business conclusions.

### 9. Export Reports
- Generate PDF reports using ReportLab when installed.
- Generate PowerPoint presentations using `python-pptx` when installed.
- Include dataset overview, cleaning activity, data-quality summary, KPIs, findings, recommendations, and charts where chart-image export is supported.

---

## Technology Stack

| Technology / Library | Role |
|---|---|
| Python | Core application language |
| Streamlit | Interactive web application |
| Pandas | Dataset loading, transformation, and analysis |
| NumPy | Numerical operations |
| Plotly Express | Interactive data visualizations |
| Plotly Graph Objects | Custom interactive visualizations |
| ReportLab | PDF report generation |
| python-pptx | PowerPoint report generation |

The application uses Python's built-in modules such as `io`, `re`, `warnings`, and `datetime` for supporting tasks.

## Project Structure

```text
Automated-ETL-Dashboard-System/
│
├── app.py
├── README.md
├── requirements.txt
├── .gitignore
│
├── data/
│   ├── raw/
│   └── processed/
│
├── reports/
│   ├── pdf/
│   └── ppt/
│
├── assets/
│   ├── screenshots/
│   └── images/
│
├── docs/
│   └── project_documentation.md
│
└── venv/                  # Local virtual environment; do not upload
```

Generated files and folders may be created as you use the application. Keep personal, confidential, or restricted datasets out of a public GitHub repository.

## Installation and Setup

### Prerequisites
- Python installed on your computer
- pip (Python package installer)
- A terminal such as PowerShell or Command Prompt

### 1. Open the project folder

```powershell
cd "path\to\Automated-ETL-Dashboard-System"
```

### 2. Create a virtual environment

```powershell
python -m venv venv
```

### 3. Activate the virtual environment

For Windows PowerShell:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\venv\Scripts\Activate.ps1
```

If PowerShell activation is unavailable, use Command Prompt:

```cmd
venv\Scripts\activate.bat
```

### 4. Install dependencies

```powershell
python -m pip install --upgrade pip
pip install -r requirements.txt
```

### 5. Run the application

```powershell
python -m streamlit run app.py
```

Open the local URL displayed in the terminal (commonly `http://localhost:8501`) in your browser.

## Requirements

The `requirements.txt` file lists the Python packages needed by the application. Install them from the project root using:

```powershell
pip install -r requirements.txt
```

## How to Use

1. Start the Streamlit application.
2. Open **Upload Data** and upload a CSV, Excel, or JSON dataset.
3. Review **Data Overview** to understand the dataset.
4. Open **Data Cleaning** and inspect the suggested actions before applying them.
5. Use **Data Quality Validation** to check the resulting data.
6. Explore **Explore & Analyze** for summaries and custom analysis.
7. Use **Visualize Data** to create charts.
8. Open **Interactive Dashboard** and apply filters to investigate the data.
9. Review **Findings & Recommendations**.
10. Export a PDF or PowerPoint report from **Export Reports**.

## Important Notes and Limitations

- The application analyzes the columns and values in the uploaded dataset. It cannot guarantee that automatically detected domains, suggested identifiers, KPIs, or findings match every business context.
- A suggested primary key must be validated against the real-world meaning and uniqueness requirements of the data.
- Missing-value treatments should be selected based on the meaning of each column; filling or removing values can affect results.
- Potential outliers are review candidates. An unusual value is not automatically an error.
- Some visualizations require particular combinations of categorical, numeric, date, or geographic fields.
- PDF and PowerPoint export features require their respective libraries to be installed.
- Chart images in exported reports may depend on Plotly image-rendering support in the local environment.
- The application session is not a substitute for a permanent database or a versioned data pipeline. Save important source and processed files separately.

## Screenshots

Add screenshots of your own running application to `assets/screenshots/`, then embed them here. For example:

```markdown
![Application dashboard](assets/screenshots/dashboard.png)
```

Suggested screenshots:
- Dataset overview
- Data cleaning and quality validation
- Visualization page
- Interactive dashboard
- Exported report preview

## Future Enhancements

Possible future improvements:
- More configurable cleaning rules
- Additional file formats and larger-dataset handling
- Reusable data-cleaning pipelines
- More domain-specific KPI definitions
- Automated unit and data-validation tests
- Deployment with a hosted Streamlit environment

## Author

- **Karthik Kasakani**
- **https://github.com/KarthikKasakani2004**
- **https://www.linkedin.com/in/karthik-kasakani**
