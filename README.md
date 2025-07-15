# VBA Data Processing Automation

![Screenshot of VBA Macro Interface](VBA_Master_Macro.png)

[![Excel VBA](https://img.shields.io/badge/Language-VBA-green.svg)](https://docs.microsoft.com/en-us/office/vba/api/overview/excel)

---

## Business Problem

During my Summer 2025 Metrology Retention internship at Entegris, our team identified a major bottleneck in the workflow for a particle‑analysis test, which measures the amount of particles and impurities in a sample. Each new request involved manually handling large CSV exports: copying data into master templates, recalculating metrics, rebuilding charts, and formatting reports. This manual process consumed **nearly 38 minutes per request**, delaying decision‑making and tying up skilled engineers in repetitive tasks.

### Manual Workflow (Pre‑Automation)

1. **Import Data**
   - Open raw CSV in Excel
   - Copy relevant columns into the test template
2. **Clean & Filter**
   - Remove blank rows and non‑printable characters
   - Manually apply filters to isolate the target sample batch
3. **Calculate Metrics**
   - Enter formulas to compute analyte‑to‑impurity ratios and percent purity
   - Flag any samples that fall below the required thresholds
4. **Build Charts**
   - Insert and style bar charts showing particle counts
   - Manually color‑code samples that fail specifications
5. **Finalize Report**
   - Adjust headers, column widths, and conditional formatting
   - Save and distribute

This repetitive process was error‑prone, inconsistent, and strained lab throughput for a critical quality metric.

## Automation Solution

I developed a modular VBA macro suite—all accessible from a single Master UserForm—that automates each step for the particle‑analysis test:

### Data Ingestion Module
- Batch‑imports CSVs for all requested sample runs into hidden staging sheets

### Cleaning & Transformation Module
- Automatically strips extraneous rows/characters
- Applies test‑specific filter criteria in one step

### Sub‑Macros for Data Validation
- **Input Checker:** Verifies that manually entered parameters (e.g., sample IDs, thresholds) match expected formats.
- **Structure Enforcer:** Ensures pasted data aligns with template columns and inserts any missing headers.
- **Error Corrector:** Auto‑fixes common typos (extra spaces, misplaced decimals) and prompts the user when manual review is needed.

These safeguards guarantee that the main macro always receives clean, correctly organized data.

### Purity Analysis Module
- Uses in‑memory arrays to compute analyte‑to‑impurity ratios, percent purity, and standard deviations
- Flags any samples below defined thresholds

### Dynamic Reporting Module
- Generates a new workbook per request with parameterized, color‑coded charts in a uniform layout

### Error Handling & Logging Module
- Validates all required data columns and formats
- Logs anomalies (missing values, out‑of‑range readings) to a hidden “Log” sheet

### Master UserForm Interface
- One‑click execution for the entire pipeline
- Real‑time progress indicators and error messages

## Impact & Results

- **Massive Time Savings:** Reduced per‑request processing from **~38 minutes** to **under 3 minutes**—a **>1200% improvement**.
- **Zero Manual Errors:** Sub‑macros eliminated copy‑paste and formatting mistakes through automated validation and correction.
- **Consistent Reporting:** Every report now has the same styling, formatting, and pass/fail indicators.
- **Extensible Framework:** Easily adaptable to other lab tests by updating filter thresholds and metric calculations.
- **Rapid Rollout:** Adopted by three other metrology teams within two weeks, each realizing similar productivity gains.

## Technologies & Skills

- **Excel VBA:** Advanced use of the Excel Object Model, UserForms, and ChartObjects
- **Data Manipulation:** Efficient array-based processing and batch operations with `Application.ScreenUpdating = False`
- **Statistical & Purity Calculations:** Custom VBA functions to compute particle counts, purity percentages, and variance flags
- **Error Handling:** Uses strategic `On Error Resume Next` statements around pattern‑based data operations, then checks `Err.Number` immediately after to handle or log exceptions, leveraging the consistent data layout to minimize complex error logic.
- **Performance Tuning:** Minimized Select/Activate calls and optimized memory usage for large datasets
