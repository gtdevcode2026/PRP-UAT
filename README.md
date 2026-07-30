# PRP Automation Dashboard

AB InBev TPRM & Risk Intelligence dashboard.

## Quick start - browser only (no install needed)

1. Download this repo: **Code → Download ZIP**, then unzip it.
2. Keep `index.html` next to the `vendor/` and `js/` folders.
3. **Double-click `index.html`.** Build the PRP workbook in **Create PRP**, pick a report, click **Run**.

All steps run in JavaScript in your browser (SheetJS, ExcelJS, Plotly - all
vendored locally). Charts are generated to match the original Python script
output exactly. Recommended browsers: Chrome or Edge.

## Usage

1. **Create PRP** (hero section): drop the three raw exports - the filenames
   are matched automatically (`*risk-export*.xlsx`, `*assessment-export*.xlsx`,
   `*tprm*`/`*supplier*.xlsx`) - and click **Create PRP workbook**. The
   consolidated workbook (dates normalised, `Aging`/`Ageing`/`Working1`/
   `Working2`/`30sep26` formulas added) downloads and is auto-loaded into the
   pipeline below. Already have a consolidated workbook? Skip this step and
   upload it directly in Step 2.
2. Select a **report** from the report drawer (Vendor Onboarding, Response Time Reduction, etc.).
3. Choose a **variant** if multiple options are available.
4. Click **Run** to execute the report and preview charts and tables.
5. Download the generated `.xlsx` report from the **Download Options** section.

## Reports

| # | Name | Source Sheet |
|---|------|-------------|
| 1 | Vendor Onboarding (Critical Tech) | TPRM Web-Portal Export |
| 2 | Response Time Reduction | OneTrust Assessment |
| 3 | Reduce Long Over-Due Assessments | OneTrust Assessment |
| 4 | Risk Treatment | OneTrust - Risk Export |
| 5 | Suppliers in Scope | TPRM Web-Portal Export |
| 6 | Risk Assessment Progress | OneTrust Assessment |
| 7 | Risks Identified | OneTrust Assessment |
