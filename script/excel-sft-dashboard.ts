
---

```javascript
/**
 * Excel SFT Transformation Script (Office.js / Script Lab)
 *
 * What it does:
 * - Builds the Dashboard from Raw Data (formulas + formatting)
 * - Adds health classification + conditional formatting
 * - Builds a Chart Data table
 * - Creates a combo chart (margins as columns + revenue as a line on secondary axis)
 *
 * Assumptions:
 * - Worksheets exist: "Dashboard" and "Raw Data"
 * - "Dashboard" template already has the product + quarter grid in A8:B39
 * - Raw Data layout supports the SUMIFS formulas (see README)
 */

Excel.run(async (context) => {
  const dashboard = context.workbook.worksheets.getItem("Dashboard");
  const rawData = context.workbook.worksheets.getItem("Raw Data");

  const ROWS_8_TO_39 = 32;  // rows 8..39 inclusive
  const ROWS_44_TO_51 = 8;  // rows 44..51 inclusive

  /* --------------------------------------------------
     1) Total Revenue (C8:C39)
     -------------------------------------------------- */
  dashboard.getRange("C8").formulas = [[
    "=SUMIFS('Raw Data'!$E:$E,'Raw Data'!$D:$D,$A8,'Raw Data'!$B:$B,LEFT($B8,4),'Raw Data'!$C:$C,RIGHT($B8,2))"
  ]];

  dashboard.getRange("C8").autoFill(
    dashboard.getRange("C8:C39"),
    Excel.AutoFillType.fillDefault
  );

  dashboard.getRange("C8:C39").numberFormat =
    Array.from({ length: ROWS_8_TO_39 }, () => ["$#,##0"]);

  /* --------------------------------------------------
     2) Weighted Average Margin (D8:D39)
     -------------------------------------------------- */
  dashboard.getRange("D8").formulas = [[
    "=IFERROR(SUMIFS('Raw Data'!$G:$G,'Raw Data'!$D:$D,$A8,'Raw Data'!$B:$B,LEFT($B8,4),'Raw Data'!$C:$C,RIGHT($B8,2))/C8,\"N/A\")"
  ]];

  dashboard.getRange("D8").autoFill(
    dashboard.getRange("D8:D39"),
    Excel.AutoFillType.fillDefault
  );

  dashboard.getRange("D8:D39").numberFormat =
    Array.from({ length: ROWS_8_TO_39 }, () => ["0.0%"]);

  /* --------------------------------------------------
     3) Rolling Trend (E8:E39)
        - For the first quarter, show N/A
        - Otherwise, if the product repeats and both values are numeric, show delta vs previous row
     -------------------------------------------------- */
  dashboard.getRange("E8").formulas = [[
    "=IF($B8=\"2023 Q1\",\"N/A\",IF(AND($A8=$A7,ISNUMBER(D8),ISNUMBER(D7)),D8-D7,\"N/A\"))"
  ]];

  dashboard.getRange("E8").autoFill(
    dashboard.getRange("E8:E39"),
    Excel.AutoFillType.fillDefault
  );

  dashboard.getRange("E8:E39").numberFormat =
    Array.from({ length: ROWS_8_TO_39 }, () => ["0.0%"]);

  /* --------------------------------------------------
     4) Year-over-Year Margin Delta (F8:F39)
        - For 2023 rows, show N/A
        - For later years, compare to same product + same quarter in the prior year
     -------------------------------------------------- */
  dashboard.getRange("F8").formulas = [[
    "=IF(LEFT($B8,4)=\"2023\",\"N/A\",IFERROR(D8-INDEX($D$8:$D$39,MATCH(1,($A$8:$A$39=$A8)*($B$8:$B$39=(TEXT(VALUE(LEFT($B8,4))-1,\"0\")&\" \"&RIGHT($B8,2))),0)),\"N/A\"))"
  ]];

  dashboard.getRange("F8").autoFill(
    dashboard.getRange("F8:F39"),
    Excel.AutoFillType.fillDefault
  );

  dashboard.getRange("F8:F39").numberFormat =
    Array.from({ length: ROWS_8_TO_39 }, () => ["0.0%"]);

  /* --------------------------------------------------
     5) Margin Health Classification (G8:G39)
        - Strong:   > 35%
        - Moderate: 20% to 35%
        - At Risk:  < 20%
     -------------------------------------------------- */
  dashboard.getRange("G8").formulas = [[
    "=IF(ISNUMBER(D8),IF(D8>0.35,\"Strong\",IF(D8>=0.2,\"Moderate\",\"At Risk\")),\"N/A\")"
  ]];

  dashboard.getRange("G8").autoFill(
    dashboard.getRange("G8:G39"),
    Excel.AutoFillType.fillDefault
  );

  // Make sure formulas are fully calculated before we start styling/charting.
  context.workbook.application.calculate(Excel.CalculationType.full);
  await context.sync();

  /* --------------------------------------------------
     6) Conditional Formatting (G8:G39)
     -------------------------------------------------- */
  const healthRange = dashboard.getRange("G8:G39");
  healthRange.conditionalFormats.clearAll();

  const strongRule = healthRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
  strongRule.custom.rule.formula = '=$G8="Strong"';
  strongRule.custom.format.fill.color = "#C6EFCE";

  const moderateRule = healthRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
  moderateRule.custom.rule.formula = '=$G8="Moderate"';
  moderateRule.custom.format.fill.color = "#FFEB9C";

  const riskRule = healthRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
  riskRule.custom.rule.formula = '=$G8="At Risk"';
  riskRule.custom.format.fill.color = "#F4CCCC";

  /* --------------------------------------------------
     7) Chart Data Table header + quarters
     -------------------------------------------------- */
  dashboard.getRange("A42").values = [["Chart Data"]];
  dashboard.getRange("A43:F43").values = [[
    "Quarter",
    "Widget Pro",
    "Widget Standard",
    "Service Package",
    "Accessory Kit",
    "Total Revenue"
  ]];

  dashboard.getRange("A44:A51").values = [
    ["2023 Q1"], ["2023 Q2"], ["2023 Q3"], ["2023 Q4"],
    ["2024 Q1"], ["2024 Q2"], ["2024 Q3"], ["2024 Q4"]
  ];

  /* --------------------------------------------------
     8) Product margins per quarter (B44:E51)
        - Uses SUMPRODUCT against the Dashboard grid (A8:B39 + D8:D39)
     -------------------------------------------------- */
  dashboard.getRange("B44").formulas = [[
    "=SUMPRODUCT(($B$8:$B$39=$A44)*($A$8:$A$39=\"Widget Pro\")*($D$8:$D$39))"
  ]];
  dashboard.getRange("C44").formulas = [[
    "=SUMPRODUCT(($B$8:$B$39=$A44)*($A$8:$A$39=\"Widget Standard\")*($D$8:$D$39))"
  ]];
  dashboard.getRange("D44").formulas = [[
    "=SUMPRODUCT(($B$8:$B$39=$A44)*($A$8:$A$39=\"Service Package\")*($D$8:$D$39))"
  ]];
  dashboard.getRange("E44").formulas = [[
    "=SUMPRODUCT(($B$8:$B$39=$A44)*($A$8:$A$39=\"Accessory Kit\")*($D$8:$D$39))"
  ]];

  dashboard.getRange("B44:E44").autoFill(
    dashboard.getRange("B44:E51"),
    Excel.AutoFillType.fillDefault
  );

  dashboard.getRange("B44:E51").numberFormat =
    Array.from({ length: ROWS_44_TO_51 }, () => ["0.0%","0.0%","0.0%","0.0%"]);

  /* --------------------------------------------------
     9) Total Revenue by quarter (F44:F51)
        - Computed from Raw Data using JS, so it's stable even if formulas change
     -------------------------------------------------- */
  const rawUsed = rawData.getUsedRange();
  rawUsed.load("values");
  await context.sync();

  const rawValues = rawUsed.values;
  if (!rawValues || rawValues.length < 2) {
    throw new Error("Raw Data sheet looks empty (no rows found).");
  }

  const headers = rawValues[0].map(h => String(h).trim());
  const yearIdx = headers.indexOf("Year");
  const quarterIdx = headers.indexOf("Quarter");
  const revenueIdx = headers.indexOf("Revenue");

  if (yearIdx === -1 || quarterIdx === -1 || revenueIdx === -1) {
    throw new Error(
      "Raw Data must include headers: Year, Quarter, Revenue (exact spelling)."
    );
  }

  const revenueByQuarter = new Map();

  for (let i = 1; i < rawValues.length; i++) {
    const year = String(rawValues[i][yearIdx]).trim();
    const q = String(rawValues[i][quarterIdx]).trim();
    const rev = Number(rawValues[i][revenueIdx]) || 0;

    if (!year || !q) continue;

    const key = `${year} ${q}`;
    revenueByQuarter.set(key, (revenueByQuarter.get(key) || 0) + rev);
  }

  const quarterCells = dashboard.getRange("A44:A51");
  quarterCells.load("values");
  await context.sync();

  const totals = quarterCells.values.map(row => {
    const key = String(row[0]).trim();
    return [revenueByQuarter.get(key) || 0];
  });

  const totalRevenueRange = dashboard.getRange("F44:F51");
  totalRevenueRange.values = totals;
  totalRevenueRange.numberFormat =
    Array.from({ length: ROWS_44_TO_51 }, () => ["$#,##0"]);

  /* --------------------------------------------------
     10) Exclude 2023 Q1 from chart plotting
         - We keep the quarter label, but set the series cells to #N/A
         - That prevents Excel from plotting the first point/bar
     -------------------------------------------------- */
  dashboard.getRange("B44:F44").values = [[
    "#N/A", "#N/A", "#N/A", "#N/A", "#N/A"
  ]];

  /* --------------------------------------------------
     11) Chart Creation
         - Clears existing charts on Dashboard
         - Adds clustered column chart for margins
         - Converts "Total Revenue" series to line on secondary axis
     -------------------------------------------------- */
  dashboard.charts.load("items");
  await context.sync();
  dashboard.charts.items.forEach(c => c.delete());

  const chart = dashboard.charts.add(
    Excel.ChartType.columnClustered,
    dashboard.getRange("A43:F51"),
    Excel.ChartSeriesBy.columns
  );

  chart.title.text = "Quarterly Margin Trends by Product";
  chart.axes.valueAxis.title.text = "Profit Margin";

  chart.series.load("items/name");
  await context.sync();

  for (const s of chart.series.items) {
    if (s.name === "Total Revenue") {
      s.axisGroup = Excel.ChartAxisGroup.secondary;
      s.chartType = Excel.ChartType.line;
    }
  }

  await context.sync();

  // Secondary axis title (only available after the series is moved)
  if (chart.axes.secondaryValueAxis) {
    chart.axes.secondaryValueAxis.title.text = "Total Revenue ($)";
  }

  // Place + size chart (tweak to match your template)
  chart.top = 550;
  chart.left = 20;
  chart.height = 350;
  chart.width = 850;

  await context.sync();
}).catch(err => console.log("Script failed:", err));
