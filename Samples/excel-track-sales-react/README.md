---
title: Track Excel sales with Fluent UI React tabs
page_type: sample
urlFragment: excel-track-sales-react
products:
  - office-excel
  - office
  - m365
languages:
  - typescript
extensions:
  contentType: samples
  technologies:
    - Add-ins
  createdDate: "11/25/2026 10:00:00 AM"
description: "Use Fluent UI React tabs to import demo sales into Excel and customize tables and charts."
---

# Track Excel sales with Fluent UI React tabs

Explore sales trends using a React task pane with Fluent UI React v9 controls. The sample demonstrates how to:

- Navigate **Sales**, **Table**, and **Chart** tabs with Fluent UI `TabList` and `Tab`.
- Import bundled demo sales filtered by product and an inclusive date range.
- Change Excel table styles and worksheet column visibility.
- Create and update native Excel line, clustered column, and area charts.
- Preserve selections across tabs and display workbook progress and errors.

For the same workbook operations with standard browser controls, see the [TypeScript sales tracker](../excel-track-sales). Both samples use bundled demo data, not a live sales service.

## Applies to

- Excel on Windows
- Excel on Mac
- Excel on the web

The complete workflow requires **ExcelApi 1.13**, including support for resizing an existing table, and uses the unified manifest for Microsoft 365. Manifest support requires Excel on Windows version 2501 or later, or Excel on Mac version 16.103 or later. The client must meet both the API and manifest requirements. For more information, see:

- [Excel API requirement sets](https://learn.microsoft.com/javascript/api/requirement-sets/excel/excel-api-requirement-sets)
- [Unified manifest platform support](https://learn.microsoft.com/office/dev/add-ins/develop/unified-manifest-overview#client-and-platform-support)

## Prerequisites

- Node.js 24 LTS or a later supported LTS release, and npm.
- A supported Excel installation connected to a Microsoft 365 account.

## Run the sample

### Install and sideload the add-in

1. Clone or download this repository, then open a terminal in **Samples/excel-track-sales-react**.
1. Install dependencies.

   ```console
   npm install
   ```

1. Start the HTTPS development server and sideload the add-in.

   ```console
   npm start
   ```

   Trust the development certificate if prompted. Both sales tracker samples use port 3000, so run only one development server at a time.

### Import and customize sales

1. Open an Excel workbook, then select **Home** > **Sales Tracker React** > **Track sales (React)** to open the task pane.
1. On the **Sales** tab, select a product and start and end dates within November 1-12, 2026. Select **Get sales data**. The date range includes both endpoints and can cover a single day.

   The add-in creates or updates the **Sales Tracker** worksheet, **SalesTrackerData** table, and **SalesTrackerChart** chart.

1. On the **Table** tab, select the columns to show and a table color, then select **Apply table settings**. Hiding a column hides the entire worksheet column, not just the table cells.
1. On the **Chart** tab, select a data source, chart type, and chart theme, then select **Apply chart settings**.

   Selections persist when switching tabs. Settings are disabled until a sales table exists and while a workbook operation is running. Use the arrow keys to focus a tab, then press Enter or Space to select it.

> [!NOTE]
> Each import resets the table color and column visibility but retains chart selections. Both sales tracker samples use the same worksheet, table, and chart names. Use separate workbooks to compare them independently.

### Stop and remove the add-in

Run the following command to stop the development server and remove the sideloaded add-in:

```console
npm stop
```

## Key parts of the sample

The workflow starts with filtered sales records, writes them to an Excel table, and uses that table as the source for a chart. The task pane keeps the user's settings available across all three steps.

### Preserve selections across the workflow

The task pane uses [`TabList` and `Tab`](https://react.fluentui.dev/?path=/docs/components-tablist--docs) to select a workflow panel. React manages the selected tab and form values in [App.tsx](src/App.tsx). Each tab is linked to its corresponding `tabpanel`, and inactive panels are hidden from keyboard navigation and assistive technology. The values remain available across tabs, so importing sales uses the selected chart options even when the Chart panel is hidden.

[`FluentProvider`](https://react.fluentui.dev/?path=/docs/components-fluentprovider--docs) supplies the light theme. Fluent fields, selects, date inputs, checkboxes, buttons, a spinner, and a message bar provide the task pane UI. Fluent UI styles the task pane, not the workbook's native table and chart.

### Import sales into an Excel table

After the user selects a product and date range, [sales-data.ts](src/sales-data.ts) validates and filters the bundled records. [sales-workbook.ts](src/sales-workbook.ts) converts the dates to Excel date values and writes the records through [`Range.values`](https://learn.microsoft.com/javascript/api/excel/excel.range?view=excel-js-1.13#excel-excel-range-values-member).

The first import creates the named table and chart. Later imports use [`Table.resize`](https://learn.microsoft.com/javascript/api/excel/excel.table?view=excel-js-1.13#excel-excel-table-resize-member(1)) to fit the filtered records and update the existing objects rather than creating duplicates.

### Customize table styles and column visibility

Once a sales table exists, the user can change its color and visible columns. The Table panel applies [`Table.style`](https://learn.microsoft.com/javascript/api/excel/excel.table?view=excel-js-1.13#excel-excel-table-style-member) and [`Range.columnHidden`](https://learn.microsoft.com/javascript/api/excel/excel.range?view=excel-js-1.13#excel-excel-range-columnhidden-member) in [sales-workbook.ts](src/sales-workbook.ts). Column visibility applies to the entire worksheet column.

For more information, see [Create, read, and manage Excel tables with JavaScript](https://learn.microsoft.com/office/dev/add-ins/excel/excel-add-ins-tables).

### Customize native Excel charts

To explore a different view of the sales, the user selects a metric, chart type, and theme. The Chart panel updates [`Chart.chartType`](https://learn.microsoft.com/javascript/api/excel/excel.chart?view=excel-js-1.13#excel-excel-chart-charttype-member) and chart colors in [sales-workbook.ts](src/sales-workbook.ts).

[`ChartSeries.setXAxisValues`](https://learn.microsoft.com/javascript/api/excel/excel.chartseries?view=excel-js-1.13#excel-excel-chartseries-setxaxisvalues-member(1)) and [`ChartSeries.setValues`](https://learn.microsoft.com/javascript/api/excel/excel.chartseries?view=excel-js-1.13#excel-excel-chartseries-setvalues-member(1)) connect the chart to the table's date and metric columns. Setting [`Chart.plotVisibleOnly`](https://learn.microsoft.com/javascript/api/excel/excel.chart?view=excel-js-1.13#excel-excel-chart-plotvisibleonly-member) to `false` keeps the metric plotted when its worksheet column is hidden.

See [Create and customize Excel charts with JavaScript](https://learn.microsoft.com/office/dev/add-ins/excel/excel-add-ins-charts).

### Coordinate asynchronous workbook operations

[taskpane.tsx](src/taskpane.tsx) waits for Office to initialize and checks the Excel host before rendering the React app. Workbook operations in [sales-workbook.ts](src/sales-workbook.ts) batch commands, load the properties needed for decisions, and synchronize changes with Excel.

The app checks for an existing sales table on startup. A ref guards against overlapping operations across tabs, while React state disables controls and displays progress. Failures are logged and shown in the message bar, including Office error codes and locations when available.

For more information, see [Using the application-specific API model](https://learn.microsoft.com/office/dev/add-ins/develop/application-specific-api-model).

## Validate the sample

```console
npm run build
npm run validate
```

Build and manifest checks don't verify workbook behavior or sideloading. Before publishing, test imports, repeated imports, table formatting, and each chart type in Excel on Windows, Mac, and the web. Confirm that selections persist across tabs, controls are disabled during workbook operations, and errors appear in the message bar.

## Solution

| Solution | Author(s) |
| --- | --- |
| Track Excel sales with Fluent UI React tabs | Microsoft |

## Version history

| Version | Date | Comments |
| --- | --- | --- |
| 1.0 | November 25, 2026 | Initial release. |

## Additional resources

- [Use React to build an Excel task pane add-in](https://learn.microsoft.com/office/dev/add-ins/quickstarts/excel-quickstart-react)
- [Fluent UI React v9](https://react.fluentui.dev/)
- [Excel JavaScript API reference](https://learn.microsoft.com/javascript/api/excel)

## Questions and feedback

- Did you experience any problems with the sample? [Create an issue](https://github.com/OfficeDev/Office-Add-in-samples/issues/new/choose) and we'll help you out.
- For general questions about developing Office Add-ins, go to [Microsoft Q&A](https://learn.microsoft.com/answers/topics/office-js-dev.html) using the office-js-dev tag.

## Copyright

Copyright (c) 2026 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/excel-track-sales-react" />
