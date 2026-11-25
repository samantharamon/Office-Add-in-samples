/* global console, OfficeExtension */

import { useEffect, useRef, useState } from "react";
import {
  Button,
  Checkbox,
  Field,
  Input,
  MessageBar,
  MessageBarBody,
  Select,
  Spinner,
  Tab,
  TabList,
} from "@fluentui/react-components";
import {
  END_DATE,
  getSales,
  HEADERS,
  loadSalesData,
  METRICS,
  PRODUCTS,
  START_DATE,
  validateDateRange,
} from "./sales-data";
import {
  CHART_THEMES,
  CHART_TYPES,
  ChartOptions,
  formatChart,
  formatTable,
  hasSalesTable,
  importSales,
  TABLE_STYLES,
} from "./sales-workbook";

type SalesTab = "sales" | "table" | "chart";
interface Feedback {
  message: string;
  intent: "info" | "success" | "error";
}

function settingKey<T extends object>(value: string, choices: T): keyof T {
  if (!Object.prototype.hasOwnProperty.call(choices, value)) {
    throw new Error("Choose a valid setting.");
  }
  return value as keyof T;
}

function errorFeedback(error: unknown): Feedback {
  console.error(error);
  let message = error instanceof Error ? error.message : String(error);
  if (error instanceof OfficeExtension.Error) {
    const location = error.debugInfo.errorLocation;
    message = `${error.message} (${error.code}${location ? ` at ${location}` : ""})`;
  }
  return { message, intent: "error" };
}

export function App() {
  const [tab, setTab] = useState<SalesTab>("sales");
  const [product, setProduct] = useState<string>(PRODUCTS[0]);
  const [startDate, setStartDate] = useState(START_DATE);
  const [endDate, setEndDate] = useState(END_DATE);
  const [columns, setColumns] = useState<string[]>([...HEADERS]);
  const [tableStyle, setTableStyle] = useState<keyof typeof TABLE_STYLES>("blue");
  const [metric, setMetric] = useState("ProductSales");
  const [chartType, setChartType] = useState<ChartOptions["type"]>("line");
  const [theme, setTheme] = useState<ChartOptions["theme"]>("light");
  const [dataLoaded, setDataLoaded] = useState(false);
  const [busy, setBusy] = useState(true);
  const busyRef = useRef(true);
  const [feedback, setFeedback] = useState<Feedback>({
    message: "Checking for an existing sales table...",
    intent: "info",
  });

  useEffect(() => {
    let active = true;
    hasSalesTable()
      .then((exists) => {
        if (active) {
          setDataLoaded(exists);
          setFeedback({
            message: "Ready. Import demo sales, or apply settings to an existing sales table.",
            intent: "info",
          });
        }
      })
      .catch((error: unknown) => {
        const failure = errorFeedback(error);
        if (active) {
          setFeedback(failure);
        }
      })
      .finally(() => {
        if (active) {
          busyRef.current = false;
          setBusy(false);
        }
      });
    return () => {
      active = false;
    };
  }, []);

  async function perform(action: () => Promise<void>, message: string): Promise<void> {
    if (busyRef.current) {
      return;
    }
    busyRef.current = true;
    setBusy(true);
    setFeedback({ message: "Working...", intent: "info" });
    try {
      await action();
      setFeedback({ message, intent: "success" });
    } catch (error) {
      setFeedback(errorFeedback(error));
    } finally {
      busyRef.current = false;
      setBusy(false);
    }
  }

  function chartOptions(): ChartOptions {
    return { metric, type: chartType, theme };
  }

  return (
    <>
      <header>
        <h1>Sales Tracker</h1>
        <p>Explore sales in an Excel table and chart.</p>
      </header>
      <TabList
        aria-label="Sales tracker workflow"
        selectedValue={tab}
        onTabSelect={(_, data) => {
          if (data.value === "sales" || data.value === "table" || data.value === "chart") {
            setTab(data.value);
          }
        }}
      >
        <Tab id="sales-tab" aria-controls="sales-panel" value="sales">
          Sales
        </Tab>
        <Tab id="table-tab" aria-controls="table-panel" value="table">
          Table
        </Tab>
        <Tab id="chart-tab" aria-controls="chart-panel" value="chart">
          Chart
        </Tab>
      </TabList>
      <main aria-busy={busy}>
        <section
          id="sales-panel"
          role="tabpanel"
          aria-labelledby="sales-tab"
          hidden={tab !== "sales"}
          tabIndex={0}
        >
          <h2>Demo sales</h2>
          <p>Sample data: November 1-12, 2026. Not live sales.</p>
          <form
            onSubmit={(event) => {
              event.preventDefault();
              void perform(async () => {
                validateDateRange(startDate, endDate);
                const sales = getSales(product, startDate, endDate, await loadSalesData());
                await importSales(sales, chartOptions());
                setDataLoaded(true);
                setTableStyle("blue");
                setColumns([...HEADERS]);
              }, `Imported sales for ${product} from ${startDate} through ${endDate}.`);
            }}
          >
            <fieldset disabled={busy} className="fields">
              <Field label="Product" required>
                <Select
                  disabled={busy}
                  value={product}
                  onChange={(_, data) => setProduct(data.value)}
                >
                  {PRODUCTS.map((value) => (
                    <option key={value} value={value}>
                      {value}
                    </option>
                  ))}
                </Select>
              </Field>
              <Field label="Start date" required>
                <Input
                  type="date"
                  disabled={busy}
                  required
                  min={START_DATE}
                  max={END_DATE}
                  value={startDate}
                  aria-describedby="date-help"
                  onChange={(_, data) => setStartDate(data.value)}
                />
              </Field>
              <Field label="End date" required>
                <Input
                  type="date"
                  disabled={busy}
                  required
                  min={START_DATE}
                  max={END_DATE}
                  value={endDate}
                  aria-describedby="date-help"
                  onChange={(_, data) => setEndDate(data.value)}
                />
              </Field>
              <p id="date-help">
                Includes sales on both dates. Choose a range within November 1-12, 2026.
              </p>
              <Button appearance="primary" type="submit" disabled={busy}>
                Get sales data
              </Button>
            </fieldset>
          </form>
          <p>Creates or refreshes the Sales Tracker table and chart. It resets table settings.</p>
        </section>
        <section
          id="table-panel"
          role="tabpanel"
          aria-labelledby="table-tab"
          hidden={tab !== "table"}
          tabIndex={0}
        >
          <h2>Table settings</h2>
          {!dataLoaded && <p>Import sales on the Sales tab before applying settings.</p>}
          <p>Hiding a column hides that worksheet column; the chart remains visible.</p>
          <fieldset disabled={busy || !dataLoaded} className="fields">
            <legend>Columns and table color</legend>
            <div className="columns">
              {HEADERS.map((header) => (
                <Checkbox
                  key={header}
                  label={header}
                  checked={columns.includes(header)}
                  disabled={busy || !dataLoaded}
                  onChange={(_, data) =>
                    setColumns((current) =>
                      data.checked
                        ? [...current, header]
                        : current.filter((value) => value !== header)
                    )
                  }
                />
              ))}
            </div>
            <Field label="Table color">
              <Select
                disabled={busy || !dataLoaded}
                value={tableStyle}
                onChange={(_, data) => setTableStyle(settingKey(data.value, TABLE_STYLES))}
              >
                <option value="blue">Blue</option>
                <option value="green">Green</option>
                <option value="orange">Orange</option>
              </Select>
            </Field>
            <Button
              disabled={busy || !dataLoaded}
              onClick={() =>
                void perform(() => formatTable(tableStyle, columns), "Table settings applied.")
              }
            >
              Apply table settings
            </Button>
          </fieldset>
        </section>
        <section
          id="chart-panel"
          role="tabpanel"
          aria-labelledby="chart-tab"
          hidden={tab !== "chart"}
          tabIndex={0}
        >
          <h2>Chart settings</h2>
          {!dataLoaded && <p>Import sales on the Sales tab before applying settings.</p>}
          <fieldset disabled={busy || !dataLoaded} className="fields">
            <legend>Chart data and appearance</legend>
            <Field label="Data source">
              <Select
                disabled={busy || !dataLoaded}
                value={metric}
                onChange={(_, data) => setMetric(data.value)}
              >
                {METRICS.map((value) => (
                  <option key={value} value={value}>
                    {value}
                  </option>
                ))}
              </Select>
            </Field>
            <Field label="Chart type">
              <Select
                disabled={busy || !dataLoaded}
                value={chartType}
                onChange={(_, data) => setChartType(settingKey(data.value, CHART_TYPES))}
              >
                <option value="line">Line</option>
                <option value="column">Clustered column</option>
                <option value="area">Area</option>
              </Select>
            </Field>
            <Field label="Chart theme">
              <Select
                disabled={busy || !dataLoaded}
                value={theme}
                onChange={(_, data) => setTheme(settingKey(data.value, CHART_THEMES))}
              >
                <option value="light">Light</option>
                <option value="dark">Dark</option>
                <option value="blue">Blue</option>
              </Select>
            </Field>
            <Button
              disabled={busy || !dataLoaded}
              onClick={() =>
                void perform(() => formatChart(chartOptions()), "Chart settings applied.")
              }
            >
              Apply chart settings
            </Button>
          </fieldset>
        </section>
      </main>
      <div className="feedback">
        {busy && <Spinner size="tiny" label="Working..." />}
        <MessageBar intent={feedback.intent} role="status" aria-live="polite">
          <MessageBarBody>{feedback.message}</MessageBarBody>
        </MessageBar>
      </div>
    </>
  );
}
