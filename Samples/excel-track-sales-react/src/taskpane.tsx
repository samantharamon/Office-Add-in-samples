/* global console, document, Office */

import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { createRoot } from "react-dom/client";
import { App } from "./App";

const container = document.getElementById("root");
if (!container) {
  throw new Error("The task pane root element is missing.");
}
const root = createRoot(container);

function showMessage(message: string): void {
  root.render(<p role="status">{message}</p>);
}

Office.onReady()
  .then((info) => {
    if (info.host !== Office.HostType.Excel) {
      showMessage("Open this add-in in Excel to track sales.");
      return;
    }
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.13")) {
      showMessage("This sample requires ExcelApi 1.13 or later. Update Excel before continuing.");
      return;
    }
    root.render(
      <FluentProvider theme={webLightTheme} className="app">
        <App />
      </FluentProvider>
    );
  })
  .catch((error: unknown) => {
    console.error(error);
    showMessage(error instanceof Error ? error.message : String(error));
  });
