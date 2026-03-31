import { createRoot } from "react-dom/client";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { App } from "./App";

/* global Office */

function renderApp() {
  const container = document.getElementById("root");
  if (!container) {
    document.body.innerHTML = "<p>Error: #root element not found</p>";
    return;
  }
  const root = createRoot(container);
  root.render(
    <FluentProvider theme={webLightTheme}>
      <App />
    </FluentProvider>
  );
}

try {
  Office.onReady(() => {
    renderApp();
  });
} catch {
  // Office.js not available (e.g. direct browser access) — render anyway
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", renderApp);
  } else {
    renderApp();
  }
}
