import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";
import "./global.css";

/* global document, Office, module, require, HTMLElement */

const title = "Solvid Add-in";

const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

/* Render application after Office initializes */
// Wait for Office.js to be fully loaded
try {
  if (typeof Office === "undefined") {
    // Office.js not loaded, show error
    console.error("Office.js is not loaded. Make sure office.js script is included in taskpane.html");
    if (rootElement) {
      rootElement.innerHTML = "<p>Error: Office.js failed to load. Please reload the add-in.</p>";
    }
  } else {
    Office.onReady((info) => {
      if (info.host === Office.HostType.Word) {
        root?.render(<App />);
      } else {
        console.error("This add-in must be run in Word");
        if (rootElement) {
          rootElement.innerHTML = "<p>This add-in must be run in Microsoft Word.</p>";
        }
      }
    });
  }
} catch (error) {
  console.error("Error initializing Office:", error);
  if (rootElement) {
    rootElement.innerHTML = "<p>Error initializing Office.js. Please reload the add-in.</p>";
  }
}

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    root?.render(NextApp);
  });
}
