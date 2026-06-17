import { StrictMode, useCallback, useEffect, useMemo, useState } from "react";
import { createRoot } from "react-dom/client";

import { AddinConnection } from "./connection.js";

const backendUrl = new URLSearchParams(window.location.search).get("backendUrl") ?? "ws://127.0.0.1:37845/addin";

function App() {
  const [status, setStatus] = useState("OpenWorkbook is ready to connect.");
  const [connection, setConnection] = useState<AddinConnection>();

  const statusTone = useMemo(() => {
    const normalized = status.toLowerCase();
    if (normalized.includes("workbook ready")) return "success";
    if (normalized.includes("connected to local")) return "success";
    if (normalized.includes("could not") || normalized.includes("did not load")) return "error";
    if (normalized.includes("disconnected") || normalized.includes("retrying") || normalized.includes("waiting") || normalized.includes("connecting")) return "warning";
    return "neutral";
  }, [status]);

  const connectLocal = useCallback(() => {
    if (connection) return;

    setStatus("Waiting for Office.js...");

    if (!("Office" in window)) {
      setStatus("Office.js did not load. Check network access to appsforoffice.microsoft.com.");
      return;
    }

    Office.onReady(() => {
      setStatus("Office.js ready. Connecting to local runtime...");
      try {
        const nextConnection = new AddinConnection({
          backendUrl,
          heartbeatMs: 5_000,
          reconnectMs: 2_000,
          onStatus: setStatus
        });
        setConnection(nextConnection);
        nextConnection.connect();
      } catch (error) {
        const detail = error instanceof Error ? error.message : String(error);
        setStatus(`Could not load taskpane runtime: ${detail}`);
      }
    });
  }, [connection]);

  useEffect(() => {
    connectLocal();
  }, [connectLocal]);

  return (
    <main className="taskpane-shell">
      <LandingView status={status} statusTone={statusTone} />
    </main>
  );
}

function BrandLockup() {
  return (
    <div className="brand-lockup" aria-label="ComponentsKit OpenWorkbook">
      <img className="brand-logo" alt="ComponentsKit" src="/assets/icon-80.png" />
      <span className="brand-divider">/</span>
      <span className="brand-product">OpenWorkbook</span>
    </div>
  );
}

function LandingView({ status, statusTone }: { status: string; statusTone: string }) {
  return (
    <section className="landing-view" aria-labelledby="landing-title">
      <div className="landing-content">
        <BrandLockup />

        <div className="landing-copy">
          <h1 id="landing-title">Excel your spreadsheet with any AI agent</h1>
          <p>Local-first MCP for live Excel automation with safe edits, backups, rollback, formulas, tables, and style preservation.</p>
        </div>

        <StatusBadge status={status} statusTone={statusTone} />
      </div>

      <WorkbookPreview />
    </section>
  );
}

function StatusBadge({ status, statusTone }: { status: string; statusTone: string }) {
  return <div className={`connection-badge connection-${statusTone}`} role="status">Status: {formatStatus(status)}</div>;
}

function formatStatus(status: string) {
  const normalized = status.toLowerCase();
  if (normalized.includes("workbook ready")) return "Workbook ready";
  if (normalized.includes("connected to local")) return "Connected to local runtime";
  if (normalized.includes("office.js did not load")) return "Office.js did not load";
  if (normalized.includes("could not")) return "Could not connect";
  if (normalized.includes("disconnected")) return "Disconnected, retrying";
  if (normalized.includes("retrying")) return "Retrying local runtime";
  return "Connecting to local runtime";
}

const previewRows = [
  ["Owner", "Variance", "Status", "Note"],
  ["A. Chen", "+$18K", "Review", "Travel spike"],
  ["M. Lee", "-$7K", "Ready", "Budget matched"],
  ["N. Park", "+$42K", "Updated", "Contractor accrual"]
];

function WorkbookPreview() {
  return (
    <div className="workbook-preview-card" aria-hidden="true">
      <div className="workbook-preview-window">
        <div className="workbook-preview-titlebar">
          <div className="window-dots">
            <span className="dot-red" />
            <span className="dot-yellow" />
            <span className="dot-green" />
          </div>
          <span className="autosave-label">AutoSave</span>
          <span className="autosave-toggle" />
          <span className="preview-tool">Home</span>
          <span className="preview-tool muted">Insert</span>
          <span className="preview-tool muted">Data</span>
        </div>

        <div className="workbook-preview-ribbon">
          <span className="ribbon-button wide" />
          <span className="ribbon-divider" />
          <span className="ribbon-button" />
          <span className="ribbon-button short" />
          <span className="ribbon-pill">Wrap Text</span>
        </div>

        <div className="workbook-preview-formula">
          <span className="name-box">C3</span>
          <span className="formula-mark">fx</span>
          <span className="formula-box">Controller note updated</span>
        </div>

        <div className="workbook-preview-sheet">
          <div className="sheet-grid">
            <div className="sheet-corner" />
            {["A", "B", "C", "D"].map((column) => (
              <div className="sheet-column" key={column}>
                {column}
              </div>
            ))}
            {previewRows.map((row, rowIndex) => [
              <div className="sheet-row-number" key={`row-${rowIndex}`}>
                {rowIndex + 1}
              </div>,
              ...row.map((cell, cellIndex) => (
                <div
                  className={`sheet-cell ${rowIndex === 0 ? "sheet-header" : ""} ${rowIndex === 2 && cellIndex === 2 ? "sheet-selected" : ""} ${rowIndex > 0 && cellIndex > 0 ? "sheet-agent-cell" : ""}`}
                  key={`${rowIndex}-${cellIndex}`}
                >
                  <span>{cell}</span>
                </div>
              ))
            ])}
          </div>
        </div>
      </div>
    </div>
  );
}

const rootElement = document.getElementById("root");

if (!rootElement) {
  throw new Error("Missing #root element for OpenWorkbook taskpane.");
}

createRoot(rootElement).render(
  <StrictMode>
    <App />
  </StrictMode>
);
