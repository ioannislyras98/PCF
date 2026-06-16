import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as XLSX from "xlsx";

type Row = Record<string, unknown>;

interface ColumnDef {
    key: string;
    header: string;
    // Optional Excel cell typing/formatting, driven from the canvas ColumnsConfig.
    type?: "text" | "number" | "date";
    format?: string;
}

const DEFAULT_FILE_NAME = "Export";
const SHEET_NAME = "Data";

export class DataExport
    implements ComponentFramework.StandardControl<IInputs, IOutputs>
{
    private context!: ComponentFramework.Context<IInputs>;
    private notifyOutputChanged!: () => void;
    private container!: HTMLDivElement;

    // In-memory buffer of rows. Lives in plain JS memory, so it is NOT subject to
    // the canvas ~2000 row limit; it can hold the full (paged) result set.
    private buffer: Row[] = [];

    // Change-trackers for the increment-style trigger properties.
    private lastTrigger = 0;
    private lastDataVersion = 0;
    private lastReset = 0;

    // Output state.
    private status = "";
    private rowCount = 0;
    private lastError = "";

    // UI elements.
    private buttonsEl!: HTMLDivElement;
    private statusEl!: HTMLDivElement;
    private downloadEl!: HTMLDivElement;
    private downloadBtn!: HTMLButtonElement;

    // File prepared by a host-app trigger, waiting for the user's Download click.
    private pendingBlob: Blob | null = null;
    private pendingFileName = "";

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        _state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        this.context = context;
        this.notifyOutputChanged = notifyOutputChanged;
        this.container = container;
        // Baseline the increment-triggers to their CURRENT values so a persisted /
        // non-zero initial value (e.g. a context/global var that survived navigation)
        // does NOT auto-fire an export on first render. Only a later increment fires.
        this.lastTrigger = context.parameters.triggerExport.raw ?? 0;
        this.lastDataVersion = context.parameters.dataVersion.raw ?? 0;
        this.lastReset = context.parameters.resetBuffer.raw ?? 0;
        this.renderUI();
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this.context = context;
        const p = context.parameters;

        // 1) Reset buffer (increment to fire).
        const reset = p.resetBuffer.raw ?? 0;
        if (reset !== this.lastReset) {
            this.lastReset = reset;
            if (reset) {
                this.buffer = [];
                this.setStatus(`Buffer cleared`);
            }
        }

        // 2) Append the current chunk (increment dataVersion to fire). Used for paged feeding.
        const dataVersion = p.dataVersion.raw ?? 0;
        if (dataVersion !== this.lastDataVersion) {
            this.lastDataVersion = dataVersion;
            if (dataVersion) {
                this.ingest(p.inputData.raw ?? "", /*append*/ true);
            }
        }

        // 3) Export (increment triggerExport to fire).
        const trigger = p.triggerExport.raw ?? 0;
        if (trigger !== this.lastTrigger) {
            this.lastTrigger = trigger;
            if (trigger) {
                this.log(
                    "triggerExport fired →",
                    trigger,
                    "| format:",
                    p.exportFormat.raw,
                    "| appendMode:",
                    p.appendMode.raw,
                    "| inputData length:",
                    (p.inputData.raw ?? "").length
                );
                this.prepareExport();
            }
        }

        // Toggle the built-in buttons: shown unless explicitly set to false.
        this.buttonsEl.style.display = p.showButtons.raw === false ? "none" : "flex";

        this.refreshStatusUI();
    }

    public getOutputs(): IOutputs {
        return {
            status: this.status,
            rowCount: this.rowCount,
            lastError: this.lastError,
        };
    }

    public destroy(): void {
        // No external listeners to clean up.
    }

    // -- Export orchestration -------------------------------------------------

    /**
     * Triggered from the HOST APP (triggerExport increment). Browsers / the Power
     * Apps player iframe only allow a file download that is started by a real user
     * gesture (sandbox `allow-downloads`). A download fired here — after the async
     * flow, with no fresh gesture — gets silently blocked in stricter runtimes
     * (e.g. published/managed). So we only BUILD the file and reveal a Download
     * button; the actual save happens on the user's click (savePending).
     */
    private prepareExport(): void {
        const built = this.runBuild();
        if (!built) {
            this.hideDownloadButton();
            return;
        }
        this.pendingBlob = built.blob;
        this.pendingFileName = built.fileName;
        this.log(
            "prepareExport: file READY →",
            built.fileName,
            "(",
            this.buffer.length,
            "rows ) — waiting for the user's Download click"
        );
        this.setStatus(`Ready: ${this.buffer.length} rows — click Download`);
        this.showDownloadButton();
        this.notifyOutputChanged();
    }

    /** Built-in buttons: a direct user click, so the file can download immediately. */
    private runExportNow(formatOverride: "excel" | "csv"): void {
        const built = this.runBuild(formatOverride);
        if (!built) {
            return;
        }
        this.downloadBlob(built.blob, built.fileName);
        this.setStatus(`Exported ${this.buffer.length} rows (${formatOverride})`);
        this.notifyOutputChanged();
    }

    /** Save the file prepared by prepareExport() — fired by the user's Download click. */
    private savePending(): void {
        if (!this.pendingBlob) {
            this.log("savePending: no file prepared — ignoring click");
            return;
        }
        this.log("savePending: user clicked Download →", this.pendingFileName);
        this.downloadBlob(this.pendingBlob, this.pendingFileName);
        this.setStatus(`Downloaded ${this.pendingFileName}`);
        // Hide the Download button once used, so the host-app Export button
        // (which may sit behind it) becomes visible/clickable again.
        this.hideDownloadButton();
        this.notifyOutputChanged();
    }

    /** Ingest the current input (single-shot) and build the file; null on no data/error. */
    private runBuild(
        formatOverride?: "excel" | "csv"
    ): { blob: Blob; fileName: string } | null {
        const p = this.context.parameters;
        const append = p.appendMode.raw === true;
        // In single-shot (non-append) mode, treat inputData as the full dataset.
        if (!append) {
            this.ingest(p.inputData.raw ?? "", /*append*/ false);
        }
        const fmt =
            formatOverride ??
            ((p.exportFormat.raw ?? "excel").toString().trim().toLowerCase() === "csv"
                ? "csv"
                : "excel");
        const baseName =
            (p.fileName.raw ?? "").toString().trim() || DEFAULT_FILE_NAME;
        return this.buildFile(fmt, baseName);
    }

    private buildFile(
        format: "excel" | "csv",
        baseName: string
    ): { blob: Blob; fileName: string } | null {
        try {
            if (this.buffer.length === 0) {
                this.log("buildFile: NO DATA in buffer — nothing to export");
                this.setStatus("No data to export");
                this.notifyOutputChanged();
                return null;
            }
            const cols = this.resolveColumns();
            this.log(
                "buildFile:",
                format,
                "·",
                this.buffer.length,
                "rows ·",
                cols.length,
                "columns →",
                cols.map((c) => c.key)
            );
            if (format === "csv") {
                const csv = this.buildCsv(cols, this.csvDelimiter());
                // Prepend a UTF-8 BOM so Excel renders Greek/Unicode correctly.
                const blob = new Blob(["﻿" + csv], {
                    type: "text/csv;charset=utf-8;",
                });
                this.lastError = "";
                return { blob, fileName: `${baseName}.csv` };
            }
            const aoa = this.toArrayOfArrays(cols);
            const ws = XLSX.utils.aoa_to_sheet(aoa, { cellDates: true });
            this.applyColumnFormats(ws, cols);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, this.sheetName());
            const out = XLSX.write(wb, {
                bookType: "xlsx",
                type: "array",
            }) as ArrayBuffer;
            const blob = new Blob([out], {
                type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            });
            this.lastError = "";
            return { blob, fileName: `${baseName}.xlsx` };
        } catch (e) {
            this.lastError = `Export error: ${(e as Error).message}`;
            console.error("[DataExport] buildFile error:", e);
            this.setStatus(this.lastError);
            this.notifyOutputChanged();
            return null;
        }
    }

    /** CSV delimiter from the manifest property (default comma). */
    private csvDelimiter(): string {
        const d = (this.context.parameters.csvDelimiter.raw ?? "").toString();
        return d.length > 0 ? d : ",";
    }

    /** Excel sheet/tab name from the manifest property, sanitized for Excel's rules. */
    private sheetName(): string {
        const raw =
            (this.context.parameters.sheetName.raw ?? "").toString().trim() ||
            SHEET_NAME;
        // Excel tab names: max 31 chars, cannot contain \ / ? * [ ] :
        return raw.replace(/[\\/?*[\]:]/g, " ").slice(0, 31) || SHEET_NAME;
    }

    /** Apply per-column Excel number/date formats (the `format` in ColumnsConfig). */
    private applyColumnFormats(ws: XLSX.WorkSheet, cols: ColumnDef[]): void {
        const n = this.buffer.length;
        cols.forEach((col, c) => {
            if (!col.format) {
                return;
            }
            for (let i = 0; i < n; i++) {
                const addr = XLSX.utils.encode_cell({ c, r: i + 1 });
                const cell = ws[addr] as XLSX.CellObject | undefined;
                if (cell) {
                    cell.z = col.format;
                }
            }
        });
    }

    // -- Data ingestion -------------------------------------------------------

    private ingest(json: string, append: boolean): void {
        try {
            const rows = this.parseRows(json);
            if (append) {
                this.buffer.push(...rows);
            } else {
                this.buffer = rows;
            }
            this.rowCount = this.buffer.length;
            this.lastError = "";
            this.log(
                "ingest: parsed",
                rows.length,
                "rows (append:",
                append,
                ") — buffer now",
                this.buffer.length
            );
            this.setStatus(`${this.buffer.length} rows ready`);
        } catch (e) {
            this.lastError = `Parse error: ${(e as Error).message}`;
            console.error("[DataExport] ingest/parse error:", e, "| raw input:", json);
            this.setStatus(this.lastError);
        }
        this.notifyOutputChanged();
    }

    private parseRows(json: string): Row[] {
        if (!json || !json.trim()) {
            return [];
        }
        const data = JSON.parse(json);
        if (Array.isArray(data)) {
            return data as Row[];
        }
        // Accept a raw OData response shape: { value: [...] }.
        if (data && Array.isArray((data as { value?: unknown }).value)) {
            return (data as { value: Row[] }).value;
        }
        throw new Error("inputData must be a JSON array of row objects");
    }

    // -- Column resolution + cell formatting ----------------------------------

    private resolveColumns(): ColumnDef[] {
        const cfg = this.context.parameters.columnsConfig.raw;
        if (cfg && cfg.trim()) {
            const parsed = JSON.parse(cfg);
            if (Array.isArray(parsed)) {
                return parsed.map((c: unknown) =>
                    typeof c === "string"
                        ? { key: c, header: c }
                        : {
                              key: (c as ColumnDef).key,
                              header: (c as ColumnDef).header ?? (c as ColumnDef).key,
                              type: (c as ColumnDef).type,
                              format: (c as ColumnDef).format,
                          }
                );
            }
        }
        // No config: infer columns from the union of keys (first row drives order).
        if (this.buffer.length > 0) {
            return Object.keys(this.buffer[0]).map((k) => ({ key: k, header: k }));
        }
        return [];
    }

    private toArrayOfArrays(cols: ColumnDef[]): (string | number | Date)[][] {
        const rows: (string | number | Date)[][] = [cols.map((c) => c.header)];
        for (const r of this.buffer) {
            rows.push(cols.map((c) => this.typedCell(r[c.key], c)));
        }
        return rows;
    }

    /** Coerce a value to a real number / Date for Excel when the column declares a type. */
    private typedCell(v: unknown, col: ColumnDef): string | number | Date {
        if (v === null || v === undefined || v === "") {
            return "";
        }
        if (col.type === "number") {
            const n = typeof v === "number" ? v : Number(v);
            return Number.isNaN(n) ? this.cell(v) : n;
        }
        if (col.type === "date") {
            const d = v instanceof Date ? v : new Date(String(v));
            return Number.isNaN(d.getTime()) ? this.cell(v) : d;
        }
        return this.cell(v);
    }

    private cell(v: unknown): string | number {
        if (v === null || v === undefined) {
            return "";
        }
        if (typeof v === "number" || typeof v === "string") {
            return v;
        }
        if (typeof v === "boolean") {
            return v ? "TRUE" : "FALSE";
        }
        return String(v);
    }

    private buildCsv(cols: ColumnDef[], delim: string): string {
        const esc = (val: string | number): string => {
            const s = String(val ?? "");
            // Quote if it contains the delimiter, a quote, or a newline.
            const needsQuote =
                s.includes(delim) ||
                s.includes('"') ||
                s.includes("\n") ||
                s.includes("\r");
            return needsQuote ? `"${s.replace(/"/g, '""')}"` : s;
        };
        const lines: string[] = [];
        lines.push(cols.map((c) => esc(c.header)).join(delim));
        for (const r of this.buffer) {
            lines.push(cols.map((c) => esc(this.csvCell(r[c.key], c))).join(delim));
        }
        return lines.join("\r\n");
    }

    /** CSV cell value: numbers stay numeric; dates/text use the plain string form. */
    private csvCell(v: unknown, col: ColumnDef): string | number {
        if (col.type === "number") {
            const n = typeof v === "number" ? v : Number(v);
            return Number.isNaN(n) ? this.cell(v) : n;
        }
        return this.cell(v);
    }

    // -- Download -------------------------------------------------------------

    private downloadBlob(blob: Blob, fileName: string): void {
        this.log("downloadBlob: saving", fileName, "(", blob.size, "bytes )");
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = fileName;
        a.style.display = "none";
        document.body.appendChild(a);
        a.click();
        // Defer cleanup so the browser has time to start the download.
        window.setTimeout(() => {
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }, 1500);
    }

    // -- UI -------------------------------------------------------------------

    private renderUI(): void {
        this.container.classList.add("vms-pte-root");

        this.buttonsEl = document.createElement("div");
        this.buttonsEl.className = "vms-pte-buttons";

        const excelBtn = this.makeButton("Export to Excel", "vms-pte-btn--excel", () =>
            this.runExportNow("excel")
        );
        const csvBtn = this.makeButton("Export to CSV", "vms-pte-btn--csv", () =>
            this.runExportNow("csv")
        );

        this.buttonsEl.appendChild(excelBtn);
        this.buttonsEl.appendChild(csvBtn);

        // Prepared-download button: shown after a host-app trigger built a file.
        // The user clicks it so the actual download has a real user gesture.
        this.downloadEl = document.createElement("div");
        this.downloadEl.className = "vms-pte-download";
        this.downloadEl.style.display = "none";
        this.downloadBtn = this.makeButton("⬇ Download", "vms-pte-btn--download", () =>
            this.savePending()
        );
        this.downloadEl.appendChild(this.downloadBtn);

        this.statusEl = document.createElement("div");
        this.statusEl.className = "vms-pte-status";

        this.container.appendChild(this.buttonsEl);
        this.container.appendChild(this.downloadEl);
        this.container.appendChild(this.statusEl);
    }

    private showDownloadButton(): void {
        const p = this.context.parameters;
        // Label, background and text color are all canvas-driven so the button can
        // match the host app's theme. Empty values fall back to the CSS defaults.
        const label = (p.downloadLabel.raw ?? "").toString().trim();
        this.downloadBtn.textContent =
            label || `⬇ Download ${this.pendingFileName}`;
        const bg = (p.downloadColor.raw ?? "").toString().trim();
        const fg = (p.downloadTextColor.raw ?? "").toString().trim();
        this.downloadBtn.style.background = bg;
        this.downloadBtn.style.borderColor = bg;
        this.downloadBtn.style.color = fg;
        this.downloadEl.style.display = "flex";
    }

    private hideDownloadButton(): void {
        this.pendingBlob = null;
        this.downloadEl.style.display = "none";
    }

    private makeButton(
        label: string,
        modifier: string,
        onClick: () => void
    ): HTMLButtonElement {
        const btn = document.createElement("button");
        btn.type = "button";
        btn.className = `vms-pte-btn ${modifier}`;
        btn.textContent = label;
        btn.addEventListener("click", onClick);
        return btn;
    }

    private setStatus(message: string): void {
        this.status = message;
        this.rowCount = this.buffer.length;
    }

    /** Console tracing for debugging — open DevTools (F12) and filter by "DataExport". */
    private log(...args: unknown[]): void {
        console.log("[DataExport]", ...args);
    }

    private refreshStatusUI(): void {
        if (!this.statusEl) {
            return;
        }
        // Only surface errors in the UI; normal status stays silent (the host app
        // does not want any text under the button). Full tracing is in the console.
        // When there is nothing to show, collapse the row so the Download button can
        // fill the control's full Height.
        const hasError = this.lastError.length > 0;
        this.statusEl.textContent = hasError ? this.lastError : "";
        this.statusEl.style.display = hasError ? "block" : "none";
        this.statusEl.classList.toggle("vms-pte-status--error", hasError);
    }
}
