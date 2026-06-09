import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as XLSX from "xlsx";

type Row = Record<string, unknown>;

interface ColumnDef {
    key: string;
    header: string;
}

const DEFAULT_FILE_NAME = "PaymentTransactions";
const SHEET_NAME = "Payment Transactions";

export class PaymentTransactionsExport
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

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        _state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        this.context = context;
        this.notifyOutputChanged = notifyOutputChanged;
        this.container = container;
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
                this.runExport();
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

    /** Build + download using the format from the manifest property. */
    private runExport(formatOverride?: "excel" | "csv"): void {
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

        this.export(fmt, baseName);
    }

    private export(format: "excel" | "csv", baseName: string): void {
        try {
            if (this.buffer.length === 0) {
                this.setStatus("No data to export");
                this.notifyOutputChanged();
                return;
            }

            const cols = this.resolveColumns();

            if (format === "csv") {
                const csv = this.buildCsv(cols);
                // Prepend a UTF-8 BOM so Excel renders Greek/Unicode correctly.
                const blob = new Blob(["﻿" + csv], {
                    type: "text/csv;charset=utf-8;",
                });
                this.downloadBlob(blob, `${baseName}.csv`);
            } else {
                const aoa = this.toArrayOfArrays(cols);
                const ws = XLSX.utils.aoa_to_sheet(aoa);
                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, SHEET_NAME);
                const out = XLSX.write(wb, {
                    bookType: "xlsx",
                    type: "array",
                }) as ArrayBuffer;
                const blob = new Blob([out], {
                    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                });
                this.downloadBlob(blob, `${baseName}.xlsx`);
            }

            this.lastError = "";
            this.setStatus(`Exported ${this.buffer.length} rows (${format})`);
        } catch (e) {
            this.lastError = `Export error: ${(e as Error).message}`;
            this.setStatus(this.lastError);
        }
        this.notifyOutputChanged();
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
            this.setStatus(`${this.buffer.length} rows ready`);
        } catch (e) {
            this.lastError = `Parse error: ${(e as Error).message}`;
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

    private toArrayOfArrays(cols: ColumnDef[]): (string | number)[][] {
        const rows: (string | number)[][] = [cols.map((c) => c.header)];
        for (const r of this.buffer) {
            rows.push(cols.map((c) => this.cell(r[c.key])));
        }
        return rows;
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

    private buildCsv(cols: ColumnDef[]): string {
        const esc = (val: string | number): string => {
            const s = String(val ?? "");
            // Quote if it contains a delimiter, quote, or newline.
            return /[",\n\r;]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
        };
        const lines: string[] = [];
        lines.push(cols.map((c) => esc(c.header)).join(","));
        for (const r of this.buffer) {
            lines.push(cols.map((c) => esc(this.cell(r[c.key]))).join(","));
        }
        return lines.join("\r\n");
    }

    // -- Download -------------------------------------------------------------

    private downloadBlob(blob: Blob, fileName: string): void {
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
            this.runExport("excel")
        );
        const csvBtn = this.makeButton("Export to CSV", "vms-pte-btn--csv", () =>
            this.runExport("csv")
        );

        this.buttonsEl.appendChild(excelBtn);
        this.buttonsEl.appendChild(csvBtn);

        this.statusEl = document.createElement("div");
        this.statusEl.className = "vms-pte-status";

        this.container.appendChild(this.buttonsEl);
        this.container.appendChild(this.statusEl);
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

    private refreshStatusUI(): void {
        if (!this.statusEl) {
            return;
        }
        const hasError = this.lastError.length > 0;
        this.statusEl.textContent = hasError ? this.lastError : this.status;
        this.statusEl.classList.toggle("vms-pte-status--error", hasError);
    }
}
