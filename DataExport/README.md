# DataExport — Excel/CSV Export

PCF code component (`VMS.DataExport`, namespace `VMS`) that builds and downloads an
**Excel (.xlsx)** or **CSV** file from data supplied by the **canvas app**. It is **generic and
reusable on any screen** — the file name, columns and data are all passed in as properties.
Designed for a canvas host, so it does not call Dataverse itself — the app (via a Power Automate
flow) supplies the rows, and this control turns them into a file and triggers the browser
download.

> Originally built for **PBM-18 (Payment Transactions export)**; the PBM-18 column/data samples
> below are just one example of how to drive it.

## Why this design

- Host is a **canvas app** → `context.webAPI` is not available, so the control cannot query Dataverse.
- Must export **all** filtered rows (possibly > 50k) → canvas collections cap at ~2000, so a flow
  retrieves the rows server-side (Dataverse *List rows*, **Pagination ON**) and feeds them here.
- Rows are accumulated in **plain JS memory** (no 2000 cap), optionally fed in **chunks** for very
  large exports, then written to a file client-side and downloaded.

## Properties (contract)

| Property | Type | Dir | Purpose |
|---|---|---|---|
| `exportFormat` | Text | in | `"excel"` (default) or `"csv"`. |
| `fileName` | Text | in | Base name without extension, **set from the canvas app** (e.g. `"Payment Transactions"`). Default `Export`. |
| `sheetName` | Text | in | Excel worksheet/tab name. Default `Data`. |
| `csvDelimiter` | Text | in | CSV column separator. Default `,`. Use `;` for Greek/European Excel. |
| `inputData` | Multiline | in | JSON array of row objects, or an OData `{value:[...]}` object. |
| `columnsConfig` | Multiline | in | Optional `[{ "key", "header", "type", "format" }]` — controls which columns, their order, headers and **Excel cell type**. `type` = `text` (default) / `number` / `date`; `format` = optional Excel number/date format (e.g. `#,##0.00`, `dd/mm/yyyy`). If empty, inferred from the data. |
| `appendMode` | TwoOptions | in | ON = each fed chunk is appended (paged export). OFF = `inputData` is the full dataset on export. |
| `dataVersion` | Whole | in | **Increment** to ingest the current `inputData` chunk (append mode). |
| `triggerExport` | Whole | in | **Increment** to build + download the file. |
| `resetBuffer` | Whole | in | **Increment** to clear the buffer before a new paged export. |
| `showButtons` | TwoOptions | in | Show the built-in Excel/CSV buttons (default shown; turn OFF in production). |
| `status` | Text | out | Last status message. |
| `rowCount` | Whole | out | Rows currently in the buffer. |
| `lastError` | Text | out | Last error, if any. |

## Usage from canvas (Power Fx)

**Small / single-shot** (≤ a few thousand rows):
```
// Excel button OnSelect
UpdateContext({
    ctxFormat: "excel",
    ctxData: JSON(colFilteredTransactions),   // or the flow's returned JSON
    ctxTrigger: ctxTrigger + 1
});
```
Bind: `ExportFormat = ctxFormat`, `InputData = ctxData`, `TriggerExport = ctxTrigger`, `AppendMode = false`.

**Large / paged** (loop the flow page-by-page):
```
// 1) reset, 2) for each page: set InputData=page JSON and bump DataVersion,
// 3) bump TriggerExport to finish.
```
Bind `AppendMode = true`. (Exact paging loop is finalized when the flow is built.)

## Local testing (test harness)

```
npm start watch
```
In the harness right-hand panel:
1. Paste a sample array into **Input Data (JSON)** (see below).
2. Click the built-in **Export to Excel** / **Export to CSV** button → the file downloads.
3. To test the trigger path: set **Export Format** = `csv`, then set **Trigger Export** = `1`.
4. To test chunked append: turn **Append Mode** ON, paste chunk 1 + set **Data Version** = `1`,
   paste chunk 2 + set **Data Version** = `2`, then set **Trigger Export** = `1`.

### Sample `inputData`
```json
[
  {"paymentId":"PAY-001","paymentStatus":"Paid","serviceTitle":"STCW Basic","vesselName":"MV Aurora","trainingCentreName":"Athens TC","fee":350,"gpgOrderId":"GPG-9001","gpgOrderCreatedOn":"2026-05-01T10:15:00Z","invoiceId":""},
  {"paymentId":"PAY-002","paymentStatus":"Pending","serviceTitle":"Advanced Firefighting","vesselName":"MV Böreas","trainingCentreName":"Πειραιάς TC","fee":420.5,"gpgOrderId":"GPG-9002","gpgOrderCreatedOn":"2026-05-03T08:00:00Z","invoiceId":""}
]
```

### Sample `columnsConfig` (PBM-18 columns, in order)
```json
[
  {"key":"paymentId","header":"Payment ID"},
  {"key":"paymentStatus","header":"Payment Status"},
  {"key":"serviceTitle","header":"Service Title"},
  {"key":"vesselName","header":"Vessel Name"},
  {"key":"trainingCentreName","header":"Training Centre Name"},
  {"key":"fee","header":"Fee","type":"number","format":"#,##0.00"},
  {"key":"gpgOrderId","header":"GPG Order ID"},
  {"key":"gpgOrderCreatedOn","header":"GPG Order Creation Date - Time","type":"date","format":"dd/mm/yyyy hh:mm"},
  {"key":"invoiceId","header":"Invoice ID"}
]
```
> `Invoice ID` requires OAS integration — leave blank / omit the column until that is available.

## Deploy (after local test passes)

This control ships inside the shared **`VMS_PCF`** Dataverse solution together with the other
VMS components. To build and deploy see [../CONTRIBUTING.md](../CONTRIBUTING.md) (build the
solution, import the unmanaged zip into DEV, then promote to TEST/PROD via the pipeline).

For a quick local-to-dev iteration you can also use `pac pcf push --publisher-prefix vms`.
