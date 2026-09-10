# Privacy Policy — HTML Export Viewer

**Last updated: 2026-09-11**

HTML Export Viewer is a Power BI custom visual published by PBIwise.

## Data collection

This visual **does not collect, store, or transmit any data**. It does not make network
requests of any kind, does not use telemetry or analytics, and does not communicate with any
external service.

## What the visual does with your data

The visual receives one thing: the text of a single DAX measure bound to its "HTML content"
data role, exactly as any other Power BI visual receives the fields bound to it. It renders that
content inside the report, and — only when the report user clicks an export button, and only if
your Power BI tenant's "Allow downloads from custom visuals" setting permits it — converts the
currently displayed content into a PDF, Excel, or Word file, which is saved to the user's own
device via Power BI's own built-in download mechanism. No data leaves the Power BI environment
through this visual at any point.

## Permissions

The visual requests exactly one Power BI privilege: `ExportContent`, which is required to save
the exported file to the user's device. No other privilege (network access, local storage, or
otherwise) is requested.

## Contact

Questions about this policy: info@pbiwise.com
