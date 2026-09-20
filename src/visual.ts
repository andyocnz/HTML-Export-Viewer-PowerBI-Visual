"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import * as ExcelJS from "exceljs";
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";
import * as DOMPurifyModule from "dompurify";
import "./../style/visual.less";

// Without esModuleInterop, TypeScript compiles `import * as X from "dompurify"` to a bare
// require(), but webpack resolves that to dompurify's ESM build and wraps it in a namespace
// object - the actual sanitizer ends up at `.default`, not on the imported binding itself.
// (Confirmed at runtime in the packaged visual: DOMPurifyModule.sanitize was undefined.)
// Handles both shapes so this doesn't silently break again if the resolved build changes.
const DOMPurify = (DOMPurifyModule as unknown as { default?: typeof DOMPurifyModule }).default
    || DOMPurifyModule;

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import IDownloadService = powerbi.extensibility.IDownloadService;
import PrivilegeStatus = powerbi.PrivilegeStatus;
import ISelectionManager = powerbi.extensibility.ISelectionManager;
import ISelectionIdBuilder = powerbi.visuals.ISelectionIdBuilder;
import ITooltipService = powerbi.extensibility.ITooltipService;

import { VisualFormattingSettingsModel } from "./settings";

interface ParsedCellStyle {
    backgroundHex: string;
    colorHex: string;
    bold: boolean;
    italic: boolean;
    align: "left" | "center" | "right";
    borderHex: string;
}

// A named, reusable value the source measure wants available to formula cells (e.g. a base
// rent, an escalation rate). Declared once per report via a hidden JSON block (see
// findAssumptions) and written to a small labeled block on the sheet; formula templates
// reference it by `ref`, not by any built-in business term.
interface AssumptionDef {
    ref: string;
    label: string;
    value: number;
    numFmt?: string;
}

export class Visual implements IVisual {
    private target: HTMLElement;
    private toolbarEl: HTMLDivElement;
    private pdfButton: HTMLButtonElement;
    private excelButton: HTMLButtonElement;
    private wordButton: HTMLButtonElement;
    private statusEl: HTMLSpanElement;
    private contentHost: HTMLDivElement;
    private shadow: ShadowRoot;
    private placeholder: HTMLDivElement;
    private downloadService: IDownloadService;
    private selectionManager: ISelectionManager;
    private selectionIdBuilder: ISelectionIdBuilder;
    private tooltipService: ITooltipService;

    private formattingSettings: VisualFormattingSettingsModel;
    private formattingSettingsService: FormattingSettingsService;

    private lastHtml: string = null;
    private statusTimer: ReturnType<typeof setTimeout>;

    // Custom visuals in Power BI run inside a sandboxed host iframe (allow-scripts only).
    // window.print() and raw Blob/anchor downloads are silently swallowed by that sandbox -
    // there is no error, the file just never lands anywhere. Content is rendered into a
    // Shadow DOM (for style isolation, no iframe needed) and both exports are handed to
    // Power BI's own host.downloadService, which is the only sanctioned way to save a file
    // from inside a custom visual (requires the ExportContent privilege in capabilities.json).
    constructor(options: VisualConstructorOptions) {
        this.formattingSettingsService = new FormattingSettingsService();
        this.target = options.element;
        this.target.classList.add("wolc-export-visual");
        this.downloadService = options.host.downloadService;
        this.selectionManager = options.host.createSelectionManager();
        this.selectionIdBuilder = options.host.createSelectionIdBuilder();
        this.tooltipService = options.host.tooltipService;

        // Right-click context menu (Power BI certification requirement 1180.2.5). This visual
        // has no discrete data points to attach a real identity to (content is a single HTML
        // measure, not categorical data), so every right-click - on the toolbar, the rendered
        // content, or empty space - opens the host's default "empty space" context menu via an
        // identity-less selection ID, per Microsoft's guidance for non-categorical visuals.
        this.target.addEventListener("contextmenu", (event: MouseEvent) => {
            event.preventDefault();
            this.selectionManager.showContextMenu(
                this.selectionIdBuilder.createSelectionId(),
                { x: event.clientX, y: event.clientY }
            );
        });

        this.toolbarEl = document.createElement("div");
        this.toolbarEl.className = "wolc-toolbar wolc-pos-bottom-right";

        const buttonRow = document.createElement("div");
        buttonRow.className = "wolc-toolbar-buttons";

        this.pdfButton = document.createElement("button");
        this.pdfButton.type = "button";
        this.pdfButton.className = "wolc-btn wolc-btn-pdf";
        this.pdfButton.addEventListener("click", () => this.exportPdf());

        this.excelButton = document.createElement("button");
        this.excelButton.type = "button";
        this.excelButton.className = "wolc-btn wolc-btn-excel";
        this.excelButton.addEventListener("click", () => this.exportExcel());

        this.wordButton = document.createElement("button");
        this.wordButton.type = "button";
        this.wordButton.className = "wolc-btn wolc-btn-word";
        this.wordButton.addEventListener("click", () => this.exportWord());

        this.statusEl = document.createElement("span");
        this.statusEl.className = "wolc-status";

        buttonRow.appendChild(this.pdfButton);
        buttonRow.appendChild(this.excelButton);
        buttonRow.appendChild(this.wordButton);
        this.toolbarEl.appendChild(buttonRow);
        this.toolbarEl.appendChild(this.statusEl);

        this.placeholder = document.createElement("div");
        this.placeholder.className = "wolc-placeholder";
        this.placeholder.textContent = "Select a single item to view content.";

        this.contentHost = document.createElement("div");
        this.contentHost.className = "wolc-content-host";
        this.shadow = this.contentHost.attachShadow({ mode: "open" });

        // Tooltips (Power BI certification requirement 1180.2.2.2). Delegated on the shadow
        // root rather than attached per-cell, since content is replaced wholesale on every
        // data update (see renderHtml) - this survives that without needing to re-wire
        // listeners each time. Shows the cell's own text; for a data cell in a row with a
        // label in its first column, the label is shown as the tooltip's header for context.
        this.shadow.addEventListener("mouseover", (event: MouseEvent) => this.handleCellHover(event));
        this.shadow.addEventListener("mousemove", (event: MouseEvent) => this.moveTooltip(event));
        this.shadow.addEventListener("mouseout", (event: MouseEvent) => this.handleCellLeave(event));

        this.target.appendChild(this.toolbarEl);
        this.target.appendChild(this.placeholder);
        this.target.appendChild(this.contentHost);
    }

    private handleCellHover(event: MouseEvent): void {
        const cell = (event.target as HTMLElement).closest("td, th") as HTMLTableCellElement;
        if (!cell) {
            return;
        }
        const value = (cell.textContent || "").trim();
        if (!value) {
            return;
        }
        const row = cell.closest("tr");
        const rowLabel = row && row.cells.length > 0 ? (row.cells[0].textContent || "").trim() : "";
        this.tooltipService.show({
            coordinates: [event.clientX, event.clientY],
            isTouchEvent: false,
            dataItems: [{ displayName: rowLabel && rowLabel !== value ? rowLabel : "Value", value }],
            identities: []
        });
    }

    private moveTooltip(event: MouseEvent): void {
        if (!(event.target as HTMLElement).closest("td, th")) {
            return;
        }
        this.tooltipService.move({
            coordinates: [event.clientX, event.clientY],
            isTouchEvent: false,
            identities: []
        });
    }

    private handleCellLeave(event: MouseEvent): void {
        if (!(event.target as HTMLElement).closest("td, th")) {
            return;
        }
        this.tooltipService.hide({ isTouchEvent: false, immediately: true });
    }

    public update(options: VisualUpdateOptions) {
        const dataView: powerbi.DataView = options.dataViews && options.dataViews[0];
        this.formattingSettings = this.formattingSettingsService.populateFormattingSettingsModel(VisualFormattingSettingsModel, dataView);

        const toolbar = this.formattingSettings.toolbarCard;
        this.toolbarEl.style.display = toolbar.show.value ? "flex" : "none";
        this.pdfButton.textContent = toolbar.pdfButtonText.value || "Export PDF";
        this.excelButton.textContent = toolbar.excelButtonText.value || "Export Excel";
        this.wordButton.textContent = toolbar.wordButtonText.value || "Export Word";

        const position = (toolbar.position.value && toolbar.position.value.value) || "bottom-right";
        this.toolbarEl.classList.remove("wolc-pos-bottom-right", "wolc-pos-bottom-left", "wolc-pos-top-left", "wolc-pos-top-right");
        this.toolbarEl.classList.add(`wolc-pos-${position}`);

        const html = this.getHtml(dataView);

        if (!html) {
            this.contentHost.style.display = "none";
            this.placeholder.style.display = "flex";
            this.pdfButton.disabled = true;
            this.excelButton.disabled = true;
            this.wordButton.disabled = true;
            this.lastHtml = null;
            return;
        }

        this.placeholder.style.display = "none";
        this.contentHost.style.display = "block";
        this.pdfButton.disabled = false;
        this.excelButton.disabled = false;
        this.wordButton.disabled = false;

        if (html !== this.lastHtml) {
            this.lastHtml = html;
            this.renderHtml(html);
        }
    }

    // Sanitized because the string comes from a DAX measure, not code we authored - report
    // authors (or upstream data) could otherwise inject a script/event handler. Any <style>
    // tags are pulled out and parsed separately first, wherever they appear in the source
    // (inside <head> or <body>): DOMPurify's default sanitization works from <body> downward
    // and silently discards everything under <head>, so a <head>-only style block would
    // otherwise vanish with no error. The extracted CSS text is attached via a real <style>
    // element's textContent (never innerHTML), which can't be interpreted as markup. The
    // remaining body content goes through DOMPurify as a DocumentFragment and is inserted via
    // replaceChildren, rather than assigning to innerHTML - avoids the innerHTML injection
    // pattern entirely instead of just sanitizing before it (also flagged by Microsoft's own
    // powerbi-visuals/no-inner-outer-html certification lint rule).
    private renderHtml(html: string): void {
        const parsed = new DOMParser().parseFromString(html, "text/html");
        const styleEls = Array.from(parsed.querySelectorAll("style"));
        const styleText = styleEls.map((el) => el.textContent || "").join("\n");
        styleEls.forEach((el) => el.remove());
        const bodyHtml = parsed.body ? parsed.body.innerHTML : html;

        const fragment = DOMPurify.sanitize(bodyHtml, {
            WHOLE_DOCUMENT: false,
            RETURN_DOM_FRAGMENT: true
        });
        // DOMPurify builds this fragment inside its own internal document (it parses via a
        // fresh DOMParser, not this visual's document), so its nodes belong to a different
        // document than the shadow root they're about to be inserted into. adoptNode makes
        // the fragment belong to this document explicitly before insertion, rather than
        // relying on the browser to tolerate a cross-document append.
        const adopted = document.adoptNode(fragment);

        this.shadow.replaceChildren();
        if (styleText) {
            const styleEl = document.createElement("style");
            styleEl.textContent = styleText;
            this.shadow.appendChild(styleEl);
        }
        this.shadow.appendChild(adopted);
    }

    private getHtml(dataView: powerbi.DataView): string {
        if (!dataView) {
            return "";
        }
        if (dataView.single && dataView.single.value != null) {
            return String(dataView.single.value);
        }
        if (dataView.categorical && dataView.categorical.values && dataView.categorical.values[0]
            && dataView.categorical.values[0].values && dataView.categorical.values[0].values[0] != null) {
            return String(dataView.categorical.values[0].values[0]);
        }
        return "";
    }

    private showStatus(message: string, isError: boolean): void {
        clearTimeout(this.statusTimer);
        this.statusEl.textContent = message;
        this.statusEl.classList.toggle("wolc-status-error", isError);
        this.statusTimer = setTimeout(() => {
            this.statusEl.textContent = "";
        }, 4000);
    }

    private getTables(): HTMLTableElement[] {
        return Array.from(this.shadow.querySelectorAll("table"));
    }

    private baseFileName(): string {
        const fileNameSetting = this.formattingSettings.toolbarCard.fileName.value;
        return (fileNameSetting || "Export").replace(/[\\/:*?"<>|]/g, "").trim() || "Export";
    }

    private async checkExportAllowed(): Promise<boolean> {
        try {
            const status = await this.downloadService.exportStatus();
            switch (status) {
                case PrivilegeStatus.Allowed:
                    return true;
                case PrivilegeStatus.DisabledByAdmin:
                    this.showStatus("Downloads are off - ask your admin to enable them (see console for steps).", true);
                    console.error(
                        "wolcHtmlExport: file export is turned off by your Power BI tenant admin. To enable it:\n" +
                        "1. Sign in to https://app.powerbi.com as a Power BI admin.\n" +
                        "2. Click the gear icon (top right) > Admin portal.\n" +
                        "3. Open 'Tenant settings'.\n" +
                        "4. Find the 'Custom visuals' section and expand 'Allow downloads from custom visuals'.\n" +
                        "5. Toggle it to Enabled, then choose 'The entire organization' or specific security groups.\n" +
                        "6. Click Apply. It can take up to 15 minutes to take effect.\n" +
                        "Note: this is a separate switch from the regular Export & sharing tenant settings, " +
                        "and only applies once the report is used in the Power BI Service (not needed for local Desktop testing)."
                    );
                    return false;
                case PrivilegeStatus.NotSupported:
                    this.showStatus("File export isn't supported in this environment.", true);
                    return false;
                default:
                    this.showStatus("File export isn't enabled for this visual.", true);
                    return false;
            }
        } catch (err) {
            console.error("wolcHtmlExport: export status check failed", err);
            this.showStatus("Could not check export permission - see console.", true);
            return false;
        }
    }

    private async exportPdf(): Promise<void> {
        try {
            const tables = this.getTables();
            if (tables.length === 0) {
                this.showStatus("No table found in the content to export.", true);
                return;
            }
            if (!(await this.checkExportAllowed())) {
                return;
            }

            const doc = new jsPDF({ orientation: "landscape", unit: "pt" });
            let cursorY = 40;

            tables.forEach((table: HTMLTableElement, index: number) => {
                const title = this.titleFor(table, index);
                if (index > 0) {
                    cursorY += 24; // gap between this title and the previous table's end
                }
                doc.setFontSize(13);
                doc.text(title, 40, cursorY);

                // useCss reads each source cell's live computed style (background, text
                // color, bold, alignment, borders) so the PDF matches the on-screen HTML
                // instead of a generic autoTable theme.
                autoTable(doc, {
                    html: table,
                    startY: cursorY + 12,
                    useCss: true,
                    margin: { left: 40, right: 40 }
                });

                // Continue from wherever autoTable actually finished, rather than forcing a
                // page break just because this was a second table - tables flow onto one
                // continuous page and only spill onto a new one if they don't fit, same as
                // the Excel export's single continuous sheet.
                const finishedTable = (doc as unknown as { lastAutoTable?: { finalY: number } }).lastAutoTable;
                cursorY = finishedTable ? finishedTable.finalY : cursorY;
            });

            const dataUri: string = doc.output("datauristring");
            const base64 = dataUri.substring(dataUri.indexOf(",") + 1);
            const fileName = `${this.baseFileName()}.pdf`;

            const result = await this.downloadService.exportVisualsContentExtended(base64, fileName, "base64", "PDF export");
            if (result.downloadCompleted) {
                this.showStatus("PDF file downloaded.", false);
            } else {
                this.showStatus("PDF download was cancelled or blocked.", true);
            }
        } catch (err) {
            console.error("wolcHtmlExport: PDF export failed", err);
            this.showStatus("PDF export failed - see console for details.", true);
        }
    }

    // Walks every top-level rendered element in document order - not just <table>s - so the
    // property-address banner, section headings, and the trailing numbered notes all make it
    // into the workbook, laid into ONE continuous sheet (matching the PDF export's continuous
    // page). "Worked-out" cells (a row's Total = the sum
    // of its own numeric cells, or a subtotal row's column = SUM of the rows above it) are
    // written as live Excel formulas rather than baked-in numbers, wherever the arithmetic can
    // be recovered unambiguously from the rendered values. Cells the source measure explicitly
    // marks with data-formula get a live formula built from the report's own assumptions block
    // (see findAssumptions / resolveFormulaTemplate) instead of that generic sum-detection.
    private async exportExcel(): Promise<void> {
        try {
            const nodes = Array.from(this.shadow.children)
                .filter((el: Element) => el.tagName !== "STYLE" && el.getAttribute("data-role") !== "export-assumptions") as HTMLElement[];
            if (!nodes.some((el) => el.tagName === "TABLE")) {
                this.showStatus("No table found in the content to export.", true);
                return;
            }
            if (!(await this.checkExportAllowed())) {
                return;
            }

            const workbook = new ExcelJS.Workbook();
            const sheetName = this.sanitizeSheetName(this.sheetTitleFromContent(nodes));
            const worksheet = workbook.addWorksheet(sheetName);

            const totalColumns = this.computeMaxColumns(nodes);
            const colWidths: number[] = [];
            let currentRow = 1;

            const assumptions = this.findAssumptions();
            const formulaRefs = assumptions ? this.writeAssumptionsBlock(worksheet, totalColumns, assumptions) : null;

            nodes.forEach((el: HTMLElement) => {
                if (el.tagName === "TABLE") {
                    currentRow = this.writeTableBlock(worksheet, el as HTMLTableElement, currentRow, totalColumns, colWidths, formulaRefs);
                    currentRow += 1; // spacer row, mirrors the HTML table's own bottom margin
                } else if (el.tagName === "P") {
                    currentRow = this.writeFullWidthTextRow(worksheet, el, currentRow, totalColumns);
                } else {
                    // DIV banners (dark-header / section-header) and headings (h1/h2/h3)
                    currentRow = this.writeFullWidthTextRow(worksheet, el, currentRow, totalColumns);
                }
            });

            worksheet.columns.forEach((column: Partial<ExcelJS.Column>, i: number) => {
                column.width = Math.min(Math.max((colWidths[i] || 8) + 2, 10), 60);
            });

            const buffer = await workbook.xlsx.writeBuffer();
            const base64 = this.arrayBufferToBase64(buffer);
            const fileName = `${this.baseFileName()}.xlsx`;

            const result = await this.downloadService.exportVisualsContentExtended(base64, fileName, "base64", "Excel export");
            if (result.downloadCompleted) {
                this.showStatus("Excel file downloaded.", false);
            } else {
                this.showStatus("Excel download was cancelled or blocked.", true);
            }
        } catch (err) {
            console.error("wolcHtmlExport: Excel export failed", err);
            this.showStatus("Excel export failed - see console for details.", true);
        }
    }

    private sheetTitleFromContent(nodes: HTMLElement[]): string {
        const banner = nodes.find((el) => el.classList.contains("dark-header"));
        const text = banner ? (banner.textContent || "").trim() : "";
        return text || "Export";
    }

    private computeMaxColumns(nodes: HTMLElement[]): number {
        let max = 1;
        nodes.forEach((el) => {
            if (el.tagName === "TABLE") {
                max = Math.max(max, this.tableColumnCount(el as HTMLTableElement));
            }
        });
        return max;
    }

    private tableColumnCount(table: HTMLTableElement): number {
        const rows = Array.from(table.rows);
        return Math.max(1, ...rows.map((row: HTMLTableRowElement) =>
            Array.from(row.cells).reduce((sum: number, cell: HTMLTableCellElement) => sum + (cell.colSpan || 1), 0)
        ));
    }

    // A banner (dark-header/section-header), a heading, or a note paragraph - rendered as one
    // merged row spanning the full table width, styled from its own computed CSS.
    private writeFullWidthTextRow(worksheet: ExcelJS.Worksheet, el: HTMLElement, row: number, totalColumns: number): number {
        const style = this.readCellStyle(el);
        const computed = window.getComputedStyle(el);
        const cell = worksheet.getRow(row).getCell(1);
        cell.value = (el.textContent || "").trim();
        if (style.backgroundHex) {
            cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: `FF${style.backgroundHex}` } };
        }
        cell.font = {
            name: "Arial",
            size: this.pxToPt(computed.fontSize),
            bold: style.bold,
            italic: style.italic,
            color: { argb: `FF${style.colorHex}` }
        };
        cell.alignment = { horizontal: style.align, vertical: "middle", wrapText: true };
        if (totalColumns > 1) {
            worksheet.mergeCells(row, 1, row, totalColumns);
        }
        return row + 1;
    }

    private pxToPt(px: string): number {
        const n = parseFloat(px);
        if (isNaN(n)) {
            return 9;
        }
        return Math.round(n * 0.75 * 10) / 10;
    }

    // Writes one <table> starting at `startRow` into the shared worksheet (rather than its own
    // sheet). If the table has fewer native columns than the widest table on the sheet, each
    // row's last cell is stretched to fill the remaining width so the whole sheet reads as one
    // consistent grid. `formulaRefs` (optional) resolves any cell carrying a data-formula
    // attribute into a live Excel formula referencing the assumptions block, instead of the
    // generic sum-detection static value.
    private writeTableBlock(
        worksheet: ExcelJS.Worksheet,
        table: HTMLTableElement,
        startRow: number,
        totalColumns: number,
        colWidths: number[],
        formulaRefs: Map<string, string>
    ): number {
        const rows = Array.from(table.rows);
        const tableColumnCount = this.tableColumnCount(table);
        const stretch = totalColumns - tableColumnCount;

        let dataRowFirst: number = null;
        let dataRowLast: number = null;

        rows.forEach((row: HTMLTableRowElement, rIdx: number) => {
            const excelRow = startRow + rIdx;
            const cells = Array.from(row.cells);
            const isHeaderRow = cells.length > 0 && cells[0].tagName === "TH";
            const isSubtotalRow = row.classList.contains("subtotal-row");
            const isTrackedDataRow = !isHeaderRow && !isSubtotalRow && !row.classList.contains("summary-row");

            let colPointer = 1;
            const rowNumericCells: { col: number; value: number }[] = [];

            cells.forEach((cell: HTMLTableCellElement, cIdx: number) => {
                let colSpan = cell.colSpan || 1;
                if (cIdx === cells.length - 1 && stretch > 0) {
                    colSpan += stretch;
                }

                const excelCell = worksheet.getRow(excelRow).getCell(colPointer);
                const parsed = this.parseCellValue(cell.textContent || "");
                const formulaTemplate = cell.dataset.formula;
                const customFormula = formulaTemplate && formulaRefs && typeof parsed.value === "number"
                    ? this.resolveFormulaTemplate(formulaTemplate, formulaRefs, cell.dataset.formulaVars)
                    : null;

                if (customFormula) {
                    excelCell.value = { formula: customFormula, result: parsed.value as number };
                    rowNumericCells.push({ col: colPointer, value: parsed.value as number });
                } else if (typeof parsed.value === "number") {
                    const horizontalFormula = this.tryHorizontalSumFormula(rowNumericCells, colPointer, parsed.value, excelRow);
                    if (horizontalFormula) {
                        excelCell.value = { formula: horizontalFormula, result: parsed.value };
                    } else if (isSubtotalRow && dataRowFirst != null) {
                        const colLetter = this.colLetter(colPointer);
                        excelCell.value = {
                            formula: `SUM(${colLetter}${dataRowFirst}:${colLetter}${dataRowLast})`,
                            result: parsed.value
                        };
                    } else {
                        excelCell.value = parsed.value;
                    }
                    rowNumericCells.push({ col: colPointer, value: parsed.value });
                } else {
                    excelCell.value = parsed.value;
                }
                if (parsed.numFmt) {
                    excelCell.numFmt = parsed.numFmt;
                }

                const style = this.readCellStyle(cell);
                if (style.backgroundHex) {
                    excelCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: `FF${style.backgroundHex}` } };
                }
                excelCell.font = {
                    name: "Arial",
                    size: 9,
                    bold: style.bold,
                    italic: style.italic,
                    color: { argb: `FF${style.colorHex}` }
                };
                excelCell.alignment = { horizontal: style.align, vertical: "middle", wrapText: true };
                if (style.borderHex) {
                    const border: Partial<ExcelJS.Border> = { style: "thin", color: { argb: `FF${style.borderHex}` } };
                    excelCell.border = { top: border, bottom: border, left: border, right: border };
                }

                if (colSpan > 1) {
                    worksheet.mergeCells(excelRow, colPointer, excelRow, colPointer + colSpan - 1);
                }

                const len = String(parsed.value ?? "").length;
                colWidths[colPointer - 1] = Math.max(colWidths[colPointer - 1] || 0, len);
                colPointer += colSpan;
            });

            if (isTrackedDataRow) {
                if (dataRowFirst == null) {
                    dataRowFirst = excelRow;
                }
                dataRowLast = excelRow;
            }
        });

        return startRow + rows.length;
    }

    // Looks for a single hidden block the source measure can optionally emit:
    // <div hidden data-role="export-assumptions" data-assumptions='[{"ref": "baseRent",
    //   "label": "Base rent (current)", "value": 200000, "numFmt": "$#,##0"}, ...]'></div>
    // (a <div> attribute, not a <script> tag/textContent - <script> isn't in the sanitizer's
    // allowlist and would be stripped before this ever runs). Each entry becomes one row in a
    // small labeled block on the sheet, and its `ref` is what data-formula templates on table
    // cells use to point at it (see resolveFormulaTemplate). This is the whole mechanism for
    // "live formula" exports - it carries no assumptions about what the values mean (rent,
    // opex, or anything else); that meaning lives entirely in the measure that emits the HTML.
    private findAssumptions(): AssumptionDef[] {
        const holder = this.shadow.querySelector('[data-role="export-assumptions"]') as HTMLElement;
        const raw = holder && holder.dataset.assumptions;
        if (!raw) {
            return null;
        }
        try {
            const parsed = JSON.parse(raw);
            if (!Array.isArray(parsed)) {
                return null;
            }
            const valid = parsed.filter((a) =>
                a && typeof a.ref === "string" && /^[A-Za-z_]\w*$/.test(a.ref)
                && typeof a.label === "string" && typeof a.value === "number"
            );
            return valid.length > 0 ? valid : null;
        } catch {
            return null;
        }
    }

    // Writes the assumptions block to a small labeled area to the right of the main content
    // and returns a map from each assumption's `ref` to its absolute cell address, so formula
    // templates can be resolved against real Excel references.
    private writeAssumptionsBlock(worksheet: ExcelJS.Worksheet, totalColumns: number, assumptions: AssumptionDef[]): Map<string, string> {
        const labelCol = totalColumns + 2;
        const valueCol = totalColumns + 3;
        const valueLetter = this.colLetter(valueCol);

        const header = worksheet.getRow(1).getCell(labelCol);
        header.value = "Assumptions (used by formulas)";
        header.font = { name: "Arial", size: 9, bold: true, italic: true, color: { argb: "FF4A5E3A" } };
        worksheet.mergeCells(1, labelCol, 1, valueCol);

        const refs = new Map<string, string>();
        assumptions.forEach((a, i) => {
            const rowNum = i + 2;
            const labelCell = worksheet.getRow(rowNum).getCell(labelCol);
            labelCell.value = a.label;
            labelCell.font = { name: "Arial", size: 9, color: { argb: "FF4A5E3A" } };
            const valueCell = worksheet.getRow(rowNum).getCell(valueCol);
            valueCell.value = a.value;
            if (a.numFmt) {
                valueCell.numFmt = a.numFmt;
            }
            valueCell.font = { name: "Arial", size: 9 };
            refs.set(a.ref, `$${valueLetter}$${rowNum}`);
        });

        return refs;
    }

    // A cell opts into a live formula with two attributes:
    //   data-formula="{baseRent}*(1+{years}*{rentRate})"
    //   data-formula-vars="years:3"
    // {tokens} are resolved first against the assumptions ref map, then against the per-cell
    // vars string (comma-separated key:number pairs) - so the row-specific numbers (like "3
    // years") live on the cell itself while the shared constants live in one place. Returns
    // null (falling back to the normal static/sum value) if any token is unresolved or if the
    // resolved text doesn't pass isFormulaSafe.
    private resolveFormulaTemplate(template: string, refs: Map<string, string>, varsAttr: string): string {
        const vars = new Map<string, string>();
        if (varsAttr) {
            varsAttr.split(",").forEach((pair) => {
                const [key, raw] = pair.split(":").map((s) => (s || "").trim());
                if (key && raw && !isNaN(parseFloat(raw))) {
                    vars.set(key, raw);
                }
            });
        }

        let unresolved = false;
        const resolved = template.replace(/\{(\w+)\}/g, (match, token: string) => {
            if (refs.has(token)) {
                return refs.get(token);
            }
            if (vars.has(token)) {
                return vars.get(token);
            }
            unresolved = true;
            return match;
        });

        if (unresolved || !this.isFormulaSafe(resolved)) {
            return null;
        }
        return resolved;
    }

    // The resolved formula text comes from the report's own measure, but it still lands
    // directly in an .xlsx a user opens - so before trusting it, require every character to
    // tokenize as a cell reference, a number, one of a small whitelist of Excel functions, or
    // a structural character (parens/operators/punctuation). Anything else (an unresolved
    // {token}, a stray function name, an external reference) is rejected rather than written,
    // since Excel formulas have known injection patterns (e.g. DDE, external links) that a
    // plain sanitize-the-string approach wouldn't catch. Uses a sticky-flag tokenizer instead
    // of one big regex so a long/adversarial input can't cause pathological backtracking.
    private static readonly FORMULA_TOKEN = /\$?[A-Z]{1,3}\$?\d{1,7}|\d+(?:\.\d+)?|SUM|AVERAGE|MIN|MAX|ROUND|[+\-*/(),.\s]/giy;

    private isFormulaSafe(formula: string): boolean {
        if (!formula || formula.length > 200) {
            return false;
        }
        const pattern = Visual.FORMULA_TOKEN;
        pattern.lastIndex = 0;
        let pos = 0;
        while (pos < formula.length) {
            pattern.lastIndex = pos;
            const match = pattern.exec(formula);
            if (!match || match.index !== pos) {
                return false;
            }
            pos += match[0].length;
        }
        return true;
    }

    // If `targetValue` equals the sum of a contiguous run of numeric cells already written
    // earlier in the same row (e.g. Total = Rent + Opex + Fitout), returns a SUM formula
    // referencing those cells instead of the caller having to hardcode which columns to add.
    private tryHorizontalSumFormula(
        rowNumericCells: { col: number; value: number }[],
        targetCol: number,
        targetValue: number,
        excelRow: number
    ): string {
        const candidates = rowNumericCells.filter((c) => c.col < targetCol);
        if (candidates.length < 2) {
            return null;
        }
        for (let i = 0; i < candidates.length - 1; i++) {
            if (candidates[i + 1].col !== candidates[i].col + 1) {
                return null; // not a contiguous run of columns
            }
        }
        const sum = candidates.reduce((s, c) => s + c.value, 0);
        if (Math.abs(sum - targetValue) > 0.5) {
            return null;
        }
        const first = this.colLetter(candidates[0].col);
        const last = this.colLetter(candidates[candidates.length - 1].col);
        return `SUM(${first}${excelRow}:${last}${excelRow})`;
    }

    private colLetter(col: number): string {
        let letter = "";
        let n = col;
        while (n > 0) {
            const rem = (n - 1) % 26;
            letter = String.fromCharCode(65 + rem) + letter;
            n = Math.floor((n - 1) / 26);
        }
        return letter;
    }

    // Word/.docx (a zip container) isn't in the file-download API's supported extension
    // list - the host silently refuses anything outside .txt/.csv/.json/.tmplt/.xml/.pdf/
    // .xlsx (confirmed: a .doc filename came back downloadCompleted:false, no error, just
    // refused). So this builds a legacy WordprocessingML document - a plain-text .xml
    // dialect Word opens natively with full table/color/bold fidelity - and saves it with
    // the .xml extension the API actually allows. Trade-off: double-clicking it in Explorer
    // opens a browser/XML viewer, not Word directly; open it via Word > File > Open instead.
    private async exportWord(): Promise<void> {
        try {
            const tables = this.getTables();
            if (tables.length === 0) {
                this.showStatus("No table found in the content to export.", true);
                return;
            }
            if (!(await this.checkExportAllowed())) {
                return;
            }

            const xml = this.buildWordXml(tables);
            const base64 = btoa(unescape(encodeURIComponent(xml)));
            const fileName = `${this.baseFileName()}.xml`;

            const result = await this.downloadService.exportVisualsContentExtended(base64, fileName, "base64", "Word export");
            if (result.downloadCompleted) {
                this.showStatus("Downloaded. In Word, use File > Open to open it.", false);
            } else {
                this.showStatus("Word download was cancelled or blocked.", true);
            }
        } catch (err) {
            console.error("wolcHtmlExport: Word export failed", err);
            this.showStatus("Word export failed - see console for details.", true);
        }
    }

    private buildWordXml(tables: HTMLTableElement[]): string {
        const parts: string[] = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<?mso-application progid="Word.Document"?>',
            '<w:wordDocument xmlns:w="http://schemas.microsoft.com/office/word/2003/wordml" xml:space="preserve">',
            "<w:body>"
        ];

        tables.forEach((table: HTMLTableElement, index: number) => {
            const title = this.titleFor(table, index);
            parts.push(
                "<w:p><w:pPr><w:spacing w:before=\"240\" w:after=\"120\"/></w:pPr><w:r>" +
                "<w:rPr><w:b/><w:sz w:val=\"28\"/><w:rFonts w:ascii=\"Arial\" w:hAnsi=\"Arial\"/></w:rPr>" +
                `<w:t>${this.escapeXml(title)}</w:t></w:r></w:p>`
            );
            parts.push(this.tableToWordXml(table));
        });

        parts.push("</w:body>", "</w:wordDocument>");
        return parts.join("");
    }

    private tableToWordXml(table: HTMLTableElement): string {
        const totalWidthTwips = 9350;
        const rows = Array.from(table.rows);
        const columnCount = Math.max(1, ...rows.map((row: HTMLTableRowElement) =>
            Array.from(row.cells).reduce((sum: number, cell: HTMLTableCellElement) => sum + (cell.colSpan || 1), 0)
        ));
        const colWidth = Math.floor(totalWidthTwips / columnCount);
        const gridCols = new Array(columnCount).fill(`<w:gridCol w:w="${colWidth}"/>`).join("");

        const rowsXml = rows.map((row: HTMLTableRowElement) => {
            const cellsXml = Array.from(row.cells).map((cell: HTMLTableCellElement) => {
                const colSpan = cell.colSpan || 1;
                const style = this.readCellStyle(cell);
                const width = colWidth * colSpan;
                const shd = style.backgroundHex ? `<w:shd w:val="clear" w:color="auto" w:fill="${style.backgroundHex}"/>` : "";
                const gridSpan = colSpan > 1 ? `<w:gridSpan w:val="${colSpan}"/>` : "";
                const jc = style.align !== "left" ? `<w:jc w:val="${style.align}"/>` : "";
                const bold = style.bold ? "<w:b/>" : "";
                const italic = style.italic ? "<w:i/>" : "";
                const text = this.escapeXml((cell.textContent || "").trim());
                return "<w:tc><w:tcPr>" +
                    `<w:tcW w:w="${width}" w:type="dxa"/>${gridSpan}${shd}` +
                    "</w:tcPr><w:p>" +
                    `<w:pPr>${jc}</w:pPr>` +
                    `<w:r><w:rPr>${bold}${italic}<w:color w:val="${style.colorHex}"/>` +
                    `<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="16"/></w:rPr>` +
                    `<w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:tc>`;
            }).join("");
            return `<w:tr>${cellsXml}</w:tr>`;
        }).join("");

        const borderAttrs = 'w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"';
        return "<w:tbl><w:tblPr><w:tblBorders>" +
            `<w:top ${borderAttrs}/><w:left ${borderAttrs}/><w:bottom ${borderAttrs}/><w:right ${borderAttrs}/>` +
            `<w:insideH ${borderAttrs}/><w:insideV ${borderAttrs}/>` +
            "</w:tblBorders></w:tblPr>" +
            `<w:tblGrid>${gridCols}</w:tblGrid>${rowsXml}</w:tbl><w:p/>`;
    }

    private escapeXml(text: string): string {
        return text
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&apos;");
    }

    // Recovers a real number (and a matching Excel number format) from DAX's own
    // FORMAT()-produced strings like "$200,000", so totals stay summable in Excel
    // instead of importing as text that merely looks like a currency value.
    private parseCellValue(text: string): { value: string | number; numFmt?: string } {
        const trimmed = text.trim();
        const match = trimmed.match(/^(-?)\$([\d,]+)(\.\d+)?$/);
        if (match) {
            const numeric = parseFloat(`${match[1]}${match[2].replace(/,/g, "")}${match[3] || ""}`);
            return { value: numeric, numFmt: '"$"#,##0' };
        }
        return { value: trimmed };
    }

    private readCellStyle(cell: HTMLElement): ParsedCellStyle {
        const computed = window.getComputedStyle(cell);
        const weight = parseInt(computed.fontWeight, 10);
        let backgroundHex = this.rgbToHex(computed.backgroundColor);
        if (!backgroundHex && cell.parentElement) {
            // Row-level classes like .stripe/.future set background-color on the <tr>, which
            // does NOT carry into a <td>'s own computed style (background-color isn't an
            // inherited CSS property) - fall back to the row so striping/highlighting survives.
            backgroundHex = this.rgbToHex(window.getComputedStyle(cell.parentElement).backgroundColor);
        }
        return {
            backgroundHex,
            colorHex: this.rgbToHex(computed.color) || "000000",
            bold: computed.fontWeight === "bold" || computed.fontWeight === "bolder" || (!isNaN(weight) && weight >= 700),
            italic: computed.fontStyle === "italic",
            align: this.excelAlign(computed.textAlign),
            borderHex: this.rgbToHex(computed.borderTopColor)
        };
    }

    private excelAlign(textAlign: string): "left" | "center" | "right" {
        if (textAlign === "right" || textAlign === "center" || textAlign === "left") {
            return textAlign;
        }
        return "left";
    }

    private rgbToHex(rgb: string): string {
        if (!rgb) {
            return null;
        }
        const match = rgb.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?\)/);
        if (!match) {
            return null;
        }
        const alpha = match[4] !== undefined ? parseFloat(match[4]) : 1;
        if (alpha === 0) {
            return null;
        }
        const toHex = (n: string) => parseInt(n, 10).toString(16).padStart(2, "0").toUpperCase();
        return `${toHex(match[1])}${toHex(match[2])}${toHex(match[3])}`;
    }

    private arrayBufferToBase64(buffer: ArrayBuffer): string {
        const bytes = new Uint8Array(buffer);
        const chunkSize = 0x8000;
        let binary = "";
        for (let i = 0; i < bytes.length; i += chunkSize) {
            binary += String.fromCharCode(...Array.from(bytes.subarray(i, i + chunkSize)));
        }
        return btoa(binary);
    }

    private titleFor(table: HTMLTableElement, index: number): string {
        let label: string = null;
        let node: Element = table.previousElementSibling;
        while (node && !label) {
            const text = (node.textContent || "").trim();
            if (text && (node.tagName === "H1" || node.tagName === "H2" || node.tagName === "H3" ||
                node.classList.contains("section-header") || node.classList.contains("dark-header"))) {
                label = text;
            }
            node = node.previousElementSibling;
        }
        return label || `Table ${index + 1}`;
    }

    private sanitizeSheetName(name: string): string {
        return name.replace(/[\\/:*?[\]]/g, "").substring(0, 31) || "Sheet";
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }
}
