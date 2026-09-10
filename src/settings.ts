"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";

import FormattingSettingsCard = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

/**
 * Toolbar formatting card - controls the export buttons overlaid on the HTML content
 */
class ToolbarCardSettings extends FormattingSettingsCard {
    show = new formattingSettings.ToggleSwitch({
        name: "show",
        displayName: "Show export buttons",
        value: true
    });

    position = new formattingSettings.ItemDropdown({
        name: "position",
        displayName: "Button position",
        items: [
            { displayName: "Bottom right", value: "bottom-right" },
            { displayName: "Bottom left", value: "bottom-left" },
            { displayName: "Top left", value: "top-left" },
            { displayName: "Top right", value: "top-right" }
        ],
        value: { displayName: "Bottom right", value: "bottom-right" }
    });

    pdfButtonText = new formattingSettings.TextInput({
        name: "pdfButtonText",
        displayName: "PDF button text",
        placeholder: "Export PDF",
        value: "Export PDF"
    });

    excelButtonText = new formattingSettings.TextInput({
        name: "excelButtonText",
        displayName: "Excel button text",
        placeholder: "Export Excel",
        value: "Export Excel"
    });

    wordButtonText = new formattingSettings.TextInput({
        name: "wordButtonText",
        displayName: "Word button text",
        placeholder: "Export Word",
        value: "Export Word"
    });

    fileName = new formattingSettings.TextInput({
        name: "fileName",
        displayName: "Export file name",
        placeholder: "Export",
        value: "Export"
    });

    name: string = "toolbar";
    displayName: string = "Export toolbar";
    slices: Array<FormattingSettingsSlice> = [this.show, this.position, this.pdfButtonText, this.excelButtonText, this.wordButtonText, this.fileName];
}

/**
* visual settings model class
*/
export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    toolbarCard = new ToolbarCardSettings();

    cards = [this.toolbarCard];
}
