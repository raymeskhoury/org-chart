import "@shoelace-style/shoelace/dist/components/button/button";
import "@shoelace-style/shoelace/dist/components/color-picker/color-picker";
import "@shoelace-style/shoelace/dist/components/dialog/dialog";
import "@shoelace-style/shoelace/dist/components/dropdown/dropdown";
import SlDropdown from "@shoelace-style/shoelace/dist/components/dropdown/dropdown";
import "@shoelace-style/shoelace/dist/components/icon/icon";
import "@shoelace-style/shoelace/dist/components/menu-item/menu-item";
import "@shoelace-style/shoelace/dist/components/menu/menu";
import "@shoelace-style/shoelace/dist/themes/dark.css";
import "@shoelace-style/shoelace/dist/themes/light.css";
import {setBasePath} from "@shoelace-style/shoelace/dist/utilities/base-path";
import {LitElement, css, html} from "lit";
import {customElement} from "lit/decorators.js";
import {configure} from "mobx";
import Papa from "papaparse";
import "./DataDialogElement";
import {DataDialogElement} from "./DataDialogElement";
import {DialogElement, DialogElementType} from "./DialogElement";
import {OrgChartDataModel, OrgChartEntry} from "./OrgChartDataModel";
import {OrgChartElement, OrgChartElementExportType} from "./OrgChartElement";
import {Util} from "./Util";
import "./style.css";

configure({
  enforceActions: "always",
  computedRequiresReaction: true,
  disableErrorBoundaries: true,
});

const href = window.location.href;
const dir = href.substring(0, href.lastIndexOf("/")) + "/";
setBasePath(dir);

let isInOffice = false;
let isInOfficeDesktop = false;
let chartElement: OrgChartElement;
let dialog: DialogElement;
let dataDialog: DataDialogElement;

export const COLUMN_NAMES = ["id", "parentId", "name", "position", "color"];

const bgColors = ["#0d2747", "#97daff", "#fffffe"];

const templateData = [
  ["1", "", "Bianca Toscano", "Director", ""],
  ["2", "1", "Aasa Andrejev", "Manager, Marketing", ""],
  ["3", "1", "Paul Lohmus", "Manager, Products", ""],
  ["4", "2", "Sergio Udinese", "PR Coordinator", ""],
  ["5", "2", "Mattia Sabbatini", "Content Strategist", ""],
  ["6", "3", "Mai Aare", "Engineering Lead", ""],
  ["7", "3", "Aet Kangro", "Design Lead", ""],
  ["8", "4", "Aili Mihhailov", "PR Specialist", ""],
  ["9", "4", "Lemme Kangur", "PR Assistant", ""],
  ["10", "5", "Alice Cattaneo", "Copywriter", ""],
  ["11", "6", "Helbe Piip", "Software Engineer", ""],
  ["12", "6", "Riccardo Buccho", "Intern", ""],
  ["13", "7", "Jana Piip", "UX Designer", ""],
];

@customElement("org-chart-toolbar")
export class ToolbarElement extends LitElement {
  static override styles = css`
    .toolbar {
      display: flex;
      padding: 2px;
      column-gap: 5px;
    }

    .menu {
      min-width: 110px;
    }
  `;
  private dropdown?: SlDropdown;

  protected override firstUpdated(): void {
    this.dropdown = this.shadowRoot!.getElementById(
      "exportDropdown"
    ) as SlDropdown;
  }

  hideMenus(): void {
    this.dropdown!.hide();
  }

  protected override render(): unknown {
    return html`<div class="toolbar">
      ${isInOffice
        ? html`<sl-button
            size="small"
            @click=${() => {
              updateTable();
            }}
          >
            <sl-icon slot="prefix" name="arrow-clockwise"></sl-icon>
            Update
          </sl-button>`
        : html`<sl-button
            size="small"
            @click=${() => {
              renderFromInput();
            }}
          >
            <sl-icon slot="prefix" name="pencil"></sl-icon>
            Edit data
          </sl-button>`}

      <sl-dropdown id="exportDropdown">
        <sl-button size="small" slot="trigger" caret>
          <sl-icon slot="prefix" name="download"></sl-icon>
          Export
        </sl-button>
        <sl-menu class="menu">
          <sl-menu-item
            id="exportPng"
            @click=${() => {
              chartElement.exportImage(
                OrgChartElementExportType.Png,
                isInOfficeDesktop
              );
            }}
            >PNG</sl-menu-item
          >
          <sl-menu-item
            id="exportSvg"
            @click=${() => {
              chartElement.exportImage(
                OrgChartElementExportType.Svg,
                isInOfficeDesktop
              );
            }}
            >SVG</sl-menu-item
          >
          <sl-menu-item
            id="exportPdf"
            @click=${async () => {
              const result = await dialog.show(
                "Export PDF",
                'To export to PDF, select "Print" and select ' +
                  '"Save as PDF" in your browser\'s print options.',
                DialogElementType.OK_CANCEL,
                "",
                "Print",
                "Close"
              );
              if (result) {
                chartElement.exportImage(
                  OrgChartElementExportType.Print,
                  isInOfficeDesktop
                );
              }
            }}
            >PDF...</sl-menu-item
          >
        </sl-menu>
      </sl-dropdown>

      <sl-button
        size="small"
        id="printButton"
        @click=${() => {
          chartElement.exportImage(
            OrgChartElementExportType.Print,
            isInOfficeDesktop
          );
        }}
      >
        <sl-icon slot="prefix" name="printer"></sl-icon>
        Print
      </sl-button>
      <sl-button
        size="small"
        id="fitButton"
        @click=${() => {
          chartElement.fit();
        }}
      >
        <sl-icon slot="prefix" name="arrows-angle-expand"></sl-icon>
        Fit
      </sl-button>
      <sl-button
        size="small"
        id="expandButton"
        @click=${() => {
          chartElement.expandAll();
        }}
      >
        <sl-icon slot="prefix" name="plus-lg"></sl-icon>
        Expand all
      </sl-button>
    </div>`;
  }
}

document.addEventListener("DOMContentLoaded", async () => {
  dialog = document.getElementById("dialog")! as DialogElement;
  dataDialog = new DataDialogElement(new OrgChartDataModel());
  await Util.appendChildAndWait(
    document.getElementById("dataDialog")!,
    dataDialog
  );

  chartElement = document.getElementById("orgChart") as OrgChartElement;
  const toolbar = document.getElementById("toolbar") as ToolbarElement;
  chartElement.addEventListener(
    "mousedown",
    () => {
      toolbar.hideMenus();
    },
    true
  );

  Office.onReady(async (info): Promise<void> => {
    isInOffice =
      info.host === Office.HostType.Excel ||
      window.location.search === "?excelDesktop" ||
      window.location.search === "?excelOnline";
    isInOfficeDesktop =
      isInOffice &&
      (info.platform !== Office.PlatformType.OfficeOnline ||
        window.location.search === "?excelDesktop");
    if (isInOffice) {
      let csv = "";

      if (!(await tableExists())) {
        let error = "";
        do {
          await dialog.show(
            "Insert template",
            "To get started, insert a template table for entering your org " +
              "chart data. The table will be added at the currently selected " +
              "cell.",
            DialogElementType.OK,
            error,
            "Insert",
            "",
            true
          );
          error = await insertTemplate();
        } while (error.length !== 0);
      }
      csv = await readTable();

      const parsed = Papa.parse(csv!, {
        header: true,
      });
      await renderChart(parsed.data as OrgChartEntry[]);
    } else {
      await renderFromInput();
    }
  });
});

async function renderFromInput(): Promise<void> {
  let error = "";
  let data = [];
  do {
    data = await dataDialog.show(error);
    error = chartElement.setData(data);
    if (error !== "") {
      error =
        "Error: Failed to render chart. Each row must have a unique ID " +
        "field and a valid manager ID. The head of the organisation should " +
        `have an empty manager ID (problem: ${error})`;
    }
  } while (error.length !== 0);
}

async function updateTable(): Promise<void> {
  if (!(await tableExists())) {
    document.location.reload();
    return;
  }

  const csv = await readTable();
  const parsed = Papa.parse(csv!, {
    header: true,
  });
  setTimeout(() => {
    renderChart(parsed.data as OrgChartEntry[]);
  });
}

async function renderChart(data: OrgChartEntry[]): Promise<void> {
  const error = chartElement.setData(data);
  if (error !== "") {
    await dialog.show(
      "Error updating chart",
      "There is a problem with the data in your chart. Each row should have " +
        'a unique "ID" and a "Manager ID" that corresponds to a unique ID of' +
        'another entry in the data. Exactly one person should have no "ID" ' +
        "which indicates they are the head of the organisation.",
      DialogElementType.OK,
      "Error: " + error,
      "Retry"
    );
    updateTable();
  }
}

async function tableExists(): Promise<boolean> {
  let result = false;
  await Excel.run(async context => {
    const tableName = Office.context.document.settings.get("table");

    if (tableName === null) {
      result = false;
      return;
    }

    try {
      const table = context.workbook.tables.getItem(tableName);
      await context.sync();
      if (table !== null) {
        result = true;
        return;
      }
    } catch (e) {
      result = false;
      return;
    }
  });

  return result;
}

function excelValueToString(
  value: any,
  type: Excel.RangeValueType | undefined
): string {
  if (type === undefined) {
    return "";
  }
  switch (type) {
    case Excel.RangeValueType.boolean:
    case Excel.RangeValueType.double:
    case Excel.RangeValueType.empty:
    case Excel.RangeValueType.integer:
    case Excel.RangeValueType.string:
      return String(value);
  }
  return "";
}

async function readTable(): Promise<string> {
  let result = "";
  await Excel.run(async context => {
    const tableName = Office.context.document.settings.get("table");
    let table;
    if (tableName === null) {
      result = "";
      return;
    }

    try {
      table = context.workbook.tables.getItem(tableName);
      await context.sync();
    } catch (e) {
      result = "";
      return;
    }

    const bodyRange = table
      .getDataBodyRange()
      .load(["values", "valueTypes", "rowCount", "columnCount"]);

    await context.sync();

    const colors = [];
    for (let i = 0; i < bodyRange.rowCount; ++i) {
      if (bodyRange.columnCount >= COLUMN_NAMES.length) {
        const range = bodyRange
          .getCell(i, bodyRange.columnCount - 1)
          .load("format/fill/color");
        colors.push(range);
      } else {
        colors.push(undefined);
      }
    }

    await context.sync();

    const bodyValues = bodyRange.values;
    const bodyTypes = bodyRange.valueTypes;

    const entries: OrgChartEntry[] = [];
    for (let i = 0; i < bodyValues.length; ++i) {
      if (bodyValues[i].length < COLUMN_NAMES.length) {
        bodyValues[i].length = COLUMN_NAMES.length;
        bodyTypes[i].length = COLUMN_NAMES.length;
      }
      const entry: OrgChartEntry = {
        id: excelValueToString(bodyValues[i][0], bodyTypes[i][0]),
        parentId: excelValueToString(bodyValues[i][1], bodyTypes[i][1]),
        name: excelValueToString(bodyValues[i][2], bodyTypes[i][2]),
        position: excelValueToString(bodyValues[i][3], bodyTypes[i][3]),
        color:
          colors[i] === undefined ? "#FFFFFF" : colors[i]!.format.fill.color,
      };
      if (entry.name === "") {
        continue;
      }
      entries.push(entry);
    }
    const csv = Papa.unparse(entries);
    result = csv;
  });
  return result;
}

async function insertTemplate(): Promise<string> {
  let result = "";
  await Excel.run(async context => {
    let range: Excel.Range;
    try {
      range = context.workbook.getSelectedRange();
    } catch (e) {
      console.error(e);
      result =
        "Error: Select an indiviudal cell where the template data should be inserted.";
      return;
    }
    range.load(["cellCount", "worksheet"]);
    const sheet = range!.worksheet;

    await context.sync();

    if (range.cellCount !== 1) {
      result =
        "Error: Select an indiviudal cell where the template data should be inserted.";
      return;
    }

    const headerRange = range.getCell(0, 0).getResizedRange(0, 4);
    const dataRange = headerRange.getResizedRange(templateData.length, 0);
    dataRange.load("valueTypes");
    await context.sync();
    for (let i = 0; i < dataRange.valueTypes.length; ++i) {
      for (let j = 0; j < dataRange.valueTypes[i].length; ++j) {
        if (dataRange.valueTypes[i][j] !== Excel.RangeValueType.empty) {
          result = "Error: Not enough space to insert template data.";
          return;
        }
      }
    }

    let table: Excel.Table;
    try {
      table = sheet.tables.add(dataRange, true /*hasHeaders*/);
      table.getHeaderRowRange().values = [
        ["ID", "Manager ID", "Name", "Position", "Background Colour"],
      ];
      const body = table.getDataBodyRange();
      body.values = templateData;

      table.getRange().format.fill.clear();

      for (let i = 0; i < templateData.length; ++i) {
        for (let j = 0; j < templateData[i].length; ++j) {
          const colorIndex = Math.floor(Math.random() * bgColors.length);
          if (j === templateData[i].length - 1) {
            body.getCell(i, j).format.fill.color = bgColors[colorIndex];
          }
        }
      }

      if (
        Office.context.requirements.isSetSupported("ExcelApi", "1.2") === true
      ) {
        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();
      }
      table.load("name");
      await context.sync();
    } catch (e) {
      console.error(e);
      result = "Error: template data cannot overlap another table.";
      return;
    }

    Office.context.document.settings.set("table", table.name);
    Office.context.document.settings.saveAsync();
    result = "";
    return;
  });
  return result;
}
