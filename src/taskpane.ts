import "@shoelace-style/shoelace/dist/components/button/button";
import "@shoelace-style/shoelace/dist/components/color-picker/color-picker";
import "@shoelace-style/shoelace/dist/components/dialog/dialog";
import SlDialog from "@shoelace-style/shoelace/dist/components/dialog/dialog";
import "@shoelace-style/shoelace/dist/components/dropdown/dropdown";
import SlDropdown from "@shoelace-style/shoelace/dist/components/dropdown/dropdown";
import "@shoelace-style/shoelace/dist/components/icon/icon";
import "@shoelace-style/shoelace/dist/components/menu-item/menu-item";
import "@shoelace-style/shoelace/dist/components/menu/menu";
import "@shoelace-style/shoelace/dist/themes/dark.css";
import "@shoelace-style/shoelace/dist/themes/light.css";
import {setBasePath} from "@shoelace-style/shoelace/dist/utilities/base-path";
import * as d3 from "d3";
import {OrgChart} from "d3-org-chart";
import Papa from "papaparse";
import "./style.css";

const href = window.location.href;
const dir = href.substring(0, href.lastIndexOf("/")) + "/";
setBasePath(dir);

enum ExportType {
  Svg,
  Png,
  Print,
}

let chart: OrgChart<unknown>;
let exportPdfDialog: SlDialog;
const fontSizes = new Map();
let isInOffice = false;
let isInOfficeDesktop = false;
let containerElement: HTMLElement;
let templateDataDialog: SlDialog;
let templateDataDialogError: HTMLDivElement;
let chartErrorDialog: SlDialog;
let chartErrorDialogMessage: HTMLDivElement;

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

const columnNames = ["id", "parentId", "name", "position", "color"];

interface OrgEntry {
  id: string;
  parentId: string;
  name: string;
  position: string;
  color: string;
}

async function showDialogAndWait(dialog: SlDialog): Promise<void> {
  const promise = new Promise<void>(resolve => {
    function onHide(): void {
      resolve();
      dialog.removeEventListener("sl-after-hide", onHide);
    }
    dialog.addEventListener("sl-after-hide", onHide);
    dialog.show();
  });
  return promise;
}

function pickTextColorBasedOnBgColor(
  bgColor: string,
  lightColor: string,
  darkColor: string
): string {
  const color = bgColor.charAt(0) === "#" ? bgColor.substring(1, 7) : bgColor;
  const r = parseInt(color.substring(0, 2), 16); // hexToR
  const g = parseInt(color.substring(2, 4), 16); // hexToG
  const b = parseInt(color.substring(4, 6), 16); // hexToB
  const uicolors = [r / 255, g / 255, b / 255];
  const c = uicolors.map(col => {
    if (col <= 0.03928) {
      return col / 12.92;
    }
    return Math.pow((col + 0.055) / 1.055, 2.4);
  });
  const L = 0.2126 * c[0] + 0.7152 * c[1] + 0.0722 * c[2];
  return L > 0.179 ? darkColor : lightColor;
}

document.addEventListener("DOMContentLoaded", () => {
  chartErrorDialog = document.getElementById("chartErrorDialog")! as SlDialog;
  chartErrorDialogMessage = document.getElementById(
    "chartErrorDialogMessage"
  )! as HTMLDivElement;
  exportPdfDialog = document.getElementById("exportPdfDialog")! as SlDialog;

  templateDataDialog = document.getElementById(
    "templateDataDialog"
  )! as SlDialog;
  templateDataDialog.addEventListener("sl-request-close", event => {
    const e = event as CustomEvent;
    if (e.detail.source === "overlay" || e.detail.source === "close-button") {
      e.preventDefault();
    }
  });
  document
    .getElementById("templateDataDialogInsertButton")!
    .addEventListener("click", () => {
      templateDataDialog.hide();
    });
  document
    .getElementById("chartErrorDialogRetryButton")!
    .addEventListener("click", () => {
      chartErrorDialog.hide();
    });

  templateDataDialogError = document.getElementById(
    "templateDataDialogError"
  )! as HTMLDivElement;

  containerElement = document.getElementsByClassName(
    "chart-container"
  )[0]! as HTMLElement;
  const exportDropdown = document.getElementById(
    "exportDropdown"
  ) as SlDropdown;
  containerElement.addEventListener(
    "mousedown",
    () => {
      exportDropdown.hide();
    },
    true
  );
  document
    .getElementById("exportPdfDialogPrintButton")!
    .addEventListener("click", () => {
      exportPdfDialog.hide();
      exportImage(ExportType.Print);
    });
  document
    .getElementById("exportPdfDialogCloseButton")!
    .addEventListener("click", () => {
      exportPdfDialog.hide();
    });

  document.getElementById("updateButton")!.addEventListener("click", () => {
    updateTable();
  });
  document.getElementById("printButton")!.addEventListener("click", () => {
    exportImage(ExportType.Print);
  });
  document.getElementById("fitButton")!.addEventListener("click", () => {
    chart.render();
    chart.fit();
  });
  document.getElementById("expandButton")!.addEventListener("click", () => {
    chart.expandAll();
  });
  document.getElementById("exportPng")!.addEventListener("click", () => {
    exportImage(ExportType.Png);
  });
  document.getElementById("exportSvg")!.addEventListener("click", () => {
    exportImage(ExportType.Svg);
  });
  document.getElementById("exportPdf")!.addEventListener("click", () => {
    exportPdfDialog.show();
  });

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
        await showDialogAndWait(templateDataDialog);
        let error = await insertTemplate();
        console.error(error);
        while (error.length !== 0) {
          console.error("test");
          templateDataDialogError.innerText = error;
          await showDialogAndWait(templateDataDialog);
          error = await insertTemplate();
        }
      }
      csv = await readTable();

      const parsed = Papa.parse(csv!, {
        header: true,
      });
      drawChart(parsed.data as OrgEntry[]);
    }
  });
});

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
    renderChart(parsed.data as OrgEntry[]);
  });
}

async function renderChart(data: OrgEntry[]): Promise<void> {
  try {
    chart.data(data).expandAll().compact(true).fit().render();
  } catch (e) {
    chartErrorDialogMessage.innerText = "Error: " + (e as Error).message;
    await showDialogAndWait(chartErrorDialog);
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
      if (bodyRange.columnCount >= columnNames.length) {
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

    const entries: OrgEntry[] = [];
    for (let i = 0; i < bodyValues.length; ++i) {
      if (bodyValues[i].length < columnNames.length) {
        bodyValues[i].length = columnNames.length;
        bodyTypes[i].length = columnNames.length;
      }
      const entry: OrgEntry = {
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

    const tableRange = range.getCell(0, 0).getResizedRange(0, 4);
    const dataRange = tableRange.getResizedRange(templateData.length, 0);
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
      table = sheet.tables.add(tableRange, true /*hasHeaders*/);
      table.getHeaderRowRange().values = [
        ["ID", "Manager ID", "Name", "Position", "Background Colour"],
      ];
      table.rows.add(
        undefined /*add rows to the end of the table*/,
        templateData
      );

      table.getRange().format.fill.clear();
      const body = table.getDataBodyRange();

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

function doesOverflow(el: HTMLElement): boolean {
  return el.clientWidth < el.scrollWidth || el.clientHeight < el.scrollHeight;
}

function saveAs(uri: string, filename: string): void {
  if (isInOfficeDesktop) {
    const src = `
    <html>
    <body style="background-color: white">
    <img style="width: 100%;" src=${uri} />
    <a download="${filename}" href="${uri}">Download</a>
    </body>
    </html>
    `;
    const temp = window.open(document.location.href);
    if (temp) {
      temp.document.write(src);
      temp.document.close();
      temp.focus();
    }
  } else {
    const link = document.createElement("a");
    if (typeof link.download === "string") {
      document.body.appendChild(link);
      link.download = filename;
      link.href = uri;
      link.click();
      document.body.removeChild(link);
    } else {
      location.replace(uri);
    }
  }
}

function exportPng(url: string, width: number, height: number): void {
  const targetWidth = 3000;
  const scale = targetWidth / width;
  const image = document.createElement("img");
  image.onload = function () {
    // Create image canvas
    const canvas = document.createElement("canvas");
    // Set width and height based on SVG node
    canvas.width = width * scale;
    canvas.height = height * scale;
    // Draw background
    const context = canvas.getContext("2d")!;
    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, width * scale, height * scale);
    context.drawImage(image, 0, 0, width * scale, height * scale);

    saveAs(canvas.toDataURL("image/png"), "orgchart.png");
  };

  image.src = url;
}

function print(url: string): void {
  const src =
    `
      <html>
      <head>
      <style>
        @media print {
          @page {size: landscape}
          .svg-chart-container {
            width:100%;
            height: 100%;
            page-break-after:always
          }
        }
    }
      </style>
      <script>
        window.addEventListener("load", () => {
          print();
          window.close();
        });
      </script>
      </head>
      <body>
        <img class="svg-chart-container" src="
      ` +
    url +
    '" /></body></html>';

  if (isInOfficeDesktop) {
    const temp = window.open(document.location.href);
    if (temp) {
      temp.document.write(src);
      temp.document.close();
      temp.focus();
    }
  } else {
    const frame = document.getElementById("print")! as HTMLIFrameElement;
    frame.srcdoc = src;
  }
}

async function exportImage(type: ExportType): Promise<void> {
  const svg = document.getElementsByClassName("svg-chart-container")[0];

  const xmlns = "http://www.w3.org/2000/xmlns/";
  const xlinkns = "http://www.w3.org/1999/xlink";
  const svgns = "http://www.w3.org/2000/svg";

  const oldDuration = chart.duration();
  chart.duration(0);
  chart.expandAll();
  chart.fit({animate: false});

  setTimeout(() => {
    const elements = svg.getElementsByClassName("node-foreign-object-div");
    let minX;
    let minY;
    let maxX;
    let maxY;
    for (const element of elements) {
      const rect = element.getBoundingClientRect();
      if (minX === undefined || rect.x < minX) {
        minX = rect.x - 5;
      }
      if (minY === undefined || rect.y < minY) {
        minY = rect.y - 5;
      }
      if (maxX === undefined || rect.right > maxX) {
        maxX = rect.right + 10;
      }
      if (maxY === undefined || rect.bottom > maxY) {
        maxY = rect.bottom + 10;
      }
    }
    const svgRect = svg.getBoundingClientRect();

    minX = minX! - svgRect.x;
    minY = minY! - svgRect.y;
    maxX = maxX! - svgRect.x;
    maxY = maxY! - svgRect.y;
    console.error(minX + " " + maxX + " " + minY + " " + maxY);
    const width = maxX - minX;
    const height = maxY - minY;

    const cloned = svg.cloneNode(true) as HTMLElement;
    cloned.setAttributeNS(xmlns, "xmlns", svgns);
    cloned.setAttributeNS(xmlns, "xmlns:xlink", xlinkns);

    cloned.setAttribute(
      "viewBox",
      minX + " " + minY + " " + width + " " + height
    );
    cloned.setAttribute("width", String(width));
    cloned.setAttribute("height", String(height));
    const serializer = new XMLSerializer();
    let src = serializer.serializeToString(cloned);
    src = '<?xml version="1.0" standalone="no"?>\r\n' + src;

    const url = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(src);

    switch (type) {
      case ExportType.Png:
        exportPng(url, width, height);
        break;
      case ExportType.Svg:
        saveAs(url, "orgchart.svg");
        break;
      case ExportType.Print:
        print(url);
        break;
    }

    chart.duration(oldDuration);
  });
}

function drawChart(data: OrgEntry[]): void {
  console.error(data);
  d3.csv("./assets/data-oracle.csv").then(oldData => {
    OrgChart.prototype.diagonal = function (s, t, m, offsets = {sy: 0}) {
      const x = s.x;
      let y = s.y;

      const ex = t.x;
      const ey = t.y;

      const mx = m !== null && m !== undefined && m.x !== null ? m.x : x; // This is a changed line
      const my = m !== null && m !== undefined && m.y !== null ? m.y : y; // This also is a changed line

      const xrvs = ex - x < 0 ? -1 : 1;
      const yrvs = ey - y < 0 ? -1 : 1;

      y += offsets.sy;

      const rdef = 0;
      let r = Math.abs(ex - x) / 2 < rdef ? Math.abs(ex - x) / 2 : rdef;

      r = Math.abs(ey - y) / 2 < r ? Math.abs(ey - y) / 2 : r;

      const h = Math.abs(ey - y) / 2 - r;
      const w = Math.abs(ex - x) - r * 2;
      //w=0;
      const path = `
                M ${mx} ${my}
                L ${x} ${my}
                L ${x} ${y}
                L ${x} ${y + h * yrvs}
                C  ${x} ${y + h * yrvs + r * yrvs} ${x} ${
                  y + h * yrvs + r * yrvs
                } ${x + r * xrvs} ${y + h * yrvs + r * yrvs}
                L ${x + w * xrvs + r * xrvs} ${y + h * yrvs + r * yrvs}
                C  ${ex}  ${y + h * yrvs + r * yrvs} ${ex}  ${
                  y + h * yrvs + r * yrvs
                } ${ex} ${ey - h * yrvs}
                L ${ex} ${ey}
     `;
      return path;
    };
    chart = new OrgChart();

    chart
      .nodeHeight(d => 120)
      .nodeWidth(d => 160)
      .childrenMargin(d => 50)
      .compactMarginBetween(d => 50)
      .compactMarginPair(d => 50)
      .neighbourMargin((a, b) => 30)
      .linkUpdate(function (d, i, arr) {
        // @ts-ignore
        d3.select(this)
          .attr("stroke", d => "#000000")
          .attr("stroke-width", d => 2);
      })

      .nodeContent((d, i, arr, state) => {
        const a = d as any;
        let fontSize = fontSizes.get(d.id);
        const fontColor = pickTextColorBasedOnBgColor(
          a.data.color,
          "#FFFFFF",
          "#000000"
        );
        function nodeHtml(size: number): string {
          return `
                <div style='width:${a.width}px;height:${d.height}px;' >
                  <div style="display: flex; flex-direction:column; justify-content: center;  font-family: 'Inter', sans-serif;background-color:${
                    a.data.color
                  };  width:${a.width - 4}px; height: ${
                    d.height - 4
                  }px;border-radius:10px;border: 2px solid #000000; position:relative;">
                    <div style="padding: 5px 8px 5px 10px;overflow:hidden;">
                    <div style="font-size: ${size}px;color:${fontColor}; font-weight: 500;">  ${
                      a.data.name
                    } </div>
                    <div style="color:${fontColor};margin-top:3px;font-size:15px;"> ${
                      a.data.position
                    } </div>
                    </div>
                  </div>
                </div>`;
        }
        if (fontSize === undefined) {
          fontSize = 25;

          const resizer = document.getElementById("font-resizer");
          resizer!.innerHTML = nodeHtml(fontSize);
          while (
            doesOverflow(
              resizer!.firstElementChild!.firstElementChild!
                .firstElementChild! as HTMLElement
            ) &&
            fontSize > 5
          ) {
            fontSize--;
            resizer!.innerHTML = nodeHtml(fontSize);
          }
          fontSizes.set(d.id, fontSize);
        }

        return nodeHtml(fontSize);
      })
      .buttonContent(({node, state}) => {
        function top(d: boolean): string {
          return d
            ? `<div style="display:flex;"><span style="align-items:center;display:flex;"><svg width="8" height="8" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
              <path d="M11.457 8.07005L3.49199 16.4296C3.35903 16.569 3.28485 16.7543 3.28485 16.9471C3.28485 17.1398 3.35903 17.3251 3.49199 17.4646L3.50099 17.4736C3.56545 17.5414 3.64304 17.5954 3.72904 17.6324C3.81504 17.6693 3.90765 17.6883 4.00124 17.6883C4.09483 17.6883 4.18745 17.6693 4.27344 17.6324C4.35944 17.5954 4.43703 17.5414 4.50149 17.4736L12.0015 9.60155L19.4985 17.4736C19.563 17.5414 19.6405 17.5954 19.7265 17.6324C19.8125 17.6693 19.9052 17.6883 19.9987 17.6883C20.0923 17.6883 20.1849 17.6693 20.2709 17.6324C20.3569 17.5954 20.4345 17.5414 20.499 17.4736L20.508 17.4646C20.641 17.3251 20.7151 17.1398 20.7151 16.9471C20.7151 16.7543 20.641 16.569 20.508 16.4296L12.543 8.07005C12.4729 7.99653 12.3887 7.93801 12.2954 7.89801C12.202 7.85802 12.1015 7.8374 12 7.8374C11.8984 7.8374 11.798 7.85802 11.7046 7.89801C11.6113 7.93801 11.527 7.99653 11.457 8.07005Z" fill="#000000" stroke="#000000"/>
              </svg></span><span style="margin-left:1px;color:#000000">${
                (node.data as any)._directSubordinatesPaging
              } </span></div>
              `
            : `<div style="display:flex;"><span style="align-items:center;display:flex;"><svg width="8" height="8" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
              <path d="M19.497 7.98903L12 15.297L4.503 7.98903C4.36905 7.85819 4.18924 7.78495 4.002 7.78495C3.81476 7.78495 3.63495 7.85819 3.501 7.98903C3.43614 8.05257 3.38462 8.12842 3.34944 8.21213C3.31427 8.29584 3.29615 8.38573 3.29615 8.47653C3.29615 8.56733 3.31427 8.65721 3.34944 8.74092C3.38462 8.82463 3.43614 8.90048 3.501 8.96403L11.4765 16.74C11.6166 16.8765 11.8044 16.953 12 16.953C12.1956 16.953 12.3834 16.8765 12.5235 16.74L20.499 8.96553C20.5643 8.90193 20.6162 8.8259 20.6517 8.74191C20.6871 8.65792 20.7054 8.56769 20.7054 8.47653C20.7054 8.38537 20.6871 8.29513 20.6517 8.21114C20.6162 8.12715 20.5643 8.05112 20.499 7.98753C20.3651 7.85669 20.1852 7.78345 19.998 7.78345C19.8108 7.78345 19.6309 7.85669 19.497 7.98753V7.98903Z" fill="#000000" stroke="#000000"/>
              </svg></span><span style="margin-left:1px;color:#000000">${
                (node.data as any)._directSubordinatesPaging
              } </span></div>
          `;
        }
        return `<div style="border:1px solid #000000;border-radius:3px;padding: 3px 5px 1px 5px;font-size:12px;margin:auto auto;background-color:white"> ${top(
          node.children !== undefined && node.children !== null
        )}  </div>`;
      })
      .container(".chart-container")
      .svgHeight(containerElement.clientHeight);

    window.addEventListener("resize", () => {
      chart.svgHeight(containerElement.clientHeight);
    });

    renderChart(data);
  });
}
