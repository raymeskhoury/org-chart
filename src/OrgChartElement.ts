import * as d3 from "d3";
import {OrgChart} from "d3-org-chart";
import {elementToSVG} from "dom-to-svg";
import {LitElement, css, html} from "lit-element";
import {customElement} from "lit/decorators.js";
import {OrgChartEntry} from "./OrgChartDataModel";

function isColor(strColor: string): boolean {
  const s = new Option().style;
  s.color = strColor;
  return s.color !== "";
}

function exportPng(
  url: string,
  width: number,
  height: number,
  openNewWindow: boolean
): void {
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

    saveAs(canvas.toDataURL("image/png"), "orgchart.png", openNewWindow);
  };

  image.src = url;
}

function saveAs(uri: string, filename: string, openNewWindow: boolean): void {
  if (openNewWindow) {
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

function print(
  url: string,
  openNewWindow: boolean,
  iframe: HTMLIFrameElement,
  width: number,
  height: number,
  toPdf: boolean
): void {
  const src = `
      <html>
      <head>
      <style>
        @media print {
          @page {size: ${width > height ? "landscape" : "portrait"}; ${
            toPdf ? "margin: 0mm;" : ""
          }}
          .svg-chart-container {
            height: 100%;
            width: 100%;
            page-break-after:always
          }
          html, body {
            height: 98%;
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
        <img class="svg-chart-container" src="${url}" /></body></html>`;

  if (openNewWindow) {
    const temp = window.open(document.location.href);
    if (temp) {
      temp.document.write(src);
      temp.document.close();
      temp.focus();
    }
  } else {
    iframe.srcdoc = src;
  }
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

function doesOverflow(el: HTMLElement): boolean {
  return el.clientWidth < el.scrollWidth || el.clientHeight < el.scrollHeight;
}

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

export enum OrgChartElementExportType {
  Svg,
  Png,
  Print,
  Pdf,
}

@customElement("org-chart")
export class OrgChartElement extends LitElement {
  static override styles = css`
    .chart-container {
      overflow: hidden;
      height: 100%;
    }

    .helper-elements {
      position: absolute;
      top: 0;
      width: 0;
      height: 0;
      visibility: hidden;
    }
  `;

  private chart: OrgChart<unknown>;
  private fontSizes: Map<string, number>;
  private containerElement?: HTMLDivElement;

  constructor() {
    super();
    this.chart = new OrgChart();
    this.fontSizes = new Map();
  }

  setData(data: OrgChartEntry[]): string {
    const strippedData = [];
    for (const entry of data) {
      if (
        !Boolean(entry.id) &&
        !Boolean(entry.parentId) &&
        !Boolean(entry.name) &&
        !Boolean(entry.position) &&
        !Boolean(entry.color)
      ) {
        continue;
      }
      strippedData.push(entry);
    }

    if (strippedData.length === 0) {
      return "There must be at least one person in the organisation.";
    }

    try {
      this.chart.data(strippedData);
      if (this.containerElement !== undefined) {
        this.chart.expandAll().compact(true).fit().render();
      }
    } catch (e) {
      console.error(e);
      return (e as Error).message;
    }

    return "";
  }

  protected override firstUpdated(): void {
    this.containerElement = this.shadowRoot!.querySelector(
      ".chart-container"
    )! as HTMLDivElement;

    this.chart
      .nodeHeight(d => 120)
      .nodeWidth(d => 160)
      .childrenMargin(d => 50)
      .compactMarginBetween(d => 50)
      .compactMarginPair(d => 50)
      .neighbourMargin((a, b) => 30)
      .linkUpdate(function (d, i, arr) {
        // @ts-ignore
        d3.select(this)
          .attr("stroke", () => "#000000")
          .attr("stroke-width", () => 2);
      })

      .nodeContent((d, i, arr, state) => {
        const a = d as any;
        let color = a.data.color;
        if (!isColor(color)) {
          color = "#ffffff";
        }
        let fontSize = this.fontSizes.get(d.id!);
        const fontColor = pickTextColorBasedOnBgColor(
          color,
          "#FFFFFF",
          "#000000"
        );
        function nodeHtml(size: number): string {
          return `
                <div class="node-html-container" style='width:${
                  a.width
                }px;height:${d.height}px;' >
                  <div style="display: flex; flex-direction:column; justify-content: center; font-family: 'Arial', sans-serif;background-color:${color};  width:${
                    a.width - 4
                  }px; height: ${
                    d.height - 4
                  }px;border-radius:10px;border: 2px solid #000000; position:relative;">
                    <div style="padding: 5px 8px 5px 10px;">
                    <div style="font-size: ${size}px;color:${fontColor}; font-weight: 500;">${
                      a.data.name
                    } </div>
                    <div style="color:${fontColor};margin-top:3px;font-size:15px;"> ${
                      a.data.position
                    } </div>
                    </div>
                  </div>
                </div>`;
        }
        const resizer = this.shadowRoot!.getElementById("font-resizer");

        if (fontSize === undefined) {
          fontSize = 25;

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
          this.fontSizes.set(d.id!, fontSize);
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
        const resultHtml = `<div class="node-button-container" style="font-family: 'Arial', sans-serif; border:1px solid #000000;border-radius:3px;padding: 3px 5px 1px 5px;font-size:12px;margin:auto auto;background-color:white"> ${top(
          node.children !== undefined && node.children !== null
        )}  </div>`;

        return resultHtml;
      })
      // @ts-ignore
      .container(this.containerElement)
      .svgHeight(this.containerElement.clientHeight);

    if (this.chart.data() !== null) {
      this.chart.expandAll().compact(true).fit().render();
    }

    window.addEventListener("resize", this.onResize);
  }

  onResize = (): void => {
    console.error(this.containerElement!.clientHeight);
    this.chart.svgHeight(this.containerElement!.clientHeight);
  };

  override disconnectedCallback(): void {
    window.removeEventListener("resize", this.onResize);
  }

  async exportImage(
    type: OrgChartElementExportType,
    openNewWindow: boolean
  ): Promise<void> {
    if (this.containerElement === undefined) {
      return;
    }
    const svg = this.shadowRoot!.querySelector(
      ".svg-chart-container"
    ) as SVGElement;

    const xmlns = "http://www.w3.org/2000/xmlns/";
    const xlinkns = "http://www.w3.org/1999/xlink";
    const svgns = "http://www.w3.org/2000/svg";

    const oldDuration = this.chart.duration();
    this.chart.duration(0);
    this.chart.expandAll();
    this.chart.fit({animate: false});

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

      if (type === OrgChartElementExportType.Svg) {
        this.removeForeignObjects(cloned);
      }

      const serializer = new XMLSerializer();
      let src = serializer.serializeToString(cloned);
      src = '<?xml version="1.0" standalone="no"?>\r\n' + src;

      const url = "data:image/svg+xml;charset=utf-8," + encodeURIComponent(src);

      switch (type) {
        case OrgChartElementExportType.Png:
          exportPng(url, width, height, openNewWindow);
          break;
        case OrgChartElementExportType.Svg:
          saveAs(url, "orgchart.svg", openNewWindow);
          break;
        case OrgChartElementExportType.Print:
          print(
            url,
            openNewWindow,
            this.shadowRoot!.getElementById("print")! as HTMLIFrameElement,
            width,
            height,
            false
          );
          break;
        case OrgChartElementExportType.Pdf:
          print(
            url,
            openNewWindow,
            this.shadowRoot!.getElementById("print")! as HTMLIFrameElement,
            width,
            height,
            true
          );
          break;
      }

      this.chart.duration(oldDuration);
    });
  }

  public fit(): void {
    this.chart.fit();
  }

  public expandAll(): void {
    this.chart.expandAll();
  }

  private removeForeignObjects(cloned: HTMLElement): void {
    const svgConversionFrame = this.shadowRoot!.getElementById(
      "svg"
    ) as HTMLIFrameElement;
    svgConversionFrame.contentWindow!.document.body.style.margin = "0";

    let nodes = cloned.getElementsByClassName("node-html-container");
    const replacementMap = new Map<HTMLElement, SVGGElement>();
    for (const node of nodes) {
      svgConversionFrame.contentWindow!.document.body.innerHTML =
        node.parentElement!.innerHTML;

      const svgDocument = elementToSVG(
        svgConversionFrame.contentWindow!.document.body
      );

      const toReplace = node.parentElement!.parentElement!;
      const replacemenet = document.createElementNS(
        "http://www.w3.org/2000/svg",
        "g"
      );

      replacemenet.setAttribute("width", toReplace.getAttribute("width")!);
      replacemenet.setAttribute("height", toReplace.getAttribute("height")!);
      replacemenet.setAttribute("x", toReplace.getAttribute("x")!);
      replacemenet.setAttribute("y", toReplace.getAttribute("y")!);
      replacemenet.setAttribute("style", toReplace.getAttribute("style")!);
      replacemenet.innerHTML =
        svgDocument.firstElementChild!.getElementsByClassName(
          "node-html-container"
        )[0]!.innerHTML;
      replacementMap.set(toReplace, replacemenet);
    }

    nodes = cloned.getElementsByClassName("node-button-container");
    for (const node of nodes) {
      svgConversionFrame.contentWindow!.document.body.innerHTML =
        node.parentElement!.innerHTML;
      (
        svgConversionFrame.contentWindow!.document.body
          .firstElementChild! as HTMLElement
      ).style.display = "inline-block";

      const svgDocument = elementToSVG(
        svgConversionFrame.contentWindow!.document.body
      );

      const rect = (
        svgConversionFrame.contentWindow!.document.getElementsByClassName(
          "node-button-container"
        )[0] as HTMLElement
      ).getBoundingClientRect();

      const toReplace = node.parentElement!.parentElement!;
      const replacemenet = document.createElementNS(
        "http://www.w3.org/2000/svg",
        "g"
      );

      replacemenet.setAttribute("width", String(rect.width));
      replacemenet.setAttribute("height", String(rect.height));
      replacemenet.setAttribute(
        "transform",
        `translate(${rect.width / -2},${rect.height / -2})`
      );
      replacemenet.setAttribute("style", toReplace.getAttribute("style")!);
      replacemenet.innerHTML =
        svgDocument.firstElementChild!.getElementsByClassName(
          "node-button-container"
        )[0]!.innerHTML;
      replacementMap.set(toReplace, replacemenet);
    }

    for (const [key, value] of replacementMap) {
      key.replaceWith(value);
    }

    for (const node of cloned.getElementsByTagName("tspan")) {
      node.removeAttribute("textLength");
    }
  }

  override render(): unknown {
    return html` <div class="chart-container"></div>
      <div class="helper-elements">
        <div id="font-resizer"></div>
        <iframe id="svg"></iframe>
        <iframe id="print"></iframe>
      </div>`;
  }
}
