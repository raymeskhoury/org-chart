import "@shoelace-style/shoelace/dist/components/button/button";
import "@shoelace-style/shoelace/dist/components/dialog/dialog";
import SlDialog from "@shoelace-style/shoelace/dist/components/dialog/dialog";
import "@shoelace-style/shoelace/dist/themes/dark.css";
import "@shoelace-style/shoelace/dist/themes/light.css";
import * as d3 from "d3";
import {OrgChart} from "d3-org-chart";
import "./style.css";

const dialog = document.querySelector(".dialog-overview") as SlDialog;
const openButton = dialog!.nextElementSibling;
const closeButton = dialog!.querySelector('sl-button[slot="footer"]');

openButton!.addEventListener("click", () => dialog!.show());
closeButton!.addEventListener("click", () => dialog!.hide());

let chart: OrgChart<unknown>;
const fontSizes = new Map();

document.addEventListener("DOMContentLoaded", () => {});

Office.onReady(info => {
  // Check that we loaded into Excel
  // if (info.host === Office.HostType.Excel) {
  // }
});
document!.getElementById("helloButton")!.onclick = sayHello;

function doesOverflow(el: HTMLElement): boolean {
  return el.clientWidth < el.scrollWidth || el.clientHeight < el.scrollHeight;
}

async function sayHello(): Promise<void> {
  const svg = document.getElementsByClassName("svg-chart-container")[0];

  const xmlns = "http://www.w3.org/2000/xmlns/";
  const xlinkns = "http://www.w3.org/1999/xlink";
  const svgns = "http://www.w3.org/2000/svg";
  // const fragment = window.location.href + "#";

  // document.getElementsByClassName("chart-container")[0]!.style.width =
  //   maxX - minX + "px";
  // document.getElementsByClassName("chart-container")[0]!.style.height =
  //   maxY - minY + "px";
  // document.getElementsByClassName("svg-chart-container")[0]!.style.width =
  //   maxX - minX + "px";
  // document.getElementsByClassName("svg-chart-container")[0]!.style.height =
  //   maxY - minY + "px";

  // setTimeout(() => {
  // chart.expandAll();
  // chart.fit();
  // chart.render();
  const oldDuration = chart.duration();
  chart.duration(0);
  chart.expandAll();
  chart.fit();

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
    console.error(svgRect);
    console.error(minY);
    minX = minX! - svgRect.x;
    minY = minY! - svgRect.y;
    maxX = maxX! - svgRect.x;
    maxY = maxY! - svgRect.y;
    console.error(minX + " " + maxX + " " + minY + " " + maxY);
    const width = maxX - minX;
    const height = maxY - minY;

    const cloned = svg.cloneNode(true);
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

    src =
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
        });
      </script>
      </head>
      <body>
        <img class="svg-chart-container" src="
      ` +
      url +
      `" /></body></html>`;

    const frame = document.getElementById("print")! as HTMLIFrameElement;

    frame.srcdoc = src;
    chart.duration(oldDuration);
  });
}

document.addEventListener("DOMContentLoaded", () => {
  d3.csv("./assets/data-oracle.csv").then(data => {
    OrgChart.prototype.diagonal = function (s, t, m, offsets = {sy: 0}) {
      const x = s.x;
      let y = s.y;

      const ex = t.x;
      const ey = t.y;

      let mx = m && m.x != null ? m.x : x; // This is a changed line
      let my = m && m.y != null ? m.y : y; // This also is a changed line

      let xrvs = ex - x < 0 ? -1 : 1;
      let yrvs = ey - y < 0 ? -1 : 1;

      y += offsets.sy;

      let rdef = 0;
      let r = Math.abs(ex - x) / 2 < rdef ? Math.abs(ex - x) / 2 : rdef;

      r = Math.abs(ey - y) / 2 < r ? Math.abs(ey - y) / 2 : r;

      let h = Math.abs(ey - y) / 2 - r;
      let w = Math.abs(ex - x) - r * 2;
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
        d3.select(this)
          .attr("stroke", d => "#000000")
          .attr("stroke-width", d => 2);
      })

      .nodeContent((d, i, arr, state) => {
        const a = d as any;
        let fontSize = fontSizes.get(d.id);
        function nodeHtml(size: number): string {
          return `
                <div style='width:${a.width}px;height:${d.height}px;' >
                  <div style="display: flex; flex-direction:column; justify-content: center;  font-family: 'Inter', sans-serif;background-color:#ffffff;  width:${
                    d.width - 4
                  }px; height: ${
                    d.height - 4
                  }px;border-radius:10px;border: 2px solid #000000; position:relative">
                    <div style="padding: 5px 8px 5px 10px;overflow:hidden;">
                    <div style="font-size: ${size}px;color:#08011E; font-weight: 500;">  ${
                      a.data.name
                    } </div>
                    <div style="color:#000000;margin-top:3px;font-size:15px;"> ${
                      a.data.position
                    } </div>
                    </div>
                  </div>
                </div>`;
        }
        if (fontSize === undefined) {
          fontSize = 25;

          const resizer = document.getElementById("fontResizer");
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
        const icons = {
          left: d =>
            d
              ? `<div style="display:flex;"><span style="align-items:center;display:flex;"><svg width="8" height="8" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
            <path d="M14.283 3.50094L6.51 11.4749C6.37348 11.615 6.29707 11.8029 6.29707 11.9984C6.29707 12.194 6.37348 12.3819 6.51 12.5219L14.283 20.4989C14.3466 20.5643 14.4226 20.6162 14.5066 20.6516C14.5906 20.6871 14.6808 20.7053 14.772 20.7053C14.8632 20.7053 14.9534 20.6871 15.0374 20.6516C15.1214 20.6162 15.1974 20.5643 15.261 20.4989C15.3918 20.365 15.4651 20.1852 15.4651 19.9979C15.4651 19.8107 15.3918 19.6309 15.261 19.4969L7.9515 11.9984L15.261 4.50144C15.3914 4.36756 15.4643 4.18807 15.4643 4.00119C15.4643 3.81431 15.3914 3.63482 15.261 3.50094C15.1974 3.43563 15.1214 3.38371 15.0374 3.34827C14.9534 3.31282 14.8632 3.29456 14.772 3.29456C14.6808 3.29456 14.5906 3.31282 14.5066 3.34827C14.4226 3.38371 14.3466 3.43563 14.283 3.50094V3.50094Z" fill="#000000" stroke="#000000"/>
            </svg></span><span style="color:#000000">${node.data._directSubordinatesPaging} </span></div>`
              : `<div style="display:flex;"><span style="align-items:center;display:flex;"><svg width="8" height="8" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
                <path d="M7.989 3.49944C7.85817 3.63339 7.78492 3.8132 7.78492 4.00044C7.78492 4.18768 7.85817 4.36749 7.989 4.50144L15.2985 11.9999L7.989 19.4969C7.85817 19.6309 7.78492 19.8107 7.78492 19.9979C7.78492 20.1852 7.85817 20.365 7.989 20.4989C8.05259 20.5643 8.12863 20.6162 8.21261 20.6516C8.2966 20.6871 8.38684 20.7053 8.478 20.7053C8.56916 20.7053 8.6594 20.6871 8.74338 20.6516C8.82737 20.6162 8.90341 20.5643 8.967 20.4989L16.74 12.5234C16.8765 12.3834 16.9529 12.1955 16.9529 11.9999C16.9529 11.8044 16.8765 11.6165 16.74 11.4764L8.967 3.50094C8.90341 3.43563 8.82737 3.38371 8.74338 3.34827C8.6594 3.31282 8.56916 3.29456 8.478 3.29456C8.38684 3.29456 8.2966 3.31282 8.21261 3.34827C8.12863 3.38371 8.05259 3.43563 7.989 3.50094V3.49944Z" fill="#000000" stroke="#000000"/>
                </svg></span><span style="color:#000000">${node.data._directSubordinatesPaging} </span></div>`,
          bottom: d =>
            d
              ? `<div style="display:flex;"><span style="align-items:center;display:flex;"><svg width="8" height="8" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
             <path d="M19.497 7.98903L12 15.297L4.503 7.98903C4.36905 7.85819 4.18924 7.78495 4.002 7.78495C3.81476 7.78495 3.63495 7.85819 3.501 7.98903C3.43614 8.05257 3.38462 8.12842 3.34944 8.21213C3.31427 8.29584 3.29615 8.38573 3.29615 8.47653C3.29615 8.56733 3.31427 8.65721 3.34944 8.74092C3.38462 8.82463 3.43614 8.90048 3.501 8.96403L11.4765 16.74C11.6166 16.8765 11.8044 16.953 12 16.953C12.1956 16.953 12.3834 16.8765 12.5235 16.74L20.499 8.96553C20.5643 8.90193 20.6162 8.8259 20.6517 8.74191C20.6871 8.65792 20.7054 8.56769 20.7054 8.47653C20.7054 8.38537 20.6871 8.29513 20.6517 8.21114C20.6162 8.12715 20.5643 8.05112 20.499 7.98753C20.3651 7.85669 20.1852 7.78345 19.998 7.78345C19.8108 7.78345 19.6309 7.85669 19.497 7.98753V7.98903Z" fill="#000000" stroke="#000000"/>
             </svg></span><span style="margin-left:1px;color:#000000" >${node.data._directSubordinatesPaging} </span></div>
             `
              : `<div style="display:flex;"><span style="align-items:center;display:flex;"><svg width="8" height="8" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
             <path d="M11.457 8.07005L3.49199 16.4296C3.35903 16.569 3.28485 16.7543 3.28485 16.9471C3.28485 17.1398 3.35903 17.3251 3.49199 17.4646L3.50099 17.4736C3.56545 17.5414 3.64304 17.5954 3.72904 17.6324C3.81504 17.6693 3.90765 17.6883 4.00124 17.6883C4.09483 17.6883 4.18745 17.6693 4.27344 17.6324C4.35944 17.5954 4.43703 17.5414 4.50149 17.4736L12.0015 9.60155L19.4985 17.4736C19.563 17.5414 19.6405 17.5954 19.7265 17.6324C19.8125 17.6693 19.9052 17.6883 19.9987 17.6883C20.0923 17.6883 20.1849 17.6693 20.2709 17.6324C20.3569 17.5954 20.4345 17.5414 20.499 17.4736L20.508 17.4646C20.641 17.3251 20.7151 17.1398 20.7151 16.9471C20.7151 16.7543 20.641 16.569 20.508 16.4296L12.543 8.07005C12.4729 7.99653 12.3887 7.93801 12.2954 7.89801C12.202 7.85802 12.1015 7.8374 12 7.8374C11.8984 7.8374 11.798 7.85802 11.7046 7.89801C11.6113 7.93801 11.527 7.99653 11.457 8.07005Z" fill="#000000" stroke="#000000"/>
             </svg></span><span style="margin-left:1px;color:#000000" >${node.data._directSubordinatesPaging} </span></div>
          `,
          right: d =>
            d
              ? `<div style="display:flex;"><span style="align-items:center;display:flex;"><svg width="8" height="8" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
             <path d="M7.989 3.49944C7.85817 3.63339 7.78492 3.8132 7.78492 4.00044C7.78492 4.18768 7.85817 4.36749 7.989 4.50144L15.2985 11.9999L7.989 19.4969C7.85817 19.6309 7.78492 19.8107 7.78492 19.9979C7.78492 20.1852 7.85817 20.365 7.989 20.4989C8.05259 20.5643 8.12863 20.6162 8.21261 20.6516C8.2966 20.6871 8.38684 20.7053 8.478 20.7053C8.56916 20.7053 8.6594 20.6871 8.74338 20.6516C8.82737 20.6162 8.90341 20.5643 8.967 20.4989L16.74 12.5234C16.8765 12.3834 16.9529 12.1955 16.9529 11.9999C16.9529 11.8044 16.8765 11.6165 16.74 11.4764L8.967 3.50094C8.90341 3.43563 8.82737 3.38371 8.74338 3.34827C8.6594 3.31282 8.56916 3.29456 8.478 3.29456C8.38684 3.29456 8.2966 3.31282 8.21261 3.34827C8.12863 3.38371 8.05259 3.43563 7.989 3.50094V3.49944Z" fill="#000000" stroke="#000000"/>
             </svg></span><span style="color:#000000">${node.data._directSubordinatesPaging} </span></div>`
              : `<div style="display:flex;"><span style="align-items:center;display:flex;"><svg width="8" height="8" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
             <path d="M14.283 3.50094L6.51 11.4749C6.37348 11.615 6.29707 11.8029 6.29707 11.9984C6.29707 12.194 6.37348 12.3819 6.51 12.5219L14.283 20.4989C14.3466 20.5643 14.4226 20.6162 14.5066 20.6516C14.5906 20.6871 14.6808 20.7053 14.772 20.7053C14.8632 20.7053 14.9534 20.6871 15.0374 20.6516C15.1214 20.6162 15.1974 20.5643 15.261 20.4989C15.3918 20.365 15.4651 20.1852 15.4651 19.9979C15.4651 19.8107 15.3918 19.6309 15.261 19.4969L7.9515 11.9984L15.261 4.50144C15.3914 4.36756 15.4643 4.18807 15.4643 4.00119C15.4643 3.81431 15.3914 3.63482 15.261 3.50094C15.1974 3.43563 15.1214 3.38371 15.0374 3.34827C14.9534 3.31282 14.8632 3.29456 14.772 3.29456C14.6808 3.29456 14.5906 3.31282 14.5066 3.34827C14.4226 3.38371 14.3466 3.43563 14.283 3.50094V3.50094Z" fill="#000000" stroke="#000000"/>
             </svg></span><span style="color:#000000">${node.data._directSubordinatesPaging} </span></div>`,
          top: d =>
            d
              ? `<div style="display:flex;"><span style="align-items:center;display:flex;"><svg width="8" height="8" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
              <path d="M11.457 8.07005L3.49199 16.4296C3.35903 16.569 3.28485 16.7543 3.28485 16.9471C3.28485 17.1398 3.35903 17.3251 3.49199 17.4646L3.50099 17.4736C3.56545 17.5414 3.64304 17.5954 3.72904 17.6324C3.81504 17.6693 3.90765 17.6883 4.00124 17.6883C4.09483 17.6883 4.18745 17.6693 4.27344 17.6324C4.35944 17.5954 4.43703 17.5414 4.50149 17.4736L12.0015 9.60155L19.4985 17.4736C19.563 17.5414 19.6405 17.5954 19.7265 17.6324C19.8125 17.6693 19.9052 17.6883 19.9987 17.6883C20.0923 17.6883 20.1849 17.6693 20.2709 17.6324C20.3569 17.5954 20.4345 17.5414 20.499 17.4736L20.508 17.4646C20.641 17.3251 20.7151 17.1398 20.7151 16.9471C20.7151 16.7543 20.641 16.569 20.508 16.4296L12.543 8.07005C12.4729 7.99653 12.3887 7.93801 12.2954 7.89801C12.202 7.85802 12.1015 7.8374 12 7.8374C11.8984 7.8374 11.798 7.85802 11.7046 7.89801C11.6113 7.93801 11.527 7.99653 11.457 8.07005Z" fill="#000000" stroke="#000000"/>
              </svg></span><span style="margin-left:1px;color:#000000">${node.data._directSubordinatesPaging} </span></div>
              `
              : `<div style="display:flex;"><span style="align-items:center;display:flex;"><svg width="8" height="8" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
              <path d="M19.497 7.98903L12 15.297L4.503 7.98903C4.36905 7.85819 4.18924 7.78495 4.002 7.78495C3.81476 7.78495 3.63495 7.85819 3.501 7.98903C3.43614 8.05257 3.38462 8.12842 3.34944 8.21213C3.31427 8.29584 3.29615 8.38573 3.29615 8.47653C3.29615 8.56733 3.31427 8.65721 3.34944 8.74092C3.38462 8.82463 3.43614 8.90048 3.501 8.96403L11.4765 16.74C11.6166 16.8765 11.8044 16.953 12 16.953C12.1956 16.953 12.3834 16.8765 12.5235 16.74L20.499 8.96553C20.5643 8.90193 20.6162 8.8259 20.6517 8.74191C20.6871 8.65792 20.7054 8.56769 20.7054 8.47653C20.7054 8.38537 20.6871 8.29513 20.6517 8.21114C20.6162 8.12715 20.5643 8.05112 20.499 7.98753C20.3651 7.85669 20.1852 7.78345 19.998 7.78345C19.8108 7.78345 19.6309 7.85669 19.497 7.98753V7.98903Z" fill="#000000" stroke="#000000"/>
              </svg></span><span style="margin-left:1px;color:#000000">${node.data._directSubordinatesPaging} </span></div>
          `,
        };
        return `<div style="border:1px solid #000000;border-radius:3px;padding: 3px 5px 1px 5px;font-size:12px;margin:auto auto;background-color:white"> ${icons[
          state.layout
        ](node.children)}  </div>`;
      })
      .container(".chart-container")
      .data(data)
      .expandAll()
      .compact(true)
      .fit()
      .render();
  });
});
