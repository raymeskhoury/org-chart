import {MobxLitElement} from "@adobe/lit-mobx";
import "@shoelace-style/shoelace/dist/components/color-picker/color-picker";
import SlColorPicker from "@shoelace-style/shoelace/dist/components/color-picker/color-picker";
import "@shoelace-style/shoelace/dist/components/dialog/dialog";
import SlDialog from "@shoelace-style/shoelace/dist/components/dialog/dialog";
import "@shoelace-style/shoelace/dist/components/icon-button/icon-button";
import "@shoelace-style/shoelace/dist/components/input/input";
import SlInput from "@shoelace-style/shoelace/dist/components/input/input";
import "@shoelace-style/shoelace/dist/components/tab-group/tab-group";
import "@shoelace-style/shoelace/dist/components/tab-panel/tab-panel";
import "@shoelace-style/shoelace/dist/components/tab/tab";
import "@shoelace-style/shoelace/dist/components/textarea/textarea";
import SlTextarea from "@shoelace-style/shoelace/dist/components/textarea/textarea";
import {LitElement, css, html, nothing} from "lit-element";
import {customElement, property} from "lit/decorators.js";
import {action, autorun, makeObservable, toJS} from "mobx";
import {OrgChartDataModel, OrgChartEntry} from "./OrgChartDataModel";
import {ASSERT} from "./Util";

const sharedCss = css`
  .id-column {
    min-width: 60px;
    width: 60px;
  }

  .name-column,
  .position-column {
    min-width: 200px;
    width: 200px;
  }

  .color-picker-column {
    min-width: 40px;
    width: 40px;
  }
`;

@customElement("org-chart-input-row")
export class InputElementRow extends MobxLitElement {
  static override styles = css`
    .fields {
      display: flex;
      flex-direction: row;
      column-gap: 10px;
    }

    ${sharedCss}

    .fields .color-picker-column {
      display: flex;
      align-items: center;
      justify-content: center;
    }
  `;

  @property({attribute: false})
  orgChartDataEntry?: OrgChartEntry;

  constructor() {
    super();
    makeObservable(this);
  }

  protected override render(): unknown {
    console.error("render2");
    if (!this.orgChartDataEntry) {
      return nothing;
    }
    return html` <div class="fields">
      <sl-input
        id="id"
        class="id-column"
        value="${this.orgChartDataEntry!.id}"
        @sl-input="${action((e: CustomEvent) => {
          this.orgChartDataEntry!.id = (e.target as SlInput).value;
        })}"
      ></sl-input>
      <sl-input
        id="manager"
        class="id-column"
        value="${this.orgChartDataEntry!.parentId}"
        @sl-input="${action((e: CustomEvent) => {
          this.orgChartDataEntry!.parentId = (e.target as SlInput).value;
        })}"
      ></sl-input>
      <sl-input
        id="name"
        class="name-column"
        value="${this.orgChartDataEntry!.name}"
        @sl-input="${action((e: CustomEvent) => {
          this.orgChartDataEntry!.name = (e.target as SlInput).value;
        })}"
      ></sl-input>
      <sl-input
        id="position"
        class="position-column"
        value="${this.orgChartDataEntry!.position}"
        @sl-input="${action((e: CustomEvent) => {
          this.orgChartDataEntry!.position = (e.target as SlInput).value;
        })}"
      ></sl-input>

      <div class="color-picker-column">
        <sl-color-picker
          id="color-picker"
          label="Color"
          value="white"
          size="small"
          ?hoist=${false}
          value="${this.orgChartDataEntry!.color}"
          @sl-input="${action((e: CustomEvent) => {
            this.orgChartDataEntry!.color = (e.target as SlColorPicker).value;
          })}"
        ></sl-color-picker>
      </div>
    </div>`;
  }
}

@customElement("org-chart-input")
export class InputElement extends MobxLitElement {
  static override styles = css`
    #container {
      padding: 1px;
      display: flex;
      flex-direction: column;
    }
    #error {
      margin-top: 20px;
      color: red;
    }

    .headers {
      display: flex;
      flex-direction: row;
      column-gap: 10px;
    }

    .headers {
      margin-bottom: 10px;
    }

    .entries {
      display: flex;
      flex-direction: column;
      row-gap: 10px;
    }

    ${sharedCss}

    .buttons {
      width: 100%;
      display: flex;
      margin-top: 20px;
      justify-content: right;
      column-gap: 10px;
    }

    .entry {
      display: flex;
      flex-direction: row;
      align-items: center;
      column-gap: 10px;
    }

    #container {
      height: 100%;
    }

    .headers-and-entries {
      overflow: auto;
      height: 100%;
    }

    .delete-button-column {
      min-width: 30px;
      width: 30px;
    }
  `;

  @property({attribute: false})
  orgChartData?: OrgChartDataModel;

  private scrollRegion?: HTMLDivElement;

  constructor() {
    super();
    console.error(this.orgChartData);
    makeObservable(this);
  }

  protected override firstUpdated(): void {
    this.scrollRegion = this.shadowRoot!.querySelector(
      ".headers-and-entries"
    ) as HTMLDivElement;
  }

  protected override render(): unknown {
    if (!this.orgChartData) {
      return nothing;
    }
    console.error("render");
    const entries = [];
    for (let i = 0; i < this.orgChartData.data.length; ++i) {
      entries.push(html`
        <div class="entry">
          <div class="delete-button-column">
            <sl-icon-button
              name="x-lg"
              @click="${action(() => {
                if (this.orgChartData!.data.length <= 1) {
                  return;
                }
                this.orgChartData!.data.splice(i, 1);
              })}"
              
            ></sl-icon-button>
          </div>
            <org-chart-input-row
              .orgChartDataEntry=${this.orgChartData.data[i]}
            ></org-chart-input-row>
          </div>
        </div>
      `);
    }
    return html`<div id="container">
      <div class="headers-and-entries">
        <div class="headers">
          <div class="delete-button-column"></div>
          <div class="id-column">ID</div>
          <div class="id-column">Manager ID</div>
          <div class="name-column">Name</div>
          <div class="position-column">Position</div>
          <div class="color-picker-column">Color</div>
        </div>
        <div class="entries">${entries}</div>
      </div>
      <div class="buttons">
        <sl-button
          @click="${action(() => {
            this.orgChartData!.resetToSample();
          })}"
        >
          Reset to sample data
        </sl-button>
        <sl-button
          variant="primary"
          @click="${action(() => {
            this.orgChartData!.data.push(new OrgChartEntry());
            this.scrollRegion!.scrollTop = this.scrollRegion!.scrollHeight;
          })}"
        >
          <sl-icon slot="prefix" name="plus-lg"></sl-icon>
          Add</sl-button
        >
      </div>
    </div>`;
  }
}

@customElement("org-chart-import")
export class ImportElement extends MobxLitElement {
  static override styles = css`
    #container {
      padding: 1px;
      display: flex;
      flex-direction: column;
      height: 100%;
    }

    #import-text,
    sl-textarea::part(form-control),
    sl-textarea::part(form-control-input),
    sl-textarea::part(base),
    sl-textarea::part(textarea) {
      height: 100%;
    }

    #error {
      margin-top: 20px;
      height: 20px;
      color: red;
    }

    .buttons {
      display: flex;
      column-gap: 10px;
      justify-content: right;
      margin-top: 20px;
    }
  `;

  private initialTextSet: boolean;

  @property({attribute: false})
  orgChartData?: OrgChartDataModel;

  @property({state: true, attribute: false})
  private text = "";

  @property({state: true, attribute: false})
  private error = "";

  constructor() {
    super();
    this.initialTextSet = false;
    makeObservable(this);
  }

  protected override willUpdate(): void {
    if (!this.initialTextSet) {
      this.initialTextSet = true;
      autorun(() => {
        this.text = this.orgChartData!.getCsv();
        this.error = "";
      });
    }
  }

  protected override render(): unknown {
    if (!this.orgChartData) {
      return nothing;
    }
    return html`<div id="container">
      <sl-textarea
        id="import-text"
        resize="none"
        value=${this.text}
        @sl-input=${(e: CustomEvent) => {
          this.text = (e.target as SlTextarea).value;
          this.error = this.orgChartData!.fromCsv(this.text);
        }}
      ></sl-textarea>
      <div id="error">${this.error}</div>

      <div class="buttons">
        <sl-button
          ?disabled=${this.error === ""}
          @click=${() => {
            this.text = this.orgChartData!.getCsv();
            this.error = "";
          }}
          >Reset to Input data</sl-button
        >
        <sl-button
          @click=${() => {
            const input = document.createElement("input");
            input.type = "file";
            input.onchange = e => {
              const files = (e.target! as HTMLInputElement).files;
              if (files === undefined || files!.length === 0) {
                return;
              }
              const reader = new FileReader();
              reader.readAsText(files![0]);
              reader.onload = readerEvent => {
                this.text = readerEvent.target!.result as string;
                this.error = this.orgChartData!.fromCsv(this.text);
              };
            };

            input.click();
          }}
          >Import from file...</sl-button
        >
        <sl-button
          download="orgchart.csv"
          href="data:text/plain;charset=utf-8,${encodeURIComponent(this.text)}"
          >Export to file...</sl-button
        >
      </div>
    </div>`;
  }
}

@customElement("org-chart-data-dialog")
export class DataDialogElement extends LitElement {
  static override styles = css`
    #root {
      --width: 710px;
    }

    .dialog-content {
      height: 100%;
    }

    .tabs {
      height: calc(100% - 100px);
    }

    #input {
      display: block;
      height: 100%;
    }

    #error {
      margin-bottom: 20px;
      color: red;
      height: 35px;
    }

    sl-dialog::part(base),
    sl-dialog::part(panel),
    sl-tab-group::part(base),
    sl-tab-group::part(body),
    sl-tab-panel::part(base),
    sl-tab-panel::part(body),
    sl-tab-panel {
      height: 100%;
    }
  `;

  private dialog?: SlDialog;
  private resolveFunction?: (result: OrgChartEntry[]) => void;
  private showPromise?: Promise<OrgChartEntry[]>;
  private hiding = false;

  constructor(orgChartData: OrgChartDataModel) {
    super();
    this.orgChartData = orgChartData;
  }

  @property({state: true})
  private errorMessage = "";

  @property({attribute: false})
  orgChartData?: OrgChartDataModel;

  async show(error = ""): Promise<OrgChartEntry[]> {
    this.errorMessage = error;
    while (this.showPromise) {
      console.error("waiting");
      await this.showPromise;
      console.error("finished waiting");
    }
    ASSERT(this.showPromise === undefined && this.dialog !== undefined);

    this.showPromise = new Promise<OrgChartEntry[]>(resolve => {
      this.resolveFunction = async (result: OrgChartEntry[]): Promise<void> => {
        if (this.hiding) {
          return;
        }
        this.hiding = true;

        await this.dialog!.hide();
        this.showPromise = undefined;

        this.hiding = false;
        resolve(result);
      };
    });
    await this.dialog!.show();
    return this.showPromise;
  }

  private onOk(): void {
    this.resolveFunction!(toJS(this.orgChartData!.data));
  }

  private onAttemptClose = (event: Event): void => {
    this.onOk();
  };

  protected override firstUpdated(): void {
    this.dialog = this.shadowRoot!.getElementById("root")! as SlDialog;
  }

  protected override render(): unknown {
    return html`<sl-dialog
      @sl-request-close="${this.onAttemptClose}"
      id="root"
      no-header
    >
      <div class="dialog-content">
        <sl-tab-group class="tabs">
          <sl-tab slot="nav" panel="input">Input</sl-tab>
          <sl-tab slot="nav" panel="import">Import/Export</sl-tab>

          <sl-tab-panel name="input"
            ><org-chart-input
              id="input"
              .orgChartData=${this.orgChartData}
            ></org-chart-input
          ></sl-tab-panel>
          <sl-tab-panel name="import"
            ><org-chart-import
              id="import"
              .orgChartData=${this.orgChartData}
            ></org-chart-import>
          </sl-tab-panel>
        </sl-tab-group>

        <div id="error">${this.errorMessage}</div>

        <sl-button @click="${this.onOk}" slot="footer" variant="primary" id="ok"
          >Close</sl-button
        >
      </div>
    </sl-dialog>`;
  }
}
