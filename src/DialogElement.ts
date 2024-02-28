import "@shoelace-style/shoelace/dist/components/dialog/dialog";
import SlDialog from "@shoelace-style/shoelace/dist/components/dialog/dialog";
import {LitElement, css, html, nothing} from "lit-element";
import {customElement, property} from "lit/decorators.js";
import {ASSERT} from "./Util";

export enum DialogElementType {
  OK,
  OK_CANCEL,
}

@customElement("org-chart-dialog")
export class DialogElement extends LitElement {
  static override styles = css`
    #error {
      margin-top: 20px;
      color: red;
    }
  `;

  private dialog?: SlDialog;
  private modal = false;
  private resolveFunction?: (result: boolean) => void;
  private showPromise?: Promise<boolean>;
  private hiding = false;

  @property({state: true})
  private dialogTitle = "";

  @property({state: true})
  private message = "";

  @property({state: true})
  private errorMessage = "";

  @property({state: true, type: Number})
  private type = DialogElementType.OK_CANCEL;

  @property({state: true})
  private okButton = "OK";

  @property({state: true})
  private cancelButton = "Cancel";

  async show(
    title: string,
    message: string,
    type: DialogElementType,
    errorMessage = "",
    okButton = "",
    cancelButton = "",
    modal = false
  ): Promise<boolean> {
    console.error("here");
    while (this.showPromise) {
      console.error("waiting");
      await this.showPromise;
      console.error("finished waiting");
    }
    ASSERT(this.showPromise === undefined && this.dialog !== undefined);
    this.dialogTitle = title;
    this.message = message;
    this.type = type;
    this.errorMessage = errorMessage;
    this.okButton = okButton;
    this.cancelButton = cancelButton;
    this.modal = modal;

    this.showPromise = new Promise<boolean>(resolve => {
      this.resolveFunction = async (result: boolean): Promise<void> => {
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
    this.resolveFunction!(true);
  }

  private onCancel(): void {
    this.resolveFunction!(false);
  }

  private onAttemptClose = (event: Event): void => {
    const e = event as CustomEvent;
    if (e.detail.source === "overlay" || e.detail.source === "close-button") {
      if (this.modal) {
        e.preventDefault();
        return;
      }
    }
    this.onCancel();
    e.preventDefault();
  };

  protected override firstUpdated(): void {
    this.dialog = this.shadowRoot!.getElementById("root")! as SlDialog;
  }

  override disconnectedCallback(): void {
    this.dialog!.removeEventListener("sl-request-close", this.onAttemptClose);
  }

  protected override render(): unknown {
    return html`<sl-dialog
      @sl-request-close="${this.onAttemptClose}"
      label="${this.dialogTitle}"
      id="root"
    >
      ${this.message}
      <div id="error">${this.errorMessage}</div>

      <sl-button @click="${this.onOk}" slot="footer" variant="primary" id="ok"
        >${this.okButton}</sl-button
      >
      ${this.type === DialogElementType.OK_CANCEL
        ? html`<sl-button @click="${this.onCancel}" slot="footer" id="cancel"
            >${this.cancelButton}</sl-button
          >`
        : nothing}
    </sl-dialog>`;
  }
}
