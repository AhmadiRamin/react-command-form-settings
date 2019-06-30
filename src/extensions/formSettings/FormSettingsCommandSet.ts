import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters,
  Command
} from '@microsoft/sp-listview-extensibility';
import * as $ from 'jquery';
import ListService from './services/list-service';
import IFormItem from './models/form-item';
/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IFormsSettingsCommandSetProperties {

}

const LOG_SOURCE: string = 'FormsSettingsCommandSet';

export default class FormsSettingsCommandSet extends BaseListViewCommandSet<IFormsSettingsCommandSetProperties> {
  private listService = new ListService();
  private formSettings: IFormItem[] = [];
  private contentTypes: any[] = [];
  private ovverideEditClick = false;
  private ovverideDispClick = false;
  private editForms: IFormItem[];
  private displayForms: IFormItem[];
  private selectedRow = null;
  private itemId: number;
  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized FormsSettingsCommandSet');

    return Promise.resolve();
  }

  @override
  public async onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): Promise<void> {
    const listId = String(this.context.pageContext.list.id);
    if (this.formSettings.length <= 0)
      this.formSettings = await this.listService.getEnabledFormSettings(listId);

    if (this.contentTypes.length <= 0)
      this.contentTypes = await this.listService.getListContentTypes(listId);

    const newForms = this.formSettings.filter(i => i.FormType === "New");
    this.ovverideNewFormSettings(newForms, this.contentTypes.length);

    $("body").on("click", `button[data-automationid='FieldRenderer-title']`, (e) => {
      this.selectedRow = $(e.target).parents().closest("div[data-automationid='DetailsRow']");
      this.selectedRow.trigger("click");
    });

    if (event.selectedRows.length > 0) {
      this.itemId = event.selectedRows[0].getValueByName("ID");

      if (this.contentTypes.length > 1) {
        const contentType = event.selectedRows[0].getValueByName("ContentType");
        this.editForms = this.formSettings.filter(i => i.ContentTypeName === contentType && i.FormType === "Edit");
        this.displayForms = this.formSettings.filter(i => i.ContentTypeName === contentType && i.FormType === "Display");
      }
      else {
        this.editForms = this.formSettings.filter(i => i.FormType === "Edit");
        this.displayForms = this.formSettings.filter(i => i.FormType === "Display");
      }

      if (this.editForms.length > 0) {
        this.ovverideEditClick = true;
        this.overrideOnClick("Edit", this.editForms[0]);
      }
      else {
        this.ovverideEditClick = false;
      }
      if (this.displayForms.length > 0) {
        this.ovverideDispClick = true;
        this.overrideOnClick("Open", this.displayForms[0]);
        if (this.selectedRow) {
          this.selectedRow = null;
          this.redirect(this.displayForms[0]);
          $(document.body.firstChild).trigger("click");

        }
      }
      else {
        this.selectedRow = null;
        this.ovverideDispClick = false;
      }

    }
  }

  @override
  public async onExecute(event: IListViewCommandSetExecuteEventParameters): Promise<void> {
    if (event.itemId === "COMMAND_Form_Settings") {
      const component = await import(
        /* webpackMode: "lazy" */
        /* webpackChunkName: 'multisharedialog-component' */
        './components/container/container'
      );

      const panel = new component.container();
      panel.listId = this.context.pageContext.list.id.toString();
      panel.formSettings = this.formSettings;
      panel.contentTypes = this.contentTypes;
      panel.render();

    }
  }
  private async ovverideNewFormSettings(formSettings: IFormItem[], ctCount: number) {
    // Ovveride if only one content type exists in the list
    if (ctCount < 2)
      this.overrideOnClick("New", formSettings[0]);
    else {
      formSettings.map(form => {
        this.overrideOnClick(form.ContentTypeName, form);
      });
    }
  }

  private overrideOnClick(tagName: string, settings: IFormItem) {
    $("body").on("click", `button[name='${tagName}']`, (e) => {
      switch (tagName) {
        case "Edit":
          this.ovverideEditClick && this.redirect(settings); return e.stopPropagation();
        case "Open":
          this.ovverideDispClick && this.redirect(settings); return e.stopPropagation();
        default:
          this.redirect(settings); return e.stopPropagation();
      }
    });
  }

  private redirect(settings: IFormItem) {
    const { OpenIn, RedirectURL, Parameters } = settings;
    let tokens = "";
    tokens = Parameters && Parameters.length > 0 ? `?${this.replaceTokens(Parameters)}` : "";
    switch (OpenIn) {
      case "Current Window":
        window.location.href = `${RedirectURL}${tokens}`;
        break;
      case "New Tab":
        window.open(`${RedirectURL}${tokens}`, "_blank");
        break;
    }
  }

  private replaceTokens(tokens: string) {
    if (!tokens)
      return "";
    return tokens.replace("{ListId}", String(this.context.pageContext.list.id))
      .replace("{WebUrl}", this.context.pageContext.web.absoluteUrl)
      .replace("{SiteUrl}", this.context.pageContext.site.absoluteUrl)
      .replace("{UserLoginName}", this.context.pageContext.user.loginName)
      .replace("{UserDisplayName}", this.context.pageContext.user.displayName)
      .replace("{UserEmail}", this.context.pageContext.user.email)
      .replace("{ItemId}", String(this.itemId));
  }
}
