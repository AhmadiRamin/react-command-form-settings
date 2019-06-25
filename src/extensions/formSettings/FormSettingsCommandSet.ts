import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import * as jquery from 'jquery';
import ListService from './services/list-service';
import IFormItem from './components/models/form-item';
/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IFormsSettingsCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'FormsSettingsCommandSet';

export default class FormsSettingsCommandSet extends BaseListViewCommandSet<IFormsSettingsCommandSetProperties> {
  private listService = new ListService();
  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized FormsSettingsCommandSet');
    return Promise.resolve();
  }

  @override
  public async onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): Promise<void> {
    const formSettings = await this.listService.getFormSettings(String(this.context.pageContext.list.id));
    this.loadFormSettings(formSettings);

    jquery("body").on("click", `button[data-automationid='FieldRenderer-title']`,  (e) => {
      jquery(this).attr("data-selection-toggle","true");
    });
    if (event.selectedRows.length > 0) {
      console.log("Selected");
      const contentType = event.selectedRows[0].getValueByName("ContentType");
      const editForms = formSettings.filter(i=>i.ContentTypeName===contentType && i.Form==="Edit");
      const displayForms = formSettings.filter(i=>i.ContentTypeName===contentType && i.Form==="Display");
      
      if(editForms.length>0){
        this.overrideOnClick("Edit",editForms[0].OpenIn,editForms[0].RedirectURL);        
      }

      if(displayForms.length>0){
        this.ovverideDisplayClick(displayForms[0].OpenIn,displayForms[0].RedirectURL);
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

      const panel = new component.container(
        this.context.pageContext.list.id.toString()

      );
      panel.render();

    }
  }

  private async loadFormSettings(formSettings: IFormItem[]) {
    formSettings.map(form => {
      switch (form.Form) {
        case "New":
          this.overrideOnClick(form.ContentTypeName,form.OpenIn,form.RedirectURL);
          break;
      }
    });
  }

  private overrideOnClick(tagName:string,openIn:string,redirectURL:string){
    jquery("body").on("click", `button[name='${tagName}']`,  (e)=> {
      switch (openIn) {
        case "Current Window":
          window.location.href = redirectURL;
          break;
        case "New Tab":
          window.open(redirectURL, "_blank");
          break;
      }
      e.stopPropagation();
    });
  }

  private ovverideDisplayClick(openIn:string,redirectURL:string){
    jquery("body").on("click", `button[data-automationid='FieldRenderer-title'],[name='Open']`,  (e) => {
      alert('s1');
      switch (openIn) {
        case "Current Window":
          window.location.href = redirectURL;
          break;
        case "New Tab":
          window.open(redirectURL, "_blank");
          break;
      }
      e.stopPropagation();
    });
  }
}
