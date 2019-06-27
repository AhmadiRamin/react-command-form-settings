import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import * as jquery from 'jquery';
import ListService from './services/list-service';
import IFormItem from './models/form-item';
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
  private ovverideClick = false;
  private selectedRow = null;
  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized FormsSettingsCommandSet');
    return Promise.resolve();
  }

  @override
  public async onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): Promise<void> {
    const formSettings = await this.listService.getEnabledFormSettings(String(this.context.pageContext.list.id));
    
    this.loadFormSettings(formSettings);

    jquery("body").on("click", `button[data-automationid='FieldRenderer-title']`,  (e) => {
      this.selectedRow = jquery(e.target).parents().closest("div[data-automationid='DetailsRowCell']");
      this.selectedRow.trigger("click");
      
      //e.stopPropagation();
    });

    if (event.selectedRows.length > 0) {
      
      const contentType = event.selectedRows[0].getValueByName("ContentType");
      const editForms = formSettings.filter(i=>i.ContentTypeName===contentType && i.Form==="Edit");
      const displayForms = formSettings.filter(i=>i.ContentTypeName===contentType && i.Form==="Display");
      
      if(editForms.length>0){
        
        this.ovverideClick=true;
        this.overrideOnClick("Edit",editForms[0].OpenIn,editForms[0].RedirectURL,editForms[0].Parameters);        
      }
      else{
        this.ovverideClick=false;
      }

      if(displayForms.length>0){
        this.ovverideClick=true;
        this.overrideOnClick("Open",displayForms[0].OpenIn,displayForms[0].RedirectURL,displayForms[0].Parameters);
        if(this.selectedRow){
          window.open(`http://google.com`, "_blank");
          this.selectedRow=null;
        }
        //this.selectedRow.find("button[data-automationid='FieldRenderer-title']").click();
      }
      else
        this.ovverideClick=false;
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
          this.overrideOnClick(form.ContentTypeName,form.OpenIn,form.RedirectURL,form.Parameters);
          break;
      }
    });
  }

  private overrideOnClick(tagName:string,openIn:string,redirectURL:string,tokens:string){
    jquery("body").on("click", `button[name='${tagName}']`,  (e)=> {
      if(this.ovverideClick){
        switch (openIn) {
          case "Current Window":
            window.location.href = `${redirectURL}?${this.replaceTokens(tokens)}`;
            break;
          case "New Tab":
            window.open(`${redirectURL}?${this.replaceTokens(tokens)}`, "_blank");
            break;
        }
        e.stopPropagation();
      }
      
    });
  }

  private replaceTokens(tokens:string){
    if(!tokens)
      return "";
    return tokens.replace("{ListId}",String(this.context.pageContext.list.id))
            .replace("{WebUrl}",this.context.pageContext.web.absoluteUrl)
            .replace("{SiteUrl}",this.context.pageContext.site.absoluteUrl)
            .replace("{UserLoginName}",this.context.pageContext.user.loginName)
            .replace("{UserDisplayName}",this.context.pageContext.user.displayName)
            .replace("{UserEmail}",this.context.pageContext.user.email);
  }
}
