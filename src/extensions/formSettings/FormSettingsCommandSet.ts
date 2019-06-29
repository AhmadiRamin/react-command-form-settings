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
  private ovverideClick = false;
  private selectedRow = null;
  private itemId:number;
  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized FormsSettingsCommandSet');
    
    return Promise.resolve();
  }

  @override
  public async onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): Promise<void> {
    const formSettings = await this.listService.getEnabledFormSettings(String(this.context.pageContext.list.id));
    const newForms = formSettings.filter(i=>i.FormType==="New");
    this.ovverideNewFormSettings(newForms);
    
    $("body").on("click", `button[data-automationid='FieldRenderer-title']`,  (e) => {
      
      this.selectedRow = $(e.target).parents().closest("div[data-automationid='DetailsRow']");
      this.selectedRow.trigger("click");
    });

    if (event.selectedRows.length > 0) {
      
      const contentType = event.selectedRows[0].getValueByName("ContentType");
      const editForms = formSettings.filter(i=>i.ContentTypeName===contentType && i.FormType==="Edit");
      const displayForms = formSettings.filter(i=>i.ContentTypeName===contentType && i.FormType==="Display");
      this.itemId = event.selectedRows[0].getValueByName("ID");

      if(editForms.length>0){
        
        this.ovverideClick=true;
        this.overrideOnClick("Edit",editForms[0]);        
      }
      else{
        this.ovverideClick=false;
      }
      if(displayForms.length>0){
        this.ovverideClick=true;
        this.overrideOnClick("Open",displayForms[0]);
        if(this.selectedRow){
          this.selectedRow=null;
          this.redirect(displayForms[0]);
          $(document.body.firstChild).trigger("click");
          
        }
      }
      else{
        this.selectedRow=null;
        this.ovverideClick=false;
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

  private async ovverideNewFormSettings(formSettings: IFormItem[]) {
    // Ovveride if only one content type exists in the list
    this.overrideOnClick("New",formSettings[0]);
    formSettings.map(form => {
          this.overrideOnClick(form.ContentTypeName,form);
    });
    
  }

  private overrideOnClick(tagName:string,settings:IFormItem){
    $("body").on("click", `button[name='${tagName}']`,  (e)=> {
      if(this.ovverideClick || tagName==="New"){
        this.redirect(settings);
        return e.stopPropagation();
      }
      
    });
  }

  private redirect(settings:IFormItem){
    const {OpenIn,RedirectURL,Parameters} = settings;
    switch (OpenIn) {
      case "Current Window":
        window.location.href = `${RedirectURL}?${this.replaceTokens(Parameters)}`;
        break;
      case "New Tab":
        window.open(`${RedirectURL}?${this.replaceTokens(Parameters)}`, "_blank");
        break;
    }
  }

  private replaceTokens(tokens:string){
    if(!tokens)
      return "";
    return tokens.replace("{ListId}",String(this.context.pageContext.list.id))
            .replace("{WebUrl}",this.context.pageContext.web.absoluteUrl)
            .replace("{SiteUrl}",this.context.pageContext.site.absoluteUrl)
            .replace("{UserLoginName}",this.context.pageContext.user.loginName)
            .replace("{UserDisplayName}",this.context.pageContext.user.displayName)
            .replace("{UserEmail}",this.context.pageContext.user.email)
            .replace("{ItemId}",String(this.itemId));
  }
}
