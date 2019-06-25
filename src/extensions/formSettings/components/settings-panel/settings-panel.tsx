import * as React from 'react';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';
import ISettingsPanelState from './settings-panel-state';
import ISettingsPanelProps from './settings-panel-props';
import styles from './settings-panel.module.scss';
import { isArray } from '@pnp/common';
import {
    Dropdown, Label, TextField, Toggle, IDropdownOption
} from 'office-ui-fabric-react';
import ListService from '../../services/list-service';
import { ChoiceGroup, IChoiceGroupOption } from 'office-ui-fabric-react/lib/ChoiceGroup';
import IFormItem from '../models/form-item';

export default class SettingsPanel extends React.Component<ISettingsPanelProps, ISettingsPanelState>{
    private listService: ListService;
    private formOptions = [
        { key: "Display", text: "Display" },
        { key: "New", text: "New" },
        { key: "Edit", text: "Edit" }
    ];
    constructor(props) {
        super(props);
        this.state = {
            contentTypes: [],
            formSettings: [],
            form: {},
            showTemplatePanel: true
        };
        this.listService = new ListService();
        this._onRenderFooterContent = this._onRenderFooterContent.bind(this);
    }

    public async componentDidMount() {
        let formSettings = await this.listService.getFormSettings(this.props.listId);
        const contentTypes = await this.listService.getListContentTypes(this.props.listId);

        this.setState({
            contentTypes: contentTypes.filter(t => t.Name != "Folder").map(t => ({ key: t.Id.StringValue, text: t.Name })),
            formSettings: formSettings
        });
    }

    public render() {
        return (
            <div className={styles.SettingsPanel}>
                <Panel isOpen={this.props.showPanel}
                    onDismissed={() => this.props.setShowPanel(false)}
                    type={PanelType.medium} headerText="Form Settings"
                    onRenderFooterContent={this._onRenderFooterContent}>
                    <Label>Content type:</Label>
                    <Dropdown onChanged={this._onDropDownChanged} placeholder="Select content type..." options={this.state.contentTypes} />
                    {
                        this.state.form.ContentTypeName &&
                        <Dropdown selectedKey={this.state.form.Form} onChanged={this._onFormDropDownChanged} label="Form:" placeholder="Select form..." options={this.formOptions} />
                    }

                    <Toggle
                        label="Enabled"
                        onText="Yes"
                        offText="No"
                        checked={this.state.form.Enabled}
                        onChanged={this._enabledToggleChange}
                        hidden={this.state.form.Form === undefined}
                    />
                    {
                        this.state.form.Enabled &&
                        <TextField label="Redirect URL:" value={this.state.form.RedirectURL} onChanged={this._onUrlChanged} />
                    }
                    {
                        this.state.form.Enabled &&
                        <ChoiceGroup
                        className="defaultChoiceGroup"
                        selectedKey={this.state.form.OpenIn}
                        options={[
                            {
                                key: 'Current Window',
                                text: 'Current Window'
                            },
                            {
                                key: 'New Tab',
                                text: 'New Tab'
                            },
                            {
                                key: 'Dialog',
                                text: 'Dialog'
                            }
                        ]}
                        onChanged={this._onChoiceChanged}
                        label="Open in"
                    />
                    }
                    
                </Panel>
            </div>
        );
    }

    private _onRenderFooterContent = (): JSX.Element => {
        return (
            <div>
                <PrimaryButton onClick={this._saveTemplate} style={{ marginRight: '8px' }}>Save</PrimaryButton>
            </div>
        );
    }

    private _onDropDownChanged = (option: IDropdownOption, index?: number) => {

        this.setState({
            form: {
                ContentTypeName: option.text,
                Form: null,
                Enabled: false
            }
        }
        );
    }

    private _onFormDropDownChanged = (option: IDropdownOption, index?: number) => {


        const forms: IFormItem[] = this.state.formSettings.filter(ct => ct.Form === option.text && ct.ContentTypeName === this.state.form.ContentTypeName);

        if (forms.length > 0) {

            const form = forms[0] as IFormItem;
            this.setState({
                form
            });
        }
        else {
            this.setState(prevState => (
                {
                    form: {
                        ContentTypeName: prevState.form.ContentTypeName,
                        Enabled: false,
                        Form: option.text
                    }
                }
            )
            );
        }
    }

    private _enabledToggleChange = (value) => {
        this.setState(
            {
                form: {
                    ...this.state.form,
                    Enabled: value
                }
            }
        );
    }

    private _onUrlChanged = (value) => {
        this.setState({
            form: {
                ...this.state.form,
                RedirectURL: value
            }
        });
    }

    private _onChoiceChanged = (option: IChoiceGroupOption, evt?: React.FormEvent<HTMLElement | HTMLInputElement>): void => {

        this.setState({
            form: {
                ...this.state.form,
                OpenIn: option.text
            }
        });
    }

    private _saveTemplate = () => {
        const { form } = this.state;
        const formObject: IFormItem = {
            Id: form.Id,
            Title: this.props.listId,
            Enabled: form.Enabled,
            ContentTypeName: form.ContentTypeName,
            Form: form.Form,
            OpenIn: form.OpenIn,
            RedirectURL: form.RedirectURL
        };

        const forms: IFormItem[] = this.state.formSettings.filter(ct => ct.Form === form.Form && ct.ContentTypeName === form.ContentTypeName);

        if (forms.length > 0) {
            this.listService.UpdateForm(form);
            let newFormSettings = this.state.formSettings;
            newFormSettings = newFormSettings.map(f => {
                return f.Id === formObject.Id ? formObject : f;
            });
            this.setState({
                formSettings: newFormSettings
            });

        }
        else {
            // Add new form settings

            console.log(formObject);
            this.listService.SaveForm(formObject);

            this.setState(prevState => ({
                formSettings: prevState.formSettings.concat(formObject)
            }));
        }

    }

}