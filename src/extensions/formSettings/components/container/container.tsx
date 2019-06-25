import * as React from 'react';
import * as ReactDOM from 'react-dom';
import SettingsPanel from '../settings-panel/settings-panel';

class container {
    private showPanel:boolean = true;
    constructor(private listId){
        
    }
    public render() {
        const settingsPanel = (
            <SettingsPanel showPanel={this.showPanel} setShowPanel={this._setShowPanel} listId={this.listId} />
        );
        
        ReactDOM.render([settingsPanel],document.body.firstChild as Element);
    }
    public _setShowPanel = (showSettingsPanel: boolean): void => {
        this.showPanel=showSettingsPanel;
    }
}

export{
    container
};