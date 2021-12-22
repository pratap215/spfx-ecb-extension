import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import { ProgressIndicator } from 'office-ui-fabric-react/lib/ProgressIndicator';
import { DefaultButton, DialogContent, DialogFooter, DialogType, IStackTokens, Label, PrimaryButton, Spinner, Stack } from 'office-ui-fabric-react';
import { Dialog as D1 } from 'office-ui-fabric-react'
// import { Spinner } from '@fluentui/react/lib/Spinner';
import CustomEcbCommandSet from '../customEcb/CustomEcbCommandSet';

import { sp } from '@pnp/sp';
import { SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration } from '@microsoft/sp-http';
import { Guid } from '@microsoft/sp-core-library';
import { IClientsidePage, ClientsidePageFromFile, ColumnControl, ClientsideText, ClientsideWebpart } from '@pnp/sp/clientside-pages';
import { environment } from '../../environments/environment';
import { ITranslationService } from '../../services/ITranslationService';
import { TranslationService } from '../../services/TranslationService';
import { Dialog } from '@microsoft/sp-dialog';
import { PageContext } from '@microsoft/sp-page-context';
interface IProgressDialogContentProps {
    DefaultProgress?: number;
    close: () => void;
    labelname: string;
    description: string;
    submit: (labelname: string) => void;

}

interface IProgressDialogContentState {
    Progress: number;
    labelname: string;
    description: string;
}
const stackTokens: IStackTokens = {
    childrenGap: 20,
    maxWidth: 250,
  };
export class ConfirmDialogContent extends React.Component<any, any>  {
    public labelName: string;
    constructor(props) {
        super(props);
        this.state = {
            showDialog: true,
        };
    }

    public componentDidMount() {
        // Sleep in loop
        // sp.web.lists.getByTitle('kkkk').items.getAll().then(res => {
        //   console.log(res[0]['ID']);
        //   this.setState({
        //     Progress: 0.5
        //   });
        //   console.log('hh');
        // });
        // commented this for loop for testing on 16/12
        // for (let i = 2; i < 11; i++) {
        //     setTimeout(() => {
        //         this.setState({
        //             Progress: i / 10
        //         });

        //         if (this.state.Progress == 1) {
        //             this.props.close();
        //         }

        //     }, 1000);
        // }
    }

    public render(){
        return (
            this.state.showDialog ?
            <>
                    <DialogContent
            type={DialogType.normal}
                title='Translation'
                subText={`You are about to overwrite the content on this page with \nan automatic translation of the original language. Please confirm`}
                showCloseButton={false}
                isMultiline={true}

            >
            <DialogFooter>
            <PrimaryButton onClick={() => {
                //this.ceb._onTranslate()
                this.props.submit("Yes");

            }}>Yes
            </PrimaryButton>
            <DefaultButton onClick={() => {
        this.props.submit("No");           

            }}>No
            </DefaultButton>
            </DialogFooter>
        </DialogContent>
            </>
            :
            null


        )
    }

}

export default class ConfirmDialog extends BaseDialog {
    public initprogress: number;
    public labelname: string;
    public description: string;

    constructor() {
        super({ isBlocking: true });
       
    }

    public render(): void {
        ReactDOM.render(<ConfirmDialogContent
            // DefaultProgress={this.initprogress}
            // close={this.close}
            labelname={this.labelname}
            submit={ this._submit }
            // description={this.description}

        />, this.domElement);
    }

    public getConfig(): IDialogConfiguration {
        return {
            isBlocking: true
        };
    }
    protected onAfterClose(): void {
        super.onAfterClose();
        
        // Clean up the element for the next dialog
        ReactDOM.unmountComponentAtNode(this.domElement);
    }
    private _submit = (labelName: string) => {
        this.labelname = labelName;
        this.close();
    }
}