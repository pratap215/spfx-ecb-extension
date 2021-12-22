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
    public ceb: CustomEcbCommandSet;
    // private _dialog: ProgressDialogContent;
    private _confirmDialog: ConfirmDialog;
    private _multilingual: boolean;
    private _pageName: string | undefined;
    private _getUserPermissions: boolean | undefined;
    private _listId: string | undefined;
    private _listItemId: string | undefined;
    private _targetPageurl: string | undefined;
    private _sourcePageurl: string | undefined;
    private _sourcePageId: string | undefined;

    private _sPTranslationSourceItemId: Guid | undefined;
    private _sPTranslationLanguage: string | undefined;
    private _sPTranslatedLanguages: Array<string> | undefined;
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

            }}>Yes
            </PrimaryButton>
            <DefaultButton onClick={() => {
        return Promise.resolve(true);            

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
            // labelname={this.labelname}
            // description={this.description}

        />,{
            context: "context"
        }, this.domElement);
    }

    public getConfig(): IDialogConfiguration {
        return {
            isBlocking: true
        };
    }
}