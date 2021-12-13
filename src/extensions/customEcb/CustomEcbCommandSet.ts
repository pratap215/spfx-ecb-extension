import { override } from '@microsoft/decorators';
import { Log, Guid, UrlQueryParameterCollection } from '@microsoft/sp-core-library';
import {
    BaseListViewCommandSet,
    Command,
    IListViewCommandSetListViewUpdatedParameters,
    IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'CustomEcbCommandSetStrings';


import { ILanguage } from "../../models/ILanguage";
import { INavigation } from "@pnp/sp/navigation";
import { IContextualMenuItem } from "office-ui-fabric-react/lib/ContextualMenu";
import { Layer } from "office-ui-fabric-react/lib/Layer";
import { MessageBar, MessageBarType } from "office-ui-fabric-react/lib/MessageBar";
import { Overlay } from "office-ui-fabric-react/lib/Overlay";
import { IDetectedLanguage } from "../../models/IDetectedLanguage";
import * as _ from "lodash";
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/navigation";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/features";
import { ColumnControl, ClientsideText, ClientsideWebpart, IClientsidePage, CreateClientsidePage, ClientsidePageFromFile } from "@pnp/sp/clientside-pages";
import { ITranslationResult } from "../../models/ITranslationResult";
import { Navigation } from "@pnp/sp/navigation";

import { ITranslationService } from "../../services/ITranslationService";
import { TranslationService } from "../../services/TranslationService";
import { environment } from '../../environments/environment';
import { SPPermission } from '@microsoft/sp-page-context';
//import ProgressDialog from '../components/ProgressDialog';
/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ICustomEcbCommandSetProperties {
    targetUrl: string;
}

const LOG_SOURCE: string = 'CustomEcbCommandSet';

export default class CustomEcbCommandSet extends BaseListViewCommandSet<ICustomEcbCommandSetProperties> {

    private _pageName: string | undefined;
    private _multilingual: boolean;
   // private dialog: ProgressDialog;

    @override
    public async onInit(): Promise<void> {
        Log.info(LOG_SOURCE, 'Initialized CustomEcbCommandSet');
        this._multilingual = await this.getMultiLingualFeatureEnabled();
        return Promise.resolve();
    }

    @override
    public async onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): Promise<void> {
        const compareOneCommand: Command = this.tryGetCommand('ShowDetails');
        const validPage = (pageName): boolean => {
            return pageName.slice(10, pageName.lastIndexOf('/')).length > 0 && pageName.indexOf('.aspx') !== -1 ? true : false;
        };
        if (compareOneCommand) {
            if (event.selectedRows.length === 1) {             
                const pageName: string = event.selectedRows[0].getValueByName("FileRef");
                compareOneCommand.visible = validPage(pageName) && this._multilingual && event.selectedRows.length === 1;
            }
        }
    }

    @override
    public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
        switch (event.itemId) {
            case 'ShowDetails':
                console.log('onExecute start');
                this._pageName = event.selectedRows[0].getValueByName('FileLeafRef');
                if (confirm('Are you sure you want to translate this page[' + this._pageName + ']')) {
                    this._onTranslate('de');
                }
                break;
            default:
                throw new Error('Unknown command');
        }
    }


    private _onTranslate = (languagecode: string): void => {
        console.log('_onTranslate start');

        (async () => {
            try {
                console.log('Copying......... ');
                const sourceRelativePageUrl: string = '/SitePages/' + this._pageName;
                let sourcepage: IClientsidePage = undefined;
                try {
                    sourcepage = await ClientsidePageFromFile(sp.web.getFileByServerRelativeUrl(sourceRelativePageUrl));
                } catch (error) {
                    console.dir(error);
                    console.log('source page not found ' + this._pageName);
                    Dialog.alert('Original page [' + this._pageName + '] not exists.Contact Admin');
                    return;
                }
                console.log('async/await source -> ', sourcepage);
               
                if (sourcepage != undefined) {

                        const targetRelativePageUrl: string = '/SitePages/' + languagecode + '/' + this._pageName;
                        const targetpage = await ClientsidePageFromFile(sp.web.getFileByServerRelativeUrl(targetRelativePageUrl));
                        await sourcepage.copyTo(targetpage, true);

                        console.log('Copy Completed.......');

                        const translationService: ITranslationService = environment.config.regionSpecifier
                            ? new TranslationService(this.context.httpClient, environment.config.translatorApiKey, `-${environment.config.regionSpecifier}`)
                            : new TranslationService(this.context.httpClient, environment.config.translatorApiKey);

                        Dialog.alert(`Starting Translation............`);

                        await new Promise(resolve => setTimeout(resolve, 5000));

                        sp.web.loadClientsidePage(targetRelativePageUrl).then(async (clientSidePage: IClientsidePage) => {

                            try {
                                console.log('translation started');

                                var clientControls: ColumnControl<any>[] = [];
                                clientSidePage.findControl((c) => {
                                    if (c instanceof ClientsideText) {
                                        clientControls.push(c);
                                    }
                                    else if (c instanceof ClientsideWebpart) {
                                        clientControls.push(c);
                                    }
                                    return false;
                                });

                                await this._alltranslateClientSideControl(translationService, clientControls, languagecode);

                                //const nav = sp.web.navigation.topNavigationBar;
                                //Dialog.alert(nav.length.toString());
                                //const childrenData = await nav.getById(1).children();
                                //await nav.getById(1).update({
                                //    Title: "A new title",
                                //});

                                //clientSidePage.title = this._getTranslatedText(clientSidePage.title, languagecode, false);

                                clientSidePage.save();

                                console.log('translation complete');

                                Dialog.alert(`Translation Completed........`);

                            } catch (error) {
                                console.dir(error);
                                

                            }
                        }).catch((error: Error) => {
                            console.dir(error);
                            
                        });
                    
                }

            } catch (err) {
                console.dir('aynsc error');
                console.log(err);
                
            }

        })();


    }

    private _alltranslateClientSideControl = async (translationService: ITranslationService, clientsideControls: ColumnControl<any>[], languagecode: string): Promise<void> => {
        try {
            for (const c of clientsideControls) {
                if (c instanceof ClientsideWebpart) {
                    if (c.data.webPartData?.serverProcessedContent?.searchablePlainTexts) {
                        let propkeys = Object.keys(c.data.webPartData?.serverProcessedContent?.searchablePlainTexts);
                        for (const key of propkeys) {
                            const propvalue = c.data.webPartData?.serverProcessedContent?.searchablePlainTexts[key];
                            if (propvalue) {
                                let translationResult = await translationService.translate(propvalue, languagecode, false);
                                const translatedText = translationResult.translations[0].text;
                                c.data.webPartData.serverProcessedContent.searchablePlainTexts[key] = translatedText;
                            }
                        }
                    }
                }
                else if (c instanceof ClientsideText) {
                    const propvalue = c.text;
                    if (propvalue) {
                        let translationResult = await translationService.translate(propvalue, languagecode, true);
                        const translatedText = translationResult.translations[0].text;
                        c.text = translatedText;
                    }
                }
            }
        } catch (err) {
            console.dir('aynsc error');
            console.log(err);
            
        }
    }

    //private _getTranslatedText = (text: string, languagecode: string, asHtml: boolean): string => {


    //    let translatedText: string = "";
    //    if (text) {
    //        // console.log('start');
    //        const translationService: ITranslationService = environment.config.regionSpecifier
    //            ? new TranslationService(this.context.httpClient, environment.config.translatorApiKey, `-${environment.config.regionSpecifier}`)
    //            : new TranslationService(this.context.httpClient, environment.config.translatorApiKey);

    //        //TODO : uncomment the below code 
    //        //(async () => {

    //        //    let translationResult = await translationService.translate(text, languagecode, asHtml);
    //        //    translatedText = translationResult.translations[0].text

    //console.log('end');

    //return translatedText;
    //}
    //*************Function to get Multilingual Feature Enabled************************************* */
    public getMultiLingualFeatureEnabled = (): Promise<boolean> => {
        return new Promise<boolean>(async (resolve, reject) => {
            let features = await sp.web.features.select("DisplayName", "DefinitionId").get().then(f => {
                if (_.find(f, { "DisplayName": "MultilingualPages" })) {
                    return resolve(true);
                }
                else {
                    return resolve(false);
                }
                //test comment for push

            }).catch(error => {
                console.log(error);
                return reject(false);
            });
            return resolve(false);
        });

    }

}

export interface cTypedHash<T> {
    [key: string]: T;
}

