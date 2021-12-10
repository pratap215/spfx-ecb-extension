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

    @override
    public onInit(): Promise<void> {
        Log.info(LOG_SOURCE, 'Initialized CustomEcbCommandSet');
        return Promise.resolve();
    }

    @override
    public async onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): Promise<void> {

        console.log('====================================');
        console.log('onListViewUpdated');
        console.log('====================================');
        const compareOneCommand: Command = this.tryGetCommand('ShowDetails');
        const langFeature: boolean = await this.getMultiLingualFeatureEnabled();
        console.log(langFeature);
        if (compareOneCommand) {
            if (event.selectedRows.length == 1) {
                //let pagename = event.selectedRows[0].getValueByName('FileLeafRef');
                //Dialog.alert(pagename);

                    // This command should be hidden unless exactly one row is selected.
                    compareOneCommand.visible = event.selectedRows.length === 1 && langFeature ;
            }
        }
    }

    @override
    public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
        switch (event.itemId) {
            case 'ShowDetails':
                this._pageName = event.selectedRows[0].getValueByName('FileLeafRef');
                if (confirm('Are you sure you want to translate this page')) {
                    this._onTranslate('de');
                }
                else {
                    console.log("No translation");
                }

                break;
            default:
                throw new Error('Unknown command');
        }
    }



    private _onTranslate = (languagecode: string): void => {

        console.log('_onTranslate');

        //var siteUrl = this.context.pageContext.web.serverRelativeUrl;

        // var sourcePageUrl = siteUrl + "/SitePages/" + selectedPage;
        // var sourcePageUrl = siteUrl + "/SitePages/Home.aspx";
        //sourcePageUrl = 'https://8p5g5n.sharepoint.com/SitePages/Home.aspx';
        //console.log(sourcePageUrl);

        //var targetPageUrl = 'https://8p5g5n.sharepoint.com/SitePages/de/Home.aspx';

        //TODO we should use this._pageName
        //const sourceRelativePageUrl: string = '/SitePages/Home.aspx';
        //const targetRelativePageUrl: string = '/SitePages/de/Home.aspx';

        const sourceRelativePageUrl: string = '/SitePages/' + this._pageName;
        const targetRelativePageUrl: string = '/SitePages/' + languagecode + '/' + this._pageName;

        (async () => {

            Dialog.alert(this._pageName);
            try {
                console.log('Copying......... ');
                //const result = await sp.web.loadClientsidePage(deRelativePageUrl);
                const sourcepage = await ClientsidePageFromFile(sp.web.getFileByServerRelativeUrl(sourceRelativePageUrl));
                const targetpage = await ClientsidePageFromFile(sp.web.getFileByServerRelativeUrl(targetRelativePageUrl));
                //console.log('async/await source -> ', sourcepage);
                //console.log('async/await target -> ', targetpage);
                sourcepage.copyTo(targetpage, true);

                console.log('Copy Completed.......');

                Dialog.alert(`Copy Completed.......Starting Translation.`);

                await new Promise(resolve => setTimeout(resolve, 2000));

                //const deRelativePageUrl: string = '/SitePages/de/Home.aspx';

                sp.web.loadClientsidePage(targetRelativePageUrl).then(async (clientSidePage: IClientsidePage) => {

                    try {
                        console.log('translation started');
                        // Translate title

                        //clientSidePage.copy(sp.web, "home_page_de", "Home Page", true);

                        clientSidePage.title = this._getTranslatedText(clientSidePage.title, languagecode, false);

                        clientSidePage.findControl((c) => {
                            if (c instanceof ClientsideText) {
                                //Dialog.alert(c.text);
                                //const translatedText = this._getTranslatedText(c.text, languagecode,true);
                                //c.text = c.text + translatedText;
                            }
                            else if (c instanceof ClientsideWebpart) {

                                //const spt = c.data.webPartData?.serverProcessedContent?.searchablePlainTexts;
                                //let spt: cTypedHash<string> = c.data.webPartData?.serverProcessedContent?.searchablePlainTexts;
                                if (c.data.webPartData?.serverProcessedContent?.searchablePlainTexts) {
                                    let propkeys = Object.keys(c.data.webPartData?.serverProcessedContent?.searchablePlainTexts);
                                    //console.log("wait...");
                                    //console.log(keys.length + "    " + keys);
                                    propkeys.forEach(key => {
                                        const propvalue = c.data.webPartData?.serverProcessedContent?.searchablePlainTexts[key];
                                        const translatedText = this._getTranslatedText(propvalue, languagecode, false);
                                        c.data.webPartData.serverProcessedContent.searchablePlainTexts[key] = propvalue + translatedText;
                                        //console.log(spt[key])
                                    });
                                }
                            }
                            return false;
                        });

                        //const nav = sp.web.navigation.topNavigationBar;
                        //Dialog.alert(nav.length.toString());
                        //const childrenData = await nav.getById(1).children();
                        //await nav.getById(1).update({
                        //    Title: "A new title",
                        //});



                        console.log('translation complete');

                        clientSidePage.save();

                        Dialog.alert(`Translation Completed........`);

                    } catch (error) {
                        console.dir(error);

                    }
                }).catch((error: Error) => {
                    console.dir(error);

                });


            } catch (err) {
                console.dir('aynsc error');
                console.log(err);
            }
        })();


    }



    private _getTranslatedText = (text: string, languagecode: string, asHtml: boolean): string => {


        let translatedText: string = "";
        if (text) {
            // console.log('start');
            const translationService: ITranslationService = environment.config.regionSpecifier
                ? new TranslationService(this.context.httpClient, environment.config.translatorApiKey, `-${environment.config.regionSpecifier}`)
                : new TranslationService(this.context.httpClient, environment.config.translatorApiKey);

            //TODO : uncomment the below code 
            //translationService.translate(text, languagecode, false).then(translationResult =>
            //    translatedText=translationResult.translations[0].text
            //);

            //TODO remove below code.
            translatedText = "_ed";
            // console.log('end');
        }

        console.log('end');
        
        return translatedText;
    }

    public getMultiLingualFeatureEnabled = () : Promise<boolean> => {
        return new Promise<boolean>(async (resolve, reject) => {
            let features = await sp.web.features.select("DisplayName", "DefinitionId").get().then(f => {
                // console.log('====================================');
                // f.map(f => {
                //     console.log(f["DisplayName"] + " - " + f.DefinitionId);
                // });
                // console.log(f);
                // console.log('====================================');
                if(_.find(f, { "DisplayName": "MultilingualPages" })){
                    return resolve(true);
                }
                else{
                    return resolve(false);  
                }
            
        }).catch(error => {
            console.log(error);
            return reject(false);
        });
        return resolve(false)
    })

    }

}

export interface cTypedHash<T> {
    [key: string]: T;
}

