import { override } from '@microsoft/decorators';
import { Log, Guid } from '@microsoft/sp-core-library';
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

import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/navigation";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { ColumnControl, ClientsideText, ClientsideWebpart, IClientsidePage, CreateClientsidePage } from "@pnp/sp/clientside-pages";
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

    @override
    public onInit(): Promise<void> {
        Log.info(LOG_SOURCE, 'Initialized CustomEcbCommandSet');

        sp.setup(this.context);


        return Promise.resolve();
    }

    @override
    public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
        const compareOneCommand: Command = this.tryGetCommand('ShowDetails');
        if (compareOneCommand) {
            if (event.selectedRows.length == 1) {
                //let pagename = event.selectedRows[0].getValueByName('FileLeafRef');
                //Dialog.alert(pagename);

                    // This command should be hidden unless exactly one row is selected.
                    compareOneCommand.visible = event.selectedRows.length === 1;
            }
        }
    }

    @override
    public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
        switch (event.itemId) {
            case 'ShowDetails':
                if (confirm('Are you sure you want to translate this page')) {
                    this._onTranslate('de');
                }
                else {
                    console.log("No translation");
                }
                //siteUrl: this.context.pageContext.web.absoluteUrl;
                //Dialog.alert(`Done`);
                break;
            default:
                throw new Error('Unknown command');
        }
    }



    private _onTranslate = (languagecode: string): void => {

        console.log('start');
       
        
        const relativePageUrl: string = '/SitePages/Home.aspx';

        //sp.web.loadClientsidePage(relativePageUrl).then(async (homePage: IClientsidePage) => {
        //    try {
        //        console.log(`Page 2 creating`);
        //        //let targetpage2 = await sp.web.loadClientsidePage('/SitePages/de/Home.aspx');
        //        //homePage.copyTo(targetpage2, true);

               

               
        //        console.log(`Page 2 created` + homePage.sections.length);

        //    } catch (error) {
        //        console.dir(error);
        //        Dialog.alert((error as Error).message);
        //    }
        //}).catch((error: Error) => {
        //    console.dir(error);
        //    Dialog.alert((error as Error).message);
        //});

        Dialog.alert(`Starting Translation........`);

        //sp.web.loadClientsidePage(relativePageUrl).then(async (page: IClientsidePage) => {

        //    try {
        //        //const pageCopy2 = await page.copy(sp.web, "/SitePages/de/Home.aspx", "De Page", true);
        //       // pageCopy2.save();
               
        //        console.log("page 2 created");

        //    } catch (error) {
        //        console.dir(error);
        //        Dialog.alert((error as Error).message);
        //    }
        //}).catch((error: Error) => {
        //    console.dir(error);
        //    Dialog.alert((error as Error).message);
        //});

        const deRelativePageUrl: string = '/SitePages/de/Home.aspx';

        sp.web.loadClientsidePage(deRelativePageUrl).then(async (clientSidePage: IClientsidePage) => {

            try {
                console.log('translation started');
                // Translate title
                // await this._translatePageTitle(clientSidePage.title, language.code);
                //  await this._translatePageNav(sp.web.navigation.topNavigationBar.toString, language.code);

                // Get all text controls
                var clientControls: ColumnControl<any>[] = [];
                clientSidePage.findControl((c) => {
                    if (c instanceof ClientsideText) {
                        //clientControls.push(c);
                        const translatedText = this._getTranslatedText(c.text, languagecode);
                        c.text = c.text + translatedText;
                       
                    }
                    else if (c instanceof ClientsideWebpart) {
                        //clientControls.push(c);
                        //const spt = c.data.webPartData?.serverProcessedContent?.searchablePlainTexts;
                        let spt: cTypedHash<string> = c.data.webPartData?.serverProcessedContent?.searchablePlainTexts;
                        if (c.data.webPartData?.serverProcessedContent?.searchablePlainTexts!=null ) {
                            let propkeys = Object.keys(c.data.webPartData?.serverProcessedContent?.searchablePlainTexts);
                            //console.log("wait...");
                            //console.log(keys.length + "    " + keys);
                            propkeys.forEach(key => {
                                const propvalue = c.data.webPartData?.serverProcessedContent?.searchablePlainTexts[key];
                                const translatedText = this._getTranslatedText(propvalue, languagecode);
                                //c.data.webPartData.serverProcessedContent.searchablePlainTexts[key] = propvalue + translatedText;
                                //console.log(spt[key])
                            });
                        }
                    }
                    return false;
                });

                //await this._alltranslateClientSideControl(clientControls, language.code);

                console.log('translation complete');
              
                clientSidePage.save();

                Dialog.alert(`Translation Completed........`);

            } catch (error) {
                console.dir(error);
                Dialog.alert((error as Error).message);
            }
        }).catch((error: Error) => {
            console.dir(error);
            Dialog.alert((error as Error).message);
        });


    }



    private _getTranslatedText = (text:string,languagecode: string): string => {

        console.log('start');
        let translatedText:string = "noValue";

        const translationService: ITranslationService = environment.config.regionSpecifier
            ? new TranslationService(this.context.httpClient, environment.config.translatorApiKey, `-${environment.config.regionSpecifier}`)
            : new TranslationService(this.context.httpClient, environment.config.translatorApiKey);

        //translationService.translate(text, languagecode, false).then(translationResult =>
        //    translatedText=translationResult.translations[0].text
        //);

        translatedText = "_de";

        console.log('end');
        return translatedText;
    }



}

export interface cTypedHash<T> {
    [key: string]: T;
}

