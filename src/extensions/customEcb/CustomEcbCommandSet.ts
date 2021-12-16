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
import "@pnp/sp/site-users/web";
// import { SPPermission } from '@microsoft/sp-page-context'
import { ColumnControl, ClientsideText, ClientsideWebpart, IClientsidePage, CreateClientsidePage, ClientsidePageFromFile } from "@pnp/sp/clientside-pages";
import { ITranslationResult } from "../../models/ITranslationResult";
import { Navigation } from "@pnp/sp/navigation";

import { ITranslationService } from "../../services/ITranslationService";
import { TranslationService } from "../../services/TranslationService";
import { environment } from '../../environments/environment';
import { SPPermission } from '@microsoft/sp-page-context';
//import ProgressDialog from '../components/ProgressDialog';
import { SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration } from '@microsoft/sp-http';
/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
import ProgressDialogContent from './../components/ProgressDialog';
export interface ICustomEcbCommandSetProperties {
    targetUrl: string;
}

const LOG_SOURCE: string = 'CustomEcbCommandSet';

export default class CustomEcbCommandSet extends BaseListViewCommandSet<ICustomEcbCommandSetProperties> {

    constructor() {
        super();
        this._dialog = new ProgressDialogContent();
        // Log.info(LOG_SOURCE, 'Initialized CustomEcbCommandSet');
    }
    private _dialog: ProgressDialogContent;
    private _multilingual: boolean;
    private _pageName: string | undefined;
    private _getUserPermissions: boolean | undefined;
    private _listId: string | undefined;
    private _listItemId: string | undefined;
    private _targetPageurl: string | undefined;
    private _sourcePageurl: string | undefined;

    private _sPTranslationSourceItemId: Guid | undefined;
    private _sPTranslationLanguage: string | undefined;
    private _sPTranslatedLanguages: Array<string> | undefined;

    @override
    public async onInit(): Promise<void> {
        Log.info(LOG_SOURCE, 'Initialized CustomEcbCommandSet');
        this._multilingual = await this.getMultiLingualFeatureEnabled();
        this._getUserPermissions = await this.getUsersPermissions();
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
                compareOneCommand.visible = this._getUserPermissions && validPage(pageName) && this._multilingual && event.selectedRows.length === 1;
            }
        }
    }

    @override
    public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
        switch (event.itemId) {
            case 'ShowDetails':
                console.log('onExecute start');
                
                this._dialog.show();
                this._pageName = event.selectedRows[0].getValueByName('FileLeafRef');
                if (confirm('Are you sure you want to translate this page[' + this._pageName + ']')) {
                    
                    // ProgressDialogContent.show(dialog);
                    this._listId = this.context.pageContext.list.id.toString();
                    this._listItemId = event.selectedRows[0].getValueByName('ID').toString();

                    this._onTranslate();
                    // this._dialog.close();
                }
                break;
            default:
                throw new Error('Unknown command');
        }
    }


    private _onTranslate = (): void => {
        console.log('_onTranslate start');

        (async () => {
            try {

                const isValidTargetFile = await this.getTranslationPageMetaData();

                console.log(this._targetPageurl);

                if (isValidTargetFile == false) {
                    Dialog.alert('Page cannot be translated.Contact Admin');
                    return;
                }

               

                const isValidSourceFile = await this.getSourcePageMetaData(this._sPTranslationSourceItemId);

                if (isValidSourceFile == false) {
                    Dialog.alert('Original page not exists.Contact Admin');
                    return;
                }

                console.log('Copying......... ');
                // const sourceRelativePageUrl: string = '/SitePages/' + this._pageName;
                const sourceRelativePageUrl: string = this._sourcePageurl;
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

                    const languagecode: string = this._sPTranslationLanguage;

                    // const targetRelativePageUrl: string = '/SitePages/' + languagecode + '/' + this._pageName;
                    const targetRelativePageUrl: string = this._targetPageurl;
                    const targetpage = await ClientsidePageFromFile(sp.web.getFileByServerRelativeUrl(targetRelativePageUrl));
                    await sourcepage.copyTo(targetpage, true);

                    console.log('Copy Completed.......');

                    const translationService: ITranslationService = environment.config.regionSpecifier
                        ? new TranslationService(this.context.httpClient, environment.config.translatorApiKey, `-${environment.config.regionSpecifier}`)
                        : new TranslationService(this.context.httpClient, environment.config.translatorApiKey);

                    // Dialog.alert(`Starting Translation............ ` + languagecode);

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
                            this._dialog.close();

                        }
                    }).catch((error: Error) => {
                        console.dir(error);
                        this._dialog.close();

                    });

                }

            } catch (err) {
                console.dir('aynsc error');
                console.log(err);
                this._dialog.close();

            }
            this._dialog.close();
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
    //*************Function to check user's effective permissions************************************* */
    //Will work if user belongs to Owners or Members group => Manage list permissions
    //Promise can be removed, however doesn't harm if used with async
    public getUsersPermissions = (): Promise<boolean> => {
        return new Promise<boolean>(async (resolve, reject) => {
            try{
                let userHasPermissions: boolean = false;
                userHasPermissions = this.context.pageContext.list.permissions.hasPermission(SPPermission.manageLists);
                return resolve(userHasPermissions);
            }
            catch(error){
                console.log(error);
                return reject(false);
            }
        });

    }


    //Metadata start

    public async getTranslationPageMetaData(): Promise<boolean> {
        console.log('getTranslationPageMetaData');
        try {
            const absoluteurl = this.context.pageContext.web.absoluteUrl;
            const siteurl = `${absoluteurl}/_api/web/Lists/GetById('${this._listId}')/RenderListDataAsStream`;
          //  const siteurl = `https://8p5g5n.sharepoint.com/_api/web/Lists/GetById('${this._listId}')/RenderListDataAsStream`;
            const result = await this.context.spHttpClient.post(siteurl, SPHttpClient.configurations.v1, {
                body: JSON.stringify({
                    parameters: {
                        ViewXml: `<View Scope="RecursiveAll">
                  <ViewFields>
                    <FieldRef Name="_SPIsTranslation" />
                    <FieldRef Name="_SPTranslatedLanguages" />
                    <FieldRef Name="_SPTranslationLanguage" />
                    <FieldRef Name="_SPTranslationSourceItemId" />
                  </ViewFields>
                  <Query>
                    <Where>
                    <Eq>
                        <FieldRef Name="ID" />
                        <Value Type="Number">${this._listItemId}</Value>
                    </Eq>
                </Where>
                  </Query>
                  <RowLimit />
                </View>`
                    }
                })
            });

            if (!result.ok) {
                console.log('failed getTranslationPageMetaData');
                const resultData: any = await result.json();
                console.log(resultData.error);
                return false;
            }
            else {
                console.log("success getTranslationPageMetaData");
                const data: any = await result.json();
                // console.log(data);
                if (data && data.Row && data.Row.length > 0) {
                    const row = data.Row[0];
                    console.log("target page info");
                    console.log(row);
                    if (row["_SPIsTranslation"] == "Yes") {
                        //  this._sPTranslationSourceItemId = row["_SPTranslationSourceItemId"].toString().replace("{", "").replace("}", "").trim();
                        this._sPTranslationSourceItemId = row["_SPTranslationSourceItemId"].toString();
                        this._sPTranslationLanguage = row["_SPTranslationLanguage"];
                        this._targetPageurl = row["FileRef"];

                        //console.log(Object.keys(row));
                        return true;
                    }
                }
            }

        } catch (e) {
            console.log('error getTranslationPageMetaData');
            console.log(e);
            return false;
        }

        return false;
    }



    public async getSourcePageMetaData(pageid: Guid): Promise<boolean> {
        console.log("");
        console.log('getSourcePageMetaData :' + pageid);

        // const uniqid = "{9956AB6B-9C81-4448-88D3-634BC9536D34}";
        //var currentPageUrl = this.context.pageContext.site.serverRequestPath;

        //sp.web.lists.getByTitle("Site Pages").items.get().then((items: any[]) => {
        //   console.log(items[0]);
        //});

        //sp.web.lists.getById("${this._listId}").items.get().then((items: any[]) => {
        //    console.log(items[0]);
        //});

        //const siteAssetsList = await sp.web.lists.ensureSitePagesLibrary();
        //const r = await siteAssetsList.select("Title")();
        //    console.log(r);

        try {
           // const siteurl = `https://8p5g5n.sharepoint.com/_api/web/Lists/GetById('${this._listId}')/RenderListDataAsStream`;
            const absoluteurl = this.context.pageContext.web.absoluteUrl;
            const siteurl = `${absoluteurl}/_api/web/Lists/GetById('${this._listId}')/RenderListDataAsStream`;

            const result = await this.context.spHttpClient.post(siteurl, SPHttpClient.configurations.v1, {
                body: JSON.stringify({
                    parameters: {
                        ViewXml: `<View Scope="RecursiveAll">
                  <ViewFields>
                    <FieldRef Name="_SPIsTranslation" />
                    <FieldRef Name="_SPTranslatedLanguages" />
                    <FieldRef Name="_SPTranslationLanguage" />
                    <FieldRef Name="_SPTranslationSourceItemId" />
                  </ViewFields>
                  <Query>
                    <Where>
                    <Eq>
                        <FieldRef Name="UniqueId" />
                        <Value Type="Guid">${pageid}</Value>
                    </Eq>
                </Where>
                  </Query>
                  <RowLimit />
                </View>`
                    }
                })
            });

            if (!result.ok) {
                console.log('failed getSourcePageMetaData');
                const resultData: any = await result.json();
                console.log(resultData.error);
                return false;
            }
            else {
                console.log("success getSourcePageMetaData2");
                const data: any = await result.json();
                // console.log(data);
                if (data && data.Row && data.Row.length > 0) {
                    const row = data.Row[0];
                    console.log("source page info");
                    console.log(row);
                    this._sourcePageurl = row["FileRef"];
                    this._sPTranslatedLanguages = row["_SPTranslatedLanguages"];
                    console.log(this._sPTranslatedLanguages);
                    return true;
                }
            }

        } catch (e) {
            console.log('error getTranslationPageMetaData');
            console.log(e);
            return false;
        }

        return false;
    }

    private getLanguageName(code: string): string {
        console.log("getLanguageName " + code);
        const regionalLanguages = `{"ar-sa":"Arabic",
"az-latn-az":"Azerbaijani",
"eu-es":"Basque",
"bs-latn-ba":"Bosnian (Latin)",
"bg-bg":"Bulgarian",
"ca-es":"Catalan",
"zh-cn":"Chinese (Simplified)",
"zh-tw":"Chinese (Traditional)",
"hr-hr":"Croatian",
"cs-cz":"Czech",
"da-dk":"Danish",
"prs-af":"Dari",
"nl-nl":"Dutch",
"en-us":"English",
"et-ee":"Estonian",
"fi-fi":"Finnish",
"fr-fr":"French",
"gl-es":"Galician",
"de-de":"German",
"el-gr":"Greek",
"he-il":"Hebrew",
"hi-in":"Hindi",
"hu-hu":"Hungarian",
"id-id":"Indonesian",
"ga-ie":"Irish",
"it-it":"Italian",
"ja-jp":"Japanese",
"kk-kz":"Kazakh",
"ko-kr":"Korean",
"lv-lv":"Latvian",
"lt-lt":"Lithuanian",
"mk-mk":"Macedonian",
"ms-my":"Malay",
"nb-no":"Norwegian (Bokmål)",
"pl-pl":"Polish",
"pt-br":"Portuguese (Brazil)",
"pt-pt":"Portuguese (Portugal)",
"ro-ro":"Romanian",
"ru-ru":"Russian",
"sr-cyrl-rs":"Serbian (Cyrillic, Serbia)",
"sr-latn-cs":"Serbian (Latin)",
"sr-latn-rs":"Serbian (Latin, Serbia)",
"sk-sk":"Slovak",
"sl-si":"Slovenian",
"es-es":"Spanish",
"sv-se":"Swedish",
"th-th":"Thai",
"tr-tr":"Turkish",
"uk-ua":"Ukrainian",
"vi-vn":"Vietnamese",
"cy-gb":"Welsh"}`;

        const languageNames = JSON.parse(regionalLanguages);

        console.log("getLanguageName name " + languageNames["de-de"]);

        return languageNames[code.toLowerCase()];

    }
    //Metadata end


}

export interface cTypedHash<T> {
    [key: string]: T;
}

