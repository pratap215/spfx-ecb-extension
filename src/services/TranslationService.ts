import { ITranslationService } from "./ITranslationService";
import { HttpClient, HttpClientConfiguration } from "@microsoft/sp-http";
import { ILanguage } from "../models/ILanguage";
import { IDictionary } from "../models/IDictionary";
import { ITranslatorLanguage } from "../models/ITranslatorLanguage";
import { IDetectedLanguage } from "../models/IDetectedLanguage";
import { ITranslationResult } from "../models/ITranslationResult";
import { IBreakSentenceResult } from "../models/IBreakSentenceResult";
import { SPHttpClient } from '@microsoft/sp-http';
export class TranslationService implements ITranslationService {

  private httpClient: HttpClient;
  private apiKey: string;
  private headers: Headers;
  private host: string;
  private sphttpclient: SPHttpClient;
  constructor(httpClient: HttpClient, apiKey: string, regionSpecifier: string = "") {
    this.httpClient = httpClient;
    this.apiKey = apiKey;
    this.host = `api${regionSpecifier}.cognitive.microsofttranslator.com`;
    this.headers = new Headers();
    this.headers.append("Content-type", "application/json");
    this.headers.append("Ocp-Apim-Subscription-Key", this.apiKey);
  }
  public async getPageMode(restApi: string): Promise<string> {
    console.log("");
    console.log('getPageMode :' + restApi);
    try {
      //const restApi = `${this.context.pageContext.web.absoluteUrl}/_api/sitepages/pages(${pageId})/checkoutpage`;
      const result = await this.sphttpclient.post(restApi, SPHttpClient.configurations.v1, {});

      if (!result.ok) {
        console.log('failed getPageMode');
        const resultData: any = await result.json();
        console.log(resultData.error);
       // Dialog.alert(resultData.error.message);
        return resultData.error.message;
      }
      else {
        console.log("success getPageMode");
        const data: any = await result.json();
        console.log(data);
        return "";
      }
    } catch (e) {
      console.log('error getPageMode');
      console.log(e);
      return "";
    }
  }
  public async getAvailableLanguages(supportedLanguages: string[]): Promise<ILanguage[]> {
    const httpClient = this.httpClient;
    const path: string = "languages?api-version=3.0&scope=dictionary";

    const result = await httpClient.get(
      `https://${this.host}/${path}`,
      new HttpClientConfiguration({}),
      { headers: this.headers }
    );

    if (!result.ok) {
      const resultData: any = await result.json();
      throw new Error(resultData.error.message);
    }

    const translatorLanguages: IDictionary<ITranslatorLanguage> = (await result.json()).dictionary;
    const languages: ILanguage[] = supportedLanguages.map((languageCode: string) => {
      if (translatorLanguages[languageCode]) {
        return {
          label: translatorLanguages[languageCode].nativeName,
          code: languageCode
        };
      }
    });

    return languages;
  }

  public async detectLanguage(text: string): Promise<IDetectedLanguage> {
    const httpClient = this.httpClient;
    const path: string = "detect?api-version=3.0";

    const body: string = JSON.stringify([{ Text: text }]);

    const result = await httpClient.post(
      `https://${this.host}/${path}`,
      new HttpClientConfiguration({}),
      {
        headers: this.headers,
        body: body
      }
    );

    if (!result.ok) {
      const resultData: any = await result.json();
      throw new Error(resultData.error.message);
    }

    const detectLanguageInfo: IDetectedLanguage[] = await result.json();
    if (detectLanguageInfo.some((langInfo: IDetectedLanguage) => langInfo.score >= 0.8 && langInfo.isTranslationSupported)) {
      return detectLanguageInfo.filter((langInfo: IDetectedLanguage) => langInfo.score >= 0.8 && langInfo.isTranslationSupported)[0];
    }

    return null;
  }



    public async translate(sourceText: string, languageCode: string, asHtml: boolean): Promise<ITranslationResult> {

        console.log("In translate method");
    const httpClient = this.httpClient;
    const path: string = `translate?api-version=3.0&to=${languageCode}&textType=${asHtml ? "html" : "plain"}`;
        console.log("In translate method2");
    const body: string = JSON.stringify([{ Text: sourceText }]);

    const result = await httpClient.post(
      `https://${this.host}/${path}`,
      new HttpClientConfiguration({}),
      {
        headers: this.headers,
        body: body
      }
    );
        console.log("In translate method3");
    if (!result.ok) {
      const resultData: any = await result.json();
      throw new Error(resultData.error.message);
    }

    const translationInfo: ITranslationResult[] = await result.json();

    if (translationInfo.length > 0) {
      return translationInfo[0];
    } else {
      return null;
    }
  }

  public async breakSentence(sourceText: string): Promise<IBreakSentenceResult> {
    const httpClient = this.httpClient;
    const path: string = `breaksentence?api-version=3.0`;

    const body: string = JSON.stringify([{ Text: sourceText }]);

    const result = await httpClient.post(
      `https://${this.host}/${path}`,
      new HttpClientConfiguration({}),
      {
        headers: this.headers,
        body: body
      }
    );

    if (!result.ok) {
      const resultData: any = await result.json();
      throw new Error(resultData.error.message);
    }

    const breakSentenceInfo: IBreakSentenceResult[] = await result.json();

    if (breakSentenceInfo.length > 0) {
      return breakSentenceInfo[0];
    } else {
      return null;
    }
  }
}
