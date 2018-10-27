import { Text } from '@microsoft/sp-core-library';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { Web } from 'sp-pnp-js';
import { IFieldConfiguration } from '../../../lib/webparts/lookup/components/IFieldConfiguration';

export class ListService {

    private spHttpClient: SPHttpClient;
    private web: Web;

    constructor(spHttpClient: SPHttpClient, web: Web) {
        this.web = web;
        this.spHttpClient = spHttpClient;
    }

    public getListsFromWeb(webUrl: string): Promise<Array<{url: string, title: string}>> {
        return new Promise<Array<{url: string, title: string}>>((resolve, reject) => {
            // const endpoint = Text.format("{0}/_api/web/lists?$select=Title,RootFolder/ServerRelativeUrl&$filter=(IsPrivate eq false) and (IsCatalog eq false) and (Hidden eq false)&$expand=RootFolder", webUrl);
            // this.spHttpClient.get(endpoint, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
            //     if (response.ok) {
            //         response.json().then((data: any) => {
            //             const listTitles: Array<{url: string, title: string}> = data.value.map((list) => {
            //                     return {url: list.RootFolder.ServerRelativeUrl, title: list.Title};
            //                 });
            //             resolve( listTitles.sort( (a, b) => a.title.localeCompare(b.title)) );
            //         })
            //         .catch((error) => { reject(error); });
            //     } else {
            //         reject(response);
            //     }
            // })
            // .catch((error) => { reject(error); });

            this.web.lists.get().then((data: any) => {
                const listTitles: Array<{url: string, title: string}> = data.map((list) => {
                    return {url: list.Title, title: list.Title};
                });
            resolve( listTitles.sort( (a, b) => a.title.localeCompare(b.title)) );
            });
        });
    }

    public getFields(listTitle: string): Promise<Array<IFieldConfiguration>> {
        return this.web.lists.getByTitle(listTitle).fields.get();
    }
}
