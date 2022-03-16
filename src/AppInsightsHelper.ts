import { HttpClient, IHttpClientOptions, HttpClientResponse } from "@microsoft/sp-http";
import { Logger, LogLevel } from "@pnp/logging";



export default class AppInsightsHelper {
    public getRowRestApiResponse = async (queryUrl: string): Promise<any> => {
        let response: HttpClientResponse = await this.httpClient.get(queryUrl, HttpClient.configurations.v1, this.httpClientOptions);
        return await response.json();
    }

    public getQueryResponse = async (query: string, timespan?: TimeSpan): Promise<any[]>=>{
        Logger.log({ message: timespan, level: LogLevel.Verbose});
        let queryUrl: string = timespan ? `timespan=${timespan}&query=${encodeURIComponent(query)}` : `query=${encodeURIComponent(query)}`;
        let url: string = this.appInsightsEndpoint + `/query?${queryUrl}`; 

        let resp: any = await this.getRowRestApiResponse(url);
        let result: any[] = [];
        if (resp.tables.length > 0){
            result = resp.tables[0].rows;
        }
        return result;
    }

    constructor(httpClient: HttpClient, appId: string, appKey: string){
        this.httpClient = httpClient;
        this.appInsightsEndpoint += `/${appId}`;
        
        this.requestHeaders.append('Content-type', 'application/json; charset=utf-8');
        this.requestHeaders.append('x-api-key', appKey);
        this.httpClientOptions = { headers: this.requestHeaders };

        Logger.writeJSON({ 
            appInsightsEndpoint: this.appInsightsEndpoint, 
            appKey: appKey }, LogLevel.Info);
    }

    private httpClient: HttpClient;
    private httpClientOptions: IHttpClientOptions;
    private appInsightsEndpoint: string = 'https://api.applicationinsights.io/v1/apps';
    private requestHeaders: Headers = new Headers();
}

export enum TimeSpan {
    "1 hour" = "PT1H",
    "6 hours" = "PT6H",
    "12 hours" = "PT12H",
    "1 day" = "P1D",
    "3 days" = "P3D",
    "7 days" = "P7D",
    "15 days" = "P15D",
    "30 days" = "P30D",
    "45 days" = "P45D",
    "60 days" = "P60D",
    "75 days" = "P75D",
    "90 days" = "P90D",
}