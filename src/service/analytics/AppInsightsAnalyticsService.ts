import { HttpClient, IHttpClientOptions, HttpClientResponse } from "@microsoft/sp-http";
import { Logger, LogLevel } from "@pnp/logging";
import { TimeSpan } from "./TimeSpan";

interface QueryResponse {
    tables: Table[];
}
  
interface Table {
    name: string;
    columns: Column[];
    rows: string[][];
}
  
interface Column {
    name: string;
    type: string;
}

export default class AppInsightsAnalyticsService {
    private appInsightsEndpoint: string = 'https://api.applicationinsights.io/v1/apps';
    private httpClient: HttpClient;
    private httpClientOptions: IHttpClientOptions;
    private requestHeaders: Headers = new Headers(); 
    
    constructor(httpClient: HttpClient, appId: string, appKey: string){
        this.httpClient = httpClient;
        this.appInsightsEndpoint += `/${appId}`;
        
        this.requestHeaders.append('Content-type', 'application/json; charset=utf-8');
        this.requestHeaders.append('x-api-key', appKey);
        this.httpClientOptions = { headers: this.requestHeaders };
    }
    
    // {"tables":[{"name":"PrimaryResult","columns":[{"name":"dcount_session_Id","type":"long"}],"rows":[[4]]}]}

    private executeQuery = async (queryUrl: string): Promise<QueryResponse> => {
        const response: HttpClientResponse = await this.httpClient.get(queryUrl, HttpClient.configurations.v1, this.httpClientOptions);
        return await response.json();
    }

    public getSingleNumberQueryResultAsync = async (query: string, timespan?: TimeSpan): Promise<number>=>{
        Logger.log({ message: timespan, level: LogLevel.Verbose});
        const queryUrl: string = timespan ? `timespan=${timespan}&query=${encodeURIComponent(query)}` : `query=${encodeURIComponent(query)}`;
        const url: string = this.appInsightsEndpoint + `/query?${queryUrl}`; 

        const resp: QueryResponse = await this.executeQuery(url);
        let result: number = 0;
        if (resp.tables.length > 0){
            //?.length === 1 ? resultToday[0] : 0
            result = parseInt(resp.tables[0].rows[0][0]);
        }
        return result;
    }


    public getQueryResult = (query: string, timespan?: TimeSpan): string =>{
        Logger.log({ message: timespan, level: LogLevel.Verbose});
        const queryUrl: string = timespan ? `timespan=${timespan}&query=${encodeURIComponent(query)}` : `query=${encodeURIComponent(query)}`;
        const url: string = this.appInsightsEndpoint + `/query?${queryUrl}`; 

        let result: string = '0';
        this.executeQuery(url).then((resp: QueryResponse) => {
            if (resp.tables.length > 0){
                result = resp.tables[0].rows[0][0];
            }
            return result;
        }).catch(()=> { return '0'}  
        ); 
        
        return result;
    }
}