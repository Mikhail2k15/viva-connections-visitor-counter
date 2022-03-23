import AppInsightsService from "./AppInsightsService";
import { TimeSpan } from "./TimeSpan";


export default class VivaConnectionsInsights {
    public static async getTodaySessions(service: AppInsightsService): Promise<any[]> {
        const uniqueSessions: string = "customEvents | summarize dcount(session_Id)";        
        return await service.getQueryResultAsync(uniqueSessions, TimeSpan['1 day']);
    }

    public static async getMonthlySessions(service: AppInsightsService) {
        const uniqueSessions: string = "customEvents | summarize dcount(session_Id)";        
        return await service.getQueryResultAsync(uniqueSessions, TimeSpan['30 day']);
    }

    public static async getMobileSessions(service: AppInsightsService, timeSpan: TimeSpan) {
        const queryMobile: string = "customEvents | where name == 'Mobile' | summarize dcount(session_Id)";       
        return await service.getQueryResultAsync(queryMobile, timeSpan);
    }

    public static async getDesktopSessions(service: AppInsightsService, timeSpan: TimeSpan) {
        const queryDesktop: string = "customEvents | where client_Browser startswith 'Electron' | summarize dcount(session_Id)";  
        return await service.getQueryResultAsync(queryDesktop, timeSpan);
    }

    public static async getWebSessions(service: AppInsightsService, timeSpan: TimeSpan) {
        const queryWeb: string = "customEvents | extend web = tostring(customDimensions['ancestorOrigins']) | where name == 'WebView' and web contains 'teams.microsoft.com' | summarize dcount(session_Id)";   
        return await service.getQueryResultAsync(queryWeb, timeSpan);
    }

}