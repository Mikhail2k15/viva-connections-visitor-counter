import { ReactPlugin } from "@microsoft/applicationinsights-react-js";
import { ApplicationInsights, SeverityLevel } from "@microsoft/applicationinsights-web";
import { ILogEntry, ILogListener, LogLevel } from "@pnp/logging";

export default class AppInsightsListener implements ILogListener {
    constructor(iKey: string){
        this.iKey = iKey;
        this.appInsights = this.getAppInsights();
    }

    public log(entry: ILogEntry): void {
        if (entry.level === LogLevel.Error){
            this.appInsights.trackException({
                error: new Error(entry.message),
                severityLevel: SeverityLevel.Error
            });
        }
        else if (entry.level === LogLevel.Warning){
            this.appInsights.trackTrace({ message: entry.message, severityLevel: SeverityLevel.Warning});
        }
        else if (entry.level === LogLevel.Verbose){
            this.appInsights.trackTrace({ message: entry.message, severityLevel: SeverityLevel.Verbose});
        }
        else if (entry.level === LogLevel.Info){            
            this.appInsights.trackEvent({ name: entry.message});
        }        
    }

    public trackEvent(name: string): void{
        this.appInsights.trackEvent({name: name});
    }

    private getAppInsights(): ApplicationInsights {
        const reactPlugin = new ReactPlugin();
        const appInsights = new ApplicationInsights({
            config: {
                maxBatchInterval: 0,
                instrumentationKey: this.iKey,
                extensions: [reactPlugin]
            }
        });
        appInsights.loadAppInsights();
        return appInsights;
    }

    private appInsights: ApplicationInsights;
    private iKey: string;
}