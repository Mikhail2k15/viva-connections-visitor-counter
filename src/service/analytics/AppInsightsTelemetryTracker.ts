
import {
    LogLevel,
    ILogListener,
    ILogEntry
} from "@pnp/logging";
import { ApplicationInsights, SeverityLevel } from '@microsoft/applicationinsights-web';
import { ReactPlugin } from '@microsoft/applicationinsights-react-js';
import { createBrowserHistory } from "history";

const APP_NAME = 'VISITOR_COUNTER_ACE';
const APP_VERSION = '2.0.0';

export class AppInsightsTelemetryTracker implements ILogListener {
    private static appInsightsInstance: ApplicationInsights;
    private static reactPluginInstance: ReactPlugin;

    private static BaseProperties = {
        CustomProps: {
            ancestorOrigins: (window && window.location && window.location.ancestorOrigins) ? window.location.ancestorOrigins : "UNKNOWN", 
            App_Name: APP_NAME, 
        }
    };

    constructor(instrumentationKey: string) {
        if (!AppInsightsTelemetryTracker.appInsightsInstance)
        AppInsightsTelemetryTracker.appInsightsInstance = AppInsightsTelemetryTracker.initializeAI(instrumentationKey);
    }

    public log(entry: ILogEntry): void {
        const msg = this.logMessageFormat(entry);
        if (entry.level === LogLevel.Off) {
            return;
        }

        if (AppInsightsTelemetryTracker.appInsightsInstance)
            switch (entry.level) {
                case LogLevel.Verbose:
                    AppInsightsTelemetryTracker.appInsightsInstance.trackTrace({ message: msg, severityLevel: SeverityLevel.Verbose });
                    break;
                case LogLevel.Info:
                    AppInsightsTelemetryTracker.appInsightsInstance.trackTrace({ message: msg, severityLevel: SeverityLevel.Information });
                    console.log({ Message: msg });
                    break;
                case LogLevel.Warning:
                    AppInsightsTelemetryTracker.appInsightsInstance.trackTrace({ message: msg, severityLevel: SeverityLevel.Warning });
                    console.warn({ Message: msg });
                    break;
                case LogLevel.Error:
                    AppInsightsTelemetryTracker.appInsightsInstance.trackException({ error: new Error(msg), severityLevel: SeverityLevel.Error });
                    console.error({ Message: msg });
                    break;
            }
    }

    public trackEvent(name: string, json?: string): void {
        if (AppInsightsTelemetryTracker.appInsightsInstance) {
            const obj = json ? {...AppInsightsTelemetryTracker.BaseProperties.CustomProps, ...JSON.parse(json)} : AppInsightsTelemetryTracker.BaseProperties.CustomProps;
            AppInsightsTelemetryTracker.appInsightsInstance.trackEvent({ name: name}, obj);
        }
    }

    private logMessageFormat(entry: ILogEntry): string {
        const msg: string[] = [];
        msg.push(entry.message);
    
        if (entry.data) {
            try {
                msg.push('Data: ' + JSON.stringify(entry.data));
            } catch (e) {
                msg.push(`Data: Error in stringify of supplied data ${e}`);
            }
        }
        return msg.join(' | ');
    }

    private static initializeAI(instrumentationKey?: string): ApplicationInsights {
        const browserHistory = createBrowserHistory({ basename: '' });
        AppInsightsTelemetryTracker.reactPluginInstance = new ReactPlugin();
        const appInsights = new ApplicationInsights({
            config: {
                maxBatchInterval: 0,
                instrumentationKey: instrumentationKey,
                namePrefix: AppInsightsTelemetryTracker.BaseProperties.CustomProps.App_Name,
                disableFetchTracking: false,
                disableAjaxTracking: true,
                extensions: [AppInsightsTelemetryTracker.reactPluginInstance],
                extensionConfig: {
                    [AppInsightsTelemetryTracker.reactPluginInstance.identifier]: { history: browserHistory }
                }
            }
        });

        appInsights.loadAppInsights();
        appInsights.context.application.ver = APP_VERSION;
        return appInsights;
    }
}