
import {
    LogLevel,
    ILogListener,
    ILogEntry
} from "@pnp/logging";
import { ApplicationInsights, SeverityLevel } from '@microsoft/applicationinsights-web';
import { ReactPlugin } from '@microsoft/applicationinsights-react-js';
import { createBrowserHistory } from "history";
import { CONST, _logEventFormat, _logMessageFormat } from "./Utilities";


export class AppInsightsLogListener implements ILogListener {
    private static _appInsightsInstance: ApplicationInsights;
    private static _reactPluginInstance: ReactPlugin;

    constructor(instrumentationKey: string) {
        if (!AppInsightsLogListener._appInsightsInstance)
        AppInsightsLogListener._appInsightsInstance = AppInsightsLogListener._initializeAI(instrumentationKey);
    }

    private static _initializeAI(instrumentationKey?: string): ApplicationInsights {
        const browserHistory = createBrowserHistory({ basename: '' });
        AppInsightsLogListener._reactPluginInstance = new ReactPlugin();
        const appInsights = new ApplicationInsights({
            config: {
                maxBatchInterval: 0,
                instrumentationKey: instrumentationKey,
                namePrefix: 'VISITOR_COUNTER_ACE', // Used as Postfix for cookie and localStorage 
                disableFetchTracking: false,  // To avoid tracking on all fetch
                disableAjaxTracking: true,    // Not to autocollect Ajax calls
                extensions: [AppInsightsLogListener._reactPluginInstance],
                extensionConfig: {
                    [AppInsightsLogListener._reactPluginInstance.identifier]: { history: browserHistory }
                }
            }
        });

        appInsights.loadAppInsights();
        appInsights.context.application.ver = '1.0.3'; // application_Version
        //appInsights.setAuthenticatedUserContext(_hashUser(currentUser)); // user_AuthenticateId
        return appInsights;
    }

    public static getReactPluginInstance(): ReactPlugin {
        if (!AppInsightsLogListener._reactPluginInstance) {
            AppInsightsLogListener._reactPluginInstance = new ReactPlugin();
        }
        return AppInsightsLogListener._reactPluginInstance;
    }

    public trackEvent(name: string): void {
        if (AppInsightsLogListener._appInsightsInstance)
        AppInsightsLogListener._appInsightsInstance.trackEvent(_logEventFormat(name), CONST.ApplicationInsights.CustomProps);
    }

    public log(entry: ILogEntry): void {
        const msg = _logMessageFormat(entry);
        if (entry.level === LogLevel.Off) {
            // No log required since the level is Off
            return;
        }

        if (AppInsightsLogListener._appInsightsInstance)
            switch (entry.level) {
                case LogLevel.Verbose:
                    AppInsightsLogListener._appInsightsInstance.trackTrace({ message: msg, severityLevel: SeverityLevel.Verbose }, CONST.ApplicationInsights.CustomProps);
                    break;
                case LogLevel.Info:
                    AppInsightsLogListener._appInsightsInstance.trackTrace({ message: msg, severityLevel: SeverityLevel.Information }, CONST.ApplicationInsights.CustomProps);
                    console.log({ ...CONST.ApplicationInsights.CustomProps, Message: msg });
                    break;
                case LogLevel.Warning:
                    AppInsightsLogListener._appInsightsInstance.trackTrace({ message: msg, severityLevel: SeverityLevel.Warning }, CONST.ApplicationInsights.CustomProps);
                    console.warn({ ...CONST.ApplicationInsights.CustomProps, Message: msg });
                    break;
                case LogLevel.Error:
                    AppInsightsLogListener._appInsightsInstance.trackException({ error: new Error(msg), severityLevel: SeverityLevel.Error });
                    console.error({ ...CONST.ApplicationInsights.CustomProps, Message: msg });
                    break;
            }
    }
}