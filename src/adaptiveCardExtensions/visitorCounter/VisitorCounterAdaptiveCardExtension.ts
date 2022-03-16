import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { VisitorCounterPropertyPane } from './VisitorCounterPropertyPane';
import { Logger, LogLevel } from '@pnp/logging';
//import AppInsightsListener from '../../AppInsightsListener';
import AppInsightsHelper, { TimeSpan } from '../../AppInsightsHelper';
import { AppInsightsLogListener } from '../../AppInsightsLogListener';

export interface IVisitorCounterAdaptiveCardExtensionProps {
  title: string;
  aiKey: string;
  aiAppId: string;
  aiAppKey: string;
}

export interface IVisitorCounterAdaptiveCardExtensionState {
  uniqueSessions: number;
  desktop: number;
  mobile: number;
  web: number;
}

const CARD_VIEW_REGISTRY_ID: string = 'VisitorCounter_CARD_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'VisitorCounter_QUICK_VIEW';

export default class VisitorCounterAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  IVisitorCounterAdaptiveCardExtensionProps,
  IVisitorCounterAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: VisitorCounterPropertyPane | undefined;

  public onInit(): Promise<void> {
    this.state = {
      uniqueSessions: 0,
      desktop: 0,
      mobile: 0,
      web: 0
     };

    Logger.activeLogLevel = LogLevel.Verbose;
    //Logger.subscribe(ConsoleListener());
    console.log(this.properties);
    if (this.properties !== undefined && this.properties.aiKey !== undefined){
      console.log(this.properties.aiKey);
      Logger.subscribe(new AppInsightsLogListener(this.properties.aiKey));
      
    }
    Logger.log({ message: 'VisitorCounterAdaptiveCardExtension::onInit()', level: LogLevel.Verbose});
    
    if (this.properties.aiAppId && this.properties.aiAppKey){
      const appInsightsSvc = new AppInsightsHelper(this.context.httpClient, this.properties.aiAppId, this.properties.aiAppKey);
      this.getInsights(appInsightsSvc);
    }

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());

    return Promise.resolve();
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'VisitorCounter-property-pane'*/
      './VisitorCounterPropertyPane'
    )
      .then(
        (component) => {
          this._deferredPropertyPane = new component.VisitorCounterPropertyPane();
        }
      );
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane!.getPropertyPaneConfiguration();
  }

  private getInsights = async (appInsightsSvc: AppInsightsHelper) => {
    const query: string = "customEvents | summarize dcount(session_Id)";
    const result: any[] = await appInsightsSvc.getQueryResponse(query, TimeSpan['1 day']);    

    const queryMobile: string = "customEvents | where operation_Name == '/_layouts/15/meebridge.aspx' | summarize dcount(session_Id)";
    const resultMobile: any[] = await appInsightsSvc.getQueryResponse(queryMobile, TimeSpan['30 days']);    

    const queryDesktop: string = "customEvents | where client_Browser startswith 'Electron' | summarize dcount(session_Id)";
    const resultDesktop: any[] = await appInsightsSvc.getQueryResponse(queryDesktop, TimeSpan['30 days']);    

    const queryWeb: string = "customEvents | where operation_Name == '/' | summarize dcount(session_Id)";
    const resultWeb: any[] = await appInsightsSvc.getQueryResponse(queryWeb, TimeSpan['30 days']);   

    Promise.all([result, resultDesktop, resultMobile, resultWeb]).then(()=>{
      this.setState(
        {
          uniqueSessions: result?.length == 1 ? result[0] : 0,
          desktop: resultDesktop?.length == 1 ? resultDesktop[0] : 0,
          mobile: resultMobile?.length == 1 ? resultMobile[0] : 0,
          web: resultWeb?.length == 1 ? resultWeb[0] : 0,
        });

       console.log(this.state);
    });

    
  }
}
