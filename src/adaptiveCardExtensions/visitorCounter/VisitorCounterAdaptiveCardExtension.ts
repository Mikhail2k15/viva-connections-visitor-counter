import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { VisitorCounterPropertyPane } from './VisitorCounterPropertyPane';
import { Logger, LogLevel } from '@pnp/logging';
//import AppInsightsListener from '../../AppInsightsListener';
import AppInsightsHelper, { TimeSpan } from '../../AppInsightsHelper';
import { AppInsightsLogListener } from '../../AppInsightsLogListener';
import { Log } from '@microsoft/sp-core-library';

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
    try {
      Logger.activeLogLevel = LogLevel.Verbose;
      //Logger.subscribe(ConsoleListener());
      console.log('this.properties ', this.properties);
      
      
      this.state = {
        uniqueSessions: 0,
        desktop: 0,
        mobile: 0,
        web: 0
       };
  
      Log.info('ACE', 'onInit standard log output');
  
         
      
      if (this.properties.aiAppId && this.properties.aiAppKey){
        const appInsightsSvc = new AppInsightsHelper(this.context.httpClient, this.properties.aiAppId, this.properties.aiAppKey);
        this.getInsights(appInsightsSvc);
      }
  
      this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
      this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());
  

      if (this.properties.aiKey){
        console.log('this.properties.aiKey',this.properties.aiKey);
        let ai = new AppInsightsLogListener(this.properties.aiKey); 
        ai.trackEvent("onInit");     
      } 

      return Promise.resolve();
    }
    catch (error){
      console.log(error.message);
      Logger.write(`Error in onInit: ${error.message}`, LogLevel.Error);
    }    
  }

  protected loadPropertyPaneResources(): Promise<void> {
    console.log('begin loadPropertyPaneResources()');
    return import(
      /* webpackChunkName: 'VisitorCounter-property-pane'*/
      './VisitorCounterPropertyPane'
    )
      .then(
        (component) => {
          this._deferredPropertyPane = new component.VisitorCounterPropertyPane();
          console.log('end loadPropertyPaneResources()');
        }
      );
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    console.log('begin getPropertyPaneConfiguration()');
    return this._deferredPropertyPane!.getPropertyPaneConfiguration();
  }

  protected renderCard(): string | undefined {
    console.log('begin render card',this.properties.aiKey);
    
    Logger.log({ message: 'renderCard()', level: LogLevel.Info});
    return CARD_VIEW_REGISTRY_ID;
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
