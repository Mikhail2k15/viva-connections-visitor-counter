import {
    BaseImageCardView, IExternalLinkCardAction, IImageCardParameters, IQuickViewCardAction
  } from '@microsoft/sp-adaptive-card-extension-base';
  import * as strings from 'VisitorCounterAdaptiveCardExtensionStrings';
  import { IVisitorCounterAdaptiveCardExtensionProps, IVisitorCounterAdaptiveCardExtensionState, QUICK_VIEW_REGISTRY_ID } from '../VisitorCounterAdaptiveCardExtension';
  
  export class ImageCardView extends BaseImageCardView<IVisitorCounterAdaptiveCardExtensionProps, IVisitorCounterAdaptiveCardExtensionState> {
    get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction {
       if (this.state.showAnalytics) {
        return {
          type: 'QuickView',
          parameters: {
            view: QUICK_VIEW_REGISTRY_ID
          }            
        };
       }        
    }
  
    get data(): IImageCardParameters {
        return {
            primaryText: this.properties.primaryText || strings.PrimaryText,
            imageUrl: this.properties.imageUrl || require('../assets/vivaConnectionsLogo.png'),
        };
    }
  }
  