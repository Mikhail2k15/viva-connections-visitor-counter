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
        const today: number = this.state.today;
        return {
            primaryText: `${today} visits today `,
            imageUrl: 'https://statics.teams.cdn.office.net/evergreen-assets/apps/d2c6f111-ffad-42a0-b65e-ee00425598aa_largeImage.png'
        };
    }
  }
  