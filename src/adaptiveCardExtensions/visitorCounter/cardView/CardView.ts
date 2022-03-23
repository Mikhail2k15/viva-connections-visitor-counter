import {
  BasePrimaryTextCardView,
  IPrimaryTextCardParameters,
  IExternalLinkCardAction,
  IQuickViewCardAction,
  ICardButton
} from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'VisitorCounterAdaptiveCardExtensionStrings';
import { IVisitorCounterAdaptiveCardExtensionProps, IVisitorCounterAdaptiveCardExtensionState, QUICK_VIEW_REGISTRY_ID } from '../VisitorCounterAdaptiveCardExtension';

export class CardView extends BasePrimaryTextCardView<IVisitorCounterAdaptiveCardExtensionProps, IVisitorCounterAdaptiveCardExtensionState> {
  public get cardButtons(): [ICardButton] | [ICardButton, ICardButton] | undefined {
    return [
      {
        title: strings.QuickViewButton,
        action: {
          type: 'QuickView',
          parameters: {
            view: QUICK_VIEW_REGISTRY_ID
          }
        }
      }
    ];
  }

  public get data(): IPrimaryTextCardParameters {
    const today: number = this.state.today;
    const msteams: number = +this.state.desktop + +this.state.mobile + +this.state.web;
    return {
      primaryText: `${today} visits today`,
      description: `Click Details for more information`,
      title: this.properties.title
    };
  }

  
}
