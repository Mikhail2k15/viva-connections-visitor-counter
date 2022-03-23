import { ISPFxAdaptiveCard, BaseAdaptiveCardView } from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'VisitorCounterAdaptiveCardExtensionStrings';
import { IVisitorCounterAdaptiveCardExtensionProps, IVisitorCounterAdaptiveCardExtensionState } from '../VisitorCounterAdaptiveCardExtension';

export interface IQuickViewData {
  subTitle: string;
  title: string;
  msteams: string;
  spo: string;
  desktop: string;
  mobile: string;
  web: string;
}

export class QuickView extends BaseAdaptiveCardView<
  IVisitorCounterAdaptiveCardExtensionProps,
  IVisitorCounterAdaptiveCardExtensionState,
  IQuickViewData
> {
  public get data(): IQuickViewData {
    const monthly: number = this.state.monthly;
    const msteams: number = +this.state.desktop + +this.state.mobile + +this.state.web;
    const spo: number = monthly - msteams;
    return {
      subTitle: strings.SubTitle,
      title: strings.Title,
      msteams: msteams.toString(),
      spo: spo.toString(),
      desktop: this.state.desktop.toString(),
      mobile: this.state.mobile.toString(),
      web: this.state.web.toString() 
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return require('./template/QuickViewTemplate.json');
  }
}