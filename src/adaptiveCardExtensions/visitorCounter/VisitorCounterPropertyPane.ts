import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import * as strings from 'VisitorCounterAdaptiveCardExtensionStrings';

export class VisitorCounterPropertyPane {
  public getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.GeneralFieldsGroupName,
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                }),
                PropertyPaneTextField('primaryText', {
                  label: strings.PrimaryTextFieldLabel
                }),
                PropertyPaneTextField('imageUrl', {
                  label: strings.CustomImageFieldLabel
                }),
                PropertyPaneTextField('analytics', {
                  label: strings.AnalystsFieldLabel
                }),
                PropertyPaneToggle('trackPII', {
                  label: strings.TrackPIIFieldLabel
                }),                
              ]
            },
            {
              groupName: strings.AppInsightsFieldsGroupName,
              groupFields: [
                PropertyPaneTextField('aiKey', {
                  label: strings.AppInsightsInstrumentationKeyFieldLabel
                }),
                PropertyPaneTextField('aiAppId', {
                  label: strings.AppInsightsApplicationIDFieldLabel
                }),
                PropertyPaneTextField('aiAppKey', {
                  label: strings.AppInsightsAPIKeyFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
