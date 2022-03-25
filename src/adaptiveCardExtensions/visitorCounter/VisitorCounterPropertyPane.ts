import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import * as strings from 'VisitorCounterAdaptiveCardExtensionStrings';

export class VisitorCounterPropertyPane {
  public getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: 'General',
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                }),
                PropertyPaneTextField('primaryText', {
                  label: 'Primary Text'
                }),
                PropertyPaneTextField('imageUrl', {
                  label: 'Image Url'
                }),
                PropertyPaneTextField('analytics', {
                  label: 'Analytics'
                })                
              ]
            },
            {
              groupName: 'Application Insights',
              groupFields: [
                PropertyPaneTextField('aiKey', {
                  label: 'Instrumentation Key'
                }),
                PropertyPaneTextField('aiAppId', {
                  label: 'Application ID'
                }),
                PropertyPaneTextField('aiAppKey', {
                  label: 'API key'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
