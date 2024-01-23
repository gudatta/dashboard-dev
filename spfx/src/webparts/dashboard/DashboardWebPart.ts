import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'DashboardWebPartStrings';

// Reference the solution
import "../../../../dist/sp-dashboard.js";
declare const SPDashboard: {
  getVersion: () => string;
  render: (props: {
    el: HTMLElement;
    context?: any;
    dashboardType?: string;
    listName?: string;
    sourceUrl?: string
  }) => void;
  updateTheme: (themeInfo: any) => void;
};

export interface IDashboardWebPartProps {
  dashboardType: string;
  webUrl: string;
}

export default class DashboardWebPart extends BaseClientSideWebPart<IDashboardWebPartProps> {

  public render(): void {
    // Render the application
    SPDashboard.render({
      el: this.domElement,
      context: this.context,
      dashboardType: this.properties.dashboardType,
      sourceUrl: this.properties.webUrl
    });
  }

  /*
  protected onInit(): Promise<void> {
  }
  */

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      // Handle theme update
      SPDashboard.updateTheme(semanticColors);
    }

  }

  protected get dataVersion(): Version {
    return Version.parse(SPDashboard.getVersion());
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('webUrl', {
                  label: strings.WebUrlFieldLabel
                }),
                PropertyPaneDropdown('dashboardType', {
                  label: strings.DashboardTypeFieldLabel,
                  selectedKey: "table",
                  options: [
                    { key: "accordion", text: "Accordion" },
                    { key: "table", text: "Table" },
                    { key: "tiles", text: "Tiles" }
                  ]
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
