import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
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
  description: string;
}

export default class DashboardWebPart extends BaseClientSideWebPart<IDashboardWebPartProps> {

  public render(): void {
    // Render the application
    SPDashboard.render({
      el: this.domElement,
      context: this.context
    });
  }

  /*
  protected onInit(): Promise<void> {
  }
  */

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
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
