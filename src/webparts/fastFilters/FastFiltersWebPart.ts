import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'FastFiltersWebPartStrings';
import FastFilters from './components/FastFilters';
import { IFastFiltersProps, ISourceProps } from './components/IFastFiltersProps';

export interface IFastFiltersWebPartProps {
  description: string;

  webUrl: string;
  listTitle: string;
  webRelativeLink: string;
  viewItemLink?: string;
  columns: string;
  searchProps: string;
  selectThese?: string;
  restFilter?: string;
  searchSourceDesc: string;
  itemFetchCol?: string; //higher cost columns to fetch on opening panel
  orderByProp: string;
  orderByAsc: boolean;
  defSearchButtons: string;  //These are default buttons always on that source page.  Use case for Manual:  Policy, Instruction etc...


}

export default class FastFiltersWebPart extends BaseClientSideWebPart<IFastFiltersWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {


    let sourceProps: ISourceProps = {
      webUrl: this.properties.webUrl,
      listTitle: this.properties.webUrl,
      webRelativeLink: this.properties.webUrl,
      viewItemLink: this.properties.webUrl,
      columns: this.properties.columns ? this.properties.columns.split(';') : [],
      searchProps: this.properties.searchProps ? this.properties.searchProps.split(';') : [],
      selectThese: this.properties.selectThese ? this.properties.selectThese.split(';') : [],
      restFilter: this.properties.webUrl,
      searchSourceDesc: this.properties.webUrl,
      itemFetchCol: this.properties.itemFetchCol ? this.properties.itemFetchCol.split(';') : [], //higher cost columns to fetch on opening panel
      orderBy: {
          prop: this.properties.orderByProp,
          asc: this.properties.orderByAsc,
      },
      defSearchButtons: this.properties.defSearchButtons ? this.properties.defSearchButtons.split(';') : [],  //These are default buttons always on that source page.  Use case for Manual:  Policy, Instruction etc...

    }


    const element: React.ReactElement<IFastFiltersProps> = React.createElement(
      FastFilters,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,

        sourceProps: sourceProps,

      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
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
