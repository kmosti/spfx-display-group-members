import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IWebPartContext,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown
} from '@microsoft/sp-webpart-base';

import * as strings from 'DisplaySiteGroupMembersWebPartStrings';
import DisplaySiteGroupMembers from './components/DisplaySiteGroupMembers';
import { IDisplaySiteGroupMembersProps, IGroup } from './components/IDisplaySiteGroupMembersProps';
import pnp from "sp-pnp-js";

export interface IDisplaySiteGroupMembersWebPartProps {
  description: string;
  siteGroup: number;
  groupTitle: string;
}

export default class DisplaySiteGroupMembersWebPart extends BaseClientSideWebPart<IDisplaySiteGroupMembersWebPartProps> {

  public constructor(context?: IWebPartContext) {
    super();
    this.onPropertyPaneFieldChanged = this.onPropertyPaneFieldChanged.bind(this);
  }

  public render(): void {
    const element: React.ReactElement<IDisplaySiteGroupMembersProps > = React.createElement(
      DisplaySiteGroupMembers,
      {
        groupTitle: this.properties.groupTitle,
        siteGroup: this.properties.siteGroup,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  public onInit(): Promise<void> {

    return super.onInit().then(_ => {

      pnp.setup({
        spfxContext: this.context
      });

    });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private dropdownOptions: IPropertyPaneDropdownOption[];
  private groupsFetched: boolean;

  private fetchOptions() : Promise<IPropertyPaneDropdownOption[]> {
    //return pnp.sp.web.siteGroups.filter("substringof('" + this.context.pageContext.web.title + "',Title)").get().then( response => {
    return pnp.sp.web.siteGroups.get().then( response => {
      let options: Array<IPropertyPaneDropdownOption> = new Array<IPropertyPaneDropdownOption>();
      options = response.map( group => {
        return {
          key: group.Id,
          text: group.Title
        };
      });
      return options;
    }).catch( e => {
      console.error(e);
      return null;
    });
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    if (!this.groupsFetched) {
      this.fetchOptions().then( response => {
        this.dropdownOptions = response;
        this.groupsFetched = true;
        // now refresh the property pane, now that the promise has been resolved..
        this.context.propertyPane.refresh();
        this.render();
      });
   }
    return {
      pages: [
        {
          header: {
            description: "Get group members"
          },
          groups: [
            {
              groupName: "Configure basic settings",
              groupFields: [
                PropertyPaneTextField('groupTitle', {
                  label: "Display name"
                }),
                PropertyPaneDropdown('siteGroup', {
                  label: "Site groups",
                  options: this.dropdownOptions
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
