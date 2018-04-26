import { IPropertyPaneDropdownProps, IWebPartContext } from '@microsoft/sp-webpart-base';

export interface IDisplaySiteGroupMembersProps {
  description?: string;
  siteGroup?: number;
  groupTitle?: string;
  context: IWebPartContext;
}

export interface IGetSiteMembersState {
  loading?: boolean;
  error?: string;
  groupTitle?: string;
  columns?: string[];
  showError?: boolean;
  rows?: IGroupMember[];
  notconfigured?: boolean;
}

export interface IGroup {
  Id?: number;
  Title?: string;
}

export interface IGroupMember {
  Id?: number;
  Title?: string;
  email?: string;
}
