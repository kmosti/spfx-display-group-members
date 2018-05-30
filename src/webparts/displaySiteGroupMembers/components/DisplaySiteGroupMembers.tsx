import * as React from 'react';
import styles from './DisplaySiteGroupMembers.module.scss';
import {
  IDisplaySiteGroupMembersProps,
  IGetSiteMembersState,
  IGroup,
  IGroupMember
} from './IDisplaySiteGroupMembersProps';
import { Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType
} from 'office-ui-fabric-react';
import { escape } from '@microsoft/sp-lodash-subset';
import pnp from "sp-pnp-js";
import { ListView, IViewField, SelectionMode, GroupOrder, IGrouping } from "@pnp/spfx-controls-react/lib/ListView";
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";

export default class DisplaySiteGroupMembers extends React.Component<IDisplaySiteGroupMembersProps, IGetSiteMembersState> {

  constructor(props: IDisplaySiteGroupMembersProps, state: IGetSiteMembersState) {
    super(props);

    this.state = {
      loading: true,
      error: "",
      showError: false,
      rows: [],
      groupTitle: "",
      notconfigured: false
    };

    this._onConfigure = this._onConfigure.bind(this);
  }
  public componentDidMount(): void {
    this._processTasks();
  }

  public componentDidUpdate( prevProps: IDisplaySiteGroupMembersProps, prevState: IGetSiteMembersState): void {
    if( prevProps.siteGroup !== this.props.siteGroup ) {
      this._resetLoadingState();
      this._processTasks();
    }
  }

  private _resetLoadingState() {
    this.setState({
        loading: true,
        rows: [],
        error: "",
        showError: false
    });
  }

  private _onConfigure() {
    // Context of the web part
    this.props.context.propertyPane.open();
  }

  private _getSelection(items: any[]) {
    // bug in 1.2.5 of ListView control needs to have selection handling, can be removed in next version per
    // https://github.com/SharePoint/sp-dev-fx-controls-react/issues/65
  }

  private _processTasks() {
    if(Number(this.props.siteGroup) ) {
      pnp.sp.web.siteGroups.getById(this.props.siteGroup).users.get().then( res => {
        let groupMembers: any[] = [{ name: "Empty", email: null }];

        if (res.length > 0) {
          groupMembers = res.map(person => ({ name: person.Title, email: person.Email }));
        }

        this.setState({
          loading: false,
          rows: groupMembers,
          notconfigured: false
        });
      }).catch( err => {
        this.setState({
          loading: false,
          error: JSON.stringify(err)
        });
      });
    } else {
      this.setState({
        loading: false,
        notconfigured: true
      });
    }

  }

  private _toggleError() {
    this.setState({
        showError: !this.state.showError
    });
  }

  public render(): React.ReactElement<IDisplaySiteGroupMembersProps> {
    let view = <Spinner size={SpinnerSize.large} label="Loading" />;
    if (!this.state.loading && this.state.rows && this.props.siteGroup) {
      const viewFields: IViewField[] = [
        {
          name: 'name',
          displayName: 'Name',
          sorting: true,
          maxWidth: 130
        },
        {
          name: 'email',
          displayName: 'Email',
          sorting: true,
          maxWidth: 120,
          render: (item: any) => {
            return <a href={"mailto:" + item['email']}>{item['email']}</a>;
          }
        }
      ];
      view = <div className={styles.container}>
        <span className={styles.title}>{this.props.groupTitle}</span>
        <ListView
          items={this.state.rows}
          viewFields={viewFields}
          compact={true}
          selection={this._getSelection}
        />
        </div>;
    }
    if (this.state.notconfigured) {
      return (
        <Placeholder
          iconName='Edit'
          iconText='Configure your web part'
          description='Select the site group you wish to display.'
          buttonLabel='Configure'
          onConfigure={this._onConfigure} />
      );
    }
    if (this.state.error !== "") {
      return (
        <MessageBar messageBarType={MessageBarType.error} className={styles.error}>
            <span>There was an error</span>
            {
                (() => {
                    if (this.state.showError) {
                        return (
                            <div>
                                <p>
                                    <a href="javascript:;" onClick={this._toggleError.bind(this)} className="ms-fontColor-neutralPrimary ms-font-m"><i className={`ms-Icon ms-Icon--ChevronUp`} aria-hidden="true"></i> Hide error message</a>
                                </p>
                                <p className="ms-font-m">{this.state.error}</p>
                            </div>
                        );
                    } else {
                        return (
                            <p>
                                <a href="javascript:;" onClick={this._toggleError.bind(this)} className="ms-fontColor-neutralPrimary ms-font-m"><i className={`ms-Icon ms-Icon--ChevronDown`} aria-hidden="true"></i> Show error message</a>
                            </p>
                        );
                    }
                })()
            }
        </MessageBar>
      );
    }
    return (
    <div className={ styles.displaySiteGroupMembers }>
        {view}
      </div>
    );
  }
}
