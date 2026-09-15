import * as React from 'react';
import { Dialog, DialogType } from 'office-ui-fabric-react/lib/Dialog';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { DefaultButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { IBreadcrumbConfiguration } from '../models/IBreadcrumbConfiguration';
import { INavigationListReference } from '../models/IBreadcrumbItem';
import { ConfigurationService, DEFAULT_NAVIGATION_LIST_TITLE } from '../services/ConfigurationService';
import { IDynamicBreadcrumbStrings } from '../components/IDynamicBreadcrumbProps';
import styles from '../styles/DynamicBreadcrumb.module.scss';

export interface IBreadcrumbConfigurationDialogProps {
  hidden: boolean;
  isRtl: boolean;
  strings: IDynamicBreadcrumbStrings;
  configuration: IBreadcrumbConfiguration | undefined;
  configurationService: ConfigurationService;
  onDismiss: () => void;
  onSaved: () => void;
}

export interface IBreadcrumbConfigurationDialogState {
  lists: INavigationListReference[];
  selectedIdentifier: string;
  newListTitle: string;
  busy: boolean;
  error?: string;
  confirmation?: string;
}

/** Admin-only site configuration UI hosted by the Application Customizer itself. */
export class BreadcrumbConfigurationDialog extends React.Component<
  IBreadcrumbConfigurationDialogProps,
  IBreadcrumbConfigurationDialogState> {

  public constructor(props: IBreadcrumbConfigurationDialogProps) {
    super(props);
    this.state = this._initialState(props);
  }

  public componentDidMount(): void {
    if (!this.props.hidden) {
      this._loadLists();
    }
  }

  public componentWillReceiveProps(nextProps: IBreadcrumbConfigurationDialogProps): void {
    if (this.props.hidden && !nextProps.hidden) {
      this.setState(this._initialState(nextProps), () => this._loadLists());
    }
  }

  public render(): JSX.Element {
    const options: IDropdownOption[] = this.state.lists.map((list: INavigationListReference) => ({
      key: list.id,
      text: list.title
    }));
    const strings: IDynamicBreadcrumbStrings = this.props.strings;

    return (
      <Dialog
        hidden={ this.props.hidden }
        onDismiss={ this.props.onDismiss }
        modalProps={ { isBlocking: this.state.busy } }
        dialogContentProps={ {
          type: DialogType.normal,
          title: strings.ConfigurationDialogTitle,
          subText: strings.ConfigurationDialogDescription
        } }>
        <div className={ styles.configurationDialog } dir={ this.props.isRtl ? 'rtl' : 'ltr' }>
          <Dropdown
            label={ strings.NavigationListLabel }
            options={ options }
            selectedKey={ this._selectedListKey() }
            placeHolder={ strings.NavigationListPlaceholder }
            disabled={ this.state.busy }
            onChanged={ this._onListChanged } />
          <TextField
            label={ strings.NavigationListHelp }
            value={ this.state.selectedIdentifier }
            disabled={ this.state.busy }
            onChanged={ this._onIdentifierChanged } />
          <p className={ styles.fieldHelp }>{ strings.RequiredFieldsHelp }</p>
          <TextField
            label={ strings.NavigationListNameLabel }
            value={ this.state.newListTitle }
            disabled={ this.state.busy }
            onChanged={ this._onNewListTitleChanged } />
          <div className={ styles.dialogActions }>
            <PrimaryButton
              text={ this.state.busy ? strings.Creating : strings.CreateNavigationList }
              disabled={ this.state.busy }
              onClick={ this._onCreateList } />
          </div>
          { this.state.error ?
            <div className={ styles.dialogError } role='alert'>{ this.state.error }</div> : undefined }
          { this.state.confirmation ?
            <div className={ styles.dialogConfirmation } role='status'>{ this.state.confirmation }</div> : undefined }
          <div className={ styles.dialogActions }>
            <PrimaryButton
              text={ this.state.busy ? strings.Saving : strings.Save }
              disabled={ this.state.busy }
              onClick={ this._onSave } />
            <DefaultButton text={ strings.Cancel } disabled={ this.state.busy } onClick={ this.props.onDismiss } />
          </div>
        </div>
      </Dialog>
    );
  }

  private _initialState(props: IBreadcrumbConfigurationDialogProps): IBreadcrumbConfigurationDialogState {
    return {
      lists: [],
      selectedIdentifier: props.configuration ? props.configuration.navigationListId : '',
      newListTitle: DEFAULT_NAVIGATION_LIST_TITLE,
      busy: false
    };
  }

  private _selectedListKey(): string | undefined {
    const identifier: string = this.state.selectedIdentifier;
    for (let index: number = 0; index < this.state.lists.length; index++) {
      if (this.state.lists[index].id === identifier) {
        return identifier;
      }
    }
    return undefined;
  }

  private _loadLists(): void {
    this.props.configurationService.getNavigationLists()
      .then((lists: INavigationListReference[]) => this.setState({ lists: lists }))
      .catch((error: Error) => this.setState({ error: error.message }));
  }

  private _onListChanged = (option: IDropdownOption): void => {
    this.setState({ selectedIdentifier: String(option.key), error: undefined, confirmation: undefined });
  }

  private _onIdentifierChanged = (value: string): void => {
    this.setState({ selectedIdentifier: value, error: undefined, confirmation: undefined });
  }

  private _onNewListTitleChanged = (value: string): void => {
    this.setState({ newListTitle: value, error: undefined, confirmation: undefined });
  }

  private _onCreateList = (): void => {
    this.setState({ busy: true, error: undefined, confirmation: undefined });
    this.props.configurationService.provisionNavigationList(this.state.newListTitle)
      .then((list: INavigationListReference) => {
        this.setState({
          selectedIdentifier: list.id,
          busy: false,
          confirmation: '"' + list.title + '" is ready. Add navigation items, then save this configuration.'
        }, () => this._loadLists());
      })
      .catch((error: Error) => this.setState({ busy: false, error: error.message }));
  }

  private _onSave = (): void => {
    if (!this.state.selectedIdentifier.replace(/^\s+|\s+$/g, '')) {
      this.setState({ error: this.props.strings.SelectNavigationList });
      return;
    }

    this.setState({ busy: true, error: undefined, confirmation: undefined });
    this.props.configurationService.resolveNavigationList(this.state.selectedIdentifier)
      .then((list: INavigationListReference) => this.props.configurationService.saveConfiguration(list))
      .then(() => {
        this.setState({ busy: false, confirmation: this.props.strings.ConfigurationSaved });
        this.props.onSaved();
      })
      .catch((error: Error) => this.setState({ busy: false, error: error.message }));
  }
}
