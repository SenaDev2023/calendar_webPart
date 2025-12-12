import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import Calendar from './components/Calendar/Calendar';
import { ICalendarProps } from './components/Calendar/ICalendarProps';
import spservices from '../../services/spservices';
import { PropertyFieldDateTimePicker, DateConvention, IDateTimeFieldValue } from '@pnp/spfx-property-controls/lib/PropertyFieldDateTimePicker';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';

export interface ICalendarWebPartProps {
  siteUrl: string;
  list: string;
  eventStartDate: IDateTimeFieldValue;
  eventEndDate: IDateTimeFieldValue;
  showPersonaBadge: boolean;
  personaBadgeSize: string;
  showEventOwner: boolean;
  showInMobileView: boolean;
  enableDebugMode: boolean;
}

export default class CalendarWebPart extends BaseClientSideWebPart<ICalendarWebPartProps> {
  private _isDarkTheme: boolean = false;
  private spService: spservices = null;

  public render(): void {
    const element: React.ReactElement<ICalendarProps> = React.createElement(
      Calendar,
      {
        siteUrl: this.properties.siteUrl || this.context.pageContext.web.absoluteUrl,
        list: this.properties.list,
        eventStartDate: this.properties.eventStartDate || this._getDefaultStartDate(),
        eventEndDate: this.properties.eventEndDate || this._getDefaultEndDate(),
        showPersonaBadge: this.properties.showPersonaBadge !== undefined ? this.properties.showPersonaBadge : true,
        personaBadgeSize: this.properties.personaBadgeSize || 'small',
        showEventOwner: this.properties.showEventOwner !== undefined ? this.properties.showEventOwner : true,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: '',
        hasTeamsContext: Boolean(this.context.sdks.microsoftTeams),
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        displayMode: this.displayMode,
        updateProperty: (value: string) => {
          this.properties.siteUrl = value;
        }
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private _getDefaultStartDate(): IDateTimeFieldValue {
    const date = new Date();
    return {
      value: date,
      displayValue: date.toLocaleDateString()
    };
  }

  private _getDefaultEndDate(): IDateTimeFieldValue {
    const date = new Date();
    date.setFullYear(date.getFullYear() + 2);
    return {
      value: date,
      displayValue: date.toLocaleDateString()
    };
  }

  protected async onInit(): Promise<void> {
    this.spService = new spservices(this.context);
    
    if (!this.properties.siteUrl) {
      this.properties.siteUrl = this.context.pageContext.web.absoluteUrl;
    }
    
    if (!this.properties.eventStartDate) {
      this.properties.eventStartDate = this._getDefaultStartDate();
    }
    
    if (!this.properties.eventEndDate) {
      this.properties.eventEndDate = this._getDefaultEndDate();
    }
    
    if (this.properties.showPersonaBadge === undefined) {
      this.properties.showPersonaBadge = true;
    }
    
    if (!this.properties.personaBadgeSize) {
      this.properties.personaBadgeSize = 'small';
    }
    
    if (this.properties.showEventOwner === undefined) {
      this.properties.showEventOwner = true;
    }

    return Promise.resolve();
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
    const personaBadgeSizeOptions: IPropertyPaneDropdownOption[] = [
      { key: 'small', text: 'Small (20px)' },
      { key: 'medium', text: 'Medium (32px)' },
      { key: 'large', text: 'Large (48px)' }
    ];

    return {
      pages: [
        {
          header: {
            description: 'Configure your Events Calendar with modern settings. All changes are applied immediately.'
          },
          displayGroupsAsAccordion: true,
          groups: [
            {
              groupName: '📊 Data Source',
              isCollapsed: false,
              groupFields: [
                PropertyPaneTextField('siteUrl', {
                  label: 'Site URL',
                  description: 'The SharePoint site containing your calendar list. Leave blank to use the current site.',
                  placeholder: 'https://contoso.sharepoint.com/sites/team',
                  multiline: false
                }),
                PropertyFieldListPicker('list', {
                  label: 'Calendar List',
                  selectedList: this.properties.list,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context as never,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'listPickerFieldId',
                  webAbsoluteUrl: this.properties.siteUrl || this.context.pageContext.web.absoluteUrl,
                  baseTemplate: 106
                })
              ]
            },
            {
              groupName: '📅 Date Range',
              isCollapsed: false,
              groupFields: [
                PropertyFieldDateTimePicker('eventStartDate', {
                  label: 'Start Date',
                  dateConvention: DateConvention.Date,
                  initialDate: this.properties.eventStartDate || this._getDefaultStartDate(),
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'eventStartDateFieldId',
                  showLabels: false
                }),
                PropertyFieldDateTimePicker('eventEndDate', {
                  label: 'End Date',
                  dateConvention: DateConvention.Date,
                  initialDate: this.properties.eventEndDate || this._getDefaultEndDate(),
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'eventEndDateFieldId',
                  showLabels: false
                })
              ]
            },
            {
              groupName: '🎨 Display Options',
              isCollapsed: false,
              groupFields: [
                PropertyPaneToggle('showPersonaBadge', {
                  label: 'Show Persona Badges',
                  onText: 'Visible',
                  offText: 'Hidden',
                  checked: this.properties.showPersonaBadge !== undefined ? this.properties.showPersonaBadge : true
                }),
                PropertyPaneDropdown('personaBadgeSize', {
                  label: 'Badge Size',
                  options: personaBadgeSizeOptions,
                  selectedKey: this.properties.personaBadgeSize || 'small',
                  disabled: !this.properties.showPersonaBadge
                }),
                PropertyPaneToggle('showEventOwner', {
                  label: 'Show Event Owner',
                  onText: 'Visible',
                  offText: 'Hidden',
                  checked: this.properties.showEventOwner !== undefined ? this.properties.showEventOwner : true
                })
              ]
            },
            {
              groupName: '⚙️ Advanced',
              isCollapsed: true,
              groupFields: [
                PropertyPaneToggle('showInMobileView', {
                  label: 'Mobile and Email Visibility',
                  onText: 'Visible',
                  offText: 'Hidden',
                  checked: this.properties.showInMobileView !== undefined ? this.properties.showInMobileView : true
                }),
                PropertyPaneToggle('enableDebugMode', {
                  label: 'Debug Mode',
                  onText: 'Enabled',
                  offText: 'Disabled',
                  checked: this.properties.enableDebugMode || false
                })
              ]
            }
          ]
        }
      ]
    };
  }
}