import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  type IPropertyPaneField,
  PropertyPaneChoiceGroup,
  PropertyPaneDropdown,
  PropertyPaneSlider,
  PropertyPaneCheckbox,
  PropertyPaneLabel,
  PropertyPaneToggle,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';

import * as strings from 'EventCalendarWebPartStrings';
import EventCalendar from './components/EventCalendar';
import { IEventCalendarProps } from './components/IEventCalendarProps';
import { IEventCalendarWebPartProps } from './models/IWebPartProps';
import { IFieldInfo } from './models/IFieldInfo';
import { initPnP } from './services/PnPSetup';
import { fetchAllListFields, getDisplayFields, autoDetectFieldMappings } from './services/FieldService';

export default class EventCalendarWebPart extends BaseClientSideWebPart<IEventCalendarWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _allFields: IFieldInfo[] = [];

  public async onInit(): Promise<void> {
    initPnP(this.context);

    // Set defaults
    if (!this.properties.viewMode) this.properties.viewMode = 'both';
    if (!this.properties.defaultView) this.properties.defaultView = 'list';
    if (!this.properties.calendarViewType) this.properties.calendarViewType = 'dayGridMonth';
    if (!this.properties.listLayout) this.properties.listLayout = 'filmstrip';
    if (!this.properties.maxEvents) this.properties.maxEvents = 50;
    if (!this.properties.selectedFields) this.properties.selectedFields = [];
    if (this.properties.showCategory === undefined) this.properties.showCategory = true;
    if (this.properties.showLocation === undefined) this.properties.showLocation = true;
    if (this.properties.showTime === undefined) this.properties.showTime = true;
    if (this.properties.showImage === undefined) this.properties.showImage = true;

    // Load fields if list already selected
    if (this.properties.selectedListId) {
      await this._loadFields(this.properties.selectedListId);
    }
  }

  public render(): void {
    const mappedCoreFields = [
      this.properties.titleField,
      this.properties.startDateField,
      this.properties.endDateField,
      this.properties.allDayField,
      this.properties.categoryField,
      this.properties.locationField,
    ].filter(Boolean);
    const displayFields = getDisplayFields(this._allFields, mappedCoreFields);

    const element: React.ReactElement<IEventCalendarProps> = React.createElement(
      EventCalendar,
      {
        listId: this.properties.selectedListId,
        fieldMapping: {
          titleField: this.properties.titleField || 'Title',
          startDateField: this.properties.startDateField || 'EventDate',
          endDateField: this.properties.endDateField || 'EndDate',
          allDayField: this.properties.allDayField || '',
          categoryField: this.properties.categoryField || '',
          locationField: this.properties.locationField || '',
        },
        selectedFields: this.properties.selectedFields || [],
        availableFields: displayFields,
        viewMode: this.properties.viewMode || 'both',
        defaultView: this.properties.defaultView || 'list',
        calendarViewType: this.properties.calendarViewType || 'dayGridMonth',
        listLayout: this.properties.listLayout || 'filmstrip',
        maxEvents: this.properties.maxEvents || 50,
        cardDisplay: {
          showCategory: this.properties.showCategory !== false,
          showLocation: this.properties.showLocation !== false,
          showTime: this.properties.showTime !== false,
          showImage: this.properties.showImage !== false,
        },
        isDarkTheme: this._isDarkTheme,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) return;
    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;
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

  protected async onPropertyPaneFieldChanged(propertyPath: string, _oldValue: unknown, _newValue: unknown): Promise<void> {
    if (propertyPath === 'selectedListId' && _newValue) {
      // Reset everything when list changes
      this.properties.selectedFields = [];
      this.properties.titleField = '';
      this.properties.startDateField = '';
      this.properties.endDateField = '';
      this.properties.allDayField = '';
      this.properties.categoryField = '';
      this.properties.locationField = '';

      await this._loadFields(_newValue as string);

      // Auto-detect field mappings
      const detected = autoDetectFieldMappings(this._allFields);
      this.properties.titleField = detected.titleField;
      this.properties.startDateField = detected.startDateField;
      this.properties.endDateField = detected.endDateField;
      this.properties.allDayField = detected.allDayField;
      this.properties.categoryField = detected.categoryField;
      this.properties.locationField = detected.locationField;

      this.context.propertyPane.refresh();
      this.render();
    }

    if (propertyPath === 'viewMode') {
      this.context.propertyPane.refresh();
      this.render();
    }

    if (['titleField', 'startDateField', 'endDateField', 'allDayField',
         'categoryField', 'locationField'].indexOf(propertyPath) >= 0) {
      this.context.propertyPane.refresh();
      this.render();
    }

    if (['showCategory', 'showLocation', 'showTime', 'showImage',
         'listLayout'].indexOf(propertyPath) >= 0) {
      this.render();
    }

    if (propertyPath.startsWith('field_')) {
      const fieldName = propertyPath.substring(6);
      const selected = this.properties.selectedFields || [];
      if (_newValue) {
        if (selected.indexOf(fieldName) < 0) {
          selected.push(fieldName);
        }
      } else {
        const idx = selected.indexOf(fieldName);
        if (idx >= 0) selected.splice(idx, 1);
      }
      this.properties.selectedFields = [...selected];
      this.render();
    }
  }

  private async _loadFields(listId: string): Promise<void> {
    try {
      this._allFields = await fetchAllListFields(listId);
    } catch {
      this._allFields = [];
    }
  }

  private _getFieldDropdownOptions(filterTypes?: string[]): Array<{ key: string; text: string }> {
    const options: Array<{ key: string; text: string }> = [{ key: '', text: '(none)' }];
    for (const f of this._allFields) {
      if (!filterTypes || filterTypes.indexOf(f.fieldType) >= 0) {
        options.push({ key: f.internalName, text: f.displayName + ' (' + f.fieldType + ')' });
      }
    }
    return options;
  }

  private _getTextFieldOptions(): Array<{ key: string; text: string }> {
    const textTypes = ['Text', 'Note', 'Choice', 'Computed'];
    const options: Array<{ key: string; text: string }> = [];
    for (const f of this._allFields) {
      if (textTypes.indexOf(f.fieldType) >= 0) {
        options.push({ key: f.internalName, text: f.displayName });
      }
    }
    return options;
  }

  private _getCategoryFieldOptions(): Array<{ key: string; text: string }> {
    const types = ['Text', 'Choice', 'Computed'];
    const options: Array<{ key: string; text: string }> = [{ key: '', text: '(none)' }];
    for (const f of this._allFields) {
      if (types.indexOf(f.fieldType) >= 0) {
        options.push({ key: f.internalName, text: f.displayName });
      }
    }
    return options;
  }

  private _getLocationFieldOptions(): Array<{ key: string; text: string }> {
    const types = ['Text', 'Note', 'Choice', 'Computed'];
    const options: Array<{ key: string; text: string }> = [{ key: '', text: '(none)' }];
    for (const f of this._allFields) {
      if (types.indexOf(f.fieldType) >= 0) {
        options.push({ key: f.internalName, text: f.displayName });
      }
    }
    return options;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    // Get display fields (excluding mapped core fields)
    const mappedCoreFields = [
      this.properties.titleField,
      this.properties.startDateField,
      this.properties.endDateField,
      this.properties.allDayField,
      this.properties.categoryField,
      this.properties.locationField,
    ].filter(Boolean);
    const displayFields = getDisplayFields(this._allFields, mappedCoreFields);

    const fieldCheckboxes = displayFields.map(field =>
      PropertyPaneCheckbox(`field_${field.internalName}`, {
        text: field.displayName,
        checked: (this.properties.selectedFields || []).indexOf(field.internalName) >= 0,
      })
    );

    const viewMode = this.properties.viewMode || 'both';
    const showCalendarOptions = viewMode === 'both' || viewMode === 'calendar';
    const showListOptions = viewMode === 'both' || viewMode === 'list';

    // Column mapping fields
    const mappingFields: IPropertyPaneField<unknown>[] = [];

    if (this._allFields.length > 0) {
      mappingFields.push(
        PropertyPaneDropdown('titleField', {
          label: strings.TitleFieldLabel,
          options: this._getTextFieldOptions(),
          selectedKey: this.properties.titleField || '',
        }) as IPropertyPaneField<unknown>,
        PropertyPaneDropdown('startDateField', {
          label: strings.StartDateFieldLabel,
          options: this._getFieldDropdownOptions(['DateTime', 'Date']),
          selectedKey: this.properties.startDateField || '',
        }) as IPropertyPaneField<unknown>,
        PropertyPaneDropdown('endDateField', {
          label: strings.EndDateFieldLabel,
          options: this._getFieldDropdownOptions(['DateTime', 'Date']),
          selectedKey: this.properties.endDateField || '',
        }) as IPropertyPaneField<unknown>,
        PropertyPaneDropdown('allDayField', {
          label: strings.AllDayFieldLabel,
          options: this._getFieldDropdownOptions(['Boolean', 'AllDayEvent']),
          selectedKey: this.properties.allDayField || '',
        }) as IPropertyPaneField<unknown>,
        PropertyPaneDropdown('categoryField', {
          label: strings.CategoryFieldLabel,
          options: this._getCategoryFieldOptions(),
          selectedKey: this.properties.categoryField || '',
        }) as IPropertyPaneField<unknown>,
        PropertyPaneDropdown('locationField', {
          label: strings.LocationFieldLabel,
          options: this._getLocationFieldOptions(),
          selectedKey: this.properties.locationField || '',
        }) as IPropertyPaneField<unknown>,
      );
    } else {
      mappingFields.push(
        PropertyPaneLabel('noMappingFields', {
          text: this.properties.selectedListId ? 'Loading fields...' : 'Select a list first',
        }) as IPropertyPaneField<unknown>,
      );
    }

    // View settings
    const viewFields: IPropertyPaneField<unknown>[] = [
      PropertyPaneChoiceGroup('viewMode', {
        label: strings.ViewModeLabel,
        options: [
          { key: 'both', text: 'Both (Calendar & List)', iconProps: { officeFabricIconFontName: 'Switch' } },
          { key: 'calendar', text: 'Calendar Only', iconProps: { officeFabricIconFontName: 'Calendar' } },
          { key: 'list', text: 'List Only', iconProps: { officeFabricIconFontName: 'List' } },
        ],
      }) as IPropertyPaneField<unknown>,
    ];

    if (viewMode === 'both') {
      viewFields.push(
        PropertyPaneChoiceGroup('defaultView', {
          label: strings.DefaultViewLabel,
          options: [
            { key: 'calendar', text: 'Calendar' },
            { key: 'list', text: 'List' },
          ],
        }) as IPropertyPaneField<unknown>,
      );
    }

    if (showCalendarOptions) {
      viewFields.push(
        PropertyPaneChoiceGroup('calendarViewType', {
          label: strings.CalendarViewTypeLabel,
          options: [
            { key: 'dayGridMonth', text: 'Month' },
            { key: 'timeGridWeek', text: 'Week' },
            { key: 'timeGridDay', text: 'Day' },
          ],
        }) as IPropertyPaneField<unknown>,
      );
    }

    if (showListOptions) {
      viewFields.push(
        PropertyPaneChoiceGroup('listLayout', {
          label: strings.ListLayoutLabel,
          options: [
            { key: 'filmstrip', text: 'Filmstrip', iconProps: { officeFabricIconFontName: 'Slideshow' } },
            { key: 'compact', text: 'Compact list', iconProps: { officeFabricIconFontName: 'List' } },
          ],
        }) as IPropertyPaneField<unknown>,
      );
    }

    viewFields.push(
      PropertyPaneSlider('maxEvents', {
        label: strings.MaxEventsLabel,
        min: 10,
        max: 200,
        step: 10,
        value: this.properties.maxEvents || 50,
      }) as IPropertyPaneField<unknown>,
    );

    // Card display toggles
    const cardDisplayFields: IPropertyPaneField<unknown>[] = [
      PropertyPaneToggle('showImage', {
        label: strings.ShowImageLabel,
        checked: this.properties.showImage !== false,
        onText: 'On',
        offText: 'Off',
      }) as IPropertyPaneField<unknown>,
      PropertyPaneToggle('showCategory', {
        label: strings.ShowCategoryLabel,
        checked: this.properties.showCategory !== false,
        onText: 'On',
        offText: 'Off',
      }) as IPropertyPaneField<unknown>,
      PropertyPaneToggle('showTime', {
        label: strings.ShowTimeLabel,
        checked: this.properties.showTime !== false,
        onText: 'On',
        offText: 'Off',
      }) as IPropertyPaneField<unknown>,
      PropertyPaneToggle('showLocation', {
        label: strings.ShowLocationLabel,
        checked: this.properties.showLocation !== false,
        onText: 'On',
        offText: 'Off',
      }) as IPropertyPaneField<unknown>,
    ];

    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.DataSourceGroupName,
              groupFields: [
                PropertyFieldListPicker('selectedListId', {
                  label: strings.ListPickerLabel,
                  selectedList: this.properties.selectedListId,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context as never,
                  key: 'listPickerFieldId',
                }),
              ],
            },
            {
              groupName: strings.ColumnMappingGroupName,
              groupFields: mappingFields,
            },
            {
              groupName: strings.FieldsGroupName,
              groupFields: displayFields.length > 0
                ? fieldCheckboxes
                : [PropertyPaneLabel('noFields', { text: this.properties.selectedListId ? 'Configure column mapping above first' : 'Select a list first' })],
            },
          ],
        },
        {
          header: { description: strings.DisplaySettingsDescription },
          groups: [
            {
              groupName: strings.ViewSettingsGroupName,
              groupFields: viewFields,
            },
            {
              groupName: strings.CardDisplayGroupName,
              groupFields: cardDisplayFields,
            },
          ],
        },
      ],
    };
  }
}
