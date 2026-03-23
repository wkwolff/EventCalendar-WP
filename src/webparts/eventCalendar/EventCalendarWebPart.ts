/**
 * @file EventCalendarWebPart.ts
 * @description SPFx web part entry point for the Event Calendar. Handles property pane
 *   configuration, list selection, field auto-detection, theme integration, and React
 *   component bootstrapping. The property pane is split into two pages: data source
 *   configuration (list picker, column mapping, extra field selection) and display
 *   settings (view mode, calendar type, list layout, card toggles).
 * @author W. Kevin Wolff
 * @copyright TidalHealth
 */

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

/**
 * SharePoint Framework web part that renders an event calendar sourced from
 * a SharePoint list. Supports calendar view (FullCalendar), filmstrip cards,
 * and compact list layouts with configurable field mappings.
 */
export default class EventCalendarWebPart extends BaseClientSideWebPart<IEventCalendarWebPartProps> {
  /** Tracks the current SharePoint theme inversion state (dark vs. light). */
  private _isDarkTheme: boolean = false;

  /** Cached field metadata for the currently selected SharePoint list. */
  private _allFields: IFieldInfo[] = [];

  /**
   * Initializes PnPjs context and sets sensible defaults for any web part
   * properties that have not yet been persisted. If a list is already
   * selected, pre-loads its field metadata so the property pane can render
   * column mapping dropdowns immediately.
   * @returns A promise that resolves once initialization is complete.
   */
  public async onInit(): Promise<void> {
    initPnP(this.context);

    // Apply default values for first-time configuration
    if (!this.properties.viewMode) this.properties.viewMode = 'both';
    if (!this.properties.defaultView) this.properties.defaultView = 'list';
    if (!this.properties.calendarViewType) this.properties.calendarViewType = 'dayGridMonth';
    if (!this.properties.listLayout) this.properties.listLayout = 'filmstrip';
    if (!this.properties.maxEvents) this.properties.maxEvents = 50;
    if (!this.properties.selectedFields) this.properties.selectedFields = [];
    // Boolean defaults must check for `undefined` explicitly (falsy check would override `false`)
    if (this.properties.showCategory === undefined) this.properties.showCategory = true;
    if (this.properties.showLocation === undefined) this.properties.showLocation = true;
    if (this.properties.showTime === undefined) this.properties.showTime = true;
    if (this.properties.showImage === undefined) this.properties.showImage = true;

    // Pre-load field metadata when the web part already has a list configured
    if (this.properties.selectedListId) {
      await this._loadFields(this.properties.selectedListId);
    }
  }

  /**
   * Renders the React component tree into the web part DOM element.
   * Computes which fields are "display" fields (i.e., not mapped to a core
   * slot like title/start/end) and passes the full configuration down as props.
   */
  public render(): void {
    // Core fields already mapped to dedicated slots — exclude from "extra" display fields
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
          // `!== false` ensures true when undefined (opt-in default)
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

  /**
   * Responds to SharePoint theme changes by updating the dark-theme flag
   * and injecting semantic color CSS variables onto the web part root element.
   * These variables are consumed by component SCSS modules.
   * @param currentTheme - The incoming SharePoint theme, or undefined.
   */
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

  /**
   * Cleans up the React component tree when the web part is disposed.
   */
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  /**
   * Returns the semantic data version for serialization.
   * @returns The current data version (1.0).
   */
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  /**
   * Handles property pane field changes with cascading logic:
   *  - List change: resets all mappings, reloads fields, runs auto-detection.
   *  - View mode change: refreshes pane to show/hide conditional sections.
   *  - Core field mapping change: refreshes pane and re-renders.
   *  - Card display toggle or layout change: re-renders only.
   *  - Extra field checkbox (`field_*`): toggles the field in selectedFields.
   *
   * @param propertyPath - The property bag key that changed.
   * @param _oldValue - The previous value (unused but required by signature).
   * @param _newValue - The new value assigned to the property.
   */
  protected async onPropertyPaneFieldChanged(propertyPath: string, _oldValue: unknown, _newValue: unknown): Promise<void> {
    if (propertyPath === 'selectedListId' && _newValue) {
      // Reset all field mappings when the source list changes
      this.properties.selectedFields = [];
      this.properties.titleField = '';
      this.properties.startDateField = '';
      this.properties.endDateField = '';
      this.properties.allDayField = '';
      this.properties.categoryField = '';
      this.properties.locationField = '';

      await this._loadFields(_newValue as string);

      // Auto-detect maps common SharePoint field names (e.g., "EventDate", "Location")
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

    // View mode drives conditional visibility of calendar/list option groups
    if (propertyPath === 'viewMode') {
      this.context.propertyPane.refresh();
      this.render();
    }

    // Core column mapping changes affect which fields appear in the "extra fields" group
    if (['titleField', 'startDateField', 'endDateField', 'allDayField',
         'categoryField', 'locationField'].indexOf(propertyPath) >= 0) {
      this.context.propertyPane.refresh();
      this.render();
    }

    // Card display toggles and list layout only need a visual re-render
    if (['showCategory', 'showLocation', 'showTime', 'showImage',
         'listLayout'].indexOf(propertyPath) >= 0) {
      this.render();
    }

    // Extra field checkboxes use a `field_` prefix convention to namespace
    // dynamic property keys without colliding with static web part properties
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
      // Spread into a new array to ensure React detects the prop change
      this.properties.selectedFields = [...selected];
      this.render();
    }
  }

  /**
   * Fetches all user-visible fields from the selected SharePoint list and
   * caches them in `_allFields`. Silently swallows errors so the property
   * pane can still render a fallback message.
   * @param listId - The GUID of the SharePoint list to query.
   */
  private async _loadFields(listId: string): Promise<void> {
    try {
      this._allFields = await fetchAllListFields(listId);
    } catch {
      this._allFields = [];
    }
  }

  /**
   * Builds dropdown options from cached fields, optionally filtering by
   * SharePoint field type (e.g., "DateTime", "Boolean"). Includes a
   * leading "(none)" option so the user can clear a mapping.
   * @param filterTypes - Optional array of field type strings to include.
   * @returns Dropdown option array with `key` (internal name) and `text` (display label).
   */
  private _getFieldDropdownOptions(filterTypes?: string[]): Array<{ key: string; text: string }> {
    const options: Array<{ key: string; text: string }> = [{ key: '', text: '(none)' }];
    for (const f of this._allFields) {
      if (!filterTypes || filterTypes.indexOf(f.fieldType) >= 0) {
        options.push({ key: f.internalName, text: f.displayName + ' (' + f.fieldType + ')' });
      }
    }
    return options;
  }

  /**
   * Returns dropdown options restricted to text-like field types suitable
   * for the Title mapping (Text, Note, Choice, Computed).
   * @returns Dropdown option array without a "(none)" entry — title is required.
   */
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

  /**
   * Returns dropdown options for the Category field mapping, including a
   * "(none)" fallback since category is optional.
   * @returns Dropdown option array filtered to Text, Choice, and Computed types.
   */
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

  /**
   * Returns dropdown options for the Location field mapping, including a
   * "(none)" fallback since location is optional.
   * @returns Dropdown option array filtered to Text, Note, Choice, and Computed types.
   */
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

  /**
   * Builds the two-page property pane configuration:
   *
   * **Page 1 — Data Source**
   *  - List picker (PnP control)
   *  - Column mapping dropdowns (title, start, end, allDay, category, location)
   *  - Extra field checkboxes (all non-mapped fields the user can opt in to display)
   *
   * **Page 2 — Display Settings**
   *  - View mode (both / calendar only / list only)
   *  - Default view when "both" is selected
   *  - Calendar sub-type (month / week / day)
   *  - List layout (filmstrip / compact)
   *  - Max events slider
   *  - Card display toggles (image, category, time, location)
   *
   * @returns The property pane configuration object consumed by SPFx.
   */
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    // Exclude fields already assigned to a core mapping slot
    const mappedCoreFields = [
      this.properties.titleField,
      this.properties.startDateField,
      this.properties.endDateField,
      this.properties.allDayField,
      this.properties.categoryField,
      this.properties.locationField,
    ].filter(Boolean);
    const displayFields = getDisplayFields(this._allFields, mappedCoreFields);

    // Generate one checkbox per non-mapped field for the "Extra Fields" group
    const fieldCheckboxes = displayFields.map(field =>
      PropertyPaneCheckbox(`field_${field.internalName}`, {
        text: field.displayName,
        checked: (this.properties.selectedFields || []).indexOf(field.internalName) >= 0,
      })
    );

    const viewMode = this.properties.viewMode || 'both';
    const showCalendarOptions = viewMode === 'both' || viewMode === 'calendar';
    const showListOptions = viewMode === 'both' || viewMode === 'list';

    // Column mapping fields — only shown after fields are loaded from the list
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
      // Show a placeholder label until the user picks a list
      mappingFields.push(
        PropertyPaneLabel('noMappingFields', {
          text: this.properties.selectedListId ? 'Loading fields...' : 'Select a list first',
        }) as IPropertyPaneField<unknown>,
      );
    }

    // View settings — conditionally includes calendar/list options based on viewMode
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

    // Default view toggle only makes sense when both views are enabled
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

    // Card display toggles — control which metadata appears on event cards
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
