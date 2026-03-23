declare interface IEventCalendarWebPartStrings {
  PropertyPaneDescription: string;
  DataSourceGroupName: string;
  ListPickerLabel: string;
  ColumnMappingGroupName: string;
  TitleFieldLabel: string;
  StartDateFieldLabel: string;
  EndDateFieldLabel: string;
  AllDayFieldLabel: string;
  CategoryFieldLabel: string;
  LocationFieldLabel: string;
  FieldsGroupName: string;
  DisplaySettingsDescription: string;
  ViewSettingsGroupName: string;
  ViewModeLabel: string;
  DefaultViewLabel: string;
  CalendarViewTypeLabel: string;
  ListLayoutLabel: string;
  MaxEventsLabel: string;
  CardDisplayGroupName: string;
  ShowCategoryLabel: string;
  ShowLocationLabel: string;
  ShowTimeLabel: string;
  ShowImageLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  UnknownEnvironment: string;
}

declare module 'EventCalendarWebPartStrings' {
  const strings: IEventCalendarWebPartStrings;
  export = strings;
}
