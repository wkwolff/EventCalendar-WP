import { DefaultView, CalendarViewType, ViewMode, ListLayout } from '../models/IWebPartProps';
import { IFieldInfo } from '../models/IFieldInfo';
import { IEventFieldMapping } from '../services/EventService';

export interface ICardDisplayOptions {
  showCategory: boolean;
  showLocation: boolean;
  showTime: boolean;
  showImage: boolean;
}

export interface IEventCalendarProps {
  listId: string;
  fieldMapping: IEventFieldMapping;
  selectedFields: string[];
  availableFields: IFieldInfo[];
  viewMode: ViewMode;
  defaultView: DefaultView;
  calendarViewType: CalendarViewType;
  listLayout: ListLayout;
  maxEvents: number;
  cardDisplay: ICardDisplayOptions;
  isDarkTheme: boolean;
  hasTeamsContext: boolean;
}
