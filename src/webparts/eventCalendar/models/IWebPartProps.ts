export type CalendarViewType = 'dayGridMonth' | 'timeGridWeek' | 'timeGridDay';
export type DefaultView = 'calendar' | 'list';
export type ViewMode = 'both' | 'calendar' | 'list';
export type ListLayout = 'filmstrip' | 'compact';

export interface IEventCalendarWebPartProps {
  selectedListId: string;
  titleField: string;
  startDateField: string;
  endDateField: string;
  allDayField: string;
  categoryField: string;
  locationField: string;
  selectedFields: string[];
  viewMode: ViewMode;
  defaultView: DefaultView;
  calendarViewType: CalendarViewType;
  listLayout: ListLayout;
  maxEvents: number;
  showCategory: boolean;
  showLocation: boolean;
  showTime: boolean;
  showImage: boolean;
}
