export interface IEventItem {
  id: number;
  title: string;
  startDate: string;
  endDate: string;
  allDay: boolean;
  category: string;
  location: string;
  imageUrl: string;
  fields: Record<string, unknown>;
}
