import * as React from 'react';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { Icon } from '@fluentui/react/lib/Icon';
import { IEventItem } from '../models/IEventItem';
import { IFieldInfo } from '../models/IFieldInfo';
import FieldBadge, { isEmptyValue, isImageFieldValue, getImageUrl } from './FieldBadge';

export interface IEventDetailPanelProps {
  event: IEventItem | undefined;
  availableFields: IFieldInfo[];
  isOpen: boolean;
  onDismiss: () => void;
}

function formatDateTime(dateStr: string): string {
  const d = new Date(dateStr);
  return d.toLocaleDateString(undefined, {
    weekday: 'long',
    month: 'long',
    day: 'numeric',
    year: 'numeric',
  }) + ' at ' + d.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' });
}

function formatDateOnly(dateStr: string): string {
  const d = new Date(dateStr);
  return d.toLocaleDateString(undefined, {
    weekday: 'long',
    month: 'long',
    day: 'numeric',
    year: 'numeric',
  });
}

const EventDetailPanel: React.FC<IEventDetailPanelProps> = ({
  event,
  availableFields,
  isOpen,
  onDismiss,
}) => {
  if (!event) return null;

  const fieldMap = new Map(availableFields.map(f => [f.internalName, f]));

  // Separate images and regular fields
  let heroImageUrl: string | undefined;
  const imageFields: Array<{ key: string; field: IFieldInfo; value: unknown }> = [];
  const visibleFields: Array<{ key: string; field: IFieldInfo; value: unknown }> = [];

  Object.keys(event.fields).forEach((key: string) => {
    const value = event.fields[key];
    const field = fieldMap.get(key);
    if (!field || isEmptyValue(value, field.fieldType)) return;

    if (isImageFieldValue(field, value)) {
      const url = getImageUrl(value);
      if (url) {
        if (!heroImageUrl) {
          heroImageUrl = url;
        } else {
          imageFields.push({ key, field, value });
        }
        return;
      }
    }
    visibleFields.push({ key, field, value });
  });

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      type={PanelType.medium}
      headerText={event.title}
      closeButtonAriaLabel="Close"
    >
      <div style={{ padding: '8px 0' }}>
        {/* Hero image */}
        {heroImageUrl && (
          <div style={{
            margin: '0 -24px 20px',
            height: 200,
            overflow: 'hidden',
            background: '#f3f2f1',
          }}>
            <img
              src={heroImageUrl}
              alt={event.title}
              style={{
                width: '100%',
                height: '100%',
                objectFit: 'cover',
                display: 'block',
              }}
            />
          </div>
        )}

        {/* Date/time section */}
        <div style={{
          display: 'flex',
          alignItems: 'flex-start',
          gap: 12,
          marginBottom: 24,
        }}>
          <Icon iconName="Clock" style={{ fontSize: 16, color: '#605e5c', marginTop: 2 }} />
          <div>
            <div style={{ fontSize: 14, color: '#323130' }}>
              {event.allDay
                ? formatDateOnly(event.startDate)
                : formatDateTime(event.startDate) + ' \u2013 ' + formatDateTime(event.endDate)}
            </div>
            {event.allDay && (
              <div style={{ fontSize: 12, color: '#605e5c', marginTop: 2 }}>
                All day
              </div>
            )}
          </div>
        </div>

        {/* Additional images */}
        {imageFields.length > 0 && (
          <div style={{ marginBottom: 20 }}>
            {imageFields.map((item) => (
              <FieldBadge
                key={item.key}
                field={item.field}
                value={item.value}
                detailed
              />
            ))}
          </div>
        )}

        {/* Field details */}
        {visibleFields.length > 0 && (
          <div>
            {visibleFields.map((item) => (
              <FieldBadge
                key={item.key}
                field={item.field}
                value={item.value}
                detailed
              />
            ))}
          </div>
        )}
      </div>
    </Panel>
  );
};

export default EventDetailPanel;
