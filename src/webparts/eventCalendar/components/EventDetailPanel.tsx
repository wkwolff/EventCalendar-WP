/**
 * @file EventDetailPanel.tsx
 * @description Fluent UI side panel that displays full event details when an event
 *   is clicked. Renders a hero image (first image-type field), date/time section,
 *   and all user-selected extra fields using the FieldBadge component in "detailed"
 *   mode. Image fields are separated from regular fields for layout prioritization.
 * @author W. Kevin Wolff
 * @copyright Wolff Creative
 */

import * as React from 'react';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { Icon } from '@fluentui/react/lib/Icon';
import { IEventItem } from '../models/IEventItem';
import { IFieldInfo } from '../models/IFieldInfo';
import FieldBadge, { isEmptyValue, isImageFieldValue, getImageUrl } from './FieldBadge';

/**
 * Props for the EventDetailPanel component.
 */
export interface IEventDetailPanelProps {
  /** The event to display, or undefined when no event is selected. */
  event: IEventItem | undefined;
  /** Metadata for user-selected display fields. */
  availableFields: IFieldInfo[];
  /** Whether the panel is currently open. */
  isOpen: boolean;
  /** Callback to close the panel and deselect the event. */
  onDismiss: () => void;
}

/**
 * Formats an ISO date string into a full date-time string.
 * Example output: "Wednesday, January 15, 2025 at 2:00 PM"
 * @param dateStr - ISO date string to format.
 * @returns Formatted date-time string with day of week.
 */
function formatDateTime(dateStr: string): string {
  const d = new Date(dateStr);
  return d.toLocaleDateString(undefined, {
    weekday: 'long',
    month: 'long',
    day: 'numeric',
    year: 'numeric',
  }) + ' at ' + d.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' });
}

/**
 * Formats an ISO date string into a date-only string (no time component).
 * Used for all-day events where displaying a time would be misleading.
 * @param dateStr - ISO date string to format.
 * @returns Formatted date string with day of week.
 */
function formatDateOnly(dateStr: string): string {
  const d = new Date(dateStr);
  return d.toLocaleDateString(undefined, {
    weekday: 'long',
    month: 'long',
    day: 'numeric',
    year: 'numeric',
  });
}

/**
 * Renders a Fluent UI medium-width Panel with full event details including:
 *  1. Hero image (first image-type extra field, if available)
 *  2. Date/time with clock icon
 *  3. Additional image fields (rendered via FieldBadge in detailed mode)
 *  4. All non-image extra fields (rendered via FieldBadge in detailed mode)
 *
 * @param props - Panel configuration with event data and field metadata.
 * @returns A Fluent UI Panel element, or null when no event is selected.
 */
const EventDetailPanel: React.FC<IEventDetailPanelProps> = ({
  event,
  availableFields,
  isOpen,
  onDismiss,
}) => {
  if (!event) return null;

  // Build a lookup map for O(1) field metadata access
  const fieldMap = new Map(availableFields.map(f => [f.internalName, f]));

  // Separate image fields from regular fields:
  // - The first image becomes the hero banner at the top of the panel.
  // - Additional images are rendered below the date/time section.
  // - Non-image fields are rendered as detailed rows at the bottom.
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
          // First image field becomes the hero banner
          heroImageUrl = url;
        } else {
          // Subsequent images go to the secondary images section
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
        {/* Hero image — full-bleed banner with negative margin to span panel padding */}
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

        {/* Date/time section with clock icon */}
        <div style={{
          display: 'flex',
          alignItems: 'flex-start',
          gap: 12,
          marginBottom: 24,
        }}>
          <Icon iconName="Clock" style={{ fontSize: 16, color: '#605e5c', marginTop: 2 }} />
          <div>
            <div style={{ fontSize: 14, color: '#323130' }}>
              {/* All-day events show date only; timed events show start-to-end range */}
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

        {/* Additional image fields (beyond the hero) */}
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

        {/* Non-image field details — rendered as labeled rows */}
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
