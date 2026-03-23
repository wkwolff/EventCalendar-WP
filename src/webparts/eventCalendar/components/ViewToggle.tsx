/**
 * @file ViewToggle.tsx
 * @description A pair of icon buttons that allow the user to switch between
 *   list view and calendar view. Only rendered when the web part is configured
 *   to show both views. Uses Fluent UI IconButton with inline theme-aware styling.
 * @author W. Kevin Wolff
 * @copyright Wolff Creative
 */

import * as React from 'react';
import { IconButton } from '@fluentui/react/lib/Button';
import { DefaultView } from '../models/IWebPartProps';

/**
 * Props for the ViewToggle component.
 */
export interface IViewToggleProps {
  /** The currently active view ('list' or 'calendar'). */
  currentView: DefaultView;
  /** Callback to switch the active view. */
  onViewChange: (view: DefaultView) => void;
}

/**
 * Renders two side-by-side icon buttons (list and calendar) with active-state
 * highlighting. The active button uses Fluent UI's brand blue (#0078d4) with a
 * light blue background; the inactive button is neutral gray.
 *
 * @param props - Current view state and change handler.
 * @returns A flex container with two icon toggle buttons.
 */
const ViewToggle: React.FC<IViewToggleProps> = ({ currentView, onViewChange }) => {
  return (
    <div style={{ display: 'flex', gap: 2 }}>
      <IconButton
        iconProps={{ iconName: 'BulletedList' }}
        title="List view"
        ariaLabel="List view"
        checked={currentView === 'list'}
        onClick={() => onViewChange('list')}
        styles={{
          root: {
            color: currentView === 'list' ? '#0078d4' : '#605e5c',
            backgroundColor: currentView === 'list' ? '#eff6fc' : 'transparent',
            borderRadius: 2,
            width: 32,
            height: 32,
          },
          rootHovered: { backgroundColor: '#f3f2f1' },
        }}
      />
      <IconButton
        iconProps={{ iconName: 'Calendar' }}
        title="Calendar view"
        ariaLabel="Calendar view"
        checked={currentView === 'calendar'}
        onClick={() => onViewChange('calendar')}
        styles={{
          root: {
            color: currentView === 'calendar' ? '#0078d4' : '#605e5c',
            backgroundColor: currentView === 'calendar' ? '#eff6fc' : 'transparent',
            borderRadius: 2,
            width: 32,
            height: 32,
          },
          rootHovered: { backgroundColor: '#f3f2f1' },
        }}
      />
    </div>
  );
};

export default ViewToggle;
