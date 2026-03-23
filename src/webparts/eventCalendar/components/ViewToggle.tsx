import * as React from 'react';
import { IconButton } from '@fluentui/react/lib/Button';
import { DefaultView } from '../models/IWebPartProps';

export interface IViewToggleProps {
  currentView: DefaultView;
  onViewChange: (view: DefaultView) => void;
}

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
