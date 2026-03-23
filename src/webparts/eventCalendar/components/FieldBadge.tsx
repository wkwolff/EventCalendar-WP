import * as React from 'react';
import { IFieldInfo } from '../models/IFieldInfo';

export interface IFieldBadgeProps {
  field: IFieldInfo;
  value: unknown;
  /** Render in detail mode (full width, richer display) vs compact badge */
  detailed?: boolean;
}

// ── Styles ──

const badgeStyle: React.CSSProperties = {
  display: 'inline-flex',
  alignItems: 'center',
  padding: '3px 10px',
  margin: '2px 4px 2px 0',
  borderRadius: 12,
  fontSize: 12,
  lineHeight: '18px',
  maxWidth: '100%',
  overflow: 'hidden',
  textOverflow: 'ellipsis',
  whiteSpace: 'nowrap',
};

const detailRowStyle: React.CSSProperties = {
  marginBottom: 16,
  paddingBottom: 12,
  borderBottom: '1px solid #edebe9',
};

const detailLabelStyle: React.CSSProperties = {
  fontWeight: 600,
  fontSize: 12,
  color: '#605e5c',
  marginBottom: 6,
  textTransform: 'uppercase',
  letterSpacing: '0.5px',
};

// ── Badge color palette by field type ──

// All badges use the site theme for a cohesive look
function getBadgeColors(): { bg: string; color: string } {
  // Resolve SPFx theme tokens at runtime via CSS variables / window.__themeState__
  const themeState = (window as unknown as Record<string, unknown>).__themeState__ as
    { theme?: Record<string, string> } | undefined;
  const theme = themeState?.theme;

  const bg = theme?.themeLighter || '#edebe9';
  const color = theme?.themeDarkAlt || '#323130';
  return { bg, color };
}

// ── Helpers ──

/** Strip HTML tags and return plain text */
function stripHtml(html: string): string {
  const div = document.createElement('div');
  div.innerHTML = html;
  return div.textContent || div.innerText || '';
}

/** Check if a string contains HTML tags */
function isHtml(str: string): boolean {
  return /<[a-z][\s\S]*>/i.test(str);
}

/** Check if a URL points to an image */
function isImageUrl(url: string): boolean {
  if (!url) return false;
  const lower = url.toLowerCase();
  return /\.(jpg|jpeg|png|gif|bmp|webp|svg)(\?|$)/i.test(lower);
}

/** Check if a field is likely an image field based on name */
function isImageField(field: IFieldInfo): boolean {
  const name = (field.internalName + ' ' + field.displayName).toLowerCase();
  return name.indexOf('image') >= 0 || name.indexOf('banner') >= 0 || name.indexOf('photo') >= 0 || name.indexOf('picture') >= 0 || name.indexOf('thumbnail') >= 0;
}

/** Get URL string from a URL field value */
function getUrlFromValue(value: unknown): string {
  if (typeof value === 'string') return value;
  if (typeof value === 'object' && value !== null) {
    return (value as Record<string, unknown>).Url as string || '';
  }
  return '';
}

/** Returns true if the value is empty, null, undefined, or a meaningless default */
export function isEmptyValue(value: unknown, fieldType: string): boolean {
  if (value === null || value === undefined || value === '') return true;
  if (fieldType === 'Note' && typeof value === 'string') {
    const stripped = stripHtml(value).trim();
    return stripped.length === 0;
  }
  return false;
}

/** Check if a field+value represents an image that should be rendered */
export function isImageFieldValue(field: IFieldInfo, value: unknown): boolean {
  if (field.fieldType === 'URL' || field.fieldType === 'Image') {
    const url = getUrlFromValue(value);
    return isImageUrl(url) || isImageField(field);
  }
  if (field.fieldType === 'Thumbnail') return true;
  return false;
}

/** Get the image URL from a field value */
export function getImageUrl(value: unknown): string {
  return getUrlFromValue(value);
}

/** Format a value into a short text string for badges */
function formatCompact(value: unknown, fieldType: string): string {
  switch (fieldType) {
    case 'DateTime':
      return new Date(value as string).toLocaleDateString();
    case 'Boolean':
    case 'AllDayEvent':
      return value ? 'Yes' : 'No';
    case 'Currency':
      return '$' + Number(value).toFixed(2);
    case 'Number':
      return String(value);
    case 'Lookup':
    case 'User':
      if (typeof value === 'object' && value !== null) {
        return (value as Record<string, unknown>).Title as string || String(value);
      }
      return String(value);
    case 'MultiChoice':
      if (Array.isArray(value)) return (value as string[]).join(', ');
      return String(value);
    case 'URL':
      if (typeof value === 'object' && value !== null) {
        const urlObj = value as Record<string, unknown>;
        return (urlObj.Description as string) || (urlObj.Url as string) || '';
      }
      return String(value);
    case 'Note':
      if (typeof value === 'string' && isHtml(value)) {
        return stripHtml(value).substring(0, 120);
      }
      return String(value).substring(0, 120);
    default:
      return String(value);
  }
}

// ── Auto-link detection ──

const EMAIL_REGEX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
const URL_REGEX = /^https?:\/\/\S+$/i;
const SP_PATH_REGEX = /^\/sites\/\S+/i;

/** Check if a plain string value looks like an email */
function isEmail(str: string): boolean {
  return EMAIL_REGEX.test(str.trim());
}

/** Check if a plain string value looks like a URL or SharePoint file path */
function isUrl(str: string): boolean {
  const trimmed = str.trim();
  return URL_REGEX.test(trimmed) || SP_PATH_REGEX.test(trimmed);
}

const linkStyle: React.CSSProperties = {
  color: '#0078d4',
  textDecoration: 'none',
};

// ── Components ──

/** Compact badge for card view */
const CompactBadge: React.FC<{ field: IFieldInfo; value: unknown }> = ({ field, value }) => {
  const fieldType = field.fieldType;
  const colors = getBadgeColors();

  // Image fields — render as small thumbnail in card
  if (isImageFieldValue(field, value)) {
    const url = getImageUrl(value);
    if (url) {
      return (
        <img
          src={url}
          alt={field.displayName}
          style={{
            width: '100%',
            maxHeight: 120,
            objectFit: 'cover',
            borderRadius: 4,
            marginTop: 4,
          }}
        />
      );
    }
  }

  // URL fields → clickable link
  if (fieldType === 'URL' && typeof value === 'object' && value !== null) {
    const urlObj = value as Record<string, unknown>;
    const url = urlObj.Url as string;
    const desc = (urlObj.Description as string) || url;
    if (url) {
      return (
        <span style={{ ...badgeStyle, backgroundColor: colors.bg, color: colors.color }}>
          <strong style={{ marginRight: 4 }}>{field.displayName}:</strong>
          <a
            href={url}
            target="_blank"
            rel="noopener noreferrer"
            onClick={(e) => e.stopPropagation()}
            style={{ color: '#0078d4', textDecoration: 'none' }}
          >
            {desc}
          </a>
        </span>
      );
    }
  }

  const display = formatCompact(value, fieldType);
  const strVal = typeof value === 'string' ? value.trim() : '';

  // Auto-link emails
  if (strVal && isEmail(strVal)) {
    return (
      <span style={{ ...badgeStyle, backgroundColor: colors.bg, color: colors.color }}>
        <strong style={{ marginRight: 4 }}>{field.displayName}:</strong>
        <a
          href={'mailto:' + strVal}
          onClick={(e) => e.stopPropagation()}
          style={linkStyle}
        >
          {strVal}
        </a>
      </span>
    );
  }

  // Auto-link URLs in text fields
  if (strVal && isUrl(strVal)) {
    return (
      <span style={{ ...badgeStyle, backgroundColor: colors.bg, color: colors.color }}>
        <strong style={{ marginRight: 4 }}>{field.displayName}:</strong>
        <a
          href={strVal}
          target="_blank"
          rel="noopener noreferrer"
          onClick={(e) => e.stopPropagation()}
          style={linkStyle}
        >
          {strVal}
        </a>
      </span>
    );
  }

  return (
    <span
      style={{ ...badgeStyle, backgroundColor: colors.bg, color: colors.color }}
      title={field.displayName + ': ' + display}
    >
      <strong style={{ marginRight: 4 }}>{field.displayName}:</strong> {display}
    </span>
  );
};

/** Detail row for panel view */
const DetailRow: React.FC<{ field: IFieldInfo; value: unknown }> = ({ field, value }) => {
  const fieldType = field.fieldType;

  // Image fields — render full-width image
  if (isImageFieldValue(field, value)) {
    const url = getImageUrl(value);
    if (url) {
      return (
        <div style={detailRowStyle}>
          <div style={detailLabelStyle}>{field.displayName}</div>
          <img
            src={url}
            alt={field.displayName}
            style={{
              width: '100%',
              maxHeight: 300,
              objectFit: 'cover',
              borderRadius: 6,
            }}
          />
        </div>
      );
    }
  }

  // URL field → clickable link
  if (fieldType === 'URL' && typeof value === 'object' && value !== null) {
    const urlObj = value as Record<string, unknown>;
    const url = urlObj.Url as string;
    const desc = (urlObj.Description as string) || url;
    if (url) {
      return (
        <div style={detailRowStyle}>
          <div style={detailLabelStyle}>{field.displayName}</div>
          <div>
            <a
              href={url}
              target="_blank"
              rel="noopener noreferrer"
              style={{ color: '#0078d4' }}
            >
              {desc}
            </a>
          </div>
        </div>
      );
    }
  }

  // Rich text (Note) → render HTML
  if (fieldType === 'Note' && typeof value === 'string' && isHtml(value)) {
    return (
      <div style={detailRowStyle}>
        <div style={detailLabelStyle}>{field.displayName}</div>
        <div
          style={{ fontSize: 13, lineHeight: '1.5' }}
          dangerouslySetInnerHTML={{ __html: value }}
        />
      </div>
    );
  }

  // Boolean
  if (fieldType === 'Boolean' || fieldType === 'AllDayEvent') {
    return (
      <div style={detailRowStyle}>
        <div style={detailLabelStyle}>{field.displayName}</div>
        <div style={{ fontSize: 13 }}>{value ? 'Yes' : 'No'}</div>
      </div>
    );
  }

  // MultiChoice
  if (fieldType === 'MultiChoice' && Array.isArray(value)) {
    const pillColors = getBadgeColors();
    return (
      <div style={detailRowStyle}>
        <div style={detailLabelStyle}>{field.displayName}</div>
        <div style={{ fontSize: 13, display: 'flex', flexWrap: 'wrap', gap: 4 }}>
          {(value as string[]).map((v: string, i: number) => (
            <span
              key={i}
              style={{
                display: 'inline-block',
                padding: '3px 10px',
                borderRadius: 12,
                backgroundColor: pillColors.bg,
                color: pillColors.color,
                fontSize: 12,
              }}
            >
              {v}
            </span>
          ))}
        </div>
      </div>
    );
  }

  // Choice — colored pill
  if (fieldType === 'Choice') {
    const pillColors = getBadgeColors();
    return (
      <div style={detailRowStyle}>
        <div style={detailLabelStyle}>{field.displayName}</div>
        <span style={{
          display: 'inline-block',
          padding: '3px 12px',
          borderRadius: 12,
          backgroundColor: pillColors.bg,
          color: pillColors.color,
          fontSize: 13,
          fontWeight: 500,
        }}>
          {String(value)}
        </span>
      </div>
    );
  }

  const display = formatCompact(value, fieldType);
  const strVal = typeof value === 'string' ? value.trim() : '';

  // Auto-link emails
  if (strVal && isEmail(strVal)) {
    return (
      <div style={detailRowStyle}>
        <div style={detailLabelStyle}>{field.displayName}</div>
        <div style={{ fontSize: 13 }}>
          <a href={'mailto:' + strVal} style={linkStyle}>{strVal}</a>
        </div>
      </div>
    );
  }

  // Auto-link URLs (including links to files)
  if (strVal && isUrl(strVal)) {
    return (
      <div style={detailRowStyle}>
        <div style={detailLabelStyle}>{field.displayName}</div>
        <div style={{ fontSize: 13 }}>
          <a href={strVal} target="_blank" rel="noopener noreferrer" style={linkStyle}>{strVal}</a>
        </div>
      </div>
    );
  }

  return (
    <div style={detailRowStyle}>
      <div style={detailLabelStyle}>{field.displayName}</div>
      <div style={{ fontSize: 13 }}>{display}</div>
    </div>
  );
};

const FieldBadge: React.FC<IFieldBadgeProps> = ({ field, value, detailed }) => {
  if (isEmptyValue(value, field.fieldType)) return null;

  if (detailed) {
    return <DetailRow field={field} value={value} />;
  }
  return <CompactBadge field={field} value={value} />;
};

export default FieldBadge;
