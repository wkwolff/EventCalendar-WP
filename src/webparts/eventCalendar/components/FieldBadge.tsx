/**
 * @file FieldBadge.tsx
 * @description Polymorphic field renderer that displays SharePoint list field values
 *   in two modes: "compact" (single-line badge for event cards) and "detailed"
 *   (full-width labeled row for the detail panel). Handles type-specific rendering
 *   for images, URLs, rich text (Note), booleans, choices, multi-choice pills,
 *   currency, lookups, and plain text. Includes auto-link detection for emails,
 *   HTTP URLs, and SharePoint file paths. Badge colors are resolved at runtime
 *   from the SPFx theme.
 * @author W. Kevin Wolff
 * @copyright Wolff Creative
 */

import * as React from 'react';
import { IFieldInfo } from '../models/IFieldInfo';

/**
 * Props for the FieldBadge component.
 */
export interface IFieldBadgeProps {
  /** SharePoint field metadata (internal name, display name, field type). */
  field: IFieldInfo;
  /** The raw field value from the SharePoint list item. */
  value: unknown;
  /** Render in detail mode (full width, richer display) vs compact badge */
  detailed?: boolean;
}

// ── Inline Styles ──────────────────────────────────────────────────────────────

/** Base style for compact badge text — single-line with ellipsis overflow. */
const badgeStyle: React.CSSProperties = {
  display: 'block',
  fontSize: 12,
  lineHeight: '1.4',
  color: '#605e5c',
  overflow: 'hidden',
  textOverflow: 'ellipsis',
  whiteSpace: 'nowrap',
};

/** Container style for detail panel rows — includes a bottom separator. */
const detailRowStyle: React.CSSProperties = {
  marginBottom: 16,
  paddingBottom: 12,
  borderBottom: '1px solid #edebe9',
};

/** Label style for detail panel rows — uppercase, subtle gray. */
const detailLabelStyle: React.CSSProperties = {
  fontWeight: 600,
  fontSize: 12,
  color: '#605e5c',
  marginBottom: 6,
  textTransform: 'uppercase',
  letterSpacing: '0.5px',
};

// ── Theme-Aware Badge Colors ───────────────────────────────────────────────────

/**
 * Resolves badge background and text colors from the SPFx runtime theme.
 * Falls back to neutral gray tones when theme tokens are unavailable
 * (e.g., in local workbench or non-SharePoint hosts).
 *
 * Accesses `window.__themeState__` which is injected by the SharePoint
 * page framework at runtime — not available at build time.
 *
 * @returns An object with `bg` (background) and `color` (text) hex values.
 */
function getBadgeColors(): { bg: string; color: string } {
  const themeState = (window as unknown as Record<string, unknown>).__themeState__ as
    { theme?: Record<string, string> } | undefined;
  const theme = themeState?.theme;

  const bg = theme?.themeLighter || '#edebe9';
  const color = theme?.themeDarkAlt || '#323130';
  return { bg, color };
}

// ── HTML & Content Helpers ─────────────────────────────────────────────────────

/**
 * Strips HTML tags from a string and returns plain text content.
 * Uses a temporary DOM element for safe parsing.
 * @param html - HTML string to strip.
 * @returns Plain text content.
 */
function stripHtml(html: string): string {
  const div = document.createElement('div');
  div.innerHTML = html;
  return div.textContent || div.innerText || '';
}

/**
 * Tests whether a string contains HTML tags.
 * Used to determine if a Note field value needs HTML stripping or rendering.
 * @param str - String to test.
 * @returns True if the string contains HTML markup.
 */
function isHtml(str: string): boolean {
  return /<[a-z][\s\S]*>/i.test(str);
}

/**
 * Tests whether a URL string points to a common image file format.
 * Supports jpg, jpeg, png, gif, bmp, webp, and svg extensions.
 * @param url - URL string to test.
 * @returns True if the URL ends with an image file extension.
 */
function isImageUrl(url: string): boolean {
  if (!url) return false;
  const lower = url.toLowerCase();
  return /\.(jpg|jpeg|png|gif|bmp|webp|svg)(\?|$)/i.test(lower);
}

/**
 * Heuristic check: determines if a field is likely an image field based on
 * its internal or display name containing keywords like "image", "banner",
 * "photo", "picture", or "thumbnail".
 * @param field - SharePoint field metadata.
 * @returns True if the field name suggests image content.
 */
function isImageField(field: IFieldInfo): boolean {
  const name = (field.internalName + ' ' + field.displayName).toLowerCase();
  return name.indexOf('image') >= 0 || name.indexOf('banner') >= 0 || name.indexOf('photo') >= 0 || name.indexOf('picture') >= 0 || name.indexOf('thumbnail') >= 0;
}

/**
 * Extracts a URL string from a field value that may be a plain string or
 * a SharePoint URL field object (`{ Url: string, Description: string }`).
 * @param value - Raw field value.
 * @returns The extracted URL string, or empty string if not found.
 */
function getUrlFromValue(value: unknown): string {
  if (typeof value === 'string') return value;
  if (typeof value === 'object' && value !== null) {
    return (value as Record<string, unknown>).Url as string || '';
  }
  return '';
}

/**
 * Determines if a field value is effectively empty and should not be rendered.
 * Handles null, undefined, empty string, and Note fields whose HTML content
 * reduces to empty text after stripping tags.
 * @param value - Raw field value.
 * @param fieldType - SharePoint field type string.
 * @returns True if the value should be considered empty.
 */
export function isEmptyValue(value: unknown, fieldType: string): boolean {
  if (value === null || value === undefined || value === '') return true;
  if (fieldType === 'Note' && typeof value === 'string') {
    const stripped = stripHtml(value).trim();
    return stripped.length === 0;
  }
  return false;
}

/**
 * Determines if a field+value combination represents an image that should
 * be rendered as an `<img>` element rather than as text. Checks URL/Image
 * field types by extension and field name heuristics, and always treats
 * Thumbnail fields as images.
 * @param field - SharePoint field metadata.
 * @param value - Raw field value.
 * @returns True if the value should be rendered as an image.
 */
export function isImageFieldValue(field: IFieldInfo, value: unknown): boolean {
  if (field.fieldType === 'URL' || field.fieldType === 'Image') {
    const url = getUrlFromValue(value);
    return isImageUrl(url) || isImageField(field);
  }
  if (field.fieldType === 'Thumbnail') return true;
  return false;
}

/**
 * Extracts the image URL from a field value (string or URL object).
 * @param value - Raw field value.
 * @returns The image URL string.
 */
export function getImageUrl(value: unknown): string {
  return getUrlFromValue(value);
}

/**
 * Formats a raw field value into a short display string suitable for compact
 * badges. Applies type-specific formatting (dates, currency, booleans, lookups,
 * multi-choice joins, URL descriptions, and Note truncation to 120 chars).
 * @param value - Raw field value.
 * @param fieldType - SharePoint field type string.
 * @returns A short formatted string representation.
 */
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
      // SharePoint lookup/user values are objects with a Title property
      if (typeof value === 'object' && value !== null) {
        return (value as Record<string, unknown>).Title as string || String(value);
      }
      return String(value);
    case 'MultiChoice':
      if (Array.isArray(value)) return (value as string[]).join(', ');
      return String(value);
    case 'URL':
      // URL fields have { Url, Description } — prefer Description for display
      if (typeof value === 'object' && value !== null) {
        const urlObj = value as Record<string, unknown>;
        return (urlObj.Description as string) || (urlObj.Url as string) || '';
      }
      return String(value);
    case 'Note':
      // Truncate rich text to 120 characters for card display
      if (typeof value === 'string' && isHtml(value)) {
        return stripHtml(value).substring(0, 120);
      }
      return String(value).substring(0, 120);
    default:
      return String(value);
  }
}

// ── Auto-Link Detection ────────────────────────────────────────────────────────

/** Regex for validating email addresses. */
const EMAIL_REGEX = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
/** Regex for validating HTTP/HTTPS URLs. */
const URL_REGEX = /^https?:\/\/\S+$/i;
/** Regex for detecting SharePoint relative file paths (e.g., /sites/...). */
const SP_PATH_REGEX = /^\/sites\/\S+/i;

/**
 * Tests whether a string value looks like an email address.
 * @param str - String to test.
 * @returns True if the string matches email format.
 */
function isEmail(str: string): boolean {
  return EMAIL_REGEX.test(str.trim());
}

/**
 * Tests whether a string value looks like a URL or SharePoint file path.
 * @param str - String to test.
 * @returns True if the string is a URL or SP path.
 */
function isUrl(str: string): boolean {
  const trimmed = str.trim();
  return URL_REGEX.test(trimmed) || SP_PATH_REGEX.test(trimmed);
}

/** Inline style for auto-detected hyperlinks. */
const linkStyle: React.CSSProperties = {
  color: '#0078d4',
  textDecoration: 'none',
};

// ── Compact Badge Component ────────────────────────────────────────────────────

/**
 * Renders a field value as a compact, single-line badge for use inside event cards.
 * Handles special rendering for images (thumbnail), Notes (truncated text block),
 * URLs (clickable link), emails (mailto link), and generic text values.
 *
 * @param props - Field metadata and raw value.
 * @returns A compact badge element appropriate for the field type.
 */
const CompactBadge: React.FC<{ field: IFieldInfo; value: unknown }> = ({ field, value }) => {
  const fieldType = field.fieldType;

  // Image fields — render as a small thumbnail in the card
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

  // Note / multi-line text — render as a short text block with line clamping
  if (fieldType === 'Note') {
    const text = typeof value === 'string' && isHtml(value)
      ? stripHtml(value).substring(0, 120)
      : String(value).substring(0, 120);
    if (!text) return null;
    return (
      <div style={{
        fontSize: 12,
        color: '#605e5c',
        lineHeight: '1.4',
        display: '-webkit-box',
        WebkitLineClamp: 2,
        WebkitBoxOrient: 'vertical' as React.CSSProperties['WebkitBoxOrient'],
        overflow: 'hidden',
      }}>
        <strong>{field.displayName}:</strong> {text}
      </div>
    );
  }

  // URL fields — render as a clickable link, stopPropagation prevents card click
  if (fieldType === 'URL' && typeof value === 'object' && value !== null) {
    const urlObj = value as Record<string, unknown>;
    const url = urlObj.Url as string;
    const desc = (urlObj.Description as string) || url;
    if (url) {
      return (
        <span style={badgeStyle}>
          <strong>{field.displayName}:</strong>
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

  // Auto-link: email addresses get a mailto: link
  if (strVal && isEmail(strVal)) {
    return (
      <span style={badgeStyle}>
        <strong>{field.displayName}:</strong>
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

  // Auto-link: HTTP URLs and SharePoint paths get a standard hyperlink
  if (strVal && isUrl(strVal)) {
    return (
      <span style={badgeStyle}>
        <strong>{field.displayName}:</strong>
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

  // Default: render as labeled text with tooltip for overflow
  return (
    <span
      style={badgeStyle}
      title={field.displayName + ': ' + display}
    >
      <strong>{field.displayName}:</strong> {display}
    </span>
  );
};

// ── Detail Row Component ───────────────────────────────────────────────────────

/**
 * Renders a field value as a full-width labeled row for the event detail panel.
 * Provides richer rendering than CompactBadge: full-size images, rendered HTML
 * for Note fields (via dangerouslySetInnerHTML), themed Choice/MultiChoice pills,
 * clickable URLs, and auto-linked emails and paths.
 *
 * @param props - Field metadata and raw value.
 * @returns A detail row element appropriate for the field type.
 */
const DetailRow: React.FC<{ field: IFieldInfo; value: unknown }> = ({ field, value }) => {
  const fieldType = field.fieldType;

  // Image fields — render full-width with cover fit
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

  // URL field — clickable link with description fallback
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

  // Rich text (Note) — render HTML content directly for full fidelity in the panel
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

  // Boolean / AllDayEvent — display as "Yes" or "No"
  if (fieldType === 'Boolean' || fieldType === 'AllDayEvent') {
    return (
      <div style={detailRowStyle}>
        <div style={detailLabelStyle}>{field.displayName}</div>
        <div style={{ fontSize: 13 }}>{value ? 'Yes' : 'No'}</div>
      </div>
    );
  }

  // MultiChoice — render each selected value as a themed pill/chip
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

  // Choice — single themed pill with slightly larger styling
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

  // Auto-link: email addresses
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

  // Auto-link: HTTP URLs and SharePoint file paths
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

  // Default: plain text value
  return (
    <div style={detailRowStyle}>
      <div style={detailLabelStyle}>{field.displayName}</div>
      <div style={{ fontSize: 13 }}>{display}</div>
    </div>
  );
};

// ── Main FieldBadge Component ──────────────────────────────────────────────────

/**
 * Polymorphic field renderer that delegates to either CompactBadge (for card views)
 * or DetailRow (for the detail panel) based on the `detailed` prop. Returns null
 * for empty values to avoid rendering blank badges.
 *
 * @param props - Field metadata, raw value, and display mode flag.
 * @returns A CompactBadge or DetailRow element, or null if the value is empty.
 */
const FieldBadge: React.FC<IFieldBadgeProps> = ({ field, value, detailed }) => {
  if (isEmptyValue(value, field.fieldType)) return null;

  if (detailed) {
    return <DetailRow field={field} value={value} />;
  }
  return <CompactBadge field={field} value={value} />;
};

export default FieldBadge;
