/**
 * FieldPulse Pricebook 2.0 — Shared Dialog Icons
 *
 * Inline Material-style SVG icons + shared CSS for every HtmlService modal,
 * so all Pricebook Tools dialogs share one visual language. No external
 * font/CDN — icons are inline SVG (no load flash, works offline).
 *
 * Usage in a modal builder:
 *   '<style>' + DIALOG_ICON_CSS + '.summary{...}' + ...
 *   dialogSummary_('ok', escapeHtml_(text))                 // banner w/ state icon
 *   '<div class="group-header">' + dialogIcon_('layers') + 'Title</div>'
 *
 * Banner states: 'ok' | 'warn' | 'err' | 'info'.
 * Add new icon paths to DIALOG_ICON_PATHS as dialogs are converted.
 */

const DIALOG_ICON_CSS =
  '.dico{width:18px;height:18px;flex:0 0 auto;fill:currentColor;vertical-align:-4px;}' +
  '.summary .dico{width:20px;height:20px;}' +
  '.group-header .dico{width:16px;height:16px;margin-right:6px;vertical-align:-3px;}';

// Material Design icon paths (viewBox 0 0 24 24).
const DIALOG_ICON_PATHS = {
  check_circle:  'M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm-2 15l-5-5 1.41-1.41L10 14.17l7.59-7.59L19 8l-9 9z',
  warning:       'M1 21h22L12 2 1 21zm12-3h-2v-2h2v2zm0-4h-2v-4h2v4z',
  error:         'M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-2h2v2zm0-4h-2V7h2v6z',
  info:          'M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-6h2v6zm0-8h-2V7h2v2z',
  layers:        'M11.99 18.54l-7.37-5.73L3 14.07l9 7 9-7-1.63-1.27-7.38 5.74zM12 16l7.36-5.73L21 9l-9-7-9 7 1.63 1.27L12 16z',
  visibility_off:'M12 7c2.76 0 5 2.24 5 5 0 .65-.13 1.26-.36 1.83l2.92 2.92c1.51-1.26 2.7-2.89 3.43-4.75-1.73-4.39-6-7.5-11-7.5-1.4 0-2.74.25-3.98.7l2.16 2.16C10.74 7.13 11.35 7 12 7zM2 4.27l2.28 2.28.46.46C3.08 8.3 1.78 10.02 1 12c1.73 4.39 6 7.5 11 7.5 1.55 0 3.03-.3 4.38-.84l.42.42L19.73 22 21 20.73 3.27 3 2 4.27zm5.53 5.53l1.55 1.55c-.05.21-.08.43-.08.65 0 1.66 1.34 3 3 3 .22 0 .44-.03.65-.08l1.55 1.55c-.67.33-1.41.53-2.2.53-2.76 0-5-2.24-5-5 0-.79.2-1.53.53-2.2zm4.31-.78l3.15 3.15.02-.16c0-1.66-1.34-3-3-3l-.17.01z'
};

function dialogIcon_(name, extraClass) {
  const path = DIALOG_ICON_PATHS[name] || DIALOG_ICON_PATHS.info;
  const cls = 'dico' + (extraClass ? ' ' + extraClass : '');
  return '<svg class="' + cls + '" viewBox="0 0 24 24" aria-hidden="true"><path d="' + path + '"/></svg>';
}

// Standard summary banner with a state icon. text should be pre-escaped.
function dialogSummary_(state, text) {
  const iconByState = { ok: 'check_circle', warn: 'warning', err: 'error', info: 'info' };
  const icon = dialogIcon_(iconByState[state] || 'info');
  return '<div class="summary ' + state + '">' + icon + '<span>' + text + '</span></div>';
}
