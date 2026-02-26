// Security utilities for sanitizing external data before DOM rendering

/**
 * Escapes HTML special characters to prevent XSS when inserting into innerHTML.
 * Use this for ALL data from external sources (Google Sheets, Firestore, URL params).
 */
function escapeHtml(str) {
  if (!str) return '';
  const div = document.createElement('div');
  div.textContent = String(str);
  return div.innerHTML;
}

/**
 * Validates and sanitizes URLs. Only allows https://, http://, and relative paths.
 * Blocks javascript:, data:, vbscript: and other dangerous protocols.
 */
function sanitizeUrl(url) {
  if (!url) return '';
  const u = String(url).trim();
  if (/^(https?:\/\/|\/|images\/|\.\/)/i.test(u)) return u;
  return '';
}
