// Logos.gs
// Loads institution logos by calling getLogosData_() defined in LogosData.gs.
// No HtmlService / file reading — guaranteed to work regardless of file size.

var _LOGOS_CACHE = null;

function getLogos_() {
  if (_LOGOS_CACHE) return _LOGOS_CACHE;
  _LOGOS_CACHE = getLogosData_();   // defined in LogosData.gs
  return _LOGOS_CACHE;
}

// code may have trailing whitespace when read from the spreadsheet, so trim it
function logoFor_(code) {
  return getLogos_()[String(code || '').trim()] || '';
}
