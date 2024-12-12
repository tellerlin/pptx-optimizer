export const XML_NAMESPACES = {
  presentation: {
    'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'xmlns:p': 'http://schemas.openxmlformats.org/presentationml/2006/main'
  },
  relationships: {
    'xmlns': 'http://schemas.openxmlformats.org/package/2006/relationships'
  },
  contentTypes: {
    'xmlns': 'http://schemas.openxmlformats.org/package/2006/content-types'
  }
};

export const XML_VALIDATION_RULES = {
  elementName: /^[a-zA-Z_][\w\-.:]*$/,
  attributeName: /^[a-zA-Z_][\w\-.:]*$/,
  invalidChars: /[^\x09\x0A\x0D\x20-\uD7FF\uE000-\uFFFD\u{10000}-\u{10FFFF}]/u
};