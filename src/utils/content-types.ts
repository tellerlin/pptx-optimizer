import JSZip from 'jszip';
import { parseStringPromise } from 'xml2js';
import { createXMLBuilder, buildSafeXML } from './xml/builder';
import { validateXMLString } from './xml/validation';

const CONTENT_TYPES_NAMESPACES = {
  'xmlns': 'http://schemas.openxmlformats.org/package/2006/content-types'
};

export async function rebuildContentTypesFile(zip: JSZip): Promise<void> {
  try {
    const contentTypesXml = await zip.file('[Content_Types].xml')?.async('string');
    if (!contentTypesXml) return;

    await validateXMLString(contentTypesXml);

    const contentTypesObj = await parseStringPromise(contentTypesXml, {
      explicitArray: false,
      mergeAttrs: true
    });

    // Remove references to non-existent files
    if (contentTypesObj.Types?.Override) {
      contentTypesObj.Types.Override = Array.isArray(contentTypesObj.Types.Override)
        ? contentTypesObj.Types.Override
        : [contentTypesObj.Types.Override];

      contentTypesObj.Types.Override = contentTypesObj.Types.Override
        .filter((override: any) => {
          const partName = override.$.PartName.replace(/^\//, '');
          return zip.files[partName] !== undefined;
        });
    }

    const builder = createXMLBuilder({
      rootName: 'Types',
      namespaces: CONTENT_TYPES_NAMESPACES
    });

    const updatedXml = buildSafeXML(builder, contentTypesObj);
    await validateXMLString(updatedXml);
    
    zip.file('[Content_Types].xml', updatedXml);
  } catch (error) {
    console.error('Failed to rebuild content types:', error);
    throw new Error(
      'Failed to rebuild content types: ' + 
      (error instanceof Error ? error.message : 'Unknown error')
    );
  }
}