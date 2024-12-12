import JSZip from 'jszip';
import { parseStringPromise } from 'xml2js';
import { validateXMLString } from '../xml/validation';
import { createXMLBuilder, buildSafeXML } from '../xml/builder';

const PRESENTATION_NAMESPACES = {
  'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
  'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  'xmlns:p': 'http://schemas.openxmlformats.org/presentationml/2006/main'
};

export async function updatePresentationXML(
  zip: JSZip, 
  modifyCallback: (presentationObj: any) => void
): Promise<void> {
  try {
    const presentationXml = await zip.file('ppt/presentation.xml')?.async('string');
    if (!presentationXml) {
      throw new Error('Presentation XML not found');
    }

    // Validate XML before processing
    await validateXMLString(presentationXml);

    const presentationObj = await parseStringPromise(presentationXml, {
      explicitArray: false,
      mergeAttrs: true,
      xmlns: true
    });

    // Deep clone to avoid modifying the original object
    const modifiedObj = JSON.parse(JSON.stringify(presentationObj));
    
    // Apply modifications
    modifyCallback(modifiedObj);

    // Ensure namespace declarations
    if (!modifiedObj['p:presentation'].$) {
      modifiedObj['p:presentation'].$ = {};
    }
    Object.assign(modifiedObj['p:presentation'].$, PRESENTATION_NAMESPACES);

    // Build XML with proper namespaces
    const builder = createXMLBuilder({
      rootName: 'p:presentation',
      namespaces: PRESENTATION_NAMESPACES
    });

    const updatedXml = await buildSafeXML(builder, modifiedObj);
    
    // Final validation before saving
    await validateXMLString(updatedXml);
    
    zip.file('ppt/presentation.xml', updatedXml);
  } catch (error) {
    console.error('Failed to update presentation XML:', error);
    throw new Error(
      'Failed to process presentation: ' + 
      (error instanceof Error ? error.message : 'Unknown error')
    );
  }
}