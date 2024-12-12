import { Builder } from 'xml2js';
import { validateXMLString, sanitizeXMLString } from './validation';

export interface XMLBuilderOptions {
  rootName: string;
  namespaces?: Record<string, string>;
  headless?: boolean;
  pretty?: boolean;
}

export function createXMLBuilder(options: XMLBuilderOptions): Builder {
  return new Builder({
    renderOpts: { 
      pretty: options.pretty ?? false,
      indent: '  ',
      newline: '\n'
    },
    headless: options.headless ?? false,
    rootName: options.rootName,
    xmldec: { 
      version: '1.0', 
      encoding: 'UTF-8', 
      standalone: true 
    },
    namespaceDef: options.namespaces
  });
}

export async function buildSafeXML(builder: Builder, obj: any): Promise<string> {  
  try {  
    // Pre-process object to ensure valid element names  
    const processedObj = preprocessXMLObject(obj);  
    console.log('Processed Object:', processedObj);  
    
    // Build XML  
    const xml = builder.buildObject(processedObj);  
    console.log('Generated XML:', xml);  
    
    // Sanitize and validate  
    const sanitizedXml = sanitizeXMLString(xml);  
    console.log('Sanitized XML:', sanitizedXml);  
    await validateXMLString(sanitizedXml);  
    
    return sanitizedXml;  
  } catch (error) {  
    console.error('XML building error:', error);  
    throw new Error(  
      'Failed to build XML: ' +   
      (error instanceof Error ? error.message : 'Unknown error')  
    );  
  }  
}


function preprocessXMLObject(obj: any): any {  
  if (Array.isArray(obj)) {  
    return obj.map(item => preprocessXMLObject(item));  
  }  
  
  if (obj && typeof obj === 'object') {  
    const processed: any = {};  
    for (const [key, value] of Object.entries(obj)) {  
      // Skip empty or null values  
      if (value === undefined || value === null) continue;  
      
      // Ensure valid element names  
      let safeKey = key.replace(/[^\w\-.:]/g, '_');  
      if (!isValidElementName(safeKey)) {  
        safeKey = `item_${safeKey}`;  
      }  
      processed[safeKey] = preprocessXMLObject(value);  
    }  
    return processed;  
  }  
  
  return obj;  
}

