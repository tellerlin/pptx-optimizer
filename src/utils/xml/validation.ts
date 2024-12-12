import { parseStringPromise } from 'xml2js';
import { XML_VALIDATION_RULES } from './constants';

export interface XMLValidationError extends Error {
  code: 'INVALID_CHARACTER' | 'MALFORMED_XML' | 'MISSING_NAMESPACE' | 'INVALID_ELEMENT_NAME' | 'UNKNOWN';
  location?: string;
  details?: string;
}


export function sanitizeXMLString(xml: string): string {  
  // Remove BOM and normalize line endings  
  let sanitized = xml.replace(/^\uFEFF/, '').replace(/\r\n?/g, '\n');  
  
  // Remove invalid XML characters  
  sanitized = sanitized.replace(XML_VALIDATION_RULES.invalidChars, '');  
  
  // Ensure XML declaration exists  
  if (!sanitized.trim().startsWith('<?xml')) {  
    sanitized = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + sanitized;  
  }  
  
  // Normalize whitespace in tags  
  sanitized = sanitized.replace(/\s+>/g, '>');  
  
  return sanitized;  
}

function createValidationError(
  code: XMLValidationError['code'], 
  message: string,
  location?: string
): XMLValidationError {
  const error = new Error(message) as XMLValidationError;
  error.code = code;
  error.location = location;
  error.name = 'XMLValidationError';
  return error;
}

export async function validateXMLString(xml: string | null | undefined): Promise<void> {  
  try {  
    // 严格的类型检查  
    if (!xml || typeof xml !== 'string') {  
      throw createValidationError(  
        'MALFORMED_XML',   
        'XML input must be a non-empty string'  
      );  
    }  

    // 移除 BOM 字符并去除首尾空白  
    const cleanXml = xml.replace(/^\uFEFF/, '').trim();  
    
    // 检查 XML 是否为空  
    if (!cleanXml) {  
      throw createValidationError(  
        'MALFORMED_XML',   
        'XML content is empty after cleaning'  
      );  
    }  

    // 安全的字符串处理  
    let safeXml = cleanXml.replace(/[^\x09\x0A\x0D\x20-\uD7FF\uE000-\uFFFD\u10000-\u10FFFF]/g, '');  

    // 使用更安全的字符串方法  
    if (!safeXml.match(/<\?xml.*\?>/)) {  
      // 如果没有 XML 声明，添加默认声明  
      safeXml = '<?xml version="1.0" encoding="UTF-8"?>' + safeXml;  
    }  

    // 使用 try-catch 包裹解析逻辑  
    try {  
      await parseStringPromise(safeXml, {  
        strict: true,  
        async: true,  
        explicitChildren: false,  
        preserveChildrenOrder: false  
      });  
    } catch (parseError) {  
      throw createValidationError(  
        'MALFORMED_XML',   
        `XML parsing failed: ${parseError instanceof Error ? parseError.message : 'Unknown error'}`,  
        parseError instanceof Error ? parseError.stack : undefined  
      );  
    }  
  } catch (error) {  
    // 更详细的错误处理  
    if (error instanceof Error && 'code' in error) {  
      throw error;  
    }  
    
    throw createValidationError(  
      'UNKNOWN',   
      error instanceof Error ? error.message : 'Unexpected XML validation error'  
    );  
  }  
}