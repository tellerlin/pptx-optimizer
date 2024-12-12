import JSZip from 'jszip';
import { parseStringPromise } from 'xml2js';
import { createXMLBuilder, buildSafeXML } from '../xml/builder';
import { validateXMLString } from '../xml/validation';


const RELATIONSHIP_NAMESPACES = {
  'xmlns': 'http://schemas.openxmlformats.org/package/2006/relationships'
};


export async function updateRelationships(zip: JSZip): Promise<void> {
  try {
    const relsPath = 'ppt/_rels/presentation.xml.rels';
    const relsFile = zip.file(relsPath);
    
    if (!relsFile) {
      console.warn('Relationships file not found:', relsPath);
      return;
    }


    let relsXml: string;
    try {
      relsXml = await relsFile.async('string');
    } catch (asyncError) {
      console.error('Failed to read relationships file:', asyncError);
      throw new Error('Could not read relationships file');
    }


    // 验证 XML 字符串
    await validateXMLString(relsXml);


    // 解析 XML
    const relsObj = await parseStringPromise(relsXml, {
      explicitArray: false,
      mergeAttrs: true,
      xmlns: true
    });


    if (!relsObj || typeof relsObj !== 'object') {
      throw new Error('Failed to parse relationships XML');
    }


    // 创建 XML 构建器
    const builder = createXMLBuilder({
      rootName: 'Relationships',
      namespaces: RELATIONSHIP_NAMESPACES
    });


    // 构建安全的 XML
    const updatedXml = buildSafeXML(builder, relsObj);
    
    if (!updatedXml || typeof updatedXml !== 'string') {
      throw new Error('Failed to generate updated XML');
    }


    // 再次验证更新后的 XML
    await validateXMLString(updatedXml);


    // 更新 ZIP 文件中的关系文件
    zip.file(relsPath, updatedXml);


  } catch (error) {
    console.error('Failed to update relationships:', error);
    throw new Error(
      'Failed to update relationships: ' + 
      (error instanceof Error ? error.message : 'Unknown error')
    );
  }
}