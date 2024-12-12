import JSZip from 'jszip';
import { parseStringPromise } from 'xml2js';
import { Builder } from 'xml2js';

// Types
export interface OptimizationOptions {
  removeHiddenSlides?: boolean;
  compressImages?: ImageOptimizationOptions;
  removeUnusedMedia?: boolean;
  usePlaceholders?: boolean;
}

export interface ImageOptimizationOptions {
  maxWidth?: number;
  maxHeight?: number;
  quality?: number;
}

export interface MediaFile {
  path: string;
  type: string;
  size: number;
}

export interface PlaceholderOptions {
  preserveOriginalSize?: boolean;
  format?: 'png' | 'jpg';
}

export interface XMLValidationError extends Error {
  code: 'INVALID_CHARACTER' | 'MALFORMED_XML' | 'MISSING_NAMESPACE' | 'INVALID_ELEMENT_NAME' | 'UNKNOWN';
  location?: string;
  details?: string;
}

// Constants
const PRESENTATION_NAMESPACES = {
  'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
  'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  'xmlns:p': 'http://schemas.openxmlformats.org/presentationml/2006/main'
};

const RELATIONSHIP_NAMESPACES = {
  'xmlns': 'http://schemas.openxmlformats.org/package/2006/relationships'
};

const XML_NAMESPACES = {
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

const XML_VALIDATION_RULES = {
  elementName: /^[a-zA-Z_][\w\-.:]*$/,
  attributeName: /^[a-zA-Z_][\w\-.:]*$/,
  invalidChars: /[^\x09\x0A\x0D\x20-\uD7FF\uE000-\uFFFD\u{10000}-\u{10FFFF}]/u
};

const CONTENT_TYPES_NAMESPACES = {
  'xmlns': 'http://schemas.openxmlformats.org/package/2006/content-types'
};

// XML Processing Functions
function createXMLBuilder(options: any): Builder {
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

async function buildSafeXML(builder: Builder, obj: any): Promise<string> {  
  try {  
    const processedObj = preprocessXMLObject(obj);  
    const xml = builder.buildObject(processedObj);  
    const sanitizedXml = sanitizeXMLString(xml);  
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
      if (value === undefined || value === null) continue;  
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

// Media Processing Functions
async function createPlaceholderFile(originalSize: number, extension: string): Promise<ArrayBuffer> {
  switch (extension.toLowerCase()) {
    case 'png':
      return createMinimalPNG();
    case 'jpg':
    case 'jpeg':
      return createMinimalJPEG();
    default:
      return createMinimalPNG();
  }
}

function createMinimalPNG(): ArrayBuffer {
  const minimalPNG = new Uint8Array([
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
    0x00, 0x00, 0x00, 0x0D,
    0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01,
    0x00, 0x00, 0x00, 0x01,
    0x08,
    0x06,
    0x00,
    0x00,
    0x00,
    0x1F, 0x15, 0xC4, 0x89,
    0x00, 0x00, 0x00, 0x0A,
    0x49, 0x44, 0x41, 0x54,
    0x78, 0x9C, 0x63, 0x00, 0x00, 0x00, 0x02, 0x00, 0x01,
    0xE5, 0x27, 0x0E, 0x89,
    0x00, 0x00, 0x00, 0x00,
    0x49, 0x45, 0x4E, 0x44,
    0xAE, 0x42, 0x60, 0x82
  ]);
  return minimalPNG.buffer;
}

function createMinimalJPEG(): ArrayBuffer {
  const minimalJPEG = new Uint8Array([
    0xFF, 0xD8,
    0xFF, 0xE0, 0x00, 0x10,
    0x4A, 0x46, 0x49, 0x46, 0x00,
    0x01, 0x01,
    0x00,
    0x00, 0x01, 0x00, 0x01,
    0x00, 0x00,
    0xFF, 0xDB, 0x00, 0x43, 0x00,
    ...Array(64).fill(1),
    0xFF, 0xC0, 0x00, 0x0B,
    0x08, 0x00, 0x01, 0x00, 0x01,
    0x01, 0x00,
    0xFF, 0xDA, 0x00, 0x08,
    0x01, 0x00, 0x00, 0x3F, 0x00,
    0xFF, 0xD9
  ]);
  return minimalJPEG.buffer;
}

async function compressImage(
  imageData: ArrayBuffer,
  options: ImageOptimizationOptions = {}
): Promise<ArrayBuffer> {
  const blob = new Blob([imageData]);
  const bitmap = await createImageBitmap(blob);
  const { width, height } = calculateOptimalDimensions(bitmap.width, bitmap.height);
  
  const canvas = new OffscreenCanvas(width, height);
  const ctx = canvas.getContext('2d');
  if (!ctx) {
    throw new Error('Failed to get canvas context');
  }
  
  ctx.drawImage(bitmap, 0, 0, width, height);
  const imageDataForAnalysis = ctx.getImageData(0, 0, width, height);
  const hasAlpha = checkAlphaChannel(imageDataForAnalysis);

  if (hasAlpha) {
    const blob = await canvas.convertToBlob({ 
      type: 'image/webp', 
      quality: 0.7 
    });
    return await blob.arrayBuffer();
  } else {
    const [webpBlob, jpegBlob] = await Promise.all([
      canvas.convertToBlob({ type: 'image/webp', quality: 0.7 }),
      canvas.convertToBlob({ type: 'image/jpeg', quality: 0.7 })
    ]);

    const [webpBuffer, jpegBuffer] = await Promise.all([
      webpBlob.arrayBuffer(),
      jpegBlob.arrayBuffer()
    ]);

    return webpBuffer.byteLength <= jpegBuffer.byteLength ? webpBuffer : jpegBuffer;
  }
}

// Media File Management Functions
async function collectUsedMedia(zip: JSZip): Promise<Set<string>> {
  const usedMediaFiles = new Set<string>();

  try {
    const slideFiles = Object.keys(zip.files)
      .filter(filename => filename.startsWith('ppt/slides/slide') && filename.endsWith('.xml'));

    await Promise.all(slideFiles.map(async (slideFile) => {
      try {
        await processFileMedia(zip, slideFile, usedMediaFiles);
        const slideRelsFile = `ppt/slides/_rels/${slideFile.split('/').pop()}.rels`;
        const slideRelsXml = await zip.file(slideRelsFile)?.async('string');
        if (!slideRelsXml) return;

        const slideRelsObj = await parseStringPromise(slideRelsXml);
        const relationships = slideRelsObj.Relationships?.Relationship || [];
        const layoutRel = relationships.find((rel: any) => 
          rel.$.Type === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'
        );

        if (layoutRel) {
          const layoutPath = `ppt/${layoutRel.$.Target.replace('../', '')}`;
          await processFileMedia(zip, layoutPath, usedMediaFiles);
        }
      } catch (error) {
        console.warn(`Error processing slide ${slideFile}:`, error);
      }
    }));
  } catch (error) {
    console.error('Error collecting used media:', error);
    throw new Error('Failed to collect used media files');
  }

  return usedMediaFiles;
}

function findAllMediaFiles(zip: JSZip): string[] {
  return Object.keys(zip.files)
    .filter(filename => 
      filename.startsWith('ppt/media/') && 
      /\.(png|jpg|jpeg|gif|bmp|webp|wmf|emf|svg)$/i.test(filename)
    );
}

function findUnusedMedia(
  allMediaFiles: string[], 
  usedMediaFiles: Set<string>
): string[] {
  return allMediaFiles.filter(mediaFile => !usedMediaFiles.has(mediaFile));
}

// PPTX Processing Functions
async function updatePresentationXML(
  zip: JSZip, 
  modifyCallback: (presentationObj: any) => void
): Promise<void> {
  try {
    const presentationXml = await zip.file('ppt/presentation.xml')?.async('string');
    if (!presentationXml) {
      throw new Error('Presentation XML not found');
    }

    await validateXMLString(presentationXml);

    const presentationObj = await parseStringPromise(presentationXml, {
      explicitArray: false,
      mergeAttrs: true,
      xmlns: true
    });

    const modifiedObj = JSON.parse(JSON.stringify(presentationObj));
    modifyCallback(modifiedObj);

    if (!modifiedObj['p:presentation'].$) {
      modifiedObj['p:presentation'].$ = {};
    }
    Object.assign(modifiedObj['p:presentation'].$, PRESENTATION_NAMESPACES);

    const builder = createXMLBuilder({
      rootName: 'p:presentation',
      namespaces: PRESENTATION_NAMESPACES
    });

    const updatedXml = await buildSafeXML(builder, modifiedObj);
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

async function updateRelationships(zip: JSZip): Promise<void> {
  try {
    const relsPath = 'ppt/_rels/presentation.xml.rels';
    const relsFile = zip.file(relsPath);
    
    if (!relsFile) {
      console.warn('Relationships file not found:', relsPath);
      return;
    }

    let relsXml = await relsFile.async('string');
    await validateXMLString(relsXml);

    const relsObj = await parseStringPromise(relsXml, {
      explicitArray: false,
      mergeAttrs: true,
      xmlns: true
    });

    if (!relsObj || typeof relsObj !== 'object') {
      throw new Error('Failed to parse relationships XML');
    }

    const builder = createXMLBuilder({
      rootName: 'Relationships',
      namespaces: RELATIONSHIP_NAMESPACES
    });

    const updatedXml = await buildSafeXML(builder, relsObj);
    await validateXMLString(updatedXml);
    zip.file(relsPath, updatedXml);

  } catch (error) {
    console.error('Failed to update relationships:', error);
    throw new Error(
      'Failed to update relationships: ' + 
      (error instanceof Error ? error.message : 'Unknown error')
    );
  }
}

// Main Optimization Function
export async function optimizePPTX(
  file: File, 
  options: OptimizationOptions = {}
): Promise<Blob> {
  const startTime = performance.now();
  const originalSize = file.size;

  try {
    const zip = await JSZip.loadAsync(file);
    const usedMediaFiles = await collectUsedMedia(zip);
    const allMediaFiles = findAllMediaFiles(zip);
    
    if (options.compressImages || options.usePlaceholders) {
      const mediaFolder = zip.folder('ppt/media');
      if (mediaFolder) {
        const imageFiles = allMediaFiles.filter(filename => 
          /\.(png|jpg|jpeg|gif|bmp|webp)$/i.test(filename)
        );

        await Promise.all(imageFiles.map(async (filename) => {
          try {
            const file = zip.file(filename);
            if (!file) return;

            const imageData = await file.async('arraybuffer');
            const extension = filename.split('.').pop()?.toLowerCase() || 'png';
            
            if (usedMediaFiles.has(filename)) {
              if (options.usePlaceholders) {
                const placeholderData = await createPlaceholderFile(
                  file.uncompressedSize,
                  extension
                );
                zip.file(filename, placeholderData);
              } else if (options.compressImages) {
                const optimizedImage = await compressImage(imageData, options.compressImages);
                zip.file(filename, optimizedImage);
              }
            } else {
              const placeholderData = await createPlaceholderFile(
                file.uncompressedSize,
                extension
              );
              zip.file(filename, placeholderData);
            }
          } catch (error) {
            console.warn(`Failed to process image ${filename}:`, error);
          }
        }));
      }
    }

    if (options.removeUnusedMedia) {
      const unusedMediaFiles = findUnusedMedia(allMediaFiles, usedMediaFiles);
      await Promise.all(unusedMediaFiles.map(async (mediaPath) => {
        const file = zip.file(mediaPath);
        if (!file) return;
        
        const extension = mediaPath.split('.').pop()?.toLowerCase() || 'png';
        const placeholderData = await createPlaceholderFile(
          file.uncompressedSize,
          extension
        );
        zip.file(mediaPath, placeholderData);
      }));
    }

    const optimizedBlob = await zip.generateAsync({
      type: 'blob',
      compression: 'DEFLATE',
      compressionOptions: { level: 9 }
    });

    const endTime = performance.now();
    console.log(`
      Optimization completed:
      - Processing time: ${(endTime - startTime).toFixed(2)}ms
      - Original size: ${(originalSize / 1024 / 1024).toFixed(2)}MB
      - Optimized size: ${(optimizedBlob.size / 1024 / 1024).toFixed(2)}MB
      - Size reduction: ${((1 - optimizedBlob.size / originalSize) * 100).toFixed(2)}%
    `);

    return optimizedBlob;
  } catch (error) {
    console.error('PPTX optimization failed:', error);
    throw new Error(
      'Failed to optimize PPTX: ' + 
      (error instanceof Error ? error.message : 'Unknown error')
    );
  }
}
