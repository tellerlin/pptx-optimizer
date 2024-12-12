import JSZip from 'jszip';
import { parseStringPromise, Builder } from 'xml2js';

interface OptimizationOptions {
  removeHiddenSlides?: boolean;
  compressImages?: {
    maxWidth?: number;
    maxHeight?: number;
    quality?: number;
  };
  removeUnusedMedia?: boolean;
}

async function compressImage(
  imageBuffer: ArrayBuffer, 
  options?: {
    maxWidth?: number;
    maxHeight?: number;
    quality?: number;
  }
): Promise<ArrayBuffer> {
  return new Promise((resolve, reject) => {
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');
    const img = new Image();
    
    img.onload = () => {
      const maxWidth = options?.maxWidth || 1366;
      const maxHeight = options?.maxHeight || 768;
      
      let width = img.width;
      let height = img.height;
      
      if (width > maxWidth || height > maxHeight) {
        const ratio = Math.min(maxWidth / width, maxHeight / height);
        width *= ratio;
        height *= ratio;
      }
      
      canvas.width = width;
      canvas.height = height;
      
      ctx?.drawImage(img, 0, 0, width, height);
      
      canvas.toBlob((blob) => {
        const reader = new FileReader();
        reader.onloadend = () => {
          resolve(reader.result as ArrayBuffer);
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(blob!);
      }, 'image/webp', options?.quality || 0.7);
    };
    
    img.onerror = reject;
    img.src = URL.createObjectURL(new Blob([imageBuffer]));
  });
}

async function collectUsedMedia(zip: JSZip): Promise<Set<string>> {
  const usedMediaFiles = new Set<string>();

  const slideFiles = Object.keys(zip.files)
    .filter(filename => filename.startsWith('ppt/slides/slide') && filename.endsWith('.xml'));

  await Promise.all(slideFiles.map(async (slideFile) => {
    const slideXml = await zip.file(slideFile)?.async('string');
    if (slideXml) {
      const mediaMatches = slideXml.match(/r:embed="[^"]*"/g) || [];
      
      await Promise.all(mediaMatches.map(async (match) => {
        const relPath = match.replace(/r:embed="([^"]*)"/g, '$1');
        const fullMediaPath = `ppt/slides/_rels/${slideFile.split('/').pop()}.rels`;
        
        try {
          const relsXml = await zip.file(fullMediaPath)?.async('string');
          if (relsXml) {
            const relsObj = await parseStringPromise(relsXml);
            const relationships = relsObj.Relationships?.Relationship || [];
            
            relationships.forEach((rel: any) => {
              if (rel.$.Id === relPath) {
                const mediaPath = `ppt/media/${rel.$.Target}`;
                usedMediaFiles.add(mediaPath);
              }
            });
          }
        } catch (error) {
          console.warn(`Error processing relationships for ${slideFile}:`, error);
        }
      }));
    }
  }));

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
  return allMediaFiles.filter(mediaFile => 
    !usedMediaFiles.has(mediaFile)
  );
}

async function updateRelationships(zip: JSZip) {
  try {
    const presentationRelsXml = await zip.file('ppt/_rels/presentation.xml.rels')?.async('string');
    if (!presentationRelsXml) return;

    const relsObj = await parseStringPromise(presentationRelsXml, {
      explicitArray: false,
      mergeAttrs: true,
      xmlns: true
    });

    const builder = new Builder({
      renderOpts: { pretty: false },
      xmldec: { 
        version: '1.0', 
        encoding: 'UTF-8',
        standalone: 'yes'
      },
      headless: false,
      rootName: 'Relationships',
      namespaceDef: {
        'xmlns': 'http://schemas.openxmlformats.org/package/2006/relationships'
      }
    });

    const updatedRelsXml = builder.buildObject(relsObj);
    zip.file('ppt/_rels/presentation.xml.rels', updatedRelsXml);
  } catch (error) {
    console.error('Relationship update failed:', error);
  }
}

async function validateXmlFiles(zip: JSZip) {
  const xmlFiles = [
    'ppt/presentation.xml', 
    'ppt/_rels/presentation.xml.rels', 
    '[Content_Types].xml'
  ];

  for (const file of xmlFiles) {
    try {
      const xmlContent = await zip.file(file)?.async('string');
      await parseStringPromise(xmlContent || '');
    } catch (parseError) {
      console.error(`XML parsing error in ${file}:`, parseError);
      throw new Error(`Invalid XML structure in ${file}`);
    }
  }
}

async function rebuildContentTypesFile(zip: JSZip) {
  try {
    const contentTypesXml = await zip.file('[Content_Types].xml')?.async('string');
    if (!contentTypesXml) return;

    const contentTypesObj = await parseStringPromise(contentTypesXml, {
      explicitArray: false,
      mergeAttrs: true
    });

    // 移除不存在的文件引用
    if (contentTypesObj.Types && contentTypesObj.Types.Override) {
      contentTypesObj.Types.Override = contentTypesObj.Types.Override.filter((override: any) => {
        const partName = override.$.PartName.replace(/^\//, '');
        return zip.files[partName] !== undefined;
      });
    }

    const builder = new Builder({
      renderOpts: { pretty: false },
      xmldec: { 
        version: '1.0', 
        encoding: 'UTF-8',
        standalone: 'yes'
      }
    });

    const updatedContentTypesXml = builder.buildObject(contentTypesObj);
    zip.file('[Content_Types].xml', updatedContentTypesXml);
  } catch (error) {
    console.error('Rebuilding Content Types file failed:', error);
  }
}

export async function optimizePPTX(
  file: File, 
  options: OptimizationOptions = {}
): Promise<Blob> {
  const startTime = performance.now();
  const originalSize = file.size;

  try {
    const zip = await JSZip.loadAsync(file);

    // 验证 XML 文件
    await validateXmlFiles(zip);

    // 默认配置
    const mergedOptions: OptimizationOptions = {
      removeHiddenSlides: true,
      compressImages: {
        maxWidth: 1366,
        maxHeight: 768,
        quality: 0.7
      },
      removeUnusedMedia: true,
      ...options
    };

    // 移除隐藏的幻灯片
    if (mergedOptions.removeHiddenSlides) {
      await safeUpdatePresentationXml(zip, (presentationObj) => {
        const sldIdLst = presentationObj['p:presentation']['p:sldIdLst'];
        if (sldIdLst && sldIdLst['p:sldId']) {
          // Ensure we're working with an array
          const slides = Array.isArray(sldIdLst['p:sldId']) 
            ? sldIdLst['p:sldId'] 
            : [sldIdLst['p:sldId']];
          
          // Filter hidden slides
          presentationObj['p:presentation']['p:sldIdLst']['p:sldId'] = 
            slides.filter((slide: any) => {
              // 添加安全检查
              if (!slide || !slide.$) return true; // 如果slide或slide.$不存在，保留该幻灯片
              return !slide.$.hidden;
            });
        }
      });
    }

    // 图片压缩逻辑
    if (mergedOptions.compressImages) {
      const mediaFolder = zip.folder('ppt/media');
      if (mediaFolder) {
        const imageFiles = Object.keys(mediaFolder.files).filter(filename => 
          /\.(png|jpg|jpeg|gif|bmp|webp)$/i.test(filename)
        );

        const imageProcessTasks = imageFiles.map(async (filename) => {
          try {
            const file = mediaFolder.files[filename];
            const imageData = await file.async('arraybuffer');

            const optimizedImage = await compressImage(imageData, {
              maxWidth: mergedOptions.compressImages?.maxWidth,
              maxHeight: mergedOptions.compressImages?.maxHeight,
              quality: mergedOptions.compressImages?.quality
            });

            zip.file(`ppt/media/${filename}`, optimizedImage);
          } catch (error) {
            console.warn(`Failed to compress image ${filename}:`, error);
          }
        });

        await Promise.allSettled(imageProcessTasks);
      }
    }

    // 移除未使用的媒体文件
    if (mergedOptions.removeUnusedMedia) {
      try {
        const usedMediaFiles = await collectUsedMedia(zip);
        const allMediaFiles = findAllMediaFiles(zip);
        const unusedMediaFiles = findUnusedMedia(allMediaFiles, usedMediaFiles);

        unusedMediaFiles.forEach(mediaPath => {
          delete zip.files[mediaPath];
        });

        console.log(`Removed ${unusedMediaFiles.length} unused media files`);
      } catch (error) {
        console.error('Failed to remove unused media:', error);
      }
    }

    // 重建内容类型文件
    await rebuildContentTypesFile(zip);

    // 生成优化后的文件
    const optimizedBlob = await zip.generateAsync({
      type: 'blob',
      compression: 'DEFLATE',
      compressionOptions: {
        level: 9
      }
    });

    const endTime = performance.now();
    
    console.log(`
      Optimization Details:
      - Total time: ${(endTime - startTime).toFixed(2)} ms
      - Original file size: ${originalSize} bytes
      - Optimized file size: ${optimizedBlob.size} bytes
      - Size reduction: ${((1 - optimizedBlob.size / originalSize) * 100).toFixed(2)}%
    `);

    return optimizedBlob;
  } catch (error) {
    console.error('Optimization failed:', error);
    throw error;
  }
}

async function safeUpdatePresentationXml(zip: JSZip, modifyCallback: (presentationObj: any) => void) {
  try {
    const presentationXml = await zip.file('ppt/presentation.xml')?.async('string');
    if (!presentationXml) throw new Error('Presentation XML not found');

    const parseOptions = {
      explicitArray: false,
      mergeAttrs: true,
      xmlns: true,
      preserveWhitespace: true
    };

    let presentationObj;
    try {
      presentationObj = await parseStringPromise(presentationXml, parseOptions);
    } catch (parseError) {
      console.error('Failed to parse presentation XML:', parseError);
      throw parseError;
    }

    const modifiedObj = JSON.parse(JSON.stringify(presentationObj));
    modifyCallback(modifiedObj);

    const builder = new Builder({
      renderOpts: { pretty: false },
      headless: false,
      rootName: 'p:presentation',
      namespaceDef: {
        'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'xmlns:p': 'http://schemas.openxmlformats.org/presentationml/2006/main'
      },
      xmldec: { version: '1.0', encoding: 'UTF-8', standalone: true }
    });

    // Ensure the root element has all required namespaces
    if (!modifiedObj['p:presentation'].$) {
      modifiedObj['p:presentation'].$ = {};
    }
    Object.assign(modifiedObj['p:presentation'].$, {
      'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
      'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'xmlns:p': 'http://schemas.openxmlformats.org/presentationml/2006/main'
    });

    const updatedXml = builder.buildObject(modifiedObj);

    if (!updatedXml) throw new Error('Failed to build updated presentation XML');

    zip.file('ppt/presentation.xml', updatedXml);

    await updateRelationships(zip);
  } catch (error) {
    console.error('Safe XML update failed:', error);
    throw error;
  }
}

export {
  compressImage,
  collectUsedMedia,
  findAllMediaFiles,
  findUnusedMedia,
  rebuildContentTypesFile
};
