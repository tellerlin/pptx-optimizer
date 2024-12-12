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


function compressImage(
  imageBuffer: ArrayBuffer, 
  options?: {
    maxWidth?: number;
    maxHeight?: number;
    quality?: number;
  }
): Promise<ArrayBuffer> {
  return new Promise((resolve) => {
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
        reader.readAsArrayBuffer(blob!);
      }, 'image/webp', options?.quality || 0.7);
    };
    
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


async function cleanupUnusedMedia(
  zip: JSZip, 
  options: { 
    dryRun?: boolean 
  } = {}
): Promise<string[]> {
  try {
    const usedMediaFiles = await collectUsedMedia(zip);
    const allMediaFiles = findAllMediaFiles(zip);
    const unusedMediaFiles = findUnusedMedia(allMediaFiles, usedMediaFiles);


    if (!options.dryRun) {
      unusedMediaFiles.forEach(mediaPath => {
        delete zip.files[mediaPath];
      });
    }


    return unusedMediaFiles;
  } catch (error) {
    console.error('Media cleanup failed:', error);
    return [];
  }
}


export {
  compressImage,
  collectUsedMedia,
  findAllMediaFiles,
  findUnusedMedia,
  cleanupUnusedMedia
};


export async function optimizePPTX(
  file: File, 
  options: OptimizationOptions = {}
): Promise<Blob> {
  const startTime = performance.now();
  const originalSize = file.size;


  const defaultOptions: OptimizationOptions = {
    removeHiddenSlides: true,
    compressImages: {
      maxWidth: 1366,
      maxHeight: 768,
      quality: 0.7
    },
    removeUnusedMedia: true
  };


  const mergedOptions = { ...defaultOptions, ...options };


  const arrayBuffer = await file.arrayBuffer();
  const zip = new JSZip();
  await zip.loadAsync(arrayBuffer);


  try {
    const presentationXml = await zip.file('ppt/presentation.xml')?.async('string');
    if (presentationXml) {
      const presentationObj = await parseStringPromise(presentationXml, {
        explicitArray: false,
        mergeAttrs: true
      });


      if (mergedOptions.removeHiddenSlides) {
        // 安全地访问 sldIdLst
        const presentation = presentationObj['p:presentation'] || {};
        const sldIdLst = presentation['p:sldIdLst'] || {};
        
        // 确保 sldId 是数组
        let sldIds = sldIdLst['p:sldId'] || [];
        if (!Array.isArray(sldIds)) {
          sldIds = [sldIds];
        }


        // 过滤可见的幻灯片
        const visibleSlides = sldIds.filter((slide: any) => 
          !slide['$']?.hidden
        );


        // 更新幻灯片列表
        if (presentation['p:sldIdLst']) {
          presentation['p:sldIdLst']['p:sldId'] = visibleSlides;
        }


        const builder = new Builder({
          renderOpts: { pretty: true },
          xmldec: { version: '1.0', encoding: 'UTF-8' }
        });
        const updatedXml = builder.buildObject(presentationObj);
        zip.file('ppt/presentation.xml', updatedXml);
      }
    }


    // 以下代码保持不变
    if (mergedOptions.compressImages) {
      const mediaFolder = zip.folder('ppt/media');
      if (mediaFolder) {
        const imageFiles = Object.keys(mediaFolder.files).filter(filename => 
          /\.(png|jpg|jpeg|gif|bmp|webp)$/i.test(filename)
        );


        const imageProcessTasks = imageFiles.map(async (filename) => {
          const file = mediaFolder.files[filename];
          const imageData = await file.async('arraybuffer');


          const optimizedImage = await compressImage(imageData, {
            maxWidth: mergedOptions.compressImages?.maxWidth,
            maxHeight: mergedOptions.compressImages?.maxHeight,
            quality: mergedOptions.compressImages?.quality
          });


          zip.file(`ppt/media/${filename}`, optimizedImage);
        });


        await Promise.all(imageProcessTasks);
      }
    }


    if (mergedOptions.removeUnusedMedia) {
      const usedMediaFiles = await collectUsedMedia(zip);
      const allMediaFiles = findAllMediaFiles(zip);
      const unusedMediaFiles = findUnusedMedia(allMediaFiles, usedMediaFiles);


      unusedMediaFiles.forEach(mediaPath => {
        delete zip.files[mediaPath];
      });
    }


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