import JSZip from 'jszip';  
import { parseStringPromise, Builder } from 'xml2js';  


async function processImages(zip: JSZip): Promise<void> {  
  const mediaFolder = zip.folder('ppt/media');  
  if (!mediaFolder) {  
    console.log('No media folder found');  
    return;  
  }  

  const imageFiles = Object.keys(mediaFolder.files).filter(filename =>  
    filename.match(/\.(png|jpg|jpeg|gif|bmp|webp)$/i)  
  );  

  console.log(`Found ${imageFiles.length} images to process`);  

  const imageProcessTasks = imageFiles.map(async (filename) => {  
    try {  
      const file = mediaFolder.files[filename];  
      if (!file || file.dir) {  
        console.log(`Skipping ${filename}: not a file`);  
        return;  
      }  

      console.log(`Processing image: ${filename}`);  
      const imageData = await file.async('arraybuffer');  
      const optimizedImage = await optimizeImageInBrowser(imageData);  
      zip.file(`ppt/media/${filename}`, optimizedImage);  
      console.log(`Processed image: ${filename}`);  
    } catch (error) {  
      console.warn(`Error processing image ${filename}:`, error);  
    }  
  });  

  await Promise.allSettled(imageProcessTasks);  
}  

async function collectUsedMedia(pptx: JSZip): Promise<Set<string>> {  
  const usedMediaFiles = new Set<string>();  
  const mediaTraversalTasks = [  
    { path: 'ppt/slides/_rels/', match: /\.rels$/ },  
    { path: 'ppt/slideLayouts/_rels/', match: /\.rels$/ },  
    { path: 'ppt/slideMasters/_rels/', match: /\.rels$/ },  
    { path: 'ppt/_rels/', match: /presentation\.xml\.rels$/ }  
  ];  

  const relsProcessTasks = mediaTraversalTasks.flatMap(task => {  
    const relsFiles = Object.keys(pptx.files)  
      .filter(name => name.startsWith(task.path) && task.match.test(name));  

    return relsFiles.map(async (relsFile) => {  
      try {  
        const relsContent = await pptx.file(relsFile)?.async('string');  
        if (!relsContent) return;  

        const mediaMatches = relsContent.match(/Target="\.\.\/media\/([^"]+)"/g) || [];  
        mediaMatches.forEach(match => {  
          const mediaPath = match.match(/Target="\.\.\/media\/([^"]+)"/)?.[1];  
          if (mediaPath) {  
            usedMediaFiles.add(`ppt/media/${mediaPath}`);  
          }  
        });  
      } catch (error) {  
        console.warn(`Error processing relationship file ${relsFile}:`, error);  
      }  
    });  
  });  

  const slideFiles = Object.keys(pptx.files)  
    .filter(name => name.startsWith('ppt/slides/') && name.endsWith('.xml'));  

  const slideProcessTasks = slideFiles.map(async (slideFile) => {  
    try {  
      const slideContent = await pptx.file(slideFile)?.async('string');  
      if (!slideContent) return;  

      const slideRelsFile = slideFile.replace('.xml', '.xml.rels');  
      const slideRelsContent = await pptx.file(slideRelsFile)?.async('string');  

      if (slideRelsContent) {  
        const rIdMatches = slideContent.matchAll(/r:embed="(rId\d+)"/g);  
        for (const match of rIdMatches) {  
          const rId = match[1];  
          const mediaMatch = slideRelsContent.match(new RegExp(`Id="${rId}"[^>]*Target="\.\.\/media\/([^"]+)"`));  
          if (mediaMatch) {  
            usedMediaFiles.add(`ppt/media/${mediaMatch[1]}`);  
          }  
        }  
      }  
    } catch (error) {  
      console.warn(`Error processing slide file ${slideFile}:`, error);  
    }  
  });  

  await Promise.allSettled([...relsProcessTasks, ...slideProcessTasks]);  

  return usedMediaFiles;  
}  


async function optimizeImageInBrowser(imageData: ArrayBuffer): Promise<ArrayBuffer> {  
  return new Promise((resolve, reject) => {  
    const blob = new Blob([imageData]);  
    const url = URL.createObjectURL(blob);  
    const img = new Image();  


    img.onload = () => {  
      const canvas = document.createElement('canvas');  
      const ctx = canvas.getContext('2d');  


      let { width, height } = img;  
      const maxWidth = 1366;  
      const maxHeight = 768;  


      if (width > maxWidth || height > maxHeight) {  
        const ratio = Math.min(maxWidth / width, maxHeight / height);  
        width = Math.round(width * ratio);  
        height = Math.round(height * ratio);  
      }  


      canvas.width = width;  
      canvas.height = height;  


      ctx?.drawImage(img, 0, 0, width, height);  


      // 根据图片类型选择压缩格式
      const mimeType = blob.type.startsWith('image/png') ? 'image/png' : 'image/webp';


      canvas.toBlob(  
        (blob) => {  
          if (blob) {  
            blob.arrayBuffer()  
              .then(resolve)  
              .catch(reject);  
          } else {  
            reject(new Error('Failed to create blob'));  
          }  
        },  
        mimeType,  
        mimeType === 'image/webp' ? 0.7 : 0.8  // PNG质量略高
      );  


      URL.revokeObjectURL(url);  
    };  


    img.onerror = () => {  
      URL.revokeObjectURL(url);  
      reject(new Error('Failed to load image'));  
    };  


    img.src = url;  
  });  
}


async function removeUnusedSlides(zip: JSZip): Promise<void> {  
  try {  
    const presentationXml = await zip.file('ppt/presentation.xml')?.async('string');
    if (!presentationXml) {  
      console.warn('No presentation.xml found');  
      return;  
    }  


    const presentationObj = await parseStringPromise(presentationXml, {  
      explicitArray: false,  
      mergeAttrs: true,  
      trim: true  
    });  


    const sldIdLst = presentationObj['p:presentation']?.['p:sldIdLst'];
    if (!sldIdLst || !sldIdLst['p:sldId']) {  
      console.warn('No slides found in presentation');  
      return;  
    }  


    let slides = sldIdLst['p:sldId'];
    if (!Array.isArray(slides)) {
      slides = [slides];
    }


    console.log('Total slides found:', slides.length);


    const visibleSlides = await Promise.all(slides.map(async (slide: any) => {
      const slideRId = slide['$']?.['r:id'] || slide['r:id'];
      const slideId = slide['$']?.id || slide.id;


      console.log(`Slide RId: ${slideRId}, Slide ID: ${slideId}`);


      // 使用关系ID查找对应的幻灯片文件
      const matchingSlideFiles = Object.keys(zip.files)
        .filter(filename => 
          filename.startsWith('ppt/slides/slide') && 
          filename.endsWith('.xml')
        );


      const slideFileName = matchingSlideFiles[0]; // 取第一个匹配的幻灯片文件


      if (!slideFileName) {
        console.warn(`No slide file found for RId: ${slideRId}`);
        return null;
      }


      const slideFile = zip.file(slideFileName);
      if (!slideFile) {
        console.warn(`Slide file not found: ${slideFileName}`);
        return null;
      }


      // 额外检查slide文件内容是否为空
      const slideContent = await slideFile.async('string');
      if (!slideContent || slideContent.trim() === '') {
        console.warn(`Slide file is empty: ${slideFileName}`);
        return null;
      }


      return slide;
    }));


    // 过滤掉空值
    const validSlides = visibleSlides.filter(slide => slide !== null);


    if (validSlides.length === 0) {
      console.warn('All slides would be removed, keeping original slides');
      return;
    }


    // 更新presentation.xml中的slide列表
    presentationObj['p:presentation']['p:sldIdLst']['p:sldId'] = validSlides;


    const builder = new Builder({  
      renderOpts: { pretty: true },  
      xmldec: { version: '1.0', encoding: 'UTF-8' }  
    });  


    const updatedXml = builder.buildObject(presentationObj);  
    zip.file('ppt/presentation.xml', updatedXml);  
    console.log(`Removed ${slides.length - validSlides.length} unused slides`);
  } catch (error) {  
    console.error('Error removing unused slides:', error);  
  }  
}


async function listFilesInZip(zip: JSZip): Promise<{ [filename: string]: number }> {  
  const fileList: { [filename: string]: number } = {};  


  const filePromises = Object.keys(zip.files).map(async (relativePath) => {
    const file = zip.files[relativePath];
    if (!file.dir) {
      try {
        // 尝试文本解析
        try {
          const content = await file.async('string');
          fileList[relativePath] = content.length;
        } catch (textError) {
          // 文本解析失败，尝试获取文件大小
          try {
            const binaryContent = await file.async('arraybuffer');
            fileList[relativePath] = binaryContent.byteLength;
          } catch (binaryError) {
            // 如果二进制读取也失败，使用安全的方法获取大小
            fileList[relativePath] = file.options.uncompressedSize || 0;
            console.warn(`Could not read file ${relativePath}:`, {
              textError,
              binaryError
            });
          }
        }
      } catch (error) {
        console.error(`Error processing file ${relativePath}:`, error);
      }
    }
  });


  await Promise.all(filePromises);
  return fileList;  
}


async function removeUnusedMediaFiles(zip: JSZip): Promise<void> {  
  try {  
    // 首先收集所有被使用的媒体文件
    const usedMediaFiles = await collectUsedMedia(zip);  
    
    // 找出所有媒体文件
    const allMediaFiles = findAllMediaFiles(zip);  
    
    // 找出未使用的媒体文件
    const unusedMediaFiles = findUnusedMedia(allMediaFiles, usedMediaFiles);  


    // 记录调试信息
    console.log('All Media Files:', allMediaFiles);
    console.log('Used Media Files:', Array.from(usedMediaFiles));
    console.log('Unused Media Files:', unusedMediaFiles);


    // 增加更严格的过滤和保护
    const filesToDelete = unusedMediaFiles.filter(mediaPath => {
      // 检查文件是否被保护
      if (zip.files[mediaPath]?._protected) {
        console.warn(`Prevented deletion of protected file: ${mediaPath}`);
        return false;
      }


      // 安全检查，避免删除关键文件和目录
      const safeToDelete = 
        // 排除幻灯片相关文件
        !mediaPath.includes('slides/') &&
        !mediaPath.includes('slideLayouts/') &&
        !mediaPath.includes('slideMasters/') &&
        
        // 排除关键文件
        !mediaPath.includes('presentation.xml') &&
        !mediaPath.includes('core.xml') &&
        !mediaPath.includes('app.xml') &&
        !mediaPath.includes('[Content_Types].xml') &&
        
        // 排除关系文件
        !mediaPath.endsWith('.rels') &&
        
        // 确保是媒体文件
        /\.(png|jpg|jpeg|gif|bmp|svg|webp|bin)$/i.test(mediaPath);


      // 额外的日志记录
      if (!safeToDelete) {
        console.warn(`Skipping deletion of potentially important file: ${mediaPath}`);
      }


      return safeToDelete;
    }); 


    // 记录删除前的文件数量
    const initialFileCount = Object.keys(zip.files).length;


    // 执行删除操作
    const deletionResults = filesToDelete.map(mediaPath => {
      const file = zip.files[mediaPath];
      if (file) {
        // 增加额外的安全检查
        if (!file._protected) {
          try {
            delete zip.files[mediaPath];
            console.log(`Safely deleted unused media file: ${mediaPath}`);
            return { path: mediaPath, deleted: true };
          } catch (error) {
            console.warn(`Failed to delete file ${mediaPath}:`, error);
            return { path: mediaPath, deleted: false, error };
          }
        } else {
          console.warn(`Prevented deletion of protected file: ${mediaPath}`);
          return { path: mediaPath, deleted: false, reason: 'protected' };
        }
      }
      return { path: mediaPath, deleted: false, reason: 'not found' };
    });


    // 计算并记录删除的文件数量
    const deletedFilesCount = deletionResults.filter(result => result.deleted).length;
    const finalFileCount = Object.keys(zip.files).length;


    console.log(`Removed ${deletedFilesCount} unused media files`);
    console.log(`Total files before deletion: ${initialFileCount}`);
    console.log(`Total files after deletion: ${finalFileCount}`);


    // 详细的删除结果日志
    console.log('Deletion Results:', deletionResults);


  } catch (error) {  
    console.error('Error in removeUnusedMediaFiles:', error);
    
    // 更详细的错误处理
    if (error instanceof Error) {
      console.error('Error details:', {
        name: error.name,
        message: error.message,
        stack: error.stack
      });
    }


    // 可以选择抛出错误或者只记录
    throw error;
  }  
}


function findAllMediaFiles(zip: JSZip): string[] {  
  const mediaExtensions = /\.(png|jpg|jpeg|gif|bmp|svg|bin|webp)$/i;  
  const mediaFiles = Object.keys(zip.files)  
    .filter(name =>  
      (name.startsWith('ppt/media/') && mediaExtensions.test(name)) ||  
      name === 'docProps/thumbnail.jpeg'  
    )
    .map(path => path.replace(/^(ppt\/media\/)+/, 'ppt/media/'))  
    .filter((path, index, self) => self.indexOf(path) === index);  
  return mediaFiles;  
}


function findUnusedMedia(allMediaFiles: string[], usedMediaFiles: Set<string>): string[] {  
  return allMediaFiles.filter(mediaPath =>  
    !usedMediaFiles.has(mediaPath)  
  );  
}


export async function optimizePPTX(file: File): Promise<Blob> {  
  if (typeof window === 'undefined') {  
    throw new Error('This function can only be used in a browser environment');  
  }  


  const startTime = performance.now();  
  const originalSize = file.size;  


  const arrayBuffer = await file.arrayBuffer();  
  const zip = new JSZip();  
  await zip.loadAsync(arrayBuffer);  
  console.log('PPTX file loaded into JSZip');  


  // 定义必须保留的文件和路径
  const criticalFiles = [
    'ppt/presentation.xml',
    'ppt/_rels/presentation.xml.rels',
    '[Content_Types].xml',
    'docProps/core.xml', 
    'docProps/app.xml',
    'ppt/slides/slide1.xml',
    'ppt/slides/_rels/slide1.xml.rels'
  ];


  const criticalPaths = [
    'ppt/slides/',
    'ppt/slideLayouts/',
    'ppt/slideMasters/',
    'ppt/_rels/',
    'docProps/',
    'ppt/theme/'
  ];


  try {
    // 图片处理
    console.log('Processing images');  
    await processImages(zip);  
    console.log('Images processed');  


    // 移除未使用的幻灯片
    console.log('Removing unused slides');  
    await removeUnusedSlides(zip);  
    console.log('Unused slides removed');  


    // 移除未使用的媒体文件
    console.log('Removing unused media files');  
    const usedMediaFiles = await collectUsedMedia(zip);
    const allMediaFiles = findAllMediaFiles(zip);
    const unusedMediaFiles = findUnusedMedia(allMediaFiles, usedMediaFiles);


    // 安全删除未使用的媒体文件
    const deletionResults = unusedMediaFiles
      .filter(mediaPath => {
        // 确保不删除关键文件和路径下的文件
        const isCritical = criticalFiles.includes(mediaPath) || 
          criticalPaths.some(path => mediaPath.startsWith(path));
        
        // 只删除媒体文件，且不是缩略图
        const isMediaFile = /\.(png|jpg|jpeg|gif|bmp|svg|webp|bin)$/i.test(mediaPath) && 
          !mediaPath.includes('thumbnail');


        return !isCritical && isMediaFile;
      })
      .map(mediaPath => {
        try {
          if (zip.files[mediaPath]) {
            console.log(`Deleting unused media file: ${mediaPath}`);
            delete zip.files[mediaPath];
            return { path: mediaPath, deleted: true };
          }
          return { path: mediaPath, deleted: false, reason: 'not found' };
        } catch (error) {
          console.warn(`Failed to delete file ${mediaPath}:`, error);
          return { path: mediaPath, deleted: false, error };
        }
      });


    console.log('Unused media files removal results:', deletionResults);


    // 确保关键文件存在
    criticalFiles.forEach(filePath => {
      if (!zip.files[filePath]) {
        console.warn(`Critical file missing: ${filePath}. Skipping deletion.`);
      }
    });


  } catch (error) {  
    console.error('Optimization process encountered an error:', error);
    throw error; // 直接抛出错误，不再返回原文件
  }


  // 最终文件列表检查
  const fileList = await listFilesInZip(zip);  
  console.log('Files in the optimized zip:', JSON.stringify(fileList, null, 2));  


  const options = {  
    type: 'blob',  
    compression: 'DEFLATE',  
    compressionOptions: {  
      level: 9  
    }  
  };  


  const optimizedBlob = await zip.generateAsync(options);  


  const endTime = performance.now();  
  console.log(`  
    Optimization took: ${(endTime - startTime).toFixed(2)} ms  
    Original file size: ${originalSize} bytes  
    Optimized file size: ${optimizedBlob.size} bytes  
  `);  


  return optimizedBlob;  
}
