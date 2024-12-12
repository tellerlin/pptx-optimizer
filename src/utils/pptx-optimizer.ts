import JSZip from 'jszip';  
import { parseStringPromise, Builder } from 'xml2js';  

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
        'image/webp',  
        0.7  
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


async function listFilesInZip(zip: JSZip): Promise<{ [filename: string]: number }> {  
  const fileList: { [filename: string]: number } = {};  

  await zip.forEach((relativePath, file) => {  
    if (file.dir) {  
      return;  
    }  
    zip.file(relativePath).async('string').then((content) => {  
      fileList[relativePath] = content.length;  
    });  
  });  

  return fileList;  
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

  try {  
    console.log('Processing images');  
    await processImages(zip);  
    console.log('Images processed');  

    console.log('Removing unused slides');  
    await removeUnusedSlides(zip);  
    console.log('Unused slides removed');  

    console.log('Removing unused media files');  
    await removeUnusedMediaFiles(zip);  
    console.log('Unused media files removed');  

    console.log('Removing embedded fonts');  
    await removeEmbeddedFonts(zip);  
    console.log('Embedded fonts removed');  
  } catch (error) {  
    console.error('Optimization process encountered an error:', error);  
    return new Blob([arrayBuffer], {  
      type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'  
    });  
  }  

  // List all files in the zip to check their status  
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
    Size change: ${((optimizedBlob.size - originalSize) / originalSize * 100).toFixed(2)}%  
  `);  

  return optimizedBlob;  
}

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

async function removeEmbeddedFonts(zip: JSZip): Promise<void> {  
  try {  
    const fontFolder = zip.folder('ppt/fonts');  
    if (fontFolder) {  
      console.log('Found fonts folder, removing fonts');  
      Object.keys(fontFolder.files).forEach(filename => {  
        delete fontFolder.files[filename];  
      });  
    } else {  
      console.log('No fonts folder found');  
    }  

    const presentationXml = await zip.file('ppt/presentation.xml')?.async('string');  
    if (presentationXml) {  
      const presentationObj = await parseStringPromise(presentationXml);  

      if (presentationObj['p:presentation'] && presentationObj['p:presentation']['a:extLst']) {  
        delete presentationObj['p:presentation']['a:extLst'];  
      }  

      const builder = new Builder();  
      const updatedXml = builder.buildObject(presentationObj);  
      zip.file('ppt/presentation.xml', updatedXml);  
      console.log('Removed embedded fonts');  
    } else {  
      console.warn('No presentation.xml found');  
    }  
  } catch (error) {  
    console.warn('Error removing embedded fonts:', error);  
  }  
}


async function removeUnusedSlides(zip: JSZip): Promise<void> {
  try {
    const presentationXml = await zip.file('ppt/presentation.xml')?.async('string');
    if (!presentationXml) {
      console.warn('No presentation.xml found');
      return;
    }


    console.log('Raw presentation.xml content:', presentationXml);


    const presentationObj = await parseStringPromise(presentationXml, {
      explicitArray: false,
      mergeAttrs: true,
      trim: true
    });


    console.log('Parsed presentationObj structure:', JSON.stringify(presentationObj, null, 2));


    // 详细打印 presentationObj 的关键路径
    console.log('p:presentation keys:', Object.keys(presentationObj['p:presentation'] || {}));
    console.log('p:sldIdLst:', presentationObj['p:presentation']?.['p:sldIdLst']);


    const sldIdLst = presentationObj['p:presentation']?.['p:sldIdLst'];
    if (!sldIdLst || !sldIdLst['p:sldId']) {
      console.warn('No slides found in presentation');
      console.log('Full presentationObj:', JSON.stringify(presentationObj, null, 2));
      return;
    }


    // 确保 p:sldId 总是数组
    const slides = Array.isArray(sldIdLst['p:sldId']) 
      ? sldIdLst['p:sldId'] 
      : [sldIdLst['p:sldId']];


    console.log('Slides before filtering:', JSON.stringify(slides, null, 2));


    const visibleSlides = slides.filter((slide: any) => {
      // 更健壮的判断逻辑
      const show = slide['$']?.show ?? slide.show;
      console.log('Slide show value:', show);
      return show !== '0' && show !== false;
    });


    console.log('Visible slides:', JSON.stringify(visibleSlides, null, 2));


    if (visibleSlides.length > 0) {
      presentationObj['p:presentation']['p:sldIdLst']['p:sldId'] = visibleSlides;
    } else {
      console.warn('No visible slides found after filtering');
    }


    const builder = new Builder({
      renderOpts: { pretty: true },
      xmldec: { version: '1.0', encoding: 'UTF-8' }
    });


    const updatedXml = builder.buildObject(presentationObj);
    console.log('Updated presentation.xml:', updatedXml);


    zip.file('ppt/presentation.xml', updatedXml);
  } catch (error) {
    console.error('Error removing unused slides:', error);
  }
}


async function removeUnusedMediaFiles(zip: JSZip): Promise<void> {
  const mediaFolder = zip.folder('ppt/media');
  if (!mediaFolder) {
    console.log('No media folder found');
    return;
  }


  try {
    const usedMediaFiles = new Set<string>();
    console.log('Starting media file analysis...');


    // 1. 处理幻灯片关系文件
    const slideRelsFiles = Object.keys(zip.files).filter(
      file => file.startsWith('ppt/slides/') && file.endsWith('.rels')
    );


    console.log('Slide relationship files found:', slideRelsFiles);


    for (const relsFile of slideRelsFiles) {
      try {
        const relsXml = await zip.file(relsFile)?.async('string');
        if (!relsXml) {
          console.warn(`Could not read relationship file: ${relsFile}`);
          continue;
        }


        const relsObj = await parseStringPromise(relsXml, {
          explicitArray: false,
          mergeAttrs: true,
          trim: true
        });


        console.log(`Relationships in ${relsFile}:`, JSON.stringify(relsObj, null, 2));


        const relationships = relsObj?.Relationships?.Relationship || [];
        const mediaRels = Array.isArray(relationships) 
          ? relationships 
          : [relationships];


        mediaRels.forEach((rel: any) => {
          if (rel?.$ && 
              rel.$['Type'] === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image') {
            const target = rel.$['Target'];
            if (target) {
              // 提取最后的文件名，支持多种路径格式
              const mediaFilename = target.split('/').pop();
              if (mediaFilename) {
                usedMediaFiles.add(mediaFilename);
                console.log(`Found media reference in relationships: ${mediaFilename}`);
              }
            }
          }
        });


      } catch (relsError) {
        console.warn(`Error processing relationships file ${relsFile}:`, relsError);
      }
    }


    // 2. 处理幻灯片 XML 文件
    const slideFiles = Object.keys(zip.files).filter(
      file => file.startsWith('ppt/slides/') && file.endsWith('.xml')
    );


    console.log('Slide XML files found:', slideFiles);


    for (const slideFile of slideFiles) {
      try {
        const slideXml = await zip.file(slideFile)?.async('string');
        if (!slideXml) {
          console.warn(`Could not read slide file: ${slideFile}`);
          continue;
        }


        const imageRefPatterns = [
          /r:embed="(rId\d+)"/g,
          /a:blip\s+r:embed="(rId\d+)"/g
        ];


        const slideMediaRefs = new Set<string>();
        
        imageRefPatterns.forEach(pattern => {
          const matches = slideXml.matchAll(pattern);
          for (const match of matches) {
            slideMediaRefs.add(match[1]);
          }
        });


        console.log(`Media references in ${slideFile}:`, Array.from(slideMediaRefs));


        // 查找对应的关系文件
        const relsFile = slideFile.replace('.xml', '.xml.rels');
        const relsXml = await zip.file(relsFile)?.async('string');
        
        if (relsXml) {
          const relsObj = await parseStringPromise(relsXml, {
            explicitArray: false,
            mergeAttrs: true,
            trim: true
          });


          const relationships = relsObj?.Relationships?.Relationship || [];
          const mediaRels = Array.isArray(relationships) 
            ? relationships 
            : [relationships];


          slideMediaRefs.forEach(rId => {
            const matchingRel = mediaRels.find((rel: any) => 
              rel.$?.['Id'] === rId && 
              rel.$?.['Type'] === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
            );


            if (matchingRel) {
              const target = matchingRel.$['Target'];
              const mediaFilename = target.split('/').pop();
              if (mediaFilename) {
                usedMediaFiles.add(mediaFilename);
                console.log(`Found media reference in slide relationships: ${mediaFilename}`);
              }
            }
          });
        }


      } catch (slideError) {
        console.warn(`Error processing slide ${slideFile}:`, slideError);
      }
    }


    // 3. 检查演示文稿关系文件
    const presentationRelsFile = 'ppt/_rels/presentation.xml.rels';
    const presentationRelsXml = await zip.file(presentationRelsFile)?.async('string');
    
    if (presentationRelsXml) {
      const relsObj = await parseStringPromise(presentationRelsXml, {
        explicitArray: false,
        mergeAttrs: true,
        trim: true
      });


      const relationships = relsObj?.Relationships?.Relationship || [];
      const mediaRels = Array.isArray(relationships) 
        ? relationships 
        : [relationships];


      mediaRels.forEach((rel: any) => {
        if (rel?.$ && 
            rel.$['Type'] === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image') {
          const target = rel.$['Target'];
          const mediaFilename = target.split('/').pop();
          if (mediaFilename) {
            usedMediaFiles.add(mediaFilename);
            console.log(`Found media reference in presentation relationships: ${mediaFilename}`);
          }
        }
      });
    }


    console.log('Used media files:', Array.from(usedMediaFiles));


    // 删除未使用的媒体文件
    const allMediaFiles = Object.keys(zip.files)
      .filter(file => 
        (file.startsWith('ppt/media/') || file.startsWith('docProps/')) && 
        file.match(/\.(png|jpg|jpeg|gif|bmp|svg)$/i)
      )
      .map(file => file.split('/').pop());
    
    const uniqueAllMediaFiles = [...new Set(allMediaFiles)];
    const unusedMediaFiles = uniqueAllMediaFiles.filter(
      file => !Array.from(usedMediaFiles).some(usedFile => 
        file === usedFile || 
        (usedFile && file.includes(usedFile)) || 
        (file && usedFile.includes(file))
      )
    );


    console.log('All media files:', uniqueAllMediaFiles);
    console.log('Unused media files:', unusedMediaFiles);


    // 删除未使用的媒体文件
    unusedMediaFiles.forEach(filename => {
      const mediaPath = `ppt/media/${filename}`;
      const docPropsPath = `docProps/${filename}`;
      
      if (zip.files[mediaPath]) {
        delete zip.files[mediaPath];
        console.log(`Deleted unused media file: ${mediaPath}`);
      }
      
      if (zip.files[docPropsPath]) {
        delete zip.files[docPropsPath];
        console.log(`Deleted unused media file: ${docPropsPath}`);
      }
    });


    console.log(`Removed ${unusedMediaFiles.length} unused media files`);


  } catch (error) {
    console.error('Comprehensive error in removeUnusedMediaFiles:', error);
  }
}