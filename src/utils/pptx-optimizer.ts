import JSZip from 'jszip';  
import { parseStringPromise, Builder } from 'xml2js';  

const essentialFiles = [  
  'ppt/presentation.xml',  
  'docProps/core.xml',  
  'docProps/app.xml',  
  '[Content_Types].xml'  
];  

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

function findUnusedMedia(allMediaFiles: string[], usedMediaFiles: Set<string>): string[] {  
  return allMediaFiles.filter(mediaPath =>  
    !usedMediaFiles.has(mediaPath)  
  );  
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

    const slides = Array.isArray(sldIdLst['p:sldId'])  
      ? sldIdLst['p:sldId']  
      : [sldIdLst['p:sldId']];  

    const visibleSlides = slides.filter((slide: any) => {  
      const show = slide['$']?.show ?? slide.show;  
      return show !== '0' && show !== false;  
    });  

    if (visibleSlides.length === 0) {  
      console.warn('All slides would be removed, keeping original slides');  
      return;  
    }  

    presentationObj['p:presentation']['p:sldIdLst']['p:sldId'] = visibleSlides;  

    const builder = new Builder({  
      renderOpts: { pretty: true },  
      xmldec: { version: '1.0', encoding: 'UTF-8' }  
    });  

    const updatedXml = builder.buildObject(presentationObj);  
    zip.file('ppt/presentation.xml', updatedXml);  
    console.log('Removed unused slides');  
  } catch (error) {  
    console.error('Error removing unused slides:', error);  
  }  
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

    const presentationXml = zip.file('ppt/presentation.xml');  
    if (presentationXml) {  
      const presentationXmlContent = await presentationXml.async('string');  
      const presentationObj = await parseStringPromise(presentationXmlContent);  

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

async function optimizePPTX(file: File): Promise<Blob> {  
  if (typeof window === 'undefined') {  
    throw new Error('This function can only be used in a browser environment');  
  }  

  const startTime = performance.now();  
  const originalSize = file.size;  

  const arrayBuffer = await file.arrayBuffer();  
  const zip = new JSZip();  
  await zip.loadAsync(arrayBuffer);  
  console.log('PPTX file loaded into JSZip');  

  if (!await checkEssentialFiles(zip)) {  
    throw new Error('Essential files are missing in the original PPTX.');  
  }  

  try {  
    console.log('Processing images');  
    await processImages(zip);  
    console.log('Images processed');  

    if (!await checkEssentialFiles(zip)) {  
      throw new Error('Essential files are missing after processing images.');  
    }  

    console.log('Removing unused slides');  
    await removeUnusedSlides(zip);  
    console.log('Unused slides removed');  

    if (!await checkEssentialFiles(zip)) {  
      throw new Error('Essential files are missing after removing unused slides.');  
    }  

    console.log('Removing unused media files');  
    await removeUnusedMediaFiles(zip);  
    console.log('Unused media files removed');  

    if (!await checkEssentialFiles(zip)) {  
      throw new Error('Essential files are missing after removing unused media files.');  
    }  

    console.log('Removing embedded fonts');  
    await removeEmbeddedFonts(zip);  
    console.log('Embedded fonts removed');  
  } catch (error) {  
    console.error('Optimization process encountered an error:', error);  
    return new Blob([arrayBuffer], {  
      type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'  
    });  
  }  

  const fileList = await listFilesInZip(zip);  
  console.log('Files in the optimized zip:', JSON.stringify(fileList, null, 2));  

  if (!await checkEssentialFiles(zip)) {  
    console.error('Error: Essential files are missing after optimization.');  
    console.error('Missing files:', essentialFiles.filter(file => !zip.file(file)));  
    return new Blob([arrayBuffer], {  
      type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'  
    });  
  }  

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

async function removeUnusedMediaFiles(zip: JSZip): Promise<void> {  
  try {  
    const usedMediaFiles = await collectUsedMedia(zip);  
    const allMediaFiles = findAllMediaFiles(zip);  
    const unusedMediaFiles = findUnusedMedia(allMediaFiles, usedMediaFiles);  

    const filesToDelete = unusedMediaFiles.filter(file => !essentialFiles.includes(file));  

    filesToDelete.forEach(mediaPath => {  
      if (zip.files[mediaPath]) {  
        delete zip.files[mediaPath];  
        console.log(`Deleted: ${mediaPath}`);  
      }  
    });  

    console.log(`Removed ${filesToDelete.length} unused media files`);  
  } catch (error) {  
    console.error('Error in removeUnusedMediaFiles:', error);  
  }  
}  

function findAllMediaFiles(zip: JSZip): string[] {  
  const mediaExtensions = /\.(png|jpg|jpeg|gif|bmp|svg|bin)$/i;  
  const mediaFiles = Object.keys(zip.files)  
    .filter(name =>  
      (name.startsWith('ppt/media/') && mediaExtensions.test(name)) ||  
      name === 'docProps/thumbnail.jpeg'  
    )  
    .map(path => path.replace(/^(ppt\/media\/)+/, 'ppt/media/'))  
    .filter((path, index, self) => self.indexOf(path) === index);  
  return mediaFiles;  
}  

async function checkEssentialFiles(zip: JSZip): Promise<boolean> {  
  return essentialFiles.every(file => zip.file(file) !== null);  
}