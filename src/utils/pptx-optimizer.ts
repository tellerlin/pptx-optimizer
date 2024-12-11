import JSZip from 'jszip';
import { parseStringPromise, Builder } from 'xml2js';


interface ImageData {
  name: string;
  data: ArrayBuffer;
}


export async function optimizePPTX(file: File): Promise<Blob> {
  // 检查是否在浏览器环境
  if (typeof window === 'undefined') {
    throw new Error('This function can only be used in a browser environment');
  }


  // Read the PPTX file as ArrayBuffer
  const arrayBuffer = await file.arrayBuffer();
  
  // Load the PPTX file using JSZip
  const zip = new JSZip();
  await zip.loadAsync(arrayBuffer);
  
  // Process images in the PPTX
  await processImages(zip);
  
  // Remove unused elements
  await removeUnusedElements(zip);
  
  // Remove embedded fonts
  await removeEmbeddedFonts(zip);
  
  // Remove unused media files
  await removeUnusedMediaFiles(zip);
  
  // Generate the optimized PPTX file
  const options = {
    type: 'blob',
    compression: 'DEFLATE',
    compressionOptions: {
      level: 9
    }
  };
  
  return await zip.generateAsync(options);
}


async function processImages(zip: JSZip): Promise<void> {
  const mediaFolder = zip.folder('ppt/media');
  if (!mediaFolder) return;


  await Promise.all(
    Object.keys(mediaFolder.files).map(async (filename) => {
      if (!filename.match(/\.(png|jpg|jpeg|gif)$/i)) return;


      const file = mediaFolder.files[filename];
      if (!file || file.dir) return;


      try {
        // Get image data
        const imageData = await file.async('arraybuffer');
        
        // Optimize image in browser
        const optimizedImage = await optimizeImageInBrowser(imageData);
        
        // Update the image in the zip
        zip.file(`ppt/media/${filename}`, optimizedImage);
      } catch (error) {
        console.error(`Error processing image ${filename}:`, error);
      }
    })
  );
}


async function optimizeImageInBrowser(imageData: ArrayBuffer): Promise<ArrayBuffer> {
  return new Promise((resolve, reject) => {
    const blob = new Blob([imageData]);
    const url = URL.createObjectURL(blob);
    const img = new Image();
    
    img.onload = () => {
      const canvas = document.createElement('canvas');
      const ctx = canvas.getContext('2d');
      
      // Calculate new dimensions (max 1366x768)
      let { width, height } = img;
      if (width > 1366 || height > 768) {
        const ratio = Math.min(1366 / width, 768 / height);
        width *= ratio;
        height *= ratio;
      }
      
      canvas.width = width;
      canvas.height = height;
      
      // Draw and compress image
      ctx?.drawImage(img, 0, 0, width, height);
      
      // Convert to blob with quality 0.7
      canvas.toBlob(
        (blob) => {
          if (blob) {
            blob.arrayBuffer().then(resolve).catch(reject);
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


async function removeUnusedElements(zip: JSZip) {
  try {
    const presentationXml = await zip.file('ppt/presentation.xml')?.async('string');
    if (presentationXml) {
      const presentationObj = await parseStringPromise(presentationXml);
      
      // Remove hidden slides
      if (presentationObj['p:presentation'] && presentationObj['p:presentation']['p:sldIdLst']) {
        const slides = presentationObj['p:presentation']['p:sldIdLst'][0]['p:sldId'];
        presentationObj['p:presentation']['p:sldIdLst'][0]['p:sldId'] = 
          slides.filter((slide: any) => slide['$']['show'] !== '0');
      }


      const builder = new Builder();
      const updatedXml = builder.buildObject(presentationObj);
      
      zip.file('ppt/presentation.xml', updatedXml);
    }
  } catch (error) {
    console.error('Error removing unused elements:', error);
  }
}


async function removeEmbeddedFonts(zip: JSZip) {
  const commonFonts = [
    "Arial", "Calibri", "Times New Roman", 
    "Microsoft YaHei", "Helvetica", "Verdana"
  ];


  // Process slide layouts
  const slideLayoutFiles = zip.folder('ppt/slideLayouts')?.files || {};
  
  for (const filename in slideLayoutFiles) {
    const file = slideLayoutFiles[filename];
    if (file.dir) continue;


    try {
      const xmlContent = await file.async('string');
      const layoutObj = await parseStringPromise(xmlContent);
      
      // Remove embedded fonts
      if (layoutObj['p:sldLayout'] && layoutObj['p:sldLayout']['a:theme']) {
        const fonts = layoutObj['p:sldLayout']['a:theme'][0]['a:fontScheme'][0]['a:majorFont'][0]['a:font'];
        
        layoutObj['p:sldLayout']['a:theme'][0]['a:fontScheme'][0]['a:majorFont'][0]['a:font'] = 
          fonts.filter((font: any) => !commonFonts.includes(font['$']['name']));
      }


      const builder = new Builder();
      const updatedXml = builder.buildObject(layoutObj);
      
      zip.file(`ppt/slideLayouts/${filename}`, updatedXml);
    } catch (error) {
      console.error(`Error processing layout ${filename}:`, error);
    }
  }
}


async function removeUnusedMediaFiles(zip: JSZip) {
  const mediaFolder = zip.folder('ppt/media');
  const presentationXml = await zip.file('ppt/presentation.xml')?.async('string');
  
  if (mediaFolder && presentationXml) {
    try {
      const presentationObj = await parseStringPromise(presentationXml);
      const usedMediaFiles = new Set<string>();


      // Find used media files in slides
      const slides = presentationObj['p:presentation']['p:sldIdLst'][0]['p:sldId'];
      
      for (const slide of slides) {
        const slideXml = await zip.file(`ppt/slides/slide${slide['$']['id']}.xml`)?.async('string');
        if (slideXml) {
          const slideObj = await parseStringPromise(slideXml);
          
          // Extract media references
          const pics = slideObj['p:sld'][0]['p:cSld'][0]['p:pic'] || [];
          pics.forEach((pic: any) => {
            const blip = pic['p:blipFill'][0]['a:blip'][0]['$']['r:embed'];
            if (blip) usedMediaFiles.add(blip);
          });
        }
      }


      // Remove unused media files
      Object.keys(mediaFolder.files).forEach(filename => {
        if (!usedMediaFiles.has(filename)) {
          delete mediaFolder.files[filename];
        }
      });
    } catch (error) {
      console.error('Error removing unused media files:', error);
    }
  }
}