import JSZip from 'jszip';
import { OptimizationOptions } from './types';
import { compressImage, createPlaceholderFile } from './media/image-optimizer';
import { 
  collectUsedMedia, 
  findAllMediaFiles, 
  findUnusedMedia,
  getMediaFileInfo 
} from './media/media-utils';

export async function optimizePPTX(
  file: File, 
  options: OptimizationOptions = {}
): Promise<Blob> {
  const startTime = performance.now();
  const originalSize = file.size;

  try {
    const zip = await JSZip.loadAsync(file);

    // 收集所有使用的媒体文件
    const usedMediaFiles = await collectUsedMedia(zip);
    const allMediaFiles = findAllMediaFiles(zip);
    
    // 处理媒体文件
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
            // 保持原始文件扩展名
            const extension = filename.split('.').pop()?.toLowerCase() || 'png';
            
            if (usedMediaFiles.has(filename)) {
              // 如果文件在XML中被引用
              if (options.usePlaceholders) {
                // 使用占位文件替换，但保持原始文件名
                const placeholderData = await createPlaceholderFile(
                  file.uncompressedSize,
                  extension
                );
                zip.file(filename, placeholderData);
              } else if (options.compressImages) {
                // 压缩图片，但保持原始文件名和扩展名
                const optimizedImage = await compressImage(imageData, options.compressImages);
                zip.file(filename, optimizedImage);
              }
            } else {
              // 如果文件未被引用，使用占位文件替换
              const extension = filename.split('.').pop()?.toLowerCase() || 'png';
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

    // 替换未使用的媒体文件为占位文件
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

    // 生成优化后的文件
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