import JSZip from 'jszip';
import { parseStringPromise } from 'xml2js';
import { MediaFile } from '../types';

export async function collectUsedMedia(zip: JSZip): Promise<Set<string>> {
  const usedMediaFiles = new Set<string>();

  try {
    // Get all slide files
    const slideFiles = Object.keys(zip.files)
      .filter(filename => filename.startsWith('ppt/slides/slide') && filename.endsWith('.xml'));

    await Promise.all(slideFiles.map(async (slideFile) => {
      try {
        const slideXml = await zip.file(slideFile)?.async('string');
        if (!slideXml) return;

        // Find all media references in the slide
        const mediaRefs = slideXml.match(/r:embed="[^"]*"/g) || [];
        
        // Process each media reference
        await Promise.all(mediaRefs.map(async (ref) => {
          const relId = ref.replace(/r:embed="([^"]*)"/g, '$1');
          const relsFile = `ppt/slides/_rels/${slideFile.split('/').pop()}.rels`;
          
          try {
            const relsXml = await zip.file(relsFile)?.async('string');
            if (!relsXml) return;

            const relsObj = await parseStringPromise(relsXml);
            const relationships = relsObj.Relationships?.Relationship || [];
            
            relationships.forEach((rel: any) => {
              if (rel.$.Id === relId) {
                const mediaPath = `ppt/media/${rel.$.Target.split('/').pop()}`;
                usedMediaFiles.add(mediaPath);
              }
            });
          } catch (error) {
            console.warn(`Error processing relationships in ${relsFile}:`, error);
          }
        }));
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

export function findAllMediaFiles(zip: JSZip): string[] {
  return Object.keys(zip.files)
    .filter(filename => 
      filename.startsWith('ppt/media/') && 
      /\.(png|jpg|jpeg|gif|bmp|webp|wmf|emf|svg)$/i.test(filename)
    );
}

export function findUnusedMedia(
  allMediaFiles: string[], 
  usedMediaFiles: Set<string>
): string[] {
  return allMediaFiles.filter(mediaFile => !usedMediaFiles.has(mediaFile));
}

export async function getMediaFileInfo(
  zip: JSZip, 
  mediaPath: string
): Promise<MediaFile | null> {
  const file = zip.file(mediaPath);
  if (!file) return null;

  const extension = mediaPath.split('.').pop()?.toLowerCase() || '';
  const type = getMediaType(extension);

  return {
    path: mediaPath,
    type,
    size: file.uncompressedSize
  };
}

function getMediaType(extension: string): string {
  const mediaTypes: Record<string, string> = {
    'png': 'image/png',
    'jpg': 'image/jpeg',
    'jpeg': 'image/jpeg',
    'gif': 'image/gif',
    'bmp': 'image/bmp',
    'webp': 'image/webp',
    'wmf': 'image/x-wmf',
    'emf': 'image/x-emf',
    'svg': 'image/svg+xml'
  };

  return mediaTypes[extension] || 'application/octet-stream';
}