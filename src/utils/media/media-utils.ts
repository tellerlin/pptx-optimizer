import JSZip from 'jszip';
import { parseStringPromise } from 'xml2js';
import { MediaFile } from '../types';

export async function collectUsedMedia(zip: JSZip): Promise<Set<string>> {
  const usedMediaFiles = new Set<string>();

  try {
    // 获取所有幻灯片文件
    const slideFiles = Object.keys(zip.files)
      .filter(filename => filename.startsWith('ppt/slides/slide') && filename.endsWith('.xml'));

    await Promise.all(slideFiles.map(async (slideFile) => {
      try {
        // 处理幻灯片中的媒体文件
        await processFileMedia(zip, slideFile, usedMediaFiles);

        // 获取幻灯片的 rels 文件
        const slideRelsFile = `ppt/slides/_rels/${slideFile.split('/').pop()}.rels`;
        const slideRelsXml = await zip.file(slideRelsFile)?.async('string');
        if (!slideRelsXml) return;

        // 解析 rels 文件找到关联的 layout
        const slideRelsObj = await parseStringPromise(slideRelsXml);
        const relationships = slideRelsObj.Relationships?.Relationship || [];
        const layoutRel = relationships.find((rel: any) => 
          rel.$.Type === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'
        );

        if (layoutRel) {
          // 处理 layout 中的媒体文件
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

// 处理文件中的媒体引用
async function processFileMedia(zip: JSZip, filePath: string, usedMediaFiles: Set<string>): Promise<void> {
  try {
    const fileContent = await zip.file(filePath)?.async('string');
    if (!fileContent) return;

    // 查找所有媒体引用
    const mediaRefs = fileContent.match(/r:embed="[^"]*"/g) || [];

    // 获取文件的 rels 路径
    const relsFile = `${filePath.substring(0, filePath.lastIndexOf('/'))}/_rels/${filePath.split('/').pop()}.rels`;
    const relsContent = await zip.file(relsFile)?.async('string');
    if (!relsContent) return;

    const relsObj = await parseStringPromise(relsContent);
    const relationships = relsObj.Relationships?.Relationship || [];

    // 处理每个媒体引用
    mediaRefs.forEach(ref => {
      const relId = ref.replace(/r:embed="([^"]*)"/g, '$1');
      const mediaRel = relationships.find((rel: any) => rel.$.Id === relId);
      
      if (mediaRel && mediaRel.$.Type.includes('/image')) {
        const mediaPath = `ppt/media/${mediaRel.$.Target.split('/').pop()}`;
        usedMediaFiles.add(mediaPath);
      }
    });
  } catch (error) {
    console.warn(`Error processing media in ${filePath}:`, error);
  }
}

async function getUsedLayoutsAndMasters(zip: JSZip): Promise<{
  layouts: Set<string>;
  masters: Set<string>;
}> {
  const usedLayouts = new Set<string>();
  const usedMasters = new Set<string>();

  try {
    // 获取 presentation.xml 内容
    const presentationXml = await zip.file('ppt/presentation.xml')?.async('string');
    if (!presentationXml) return { layouts: usedLayouts, masters: usedMasters };

    // 解析 presentation.xml
    const presentationObj = await parseStringPromise(presentationXml);
    const sldIdLst = presentationObj?.['p:presentation']?.['p:sldIdLst']?.[0]?.['p:sldId'] || [];

    // 遍历所有非隐藏的幻灯片
    for (const slide of sldIdLst) {
      const slideRId = slide.$?.['r:id'];
      if (!slideRId) continue;

      // 获取幻灯片关系文件
      const presentationRelsXml = await zip.file('ppt/_rels/presentation.xml.rels')?.async('string');
      if (!presentationRelsXml) continue;

      const presentationRels = await parseStringPromise(presentationRelsXml);
      const relationships = presentationRels.Relationships?.Relationship || [];

      // 找到幻灯片对应的关系
      const slideRel = relationships.find((rel: any) => rel.$.Id === slideRId);
      if (!slideRel) continue;

      const slideFile = slideRel.$.Target.replace('../', 'ppt/');
      const slideXml = await zip.file(slideFile)?.async('string');
      if (!slideXml) continue;

      const slideObj = await parseStringPromise(slideXml);
      const layoutRId = slideObj?.['p:sld']?.$?.['r:id'];
      if (!layoutRId) continue;

      // 获取 slide 的 rels 文件
      const slideRelsFile = `${slideFile.replace(/[^/]+$/, '')}_rels/${slideFile.split('/').pop()}.rels`;
      const slideRelsXml = await zip.file(slideRelsFile)?.async('string');
      if (!slideRelsXml) continue;

      const slideRels = await parseStringPromise(slideRelsXml);
      const layoutRel = slideRels.Relationships?.Relationship?.find((rel: any) => rel.$.Id === layoutRId);
      if (!layoutRel) continue;

      const layoutPath = layoutRel.$.Target.replace('../', 'ppt/');
      usedLayouts.add(layoutPath);

      // 获取 layout 对应的 master
      const layoutXml = await zip.file(layoutPath)?.async('string');
      if (!layoutXml) continue;

      const layoutObj = await parseStringPromise(layoutXml);
      const masterRId = layoutObj?.['p:sldLayout']?.$?.['r:id'];
      if (!masterRId) continue;

      const layoutRelsFile = `${layoutPath.replace(/[^/]+$/, '')}_rels/${layoutPath.split('/').pop()}.rels`;
      const layoutRelsXml = await zip.file(layoutRelsFile)?.async('string');
      if (!layoutRelsXml) continue;

      const layoutRels = await parseStringPromise(layoutRelsXml);
      const masterRel = layoutRels.Relationships?.Relationship?.find((rel: any) => rel.$.Id === masterRId);
      if (!masterRel) continue;

      const masterPath = masterRel.$.Target.replace('../', 'ppt/');
      usedMasters.add(masterPath);
    }
  } catch (error) {
    console.warn('Error getting used layouts and masters:', error);
  }

  return { layouts: usedLayouts, masters: usedMasters };
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