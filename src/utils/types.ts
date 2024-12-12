export interface OptimizationOptions {
  removeHiddenSlides?: boolean;
  compressImages?: ImageOptimizationOptions;
  removeUnusedMedia?: boolean;
  usePlaceholders?: boolean; // 新选项：是否使用占位文件替换被压缩的媒体文件
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
  preserveOriginalSize?: boolean; // 是否保持与原文件相同大小
  format?: 'png' | 'jpg'; // 占位文件格式
}