import { ImageOptimizationOptions } from '../types';

/**
 * Creates a minimal placeholder image file
 * @param originalSize - Size of the original file in bytes
 * @param extension - File extension (e.g., 'png', 'jpg')
 * @returns ArrayBuffer containing the placeholder file
 */
export async function createPlaceholderFile(originalSize: number, extension: string): Promise<ArrayBuffer> {
    // Create a minimal valid file based on extension
    switch (extension.toLowerCase()) {
        case 'png':
            return createMinimalPNG();
        case 'jpg':
        case 'jpeg':
            return createMinimalJPEG();
        default:
            // For other formats, create a 1x1 transparent PNG
            return createMinimalPNG();
    }
}

/**
 * Creates a minimal valid PNG file (1x1 transparent pixel)
 */
function createMinimalPNG(): ArrayBuffer {
    // PNG header + IHDR chunk + IDAT chunk + IEND chunk
    const minimalPNG = new Uint8Array([
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,  // PNG signature
        0x00, 0x00, 0x00, 0x0D,                          // IHDR chunk length
        0x49, 0x48, 0x44, 0x52,                          // "IHDR"
        0x00, 0x00, 0x00, 0x01,                          // width: 1
        0x00, 0x00, 0x00, 0x01,                          // height: 1
        0x08,                                            // bit depth
        0x06,                                            // color type: RGBA
        0x00,                                            // compression method
        0x00,                                            // filter method
        0x00,                                            // interlace method
        0x1F, 0x15, 0xC4, 0x89,                          // IHDR CRC
        0x00, 0x00, 0x00, 0x0A,                          // IDAT chunk length
        0x49, 0x44, 0x41, 0x54,                          // "IDAT"
        0x78, 0x9C, 0x63, 0x00, 0x00, 0x00, 0x02, 0x00, 0x01, // compressed data
        0xE5, 0x27, 0x0E, 0x89,                          // IDAT CRC
        0x00, 0x00, 0x00, 0x00,                          // IEND chunk length
        0x49, 0x45, 0x4E, 0x44,                          // "IEND"
        0xAE, 0x42, 0x60, 0x82                           // IEND CRC
    ]);
    return minimalPNG.buffer;
}

/**
 * Creates a minimal valid JPEG file (1x1 pixel)
 */
function createMinimalJPEG(): ArrayBuffer {
    // Minimal JPEG structure with a 1x1 gray pixel
    const minimalJPEG = new Uint8Array([
        0xFF, 0xD8,                   // SOI marker
        0xFF, 0xE0, 0x00, 0x10,      // APP0 segment
        0x4A, 0x46, 0x49, 0x46, 0x00,// JFIF identifier
        0x01, 0x01,                   // version
        0x00,                         // units
        0x00, 0x01, 0x00, 0x01,      // density
        0x00, 0x00,                   // thumbnail
        0xFF, 0xDB, 0x00, 0x43, 0x00,// DQT marker
        ...Array(64).fill(1),         // quantization table
        0xFF, 0xC0, 0x00, 0x0B,      // SOF0 marker
        0x08, 0x00, 0x01, 0x00, 0x01,// parameters
        0x01, 0x00,                   // components
        0xFF, 0xDA, 0x00, 0x08,      // SOS marker
        0x01, 0x00, 0x00, 0x3F, 0x00,// parameters
        0xFF, 0xD9                    // EOI marker
    ]);
    return minimalJPEG.buffer;
}

interface ImageAnalysis {
    hasAlpha: boolean;
}

function checkAlphaChannel(imageData: ImageData): boolean {
    const data = imageData.data;
    for (let i = 3; i < data.length; i += 4) {
        if (data[i] < 255) {
            return true;
        }
    }
    return false;
}

function calculateOptimalDimensions(
    originalWidth: number,
    originalHeight: number,
    maxWidth = 1366,
    maxHeight = 768
): { width: number; height: number } {
    let width = originalWidth;
    let height = originalHeight;

    if (width > maxWidth) {
        height = Math.round((height * maxWidth) / width);
        width = maxWidth;
    }

    if (height > maxHeight) {
        width = Math.round((width * maxHeight) / height);
        height = maxHeight;
    }

    return { width, height };
}

/**
 * Compresses an image based on the provided options
 */
export async function compressImage(
    imageData: ArrayBuffer,
    options: ImageOptimizationOptions = {}
): Promise<ArrayBuffer> {
    // Create blob and bitmap from input data
    const blob = new Blob([imageData]);
    const bitmap = await createImageBitmap(blob);

    // Calculate optimal dimensions
    const { width, height } = calculateOptimalDimensions(bitmap.width, bitmap.height);

    // Create canvas and resize image
    const canvas = new OffscreenCanvas(width, height);
    const ctx = canvas.getContext('2d');
    if (!ctx) {
        throw new Error('Failed to get canvas context');
    }
    ctx.drawImage(bitmap, 0, 0, width, height);

    // Check for alpha channel
    const imageDataForAnalysis = ctx.getImageData(0, 0, width, height);
    const hasAlpha = checkAlphaChannel(imageDataForAnalysis);

    // Compress based on alpha channel presence
    if (hasAlpha) {
        // Use WebP for images with transparency
        const blob = await canvas.convertToBlob({ 
            type: 'image/webp', 
            quality: 0.7 
        });
        return await blob.arrayBuffer();
    } else {
        // Compare WebP and JPEG for non-transparent images
        const [webpBlob, jpegBlob] = await Promise.all([
            canvas.convertToBlob({ type: 'image/webp', quality: 0.7 }),
            canvas.convertToBlob({ type: 'image/jpeg', quality: 0.7 })
        ]);

        const [webpBuffer, jpegBuffer] = await Promise.all([
            webpBlob.arrayBuffer(),
            jpegBlob.arrayBuffer()
        ]);

        // Return the smaller of the two formats
        return webpBuffer.byteLength <= jpegBuffer.byteLength ? webpBuffer : jpegBuffer;
    }
}