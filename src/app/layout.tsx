import type { Metadata, Viewport } from 'next';
import { Inter } from 'next/font/google';
import ThemeRegistry from '@/theme/ThemeRegistry';

const inter = Inter({ 
  subsets: ['latin'],
  display: 'swap',
  fallback: ['Arial', 'sans-serif']
});

export const metadata: Metadata = {
  title: 'PPTX Optimizer - Compress PowerPoint Files Online',
  description: 'Optimize and compress your PowerPoint presentations while maintaining quality. Free online PPTX file optimizer.',
  keywords: 'PPTX optimizer, PowerPoint compression, reduce PowerPoint size, optimize presentations',
  manifest: '/manifest.json',
};

export const viewport: Viewport = {
  width: 'device-width',
  initialScale: 1,
  maximumScale: 1,
  userScalable: false,
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="en">
      <head>
        <link rel="icon" href="/favicon.ico" sizes="any" />
        <link rel="apple-touch-icon" href="/icon-192x192.png" />
      </head>
      <body className={inter.className}>
        <ThemeRegistry>
          {children}
        </ThemeRegistry>
      </body>
    </html>
  );
}