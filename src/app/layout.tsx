import type { Metadata, Viewport } from 'next';
import { Inter } from 'next/font/google';
import { AppRouterCacheProvider } from '@mui/material-nextjs/v14-appRouter';
import { ThemeProvider } from '@mui/material/styles';
import CssBaseline from '@mui/material/CssBaseline';
import theme from '@/theme';  // 使用 @/ 别名


const inter = Inter({ 
  subsets: ['latin'],
  display: 'swap',
  fallback: ['Arial', 'sans-serif']
});


export const metadata: Metadata = {
  title: 'PPTX Optimizer - Compress PowerPoint Files Online',
  description: 'Optimize and compress your PowerPoint presentations while maintaining quality. Free online PPTX file optimizer.',
  keywords: 'PPTX optimizer, PowerPoint compression, reduce PowerPoint size, optimize presentations',
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
      <body className={inter.className}>
        <AppRouterCacheProvider>
          <ThemeProvider theme={theme}>
            <CssBaseline />
            {children}
          </ThemeProvider>
        </AppRouterCacheProvider>
      </body>
    </html>
  );
}