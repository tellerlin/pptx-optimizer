'use client';


import React, { useState } from 'react';
import { 
  Box, 
  Button, 
  CircularProgress, 
  Container, 
  Typography, 
  Alert, 
  Snackbar 
} from '@mui/material';
import { CloudUpload, CheckCircle, Error as ErrorIcon } from '@mui/icons-material';
import { optimizePPTX } from '../utils/pptx-optimizer';


export default function Home() {
  const [isProcessing, setIsProcessing] = useState(false);
  const [alertOpen, setAlertOpen] = useState(false);
  const [alertMessage, setAlertMessage] = useState('');
  const [alertSeverity, setAlertSeverity] = useState<'success' | 'error'>('success');


  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;


    // 文件类型验证
    if (!file.name.toLowerCase().endsWith('.pptx')) {
      showAlert('Please upload a valid PPTX file', 'error');
      return;
    }


    // 文件大小限制（例如：最大100MB）
    const maxSize = 300 * 1024 * 1024; // 100MB
    if (file.size > maxSize) {
      showAlert('File is too large. Maximum file size is 300MB', 'error');
      return;
    }


    setIsProcessing(true);
    try {
      const optimizedFile = await optimizePPTX(file);
      
      // 创建下载链接
      const url = URL.createObjectURL(optimizedFile);
      const link = document.createElement('a');
      link.href = url;
      link.download = `optimized_${file.name}`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);


      // 显示成功提示
      showAlert('PPTX file optimized successfully!', 'success');
    } catch (error) {
      console.error('Error optimizing PPTX:', error);
      showAlert('Error processing file. Please try again.', 'error');
    } finally {
      setIsProcessing(false);
    }
  };


  const showAlert = (message: string, severity: 'success' | 'error') => {
    setAlertMessage(message);
    setAlertSeverity(severity);
    setAlertOpen(true);
  };


  const handleAlertClose = (event?: React.SyntheticEvent | Event, reason?: string) => {
    if (reason === 'clickaway') {
      return;
    }
    setAlertOpen(false);
  };


  return (
    <Container maxWidth="md">
      <Box
        sx={{
          minHeight: '100vh',
          display: 'flex',
          flexDirection: 'column',
          alignItems: 'center',
          justifyContent: 'center',
          textAlign: 'center',
          gap: 4,
          p: 3,
        }}
      >
        <Typography 
          variant="h2" 
          component="h1" 
          gutterBottom 
          sx={{ 
            fontWeight: 'bold', 
            background: 'linear-gradient(45deg, #3f51b5, #2196f3)',
            WebkitBackgroundClip: 'text',
            WebkitTextFillColor: 'transparent'
          }}
        >
          PPTX Optimizer
        </Typography>
        
        <Typography 
          variant="body1" 
          color="text.secondary" 
          paragraph 
          sx={{ maxWidth: 600, mx: 'auto' }}
        >
          Upload your PowerPoint file to optimize and compress it while maintaining 
          quality. Reduce file size, remove hidden slides, and optimize images.
        </Typography>


        <Box sx={{ position: 'relative', display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
          <input
            accept=".pptx"
            style={{ display: 'none' }}
            id="file-upload"
            type="file"
            onChange={handleFileUpload}
            disabled={isProcessing}
          />
          <label htmlFor="file-upload">
            <Button
              variant="contained"
              component="span"
              startIcon={<CloudUpload />}
              disabled={isProcessing}
              sx={{ 
                py: 2, 
                px: 4, 
                borderRadius: 2,
                transition: 'transform 0.2s',
                '&:hover': {
                  transform: 'scale(1.05)'
                }
              }}
            >
              {isProcessing ? 'Processing...' : 'Upload PPTX'}
            </Button>
          </label>
          {isProcessing && (
            <CircularProgress
              size={24}
              sx={{
                position: 'absolute',
                top: '50%',
                left: '50%',
                marginTop: '10px',
                marginLeft: '-12px',
              }}
            />
          )}
        </Box>


        {/* 弹出提示 */}
        <Snackbar 
          open={alertOpen} 
          autoHideDuration={6000} 
          onClose={handleAlertClose}
          anchorOrigin={{ vertical: 'bottom', horizontal: 'center' }}
        >
          <Alert 
            onClose={handleAlertClose}
            severity={alertSeverity}
            sx={{ width: '100%' }}
            iconMapping={{
              success: <CheckCircle fontSize="inherit" />,
              error: <ErrorIcon fontSize="inherit" />
            }}
          >
            {alertMessage}
          </Alert>
        </Snackbar>
      </Box>
    </Container>
  );
}