'use client';


import React, { useState } from 'react';
import { 
  Box, Button, CircularProgress, Container, Typography, 
  Alert, Snackbar, Dialog, DialogTitle, DialogContent, 
  DialogActions, Checkbox, FormControlLabel, Slider, Grid 
} from '@mui/material';
import { CloudUpload, CheckCircle, Error as ErrorIcon } from '@mui/icons-material';
import { optimizePPTX } from '../utils/pptx-optimizer';


export default function Home() {
  const [isProcessing, setIsProcessing] = useState(false);
  const [alertOpen, setAlertOpen] = useState(false);
  const [alertMessage, setAlertMessage] = useState('');
  const [alertSeverity, setAlertSeverity] = useState<'success' | 'error'>('success');
  const [optimizationDialogOpen, setOptimizationDialogOpen] = useState(false);
  const [removeHiddenSlides, setRemoveHiddenSlides] = useState(true);
  const [compressImages, setCompressImages] = useState(true);
  const [imageQuality, setImageQuality] = useState(70);
  const [removeUnusedMedia, setRemoveUnusedMedia] = useState(true);


  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;


    if (!file.name.toLowerCase().endsWith('.pptx')) {
      showAlert('Please upload a valid PPTX file', 'error');
      return;
    }


    const maxSize = 300 * 1024 * 1024;
    if (file.size > maxSize) {
      showAlert('File is too large. Maximum file size is 300MB', 'error');
      return;
    }


    setOptimizationDialogOpen(true);
  };


  const performOptimization = async () => {
    const fileInput = document.getElementById('file-upload') as HTMLInputElement;
    const file = fileInput.files?.[0];
    if (!file) return;


    setOptimizationDialogOpen(false);
    setIsProcessing(true);


    try {
      const optimizedFile = await optimizePPTX(file, {
        removeHiddenSlides,
        compressImages: compressImages ? {
          quality: imageQuality / 100,
          maxWidth: 1920,
          maxHeight: 1080
        } : undefined,
        removeUnusedMedia
      });
      
      const url = URL.createObjectURL(optimizedFile);
      const link = document.createElement('a');
      link.href = url;
      link.download = `optimized_${file.name}`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);


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

        <Dialog 
          open={optimizationDialogOpen} 
          onClose={() => setOptimizationDialogOpen(false)}
          maxWidth="sm"
          fullWidth
          disableEnforceFocus
          disableRestoreFocus

        >
          <DialogTitle>Optimization Settings</DialogTitle>
          <DialogContent>
            <Grid container spacing={2}>
              <Grid item xs={12}>
                <FormControlLabel
                  control={
                    <Checkbox
                      checked={removeHiddenSlides}
                      onChange={(e) => setRemoveHiddenSlides(e.target.checked)}
                    />
                  }
                  label="Remove Hidden Slides"
                />
              </Grid>
              
              <Grid item xs={12}>
                <FormControlLabel
                  control={
                    <Checkbox
                      checked={compressImages}
                      onChange={(e) => setCompressImages(e.target.checked)}
                    />
                  }
                  label="Compress Images"
                />
                {compressImages && (
                  <Box sx={{ px: 2, pt: 1 }}>
                    <Typography gutterBottom>
                      Image Quality: {imageQuality}%
                    </Typography>
                    <Slider
                      value={imageQuality}
                      onChange={(_, newValue) => setImageQuality(newValue as number)}
                      min={10}
                      max={100}
                      valueLabelDisplay="auto"
                    />
                  </Box>
                )}
              </Grid>
              
              <Grid item xs={12}>
                <FormControlLabel
                  control={
                    <Checkbox
                      checked={removeUnusedMedia}
                      onChange={(e) => setRemoveUnusedMedia(e.target.checked)}
                    />
                  }
                  label="Remove Unused Media Files"
                />
              </Grid>
            </Grid>
          </DialogContent>
          <DialogActions>
            <Button onClick={() => setOptimizationDialogOpen(false)}>
              Cancel
              </Button>
            <Button 
              onClick={performOptimization} 
              variant="contained" 
              color="primary"
            >
              Optimize
            </Button>
          </DialogActions>
        </Dialog>




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