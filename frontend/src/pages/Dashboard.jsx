import React, { useState, useEffect } from 'react'
import {
  Box,
  Card,
  CardContent,
  Grid,
  Typography,
  Button,
  TextField,
  Alert,
  LinearProgress,
  Paper,
} from '@mui/material'
import {
  CloudQueue as CloudIcon,
  CheckCircle as SuccessIcon,
  Info as InfoIcon,
} from '@mui/icons-material'
import { useNavigate } from 'react-router-dom'
import apiService from '../services/api'

export default function Dashboard({ sessionId }) {
  const navigate = useNavigate()
  const [sheetId, setSheetId] = useState('')
  const [loading, setLoading] = useState(false)
  const [connected, setConnected] = useState(false)
  const [metadata, setMetadata] = useState(null)
  const [error, setError] = useState('')

  const handleConnect = async () => {
    if (!sheetId.trim()) {
      setError('Please enter a Google Sheet ID')
      return
    }

    setLoading(true)
    setError('')

    try {
      const result = await apiService.connectGoogleSheet(sheetId, sessionId)
      setMetadata(result.metadata)
      setConnected(true)
    } catch (err) {
      setError(err.response?.data?.error || 'Failed to connect to Google Sheet')
    } finally {
      setLoading(false)
    }
  }

  return (
    <Box>
      <Typography variant="h4" gutterBottom fontWeight={600}>
        Dashboard
      </Typography>
      <Typography variant="body1" color="text.secondary" paragraph>
        Welcome to the Media Reconciliation System. Connect to your Google Sheet to get started.
      </Typography>

      <Grid container spacing={3}>
        {/* Google Sheets Connection */}
        <Grid item xs={12} md={8}>
          <Card>
            <CardContent>
              <Box display="flex" alignItems="center" mb={2}>
                <CloudIcon color="primary" sx={{ mr: 1, fontSize: 32 }} />
                <Typography variant="h6">Google Sheets Connection</Typography>
              </Box>

              {!connected ? (
                <Box>
                  <Typography variant="body2" color="text.secondary" paragraph>
                    Enter your Google Sheet ID to connect to the LMRB data source.
                  </Typography>

                  <Alert severity="info" sx={{ mb: 2 }}>
                    <strong>Where to find Sheet ID:</strong><br />
                    Copy from URL: docs.google.com/spreadsheets/d/<strong>YOUR_SHEET_ID</strong>/edit
                  </Alert>

                  <TextField
                    fullWidth
                    label="Google Sheet ID"
                    variant="outlined"
                    value={sheetId}
                    onChange={(e) => setSheetId(e.target.value)}
                    placeholder="1ABC...XYZ"
                    disabled={loading}
                    sx={{ mb: 2 }}
                  />

                  {error && (
                    <Alert severity="error" sx={{ mb: 2 }}>
                      {error}
                    </Alert>
                  )}

                  <Button
                    variant="contained"
                    size="large"
                    onClick={handleConnect}
                    disabled={loading}
                    fullWidth
                  >
                    {loading ? 'Connecting...' : 'Connect to Google Sheet'}
                  </Button>

                  {loading && <LinearProgress sx={{ mt: 2 }} />}
                </Box>
              ) : (
                <Box>
                  <Alert severity="success" icon={<SuccessIcon />} sx={{ mb: 2 }}>
                    Successfully connected to Google Sheet!
                  </Alert>

                  {metadata && (
                    <Paper variant="outlined" sx={{ p: 2, mb: 2 }}>
                      <Typography variant="subtitle2" color="primary" gutterBottom>
                        Sheet Information
                      </Typography>
                      <Typography variant="body2">
                        <strong>Title:</strong> {metadata.title}
                      </Typography>
                      {metadata.sheets && metadata.sheets.length > 0 && (
                        <Typography variant="body2">
                          <strong>Sheets:</strong> {metadata.sheets.map(s => s.name).join(', ')}
                        </Typography>
                      )}
                    </Paper>
                  )}

                  <Button
                    variant="contained"
                    color="primary"
                    onClick={() => navigate('/upload')}
                    fullWidth
                  >
                    Continue to Upload TC PDF
                  </Button>
                </Box>
              )}
            </CardContent>
          </Card>
        </Grid>

        {/* Quick Guide */}
        <Grid item xs={12} md={4}>
          <Card>
            <CardContent>
              <Box display="flex" alignItems="center" mb={2}>
                <InfoIcon color="info" sx={{ mr: 1 }} />
                <Typography variant="h6">Quick Guide</Typography>
              </Box>

              <Typography variant="body2" paragraph>
                <strong>Step 1:</strong> Connect to Google Sheet with LMRB data
              </Typography>

              <Typography variant="body2" paragraph>
                <strong>Step 2:</strong> Upload TC PDF file
              </Typography>

              <Typography variant="body2" paragraph>
                <strong>Step 3:</strong> Configure theme mappings
              </Typography>

              <Typography variant="body2" paragraph>
                <strong>Step 4:</strong> Run reconciliation
              </Typography>

              <Typography variant="body2">
                <strong>Step 5:</strong> Review results and export
              </Typography>
            </CardContent>
          </Card>
        </Grid>

        {/* Features */}
        <Grid item xs={12}>
          <Card>
            <CardContent>
              <Typography variant="h6" gutterBottom>
                Features
              </Typography>
              <Grid container spacing={2}>
                <Grid item xs={12} sm={6} md={3}>
                  <Paper variant="outlined" sx={{ p: 2, textAlign: 'center' }}>
                    <Typography variant="h6" color="primary">
                      Google Sheets
                    </Typography>
                    <Typography variant="body2" color="text.secondary">
                      Direct integration with LMRB data
                    </Typography>
                  </Paper>
                </Grid>
                <Grid item xs={12} sm={6} md={3}>
                  <Paper variant="outlined" sx={{ p: 2, textAlign: 'center' }}>
                    <Typography variant="h6" color="primary">
                      PDF Conversion
                    </Typography>
                    <Typography variant="body2" color="text.secondary">
                      Auto-convert TC PDFs to data
                    </Typography>
                  </Paper>
                </Grid>
                <Grid item xs={12} sm={6} md={3}>
                  <Paper variant="outlined" sx={{ p: 2, textAlign: 'center' }}>
                    <Typography variant="h6" color="primary">
                      Smart Matching
                    </Typography>
                    <Typography variant="body2" color="text.secondary">
                      Intelligent reconciliation engine
                    </Typography>
                  </Paper>
                </Grid>
                <Grid item xs={12} sm={6} md={3}>
                  <Paper variant="outlined" sx={{ p: 2, textAlign: 'center' }}>
                    <Typography variant="h6" color="primary">
                      Export Reports
                    </Typography>
                    <Typography variant="body2" color="text.secondary">
                      Excel and PDF exports
                    </Typography>
                  </Paper>
                </Grid>
              </Grid>
            </CardContent>
          </Card>
        </Grid>
      </Grid>
    </Box>
  )
}
