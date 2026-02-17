import React, { useState } from 'react'
import {
  Box,
  Card,
  CardContent,
  Grid,
  Typography,
  Button,
  Alert,
  LinearProgress,
  Paper,
  Chip,
  TextField,
  MenuItem,
} from '@mui/material'
import {
  CloudUpload as UploadIcon,
  Description as PdfIcon,
  CheckCircle as SuccessIcon,
  CalendarToday as CalendarIcon,
  Tune as FilterIcon,
} from '@mui/icons-material'
import { DatePicker, LocalizationProvider } from '@mui/x-date-pickers'
import { AdapterDayjs } from '@mui/x-date-pickers/AdapterDayjs'
import dayjs from 'dayjs'
import { useNavigate } from 'react-router-dom'
import apiService from '../services/api'

export default function Upload({ sessionId }) {
  const navigate = useNavigate()

  // LMRB State
  const [startDate, setStartDate] = useState(dayjs().subtract(30, 'days'))
  const [endDate, setEndDate] = useState(dayjs())
  const [selectedChannel, setSelectedChannel] = useState('All')
  const [selectedProducts, setSelectedProducts] = useState([])
  const [channels, setChannels] = useState([])
  const [products, setProducts] = useState([])
  const [lmrbLoading, setLmrbLoading] = useState(false)
  const [lmrbData, setLmrbData] = useState(null)

  // TC State
  const [tcFile, setTcFile] = useState(null)
  const [tcLoading, setTcLoading] = useState(false)
  const [tcData, setTcData] = useState(null)

  const [error, setError] = useState('')

  const handleLoadLMRB = async () => {
    setLmrbLoading(true)
    setError('')

    try {
      const result = await apiService.getLMRBData(
        sessionId,
        startDate.format('YYYY-MM-DD'),
        endDate.format('YYYY-MM-DD'),
        selectedChannel
      )

      setLmrbData(result)
      setChannels(['All', ...result.channels])
      setProducts(result.products)
    } catch (err) {
      setError(err.response?.data?.error || 'Failed to load LMRB data')
    } finally {
      setLmrbLoading(false)
    }
  }

  const handleTCUpload = async () => {
    if (!tcFile) {
      setError('Please select a TC PDF file')
      return
    }

    setTcLoading(true)
    setError('')

    try {
      const result = await apiService.uploadTCPDF(tcFile, sessionId)
      setTcData(result)
    } catch (err) {
      setError(err.response?.data?.error || 'Failed to upload TC PDF')
    } finally {
      setTcLoading(false)
    }
  }

  const handleFileSelect = (event) => {
    const file = event.target.files[0]
    if (file && file.type === 'application/pdf') {
      setTcFile(file)
      setError('')
    } else {
      setError('Please select a valid PDF file')
    }
  }

  const canProceed = lmrbData && tcData

  return (
    <LocalizationProvider dateAdapter={AdapterDayjs}>
      <Box>
        <Typography variant="h4" gutterBottom fontWeight={600}>
          Data Upload & Configuration
        </Typography>
        <Typography variant="body1" color="text.secondary" paragraph>
          Load LMRB data from Google Sheets and upload TC PDF file
        </Typography>

        {error && (
          <Alert severity="error" sx={{ mb: 3 }} onClose={() => setError('')}>
            {error}
          </Alert>
        )}

        <Grid container spacing={3}>
          {/* LMRB Data Section */}
          <Grid item xs={12} md={6}>
            <Card>
              <CardContent>
                <Box display="flex" alignItems="center" mb={2}>
                  <CalendarIcon color="primary" sx={{ mr: 1, fontSize: 32 }} />
                  <Typography variant="h6">LMRB Data (Google Sheets)</Typography>
                </Box>

                <Grid container spacing={2}>
                  <Grid item xs={12} sm={6}>
                    <DatePicker
                      label="Start Date"
                      value={startDate}
                      onChange={setStartDate}
                      slotProps={{ textField: { fullWidth: true } }}
                    />
                  </Grid>
                  <Grid item xs={12} sm={6}>
                    <DatePicker
                      label="End Date"
                      value={endDate}
                      onChange={setEndDate}
                      slotProps={{ textField: { fullWidth: true } }}
                    />
                  </Grid>

                  {channels.length > 0 && (
                    <Grid item xs={12}>
                      <TextField
                        select
                        fullWidth
                        label="Filter by Channel"
                        value={selectedChannel}
                        onChange={(e) => setSelectedChannel(e.target.value)}
                      >
                        {channels.map((channel) => (
                          <MenuItem key={channel} value={channel}>
                            {channel}
                          </MenuItem>
                        ))}
                      </TextField>
                    </Grid>
                  )}

                  <Grid item xs={12}>
                    <Button
                      variant="contained"
                      size="large"
                      onClick={handleLoadLMRB}
                      disabled={lmrbLoading}
                      fullWidth
                      startIcon={<FilterIcon />}
                    >
                      {lmrbLoading ? 'Loading...' : 'Load LMRB Data'}
                    </Button>
                  </Grid>
                </Grid>

                {lmrbLoading && <LinearProgress sx={{ mt: 2 }} />}

                {lmrbData && (
                  <Box mt={3}>
                    <Alert severity="success" icon={<SuccessIcon />}>
                      Loaded {lmrbData.record_count} LMRB records
                    </Alert>
                    <Paper variant="outlined" sx={{ p: 2, mt: 2 }}>
                      <Typography variant="subtitle2" color="primary" gutterBottom>
                        Data Summary
                      </Typography>
                      <Typography variant="body2">
                        <strong>Records:</strong> {lmrbData.record_count}
                      </Typography>
                      <Typography variant="body2">
                        <strong>Channels:</strong> {lmrbData.channels?.length || 0}
                      </Typography>
                      <Typography variant="body2">
                        <strong>Products:</strong> {lmrbData.products?.length || 0}
                      </Typography>
                      {lmrbData.date_range && (
                        <Typography variant="body2">
                          <strong>Date Range:</strong> {lmrbData.date_range.min} to {lmrbData.date_range.max}
                        </Typography>
                      )}
                    </Paper>
                  </Box>
                )}
              </CardContent>
            </Card>
          </Grid>

          {/* TC PDF Upload Section */}
          <Grid item xs={12} md={6}>
            <Card>
              <CardContent>
                <Box display="flex" alignItems="center" mb={2}>
                  <PdfIcon color="secondary" sx={{ mr: 1, fontSize: 32 }} />
                  <Typography variant="h6">TC PDF Upload</Typography>
                </Box>

                <Typography variant="body2" color="text.secondary" paragraph>
                  Upload Telecast Certificate PDF file for automatic conversion
                </Typography>

                <input
                  type="file"
                  accept="application/pdf"
                  onChange={handleFileSelect}
                  style={{ display: 'none' }}
                  id="tc-file-input"
                />

                <Box mb={2}>
                  <label htmlFor="tc-file-input">
                    <Button
                      variant="outlined"
                      component="span"
                      fullWidth
                      startIcon={<UploadIcon />}
                    >
                      Select PDF File
                    </Button>
                  </label>
                </Box>

                {tcFile && (
                  <Paper variant="outlined" sx={{ p: 2, mb: 2 }}>
                    <Typography variant="body2">
                      <strong>Selected:</strong> {tcFile.name}
                    </Typography>
                    <Typography variant="caption" color="text.secondary">
                      Size: {(tcFile.size / 1024).toFixed(2)} KB
                    </Typography>
                  </Paper>
                )}

                <Button
                  variant="contained"
                  color="secondary"
                  size="large"
                  onClick={handleTCUpload}
                  disabled={!tcFile || tcLoading}
                  fullWidth
                >
                  {tcLoading ? 'Converting...' : 'Upload & Convert PDF'}
                </Button>

                {tcLoading && <LinearProgress sx={{ mt: 2 }} />}

                {tcData && (
                  <Box mt={3}>
                    <Alert severity="success" icon={<SuccessIcon />}>
                      Converted {tcData.record_count} TC records
                    </Alert>
                    <Paper variant="outlined" sx={{ p: 2, mt: 2 }}>
                      <Typography variant="subtitle2" color="primary" gutterBottom>
                        TC Data Summary
                      </Typography>
                      <Typography variant="body2">
                        <strong>Records:</strong> {tcData.record_count}
                      </Typography>
                      <Typography variant="body2">
                        <strong>Themes:</strong> {tcData.themes?.length || 0}
                      </Typography>
                    </Paper>
                  </Box>
                )}
              </CardContent>
            </Card>
          </Grid>

          {/* Product Filter (if LMRB loaded) */}
          {lmrbData && products.length > 0 && (
            <Grid item xs={12}>
              <Card>
                <CardContent>
                  <Box display="flex" alignItems="center" mb={2}>
                    <FilterIcon color="primary" sx={{ mr: 1 }} />
                    <Typography variant="h6">Product Filter (Optional)</Typography>
                  </Box>
                  <Typography variant="body2" color="text.secondary" paragraph>
                    Select specific products to include in reconciliation (leave empty for all)
                  </Typography>
                  <Box display="flex" flexWrap="wrap" gap={1}>
                    {products.map((product) => (
                      <Chip
                        key={product}
                        label={product}
                        onClick={() => {
                          if (selectedProducts.includes(product)) {
                            setSelectedProducts(selectedProducts.filter(p => p !== product))
                          } else {
                            setSelectedProducts([...selectedProducts, product])
                          }
                        }}
                        color={selectedProducts.includes(product) ? 'primary' : 'default'}
                        variant={selectedProducts.includes(product) ? 'filled' : 'outlined'}
                      />
                    ))}
                  </Box>
                  {selectedProducts.length > 0 && (
                    <Box mt={2}>
                      <Typography variant="caption" color="text.secondary">
                        Selected: {selectedProducts.length} products
                      </Typography>
                      <Button
                        size="small"
                        onClick={() => setSelectedProducts([])}
                        sx={{ ml: 2 }}
                      >
                        Clear All
                      </Button>
                    </Box>
                  )}
                </CardContent>
              </Card>
            </Grid>
          )}

          {/* Continue Button */}
          {canProceed && (
            <Grid item xs={12}>
              <Card sx={{ bgcolor: 'primary.main', color: 'white' }}>
                <CardContent>
                  <Box display="flex" alignItems="center" justifyContent="space-between">
                    <Box>
                      <Typography variant="h6">
                        Data Ready!
                      </Typography>
                      <Typography variant="body2">
                        LMRB and TC data loaded successfully. Continue to theme mapping.
                      </Typography>
                    </Box>
                    <Button
                      variant="contained"
                      color="secondary"
                      size="large"
                      onClick={() => navigate('/theme-mapping')}
                    >
                      Continue to Theme Mapping
                    </Button>
                  </Box>
                </CardContent>
              </Card>
            </Grid>
          )}
        </Grid>
      </Box>
    </LocalizationProvider>
  )
}
