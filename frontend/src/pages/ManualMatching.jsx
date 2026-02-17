import React, { useState, useEffect } from 'react'
import {
  Box,
  Card,
  CardContent,
  Typography,
  Button,
  Grid,
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TableHead,
  TableRow,
  Alert,
  Radio,
  Paper,
} from '@mui/material'
import {
  Link as LinkIcon,
} from '@mui/icons-material'
import apiService from '../services/api'

export default function ManualMatching({ sessionId }) {
  const [unmatchedLMRB, setUnmatchedLMRB] = useState([])
  const [unmatchedTC, setUnmatchedTC] = useState([])
  const [selectedLMRB, setSelectedLMRB] = useState(null)
  const [selectedTC, setSelectedTC] = useState(null)
  const [error, setError] = useState('')
  const [success, setSuccess] = useState('')
  const [loading, setLoading] = useState(false)

  useEffect(() => {
    loadUnmatchedData()
  }, [])

  const loadUnmatchedData = async () => {
    try {
      const lmrbResult = await apiService.getUnmatchedResults(sessionId, 'lmrb')
      const tcResult = await apiService.getUnmatchedResults(sessionId, 'tc')

      setUnmatchedLMRB(lmrbResult.data || [])
      setUnmatchedTC(tcResult.data || [])
    } catch (err) {
      setError('Failed to load unmatched records')
    }
  }

  const handleMatch = async () => {
    if (selectedLMRB === null || selectedTC === null) {
      setError('Please select one record from each table')
      return
    }

    setLoading(true)
    setError('')
    setSuccess('')

    try {
      await apiService.manualMatch(sessionId, selectedLMRB, selectedTC)
      setSuccess('Records matched successfully!')

      // Reload unmatched data
      await loadUnmatchedData()

      // Clear selections
      setSelectedLMRB(null)
      setSelectedTC(null)
    } catch (err) {
      setError(err.response?.data?.error || 'Failed to match records')
    } finally {
      setLoading(false)
    }
  }

  return (
    <Box>
      <Typography variant="h4" gutterBottom fontWeight={600}>
        Manual Matching
      </Typography>
      <Typography variant="body1" color="text.secondary" paragraph>
        Manually match unmatched LMRB and TC records
      </Typography>

      {error && (
        <Alert severity="error" sx={{ mb: 2 }} onClose={() => setError('')}>
          {error}
        </Alert>
      )}

      {success && (
        <Alert severity="success" sx={{ mb: 2 }} onClose={() => setSuccess('')}>
          {success}
        </Alert>
      )}

      <Grid container spacing={3}>
        {/* Unmatched LMRB */}
        <Grid item xs={12} md={6}>
          <Card>
            <CardContent>
              <Typography variant="h6" gutterBottom>
                Unmatched LMRB Records ({unmatchedLMRB.length})
              </Typography>

              {unmatchedLMRB.length === 0 ? (
                <Alert severity="info">No unmatched LMRB records</Alert>
              ) : (
                <TableContainer component={Paper} variant="outlined" sx={{ maxHeight: 500 }}>
                  <Table stickyHeader size="small">
                    <TableHead>
                      <TableRow>
                        <TableCell padding="checkbox"></TableCell>
                        <TableCell>Date</TableCell>
                        <TableCell>Theme</TableCell>
                        <TableCell>Time</TableCell>
                        <TableCell>Duration</TableCell>
                      </TableRow>
                    </TableHead>
                    <TableBody>
                      {unmatchedLMRB.map((row, index) => (
                        <TableRow
                          key={index}
                          selected={selectedLMRB === index}
                          onClick={() => setSelectedLMRB(index)}
                          sx={{ cursor: 'pointer' }}
                        >
                          <TableCell padding="checkbox">
                            <Radio checked={selectedLMRB === index} />
                          </TableCell>
                          <TableCell>{row.Date || 'N/A'}</TableCell>
                          <TableCell>{row.Theme || 'N/A'}</TableCell>
                          <TableCell>{row.Time || 'N/A'}</TableCell>
                          <TableCell>{row.Duration || 'N/A'}s</TableCell>
                        </TableRow>
                      ))}
                    </TableBody>
                  </Table>
                </TableContainer>
              )}
            </CardContent>
          </Card>
        </Grid>

        {/* Unmatched TC */}
        <Grid item xs={12} md={6}>
          <Card>
            <CardContent>
              <Typography variant="h6" gutterBottom>
                Unmatched TC Records ({unmatchedTC.length})
              </Typography>

              {unmatchedTC.length === 0 ? (
                <Alert severity="info">No unmatched TC records</Alert>
              ) : (
                <TableContainer component={Paper} variant="outlined" sx={{ maxHeight: 500 }}>
                  <Table stickyHeader size="small">
                    <TableHead>
                      <TableRow>
                        <TableCell padding="checkbox"></TableCell>
                        <TableCell>Date</TableCell>
                        <TableCell>Theme</TableCell>
                        <TableCell>Time</TableCell>
                        <TableCell>Duration</TableCell>
                      </TableRow>
                    </TableHead>
                    <TableBody>
                      {unmatchedTC.map((row, index) => (
                        <TableRow
                          key={index}
                          selected={selectedTC === index}
                          onClick={() => setSelectedTC(index)}
                          sx={{ cursor: 'pointer' }}
                        >
                          <TableCell padding="checkbox">
                            <Radio checked={selectedTC === index} />
                          </TableCell>
                          <TableCell>{row.Date || 'N/A'}</TableCell>
                          <TableCell>{row.Theme || 'N/A'}</TableCell>
                          <TableCell>{row.Time || 'N/A'}</TableCell>
                          <TableCell>{row.Duration || 'N/A'}s</TableCell>
                        </TableRow>
                      ))}
                    </TableBody>
                  </Table>
                </TableContainer>
              )}
            </CardContent>
          </Card>
        </Grid>

        {/* Match Button */}
        <Grid item xs={12}>
          <Card>
            <CardContent>
              <Box display="flex" alignItems="center" justifyContent="space-between">
                <Typography variant="body1">
                  {selectedLMRB !== null && selectedTC !== null
                    ? 'Ready to match selected records'
                    : 'Select one record from each table to match'}
                </Typography>
                <Button
                  variant="contained"
                  size="large"
                  onClick={handleMatch}
                  disabled={selectedLMRB === null || selectedTC === null || loading}
                  startIcon={<LinkIcon />}
                >
                  {loading ? 'Matching...' : 'Match Selected Records'}
                </Button>
              </Box>
            </CardContent>
          </Card>
        </Grid>
      </Grid>
    </Box>
  )
}
