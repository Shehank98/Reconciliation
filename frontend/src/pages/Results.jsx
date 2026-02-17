import React, { useState, useEffect } from 'react'
import {
  Box,
  Card,
  CardContent,
  Typography,
  Button,
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TableHead,
  TableRow,
  TablePagination,
  Paper,
  Tabs,
  Tab,
  Alert,
  LinearProgress,
  Grid,
  Chip,
} from '@mui/material'
import {
  Download as DownloadIcon,
  PictureAsPdf as PdfIcon,
} from '@mui/icons-material'
import apiService from '../services/api'

export default function Results({ sessionId }) {
  const [tab, setTab] = useState(0)
  const [loading, setLoading] = useState(false)
  const [matchedData, setMatchedData] = useState([])
  const [unmatchedLMRB, setUnmatchedLMRB] = useState([])
  const [unmatchedTC, setUnmatchedTC] = useState([])
  const [page, setPage] = useState(0)
  const [rowsPerPage, setRowsPerPage] = useState(25)
  const [totalRecords, setTotalRecords] = useState(0)
  const [error, setError] = useState('')
  const [exporting, setExporting] = useState(false)

  useEffect(() => {
    loadMatchedData()
    loadUnmatchedData()
  }, [])

  const loadMatchedData = async () => {
    setLoading(true)
    try {
      const result = await apiService.getMatchedResults(sessionId, page + 1, rowsPerPage)
      setMatchedData(result.data || [])
      setTotalRecords(result.pagination?.total_records || 0)
    } catch (err) {
      setError('Failed to load matched results')
    } finally {
      setLoading(false)
    }
  }

  const loadUnmatchedData = async () => {
    try {
      const lmrbResult = await apiService.getUnmatchedResults(sessionId, 'lmrb')
      const tcResult = await apiService.getUnmatchedResults(sessionId, 'tc')

      setUnmatchedLMRB(lmrbResult.data || [])
      setUnmatchedTC(tcResult.data || [])
    } catch (err) {
      console.error('Error loading unmatched data:', err)
    }
  }

  const handleExportExcel = async () => {
    setExporting(true)
    try {
      await apiService.exportToExcel(sessionId)
    } catch (err) {
      setError('Failed to export to Excel')
    } finally {
      setExporting(false)
    }
  }

  const handleExportPDF = async () => {
    setExporting(true)
    try {
      await apiService.exportToPDF(sessionId, 'All Channels')
    } catch (err) {
      setError('Failed to export to PDF')
    } finally {
      setExporting(false)
    }
  }

  const handleChangePage = (event, newPage) => {
    setPage(newPage)
  }

  const handleChangeRowsPerPage = (event) => {
    setRowsPerPage(parseInt(event.target.value, 10))
    setPage(0)
  }

  return (
    <Box>
      <Box display="flex" justifyContent="space-between" alignItems="center" mb={3}>
        <Box>
          <Typography variant="h4" fontWeight={600}>
            Reconciliation Results
          </Typography>
          <Typography variant="body1" color="text.secondary">
            View matched and unmatched records
          </Typography>
        </Box>
        <Box display="flex" gap={1}>
          <Button
            variant="outlined"
            startIcon={<DownloadIcon />}
            onClick={handleExportExcel}
            disabled={exporting}
          >
            Export Excel
          </Button>
          <Button
            variant="outlined"
            startIcon={<PdfIcon />}
            onClick={handleExportPDF}
            disabled={exporting}
          >
            Export PDF
          </Button>
        </Box>
      </Box>

      {error && (
        <Alert severity="error" sx={{ mb: 2 }} onClose={() => setError('')}>
          {error}
        </Alert>
      )}

      {exporting && (
        <Alert severity="info" sx={{ mb: 2 }}>
          Exporting data... Your download will start shortly.
        </Alert>
      )}

      {/* Summary Cards */}
      <Grid container spacing={2} mb={3}>
        <Grid item xs={12} sm={4}>
          <Paper sx={{ p: 2, textAlign: 'center', bgcolor: 'success.light' }}>
            <Typography variant="h4" color="success.dark">
              {matchedData.length}
            </Typography>
            <Typography variant="caption" color="success.dark">
              Matched Records
            </Typography>
          </Paper>
        </Grid>
        <Grid item xs={12} sm={4}>
          <Paper sx={{ p: 2, textAlign: 'center', bgcolor: 'warning.light' }}>
            <Typography variant="h4" color="warning.dark">
              {unmatchedLMRB.length}
            </Typography>
            <Typography variant="caption" color="warning.dark">
              Unmatched LMRB
            </Typography>
          </Paper>
        </Grid>
        <Grid item xs={12} sm={4}>
          <Paper sx={{ p: 2, textAlign: 'center', bgcolor: 'error.light' }}>
            <Typography variant="h4" color="error.dark">
              {unmatchedTC.length}
            </Typography>
            <Typography variant="caption" color="error.dark">
              Unmatched TC
            </Typography>
          </Paper>
        </Grid>
      </Grid>

      <Card>
        <CardContent>
          <Tabs value={tab} onChange={(e, v) => setTab(v)} sx={{ mb: 2 }}>
            <Tab label={`Matched (${matchedData.length})`} />
            <Tab label={`Unmatched LMRB (${unmatchedLMRB.length})`} />
            <Tab label={`Unmatched TC (${unmatchedTC.length})`} />
          </Tabs>

          {loading && <LinearProgress sx={{ mb: 2 }} />}

          {/* Matched Records */}
          {tab === 0 && (
            <Box>
              <TableContainer>
                <Table size="small">
                  <TableHead>
                    <TableRow>
                      <TableCell>Date</TableCell>
                      <TableCell>Channel</TableCell>
                      <TableCell>Program</TableCell>
                      <TableCell>LMRB Theme</TableCell>
                      <TableCell>TC Theme</TableCell>
                      <TableCell>Time</TableCell>
                      <TableCell>Duration</TableCell>
                      <TableCell>Status</TableCell>
                    </TableRow>
                  </TableHead>
                  <TableBody>
                    {matchedData.map((row, index) => (
                      <TableRow key={index}>
                        <TableCell>{row.Date || 'N/A'}</TableCell>
                        <TableCell>{row.Channel || 'N/A'}</TableCell>
                        <TableCell>{row.Program || 'N/A'}</TableCell>
                        <TableCell>{row.lmrb_theme || row.Theme || 'N/A'}</TableCell>
                        <TableCell>{row.tc_theme || row.Theme || 'N/A'}</TableCell>
                        <TableCell>{row.Time || 'N/A'}</TableCell>
                        <TableCell>{row.Duration || 'N/A'}s</TableCell>
                        <TableCell>
                          <Chip label="Matched" color="success" size="small" />
                        </TableCell>
                      </TableRow>
                    ))}
                  </TableBody>
                </Table>
              </TableContainer>
              <TablePagination
                component="div"
                count={totalRecords}
                page={page}
                onPageChange={handleChangePage}
                rowsPerPage={rowsPerPage}
                onRowsPerPageChange={handleChangeRowsPerPage}
              />
            </Box>
          )}

          {/* Unmatched LMRB */}
          {tab === 1 && (
            <TableContainer>
              <Table size="small">
                <TableHead>
                  <TableRow>
                    <TableCell>Date</TableCell>
                    <TableCell>Channel</TableCell>
                    <TableCell>Product</TableCell>
                    <TableCell>Theme</TableCell>
                    <TableCell>Program</TableCell>
                    <TableCell>Time</TableCell>
                    <TableCell>Duration</TableCell>
                  </TableRow>
                </TableHead>
                <TableBody>
                  {unmatchedLMRB.map((row, index) => (
                    <TableRow key={index}>
                      <TableCell>{row.Date || 'N/A'}</TableCell>
                      <TableCell>{row.Channel || 'N/A'}</TableCell>
                      <TableCell>{row.Product || 'N/A'}</TableCell>
                      <TableCell>{row.Theme || 'N/A'}</TableCell>
                      <TableCell>{row.Program || 'N/A'}</TableCell>
                      <TableCell>{row.Time || 'N/A'}</TableCell>
                      <TableCell>{row.Duration || 'N/A'}s</TableCell>
                    </TableRow>
                  ))}
                </TableBody>
              </Table>
            </TableContainer>
          )}

          {/* Unmatched TC */}
          {tab === 2 && (
            <TableContainer>
              <Table size="small">
                <TableHead>
                  <TableRow>
                    <TableCell>Date</TableCell>
                    <TableCell>Program</TableCell>
                    <TableCell>Theme</TableCell>
                    <TableCell>Time</TableCell>
                    <TableCell>Duration</TableCell>
                  </TableRow>
                </TableHead>
                <TableBody>
                  {unmatchedTC.map((row, index) => (
                    <TableRow key={index}>
                      <TableCell>{row.Date || 'N/A'}</TableCell>
                      <TableCell>{row.Program || 'N/A'}</TableCell>
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
    </Box>
  )
}
