import React, { useState } from 'react'
import {
  Box,
  Card,
  CardContent,
  Typography,
  Button,
  TextField,
  FormControlLabel,
  Switch,
  Alert,
  LinearProgress,
  Grid,
  Paper,
} from '@mui/material'
import {
  PlayArrow as RunIcon,
  CheckCircle as SuccessIcon,
} from '@mui/icons-material'
import { useNavigate } from 'react-router-dom'
import apiService from '../services/api'

export default function Reconciliation({ sessionId }) {
  const navigate = useNavigate()
  const [timeTolerance, setTimeTolerance] = useState(30)
  const [ignoreDate, setIgnoreDate] = useState(false)
  const [loading, setLoading] = useState(false)
  const [results, setResults] = useState(null)
  const [error, setError] = useState('')

  const handleRunReconciliation = async () => {
    setLoading(true)
    setError('')

    try {
      const result = await apiService.runReconciliation(
        sessionId,
        timeTolerance,
        ignoreDate,
        [] // Products filter can be added if needed
      )

      setResults(result)
    } catch (err) {
      setError(err.response?.data?.error || 'Failed to run reconciliation')
    } finally {
      setLoading(false)
    }
  }

  return (
    <Box>
      <Typography variant="h4" gutterBottom fontWeight={600}>
        Run Reconciliation
      </Typography>
      <Typography variant="body1" color="text.secondary" paragraph>
        Configure matching parameters and run the reconciliation process
      </Typography>

      {error && (
        <Alert severity="error" sx={{ mb: 3 }} onClose={() => setError('')}>
          {error}
        </Alert>
      )}

      <Grid container spacing={3}>
        {/* Configuration */}
        <Grid item xs={12} md={6}>
          <Card>
            <CardContent>
              <Typography variant="h6" gutterBottom>
                Matching Parameters
              </Typography>

              <Box mb={3}>
                <Typography variant="subtitle2" gutterBottom>
                  Time Tolerance (seconds)
                </Typography>
                <Typography variant="caption" color="text.secondary" paragraph>
                  Maximum time difference allowed for matching records
                </Typography>
                <TextField
                  type="number"
                  fullWidth
                  value={timeTolerance}
                  onChange={(e) => setTimeTolerance(parseInt(e.target.value) || 30)}
                  inputProps={{ min: 0, max: 300 }}
                />
              </Box>

              <Box mb={3}>
                <FormControlLabel
                  control={
                    <Switch
                      checked={ignoreDate}
                      onChange={(e) => setIgnoreDate(e.target.checked)}
                    />
                  }
                  label="Ignore Date in Matching"
                />
                <Typography variant="caption" display="block" color="text.secondary">
                  Enable this to match records regardless of date (only time and theme)
                </Typography>
              </Box>

              <Button
                variant="contained"
                size="large"
                fullWidth
                onClick={handleRunReconciliation}
                disabled={loading}
                startIcon={<RunIcon />}
              >
                {loading ? 'Running Reconciliation...' : 'Run Reconciliation'}
              </Button>

              {loading && <LinearProgress sx={{ mt: 2 }} />}
            </CardContent>
          </Card>
        </Grid>

        {/* Results Summary */}
        <Grid item xs={12} md={6}>
          <Card>
            <CardContent>
              <Typography variant="h6" gutterBottom>
                Results Summary
              </Typography>

              {!results ? (
                <Alert severity="info">
                  Run reconciliation to see results
                </Alert>
              ) : (
                <Box>
                  <Alert severity="success" icon={<SuccessIcon />} sx={{ mb: 2 }}>
                    Reconciliation completed successfully!
                  </Alert>

                  <Grid container spacing={2}>
                    <Grid item xs={6}>
                      <Paper variant="outlined" sx={{ p: 2, textAlign: 'center' }}>
                        <Typography variant="h4" color="primary">
                          {results.statistics.total_lmrb_records}
                        </Typography>
                        <Typography variant="caption" color="text.secondary">
                          Total LMRB Records
                        </Typography>
                      </Paper>
                    </Grid>
                    <Grid item xs={6}>
                      <Paper variant="outlined" sx={{ p: 2, textAlign: 'center' }}>
                        <Typography variant="h4" color="primary">
                          {results.statistics.total_tc_records}
                        </Typography>
                        <Typography variant="caption" color="text.secondary">
                          Total TC Records
                        </Typography>
                      </Paper>
                    </Grid>
                    <Grid item xs={6}>
                      <Paper variant="outlined" sx={{ p: 2, textAlign: 'center', bgcolor: 'success.light' }}>
                        <Typography variant="h4" color="success.dark">
                          {results.statistics.matched_records}
                        </Typography>
                        <Typography variant="caption" color="success.dark">
                          Matched Records
                        </Typography>
                      </Paper>
                    </Grid>
                    <Grid item xs={6}>
                      <Paper variant="outlined" sx={{ p: 2, textAlign: 'center', bgcolor: 'warning.light' }}>
                        <Typography variant="h4" color="warning.dark">
                          {results.statistics.unmatched_lmrb}
                        </Typography>
                        <Typography variant="caption" color="warning.dark">
                          Unmatched LMRB
                        </Typography>
                      </Paper>
                    </Grid>
                    <Grid item xs={12}>
                      <Paper variant="outlined" sx={{ p: 2, textAlign: 'center' }}>
                        <Typography variant="h3" color="primary">
                          {results.statistics.match_rate}%
                        </Typography>
                        <Typography variant="caption" color="text.secondary">
                          Match Rate
                        </Typography>
                      </Paper>
                    </Grid>
                  </Grid>
                </Box>
              )}
            </CardContent>
          </Card>

          {results && (
            <Box mt={2}>
              <Button
                variant="contained"
                color="primary"
                size="large"
                fullWidth
                onClick={() => navigate('/results')}
              >
                View Detailed Results
              </Button>
            </Box>
          )}
        </Grid>
      </Grid>
    </Box>
  )
}
