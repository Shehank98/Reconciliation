import React, { useState, useEffect } from 'react'
import {
  Box,
  Card,
  CardContent,
  Typography,
  Button,
  TextField,
  MenuItem,
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TableHead,
  TableRow,
  Paper,
  IconButton,
  Alert,
  Grid,
  Chip,
} from '@mui/material'
import {
  Delete as DeleteIcon,
  Add as AddIcon,
  Merge as MergeIcon,
} from '@mui/icons-material'
import { useNavigate } from 'react-router-dom'
import apiService from '../services/api'

export default function ThemeMapping({ sessionId }) {
  const navigate = useNavigate()
  const [mappings, setMappings] = useState([])
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState('')

  // Form state
  const [tcTheme, setTcTheme] = useState('')
  const [tcDuration, setTcDuration] = useState('30')
  const [lmrbTheme, setLmrbTheme] = useState('')
  const [lmrbDuration, setLmrbDuration] = useState('30')
  const [mappingType, setMappingType] = useState('COMMERCIAL BENEFITS')

  const durations = ['15', '30', '45', '60', '90', '120']
  const mappingTypes = ['COMMERCIAL BENEFITS', 'SPONSORSHIP']

  useEffect(() => {
    loadMappings()
  }, [])

  const loadMappings = async () => {
    try {
      const result = await apiService.getThemeMappings(sessionId)
      setMappings(result.mappings || [])
    } catch (err) {
      console.error('Error loading mappings:', err)
    }
  }

  const handleAddMapping = async () => {
    if (!tcTheme || !lmrbTheme) {
      setError('Please fill in both TC Theme and LMRB Theme')
      return
    }

    setLoading(true)
    setError('')

    try {
      await apiService.addThemeMapping(sessionId, {
        tc_theme: tcTheme,
        tc_duration: parseInt(tcDuration),
        lmrb_theme: lmrbTheme,
        lmrb_duration: parseInt(lmrbDuration),
        mapping_type: mappingType,
      })

      // Clear form
      setTcTheme('')
      setLmrbTheme('')
      setTcDuration('30')
      setLmrbDuration('30')

      // Reload mappings
      await loadMappings()
    } catch (err) {
      setError(err.response?.data?.error || 'Failed to add mapping')
    } finally {
      setLoading(false)
    }
  }

  const handleDeleteMapping = async (index) => {
    try {
      await apiService.deleteThemeMapping(sessionId, index)
      await loadMappings()
    } catch (err) {
      setError(err.response?.data?.error || 'Failed to delete mapping')
    }
  }

  return (
    <Box>
      <Typography variant="h4" gutterBottom fontWeight={600}>
        Theme Mapping
      </Typography>
      <Typography variant="body1" color="text.secondary" paragraph>
        Define how TC themes map to LMRB themes for reconciliation matching
      </Typography>

      {error && (
        <Alert severity="error" sx={{ mb: 3 }} onClose={() => setError('')}>
          {error}
        </Alert>
      )}

      <Grid container spacing={3}>
        {/* Add New Mapping */}
        <Grid item xs={12} lg={5}>
          <Card>
            <CardContent>
              <Box display="flex" alignItems="center" mb={2}>
                <AddIcon color="primary" sx={{ mr: 1 }} />
                <Typography variant="h6">Add New Mapping</Typography>
              </Box>

              <Grid container spacing={2}>
                <Grid item xs={12}>
                  <TextField
                    select
                    fullWidth
                    label="Mapping Type"
                    value={mappingType}
                    onChange={(e) => setMappingType(e.target.value)}
                  >
                    {mappingTypes.map((type) => (
                      <MenuItem key={type} value={type}>
                        {type}
                      </MenuItem>
                    ))}
                  </TextField>
                </Grid>

                <Grid item xs={12} sm={8}>
                  <TextField
                    fullWidth
                    label="TC Theme"
                    value={tcTheme}
                    onChange={(e) => setTcTheme(e.target.value)}
                    placeholder="e.g., Coca Cola"
                  />
                </Grid>
                <Grid item xs={12} sm={4}>
                  <TextField
                    select
                    fullWidth
                    label="Duration (s)"
                    value={tcDuration}
                    onChange={(e) => setTcDuration(e.target.value)}
                  >
                    {durations.map((d) => (
                      <MenuItem key={d} value={d}>
                        {d}s
                      </MenuItem>
                    ))}
                  </TextField>
                </Grid>

                <Grid item xs={12}>
                  <Box display="flex" justifyContent="center" my={1}>
                    <MergeIcon color="action" />
                  </Box>
                </Grid>

                <Grid item xs={12} sm={8}>
                  <TextField
                    fullWidth
                    label="LMRB Theme"
                    value={lmrbTheme}
                    onChange={(e) => setLmrbTheme(e.target.value)}
                    placeholder="e.g., Beverages"
                  />
                </Grid>
                <Grid item xs={12} sm={4}>
                  <TextField
                    select
                    fullWidth
                    label="Duration (s)"
                    value={lmrbDuration}
                    onChange={(e) => setLmrbDuration(e.target.value)}
                  >
                    {durations.map((d) => (
                      <MenuItem key={d} value={d}>
                        {d}s
                      </MenuItem>
                    ))}
                  </TextField>
                </Grid>

                <Grid item xs={12}>
                  <Button
                    variant="contained"
                    fullWidth
                    size="large"
                    onClick={handleAddMapping}
                    disabled={loading}
                    startIcon={<AddIcon />}
                  >
                    Add Mapping
                  </Button>
                </Grid>
              </Grid>
            </CardContent>
          </Card>

          <Alert severity="info" sx={{ mt: 2 }}>
            <strong>Tip:</strong> You can add multiple TC themes mapping to the same LMRB theme for flexible matching.
          </Alert>
        </Grid>

        {/* Existing Mappings */}
        <Grid item xs={12} lg={7}>
          <Card>
            <CardContent>
              <Box display="flex" alignItems="center" justifyContent="space-between" mb={2}>
                <Typography variant="h6">
                  Current Mappings ({mappings.length})
                </Typography>
              </Box>

              {mappings.length === 0 ? (
                <Alert severity="warning">
                  No theme mappings defined yet. Add at least one mapping to proceed.
                </Alert>
              ) : (
                <TableContainer>
                  <Table>
                    <TableHead>
                      <TableRow>
                        <TableCell>Type</TableCell>
                        <TableCell>TC Theme</TableCell>
                        <TableCell>LMRB Theme</TableCell>
                        <TableCell align="center">Duration</TableCell>
                        <TableCell align="center">Action</TableCell>
                      </TableRow>
                    </TableHead>
                    <TableBody>
                      {mappings.map((mapping, index) => (
                        <TableRow key={index}>
                          <TableCell>
                            <Chip
                              label={mapping.mapping_type}
                              size="small"
                              color={mapping.mapping_type === 'SPONSORSHIP' ? 'secondary' : 'primary'}
                            />
                          </TableCell>
                          <TableCell>{mapping.tc_theme}</TableCell>
                          <TableCell>{mapping.lmrb_theme}</TableCell>
                          <TableCell align="center">
                            {mapping.tc_duration}s / {mapping.lmrb_duration}s
                          </TableCell>
                          <TableCell align="center">
                            <IconButton
                              size="small"
                              color="error"
                              onClick={() => handleDeleteMapping(index)}
                            >
                              <DeleteIcon />
                            </IconButton>
                          </TableCell>
                        </TableRow>
                      ))}
                    </TableBody>
                  </Table>
                </TableContainer>
              )}
            </CardContent>
          </Card>

          {mappings.length > 0 && (
            <Box mt={2}>
              <Button
                variant="contained"
                color="primary"
                size="large"
                fullWidth
                onClick={() => navigate('/reconciliation')}
              >
                Continue to Reconciliation
              </Button>
            </Box>
          )}
        </Grid>
      </Grid>
    </Box>
  )
}
