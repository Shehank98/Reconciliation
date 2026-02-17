import axios from 'axios'

const API_BASE_URL = import.meta.env.VITE_API_URL || 'http://localhost:5000/api'

const api = axios.create({
  baseURL: API_BASE_URL,
  headers: {
    'Content-Type': 'application/json',
  },
})

// API Service
const apiService = {
  // Health Check
  healthCheck: async () => {
    const response = await api.get('/health')
    return response.data
  },

  // Google Sheets
  connectGoogleSheet: async (sheetId, sessionId) => {
    const response = await api.post('/google-sheets/connect', {
      sheet_id: sheetId,
      session_id: sessionId,
    })
    return response.data
  },

  getLMRBData: async (sessionId, startDate, endDate, channel) => {
    const response = await api.post('/google-sheets/lmrb-data', {
      session_id: sessionId,
      start_date: startDate,
      end_date: endDate,
      channel: channel,
    })
    return response.data
  },

  // TC Upload
  uploadTCPDF: async (file, sessionId) => {
    const formData = new FormData()
    formData.append('file', file)
    formData.append('session_id', sessionId)

    const response = await api.post('/upload-tc', formData, {
      headers: {
        'Content-Type': 'multipart/form-data',
      },
    })
    return response.data
  },

  // Theme Mappings
  getThemeMappings: async (sessionId) => {
    const response = await api.get('/theme-mappings', {
      params: { session_id: sessionId },
    })
    return response.data
  },

  addThemeMapping: async (sessionId, mapping) => {
    const response = await api.post(
      '/theme-mappings',
      mapping,
      { params: { session_id: sessionId } }
    )
    return response.data
  },

  deleteThemeMapping: async (sessionId, index) => {
    const response = await api.delete('/theme-mappings', {
      params: { session_id: sessionId },
      data: { index },
    })
    return response.data
  },

  // Reconciliation
  runReconciliation: async (sessionId, timeToler ance, ignoreDate, products) => {
    const response = await api.post('/reconcile', {
      session_id: sessionId,
      time_tolerance: timeTolerance,
      ignore_date: ignoreDate,
      products: products,
    })
    return response.data
  },

  // Results
  getMatchedResults: async (sessionId, page = 1, perPage = 50) => {
    const response = await api.get('/results/matched', {
      params: {
        session_id: sessionId,
        page,
        per_page: perPage,
      },
    })
    return response.data
  },

  getUnmatchedResults: async (sessionId, type = 'lmrb') => {
    const response = await api.get('/results/unmatched', {
      params: {
        session_id: sessionId,
        type,
      },
    })
    return response.data
  },

  // Export
  exportToExcel: async (sessionId) => {
    const response = await api.post(
      '/export/excel',
      { session_id: sessionId },
      { responseType: 'blob' }
    )

    // Create download link
    const url = window.URL.createObjectURL(new Blob([response.data]))
    const link = document.createElement('a')
    link.href = url
    link.setAttribute('download', `reconciliation_${Date.now()}.xlsx`)
    document.body.appendChild(link)
    link.click()
    link.remove()

    return { success: true }
  },

  exportToPDF: async (sessionId, channel) => {
    const response = await api.post(
      '/export/pdf',
      { session_id: sessionId, channel },
      { responseType: 'blob' }
    )

    // Create download link
    const url = window.URL.createObjectURL(new Blob([response.data]))
    const link = document.createElement('a')
    link.href = url
    link.setAttribute('download', `reconciliation_summary_${Date.now()}.pdf`)
    document.body.appendChild(link)
    link.click()
    link.remove()

    return { success: true }
  },

  // Manual Matching
  manualMatch: async (sessionId, lmrbIndex, tcIndex) => {
    const response = await api.post('/manual-match', {
      session_id: sessionId,
      lmrb_index: lmrbIndex,
      tc_index: tcIndex,
    })
    return response.data
  },
}

export default apiService
