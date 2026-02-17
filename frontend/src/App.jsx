import React, { useState } from 'react'
import { ThemeProvider, createTheme } from '@mui/material/styles'
import CssBaseline from '@mui/material/CssBaseline'
import { BrowserRouter as Router, Routes, Route, Navigate } from 'react-router-dom'
import Layout from './components/Layout'
import Dashboard from './pages/Dashboard'
import Upload from './pages/Upload'
import ThemeMapping from './pages/ThemeMapping'
import Reconciliation from './pages/Reconciliation'
import Results from './pages/Results'
import ManualMatching from './pages/ManualMatching'

// Create modern theme
const theme = createTheme({
  palette: {
    mode: 'light',
    primary: {
      main: '#1976d2',
      light: '#42a5f5',
      dark: '#1565c0',
    },
    secondary: {
      main: '#dc004e',
      light: '#e33371',
      dark: '#9a0036',
    },
    background: {
      default: '#f5f5f5',
      paper: '#ffffff',
    },
  },
  typography: {
    fontFamily: 'Roboto, Arial, sans-serif',
    h4: {
      fontWeight: 600,
    },
    h5: {
      fontWeight: 500,
    },
  },
  shape: {
    borderRadius: 8,
  },
  components: {
    MuiButton: {
      styleOverrides: {
        root: {
          textTransform: 'none',
          fontWeight: 500,
        },
      },
    },
    MuiCard: {
      styleOverrides: {
        root: {
          boxShadow: '0 2px 8px rgba(0,0,0,0.1)',
        },
      },
    },
  },
})

function App() {
  const [sessionId] = useState(`session_${Date.now()}`)

  return (
    <ThemeProvider theme={theme}>
      <CssBaseline />
      <Router>
        <Layout sessionId={sessionId}>
          <Routes>
            <Route path="/" element={<Navigate to="/dashboard" replace />} />
            <Route path="/dashboard" element={<Dashboard sessionId={sessionId} />} />
            <Route path="/upload" element={<Upload sessionId={sessionId} />} />
            <Route path="/theme-mapping" element={<ThemeMapping sessionId={sessionId} />} />
            <Route path="/reconciliation" element={<Reconciliation sessionId={sessionId} />} />
            <Route path="/results" element={<Results sessionId={sessionId} />} />
            <Route path="/manual-matching" element={<ManualMatching sessionId={sessionId} />} />
          </Routes>
        </Layout>
      </Router>
    </ThemeProvider>
  )
}

export default App
