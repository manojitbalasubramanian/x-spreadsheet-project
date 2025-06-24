import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import SpreadsheetApp from './components/SpreadsheetApp.jsx'

createRoot(document.getElementById('root')).render(
  <StrictMode>
    <SpreadsheetApp /> 
  </StrictMode>,
)
