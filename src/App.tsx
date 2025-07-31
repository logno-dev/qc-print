import { useState, useCallback } from 'react'
import * as XLSX from 'xlsx'

interface ExcelRow {
  [key: string]: string | number | boolean | null | undefined
}

interface ProcessedData {
  sequenceNumber: number
  columnB: string
  columnC: string
}

function App() {
  const [file, setFile] = useState<File | null>(null)
  const [processing, setProcessing] = useState(false)
  const [processedData, setProcessedData] = useState<ProcessedData[]>([])
  const [showPrintView, setShowPrintView] = useState(false)

  const processExcelFile = useCallback(async (file: File): Promise<ProcessedData[]> => {
    return new Promise((resolve) => {
      const reader = new FileReader()
      reader.onload = (e) => {
        const data = new Uint8Array(e.target?.result as ArrayBuffer)
        const workbook = XLSX.read(data, { type: 'array' })
        const worksheet = workbook.Sheets[workbook.SheetNames[0]]
        const jsonData: ExcelRow[] = XLSX.utils.sheet_to_json(worksheet, { header: 1 })

        const processedRows: ProcessedData[] = []
        
        jsonData.forEach((row, index) => {
          if (index === 0) return
          
          const columnB = row[1] ? String(row[1]).trim() : ''
          const columnC = row[2] ? String(row[2]).trim() : ''
          
          if (columnB || columnC) {
            processedRows.push({
              sequenceNumber: 0, // Will be assigned after sorting
              columnB,
              columnC
            })
          }
        })

        // Sort alphabetically by column C first, then by column B
        processedRows.sort((a, b) => {
          const columnCCompare = a.columnC.localeCompare(b.columnC)
          if (columnCCompare !== 0) {
            return columnCCompare
          }
          return a.columnB.localeCompare(b.columnB)
        })

        // Assign sequence numbers after sorting
        processedRows.forEach((row, index) => {
          row.sequenceNumber = index + 1
        })

        resolve(processedRows)
      }
      reader.readAsArrayBuffer(file)
    })
  }, [])

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault()
    const droppedFiles = Array.from(e.dataTransfer.files).filter(
      file => file.name.endsWith('.xlsx') || file.name.endsWith('.xls')
    )
    if (droppedFiles.length > 0) {
      setFile(droppedFiles[0])
    }
  }, [])

  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault()
  }, [])

  const processFile = useCallback(async () => {
    if (!file) return
    
    setProcessing(true)
    try {
      const data = await processExcelFile(file)
      setProcessedData(data)
      setShowPrintView(true)
    } catch (error) {
      console.error('Error processing file:', error)
    }
    setProcessing(false)
  }, [file, processExcelFile])

  const handlePrint = () => {
    window.print()
  }

  const chunkData = (data: ProcessedData[], chunkSize: number) => {
    const chunks = []
    for (let i = 0; i < data.length; i += chunkSize) {
      chunks.push(data.slice(i, i + chunkSize))
    }
    return chunks
  }

  const renderPrintView = () => {
    const pages = chunkData(processedData, 100)
    
    return (
      <div className="print-view">
        {pages.map((pageData, pageIndex) => {
          const leftColumn = pageData.slice(0, 50)
          const rightColumn = pageData.slice(50, 100)
          
          return (
            <div key={pageIndex} className={`page ${pageIndex > 0 ? 'page-break' : ''}`}>
              <div className="page-content">
                <div className="table-container">
                  <table className="print-table">
                    <thead>
                      <tr>
                        <th className="col-number">#</th>
                        <th className="col-data">Column B</th>
                        <th className="col-data">Column C</th>
                      </tr>
                    </thead>
                    <tbody>
                      {leftColumn.map((row) => (
                        <tr key={row.sequenceNumber}>
                          <td className="col-number">{row.sequenceNumber}</td>
                          <td className="col-data" title={row.columnB}>{row.columnB}</td>
                          <td className="col-data" title={row.columnC}>{row.columnC}</td>
                        </tr>
                      ))}
                      {Array.from({ length: Math.max(0, 50 - leftColumn.length) }).map((_, index) => (
                        <tr key={`empty-left-${index}`}>
                          <td className="col-number">&nbsp;</td>
                          <td className="col-data">&nbsp;</td>
                          <td className="col-data">&nbsp;</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                
                <div className="table-container">
                  <table className="print-table">
                    <thead>
                      <tr>
                        <th className="col-number">#</th>
                        <th className="col-data">Column B</th>
                        <th className="col-data">Column C</th>
                      </tr>
                    </thead>
                    <tbody>
                      {rightColumn.map((row) => (
                        <tr key={row.sequenceNumber}>
                          <td className="col-number">{row.sequenceNumber}</td>
                          <td className="col-data" title={row.columnB}>{row.columnB}</td>
                          <td className="col-data" title={row.columnC}>{row.columnC}</td>
                        </tr>
                      ))}
                      {Array.from({ length: Math.max(0, 50 - rightColumn.length) }).map((_, index) => (
                        <tr key={`empty-right-${index}`}>
                          <td className="col-number">&nbsp;</td>
                          <td className="col-data">&nbsp;</td>
                          <td className="col-data">&nbsp;</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )
        })}
      </div>
    )
  }

  if (showPrintView) {
    return (
      <div>
        <div className="no-print bg-white p-4 shadow-md">
          <div className="max-w-4xl mx-auto flex justify-between items-center">
            <h1 className="text-2xl font-bold text-gray-900">QC Print Preview</h1>
            <div className="space-x-4">
              <button
                onClick={handlePrint}
                className="bg-blue-600 hover:bg-blue-700 text-white font-semibold py-2 px-4 rounded-lg transition-colors duration-200"
              >
                Print
              </button>
              <button
                onClick={() => setShowPrintView(false)}
                className="bg-gray-600 hover:bg-gray-700 text-white font-semibold py-2 px-4 rounded-lg transition-colors duration-200"
              >
                Back to Upload
              </button>
            </div>
          </div>
        </div>
        {renderPrintView()}
      </div>
    )
  }

  return (
    <div className="min-h-screen bg-gray-50 py-8">
      <div className="max-w-4xl mx-auto px-4 text-center">
        <h1 className="text-4xl font-bold text-gray-900 mb-8">QC Print Generator</h1>

        <div
          className="border-2 border-dashed border-gray-300 rounded-lg p-12 mb-8 bg-white hover:border-blue-400 hover:bg-blue-50 transition-colors duration-300"
          onDrop={handleDrop}
          onDragOver={handleDragOver}
        >
          <p className="text-gray-600 mb-2">Drag and drop Excel file here</p>
          <p className="text-gray-500 mb-4">or</p>
          <input
            type="file"
            accept=".xlsx,.xls"
            className="block mx-auto text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-full file:border-0 file:text-sm file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100"
            onChange={(e) => {
              const selectedFile = e.target.files?.[0]
              if (selectedFile) {
                setFile(selectedFile)
              }
            }}
          />
        </div>

        {file && (
          <div className="bg-white rounded-lg shadow-md p-6 mb-8">
            <h3 className="text-xl font-semibold text-gray-800 mb-4">Selected File:</h3>
            <p className="text-gray-700 py-2 px-4 bg-gray-50 rounded border mb-6">
              {file.name}
            </p>
            <div className="space-x-4">
              <button
                onClick={processFile}
                disabled={processing}
                className="bg-blue-600 hover:bg-blue-700 disabled:bg-gray-400 text-white font-semibold py-3 px-6 rounded-lg transition-colors duration-200"
              >
                {processing ? 'Processing...' : 'Process & Preview'}
              </button>
              <button
                onClick={() => setFile(null)}
                className="bg-red-600 hover:bg-red-700 text-white font-semibold py-3 px-6 rounded-lg transition-colors duration-200"
              >
                Clear File
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  )
}

export default App
