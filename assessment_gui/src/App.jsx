import { useEffect, useMemo, useState } from 'react'
import './App.css'

const API_BASE = import.meta.env.VITE_API_BASE || ''

const acceptedExtensions = ['.xlsx', '.xlsm', '.csv', '.txt', '.md']
const assessmentModes = [
  {
    id: 'questionnaire',
    label: 'Questionnaire Review',
    detail: 'Strict scoring from completed questionnaire answers, certifications, mappings, and stated processes.',
  },
  {
    id: 'formal_evidence',
    label: 'Formal Evidence Review',
    detail: 'Stricter scoring that expects supporting artifacts such as policies, records, logs, and review evidence.',
  },
]

function isAcceptedFile(file) {
  const name = file.name.toLowerCase()
  return acceptedExtensions.some((ext) => name.endsWith(ext))
}

function formatPercent(value) {
  return typeof value === 'number' ? `${value.toFixed(2)}%` : 'N/A'
}

function App() {
  const [company, setCompany] = useState('')
  const [assessmentMode, setAssessmentMode] = useState('questionnaire')
  const [files, setFiles] = useState([])
  const [jobId, setJobId] = useState('')
  const [status, setStatus] = useState('idle')
  const [stage, setStage] = useState('')
  const [message, setMessage] = useState('')
  const [progress, setProgress] = useState(0)
  const [result, setResult] = useState(null)
  const [downloads, setDownloads] = useState(null)
  const [error, setError] = useState('')
  const [dragging, setDragging] = useState(false)

  const canSubmit = useMemo(
    () => company.trim().length > 0 && files.length > 0 && status !== 'running',
    [company, files.length, status],
  )

  useEffect(() => {
    if (!jobId || status !== 'running') {
      return undefined
    }

    const interval = window.setInterval(async () => {
      try {
        const response = await fetch(`${API_BASE}/api/jobs/${jobId}`)
        if (!response.ok) {
          throw new Error('Could not retrieve job status.')
        }

        const payload = await response.json()
        setProgress(payload.progress || 0)
        setStage(payload.stage || '')
        setMessage(payload.message || '')

        if (payload.status === 'completed') {
          setStatus('completed')
          setResult(payload.result || null)
          setDownloads(payload.downloads || null)
          window.clearInterval(interval)
        }

        if (payload.status === 'failed') {
          setStatus('failed')
          setError(payload.error || 'Assessment failed.')
          window.clearInterval(interval)
        }
      } catch (pollError) {
        setStatus('failed')
        setError(pollError.message || 'Unable to poll job progress.')
        window.clearInterval(interval)
      }
    }, 1500)

    return () => window.clearInterval(interval)
  }, [jobId, status])

  const addFiles = (incoming) => {
    const next = Array.from(incoming || []).filter(isAcceptedFile)
    if (!next.length) {
      return
    }

    setFiles((existing) => {
      const deduped = new Map(existing.map((file) => [file.name + file.size, file]))
      for (const file of next) {
        deduped.set(file.name + file.size, file)
      }
      return Array.from(deduped.values())
    })
  }

  const onDrop = (event) => {
    event.preventDefault()
    setDragging(false)
    addFiles(event.dataTransfer.files)
  }

  const removeFile = (index) => {
    setFiles((existing) => existing.filter((_, idx) => idx !== index))
  }

  const startAssessment = async () => {
    setError('')
    setResult(null)
    setDownloads(null)
    setStatus('running')
    setProgress(2)
    setStage('Preparing submission')
    setMessage('Uploading questionnaire files...')

    try {
      const formData = new FormData()
      formData.append('company', company.trim())
      formData.append('mode', assessmentMode)
      for (const file of files) {
        formData.append('questionnaires', file)
      }

      const response = await fetch(`${API_BASE}/api/assess`, {
        method: 'POST',
        body: formData,
      })

      if (!response.ok) {
        const details = await response.json().catch(() => ({}))
        throw new Error(details.detail || 'Unable to start assessment.')
      }

      const payload = await response.json()
      setJobId(payload.job_id)
      setProgress(5)
      setStage('Assessment started')
      setMessage('Your assessment is now running.')
    } catch (submitError) {
      setStatus('failed')
      setError(submitError.message || 'Unable to start assessment.')
    }
  }

  const openDownload = (path) => {
    if (!path) {
      return
    }
    window.open(`${API_BASE}${path}`, '_blank', 'noopener,noreferrer')
  }

  return (
    <main className="page-shell">
      <section className="hero-card">
        <p className="eyebrow">Security Assessment Portal</p>
        <h1>ISO Evidence Review Workspace</h1>
        <p className="hero-copy">
          Score submitted security documentation in either strict questionnaire mode or formal evidence mode,
          then export an ISO-style evidence sufficiency report.
        </p>
      </section>

      <section className="panel form-panel">
        <label htmlFor="companyName" className="field-label">
          Company Name
        </label>
        <input
          id="companyName"
          className="company-input"
          placeholder="Enter the company name"
          value={company}
          onChange={(event) => setCompany(event.target.value)}
          disabled={status === 'running'}
        />

        <fieldset className="mode-field">
          <legend>Assessment Mode</legend>
          <div className="mode-segments" role="radiogroup" aria-label="Assessment mode">
            {assessmentModes.map((mode) => (
              <label key={mode.id} className={assessmentMode === mode.id ? 'selected' : ''}>
                <input
                  type="radio"
                  name="assessmentMode"
                  value={mode.id}
                  checked={assessmentMode === mode.id}
                  onChange={(event) => setAssessmentMode(event.target.value)}
                  disabled={status === 'running'}
                />
                <span>{mode.label}</span>
                <small>{mode.detail}</small>
              </label>
            ))}
          </div>
        </fieldset>

        <div
          className={`dropzone ${dragging ? 'dragging' : ''}`}
          onDragOver={(event) => {
            event.preventDefault()
            setDragging(true)
          }}
          onDragLeave={() => setDragging(false)}
          onDrop={onDrop}
        >
          <input
            type="file"
            id="questionnaireInput"
            multiple
            accept={acceptedExtensions.join(',')}
            onChange={(event) => addFiles(event.target.files)}
            disabled={status === 'running'}
          />
          <label htmlFor="questionnaireInput" className="dropzone-label">
            <span>Drag and drop questionnaires here</span>
            <small>or click to select files (.xlsx, .xlsm, .csv, .txt, .md)</small>
          </label>
        </div>

        {files.length > 0 && (
          <ul className="file-list">
            {files.map((file, index) => (
              <li key={`${file.name}-${file.size}`}>
                <span>{file.name}</span>
                <button type="button" onClick={() => removeFile(index)} disabled={status === 'running'}>
                  Remove
                </button>
              </li>
            ))}
          </ul>
        )}

        <button className="primary-btn" onClick={startAssessment} disabled={!canSubmit}>
          {status === 'running' ? 'Assessment In Progress' : 'Start Assessment'}
        </button>

        {error && <p className="error-text">{error}</p>}
      </section>

      {(status === 'running' || status === 'completed') && (
        <section className="panel progress-panel">
          <div className="progress-header">
            <h2>Analysis Progress</h2>
            <span>{Math.round(progress)}%</span>
          </div>
          <div
            className="progress-track"
            role="progressbar"
            aria-valuemin="0"
            aria-valuemax="100"
            aria-valuenow={Math.round(progress)}
          >
            <div className="progress-fill" style={{ width: `${Math.max(2, progress)}%` }}></div>
          </div>
          <p className="progress-stage">{stage}</p>
          <p className="progress-message">{message}</p>
        </section>
      )}

      {status === 'completed' && result && (
        <section className="panel results-panel">
          <div className="results-head">
            <div>
              <h2>Assessment Results</h2>
              <p>{result.assessment_mode_label || 'Strict Questionnaire Review'}</p>
            </div>
            <div className="actions">
              <button className="secondary-btn" onClick={() => openDownload(downloads?.assessment_report)}>
                Export Assessment Report
              </button>
              <button className="secondary-btn" onClick={() => openDownload(downloads?.filled_template)}>
                Export Filled Template
              </button>
            </div>
          </div>

          <div className="metric-grid">
            <article>
              <p>Evidence Sufficiency Level</p>
              <strong>{result.metrics?.security_level || 'N/A'}</strong>
            </article>
            <article>
              <p>Evidence Sufficiency Score</p>
              <strong>{formatPercent(result.metrics?.score_percent)}</strong>
            </article>
            <article>
              <p>Letter Grade</p>
              <strong>{result.metrics?.letter_grade || 'N/A'}</strong>
            </article>
            <article>
              <p>Controls Below Target</p>
              <strong>{result.metrics?.controls_below_target ?? 'N/A'}</strong>
            </article>
          </div>

          <div className="summary-columns">
            <div>
              <h3>Section Performance</h3>
              <ul className="summary-list">
                {(result.section_summary || []).map((section) => (
                  <li key={section.Section}>
                    <div>
                      <span>{section.Section}</span>
                      <small>{section['Security Level']}</small>
                    </div>
                    <strong>{formatPercent(section['Score Percent'])}</strong>
                  </li>
                ))}
              </ul>
            </div>

            <div>
              <h3>Priority Findings</h3>
              <ul className="findings-list">
                {(result.top_findings || []).slice(0, 6).map((finding) => (
                  <li key={`${finding['Control ID']}-${finding.Priority}`}>
                    <p>
                      <span>{finding.Priority}</span> {finding['Control ID']} - {finding['Control Name']}
                    </p>
                    <small>{finding['Gap Summary']}</small>
                  </li>
                ))}
              </ul>
            </div>
          </div>
        </section>
      )}
    </main>
  )
}

export default App
