import { useRef, useState } from 'react'
import { ACCEPTED_EXTENSIONS, parseWorkbook } from '../excel/parse.ts'
import { formatBytes } from '../excel/format.ts'
import type { WorkbookData } from '../types.ts'
import { IconFile, IconTable } from './Icons.tsx'

interface Props {
  oldWb: WorkbookData | null
  newWb: WorkbookData | null
  error: string | null
  onLoaded: (side: 'old' | 'new', wb: WorkbookData) => void
  onError: (message: string) => void
  onSample: () => void
}

export function UploadScreen({ oldWb, newWb, error, onLoaded, onError, onSample }: Props) {
  return (
    <div className="upload-page">
      <header className="upload-hero">
        <div className="brand">
          <IconTable className="brand-icon" />
          <span>Diff Excel</span>
        </div>
        <h1>Comparer deux versions d’un classeur</h1>
        <p>
          Tout le calcul se fait dans le navigateur, comme un <code>git diff</code> : feuilles, lignes,
          colonnes et cellules. Aucun fichier n’est envoyé sur un serveur.
        </p>
      </header>

      <div className="drop-grid">
        <FileDrop
          label="Version de base"
          hint="Ancien fichier — HEAD~1"
          side="old"
          workbook={oldWb}
          onLoaded={onLoaded}
          onError={onError}
        />
        <div className="drop-arrow" aria-hidden="true">
          →
        </div>
        <FileDrop
          label="Nouvelle version"
          hint="Fichier à comparer — HEAD"
          side="new"
          workbook={newWb}
          onLoaded={onLoaded}
          onError={onError}
        />
      </div>

      {error && <p className="upload-error">{error}</p>}

      <p className="upload-help">
        Formats : .xlsx, .xls, .xlsm, .ods, .csv — ou{' '}
        <button type="button" className="linkish" onClick={onSample}>
          charger un exemple
        </button>
      </p>
    </div>
  )
}

interface DropProps {
  label: string
  hint: string
  side: 'old' | 'new'
  workbook: WorkbookData | null
  onLoaded: (side: 'old' | 'new', wb: WorkbookData) => void
  onError: (message: string) => void
}

function FileDrop({ label, hint, side, workbook, onLoaded, onError }: DropProps) {
  const inputRef = useRef<HTMLInputElement>(null)
  const [over, setOver] = useState(false)
  const [busy, setBusy] = useState(false)

  async function handleFile(file: File | undefined) {
    if (!file) return
    setBusy(true)
    try {
      const wb = await parseWorkbook(file)
      onLoaded(side, wb)
    } catch (err) {
      onError(err instanceof Error ? err.message : 'Impossible de lire ce fichier.')
    } finally {
      setBusy(false)
    }
  }

  return (
    <button
      type="button"
      className={`drop-card ${over ? 'is-over' : ''} ${workbook ? 'has-file' : ''} ${side}`}
      onClick={() => inputRef.current?.click()}
      onDragOver={(e) => {
        e.preventDefault()
        setOver(true)
      }}
      onDragLeave={() => setOver(false)}
      onDrop={(e) => {
        e.preventDefault()
        setOver(false)
        void handleFile(e.dataTransfer.files[0])
      }}
    >
      <input
        ref={inputRef}
        type="file"
        accept={ACCEPTED_EXTENSIONS}
        hidden
        onChange={(e) => {
          void handleFile(e.target.files?.[0])
          e.target.value = ''
        }}
      />
      <span className="drop-kicker">{label}</span>
      {busy ? (
        <span className="drop-status">Lecture…</span>
      ) : workbook ? (
        <span className="drop-file">
          <IconFile />
          <span>
            <strong>{workbook.fileName}</strong>
            <small>
              {workbook.sheets.length} feuille{workbook.sheets.length > 1 ? 's' : ''} ·{' '}
              {formatBytes(workbook.fileSize)}
            </small>
          </span>
        </span>
      ) : (
        <span className="drop-placeholder">
          Déposer un fichier
          <small>{hint}</small>
        </span>
      )}
    </button>
  )
}
