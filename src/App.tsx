import { useEffect, useMemo, useState } from 'react'
import { DiffWorkspace } from './components/DiffWorkspace.tsx'
import { UploadScreen } from './components/UploadScreen.tsx'
import { defaultOptions, diffWorkbooks } from './excel/diff.ts'
import { downloadText, exportDiffCsv } from './excel/export.ts'
import {
  acceptAll,
  downloadMergedXlsx,
  emptyDecisions,
  setCellDecision,
  setColumnDecision,
  setHunkDecision,
  setRowDecision,
  setSheetDecision,
} from './excel/merge.ts'
import { createSampleWorkbooks } from './excel/sample.ts'
import type { DiffOptions, MergeDecisions, Resolution, RowDiff, Theme, ViewMode, WorkbookData } from './types.ts'

export default function App() {
  const [oldWb, setOldWb] = useState<WorkbookData | null>(null)
  const [newWb, setNewWb] = useState<WorkbookData | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [options, setOptions] = useState<DiffOptions>(defaultOptions)
  const [view, setView] = useState<ViewMode>('split')
  const [showAll, setShowAll] = useState(false)
  const [showFormulas, setShowFormulas] = useState(false)
  const [theme, setTheme] = useState<Theme>(() => {
    const saved = localStorage.getItem('xl-diff-theme')
    return saved === 'light' || saved === 'dark' ? saved : 'dark'
  })
  const [selected, setSelected] = useState<string | null>(null)
  const [changeIndex, setChangeIndex] = useState(0)
  const [decisions, setDecisions] = useState<MergeDecisions>(emptyDecisions)
  const [saving, setSaving] = useState(false)

  const diff = useMemo(() => {
    if (!oldWb || !newWb) return null
    return diffWorkbooks(oldWb, newWb, options)
  }, [oldWb, newWb, options])

  const sheetName = selected && diff?.sheets.some((s) => s.name === selected) ? selected : diff?.sheets[0]?.name

  const changeCount = useMemo(() => {
    if (!diff || !sheetName) return 0
    const sheet = diff.sheets.find((s) => s.name === sheetName)
    if (!sheet) return 0
    const rows = options.headerRow ? sheet.allRows.slice(1) : sheet.allRows
    return rows.filter((r) => r.kind !== 'context').length
  }, [diff, sheetName, options.headerRow])

  useEffect(() => {
    document.documentElement.dataset.theme = theme
    localStorage.setItem('xl-diff-theme', theme)
  }, [theme])

  useEffect(() => {
    const prevent = (e: DragEvent) => {
      e.preventDefault()
    }
    window.addEventListener('dragover', prevent)
    window.addEventListener('drop', prevent)
    return () => {
      window.removeEventListener('dragover', prevent)
      window.removeEventListener('drop', prevent)
    }
  }, [])

  useEffect(() => {
    setChangeIndex(0)
    setDecisions(emptyDecisions())
  }, [oldWb, newWb])

  useEffect(() => {
    setChangeIndex(0)
  }, [sheetName])

  useEffect(() => {
    function onKey(e: KeyboardEvent) {
      if (!diff) return
      const tag = (e.target as HTMLElement | null)?.tagName
      if (tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT') return
      if (e.key === 'n' || e.key === 'j') {
        e.preventDefault()
        setChangeIndex((i) => i + 1)
      } else if (e.key === 'p' || e.key === 'k') {
        e.preventDefault()
        setChangeIndex((i) => i - 1)
      } else if (e.key === '1') setView('unified')
      else if (e.key === '2') setView('split')
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [diff])

  function handleLoaded(side: 'old' | 'new', wb: WorkbookData) {
    setError(null)
    if (side === 'old') setOldWb(wb)
    else setNewWb(wb)
  }

  function loadSample() {
    const sample = createSampleWorkbooks()
    setError(null)
    setOldWb(sample.oldWb)
    setNewWb(sample.newWb)
    setSelected('Ventes')
  }

  useEffect(() => {
    if (new URLSearchParams(window.location.search).get('demo') === '1') {
      loadSample()
    }
  }, [])

  function reset() {
    setOldWb(null)
    setNewWb(null)
    setSelected(null)
    setError(null)
    setChangeIndex(0)
    setDecisions(emptyDecisions())
  }

  function withSheet(update: (sheetName: string) => void) {
    if (!sheetName) return
    update(sheetName)
  }

  if (!diff || !sheetName) {
    return (
      <UploadScreen
        oldWb={oldWb}
        newWb={newWb}
        error={error}
        onLoaded={handleLoaded}
        onError={setError}
        onSample={loadSample}
      />
    )
  }

  return (
    <DiffWorkspace
      diff={diff}
      selected={sheetName}
      onSelect={setSelected}
      view={view}
      onView={setView}
      showAll={showAll}
      onShowAll={setShowAll}
      showFormulas={showFormulas}
      onShowFormulas={setShowFormulas}
      options={options}
      onOptions={setOptions}
      theme={theme}
      onTheme={setTheme}
      changeIndex={changeCount === 0 ? 0 : ((changeIndex % changeCount) + changeCount) % changeCount}
      changeCount={changeCount}
      onPrev={() => setChangeIndex((i) => i - 1)}
      onNext={() => setChangeIndex((i) => i + 1)}
      onSwap={() => {
        setOldWb(newWb)
        setNewWb(oldWb)
      }}
      onReset={reset}
      onExport={() => {
        const csv = exportDiffCsv(diff)
        const base = diff.newName.replace(/\.[^.]+$/, '')
        downloadText(`${base}-diff.csv`, csv, 'text/csv;charset=utf-8')
      }}
      decisions={decisions}
      saving={saving}
      error={error}
      onAcceptAll={(side) => setDecisions(acceptAll(diff, side))}
      onSaveMerged={() => {
        if (!oldWb || !newWb) return
        void (async () => {
          setSaving(true)
          setError(null)
          try {
            await downloadMergedXlsx(diff, decisions, options.headerRow, oldWb, newWb)
          } catch (err) {
            setError(err instanceof Error ? err.message : 'Impossible de générer le fichier fusionné')
          } finally {
            setSaving(false)
          }
        })()
      }}
      onRowDecision={(row: RowDiff, value: Resolution) =>
        withSheet((name) => setDecisions((prev) => setRowDecision(prev, name, row, value)))
      }
      onCellDecision={(row, letter, value) =>
        withSheet((name) => setDecisions((prev) => setCellDecision(prev, name, row, letter, value)))
      }
      onColumnDecision={(letter, value) =>
        withSheet((name) => setDecisions((prev) => setColumnDecision(prev, name, letter, value)))
      }
      onHunkDecision={(rows, value) =>
        withSheet((name) => setDecisions((prev) => setHunkDecision(prev, name, rows, value)))
      }
      onSheetDecision={(value) =>
        withSheet((name) => setDecisions((prev) => setSheetDecision(prev, name, value)))
      }
    />
  )
}
