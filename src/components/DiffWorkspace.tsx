import { useMemo, useState } from 'react'
import type {
  DiffOptions,
  MergeDecisions,
  MergeSide,
  Resolution,
  RowDiff,
  SheetDiff,
  Theme,
  ViewMode,
  WorkbookDiff,
} from '../types.ts'
import { countPending } from '../excel/merge.ts'
import {
  IconArrowDown,
  IconArrowUp,
  IconDownload,
  IconGear,
  IconMerge,
  IconMoon,
  IconReset,
  IconSun,
  IconSwap,
  IconTable,
} from './Icons.tsx'
import { DiffView } from './DiffView.tsx'
import { ReviewView } from './ReviewView.tsx'

interface Props {
  diff: WorkbookDiff
  selected: string
  onSelect: (name: string) => void
  view: ViewMode
  onView: (view: ViewMode) => void
  showAll: boolean
  onShowAll: (value: boolean) => void
  showFormulas: boolean
  onShowFormulas: (value: boolean) => void
  options: DiffOptions
  onOptions: (next: DiffOptions) => void
  theme: Theme
  onTheme: (theme: Theme) => void
  changeIndex: number
  changeCount: number
  onPrev: () => void
  onNext: () => void
  onSwap: () => void
  onReset: () => void
  onExport: () => void
  onExportHighlight: () => void
  highlighting?: boolean
  decisions: MergeDecisions
  onAcceptAll: (side: MergeSide) => void
  onSaveMerged: () => void
  saving?: boolean
  error?: string | null
  onRowDecision: (sheetName: string, row: RowDiff, value: Resolution) => void
  onCellDecision: (sheetName: string, row: RowDiff, letter: string, value: Resolution) => void
  onColumnDecision: (sheetName: string, letter: string, value: Resolution) => void
  onHunkDecision: (sheetName: string, rows: RowDiff[], value: Resolution) => void
  onSheetDecision: (sheetName: string, value: Resolution) => void
}

export function DiffWorkspace({
  diff,
  selected,
  onSelect,
  view,
  onView,
  showAll,
  onShowAll,
  showFormulas,
  onShowFormulas,
  options,
  onOptions,
  theme,
  onTheme,
  changeIndex,
  changeCount,
  onPrev,
  onNext,
  onSwap,
  onReset,
  onExport,
  onExportHighlight,
  highlighting = false,
  decisions,
  onAcceptAll,
  onSaveMerged,
  saving = false,
  error = null,
  onRowDecision,
  onCellDecision,
  onColumnDecision,
  onHunkDecision,
  onSheetDecision,
}: Props) {
  const [screen, setScreen] = useState<'review' | 'grid'>('review')
  const sheet = diff.sheets.find((s) => s.name === selected) ?? diff.sheets[0]
  const ordered = useMemo(() => orderSheets(diff.sheets), [diff.sheets])
  const pending = useMemo(
    () => countPending(diff, decisions, options.headerRow),
    [diff, decisions, options.headerRow],
  )

  if (!sheet) return null

  return (
    <div className="workspace">
      <header className="topbar">
        <div className="brand">
          <IconTable className="brand-icon" />
          <span>Diff Excel</span>
        </div>
        <div className="compare-names">
          <span className="file-name" title={diff.oldName}>
            {diff.oldName}
          </span>
          <button type="button" className="icon-btn" onClick={onSwap} title="Inverser les fichiers">
            <IconSwap />
          </button>
          <span className="file-name" title={diff.newName}>
            {diff.newName}
          </span>
        </div>
        <Stats stats={diff.stats} />
        <div className="topbar-actions">
          <button
            type="button"
            className="text-btn"
            onClick={onExportHighlight}
            disabled={highlighting || saving}
            title="Garde la mise en forme d’origine. Vert = ajouté, rouge = supprimé. Une colonne modifiée est dédoublée : supprimez celle que vous ne gardez pas."
          >
            {highlighting ? 'Génération…' : 'Comparatif coloré'}
          </button>
          <button type="button" className="icon-btn" onClick={onExport} title="Exporter le CSV des changements">
            <IconDownload />
          </button>
          <button
            type="button"
            className="icon-btn"
            onClick={() => onTheme(theme === 'dark' ? 'light' : 'dark')}
            title="Thème"
          >
            {theme === 'dark' ? <IconSun /> : <IconMoon />}
          </button>
          <button type="button" className="icon-btn" onClick={onReset} title="Nouveau comparatif">
            <IconReset />
          </button>
        </div>
      </header>

      <aside className="sidebar">
        <div className="sidebar-title">Feuilles</div>
        <ul className="sheet-list">
          {ordered.map((item) => (
            <li key={item.name}>
              <button
                type="button"
                className={`sheet-item ${item.name === sheet.name ? 'is-active' : ''} status-${item.status} is-${decisions.sheets[item.name] ?? 'pending'}`}
                onClick={() => onSelect(item.name)}
              >
                <span className="sheet-status">{statusLetter(item.status)}</span>
                <span className="sheet-name">{item.name}</span>
                <span className="sheet-counts">
                  {item.stats.added > 0 && <span className="add">+{item.stats.added}</span>}
                  {item.stats.removed > 0 && <span className="del">−{item.stats.removed}</span>}
                  {item.stats.changed > 0 && <span className="chg">~{item.stats.changed}</span>}
                </span>
              </button>
            </li>
          ))}
        </ul>
      </aside>

      <div className="toolbar">
        <div className="seg">
          <button type="button" className={screen === 'review' ? 'is-on' : ''} onClick={() => setScreen('review')}>
            Revue
          </button>
          <button type="button" className={screen === 'grid' ? 'is-on' : ''} onClick={() => setScreen('grid')}>
            Grille
          </button>
        </div>
        {screen === 'grid' ? (
          <>
            <div className="seg">
              <button type="button" className={view === 'unified' ? 'is-on' : ''} onClick={() => onView('unified')}>
                Unifié
              </button>
              <button type="button" className={view === 'split' ? 'is-on' : ''} onClick={() => onView('split')}>
                Côte à côte
              </button>
            </div>
            <label className="check">
              <input type="checkbox" checked={showAll} onChange={(e) => onShowAll(e.target.checked)} />
              Fichier entier
            </label>
            <label className="check">
              <input type="checkbox" checked={showFormulas} onChange={(e) => onShowFormulas(e.target.checked)} />
              Formules
            </label>
          </>
        ) : null}
        {screen === 'grid' ? (
        <div className="nav-changes">
          <button type="button" className="icon-btn" onClick={onPrev} disabled={changeCount === 0} title="Changement précédent (p)">
            <IconArrowUp />
          </button>
          <span>
            {changeCount === 0 ? '0' : `${(changeIndex % changeCount) + 1} / ${changeCount}`}
          </span>
          <button type="button" className="icon-btn" onClick={onNext} disabled={changeCount === 0} title="Changement suivant (n)">
            <IconArrowDown />
          </button>
        </div>
        ) : null}
        <details className="options">
          <summary>
            <IconGear /> Options
          </summary>
          <div className="options-panel">
            <label>
              <input
                type="checkbox"
                checked={options.headerRow}
                onChange={(e) => onOptions({ ...options, headerRow: e.target.checked })}
              />
              Première ligne = en-têtes
            </label>
            <label>
              <input
                type="checkbox"
                checked={options.ignoreWhitespace}
                onChange={(e) => onOptions({ ...options, ignoreWhitespace: e.target.checked })}
              />
              Ignorer les espaces
            </label>
            <label>
              <input
                type="checkbox"
                checked={options.ignoreCase}
                onChange={(e) => onOptions({ ...options, ignoreCase: e.target.checked })}
              />
              Ignorer la casse
            </label>
            <label>
              <input
                type="checkbox"
                checked={options.compareFormulas}
                onChange={(e) => onOptions({ ...options, compareFormulas: e.target.checked })}
              />
              Comparer les formules
            </label>
            <label className="range">
              Contexte
              <input
                type="number"
                min={0}
                max={20}
                value={options.contextLines}
                onChange={(e) => onOptions({ ...options, contextLines: Number(e.target.value) || 0 })}
              />
            </label>
          </div>
        </details>
      </div>

      <div className="mergebar">
        <IconMerge className="mergebar-icon" />
        <button type="button" onClick={() => onAcceptAll('current')}>
          Tout garder : ancienne
        </button>
        <button type="button" onClick={() => onAcceptAll('incoming')}>
          Tout garder : nouvelle
        </button>
        <span className="mergebar-count">
          {pending.total === 0
            ? 'Toutes les modifications sont tranchées'
            : `${pending.total} en attente`}
        </span>
        {error ? <span className="mergebar-error">{error}</span> : null}
        <button
          type="button"
          className="merge-save"
          onClick={onSaveMerged}
          disabled={saving || highlighting}
          title="Conserve couleurs, polices, bordures et mises en forme. Les modifications encore en attente restent celles du fichier de base."
        >
          {saving ? 'Génération du fichier…' : 'Enregistrer le classeur fusionné'}
        </button>
      </div>

      <main className="main">
        {screen === 'review' ? (
          <ReviewView
            diff={diff}
            decisions={decisions}
            headerRow={options.headerRow}
            selectedSheet={sheet.name}
            onCellDecision={onCellDecision}
            onRowDecision={onRowDecision}
            onColumnDecision={onColumnDecision}
            onSheetDecision={onSheetDecision}
          />
        ) : (
          <DiffView
            sheet={sheet}
            view={view}
            showAll={showAll}
            showFormulas={showFormulas || options.compareFormulas}
            headerRow={options.headerRow}
            changeIndex={changeIndex}
            oldName={diff.oldName}
            newName={diff.newName}
            decisions={decisions}
            merge={{
              onRow: (row, value) => onRowDecision(sheet.name, row, value),
              onCell: (row, letter, value) => onCellDecision(sheet.name, row, letter, value),
              onColumn: (letter, value) => onColumnDecision(sheet.name, letter, value),
              onHunk: (rows, value) => onHunkDecision(sheet.name, rows, value),
              onSheet: (value) => onSheetDecision(sheet.name, value),
            }}
          />
        )}
      </main>
    </div>
  )
}

function Stats({
  stats,
}: {
  stats: WorkbookDiff['stats']
}) {
  return (
    <div className="stats">
      <span className="add">+{stats.added}</span>
      <span className="del">−{stats.removed}</span>
      <span className="chg">~{stats.changed}</span>
      <span className="muted">
        {stats.sheetsModified} feuille{stats.sheetsModified > 1 ? 's' : ''} modifiée
        {stats.sheetsModified > 1 ? 's' : ''}
      </span>
    </div>
  )
}

function statusLetter(status: SheetDiff['status']): string {
  if (status === 'added') return 'A'
  if (status === 'removed') return 'D'
  if (status === 'modified') return 'M'
  return ' '
}

function orderSheets(sheets: SheetDiff[]): SheetDiff[] {
  const rank = { modified: 0, added: 1, removed: 2, unchanged: 3 }
  return [...sheets].sort((a, b) => rank[a.status] - rank[b.status])
}
