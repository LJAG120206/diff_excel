import { useMemo } from 'react'
import type { DiffOptions, SheetDiff, Theme, ViewMode, WorkbookDiff } from '../types.ts'
import {
  IconArrowDown,
  IconArrowUp,
  IconDownload,
  IconGear,
  IconMoon,
  IconReset,
  IconSun,
  IconSwap,
  IconTable,
} from './Icons.tsx'
import { DiffView } from './DiffView.tsx'

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
}: Props) {
  const sheet = diff.sheets.find((s) => s.name === selected) ?? diff.sheets[0]
  const ordered = useMemo(() => orderSheets(diff.sheets), [diff.sheets])

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
                className={`sheet-item ${item.name === sheet.name ? 'is-active' : ''} status-${item.status}`}
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

      <main className="main">
        <DiffView
          sheet={sheet}
          view={view}
          showAll={showAll}
          showFormulas={showFormulas || options.compareFormulas}
          headerRow={options.headerRow}
          changeIndex={changeIndex}
        />
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
