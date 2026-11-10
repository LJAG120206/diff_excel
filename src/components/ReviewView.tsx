import { useEffect, useMemo, useRef, useState, type RefObject } from 'react'
import {
  buildReviewItems,
  findColumn,
  findRow,
  findSheet,
  itemResolution,
  rowCaption,
  type ReviewItem,
} from '../excel/review.ts'
import { changeRowId } from '../excel/merge.ts'
import type { CellDiff, MergeDecisions, MergeSide, Resolution, RowDiff, SheetDiff, WorkbookDiff } from '../types.ts'

interface Props {
  diff: WorkbookDiff
  decisions: MergeDecisions
  headerRow: boolean
  selectedSheet: string
  onCellDecision: (sheetName: string, row: RowDiff, letter: string, value: Resolution) => void
  onRowDecision: (sheetName: string, row: RowDiff, value: Resolution) => void
  onColumnDecision: (sheetName: string, letter: string, value: Resolution) => void
  onSheetDecision: (sheetName: string, value: Resolution) => void
}

const COLLAPSE_AFTER = 8

export function ReviewView({
  diff,
  decisions,
  headerRow,
  selectedSheet,
  onCellDecision,
  onRowDecision,
  onColumnDecision,
  onSheetDecision,
}: Props) {
  const items = useMemo(() => buildReviewItems(diff, headerRow), [diff, headerRow])
  const [openId, setOpenId] = useState<string | null>(null)
  const [pendingOnly, setPendingOnly] = useState(false)
  const [expanded, setExpanded] = useState<Record<string, boolean>>({})
  const listRef = useRef<HTMLDivElement>(null)

  const visible = useMemo(() => {
    if (!pendingOnly) return items
    return items.filter((item) => {
      const resolution = itemResolution(diff, item, decisions)
      return resolution === 'pending' || resolution === 'mixed'
    })
  }, [items, pendingOnly, diff, decisions])

  const openIndex = visible.findIndex((item) => item.id === openId)
  const openItem = openIndex >= 0 ? visible[openIndex] : undefined

  useEffect(() => {
    const row = listRef.current?.querySelector(`[data-sheet-anchor="${CSS.escape(selectedSheet)}"]`)
    row?.scrollIntoView({ block: 'start' })
  }, [selectedSheet])

  useEffect(() => {
    if (!openItem || openItem.type !== 'cell') return
    const parent = items.find(
      (entry) => entry.type === 'column-cells' && entry.sheetName === openItem.sheetName && entry.letter === openItem.letter,
    )
    if (parent && parent.type === 'column-cells' && parent.rowIds.length > COLLAPSE_AFTER) {
      setExpanded((prev) => (prev[parent.id] ? prev : { ...prev, [parent.id]: true }))
    }
    const row = listRef.current?.querySelector(`[data-review-id="${CSS.escape(openItem.id)}"]`)
    row?.scrollIntoView({ block: 'nearest' })
  }, [openItem, items])

  useEffect(() => {
    function onKey(event: KeyboardEvent) {
      const tag = (event.target as HTMLElement | null)?.tagName
      if (tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT') return
      if (event.key === 'Escape' && openItem) {
        event.preventDefault()
        event.stopPropagation()
        setOpenId(null)
        return
      }
      if (!openItem && (event.key === 'Enter') && visible[0]) {
        event.preventDefault()
        event.stopPropagation()
        setOpenId(visible[0].id)
        return
      }
      if (event.key !== 'n' && event.key !== 'j' && event.key !== 'ArrowDown' && event.key !== 'p' && event.key !== 'k' && event.key !== 'ArrowUp') {
        return
      }
      if (!openItem && event.key !== 'ArrowDown' && event.key !== 'ArrowUp') return
      event.preventDefault()
      event.stopPropagation()
      const direction = event.key === 'p' || event.key === 'k' || event.key === 'ArrowUp' ? -1 : 1
      move(direction)
    }
    window.addEventListener('keydown', onKey, true)
    return () => window.removeEventListener('keydown', onKey, true)
  })

  function move(direction: number) {
    if (visible.length === 0) return
    const current = openIndex >= 0 ? openIndex : -1
    const next = Math.min(visible.length - 1, Math.max(0, current + direction))
    setOpenId(visible[next]!.id)
  }

  function advanceAfter(item: ReviewItem) {
    const index = visible.findIndex((entry) => entry.id === item.id)
    if (index < 0) return
    if (item.type === 'column-cells') {
      const after = visible.findIndex(
        (entry, entryIndex) =>
          entryIndex > index && !(entry.type === 'cell' && entry.sheetName === item.sheetName && entry.letter === item.letter),
      )
      setOpenId(after >= 0 ? visible[after]!.id : null)
      return
    }
    const next = visible[index + 1]
    setOpenId(next ? next.id : null)
  }

  function choose(item: ReviewItem, side: MergeSide) {
    const sheet = findSheet(diff, item.sheetName)
    if (!sheet) return
    if (item.type === 'sheet') onSheetDecision(item.sheetName, side)
    if (item.type === 'column') onColumnDecision(item.sheetName, item.letter, side)
    if (item.type === 'row') {
      const row = findRow(sheet, item.rowId)
      if (row) onRowDecision(item.sheetName, row, side)
    }
    if (item.type === 'cell') {
      const row = findRow(sheet, item.rowId)
      if (row) onCellDecision(item.sheetName, row, item.letter, side)
    }
    if (item.type === 'column-cells') {
      const column = findColumn(sheet, item.letter)
      if (column) {
        for (const rowId of item.rowIds) {
          const row = findRow(sheet, rowId)
          if (row) onCellDecision(item.sheetName, row, column.letter, side)
        }
      }
    }
    advanceAfter(item)
  }

  let lastSheet = ''

  return (
    <div className="review">
      <div className="review-toolbar">
        <strong>{visible.length} écart{visible.length > 1 ? 's' : ''}</strong>
        <span className="review-hint">Cliquer une ligne pour comparer. n / p pour passer à l’écart suivant.</span>
        <label className="check">
          <input type="checkbox" checked={pendingOnly} onChange={(event) => setPendingOnly(event.target.checked)} />
          À trancher seulement
        </label>
      </div>
      <div className="review-list" ref={listRef}>
        {visible.length === 0 ? <p className="review-empty">Aucun écart à afficher.</p> : null}
        {visible.map((item) => {
          const sheet = findSheet(diff, item.sheetName)
          if (!sheet) return null
          const showSheet = item.sheetName !== lastSheet
          lastSheet = item.sheetName
          if (item.type === 'cell' && !isExpanded(expanded, item, visible)) return null
          return (
            <div key={item.id}>
              {showSheet ? (
                <div className="review-sheet" data-sheet-anchor={item.sheetName}>
                  {item.sheetName}
                </div>
              ) : null}
              <ReviewRow
                diff={diff}
                item={item}
                sheet={sheet}
                decisions={decisions}
                active={item.id === openId}
                expanded={Boolean(expanded[item.id])}
                childCount={item.type === 'column-cells' ? item.rowIds.length : 0}
                onToggle={() => setExpanded((prev) => ({ ...prev, [item.id]: !prev[item.id] }))}
                onOpen={() => setOpenId(item.id)}
              />
            </div>
          )
        })}
      </div>
      {openItem ? (
        <ReviewModal
          diff={diff}
          item={openItem}
          decisions={decisions}
          headerRow={headerRow}
          index={openIndex}
          count={visible.length}
          oldName={diff.oldName}
          newName={diff.newName}
          onClose={() => setOpenId(null)}
          onPrev={() => move(-1)}
          onNext={() => move(1)}
          onChoose={(side) => choose(openItem, side)}
        />
      ) : null}
    </div>
  )
}

function isExpanded(expanded: Record<string, boolean>, item: ReviewItem, visible: ReviewItem[]): boolean {
  if (item.type !== 'cell') return true
  const parent = visible.find(
    (entry) => entry.type === 'column-cells' && entry.sheetName === item.sheetName && entry.letter === item.letter,
  )
  if (!parent || parent.type !== 'column-cells') return true
  if (parent.rowIds.length <= COLLAPSE_AFTER) return true
  return Boolean(expanded[parent.id])
}

function ReviewRow({
  diff,
  item,
  sheet,
  decisions,
  active,
  expanded,
  childCount,
  onToggle,
  onOpen,
}: {
  diff: WorkbookDiff
  item: ReviewItem
  sheet: SheetDiff
  decisions: MergeDecisions
  active: boolean
  expanded: boolean
  childCount: number
  onToggle: () => void
  onOpen: () => void
}) {
  const resolution = itemResolution(diff, item, decisions)
  const preview = previewOf(sheet, item)
  return (
    <div className={`review-row is-${preview.tone} ${active ? 'is-active' : ''}`} data-review-id={item.id}>
      <button type="button" className="review-open" onClick={onOpen}>
        <span className={`review-mark tone-${preview.tone}`}>{preview.mark}</span>
        <span className="review-title">{preview.title}</span>
        <span className="review-values">
          <span className="review-old">{preview.oldText}</span>
          <span className="review-arrow">→</span>
          <span className="review-new">{preview.newText}</span>
        </span>
        <span className={`review-choice is-${resolution}`}>{choiceLabel(resolution)}</span>
      </button>
      {item.type === 'column-cells' && childCount > COLLAPSE_AFTER ? (
        <button type="button" className="review-expand" onClick={onToggle}>
          {expanded ? 'Masquer les lignes' : `Afficher les ${childCount} lignes`}
        </button>
      ) : null}
    </div>
  )
}

function ReviewModal({
  diff,
  item,
  decisions,
  headerRow,
  index,
  count,
  oldName,
  newName,
  onClose,
  onPrev,
  onNext,
  onChoose,
}: {
  diff: WorkbookDiff
  item: ReviewItem
  decisions: MergeDecisions
  headerRow: boolean
  index: number
  count: number
  oldName: string
  newName: string
  onClose: () => void
  onPrev: () => void
  onNext: () => void
  onChoose: (side: MergeSide) => void
}) {
  const dialogRef = useRef<HTMLDivElement>(null)
  const sheet = findSheet(diff, item.sheetName)
  const resolution = itemResolution(diff, item, decisions)
  useEffect(() => {
    dialogRef.current?.focus()
  }, [item.id])
  if (!sheet) return null
  const preview = previewOf(sheet, item)
  return (
    <div className="review-backdrop" onMouseDown={onClose}>
      <div
        className="review-modal"
        role="dialog"
        aria-modal="true"
        aria-label={preview.title}
        tabIndex={-1}
        ref={dialogRef}
        onMouseDown={(event) => event.stopPropagation()}
      >
        <header className="review-modal-bar">
          <span>
            {index + 1} / {count}
          </span>
          <strong>{preview.title}</strong>
          <button type="button" className="review-close" onClick={onClose}>
            Fermer
          </button>
        </header>
        <div className="review-compare">
          <ValueCard label="Ancienne" fileName={oldName} text={preview.oldText} side="current" selected={resolution === 'current'} />
          <ValueCard label="Nouvelle" fileName={newName} text={preview.newText} side="incoming" selected={resolution === 'incoming'} />
        </div>
        <SheetPreview sheet={sheet} item={item} headerRow={headerRow} oldName={oldName} newName={newName} />
        <footer className="review-modal-actions">
          <button type="button" onClick={onPrev} disabled={index <= 0}>
            Précédent
          </button>
          <button type="button" className={`review-keep current ${resolution === 'current' ? 'is-on' : ''}`} onClick={() => onChoose('current')}>
            Garder l’ancienne
          </button>
          <button type="button" className={`review-keep incoming ${resolution === 'incoming' ? 'is-on' : ''}`} onClick={() => onChoose('incoming')}>
            Garder la nouvelle
          </button>
          <button type="button" onClick={onNext} disabled={index >= count - 1}>
            Suivant
          </button>
        </footer>
      </div>
    </div>
  )
}

function ValueCard({
  label,
  fileName,
  text,
  side,
  selected,
}: {
  label: string
  fileName: string
  text: string
  side: MergeSide
  selected: boolean
}) {
  return (
    <section className={`review-card side-${side} ${selected ? 'is-on' : ''}`}>
      <header>
        <span>{label}</span>
        <small>{fileName}</small>
      </header>
      <p>{text || '—'}</p>
    </section>
  )
}

function SheetPreview({
  sheet,
  item,
  headerRow,
  oldName,
  newName,
}: {
  sheet: SheetDiff
  item: ReviewItem
  headerRow: boolean
  oldName: string
  newName: string
}) {
  const [radius, setRadius] = useState(8)
  const [allColumns, setAllColumns] = useState(false)
  const oldScroll = useRef<HTMLDivElement>(null)
  const newScroll = useRef<HTMLDivElement>(null)
  const syncing = useRef(false)
  useEffect(() => {
    setRadius(8)
    setAllColumns(false)
  }, [item.id])

  const anchor = anchorIndex(sheet, item)
  const focusLetter = item.type === 'cell' || item.type === 'column' || item.type === 'column-cells' ? item.letter : ''
  const focusIds = new Set(item.type === 'column-cells' ? item.rowIds : item.type === 'cell' || item.type === 'row' ? [item.rowId] : [])
  const bodyStart = headerRow ? 1 : 0
  const from = Math.max(bodyStart, anchor - radius)
  const to = Math.min(sheet.allRows.length, Math.max(anchor, bodyStart) + radius + 1)
  const rows = sheet.allRows.slice(from, to)
  const columns = visibleColumns(sheet, focusLetter, rows, allColumns)
  const hiddenBefore = from - bodyStart
  const hiddenAfter = sheet.allRows.length - to
  const showOld = sheet.status !== 'added'
  const showNew = sheet.status !== 'removed'

  useEffect(() => {
    const target = (showOld ? oldScroll.current : newScroll.current)?.querySelector('.is-row')
    target?.scrollIntoView({ block: 'center', inline: 'center' })
  }, [item.id, radius, showOld])

  function sync(source: HTMLDivElement, other: HTMLDivElement | null) {
    if (!other || syncing.current) return
    syncing.current = true
    other.scrollTop = source.scrollTop
    other.scrollLeft = source.scrollLeft
    syncing.current = false
  }

  return (
    <div className="sheet-preview">
      <div className="sheet-preview-bar">
        <span>
          {hiddenBefore > 0 ? `${hiddenBefore} lignes au-dessus · ` : ''}
          {columns.length} colonne{columns.length > 1 ? 's' : ''} sur {sheet.columns.length}
          {hiddenAfter > 0 ? ` · ${hiddenAfter} lignes en dessous` : ''}
        </span>
        <span className="sheet-preview-actions">
          <button type="button" onClick={() => setAllColumns((value) => !value)} disabled={sheet.columns.length <= columns.length && !allColumns}>
            {allColumns ? 'Colonnes autour de l’écart' : 'Toutes les colonnes'}
          </button>
          <button type="button" onClick={() => setRadius((value) => Math.min(value + 12, 80))} disabled={hiddenBefore + hiddenAfter === 0}>
            Plus de lignes
          </button>
        </span>
      </div>
      <div className={`sheet-panes ${showOld && showNew ? '' : 'is-single'}`}>
        {showOld ? (
          <VersionPane
            title="Ancien"
            fileName={oldName}
            rows={rows}
            columns={columns}
            side="old"
            focusLetter={focusLetter}
            focusIds={focusIds}
            scrollRef={oldScroll}
            onScroll={(node) => sync(node, newScroll.current)}
          />
        ) : null}
        {showNew ? (
          <VersionPane
            title="Nouveau"
            fileName={newName}
            rows={rows}
            columns={columns}
            side="new"
            focusLetter={focusLetter}
            focusIds={focusIds}
            scrollRef={newScroll}
            onScroll={(node) => sync(node, oldScroll.current)}
          />
        ) : null}
      </div>
    </div>
  )
}

function VersionPane({
  title,
  fileName,
  rows,
  columns,
  side,
  focusLetter,
  focusIds,
  scrollRef,
  onScroll,
}: {
  title: string
  fileName: string
  rows: RowDiff[]
  columns: ShownColumn[]
  side: 'old' | 'new'
  focusLetter: string
  focusIds: Set<string>
  scrollRef: RefObject<HTMLDivElement | null>
  onScroll: (node: HTMLDivElement) => void
}) {
  return (
    <section className={`version-pane side-${side}`}>
      <header>
        <strong>{title}</strong>
        <small>{fileName}</small>
      </header>
      <div className="version-scroll" ref={scrollRef} onScroll={(event) => onScroll(event.currentTarget)}>
        <table>
          <thead>
            <tr>
              <th />
              {columns.map(({ column }) => (
                <th key={column.letter} className={column.letter === focusLetter ? 'is-col' : ''}>
                  {column.header}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((row) => {
              const id = changeRowId(row)
              const line = side === 'old' ? row.oldIndex : row.newIndex
              const focused = focusIds.has(id)
              return (
                <tr key={id} className={focused ? 'is-row' : ''}>
                  <th>{line == null ? '' : line + 1}</th>
                  {columns.map(({ column, index }) => {
                    const cell = row.cells[index]
                    return (
                      <td
                        key={column.letter}
                        className={`${toneClass(row, cell, side)} ${column.letter === focusLetter ? 'is-col' : ''}`}
                      >
                        {cellTextFor(row, cell, side)}
                      </td>
                    )
                  })}
                </tr>
              )
            })}
          </tbody>
        </table>
      </div>
    </section>
  )
}

interface ShownColumn {
  column: SheetDiff['columns'][number]
  index: number
}

function visibleColumns(sheet: SheetDiff, focusLetter: string, rows: RowDiff[], all: boolean): ShownColumn[] {
  const allColumns = sheet.columns.map((column, index) => ({ column, index }))
  if (all || sheet.columns.length <= 5) return allColumns
  const picked = new Set<number>()
  const key = sheet.columns.findIndex((column) => column.kind === 'same')
  picked.add(key >= 0 ? key : 0)
  const focus = sheet.columns.findIndex((column) => column.letter === focusLetter)
  if (focus >= 0) {
    picked.add(focus)
    if (focus > 0) picked.add(focus - 1)
    if (focus + 1 < sheet.columns.length) picked.add(focus + 1)
  } else {
    rows.forEach((row) => {
      row.cells.forEach((cell, index) => {
        if ((cell?.oldText || cell?.newText || '').trim()) picked.add(index)
      })
    })
  }
  const ordered = [...picked].sort((left, right) => left - right)
  const limited = focus >= 0 ? ordered : ordered.slice(0, 6)
  return limited.map((index) => allColumns[index]!)
}

function anchorIndex(sheet: SheetDiff, item: ReviewItem): number {
  if (item.type === 'cell' || item.type === 'row') {
    const index = sheet.allRows.findIndex((row) => changeRowId(row) === item.rowId)
    return index < 0 ? 0 : index
  }
  if (item.type === 'column-cells') {
    const index = sheet.allRows.findIndex((row) => changeRowId(row) === item.rowIds[0])
    return index < 0 ? 0 : index
  }
  return 0
}

function cellTextFor(row: RowDiff, cell: CellDiff | undefined, side: 'old' | 'new'): string {
  if (side === 'old') {
    if (row.oldIndex == null) return ''
    return cell?.oldText ?? ''
  }
  if (row.newIndex == null) return ''
  return cell?.newText ?? ''
}

function toneClass(row: RowDiff, cell: CellDiff | undefined, side: 'old' | 'new'): string {
  if (side === 'old' && (row.oldIndex == null || cell?.kind === 'added')) return 'is-absent'
  if (side === 'new' && (row.newIndex == null || cell?.kind === 'removed')) return 'is-absent'
  if (cell?.kind === 'changed') return 'is-chg'
  if (side === 'new' && (row.kind === 'added' || cell?.kind === 'added')) return 'is-add'
  if (side === 'old' && (row.kind === 'removed' || cell?.kind === 'removed')) return 'is-del'
  return ''
}

function previewOf(sheet: SheetDiff, item: ReviewItem): { title: string; oldText: string; newText: string; mark: string; tone: 'add' | 'del' | 'chg' } {
  if (item.type === 'sheet') {
    const added = sheet.status === 'added'
    return {
      title: added ? `Feuille ajoutée · ${sheet.name}` : `Feuille supprimée · ${sheet.name}`,
      oldText: added ? '—' : sheet.name,
      newText: added ? sheet.name : '—',
      mark: added ? 'A' : 'S',
      tone: added ? 'add' : 'del',
    }
  }
  if (item.type === 'column' || item.type === 'column-cells') {
    const column = findColumn(sheet, item.letter)
    const header = column?.header ?? item.letter
    if (item.type === 'column') {
      const added = column?.kind === 'added'
      return {
        title: `${added ? 'Colonne ajoutée' : 'Colonne supprimée'} · ${header}`,
        oldText: added ? '—' : header,
        newText: added ? header : '—',
        mark: added ? 'A' : 'S',
        tone: added ? 'add' : 'del',
      }
    }
    return {
      title: `Colonne ${header} · ${item.rowIds.length} cellule${item.rowIds.length > 1 ? 's' : ''}`,
      oldText: 'version ancienne',
      newText: 'version nouvelle',
      mark: '~',
      tone: 'chg',
    }
  }
  const row = findRow(sheet, item.rowId)
  if (!row) return { title: item.id, oldText: '', newText: '', mark: '?', tone: 'chg' }
  if (item.type === 'row') {
    const added = row.kind === 'added'
    return {
      title: `${added ? 'Ligne ajoutée' : 'Ligne supprimée'} · ${rowCaption(sheet, row)}`,
      oldText: added ? '—' : rowSummary(sheet, row, 'old'),
      newText: added ? rowSummary(sheet, row, 'new') : '—',
      mark: added ? 'A' : 'S',
      tone: added ? 'add' : 'del',
    }
  }
  const column = findColumn(sheet, item.letter)
  const cell = row.cells[sheet.columns.findIndex((entry) => entry.letter === item.letter)]
  return {
    title: `${column?.header ?? item.letter} · ${rowCaption(sheet, row)}`,
    oldText: cell?.oldText || '—',
    newText: cell?.newText || '—',
    mark: '~',
    tone: 'chg',
  }
}

function rowSummary(sheet: SheetDiff, row: RowDiff, side: 'old' | 'new'): string {
  const parts = sheet.columns
    .map((column, index) => {
      const cell = row.cells[index]
      const text = side === 'old' ? cell?.oldText : cell?.newText
      return text?.trim() ? `${column.header}: ${text}` : ''
    })
    .filter(Boolean)
    .slice(0, 4)
  return parts.join(' · ') || '—'
}

function choiceLabel(value: Resolution | 'mixed'): string {
  if (value === 'current') return 'Ancienne'
  if (value === 'incoming') return 'Nouvelle'
  if (value === 'mixed') return 'Mixte'
  return 'À trancher'
}
