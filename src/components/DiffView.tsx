import { IconChevronLeft, IconChevronRight } from './Icons.tsx'
import { useEffect, useMemo, useRef, type ReactNode } from 'react'
import {
  changeRowId,
  effectiveCell,
  getResolution,
  isRowResolved,
  rowKey,
} from '../excel/merge.ts'
import type {
  CellDiff,
  ColumnAlign,
  Hunk,
  MergeDecisions,
  MergeSide,
  Resolution,
  RowDiff,
  SheetDiff,
  ViewMode,
} from '../types.ts'

interface MergeHandlers {
  onRow: (row: RowDiff, value: Resolution) => void
  onCell: (row: RowDiff, letter: string, value: Resolution) => void
  onColumn: (letter: string, value: Resolution) => void
  onHunk: (rows: RowDiff[], value: Resolution) => void
  onSheet: (value: Resolution) => void
}

interface Props {
  sheet: SheetDiff
  view: ViewMode
  showAll: boolean
  showFormulas: boolean
  headerRow: boolean
  changeIndex: number
  oldName?: string
  newName?: string
  decisions: MergeDecisions
  merge: MergeHandlers
}

export function DiffView({
  sheet,
  view,
  showAll,
  showFormulas,
  headerRow,
  changeIndex,
  oldName,
  newName,
  decisions,
  merge,
}: Props) {
  const rootRef = useRef<HTMLDivElement>(null)
  const blocks: Array<{ hunk?: Hunk; rows: RowDiff[] }> = showAll
    ? [{ rows: dataRows(sheet, headerRow) }]
    : sheet.hunks.map((hunk) => ({ hunk, rows: hunk.rows }))

  const changeIds = useMemo(() => collectChangeIds(sheet, headerRow), [sheet, headerRow])

  useEffect(() => {
    if (changeIds.length === 0) return
    const id = changeIds[((changeIndex % changeIds.length) + changeIds.length) % changeIds.length]
    const el = rootRef.current?.querySelector(`[data-change-id="${id}"]`)
    el?.scrollIntoView({ block: 'center', behavior: 'smooth' })
  }, [changeIndex, changeIds])

  if (sheet.status === 'unchanged') {
    return (
      <div className="empty-diff">
        <p>Aucune différence sur cette feuille.</p>
      </div>
    )
  }

  const sheetRes =
    sheet.status === 'added' || sheet.status === 'removed'
      ? getResolution(decisions.sheets, sheet.name)
      : null

  const onlyColumns =
    sheet.stats.added === 0 &&
    sheet.stats.removed === 0 &&
    sheet.stats.changed === 0 &&
    (sheet.stats.columnsAdded > 0 || sheet.stats.columnsRemoved > 0)

  const hasRows = showAll ? dataRows(sheet, headerRow).length > 0 : sheet.hunks.length > 0

  return (
    <div className="diff-view" ref={rootRef}>
      {sheetRes !== null && (
        <div className="sheet-banner">
          <span>
            {sheet.status === 'added'
              ? 'Feuille ajoutée dans la nouvelle version'
              : 'Feuille absente de la nouvelle version'}
          </span>
          <LensButtons
            currentLabel={sheet.status === 'added' ? 'Ignorer la feuille' : 'Conserver la feuille'}
            incomingLabel={sheet.status === 'added' ? 'Accepter la feuille' : 'Supprimer la feuille'}
            value={sheetRes}
            onChange={merge.onSheet}
          />
        </div>
      )}

      {(sheet.stats.columnsAdded > 0 || sheet.stats.columnsRemoved > 0) && (
        <div className="column-banner">
          {sheet.columns
            .filter((c) => c.kind !== 'same')
            .map((c) => {
              const value = getResolution(decisions.columns, `${sheet.name}::col:${c.letter}`)
              return (
                <span key={`${c.kind}-${c.letter}`} className={`col-chip ${c.kind} is-${value}`}>
                  {c.kind === 'added' ? '+' : '−'} {c.header}
                  <LensButtons
                    currentLabel={c.kind === 'added' ? 'Ignorer' : 'Conserver'}
                    incomingLabel={c.kind === 'added' ? 'Accepter' : 'Supprimer'}
                    value={value}
                    onChange={(next) => merge.onColumn(c.letter, next)}
                  />
                </span>
              )
            })}
        </div>
      )}

      {onlyColumns && !showAll ? (
        <div className="empty-diff">
          <p>Seules des colonnes ont changé. Activez « Fichier entier » pour voir les valeurs.</p>
        </div>
      ) : !hasRows ? (
        <div className="empty-diff">
          <p>Aucune différence de ligne.</p>
        </div>
      ) : view === 'unified' ? (
        <UnifiedTable
          sheet={sheet}
          columns={sheet.columns}
          blocks={blocks}
          showFormulas={showFormulas}
          decisions={decisions}
          merge={merge}
        />
      ) : (
        <SplitTable
          sheet={sheet}
          columns={sheet.columns}
          blocks={blocks}
          showFormulas={showFormulas}
          decisions={decisions}
          merge={merge}
          oldName={oldName}
          newName={newName}
        />
      )}
    </div>
  )
}

function dataRows(sheet: SheetDiff, headerRow: boolean): RowDiff[] {
  if (!headerRow) return sheet.allRows
  return sheet.allRows.slice(1)
}

function collectChangeIds(sheet: SheetDiff, headerRow: boolean): string[] {
  return dataRows(sheet, headerRow)
    .filter((r) => r.kind !== 'context')
    .map(changeRowId)
}

function display(cell: CellDiff | undefined, side: 'old' | 'new', formulas: boolean): string {
  if (!cell) return ''
  if (formulas) {
    const f = side === 'old' ? cell.oldFormula : cell.newFormula
    if (f) return `=${f}`
  }
  return side === 'old' ? cell.oldText : cell.newText
}

function UnifiedTable({
  sheet,
  columns,
  blocks,
  showFormulas,
  decisions,
  merge,
}: {
  sheet: SheetDiff
  columns: ColumnAlign[]
  blocks: Array<{ hunk?: Hunk; rows: RowDiff[] }>
  showFormulas: boolean
  decisions: MergeDecisions
  merge: MergeHandlers
}) {
  const colSpan = columns.length + 3
  return (
    <div className="table-scroll">
      <table className="diff-table unified">
        <thead>
          <tr>
            <th className="gutter" />
            <th className="gutter" />
            <th className="sign" />
            {columns.map((col) => (
              <th key={col.letter} className={`col-${col.kind}`}>
                <span className="col-letter">{col.letter}</span>
                {col.header}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {blocks.map((block, bi) => (
            <BlockRows
              key={bi}
              block={block}
              colSpan={colSpan}
              onHunk={merge.onHunk}
              hunkResolution={hunkResolution(sheet, block.rows, decisions)}
              showMerge={sheet.status === 'modified'}
            >
              {block.rows.map((row) => (
                <ChangeRows
                  key={changeRowId(row)}
                  sheet={sheet}
                  row={row}
                  colSpan={colSpan}
                  decisions={decisions}
                  merge={merge}
                >
                  <tr
                    className={`row-${row.kind} is-${rowState(sheet, row, decisions)}`}
                    data-change-id={row.kind === 'context' ? undefined : changeRowId(row)}
                  >
                    <td className="gutter">{row.oldIndex != null ? row.oldIndex + 1 : ''}</td>
                    <td className="gutter">{row.newIndex != null ? row.newIndex + 1 : ''}</td>
                    <td className="sign">{signFor(row.kind)}</td>
                    {columns.map((col, i) => (
                      <td key={col.letter} className={cellClass(sheet, row, col, row.cells[i], decisions)}>
                        <CellContent
                          sheet={sheet}
                          row={row}
                          col={col}
                          cell={row.cells[i]}
                          showFormulas={showFormulas}
                          decisions={decisions}
                          onPick={(value) => merge.onCell(row, col.letter, value)}
                          unified
                        />
                      </td>
                    ))}
                  </tr>
                </ChangeRows>
              ))}
            </BlockRows>
          ))}
        </tbody>
      </table>
    </div>
  )
}

function SplitTable({
  sheet,
  columns,
  blocks,
  showFormulas,
  decisions,
  merge,
  oldName,
  newName,
}: {
  sheet: SheetDiff
  columns: ColumnAlign[]
  blocks: Array<{ hunk?: Hunk; rows: RowDiff[] }>
  showFormulas: boolean
  decisions: MergeDecisions
  merge: MergeHandlers
  oldName?: string
  newName?: string
}) {
  const leftRef = useRef<HTMLDivElement>(null)
  const rightRef = useRef<HTMLDivElement>(null)
  const railRef = useRef<HTMLDivElement>(null)
  const syncing = useRef(false)
  const canMerge = sheet.status === 'modified'

  function syncScroll(source: HTMLDivElement) {
    if (syncing.current) return
    syncing.current = true
    const top = source.scrollTop
    if (leftRef.current && leftRef.current !== source) leftRef.current.scrollTop = top
    if (rightRef.current && rightRef.current !== source) rightRef.current.scrollTop = top
    if (railRef.current && railRef.current !== source) railRef.current.scrollTop = top
    requestAnimationFrame(() => {
      syncing.current = false
    })
  }

  return (
    <div className="split-layout">
      <div className="split-head current">
        <span className="pane-kicker">Actuelle</span>
        <strong title={oldName}>{oldName ?? 'Base'}</strong>
      </div>
      <div className="split-head rail" />
      <div className="split-head incoming">
        <span className="pane-kicker">Entrante</span>
        <strong title={newName}>{newName ?? 'Nouveau'}</strong>
      </div>

      <div
        className="split-pane current"
        ref={leftRef}
        onScroll={(e) => syncScroll(e.currentTarget)}
      >
        <table className="diff-table split-side">
          <PaneCols columns={columns} />
          <tbody>
            {blocks.map((block, bi) => (
              <SplitPaneBody
                key={`L-${bi}`}
                side="current"
                block={block}
                sheet={sheet}
                columns={columns}
                showFormulas={showFormulas}
                decisions={decisions}
                merge={merge}
              />
            ))}
          </tbody>
        </table>
      </div>

      <div className="split-rail" ref={railRef} onScroll={(e) => syncScroll(e.currentTarget)}>
        <table>
          <thead>
            <tr>
              <th className="transfer-hd" />
            </tr>
          </thead>
          <tbody>
            {blocks.map((block, bi) => (
              <SplitRailBody
                key={`R-${bi}`}
                block={block}
                sheet={sheet}
                decisions={decisions}
                merge={merge}
                canMerge={canMerge}
              />
            ))}
          </tbody>
        </table>
      </div>

      <div
        className="split-pane incoming"
        ref={rightRef}
        onScroll={(e) => syncScroll(e.currentTarget)}
      >
        <table className="diff-table split-side">
          <PaneCols columns={columns} />
          <tbody>
            {blocks.map((block, bi) => (
              <SplitPaneBody
                key={`I-${bi}`}
                side="incoming"
                block={block}
                sheet={sheet}
                columns={columns}
                showFormulas={showFormulas}
                decisions={decisions}
                merge={merge}
              />
            ))}
          </tbody>
        </table>
      </div>
    </div>
  )
}

function PaneCols({ columns }: { columns: ColumnAlign[] }) {
  return (
    <thead>
      <tr>
        <th className="gutter" />
        <th className="sign" />
        {columns.map((col) => (
          <th key={col.letter} className={`col-${col.kind}`}>
            <span className="col-letter">{col.letter}</span>
            {col.header}
          </th>
        ))}
      </tr>
    </thead>
  )
}

function SplitPaneBody({
  side,
  block,
  sheet,
  columns,
  showFormulas,
  decisions,
  merge,
}: {
  side: MergeSide
  block: { hunk?: Hunk; rows: RowDiff[] }
  sheet: SheetDiff
  columns: ColumnAlign[]
  showFormulas: boolean
  decisions: MergeDecisions
  merge: MergeHandlers
}) {
  const groups = groupSplitRows(block.rows)
  return (
    <>
      {block.hunk && (
        <tr className="hunk-meta">
          <td colSpan={columns.length + 2}>
            @@ −{block.hunk.oldStart},{block.hunk.oldCount} +{block.hunk.newStart},{block.hunk.newCount} @@
          </td>
        </tr>
      )}
      {groups.map((group) =>
        group.rows.map((row) => {
          const hatch =
            (side === 'current' && group.kind === 'added') ||
            (side === 'incoming' && group.kind === 'removed')
          return (
            <tr
              key={`${side}-${changeRowId(row)}`}
              className={`row-${row.kind} is-${rowState(sheet, row, decisions)}`}
              data-change-id={side === 'incoming' && row.kind !== 'context' ? changeRowId(row) : undefined}
            >
              {hatch ? (
                <td className="cell-hatch" colSpan={columns.length + 2} />
              ) : (
                <SideCells
                  side={side}
                  sheet={sheet}
                  row={row}
                  kind={group.kind}
                  columns={columns}
                  showFormulas={showFormulas}
                  decisions={decisions}
                  merge={merge}
                />
              )}
            </tr>
          )
        }),
      )}
    </>
  )
}

function SplitRailBody({
  block,
  sheet,
  decisions,
  merge,
  canMerge,
}: {
  block: { hunk?: Hunk; rows: RowDiff[] }
  sheet: SheetDiff
  decisions: MergeDecisions
  merge: MergeHandlers
  canMerge: boolean
}) {
  const groups = groupSplitRows(block.rows)
  return (
    <>
      {block.hunk && (
        <tr className="hunk-meta">
          <td />
        </tr>
      )}
      {groups.map((group) => {
        const n = group.rows.length
        const value = groupResolution(sheet, group.rows, decisions)
        const arrows =
          canMerge && group.kind !== 'context' ? (
            <TransferArrows
              kind={group.kind}
              value={value}
              onCurrent={() => merge.onHunk(group.rows, value === 'current' ? 'pending' : 'current')}
              onIncoming={() => merge.onHunk(group.rows, value === 'incoming' ? 'pending' : 'incoming')}
            />
          ) : null
        return group.rows.map((row, i) => (
          <tr key={changeRowId(row)} className={`row-${row.kind}`}>
            {i === 0 ? (
              <td className="transfer-cell" rowSpan={n}>
                {arrows}
              </td>
            ) : null}
          </tr>
        ))
      })}
    </>
  )
}

function SideCells({
  side,
  sheet,
  row,
  kind,
  columns,
  showFormulas,
  decisions,
  merge,
}: {
  side: MergeSide
  sheet: SheetDiff
  row: RowDiff
  kind: RowDiff['kind']
  columns: ColumnAlign[]
  showFormulas: boolean
  decisions: MergeDecisions
  merge: MergeHandlers
}) {
  const displaySide = side === 'current' ? 'old' : 'new'
  const index = side === 'current' ? row.oldIndex : row.newIndex
  const sign = kind === 'removed' && side === 'current' ? '−' : kind === 'added' && side === 'incoming' ? '+' : ''
  return (
    <>
      <td className="gutter">{index != null ? index + 1 : ''}</td>
      <td className="sign">{sign}</td>
      {columns.map((col, ci) => {
        const cell = row.cells[ci]
        const changed = kind === 'changed' && cell?.kind === 'changed'
        const picked = changed ? effectiveCell(sheet.name, row, col, decisions) : null
        return (
          <td
            key={col.letter}
            className={
              kind === 'added'
                ? 'cell-added'
                : kind === 'removed'
                  ? 'cell-removed'
                  : cellClass(sheet, row, col, cell, decisions, side)
            }
          >
            {changed ? (
              <button
                type="button"
                className={`cell-pick ${picked === side ? 'is-on' : ''}`}
                onClick={() => merge.onCell(row, col.letter, side)}
              >
                {display(cell, displaySide, showFormulas)}
              </button>
            ) : (
              display(cell, displaySide, showFormulas)
            )}
          </td>
        )
      })}
    </>
  )
}

interface SplitRowGroup {
  kind: RowDiff['kind']
  rows: RowDiff[]
}

function groupSplitRows(rows: RowDiff[]): SplitRowGroup[] {
  const groups: SplitRowGroup[] = []
  for (const row of rows) {
    const last = groups[groups.length - 1]
    if (last && last.kind === row.kind && row.kind !== 'context') {
      last.rows.push(row)
    } else {
      groups.push({ kind: row.kind, rows: [row] })
    }
  }
  return groups
}

function TransferArrows({
  kind,
  value,
  onCurrent,
  onIncoming,
}: {
  kind: RowDiff['kind']
  value: Resolution
  onCurrent: () => void
  onIncoming: () => void
}) {
  return (
    <div className="transfer-btns">
      <button
        type="button"
        className={`transfer-btn incoming ${value === 'incoming' ? 'is-on' : ''}`}
        title={
          kind === 'added'
            ? 'Prendre cet ajout'
            : kind === 'removed'
              ? 'Accepter la suppression'
              : 'Prendre la version entrante'
        }
        onClick={onIncoming}
      >
        <IconChevronRight />
      </button>
      <button
        type="button"
        className={`transfer-btn current ${value === 'current' ? 'is-on' : ''}`}
        title={
          kind === 'added'
            ? 'Ignorer cet ajout'
            : kind === 'removed'
              ? 'Conserver la ligne'
              : 'Garder la version actuelle'
        }
        onClick={onCurrent}
      >
        <IconChevronLeft />
      </button>
    </div>
  )
}

function groupResolution(sheet: SheetDiff, rows: RowDiff[], decisions: MergeDecisions): Resolution {
  const states = rows.map((row) => getResolution(decisions.rows, rowKey(sheet.name, row)))
  const first = states[0]
  if (first && states.every((s) => s === first)) return first
  return 'pending'
}

function ChangeRows({
  sheet,
  row,
  colSpan,
  decisions,
  merge,
  children,
}: {
  sheet: SheetDiff
  row: RowDiff
  colSpan: number
  decisions: MergeDecisions
  merge: MergeHandlers
  children: ReactNode
}) {
  if (row.kind === 'context') return <>{children}</>
  if (sheet.status === 'added' || sheet.status === 'removed') return <>{children}</>
  const labels = actionLabels(row.kind)
  const value = getResolution(decisions.rows, rowKey(sheet.name, row))
  const mixed = value === 'pending' && isRowResolved(sheet, row, decisions)
  return (
    <>
      <tr className="merge-lens">
        <td colSpan={colSpan}>
          <LensButtons
            currentLabel={labels.current}
            incomingLabel={labels.incoming}
            value={mixed ? 'pending' : value}
            onChange={(next) => merge.onRow(row, next)}
          />
        </td>
      </tr>
      {children}
    </>
  )
}

function BlockRows({
  block,
  colSpan,
  children,
  onHunk,
  hunkResolution: hunkRes,
  showMerge,
}: {
  block: { hunk?: Hunk; rows: RowDiff[] }
  colSpan: number
  children: ReactNode
  onHunk: (rows: RowDiff[], value: Resolution) => void
  hunkResolution: Resolution | 'mixed'
  showMerge: boolean
}) {
  return (
    <>
      {block.hunk && (
        <tr className="hunk-meta">
          <td colSpan={colSpan}>
            <span className="hunk-range">
              @@ −{block.hunk.oldStart},{block.hunk.oldCount} +{block.hunk.newStart},{block.hunk.newCount} @@
            </span>
            {showMerge && (
              <LensButtons
                currentLabel="Accepter actuelles"
                incomingLabel="Accepter entrantes"
                value={hunkRes === 'mixed' ? 'pending' : hunkRes}
                onChange={(next) => onHunk(block.hunk!.rows, next)}
              />
            )}
          </td>
        </tr>
      )}
      {children}
    </>
  )
}

function LensButtons({
  currentLabel,
  incomingLabel,
  value,
  onChange,
}: {
  currentLabel: string
  incomingLabel: string
  value: Resolution
  onChange: (value: Resolution) => void
}) {
  return (
    <span className="lens">
      <button
        type="button"
        className={`lens-btn current ${value === 'current' ? 'is-on' : ''}`}
        onClick={() => onChange(value === 'current' ? 'pending' : 'current')}
      >
        {currentLabel}
      </button>
      <span className="lens-sep">|</span>
      <button
        type="button"
        className={`lens-btn incoming ${value === 'incoming' ? 'is-on' : ''}`}
        onClick={() => onChange(value === 'incoming' ? 'pending' : 'incoming')}
      >
        {incomingLabel}
      </button>
    </span>
  )
}

function CellContent({
  sheet,
  row,
  col,
  cell,
  showFormulas,
  decisions,
  onPick,
  unified,
}: {
  sheet: SheetDiff
  row: RowDiff
  col: ColumnAlign
  cell: CellDiff | undefined
  showFormulas: boolean
  decisions: MergeDecisions
  onPick: (value: MergeSide) => void
  unified?: boolean
}) {
  if (!cell) return null
  const pickable = row.kind === 'changed' && cell.kind === 'changed'
  if (unified && pickable) {
    const chosen = effectiveCell(sheet.name, row, col, decisions)
    return (
      <span className="val-pair">
        <button
          type="button"
          className={`val-old pick ${chosen === 'current' ? 'is-on' : ''}`}
          onClick={() => onPick('current')}
          title="Accepter la valeur actuelle (base)"
        >
          {display(cell, 'old', showFormulas)}
        </button>
        <button
          type="button"
          className={`val-new pick ${chosen === 'incoming' ? 'is-on' : ''}`}
          onClick={() => onPick('incoming')}
          title="Accepter la valeur entrante"
        >
          {display(cell, 'new', showFormulas)}
        </button>
      </span>
    )
  }
  const side = row.kind === 'removed' ? 'old' : 'new'
  return <>{display(cell, side, showFormulas)}</>
}

function actionLabels(kind: RowDiff['kind']): { current: string; incoming: string } {
  if (kind === 'added') return { current: 'Ignorer l’ajout', incoming: 'Accepter l’ajout' }
  if (kind === 'removed') return { current: 'Conserver la ligne', incoming: 'Accepter la suppression' }
  return { current: 'Accepter actuelle', incoming: 'Accepter entrante' }
}

function rowState(sheet: SheetDiff, row: RowDiff, decisions: MergeDecisions): Resolution | 'mixed' {
  const value = getResolution(decisions.rows, rowKey(sheet.name, row))
  if (value !== 'pending') return value
  if (row.kind === 'changed' && isRowResolved(sheet, row, decisions)) return 'mixed'
  return 'pending'
}

function hunkResolution(sheet: SheetDiff, rows: RowDiff[], decisions: MergeDecisions): Resolution | 'mixed' {
  const changes = rows.filter((r) => r.kind !== 'context')
  if (changes.length === 0) return 'pending'
  const states = changes.map((row) => getResolution(decisions.rows, rowKey(sheet.name, row)))
  const first = states[0]
  if (first && states.every((s) => s === first)) return first
  return 'mixed'
}

function cellClass(
  sheet: SheetDiff,
  row: RowDiff,
  col: ColumnAlign,
  cell: CellDiff | undefined,
  decisions: MergeDecisions,
  splitSide?: MergeSide,
): string {
  if (!cell) return 'cell-same'
  if (row.kind === 'context' && cell.kind === 'added' && !cell.newText) return 'cell-same'
  if (row.kind === 'context' && cell.kind === 'removed' && !cell.oldText) return 'cell-same'
  if (row.kind === 'changed' && cell.kind === 'changed') {
    const chosen = effectiveCell(sheet.name, row, col, decisions)
    if (chosen === 'pending') return 'cell-changed'
    if (splitSide) return chosen === splitSide ? `cell-${splitSide === 'incoming' ? 'added' : 'removed'} is-on` : 'cell-same'
    return chosen === 'incoming' ? 'cell-added' : 'cell-removed'
  }
  return `cell-${cell.kind}`
}

function signFor(kind: RowDiff['kind']): string {
  if (kind === 'added') return '+'
  if (kind === 'removed') return '−'
  if (kind === 'changed') return '~'
  return ''
}
