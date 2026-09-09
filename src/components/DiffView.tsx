import { useEffect, useMemo, useRef, type ReactNode } from 'react'
import type { CellDiff, ColumnAlign, Hunk, RowDiff, SheetDiff, ViewMode } from '../types.ts'

interface Props {
  sheet: SheetDiff
  view: ViewMode
  showAll: boolean
  showFormulas: boolean
  headerRow: boolean
  changeIndex: number
}

export function DiffView({ sheet, view, showAll, showFormulas, headerRow, changeIndex }: Props) {
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

  const onlyColumns =
    sheet.stats.added === 0 &&
    sheet.stats.removed === 0 &&
    sheet.stats.changed === 0 &&
    (sheet.stats.columnsAdded > 0 || sheet.stats.columnsRemoved > 0)

  const hasRows = showAll ? dataRows(sheet, headerRow).length > 0 : sheet.hunks.length > 0

  return (
    <div className="diff-view" ref={rootRef}>
      {(sheet.stats.columnsAdded > 0 || sheet.stats.columnsRemoved > 0) && (
        <div className="column-banner">
          {sheet.columns
            .filter((c) => c.kind !== 'same')
            .map((c) => (
              <span key={`${c.kind}-${c.letter}`} className={`col-chip ${c.kind}`}>
                {c.kind === 'added' ? '+' : '−'} {c.header}
              </span>
            ))}
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
        <UnifiedTable columns={sheet.columns} blocks={blocks} showFormulas={showFormulas} />
      ) : (
        <SplitTable columns={sheet.columns} blocks={blocks} showFormulas={showFormulas} />
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
    .map(rowId)
}

function rowId(row: RowDiff): string {
  return `${row.kind}-${row.oldIndex ?? 'x'}-${row.newIndex ?? 'x'}`
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
  columns,
  blocks,
  showFormulas,
}: {
  columns: ColumnAlign[]
  blocks: Array<{ hunk?: Hunk; rows: RowDiff[] }>
  showFormulas: boolean
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
            <BlockRows key={bi} block={block} colSpan={colSpan}>
              {block.rows.map((row) => (
                <tr
                  key={rowId(row)}
                  className={`row-${row.kind}`}
                  data-change-id={row.kind === 'context' ? undefined : rowId(row)}
                >
                  <td className="gutter">{row.oldIndex != null ? row.oldIndex + 1 : ''}</td>
                  <td className="gutter">{row.newIndex != null ? row.newIndex + 1 : ''}</td>
                  <td className="sign">{signFor(row.kind)}</td>
                  {columns.map((col, i) => {
                    const cell = row.cells[i]
                    if (row.kind === 'changed' && cell?.kind === 'changed') {
                      return (
                        <td key={col.letter} className="cell-changed">
                          <span className="val-old">{display(cell, 'old', showFormulas)}</span>
                          <span className="val-new">{display(cell, 'new', showFormulas)}</span>
                        </td>
                      )
                    }
                    const side = row.kind === 'removed' ? 'old' : 'new'
                    return (
                      <td key={col.letter} className={cellClass(row, cell)}>
                        {display(cell, side, showFormulas)}
                      </td>
                    )
                  })}
                </tr>
              ))}
            </BlockRows>
          ))}
        </tbody>
      </table>
    </div>
  )
}

function SplitTable({
  columns,
  blocks,
  showFormulas,
}: {
  columns: ColumnAlign[]
  blocks: Array<{ hunk?: Hunk; rows: RowDiff[] }>
  showFormulas: boolean
}) {
  const colSpan = columns.length * 2 + 5
  return (
    <div className="table-scroll">
      <table className="diff-table split">
        <thead>
          <tr>
            <th className="gutter" />
            <th className="sign" />
            {columns.map((col) => (
              <th key={`o-${col.letter}`} className={`col-${col.kind}`}>
                <span className="col-letter">{col.letter}</span>
                {col.header}
              </th>
            ))}
            <th className="split-gap" />
            <th className="gutter" />
            <th className="sign" />
            {columns.map((col) => (
              <th key={`n-${col.letter}`} className={`col-${col.kind}`}>
                <span className="col-letter">{col.letter}</span>
                {col.header}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {blocks.map((block, bi) => (
            <BlockRows key={bi} block={block} colSpan={colSpan}>
              {block.rows.map((row) => (
                <tr
                  key={rowId(row)}
                  className={`row-${row.kind}`}
                  data-change-id={row.kind === 'context' ? undefined : rowId(row)}
                >
                  <td className="gutter">
                    {row.kind === 'added' ? '' : row.oldIndex != null ? row.oldIndex + 1 : ''}
                  </td>
                  <td className="sign">{row.kind === 'added' || row.kind === 'context' ? '' : '−'}</td>
                  {columns.map((col, i) => {
                    const cell = row.cells[i]
                    const empty = row.kind === 'added'
                    return (
                      <td
                        key={`o-${col.letter}`}
                        className={
                          empty
                            ? 'cell-blank'
                            : `cell-${cell?.kind === 'changed' ? 'removed' : (cell?.kind ?? 'same')}`
                        }
                      >
                        {empty ? '' : display(cell, 'old', showFormulas)}
                      </td>
                    )
                  })}
                  <td className="split-gap" />
                  <td className="gutter">
                    {row.kind === 'removed' ? '' : row.newIndex != null ? row.newIndex + 1 : ''}
                  </td>
                  <td className="sign">{row.kind === 'removed' || row.kind === 'context' ? '' : '+'}</td>
                  {columns.map((col, i) => {
                    const cell = row.cells[i]
                    const empty = row.kind === 'removed'
                    return (
                      <td
                        key={`n-${col.letter}`}
                        className={
                          empty
                            ? 'cell-blank'
                            : `cell-${cell?.kind === 'changed' ? 'added' : (cell?.kind ?? 'same')}`
                        }
                      >
                        {empty ? '' : display(cell, 'new', showFormulas)}
                      </td>
                    )
                  })}
                </tr>
              ))}
            </BlockRows>
          ))}
        </tbody>
      </table>
    </div>
  )
}

function BlockRows({
  block,
  colSpan,
  children,
}: {
  block: { hunk?: Hunk; rows: RowDiff[] }
  colSpan: number
  children: ReactNode
}) {
  return (
    <>
      {block.hunk && (
        <tr className="hunk-meta">
          <td colSpan={colSpan}>
            @@ −{block.hunk.oldStart},{block.hunk.oldCount} +{block.hunk.newStart},{block.hunk.newCount} @@
          </td>
        </tr>
      )}
      {children}
    </>
  )
}

function cellClass(row: RowDiff, cell: CellDiff | undefined): string {
  if (!cell) return 'cell-same'
  if (row.kind === 'context' && cell.kind === 'added' && !cell.newText) return 'cell-same'
  if (row.kind === 'context' && cell.kind === 'removed' && !cell.oldText) return 'cell-same'
  return `cell-${cell.kind}`
}

function signFor(kind: RowDiff['kind']): string {
  if (kind === 'added') return '+'
  if (kind === 'removed') return '−'
  if (kind === 'changed') return '~'
  return ''
}
