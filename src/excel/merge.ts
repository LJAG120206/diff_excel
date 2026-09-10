import * as XLSX from 'xlsx'
import type {
  CellData,
  CellDiff,
  ColumnAlign,
  MergeDecisions,
  MergeSide,
  Resolution,
  RowDiff,
  SheetDiff,
  WorkbookData,
  WorkbookDiff,
} from '../types.ts'
import { emptyCell } from './format.ts'

export function emptyDecisions(): MergeDecisions {
  return { sheets: {}, columns: {}, rows: {}, cells: {} }
}

export function changeRowId(row: RowDiff): string {
  return `${row.kind}-${row.oldIndex ?? 'x'}-${row.newIndex ?? 'x'}`
}

export function rowKey(sheet: string, row: RowDiff): string {
  return `${sheet}::${changeRowId(row)}`
}

export function colKey(sheet: string, letter: string): string {
  return `${sheet}::col:${letter}`
}

export function cellKey(sheet: string, row: RowDiff, letter: string): string {
  return `${rowKey(sheet, row)}::${letter}`
}

export function getResolution(map: Record<string, Resolution>, key: string): Resolution {
  return map[key] ?? 'pending'
}

export function dataRows(sheet: SheetDiff, headerRow: boolean): RowDiff[] {
  if (!headerRow) return sheet.allRows
  return sheet.allRows.slice(1)
}

function pickCell(cell: CellDiff, side: MergeSide): CellData {
  if (side === 'incoming') {
    return { value: cell.newValue, text: cell.newText, formula: cell.newFormula }
  }
  return { value: cell.oldValue, text: cell.oldText, formula: cell.oldFormula }
}

export function effectiveSheet(sheet: SheetDiff, decisions: MergeDecisions): Resolution {
  if (sheet.status === 'unchanged' || sheet.status === 'modified') return 'incoming'
  return getResolution(decisions.sheets, sheet.name)
}

function effectiveColumn(sheet: string, col: ColumnAlign, decisions: MergeDecisions): Resolution {
  if (col.kind === 'same') return 'incoming'
  return getResolution(decisions.columns, colKey(sheet, col.letter))
}

function effectiveRow(sheet: string, row: RowDiff, decisions: MergeDecisions): Resolution {
  return getResolution(decisions.rows, rowKey(sheet, row))
}

export function effectiveCell(
  sheet: string,
  row: RowDiff,
  col: ColumnAlign,
  decisions: MergeDecisions,
): Resolution {
  const explicit = getResolution(decisions.cells, cellKey(sheet, row, col.letter))
  if (explicit !== 'pending') return explicit
  return effectiveRow(sheet, row, decisions)
}

export function isRowResolved(
  sheet: SheetDiff,
  row: RowDiff,
  decisions: MergeDecisions,
): boolean {
  if (effectiveRow(sheet.name, row, decisions) !== 'pending') return true
  if (row.kind !== 'changed') return false
  const changed = sheet.columns.filter((_, i) => row.cells[i]?.kind === 'changed')
  if (changed.length === 0) return true
  return changed.every((col) => getResolution(decisions.cells, cellKey(sheet.name, row, col.letter)) !== 'pending')
}

export function setRowDecision(
  prev: MergeDecisions,
  sheet: string,
  row: RowDiff,
  value: Resolution,
): MergeDecisions {
  const key = rowKey(sheet, row)
  const prefix = `${key}::`
  const cells = { ...prev.cells }
  for (const cellKeyName of Object.keys(cells)) {
    if (cellKeyName.startsWith(prefix)) delete cells[cellKeyName]
  }
  return { ...prev, rows: { ...prev.rows, [key]: value }, cells }
}

export function setCellDecision(
  prev: MergeDecisions,
  sheet: string,
  row: RowDiff,
  letter: string,
  value: Resolution,
): MergeDecisions {
  return {
    ...prev,
    cells: { ...prev.cells, [cellKey(sheet, row, letter)]: value },
  }
}

export function setColumnDecision(
  prev: MergeDecisions,
  sheet: string,
  letter: string,
  value: Resolution,
): MergeDecisions {
  return { ...prev, columns: { ...prev.columns, [colKey(sheet, letter)]: value } }
}

export function setSheetDecision(prev: MergeDecisions, name: string, value: Resolution): MergeDecisions {
  return { ...prev, sheets: { ...prev.sheets, [name]: value } }
}

export function setHunkDecision(
  prev: MergeDecisions,
  sheet: string,
  rows: RowDiff[],
  value: Resolution,
): MergeDecisions {
  let next = prev
  for (const row of rows) {
    if (row.kind === 'context') continue
    next = setRowDecision(next, sheet, row, value)
  }
  return next
}

export function acceptAll(diff: WorkbookDiff, side: MergeSide): MergeDecisions {
  let next = emptyDecisions()
  for (const sheet of diff.sheets) {
    if (sheet.status === 'added' || sheet.status === 'removed') {
      next = setSheetDecision(next, sheet.name, side)
    }
    for (const col of sheet.columns) {
      if (col.kind !== 'same') next = setColumnDecision(next, sheet.name, col.letter, side)
    }
    for (const row of sheet.allRows) {
      if (row.kind !== 'context') next = setRowDecision(next, sheet.name, row, side)
    }
  }
  return next
}

export interface PendingCounts {
  total: number
  rows: number
  columns: number
  sheets: number
}

export function countPending(diff: WorkbookDiff, decisions: MergeDecisions, headerRow: boolean): PendingCounts {
  let rows = 0
  let columns = 0
  let sheets = 0
  for (const sheet of diff.sheets) {
    if (sheet.status === 'added' || sheet.status === 'removed') {
      if (effectiveSheet(sheet, decisions) === 'pending') sheets++
      continue
    }
    for (const col of sheet.columns) {
      if (col.kind !== 'same' && effectiveColumn(sheet.name, col, decisions) === 'pending') columns++
    }
    for (const row of dataRows(sheet, headerRow)) {
      if (row.kind === 'context') continue
      if (!isRowResolved(sheet, row, decisions)) rows++
    }
  }
  return { total: rows + columns + sheets, rows, columns, sheets }
}

export function includeColumn(sheet: string, col: ColumnAlign, decisions: MergeDecisions): boolean {
  const res = effectiveColumn(sheet, col, decisions)
  if (col.kind === 'added') return res === 'incoming'
  if (col.kind === 'removed') return res !== 'incoming'
  return true
}

export function includeAddedRow(row: RowDiff, sheet: string, decisions: MergeDecisions): boolean {
  return effectiveRow(sheet, row, decisions) === 'incoming'
}

export function includeRemovedRow(row: RowDiff, sheet: string, decisions: MergeDecisions): boolean {
  return effectiveRow(sheet, row, decisions) !== 'incoming'
}

export function styleSourceSide(
  sheet: SheetDiff,
  row: RowDiff,
  col: ColumnAlign,
  decisions: MergeDecisions,
): MergeSide {
  if (row.kind === 'added') return 'incoming'
  if (row.kind === 'removed') return 'current'
  if (col.kind === 'added') return 'incoming'
  if (col.kind === 'removed') return 'current'
  if (row.kind === 'changed') {
    const cell = row.cells[sheet.columns.indexOf(col)]
    if (cell?.kind === 'changed') return resolvedCellSide(sheet, row, col, decisions)
    return effectiveRow(sheet.name, row, decisions) === 'incoming' ? 'incoming' : 'current'
  }
  return 'current'
}

export function hasIncomingDecisions(decisions: MergeDecisions): boolean {
  return [decisions.sheets, decisions.columns, decisions.rows, decisions.cells].some((map) =>
    Object.values(map).some((value) => value === 'incoming'),
  )
}

export function canOverlaySheet(sheet: SheetDiff, decisions: MergeDecisions, headerRow: boolean): boolean {
  for (const col of sheet.columns) {
    if (col.kind === 'added' && includeColumn(sheet.name, col, decisions)) return false
    if (col.kind === 'removed' && !includeColumn(sheet.name, col, decisions)) return false
  }
  for (let index = 0; index < sheet.allRows.length; index++) {
    const row = sheet.allRows[index]!
    const isHeader = headerRow && index === 0
    if (isHeader) continue
    if (row.kind === 'added' && includeAddedRow(row, sheet.name, decisions)) return false
    if (row.kind === 'removed' && !includeRemovedRow(row, sheet.name, decisions)) return false
  }
  return true
}

export type MergeSheetMode = 'clone-current' | 'clone-incoming' | 'overlay' | 'rebuild'

export interface MergeSheetPlan {
  sheet: SheetDiff
  mode: MergeSheetMode
  columns: ColumnAlign[]
  rows: RowDiff[]
}

export function getMergedPlan(
  diff: WorkbookDiff,
  decisions: MergeDecisions,
  headerRow: boolean,
): MergeSheetPlan[] {
  const plans: MergeSheetPlan[] = []
  for (const sheet of diff.sheets) {
    if (sheet.status === 'added') {
      if (effectiveSheet(sheet, decisions) !== 'incoming') continue
      plans.push({ sheet, mode: 'clone-incoming', columns: sheet.columns, rows: sheet.allRows })
      continue
    }
    if (sheet.status === 'removed') {
      if (effectiveSheet(sheet, decisions) === 'incoming') continue
      plans.push({ sheet, mode: 'clone-current', columns: sheet.columns, rows: sheet.allRows })
      continue
    }
    if (sheet.status === 'unchanged') {
      plans.push({ sheet, mode: 'clone-current', columns: sheet.columns, rows: sheet.allRows })
      continue
    }
    const columns = sheet.columns.filter((col) => includeColumn(sheet.name, col, decisions))
    const rows: RowDiff[] = []
    sheet.allRows.forEach((row, index) => {
      const isHeader = headerRow && index === 0
      if (!isHeader) {
        if (row.kind === 'added' && !includeAddedRow(row, sheet.name, decisions)) return
        if (row.kind === 'removed' && !includeRemovedRow(row, sheet.name, decisions)) return
      }
      rows.push(row)
    })
    plans.push({
      sheet,
      mode: canOverlaySheet(sheet, decisions, headerRow) ? 'overlay' : 'rebuild',
      columns,
      rows,
    })
  }
  return plans
}

function resolvedCellSide(
  sheet: SheetDiff,
  row: RowDiff,
  col: ColumnAlign,
  decisions: MergeDecisions,
): MergeSide {
  const res = effectiveCell(sheet.name, row, col, decisions)
  if (res === 'incoming') return 'incoming'
  return 'current'
}

function projectRow(
  sheet: SheetDiff,
  row: RowDiff,
  columns: ColumnAlign[],
  decisions: MergeDecisions,
): CellData[] {
  return columns.map((col, i) => {
    const cell = row.cells[i]
    if (!cell) return emptyCell()
    if (row.kind === 'added') return pickCell(cell, 'incoming')
    if (row.kind === 'removed') return pickCell(cell, 'current')
    if (col.kind === 'added') return pickCell(cell, 'incoming')
    if (col.kind === 'removed') return pickCell(cell, 'current')
    if (row.kind === 'changed' && cell.kind === 'changed') {
      return pickCell(cell, resolvedCellSide(sheet, row, col, decisions))
    }
    return pickCell(cell, 'incoming')
  })
}

function buildSheetRows(sheet: SheetDiff, decisions: MergeDecisions, headerRow: boolean): CellData[][] {
  const columns = sheet.columns.filter((col) => includeColumn(sheet.name, col, decisions))
  const out: CellData[][] = []
  sheet.allRows.forEach((row, index) => {
    const isHeader = headerRow && index === 0
    if (!isHeader) {
      if (row.kind === 'added' && !includeAddedRow(row, sheet.name, decisions)) return
      if (row.kind === 'removed' && !includeRemovedRow(row, sheet.name, decisions)) return
    }
    out.push(projectRow(sheet, row, columns, decisions))
  })
  return out
}

export function buildMergedWorkbook(
  diff: WorkbookDiff,
  decisions: MergeDecisions,
  headerRow: boolean,
): WorkbookData {
  const sheets: WorkbookData['sheets'] = []
  for (const sheet of diff.sheets) {
    const status = sheet.status
    if (status === 'added' && effectiveSheet(sheet, decisions) !== 'incoming') continue
    if (status === 'removed' && effectiveSheet(sheet, decisions) === 'incoming') continue
    sheets.push({
      name: sheet.name,
      rows: buildSheetRows(sheet, decisions, headerRow),
    })
  }
  if (sheets.length === 0) {
    sheets.push({ name: 'Vide', rows: [] })
  }
  const base = diff.newName.replace(/\.[^.]+$/, '')
  return {
    fileName: `${base}-fusion.xlsx`,
    fileSize: 0,
    sheets,
  }
}

function cellType(value: unknown): 'd' | 'n' | 'b' | 's' | 'z' {
  if (value instanceof Date) return 'd'
  if (typeof value === 'number') return 'n'
  if (typeof value === 'boolean') return 'b'
  if (value == null || value === '') return 'z'
  return 's'
}

function toCellObject(cell: CellData): XLSX.CellObject {
  if (cell.formula) {
    return {
      t: cellType(cell.value) === 'z' ? 'n' : cellType(cell.value),
      v: (cell.value ?? '') as string | number | boolean | Date,
      f: cell.formula,
      w: cell.text || undefined,
    }
  }
  if (cell.value instanceof Date) return { t: 'd', v: cell.value, w: cell.text || undefined }
  if (typeof cell.value === 'number') return { t: 'n', v: cell.value, w: cell.text || undefined }
  if (typeof cell.value === 'boolean') return { t: 'b', v: cell.value, w: cell.text || undefined }
  if (cell.value == null || cell.value === '') return { t: 'z' }
  return { t: 's', v: String(cell.value), w: cell.text || undefined }
}

function sheetName(name: string, used: Set<string>): string {
  const cleaned = name.replace(/[\\/*?:[\]]/g, ' ').slice(0, 31) || 'Feuille'
  let next = cleaned
  let i = 2
  while (used.has(next)) {
    const suffix = ` ${i}`
    next = `${cleaned.slice(0, 31 - suffix.length)}${suffix}`
    i++
  }
  used.add(next)
  return next
}

function triggerDownload(data: ArrayBuffer | Uint8Array | Blob, fileName: string): void {
  let blob: Blob
  if (data instanceof Blob) {
    blob = data
  } else {
    const view = data instanceof Uint8Array ? data : new Uint8Array(data)
    const copy = new ArrayBuffer(view.byteLength)
    new Uint8Array(copy).set(view)
    blob = new Blob([copy], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    })
  }
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = fileName
  a.click()
  URL.revokeObjectURL(url)
}

function writeValuesXlsx(wb: WorkbookData): void {
  const book = XLSX.utils.book_new()
  const used = new Set<string>()
  for (const sheet of wb.sheets) {
    const ws: XLSX.WorkSheet = {}
    let maxC = 0
    sheet.rows.forEach((row, r) => {
      row.forEach((cell, c) => {
        maxC = Math.max(maxC, c)
        ws[XLSX.utils.encode_cell({ r, c })] = toCellObject(cell)
      })
    })
    if (sheet.rows.length > 0) {
      ws['!ref'] = XLSX.utils.encode_range({
        s: { r: 0, c: 0 },
        e: { r: sheet.rows.length - 1, c: maxC },
      })
    }
    XLSX.utils.book_append_sheet(book, ws, sheetName(sheet.name, used))
  }
  const out = XLSX.write(book, { bookType: 'xlsx', type: 'array' }) as ArrayBuffer
  triggerDownload(out, wb.fileName)
}

function mergedFileName(diff: WorkbookDiff, identity: boolean, oldWb: WorkbookData): string {
  const base = diff.newName.replace(/\.[^.]+$/, '')
  if (identity) {
    const ext = oldWb.fileName.match(/\.[^.]+$/)?.[0] ?? '.xlsx'
    return `${base}-fusion${ext}`
  }
  return `${base}-fusion.xlsx`
}

export async function downloadMergedXlsx(
  diff: WorkbookDiff,
  decisions: MergeDecisions,
  headerRow: boolean,
  oldWb: WorkbookData,
  newWb: WorkbookData,
): Promise<void> {
  const identity = !hasIncomingDecisions(decisions) && Boolean(oldWb.buffer)
  const fileName = mergedFileName(diff, identity, oldWb)

  if (identity && oldWb.buffer) {
    triggerDownload(oldWb.buffer.slice(0), fileName)
    return
  }

  try {
    const { writeStyledXlsx } = await import('./styleCopy.ts')
    const out = await writeStyledXlsx(diff, decisions, headerRow, oldWb, newWb)
    triggerDownload(out, fileName)
    return
  } catch (error) {
    console.error('Export avec mise en forme impossible, repli sur les valeurs.', error)
  }

  const merged = buildMergedWorkbook(diff, decisions, headerRow)
  merged.fileName = fileName
  writeValuesXlsx(merged)
}
