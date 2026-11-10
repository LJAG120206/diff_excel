import type { Cell, CellValue, DataValidation, Fill, Row, Style, Worksheet } from 'exceljs'
import type { CellDiff, ColumnAlign, RowDiff, SheetDiff, WorkbookData, WorkbookDiff } from '../types.ts'
import { Workbook } from './exceljsClient.ts'

const FILL_ADDED: Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6EFCE' } }
const FILL_REMOVED: Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC7CE' } }
const FILL_CHANGED: Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF4B183' } }
const FILL_NEW_COLUMN: Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF5019BD' } }

type Mark = 'added' | 'removed' | 'changed'
type Side = 'old' | 'new'
type ExcelWorkbook = InstanceType<typeof Workbook>

type Slot =
  | { role: 'single'; column: ColumnAlign; index: number }
  | { role: 'old'; column: ColumnAlign; index: number }
  | { role: 'new'; column: ColumnAlign; index: number }

function fillFor(mark: Mark): Fill {
  if (mark === 'added') return FILL_ADDED
  if (mark === 'removed') return FILL_REMOVED
  return FILL_CHANGED
}

function uniqueSheetName(name: string, used: Set<string>): string {
  const cleaned = name.replace(/[\\/*?:[\]]/g, ' ').slice(0, 31) || 'Feuille'
  let next = cleaned
  let i = 2
  while (used.has(next.toLowerCase())) {
    const suffix = ` ${i}`
    next = `${cleaned.slice(0, 31 - suffix.length)}${suffix}`
    i++
  }
  used.add(next.toLowerCase())
  return next
}

function columnHasChanges(sheet: SheetDiff, index: number): boolean {
  return sheet.allRows.some((row) => row.cells[index]?.kind === 'changed')
}

function slotsFor(sheet: SheetDiff): Slot[] {
  const slots: Slot[] = []
  sheet.columns.forEach((column, index) => {
    if (column.kind === 'same' && columnHasChanges(sheet, index)) {
      slots.push({ role: 'old', column, index }, { role: 'new', column, index })
      return
    }
    slots.push({ role: 'single', column, index })
  })
  return slots
}

function cloneData<T>(value: T): T {
  if (value == null) return value
  try {
    return structuredClone(value)
  } catch {
    return JSON.parse(JSON.stringify(value)) as T
  }
}

function copyValue(value: CellValue): CellValue {
  if (value instanceof Date) return new Date(value.getTime())
  if (Array.isArray(value)) return value.map((item) => copyValue(item as CellValue)) as unknown as CellValue
  if (value && typeof value === 'object') return cloneData(value)
  return value
}

function copyCell(src: Cell, dest: Cell): void {
  dest.value = copyValue(src.value)
  dest.style = cloneData(src.style ?? ({} as Partial<Style>))
  if (src.numFmt) dest.numFmt = src.numFmt
  if (src.dataValidation) dest.dataValidation = cloneData(src.dataValidation) as DataValidation
}

function findSheet(wb: ExcelWorkbook, name: string): Worksheet | undefined {
  return wb.getWorksheet(name) ?? wb.worksheets.find((sheet) => sheet.name.toLowerCase() === name.toLowerCase())
}

function isXlsxLike(name: string): boolean {
  return /\.xlsx$/i.test(name) || /\.xlsm$/i.test(name)
}

async function loadWorkbook(buffer: ArrayBuffer | undefined, fileName: string): Promise<ExcelWorkbook | undefined> {
  if (!buffer || !isXlsxLike(fileName)) return undefined
  const wb = new Workbook()
  await wb.xlsx.load(buffer.slice(0) as unknown as Parameters<ExcelWorkbook['xlsx']['load']>[0])
  return wb
}

function resultOf(value: unknown): string | number | boolean | Date | undefined {
  if (typeof value === 'number' || typeof value === 'boolean' || typeof value === 'string') return value
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value
  return undefined
}

function writePlain(target: Cell, value: unknown, text: string, formula?: string): void {
  if (formula) {
    const body = formula.startsWith('=') ? formula.slice(1) : formula
    const result = resultOf(value)
    target.value = result === undefined ? { formula: body } : { formula: body, result }
    return
  }
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    target.value = value
    target.numFmt = 'dd/mm/yyyy'
    return
  }
  if (typeof value === 'number' || typeof value === 'boolean' || typeof value === 'string') {
    target.value = value
    return
  }
  target.value = text || null
}

function sideOf(slot: Slot, row: RowDiff, cell: CellDiff): Side {
  if (slot.role === 'old') return 'old'
  if (slot.role === 'new') return 'new'
  if (row.kind === 'removed' || cell.kind === 'removed' || slot.column.kind === 'removed') return 'old'
  if (row.kind === 'added' || cell.kind === 'added' || slot.column.kind === 'added') return 'new'
  return slot.column.oldIndex != null ? 'old' : 'new'
}

function markOf(slot: Slot, row: RowDiff, cell: CellDiff): Mark | null {
  if (row.kind === 'added' || cell.kind === 'added' || slot.column.kind === 'added') return 'added'
  if (row.kind === 'removed' || cell.kind === 'removed' || slot.column.kind === 'removed') return 'removed'
  if (slot.role !== 'single' && cell.kind === 'changed') return 'changed'
  return null
}

function labelNewHeader(target: Cell, fallback: string): void {
  const current = target.value
  const text =
    typeof current === 'string'
      ? current
      : typeof current === 'number' || typeof current === 'boolean'
        ? String(current)
        : fallback
  target.value = text ? `${text} (nouveau)` : `${fallback} (nouveau)`
}

function sourceAddress(side: Side, row: RowDiff, column: ColumnAlign): { row: number; col: number } | null {
  if (side === 'old') {
    if (row.oldIndex == null || column.oldIndex == null) return null
    return { row: row.oldIndex + 1, col: column.oldIndex + 1 }
  }
  if (row.newIndex == null || column.newIndex == null) return null
  return { row: row.newIndex + 1, col: column.newIndex + 1 }
}

function writeSlot(
  target: Cell,
  slot: Slot,
  row: RowDiff,
  cell: CellDiff,
  sources: { old?: Worksheet; new?: Worksheet },
  header: boolean,
): void {
  const side = sideOf(slot, row, cell)
  const from = sourceAddress(side, row, slot.column)
  const src = side === 'old' ? sources.old : sources.new
  if (from && src) copyCell(src.getRow(from.row).getCell(from.col), target)
  else {
    const value = side === 'old' ? cell.oldValue : cell.newValue
    const text = side === 'old' ? cell.oldText : cell.newText
    const formula = side === 'old' ? cell.oldFormula : cell.newFormula
    writePlain(target, value, text, formula)
  }
  if (header && slot.role === 'new') labelNewHeader(target, slot.column.header)
  const mark = markOf(slot, row, cell)
  if (mark) target.fill = fillFor(mark)
  if (slot.role === 'new' || slot.column.kind === 'added') {
    target.fill = FILL_NEW_COLUMN
    target.font = { ...(target.font ?? {}), color: { argb: 'FFFFFFFF' } }
  }
}

function applyRowLayout(destRow: Row, sources: { old?: Worksheet; new?: Worksheet }, row: RowDiff): void {
  const src =
    row.oldIndex != null && sources.old
      ? sources.old.getRow(row.oldIndex + 1)
      : row.newIndex != null && sources.new
        ? sources.new.getRow(row.newIndex + 1)
        : undefined
  if (!src) return
  if (src.height != null) destRow.height = src.height
  if (src.hidden != null) destRow.hidden = src.hidden
  if (src.outlineLevel != null) destRow.outlineLevel = src.outlineLevel
}

function writeLegend(ws: Worksheet, diff: WorkbookDiff): void {
  ws.getCell(1, 1).value = 'Légende du comparatif'
  ws.getCell(1, 1).font = { bold: true, size: 14 }
  ws.getCell(2, 1).value = `Ancien fichier : ${diff.oldName}`
  ws.getCell(3, 1).value = `Nouveau fichier : ${diff.newName}`

  const rows: { mark: Mark; label: string; detail: string }[] = [
    { mark: 'added', label: 'Ajouté', detail: 'Présent seulement dans le fichier nouveau.' },
    { mark: 'removed', label: 'Supprimé', detail: 'Présent seulement dans le fichier ancien. La ligne reste visible ici.' },
    {
      mark: 'changed',
      label: 'Modifié',
      detail: 'La colonne d’origine est conservée. La copie « (nouveau) » est en violet (#5019BD), texte blanc. Supprimez la colonne que vous ne gardez pas.',
    },
  ]
  rows.forEach((item, index) => {
    const sample = ws.getCell(5 + index, 1)
    sample.value = item.label
    sample.fill = item.mark === 'changed' ? FILL_NEW_COLUMN : fillFor(item.mark)
    sample.font = { bold: true, color: { argb: item.mark === 'changed' ? 'FFFFFFFF' : 'FF000000' } }
    ws.getCell(5 + index, 2).value = item.detail
  })
  ws.getCell(9, 1).value = 'Le reste de la mise en forme vient du fichier d’origine. Les cellules sans couleur sont identiques.'
  ws.getColumn(1).width = 18
  ws.getColumn(2).width = 110
}

function writeDataSheet(
  ws: Worksheet,
  sheet: SheetDiff,
  headerRow: boolean,
  sources: { old?: Worksheet; new?: Worksheet },
): void {
  const slots = slotsFor(sheet)
  const styleSheet = sources.old ?? sources.new
  if (styleSheet) {
    ws.properties = cloneData(styleSheet.properties)
    ws.pageSetup = cloneData(styleSheet.pageSetup)
    ws.headerFooter = cloneData(styleSheet.headerFooter)
    if (styleSheet.views?.length) ws.views = cloneData(styleSheet.views)
  }

  slots.forEach((slot, outIndex) => {
    const side: Side = slot.role === 'new' || slot.column.kind === 'added' ? 'new' : 'old'
    const srcWs = side === 'old' ? sources.old : sources.new
    const srcIndex = side === 'old' ? slot.column.oldIndex : slot.column.newIndex
    if (srcWs && srcIndex != null) {
      const srcCol = srcWs.getColumn(srcIndex + 1)
      const destCol = ws.getColumn(outIndex + 1)
      if (srcCol.width != null) destCol.width = srcCol.width
      if (srcCol.hidden != null) destCol.hidden = srcCol.hidden
      if (srcCol.outlineLevel != null) destCol.outlineLevel = srcCol.outlineLevel
    } else {
      ws.getColumn(outIndex + 1).width = Math.min(40, Math.max(12, slot.column.header.length + 2))
    }
  })

  sheet.allRows.forEach((row, r) => {
    const destRow = ws.getRow(r + 1)
    applyRowLayout(destRow, sources, row)
    slots.forEach((slot, c) => {
      const cell = row.cells[slot.index]
      if (!cell) return
      writeSlot(destRow.getCell(c + 1), slot, row, cell, sources, headerRow && r === 0)
    })
  })

  if (headerRow && sheet.allRows.length > 0 && !styleSheet?.views?.length) {
    ws.views = [{ state: 'frozen', ySplit: 1 }]
  }
}

function triggerDownload(data: ArrayBuffer | Uint8Array, fileName: string): void {
  const view = data instanceof Uint8Array ? data : new Uint8Array(data)
  const copy = new ArrayBuffer(view.byteLength)
  new Uint8Array(copy).set(view)
  const blob = new Blob([copy], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = fileName
  a.click()
  URL.revokeObjectURL(url)
}

export async function buildHighlightedBuffer(
  diff: WorkbookDiff,
  headerRow: boolean,
  oldWb?: WorkbookData,
  newWb?: WorkbookData,
): Promise<ArrayBuffer> {
  const [oldBook, newBook] = await Promise.all([
    loadWorkbook(oldWb?.buffer, oldWb?.fileName ?? ''),
    loadWorkbook(newWb?.buffer, newWb?.fileName ?? ''),
  ])
  const book = new Workbook()
  const base = oldBook ?? newBook
  if (base) {
    book.creator = base.creator
    book.title = base.title
  }
  const used = new Set<string>()
  writeLegend(book.addWorksheet(uniqueSheetName('Légende', used)), diff)
  for (const sheet of diff.sheets) {
    const ws = book.addWorksheet(uniqueSheetName(sheet.name, used))
    writeDataSheet(ws, sheet, headerRow, {
      old: oldBook ? findSheet(oldBook, sheet.name) : undefined,
      new: newBook ? findSheet(newBook, sheet.name) : undefined,
    })
  }
  return (await book.xlsx.writeBuffer()) as ArrayBuffer
}

export async function downloadHighlightedXlsx(
  diff: WorkbookDiff,
  headerRow: boolean,
  oldWb?: WorkbookData,
  newWb?: WorkbookData,
): Promise<void> {
  const base = diff.newName.replace(/\.[^.]+$/, '')
  triggerDownload(await buildHighlightedBuffer(diff, headerRow, oldWb, newWb), `${base}-comparatif.xlsx`)
}
