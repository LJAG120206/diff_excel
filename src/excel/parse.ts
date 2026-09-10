import * as XLSX from 'xlsx'
import type { CellData, SheetData, WorkbookData } from '../types.ts'
import { emptyCell, formatValue, isEmptyCell } from './format.ts'

function parseSheet(ws: XLSX.WorkSheet, name: string): SheetData {
  const ref = ws['!ref']
  if (!ref) return { name, rows: [] }

  const range = XLSX.utils.decode_range(ref)
  const rows: CellData[][] = []

  for (let r = 0; r <= range.e.r; r++) {
    const row: CellData[] = []
    for (let c = 0; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c })
      const cell = ws[addr] as XLSX.CellObject | undefined
      if (!cell) {
        row.push(emptyCell())
        continue
      }
      const value = cell.v ?? null
      row.push({
        value,
        text: cell.w ?? formatValue(value),
        formula: cell.f,
      })
    }
    rows.push(row)
  }

  return { name, rows: trimGrid(rows) }
}

function rowIsEmpty(row: CellData[]): boolean {
  return row.every((cell) => isEmptyCell(cell))
}

function trimGrid(rows: CellData[][]): CellData[][] {
  let lastRow = rows.length - 1
  while (lastRow >= 0 && rowIsEmpty(rows[lastRow]!)) lastRow--
  if (lastRow < 0) return []

  let lastCol = -1
  for (let r = 0; r <= lastRow; r++) {
    const row = rows[r]!
    for (let c = row.length - 1; c >= 0; c--) {
      if (!isEmptyCell(row[c])) {
        lastCol = Math.max(lastCol, c)
        break
      }
    }
  }
  if (lastCol < 0) return []

  return rows.slice(0, lastRow + 1).map((row) => {
    const next = row.slice(0, lastCol + 1)
    while (next.length <= lastCol) next.push(emptyCell())
    return next
  })
}

export async function parseWorkbook(file: File): Promise<WorkbookData> {
  const lower = file.name.toLowerCase()
  const isText = lower.endsWith('.csv') || lower.endsWith('.tsv') || lower.endsWith('.txt')

  let wb: XLSX.WorkBook
  let buffer: ArrayBuffer | undefined
  if (isText) {
    const text = await file.text()
    wb = XLSX.read(text, {
      type: 'string',
      cellDates: true,
      FS: lower.endsWith('.tsv') ? '\t' : undefined,
    })
  } else {
    const data = await file.arrayBuffer()
    buffer = data.slice(0)
    wb = XLSX.read(data, {
      type: 'array',
      cellDates: true,
      cellFormula: true,
    })
  }

  const sheets = wb.SheetNames.map((name) => parseSheet(wb.Sheets[name]!, name))
  const keepBuffer = lower.endsWith('.xlsx') || lower.endsWith('.xlsm')

  return {
    fileName: file.name,
    fileSize: file.size,
    sheets,
    buffer: keepBuffer ? buffer : undefined,
  }
}

export const ACCEPTED_EXTENSIONS = '.xlsx,.xls,.xlsm,.xlsb,.ods,.csv,.tsv'
