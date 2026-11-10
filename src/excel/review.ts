import type { ColumnAlign, MergeDecisions, Resolution, RowDiff, SheetDiff, WorkbookDiff } from '../types.ts'
import { cellKey, changeRowId, colKey, dataRows, effectiveCell, effectiveSheet, getResolution, rowKey } from './merge.ts'

export type ReviewItem =
  | { id: string; type: 'sheet'; sheetName: string }
  | { id: string; type: 'column'; sheetName: string; letter: string }
  | { id: string; type: 'column-cells'; sheetName: string; letter: string; rowIds: string[] }
  | { id: string; type: 'row'; sheetName: string; rowId: string }
  | { id: string; type: 'cell'; sheetName: string; rowId: string; letter: string }

export function buildReviewItems(diff: WorkbookDiff, headerRow: boolean): ReviewItem[] {
  const items: ReviewItem[] = []
  for (const sheet of diff.sheets) {
    if (sheet.status === 'unchanged') continue
    if (sheet.status === 'added' || sheet.status === 'removed') {
      items.push({ id: `sheet:${sheet.name}`, type: 'sheet', sheetName: sheet.name })
      continue
    }
    for (const column of sheet.columns) {
      if (column.kind === 'added' || column.kind === 'removed') {
        items.push({
          id: `column:${sheet.name}:${column.letter}`,
          type: 'column',
          sheetName: sheet.name,
          letter: column.letter,
        })
      }
    }
    const changedByColumn = new Map<string, string[]>()
    for (const row of dataRows(sheet, headerRow)) {
      if (row.kind !== 'changed') continue
      sheet.columns.forEach((column, index) => {
        if (column.kind !== 'same') return
        if (row.cells[index]?.kind !== 'changed') return
        const list = changedByColumn.get(column.letter) ?? []
        list.push(changeRowId(row))
        changedByColumn.set(column.letter, list)
      })
    }
    for (const column of sheet.columns) {
      const rowIds = changedByColumn.get(column.letter)
      if (!rowIds?.length) continue
      items.push({
        id: `column-cells:${sheet.name}:${column.letter}`,
        type: 'column-cells',
        sheetName: sheet.name,
        letter: column.letter,
        rowIds,
      })
      for (const rowId of rowIds) {
        items.push({
          id: `cell:${sheet.name}:${rowId}:${column.letter}`,
          type: 'cell',
          sheetName: sheet.name,
          rowId,
          letter: column.letter,
        })
      }
    }
    for (const row of dataRows(sheet, headerRow)) {
      if (row.kind !== 'added' && row.kind !== 'removed') continue
      items.push({
        id: `row:${sheet.name}:${changeRowId(row)}`,
        type: 'row',
        sheetName: sheet.name,
        rowId: changeRowId(row),
      })
    }
  }
  return items
}

export function findSheet(diff: WorkbookDiff, name: string): SheetDiff | undefined {
  return diff.sheets.find((sheet) => sheet.name === name)
}

export function findColumn(sheet: SheetDiff, letter: string): ColumnAlign | undefined {
  return sheet.columns.find((column) => column.letter === letter)
}

export function findRow(sheet: SheetDiff, rowId: string): RowDiff | undefined {
  return sheet.allRows.find((row) => changeRowId(row) === rowId)
}

export function rowCaption(sheet: SheetDiff, row: RowDiff): string {
  const line = (row.newIndex ?? row.oldIndex ?? 0) + 1
  const keyIndex = sheet.columns.findIndex((column) => column.kind === 'same')
  const cell = keyIndex >= 0 ? row.cells[keyIndex] : undefined
  const text = (cell?.newText || cell?.oldText || '').trim()
  return text ? `Ligne ${line} · ${text}` : `Ligne ${line}`
}

export function itemResolution(diff: WorkbookDiff, item: ReviewItem, decisions: MergeDecisions): Resolution | 'mixed' {
  const sheet = findSheet(diff, item.sheetName)
  if (!sheet) return 'pending'
  if (item.type === 'sheet') return effectiveSheet(sheet, decisions)
  if (item.type === 'column') {
    const column = findColumn(sheet, item.letter)
    return column ? getResolution(decisions.columns, colKey(sheet.name, column.letter)) : 'pending'
  }
  if (item.type === 'row') {
    const row = findRow(sheet, item.rowId)
    return row ? getResolution(decisions.rows, rowKey(sheet.name, row)) : 'pending'
  }
  if (item.type === 'cell') {
    const row = findRow(sheet, item.rowId)
    const column = findColumn(sheet, item.letter)
    if (!row || !column) return 'pending'
    const explicit = getResolution(decisions.cells, cellKey(sheet.name, row, column.letter))
    return explicit !== 'pending' ? explicit : effectiveCell(sheet.name, row, column, decisions)
  }
  const values = item.rowIds.map((rowId) => {
    const row = findRow(sheet, rowId)
    const column = findColumn(sheet, item.letter)
    return row && column ? effectiveCell(sheet.name, row, column, decisions) : 'pending'
  })
  if (values.every((value) => value === 'pending')) return 'pending'
  if (values.every((value) => value === 'current')) return 'current'
  if (values.every((value) => value === 'incoming')) return 'incoming'
  return 'mixed'
}

