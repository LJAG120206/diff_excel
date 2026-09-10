export interface CellData {
  value: unknown
  text: string
  formula?: string
}

export interface SheetData {
  name: string
  rows: CellData[][]
}

export interface WorkbookData {
  fileName: string
  fileSize: number
  sheets: SheetData[]
  buffer?: ArrayBuffer
}

export interface DiffOptions {
  headerRow: boolean
  ignoreWhitespace: boolean
  ignoreCase: boolean
  compareFormulas: boolean
  contextLines: number
}

export type ChangeKind = 'same' | 'added' | 'removed' | 'changed'
export type RowKind = 'context' | 'added' | 'removed' | 'changed'
export type SheetStatus = 'unchanged' | 'added' | 'removed' | 'modified'

export interface ColumnAlign {
  header: string
  letter: string
  oldIndex: number | null
  newIndex: number | null
  kind: 'same' | 'added' | 'removed'
}

export interface CellDiff {
  kind: ChangeKind
  oldValue: unknown
  newValue: unknown
  oldText: string
  newText: string
  oldFormula?: string
  newFormula?: string
}

export interface RowDiff {
  kind: RowKind
  oldIndex: number | null
  newIndex: number | null
  cells: CellDiff[]
}

export interface Hunk {
  oldStart: number
  oldCount: number
  newStart: number
  newCount: number
  rows: RowDiff[]
}

export interface SheetStats {
  added: number
  removed: number
  changed: number
  columnsAdded: number
  columnsRemoved: number
}

export interface SheetDiff {
  name: string
  status: SheetStatus
  columns: ColumnAlign[]
  hunks: Hunk[]
  allRows: RowDiff[]
  stats: SheetStats
}

export interface WorkbookDiff {
  oldName: string
  newName: string
  sheets: SheetDiff[]
  stats: SheetStats & { sheetsAdded: number; sheetsRemoved: number; sheetsModified: number }
}

export type ViewMode = 'unified' | 'split'
export type Theme = 'dark' | 'light'

export type MergeSide = 'current' | 'incoming'
export type Resolution = 'pending' | MergeSide

export interface MergeDecisions {
  sheets: Record<string, Resolution>
  columns: Record<string, Resolution>
  rows: Record<string, Resolution>
  cells: Record<string, Resolution>
}
