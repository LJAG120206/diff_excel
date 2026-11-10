import type {
  CellData,
  CellDiff,
  ColumnAlign,
  DiffOptions,
  Hunk,
  RowDiff,
  SheetData,
  SheetDiff,
  SheetStats,
  WorkbookData,
  WorkbookDiff,
} from '../types.ts'
import { cellText, colLetter, emptyCell, isEmptyCell, valuesEqual } from './format.ts'

type Edit = { type: 'eq' | 'del' | 'ins'; a?: number; b?: number }

function myers(a: string[], b: string[]): Edit[] {
  const n = a.length
  const m = b.length
  if (n === 0) return b.map((_, i) => ({ type: 'ins' as const, b: i }))
  if (m === 0) return a.map((_, i) => ({ type: 'del' as const, a: i }))

  const max = Math.min(n + m, 8000)
  const offset = n + m
  const v = new Int32Array(2 * offset + 1)
  const trace: Int32Array[] = []

  let dFound = -1
  outer: for (let d = 0; d <= max; d++) {
    trace.push(Int32Array.from(v))
    for (let k = -d; k <= d; k += 2) {
      const idx = k + offset
      let x: number
      if (k === -d || (k !== d && v[idx - 1]! < v[idx + 1]!)) {
        x = v[idx + 1]!
      } else {
        x = v[idx - 1]! + 1
      }
      let y = x - k
      while (x < n && y < m && a[x] === b[y]) {
        x++
        y++
      }
      v[idx] = x
      if (x >= n && y >= m) {
        dFound = d
        break outer
      }
    }
  }

  if (dFound < 0) return positionalAlign(a, b)
  return backtrack(trace, a, b, dFound, offset)
}

function positionalAlign(a: string[], b: string[]): Edit[] {
  const edits: Edit[] = []
  const n = Math.min(a.length, b.length)
  for (let i = 0; i < n; i++) {
    if (a[i] === b[i]) edits.push({ type: 'eq', a: i, b: i })
    else {
      edits.push({ type: 'del', a: i })
      edits.push({ type: 'ins', b: i })
    }
  }
  for (let i = n; i < a.length; i++) edits.push({ type: 'del', a: i })
  for (let i = n; i < b.length; i++) edits.push({ type: 'ins', b: i })
  return edits
}

function backtrack(
  trace: Int32Array[],
  a: string[],
  b: string[],
  d: number,
  offset: number,
): Edit[] {
  const edits: Edit[] = []
  let x = a.length
  let y = b.length

  for (let depth = d; depth > 0; depth--) {
    const v = trace[depth]!
    const k = x - y
    const idx = k + offset
    let prevK: number
    if (k === -depth || (k !== depth && v[idx - 1]! < v[idx + 1]!)) {
      prevK = k + 1
    } else {
      prevK = k - 1
    }
    const prevX = v[prevK + offset]!
    const prevY = prevX - prevK

    while (x > prevX && y > prevY) {
      x--
      y--
      edits.push({ type: 'eq', a: x, b: y })
    }
    if (depth === 0) break
    if (x === prevX) {
      y--
      edits.push({ type: 'ins', b: y })
    } else {
      x--
      edits.push({ type: 'del', a: x })
    }
  }

  while (x > 0 && y > 0) {
    x--
    y--
    edits.push({ type: 'eq', a: x, b: y })
  }
  while (x > 0) {
    x--
    edits.push({ type: 'del', a: x })
  }
  while (y > 0) {
    y--
    edits.push({ type: 'ins', b: y })
  }

  return edits.reverse()
}

function headerLabel(row: CellData[] | undefined, index: number): string {
  const cell = row?.[index]
  const text = cellText(cell).trim()
  return text || colLetter(index)
}

function headerToken(row: CellData[] | undefined, index: number, useHeader: boolean, side: 'old' | 'new'): string {
  if (!useHeader) return `\0${side}:${index}`
  const text = cellText(row?.[index]).trim()
  if (!text) return `\0${side}:${index}`
  return text
}

function cellToken(sheet: SheetData, rowIndex: number, columnIndex: number, options: DiffOptions): string {
  const cell = getCell(sheet.rows[rowIndex], columnIndex)
  const value = normalizeCell(cell, options, false)
  if (value) return value
  if (cell.formula) return `=${cell.formula.replace(/\s+/g, '')}`
  return ''
}

function columnSimilarity(
  oldSheet: SheetData,
  oldIndex: number,
  newSheet: SheetData,
  newIndex: number,
  options: DiffOptions,
): number {
  const start = options.headerRow ? 1 : 0
  let best = 0
  let sawValue = false
  for (let offset = -2; offset <= 2; offset++) {
    let compared = 0
    let equal = 0
    for (let rowIndex = start; rowIndex < oldSheet.rows.length; rowIndex++) {
      const newRow = rowIndex + offset
      if (newRow < start || newRow >= newSheet.rows.length) continue
      const oldToken = cellToken(oldSheet, rowIndex, oldIndex, options)
      const newToken = cellToken(newSheet, newRow, newIndex, options)
      if (!oldToken && !newToken) continue
      sawValue = true
      compared++
      if (oldToken === newToken) equal++
    }
    if (compared > 0) best = Math.max(best, equal / compared)
  }
  if (!sawValue && Math.abs(oldIndex - newIndex) <= 1) return 0.5
  return best
}

function pairShiftedColumns(
  oldIdxs: number[],
  newIdxs: number[],
  oldSheet: SheetData,
  newSheet: SheetData,
  options: DiffOptions,
): { oldIndex: number | null; newIndex: number | null }[] {
  const n = oldIdxs.length
  const m = newIdxs.length
  if (n === 0) return newIdxs.map((newIndex) => ({ oldIndex: null, newIndex }))
  if (m === 0) return oldIdxs.map((oldIndex) => ({ oldIndex, newIndex: null }))

  const gap = 0.62
  const score: number[][] = Array.from({ length: n + 1 }, () => new Array<number>(m + 1).fill(0))
  const choice: number[][] = Array.from({ length: n + 1 }, () => new Array<number>(m + 1).fill(0))
  for (let i = 1; i <= n; i++) {
    for (let j = 1; j <= m; j++) {
      const similarity = columnSimilarity(oldSheet, oldIdxs[i - 1]!, newSheet, newIdxs[j - 1]!, options)
      const matched = similarity >= gap ? score[i - 1]![j - 1]! + similarity : -1
      const dropOld = score[i - 1]![j]!
      const dropNew = score[i]![j - 1]!
      if (matched >= dropOld && matched >= dropNew) {
        score[i]![j] = matched
        choice[i]![j] = 1
      } else if (dropOld >= dropNew) {
        score[i]![j] = dropOld
        choice[i]![j] = 2
      } else {
        score[i]![j] = dropNew
        choice[i]![j] = 3
      }
    }
  }

  const reversed: { oldIndex: number | null; newIndex: number | null }[] = []
  let i = n
  let j = m
  while (i > 0 || j > 0) {
    const step = i > 0 && j > 0 ? choice[i]![j]! : i > 0 ? 2 : 3
    if (step === 1) {
      reversed.push({ oldIndex: oldIdxs[i - 1]!, newIndex: newIdxs[j - 1]! })
      i--
      j--
    } else if (step === 2) {
      reversed.push({ oldIndex: oldIdxs[i - 1]!, newIndex: null })
      i--
    } else {
      reversed.push({ oldIndex: null, newIndex: newIdxs[j - 1]! })
      j--
    }
  }
  return reversed.reverse()
}

function weakenRepeatedHeaders(tokens: string[], side: 'old' | 'new'): string[] {
  const counts = new Map<string, number>()
  for (const token of tokens) {
    if (token.startsWith('\0')) continue
    counts.set(token, (counts.get(token) ?? 0) + 1)
  }
  return tokens.map((token, index) => {
    if (token.startsWith('\0')) return token
    if ((counts.get(token) ?? 0) > 1) return `\0${side}:dup:${index}`
    return token
  })
}

function sheetWidth(sheet: SheetData): number {
  let width = 0
  for (const row of sheet.rows) width = Math.max(width, row.length)
  return width
}

function alignColumns(oldSheet: SheetData, newSheet: SheetData, options: DiffOptions): ColumnAlign[] {
  const oldWidth = sheetWidth(oldSheet)
  const newWidth = sheetWidth(newSheet)
  const oldTokens = weakenRepeatedHeaders(
    Array.from({ length: oldWidth }, (_, i) => headerToken(oldSheet.rows[0], i, options.headerRow, 'old')),
    'old',
  )
  const newTokens = weakenRepeatedHeaders(
    Array.from({ length: newWidth }, (_, i) => headerToken(newSheet.rows[0], i, options.headerRow, 'new')),
    'new',
  )
  const edits = myers(oldTokens, newTokens)
  const links: { oldIndex: number | null; newIndex: number | null }[] = []
  let cursor = 0
  while (cursor < edits.length) {
    const edit = edits[cursor]!
    if (edit.type === 'eq') {
      links.push({ oldIndex: edit.a!, newIndex: edit.b! })
      cursor++
      continue
    }
    const oldIdxs: number[] = []
    const newIdxs: number[] = []
    while (cursor < edits.length && edits[cursor]!.type !== 'eq') {
      const current = edits[cursor]!
      if (current.type === 'del') oldIdxs.push(current.a!)
      else if (current.b != null) newIdxs.push(current.b)
      cursor++
    }
    links.push(...pairShiftedColumns(oldIdxs, newIdxs, oldSheet, newSheet, options))
  }

  let letterIndex = 0
  return links.map((link) => {
    const letter = colLetter(letterIndex++)
    const newHeader = link.newIndex != null ? cellText(newSheet.rows[0]?.[link.newIndex]).trim() : ''
    const oldHeader = link.oldIndex != null ? cellText(oldSheet.rows[0]?.[link.oldIndex]).trim() : ''
    const kind = link.oldIndex == null ? 'added' : link.newIndex == null ? 'removed' : 'same'
    return {
      header: (options.headerRow ? newHeader || oldHeader : '') || letter,
      letter,
      oldIndex: link.oldIndex,
      newIndex: link.newIndex,
      kind,
    }
  })
}

function getCell(row: CellData[] | undefined, index: number | null): CellData {
  if (index == null || !row) return emptyCell()
  return row[index] ?? emptyCell()
}

function fingerprintProjected(cells: CellData[], shared: boolean[], options: DiffOptions): string {
  const parts: string[] = []
  for (let i = 0; i < cells.length; i++) {
    if (!shared[i]) continue
    parts.push(normalizeCell(cells[i]!, options, options.compareFormulas))
  }
  return parts.join('\u001f')
}

function normalizeCell(cell: CellData, options: DiffOptions, useFormulas: boolean): string {
  if (useFormulas && cell.formula) {
    let f = cell.formula
    if (options.ignoreCase) f = f.toLocaleLowerCase('fr')
    if (options.ignoreWhitespace) f = f.replace(/\s+/g, '')
    return `=${f}`
  }
  const v = cell.value
  if (v == null) return ''
  if (typeof v === 'number') return `#${v}`
  if (v instanceof Date) return `@${v.getTime()}`
  if (typeof v === 'boolean') return v ? 'T' : 'F'
  let s = String(v)
  if (options.ignoreWhitespace) s = s.trim().replace(/\s+/g, ' ')
  if (options.ignoreCase) s = s.toLocaleLowerCase('fr')
  return s
}

function projectRow(row: CellData[] | undefined, indices: (number | null)[]): CellData[] {
  return indices.map((i) => getCell(row, i))
}

function similarity(a: CellData[], b: CellData[], shared: boolean[], options: DiffOptions): number {
  let considered = 0
  let matched = 0
  for (let i = 0; i < a.length; i++) {
    if (!shared[i]) continue
    considered++
    const av = options.compareFormulas ? a[i]?.formula ?? a[i]?.value : a[i]?.value
    const bv = options.compareFormulas ? b[i]?.formula ?? b[i]?.value : b[i]?.value
    if (options.compareFormulas && (a[i]?.formula || b[i]?.formula)) {
      if ((a[i]?.formula ?? '') === (b[i]?.formula ?? '')) matched++
      continue
    }
    if (valuesEqual(av, bv, options)) matched++
  }
  if (considered === 0) return 0
  return matched / considered
}

function firstSharedEqual(a: CellData[], b: CellData[], shared: boolean[], options: DiffOptions): boolean {
  for (let i = 0; i < a.length; i++) {
    if (!shared[i]) continue
    if (isEmptyCell(a[i]) && isEmptyCell(b[i])) continue
    if (isEmptyCell(a[i]) || isEmptyCell(b[i])) return false
    return valuesEqual(a[i]!.value, b[i]!.value, options)
  }
  return false
}

interface Pair {
  oldIndex: number | null
  newIndex: number | null
  oldCells: CellData[] | null
  newCells: CellData[] | null
}

function matchBlock(
  dels: { a: number; cells: CellData[] }[],
  inss: { b: number; cells: CellData[] }[],
  shared: boolean[],
  options: DiffOptions,
): Pair[] {
  if (dels.length === 0) {
    return inss.map((row) => ({
      oldIndex: null,
      newIndex: row.b,
      oldCells: null,
      newCells: row.cells,
    }))
  }
  if (inss.length === 0) {
    return dels.map((row) => ({
      oldIndex: row.a,
      newIndex: null,
      oldCells: row.cells,
      newCells: null,
    }))
  }

  const usedDel = new Set<number>()
  const usedIns = new Set<number>()
  const matchForDel = new Array<number>(dels.length).fill(-1)

  const candidates: { di: number; ii: number; sim: number }[] = []
  for (let di = 0; di < dels.length; di++) {
    for (let ii = 0; ii < inss.length; ii++) {
      const sim = similarity(dels[di]!.cells, inss[ii]!.cells, shared, options)
      const keyBoost = firstSharedEqual(dels[di]!.cells, inss[ii]!.cells, shared, options)
      const score = keyBoost ? Math.max(sim, 0.52) : sim
      if (score >= 0.45) candidates.push({ di, ii, sim: score })
    }
  }
  candidates.sort((x, y) => y.sim - x.sim)
  for (const c of candidates) {
    if (usedDel.has(c.di) || usedIns.has(c.ii)) continue
    let crosses = false
    for (let dj = 0; dj < matchForDel.length; dj++) {
      const ij = matchForDel[dj]!
      if (ij < 0) continue
      if ((c.di - dj) * (c.ii - ij) < 0) {
        crosses = true
        break
      }
    }
    if (crosses) continue
    usedDel.add(c.di)
    usedIns.add(c.ii)
    matchForDel[c.di] = c.ii
  }

  const out: Pair[] = []
  let di = 0
  let ii = 0
  while (di < dels.length || ii < inss.length) {
    const want = di < dels.length ? matchForDel[di]! : -1
    if (di < dels.length && want === ii) {
      out.push({
        oldIndex: dels[di]!.a,
        newIndex: inss[ii]!.b,
        oldCells: dels[di]!.cells,
        newCells: inss[ii]!.cells,
      })
      di++
      ii++
    } else if (di < dels.length && want === -1) {
      out.push({
        oldIndex: dels[di]!.a,
        newIndex: null,
        oldCells: dels[di]!.cells,
        newCells: null,
      })
      di++
    } else if (di < dels.length && want > ii) {
      out.push({
        oldIndex: null,
        newIndex: inss[ii]!.b,
        oldCells: null,
        newCells: inss[ii]!.cells,
      })
      ii++
    } else if (ii < inss.length) {
      out.push({
        oldIndex: null,
        newIndex: inss[ii]!.b,
        oldCells: null,
        newCells: inss[ii]!.cells,
      })
      ii++
    } else {
      out.push({
        oldIndex: dels[di]!.a,
        newIndex: null,
        oldCells: dels[di]!.cells,
        newCells: null,
      })
      di++
    }
  }
  return out
}

function buildPairs(
  oldRows: CellData[][],
  newRows: CellData[][],
  columns: ColumnAlign[],
  options: DiffOptions,
  skipHeader: boolean,
): Pair[] {
  const start = skipHeader ? 1 : 0
  const oldSlice = oldRows.slice(start)
  const newSlice = newRows.slice(start)
  const oldIdx = columns.map((c) => c.oldIndex)
  const newIdx = columns.map((c) => c.newIndex)
  const shared = columns.map((c) => c.oldIndex != null && c.newIndex != null)

  const oldProjected = oldSlice.map((row) => projectRow(row, oldIdx))
  const newProjected = newSlice.map((row) => projectRow(row, newIdx))

  const oldKeys = oldProjected.map((cells) => fingerprintProjected(cells, shared, options))
  const newKeys = newProjected.map((cells) => fingerprintProjected(cells, shared, options))
  const edits = myers(oldKeys, newKeys)

  const pairs: Pair[] = []
  if (skipHeader) {
    pairs.push({
      oldIndex: oldRows.length ? 0 : null,
      newIndex: newRows.length ? 0 : null,
      oldCells: oldRows.length ? projectRow(oldRows[0], oldIdx) : null,
      newCells: newRows.length ? projectRow(newRows[0], newIdx) : null,
    })
  }

  let i = 0
  while (i < edits.length) {
    const edit = edits[i]!
    if (edit.type === 'eq') {
      pairs.push({
        oldIndex: edit.a! + start,
        newIndex: edit.b! + start,
        oldCells: oldProjected[edit.a!]!,
        newCells: newProjected[edit.b!]!,
      })
      i++
      continue
    }
    const dels: { a: number; cells: CellData[] }[] = []
    const inss: { b: number; cells: CellData[] }[] = []
    while (i < edits.length && edits[i]!.type !== 'eq') {
      const e = edits[i]!
      if (e.type === 'del') dels.push({ a: e.a! + start, cells: oldProjected[e.a!]! })
      else inss.push({ b: e.b! + start, cells: newProjected[e.b!]! })
      i++
    }
    pairs.push(...matchBlock(dels, inss, shared, options))
  }

  return pairs
}

function cellKindAndTexts(
  col: ColumnAlign,
  oldCell: CellData | null,
  newCell: CellData | null,
  rowKind: 'context' | 'added' | 'removed' | 'changed',
  options: DiffOptions,
): CellDiff {
  const oldText = oldCell ? cellText(oldCell) : ''
  const newText = newCell ? cellText(newCell) : ''
  const oldFormula = oldCell?.formula
  const newFormula = newCell?.formula

  if (rowKind === 'added') {
    return {
      kind: 'added',
      oldValue: null,
      newValue: newCell?.value ?? null,
      oldText: '',
      newText,
      oldFormula,
      newFormula,
    }
  }
  if (rowKind === 'removed') {
    return {
      kind: 'removed',
      oldValue: oldCell?.value ?? null,
      newValue: null,
      oldText,
      newText: '',
      oldFormula,
      newFormula,
    }
  }
  if (col.kind === 'added') {
    return {
      kind: 'added',
      oldValue: null,
      newValue: newCell?.value ?? null,
      oldText: '',
      newText,
      oldFormula,
      newFormula,
    }
  }
  if (col.kind === 'removed') {
    return {
      kind: 'removed',
      oldValue: oldCell?.value ?? null,
      newValue: null,
      oldText,
      newText: '',
      oldFormula,
      newFormula,
    }
  }

  const oldV = options.compareFormulas ? oldFormula ?? oldCell?.value : oldCell?.value
  const newV = options.compareFormulas ? newFormula ?? newCell?.value : newCell?.value
  const equal = options.compareFormulas
    ? (oldFormula ?? '') === (newFormula ?? '') && valuesEqual(oldCell?.value, newCell?.value, options)
    : valuesEqual(oldV, newV, options)

  if (equal) {
    return {
      kind: 'same',
      oldValue: oldCell?.value ?? null,
      newValue: newCell?.value ?? null,
      oldText,
      newText,
      oldFormula,
      newFormula,
    }
  }
  return {
    kind: 'changed',
    oldValue: oldCell?.value ?? null,
    newValue: newCell?.value ?? null,
    oldText,
    newText,
    oldFormula,
    newFormula,
  }
}

function rowsFromPairs(pairs: Pair[], columns: ColumnAlign[], options: DiffOptions): RowDiff[] {
  const shared = columns.map((c) => c.oldIndex != null && c.newIndex != null)
  return pairs.map((pair) => {
    let kind: RowDiff['kind']
    if (pair.oldIndex == null) kind = 'added'
    else if (pair.newIndex == null) kind = 'removed'
    else kind = 'context'

    const cells = columns.map((col, i) =>
      cellKindAndTexts(col, pair.oldCells?.[i] ?? null, pair.newCells?.[i] ?? null, kind, options),
    )

    if (kind === 'context') {
      const dataChanged = cells.some((c, i) => shared[i] && (c.kind === 'changed'))
      if (dataChanged) kind = 'changed'
    }

    return {
      kind,
      oldIndex: pair.oldIndex,
      newIndex: pair.newIndex,
      cells,
    }
  })
}

function buildHunks(rows: RowDiff[], context: number, skipHeader: boolean): Hunk[] {
  const data = skipHeader ? rows.slice(1) : rows
  if (data.length === 0) return []

  const changeIdx: number[] = []
  for (let i = 0; i < data.length; i++) {
    if (data[i]!.kind !== 'context') changeIdx.push(i)
  }
  if (changeIdx.length === 0) return []

  const ranges: { start: number; end: number }[] = []
  let start = Math.max(0, changeIdx[0]! - context)
  let last = changeIdx[0]!
  for (let i = 1; i < changeIdx.length; i++) {
    const idx = changeIdx[i]!
    if (idx - last <= context * 2 + 1) {
      last = idx
    } else {
      ranges.push({ start, end: Math.min(data.length, last + 1 + context) })
      start = Math.max(0, idx - context)
      last = idx
    }
  }
  ranges.push({ start, end: Math.min(data.length, last + 1 + context) })

  return ranges.map((range) => {
    const slice = data.slice(range.start, range.end)
    const oldIndexes = slice.map((r) => r.oldIndex).filter((n): n is number => n != null)
    const newIndexes = slice.map((r) => r.newIndex).filter((n): n is number => n != null)
    const oldStart = oldIndexes.length ? oldIndexes[0]! + 1 : (slice[0]?.newIndex ?? 0) + 1
    const newStart = newIndexes.length ? newIndexes[0]! + 1 : (slice[0]?.oldIndex ?? 0) + 1
    return {
      oldStart,
      oldCount: slice.filter((r) => r.oldIndex != null).length,
      newStart,
      newCount: slice.filter((r) => r.newIndex != null).length,
      rows: slice,
    }
  })
}

function statsFromRows(rows: RowDiff[], columns: ColumnAlign[], skipHeader: boolean): SheetStats {
  const data = skipHeader ? rows.slice(1) : rows
  let added = 0
  let removed = 0
  let changed = 0
  for (const row of data) {
    if (row.kind === 'added') added++
    else if (row.kind === 'removed') removed++
    else if (row.kind === 'changed') changed++
  }
  return {
    added,
    removed,
    changed,
    columnsAdded: columns.filter((c) => c.kind === 'added').length,
    columnsRemoved: columns.filter((c) => c.kind === 'removed').length,
  }
}

function diffSheet(
  oldSheet: SheetData | undefined,
  newSheet: SheetData | undefined,
  name: string,
  options: DiffOptions,
): SheetDiff {
  if (!oldSheet && newSheet) {
    const columns: ColumnAlign[] = Array.from(
      { length: newSheet.rows[0]?.length ?? 0 },
      (_, i) => ({
        header: options.headerRow ? headerLabel(newSheet.rows[0], i) : colLetter(i),
        letter: colLetter(i),
        oldIndex: null,
        newIndex: i,
        kind: 'same' as const,
      }),
    )
    const allRows: RowDiff[] = newSheet.rows.map((row, idx) => ({
      kind: 'added',
      oldIndex: null,
      newIndex: idx,
      cells: row.map((cell) => ({
        kind: 'added' as const,
        oldValue: null,
        newValue: cell.value,
        oldText: '',
        newText: cellText(cell),
        newFormula: cell.formula,
      })),
    }))
    const hunks =
      allRows.length === 0
        ? []
        : [
            {
              oldStart: 0,
              oldCount: 0,
              newStart: 1,
              newCount: allRows.length,
              rows: options.headerRow ? allRows.slice(1) : allRows,
            },
          ]
    return {
      name,
      status: 'added',
      columns,
      hunks,
      allRows,
      stats: statsFromRows(allRows, columns, options.headerRow),
    }
  }

  if (oldSheet && !newSheet) {
    const columns: ColumnAlign[] = Array.from(
      { length: oldSheet.rows[0]?.length ?? 0 },
      (_, i) => ({
        header: options.headerRow ? headerLabel(oldSheet.rows[0], i) : colLetter(i),
        letter: colLetter(i),
        oldIndex: i,
        newIndex: null,
        kind: 'same' as const,
      }),
    )
    const allRows: RowDiff[] = oldSheet.rows.map((row, idx) => ({
      kind: 'removed',
      oldIndex: idx,
      newIndex: null,
      cells: row.map((cell) => ({
        kind: 'removed' as const,
        oldValue: cell.value,
        newValue: null,
        oldText: cellText(cell),
        newText: '',
        oldFormula: cell.formula,
      })),
    }))
    const hunks =
      allRows.length === 0
        ? []
        : [
            {
              oldStart: 1,
              oldCount: allRows.length,
              newStart: 0,
              newCount: 0,
              rows: options.headerRow ? allRows.slice(1) : allRows,
            },
          ]
    return {
      name,
      status: 'removed',
      columns,
      hunks,
      allRows,
      stats: statsFromRows(allRows, columns, options.headerRow),
    }
  }

  const columns = alignColumns(oldSheet!, newSheet!, options)
  const pairs = buildPairs(oldSheet!.rows, newSheet!.rows, columns, options, options.headerRow)
  const allRows = rowsFromPairs(pairs, columns, options)
  const stats = statsFromRows(allRows, columns, options.headerRow)
  const status: SheetDiff['status'] =
    stats.added === 0 &&
    stats.removed === 0 &&
    stats.changed === 0 &&
    stats.columnsAdded === 0 &&
    stats.columnsRemoved === 0
      ? 'unchanged'
      : 'modified'

  return {
    name,
    status,
    columns,
    hunks: buildHunks(allRows, options.contextLines, options.headerRow),
    allRows,
    stats,
  }
}

export const defaultOptions: DiffOptions = {
  headerRow: true,
  ignoreWhitespace: true,
  ignoreCase: false,
  compareFormulas: false,
  contextLines: 3,
}

export function diffWorkbooks(
  oldWb: WorkbookData,
  newWb: WorkbookData,
  options: DiffOptions = defaultOptions,
): WorkbookDiff {
  const names: string[] = []
  const seen = new Set<string>()
  for (const s of oldWb.sheets) {
    names.push(s.name)
    seen.add(s.name)
  }
  for (const s of newWb.sheets) {
    if (!seen.has(s.name)) names.push(s.name)
  }

  const oldMap = new Map(oldWb.sheets.map((s) => [s.name, s]))
  const newMap = new Map(newWb.sheets.map((s) => [s.name, s]))

  const sheets = names.map((name) => diffSheet(oldMap.get(name), newMap.get(name), name, options))

  const totals: WorkbookDiff['stats'] = {
    added: 0,
    removed: 0,
    changed: 0,
    columnsAdded: 0,
    columnsRemoved: 0,
    sheetsAdded: 0,
    sheetsRemoved: 0,
    sheetsModified: 0,
  }
  for (const sheet of sheets) {
    totals.added += sheet.stats.added
    totals.removed += sheet.stats.removed
    totals.changed += sheet.stats.changed
    totals.columnsAdded += sheet.stats.columnsAdded
    totals.columnsRemoved += sheet.stats.columnsRemoved
    if (sheet.status === 'added') totals.sheetsAdded++
    if (sheet.status === 'removed') totals.sheetsRemoved++
    if (sheet.status === 'modified') totals.sheetsModified++
  }

  return {
    oldName: oldWb.fileName,
    newName: newWb.fileName,
    sheets,
    stats: totals,
  }
}
