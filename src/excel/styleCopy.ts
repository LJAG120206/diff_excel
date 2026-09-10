import { Workbook } from './exceljsClient.ts'
import type {
  Cell,
  CellValue,
  ConditionalFormattingOptions,
  DataValidation,
  Style,
  Worksheet,
} from 'exceljs'
import type {
  ColumnAlign,
  MergeDecisions,
  MergeSide,
  RowDiff,
  WorkbookData,
  WorkbookDiff,
} from '../types.ts'
import { colLetter } from './format.ts'
import {
  getMergedPlan,
  styleSourceSide,
  type MergeSheetPlan,
} from './merge.ts'

type ExcelWorkbook = InstanceType<typeof Workbook>

interface WorksheetModelLoose {
  id: number
  name: string
  merges?: string[]
  mergeCells?: string[]
  media?: unknown[]
  tables?: unknown[]
  rows?: unknown[]
  cols?: unknown[]
  conditionalFormattings?: ConditionalFormattingOptions[]
  [key: string]: unknown
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
  if (src.note) dest.note = cloneData(src.note)
  if (src.dataValidation) dest.dataValidation = cloneData(src.dataValidation) as DataValidation
}

function copyColumnLayout(srcCol: { width?: number; hidden?: boolean; outlineLevel?: number; style?: Partial<Style> }, destCol: { width?: number; hidden?: boolean; outlineLevel?: number; style?: Partial<Style> }): void {
  if (srcCol.width != null) destCol.width = srcCol.width
  if (srcCol.hidden != null) destCol.hidden = srcCol.hidden
  if (srcCol.outlineLevel != null) destCol.outlineLevel = srcCol.outlineLevel
  if (srcCol.style) destCol.style = cloneData(srcCol.style)
}

function findSheet(wb: ExcelWorkbook, name: string): Worksheet | undefined {
  return wb.getWorksheet(name) ?? wb.worksheets.find((sheet) => sheet.name.toLowerCase() === name.toLowerCase())
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

async function loadWorkbook(buffer: ArrayBuffer): Promise<ExcelWorkbook> {
  const wb = new Workbook()
  await wb.xlsx.load(buffer as unknown as Parameters<ExcelWorkbook['xlsx']['load']>[0])
  return wb
}

function copyWorkbookMeta(src: ExcelWorkbook, dest: ExcelWorkbook): void {
  dest.creator = src.creator
  dest.lastModifiedBy = src.lastModifiedBy
  dest.created = src.created
  dest.modified = new Date()
  dest.lastPrinted = src.lastPrinted
  dest.company = src.company
  dest.manager = src.manager
  dest.title = src.title
  dest.subject = src.subject
  dest.keywords = src.keywords
  dest.category = src.category
  dest.description = src.description
  if (src.views) dest.views = cloneData(src.views)
  if (src.properties) dest.properties = cloneData(src.properties)
  try {
    dest.calcProperties = cloneData(src.calcProperties)
  } catch {
    /* optional */
  }
  try {
    const names = src.definedNames as { model?: unknown }
    if (names.model) (dest.definedNames as { model: unknown }).model = cloneData(names.model)
  } catch {
    /* optional */
  }
}

function copyImages(src: Worksheet, srcWb: ExcelWorkbook, dest: Worksheet, destWb: ExcelWorkbook): void {
  try {
    for (const img of src.getImages()) {
      const media = srcWb.getImage(Number(img.imageId))
      if (!media) continue
      const imageId = destWb.addImage({
        extension: media.extension,
        buffer: media.buffer,
        base64: media.base64,
      })
      dest.addImage(imageId, img.range)
    }
    const backgroundId = src.getBackgroundImageId?.()
    if (backgroundId) {
      const media = srcWb.getImage(Number(backgroundId))
      if (media) {
        const imageId = destWb.addImage({
          extension: media.extension,
          buffer: media.buffer,
          base64: media.base64,
        })
        dest.addBackgroundImage(imageId)
      }
    }
  } catch {
    /* drawings not fully supported */
  }
}

function copyWorksheet(src: Worksheet, srcWb: ExcelWorkbook, destWb: ExcelWorkbook, name: string): Worksheet {
  try {
    return copyWorksheetByModel(src, srcWb, destWb, name)
  } catch {
    return copyWorksheetManual(src, srcWb, destWb, name)
  }
}

function copyWorksheetByModel(src: Worksheet, srcWb: ExcelWorkbook, destWb: ExcelWorkbook, name: string): Worksheet {
  const dest = destWb.addWorksheet(name)
  const model = cloneData(src.model as unknown as WorksheetModelLoose)
  model.id = dest.id
  model.name = dest.name
  model.mergeCells = model.merges ?? model.mergeCells ?? []
  model.media = []
  model.tables = model.tables ?? []
  model.rows = model.rows ?? []
  model.cols = model.cols ?? []
  model.conditionalFormattings = model.conditionalFormattings ?? []
  dest.model = model as unknown as Worksheet['model']
  dest.state = src.state
  copyImages(src, srcWb, dest, destWb)
  return dest
}

function copyWorksheetManual(src: Worksheet, srcWb: ExcelWorkbook, destWb: ExcelWorkbook, name: string): Worksheet {
  const dest = destWb.addWorksheet(name, {
    properties: cloneData(src.properties),
    views: cloneData(src.views),
    pageSetup: cloneData(src.pageSetup),
    headerFooter: cloneData(src.headerFooter),
    state: src.state,
  })
  dest.autoFilter = src.autoFilter
  src.columns.forEach((col) => {
    if (!col || col.number == null) return
    copyColumnLayout(col, dest.getColumn(col.number))
  })
  src.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    const destRow = dest.getRow(rowNumber)
    if (row.height != null) destRow.height = row.height
    if (row.hidden != null) destRow.hidden = row.hidden
    if (row.outlineLevel != null) destRow.outlineLevel = row.outlineLevel
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      copyCell(cell, destRow.getCell(colNumber))
    })
  })
  const merges = (src.model as unknown as WorksheetModelLoose).merges ?? []
  for (const range of merges) {
    try {
      dest.mergeCells(range)
    } catch {
      /* ignore overlapping merges */
    }
  }
  const formattings = (src as unknown as { conditionalFormattings?: ConditionalFormattingOptions[] })
    .conditionalFormattings
  for (const cf of formattings ?? []) {
    dest.addConditionalFormatting(cloneData(cf))
  }
  copyImages(src, srcWb, dest, destWb)
  return dest
}

function sourceCoord(row: RowDiff, col: ColumnAlign, side: MergeSide): { row: number; col: number } | null {
  if (side === 'incoming') {
    if (row.newIndex == null || col.newIndex == null) return null
    return { row: row.newIndex + 1, col: col.newIndex + 1 }
  }
  if (row.oldIndex == null || col.oldIndex == null) return null
  return { row: row.oldIndex + 1, col: col.oldIndex + 1 }
}

function overlayIncoming(
  dest: Worksheet,
  newWb: ExcelWorkbook | undefined,
  plan: MergeSheetPlan,
  decisions: MergeDecisions,
): void {
  const newWs = newWb ? findSheet(newWb, plan.sheet.name) : undefined
  for (const row of plan.sheet.allRows) {
    for (const col of plan.sheet.columns) {
      if (col.oldIndex == null || row.oldIndex == null) continue
      const side = styleSourceSide(plan.sheet, row, col, decisions)
      if (side !== 'incoming') continue
      const from = sourceCoord(row, col, 'incoming')
      if (!from || !newWs) continue
      copyCell(newWs.getRow(from.row).getCell(from.col), dest.getRow(row.oldIndex + 1).getCell(col.oldIndex + 1))
      const srcRow = newWs.getRow(from.row)
      if (srcRow.height) dest.getRow(row.oldIndex + 1).height = srcRow.height
    }
  }
}

function colNumber(letter: string): number {
  let n = 0
  for (const ch of letter.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64)
  return n
}

function parseA1Range(range: string): { r1: number; c1: number; r2: number; c2: number } | null {
  const match = range.match(/^\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$/i)
  if (!match) return null
  return {
    c1: colNumber(match[1]!),
    r1: Number(match[2]),
    c2: colNumber(match[3]!),
    r2: Number(match[4]),
  }
}

function a1(row: number, col: number): string {
  return `${colLetter(col - 1)}${row}`
}

function remapMerges(
  dest: Worksheet,
  src: Worksheet | undefined,
  posMap: Map<string, { row: number; col: number }>,
): void {
  if (!src) return
  const merges = (src.model as unknown as WorksheetModelLoose).merges ?? []
  for (const range of merges) {
    const parsed = parseA1Range(range)
    if (!parsed) continue
    const tl = posMap.get(`${parsed.r1},${parsed.c1}`)
    const br = posMap.get(`${parsed.r2},${parsed.c2}`)
    if (!tl || !br) continue
    if (tl.row > br.row || tl.col > br.col) continue
    try {
      dest.mergeCells(a1(tl.row, tl.col), a1(br.row, br.col))
    } catch {
      /* ignore */
    }
  }
}

function rebuildWorksheet(
  destWb: ExcelWorkbook,
  oldWb: ExcelWorkbook | undefined,
  newWb: ExcelWorkbook | undefined,
  plan: MergeSheetPlan,
  decisions: MergeDecisions,
  name: string,
): Worksheet {
  const oldWs = oldWb ? findSheet(oldWb, plan.sheet.name) : undefined
  const newWs = newWb ? findSheet(newWb, plan.sheet.name) : undefined
  const styleSrc = oldWs ?? newWs
  const dest = destWb.addWorksheet(name, styleSrc
    ? {
        properties: cloneData(styleSrc.properties),
        views: cloneData(styleSrc.views),
        pageSetup: cloneData(styleSrc.pageSetup),
        headerFooter: cloneData(styleSrc.headerFooter),
        state: styleSrc.state,
      }
    : undefined)

  plan.columns.forEach((col, index) => {
    const side: MergeSide = col.kind === 'added' ? 'incoming' : 'current'
    const srcWs = side === 'incoming' ? newWs : oldWs
    const srcIndex = side === 'incoming' ? col.newIndex : col.oldIndex
    if (!srcWs || srcIndex == null) return
    const srcCol = srcWs.getColumn(srcIndex + 1)
    const destCol = dest.getColumn(index + 1)
    copyColumnLayout(srcCol, destCol)
  })

  const oldMap = new Map<string, { row: number; col: number }>()
  const newMap = new Map<string, { row: number; col: number }>()

  plan.rows.forEach((row, rowIndex) => {
    const destRowNumber = rowIndex + 1
    const destRow = dest.getRow(destRowNumber)
    plan.columns.forEach((col, colIndex) => {
      const side = styleSourceSide(plan.sheet, row, col, decisions)
      const from = sourceCoord(row, col, side)
      const srcWs = side === 'incoming' ? newWs : oldWs
      if (from && srcWs) {
        const srcRow = srcWs.getRow(from.row)
        if (colIndex === 0) {
          if (srcRow.height != null) destRow.height = srcRow.height
          if (srcRow.hidden != null) destRow.hidden = srcRow.hidden
          if (srcRow.outlineLevel != null) destRow.outlineLevel = srcRow.outlineLevel
        }
        copyCell(srcRow.getCell(from.col), destRow.getCell(colIndex + 1))
      }
      if (row.oldIndex != null && col.oldIndex != null) {
        oldMap.set(`${row.oldIndex + 1},${col.oldIndex + 1}`, { row: destRowNumber, col: colIndex + 1 })
      }
      if (row.newIndex != null && col.newIndex != null) {
        newMap.set(`${row.newIndex + 1},${col.newIndex + 1}`, { row: destRowNumber, col: colIndex + 1 })
      }
    })
  })

  remapMerges(dest, oldWs, oldMap)
  remapMerges(dest, newWs, newMap)
  return dest
}

function isXlsxLike(name: string): boolean {
  return /\.xlsx$/i.test(name) || /\.xlsm$/i.test(name)
}

export async function writeStyledXlsx(
  diff: WorkbookDiff,
  decisions: MergeDecisions,
  headerRow: boolean,
  oldWb: WorkbookData,
  newWb: WorkbookData,
): Promise<ArrayBuffer> {
  if (!oldWb.buffer && !newWb.buffer) {
    throw new Error('Aucun classeur source avec mise en forme')
  }

  const [currentBook, incomingBook] = await Promise.all([
    oldWb.buffer && isXlsxLike(oldWb.fileName) ? loadWorkbook(oldWb.buffer.slice(0)) : Promise.resolve(undefined),
    newWb.buffer && isXlsxLike(newWb.fileName) ? loadWorkbook(newWb.buffer.slice(0)) : Promise.resolve(undefined),
  ])

  if (!currentBook && !incomingBook) {
    throw new Error('ExcelJS ne peut pas relire ces fichiers')
  }

  const out = new Workbook()
  copyWorkbookMeta(currentBook ?? incomingBook!, out)

  const plans = getMergedPlan(diff, decisions, headerRow)
  if (plans.length === 0) {
    out.addWorksheet('Vide')
  }

  const used = new Set<string>()
  for (const plan of plans) {
    const name = uniqueSheetName(plan.sheet.name, used)
    if (plan.mode === 'clone-current' || plan.mode === 'overlay') {
      const src = currentBook ? findSheet(currentBook, plan.sheet.name) : undefined
      if (src && currentBook) {
        const dest = copyWorksheet(src, currentBook, out, name)
        if (plan.mode === 'overlay') overlayIncoming(dest, incomingBook, plan, decisions)
        continue
      }
    }
    if (plan.mode === 'clone-incoming') {
      const src = incomingBook ? findSheet(incomingBook, plan.sheet.name) : undefined
      if (src && incomingBook) {
        copyWorksheet(src, incomingBook, out, name)
        continue
      }
    }
    rebuildWorksheet(out, currentBook, incomingBook, plan, decisions, name)
  }

  const buffer = await out.xlsx.writeBuffer()
  return buffer as unknown as ArrayBuffer
}
