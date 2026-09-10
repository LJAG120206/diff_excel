import * as ExcelJS from 'exceljs'

type WorkbookCtor = typeof ExcelJS.Workbook

function resolveWorkbook(): WorkbookCtor {
  const mod = ExcelJS as unknown as {
    Workbook?: WorkbookCtor
    default?: { Workbook?: WorkbookCtor } | WorkbookCtor
  }
  if (typeof mod.Workbook === 'function') return mod.Workbook
  const exported = mod.default
  if (exported && typeof (exported as { Workbook?: unknown }).Workbook === 'function') {
    return (exported as { Workbook: WorkbookCtor }).Workbook
  }
  if (typeof exported === 'function') return exported as WorkbookCtor
  throw new Error('ExcelJS.Workbook introuvable')
}

export const Workbook = resolveWorkbook()
