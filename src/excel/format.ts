import type { CellData } from '../types.ts'

export function colLetter(index: number): string {
  let n = index + 1
  let s = ''
  while (n > 0) {
    n--
    s = String.fromCharCode(65 + (n % 26)) + s
    n = Math.floor(n / 26)
  }
  return s
}

export function isEmptyCell(cell: CellData | null | undefined): boolean {
  if (!cell) return true
  if (cell.formula) return false
  const v = cell.value
  if (v == null) return true
  if (typeof v === 'string' && v.trim() === '') return true
  return false
}

export function formatValue(value: unknown): string {
  if (value == null || value === '') return ''
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) return String(value)
    return new Intl.NumberFormat('fr-FR', { maximumFractionDigits: 10 }).format(value)
  }
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    const hasTime =
      value.getHours() !== 0 || value.getMinutes() !== 0 || value.getSeconds() !== 0
    return new Intl.DateTimeFormat('fr-FR', {
      dateStyle: 'short',
      timeStyle: hasTime ? 'short' : undefined,
    }).format(value)
  }
  if (typeof value === 'boolean') return value ? 'VRAI' : 'FAUX'
  return String(value)
}

export function cellText(cell: CellData | null | undefined): string {
  if (!cell) return ''
  if (cell.text) return cell.text
  return formatValue(cell.value)
}

export function formatBytes(size: number): string {
  if (size < 1024) return `${size} o`
  if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} Ko`
  return `${(size / (1024 * 1024)).toFixed(1)} Mo`
}

export function valuesEqual(
  a: unknown,
  b: unknown,
  options: { ignoreWhitespace: boolean; ignoreCase: boolean },
): boolean {
  if (a == null && b == null) return true
  if (typeof a === 'number' && typeof b === 'number') {
    return Math.abs(a - b) < 1e-9
  }
  if (a instanceof Date && b instanceof Date) {
    return a.getTime() === b.getTime()
  }
  if (typeof a === 'boolean' && typeof b === 'boolean') return a === b
  if (typeof a === 'number' && typeof b === 'string' && b.trim() !== '') {
    const n = Number(b.replace(/\s/g, '').replace(',', '.'))
    if (!Number.isNaN(n)) return Math.abs(a - n) < 1e-9
  }
  if (typeof b === 'number' && typeof a === 'string' && a.trim() !== '') {
    const n = Number(a.replace(/\s/g, '').replace(',', '.'))
    if (!Number.isNaN(n)) return Math.abs(b - n) < 1e-9
  }

  let sa = a == null ? '' : String(a)
  let sb = b == null ? '' : String(b)
  if (options.ignoreWhitespace) {
    sa = sa.trim().replace(/\s+/g, ' ')
    sb = sb.trim().replace(/\s+/g, ' ')
  }
  if (options.ignoreCase) {
    sa = sa.toLocaleLowerCase('fr')
    sb = sb.toLocaleLowerCase('fr')
  }
  return sa === sb
}

export function emptyCell(): CellData {
  return { value: null, text: '' }
}
