import type { WorkbookDiff } from '../types.ts'

export function exportDiffCsv(diff: WorkbookDiff): string {
  const header = [
    'Feuille',
    'Type',
    'Ligne ancienne',
    'Ligne nouvelle',
    'Colonne',
    'En-tête',
    'Ancien',
    'Nouveau',
  ]
  const lines = [header.map(csvEscape).join(';')]

  for (const sheet of diff.sheets) {
    for (const row of sheet.allRows) {
      if (row.kind === 'context') continue
      sheet.columns.forEach((col, i) => {
        const cell = row.cells[i]
        if (!cell) return
        if (row.kind === 'changed' && cell.kind !== 'changed') return
        if (row.kind !== 'changed' && cell.kind === 'same') return

        const type =
          row.kind === 'added'
            ? 'ligne ajoutée'
            : row.kind === 'removed'
              ? 'ligne supprimée'
              : 'cellule modifiée'
        lines.push(
          [
            sheet.name,
            type,
            row.oldIndex != null ? String(row.oldIndex + 1) : '',
            row.newIndex != null ? String(row.newIndex + 1) : '',
            col.letter,
            col.header,
            cell.oldText,
            cell.newText,
          ]
            .map(csvEscape)
            .join(';'),
        )
      })
    }
  }

  return `\uFEFF${lines.join('\n')}`
}

function csvEscape(value: string): string {
  if (/[;"\n\r]/.test(value)) return `"${value.replace(/"/g, '""')}"`
  return value
}

export function downloadText(filename: string, content: string, mime: string): void {
  const blob = new Blob([content], { type: mime })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = filename
  a.click()
  URL.revokeObjectURL(url)
}
