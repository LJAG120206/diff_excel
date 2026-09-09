import type { CellData, SheetData, WorkbookData } from '../types.ts'
import { formatValue } from './format.ts'

function cell(value: unknown): CellData {
  return { value, text: formatValue(value) }
}

function formula(value: unknown, f: string): CellData {
  return { value, text: formatValue(value), formula: f }
}

function sheet(name: string, rows: unknown[][]): SheetData {
  return {
    name,
    rows: rows.map((row) => row.map((v) => (typeof v === 'object' && v && 'text' in v ? (v as CellData) : cell(v)))),
  }
}

export function createSampleWorkbooks(): { oldWb: WorkbookData; newWb: WorkbookData } {
  const oldWb: WorkbookData = {
    fileName: 'ventes_v1.xlsx',
    fileSize: 18432,
    sheets: [
      sheet('Ventes', [
        ['Produit', 'Quantité', 'Prix', 'Région'],
        ['Pommes', 100, 1.2, 'Nord'],
        ['Poires', 50, 1.8, 'Sud'],
        ['Bananes', 80, 0.9, 'Nord'],
        ['Kiwis', 20, 2.5, 'Est'],
        ['Oranges', 60, 1.1, 'Ouest'],
      ]),
      sheet('Stock', [
        ['SKU', 'Entrepôt', 'Qté'],
        ['A-100', 'Lyon', 240],
        ['A-200', 'Lille', 80],
        ['B-010', 'Lyon', 12],
      ]),
      sheet('Notes', [
        ['Auteur', 'Commentaire'],
        ['Alice', 'Première version du rapport'],
        ['Bob', 'Vérifier les kiwis'],
      ]),
    ],
  }

  const newWb: WorkbookData = {
    fileName: 'ventes_v2.xlsx',
    fileSize: 20110,
    sheets: [
      sheet('Ventes', [
        ['Produit', 'Quantité', 'Prix', 'Remise', 'Région'],
        ['Pommes', 120, 1.2, '5%', 'Nord'],
        ['Poires', 50, 2.0, '', 'Sud'],
        ['Bananes', 80, 0.9, '10%', 'Nord'],
        ['Oranges', 60, 1.1, '', 'Ouest'],
        ['Mangues', 40, 3.1, '', 'Ouest'],
      ]),
      sheet('Stock', [
        ['SKU', 'Entrepôt', 'Qté'],
        ['A-100', 'Lyon', 240],
        ['A-200', 'Lille', 95],
        ['B-010', 'Lyon', 12],
        ['C-300', 'Nantes', 4],
      ]),
      sheet('RH', [
        ['Nom', 'Équipe', 'Jours'],
        ['Camille', 'Ventes', formula(18, 'B2+3')],
        ['Diego', 'Logistique', 20],
      ]),
    ],
  }

  return { oldWb, newWb }
}
