export type ExcelCellValue = string | number | boolean | null

export interface ExcelColumn {
  key: string
  title: string
  index: number
}

export interface ExcelDataRow {
  __rowId: string
  __sourceRow: number
  [key: string]: ExcelCellValue | string | number
}

export interface SkippedRow {
  rowIndex: number
  reason: string
  preview: string
}

export interface ParsedWorksheet {
  sheetName: string
  headerRowIndex: number
  columns: ExcelColumn[]
  rows: ExcelDataRow[]
  skippedRows: SkippedRow[]
  rawRowCount: number
}

export interface ParseWorkbookOptions {
  sheetName?: string
  headerRowIndex?: number
  summaryKeywords?: string[]
}

export interface HasuraImportConfig {
  endpoint: string
  adminSecret: string
}

export interface DictItem {
  id: string
  dict_group_id: string
  name: string
  serial: number | null
  code: string | null
}
