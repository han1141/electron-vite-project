import * as XLSX from 'xlsx'
import type {
  ExcelCellValue,
  ExcelColumn,
  ExcelDataRow,
  ParseWorkbookOptions,
  ParsedWorksheet,
  SkippedRow,
} from '../types/excel'

const DEFAULT_SUMMARY_KEYWORDS = ['合计', '小计', '总计', '其中', '说明', '备注']
const PRIMARY_DATA_COLUMN_KEYWORDS = [
  '合同编号',
  '项目/设备名称',
  '项目_设备名称',
  '客户名称',
  '终端客户',
  '项目状态',
  '验收单日期',
  '实际开工时间',
  '实施结束时间',
  '合同金额',
]

export async function parseExcelFile(
  file: File,
  options: ParseWorkbookOptions = {},
): Promise<ParsedWorksheet> {
  const buffer = await file.arrayBuffer()
  const workbook = XLSX.read(buffer, { type: 'array' })
  const sheetName = options.sheetName ?? workbook.SheetNames[0]
  const sheet = workbook.Sheets[sheetName]

  if (!sheet) {
    throw new Error('未找到可解析的工作表')
  }

  const matrix = fillMergedCells(
    buildDisplayMatrix(sheet),
    sheet['!merges'] ?? [],
  )

  const headerRowIndex = options.headerRowIndex ?? detectHeaderRowIndex(matrix)

  if (headerRowIndex < 0) {
    throw new Error('未识别到表头，请检查 Excel 内容')
  }

  const columns = buildColumns(matrix[headerRowIndex] ?? [])

  if (!columns.length) {
    throw new Error('表头为空，无法生成表格列')
  }

  const skippedRows: SkippedRow[] = []
  const rows: ExcelDataRow[] = []
  const summaryKeywords = options.summaryKeywords ?? DEFAULT_SUMMARY_KEYWORDS

  matrix.slice(headerRowIndex + 1).forEach((rawRow, offset) => {
    const sourceRowIndex = headerRowIndex + offset + 2
    const rowValues = columns.map((column) => normalizeCellValue(rawRow[column.index]))

    if (isEmptyRow(rowValues)) {
      return
    }

    const skipReason = getSkipReason(rowValues, columns, summaryKeywords)

    if (skipReason) {
      skippedRows.push({
        rowIndex: sourceRowIndex,
        reason: skipReason,
        preview: rowValues.filter(Boolean).join(' | ').slice(0, 120),
      })
      return
    }

    rows.push(createDataRow(rowValues, columns, sourceRowIndex))
  })

  return {
    sheetName,
    headerRowIndex: headerRowIndex + 1,
    columns,
    rows,
    skippedRows,
    rawRowCount: matrix.length,
  }
}

function buildDisplayMatrix(sheet: XLSX.WorkSheet): ExcelCellValue[][] {
  const sheetRef = sheet['!ref']

  if (!sheetRef) {
    return []
  }

  const range = XLSX.utils.decode_range(sheetRef)
  const matrix: ExcelCellValue[][] = []

  for (let rowIndex = range.s.r; rowIndex <= range.e.r; rowIndex += 1) {
    const row: ExcelCellValue[] = []

    for (let columnIndex = range.s.c; columnIndex <= range.e.c; columnIndex += 1) {
      const address = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex })
      row.push(getCellDisplayValue(sheet[address]))
    }

    matrix.push(row)
  }

  return matrix
}

function fillMergedCells(
  matrix: (string | number | boolean | null)[][],
  merges: XLSX.Range[],
): (string | number | boolean | null)[][] {
  const nextMatrix = matrix.map((row) => [...row])

  merges.forEach(({ s, e }) => {
    const sourceValue = nextMatrix[s.r]?.[s.c] ?? ''

    for (let rowIndex = s.r; rowIndex <= e.r; rowIndex += 1) {
      nextMatrix[rowIndex] ??= []

      for (let columnIndex = s.c; columnIndex <= e.c; columnIndex += 1) {
        nextMatrix[rowIndex][columnIndex] = sourceValue
      }
    }
  })

  return nextMatrix
}

function detectHeaderRowIndex(matrix: (string | number | boolean | null)[][]): number {
  return matrix.findIndex((row) => countNonEmpty(row.map(normalizeCellValue)) >= 2)
}

function buildColumns(headerRow: (string | number | boolean | null)[]): ExcelColumn[] {
  const usedKeys = new Set<string>()

  return headerRow
    .map((cell, index) => {
      const title = String(normalizeCellValue(cell) ?? '').trim()

      if (!title) {
        return null
      }

      const baseKey = slugifyKey(title) || `column_${index + 1}`
      let nextKey = baseKey
      let suffix = 1

      while (usedKeys.has(nextKey)) {
        suffix += 1
        nextKey = `${baseKey}_${suffix}`
      }

      usedKeys.add(nextKey)

      return {
        key: nextKey,
        title,
        index,
      } satisfies ExcelColumn
    })
    .filter((column): column is ExcelColumn => column !== null)
}

function createDataRow(
  rowValues: ExcelCellValue[],
  columns: ExcelColumn[],
  sourceRowIndex: number,
): ExcelDataRow {
  return columns.reduce<ExcelDataRow>(
    (result, column, index) => {
      result[column.key] = rowValues[index]
      return result
    },
    {
      __rowId: crypto.randomUUID(),
      __sourceRow: sourceRowIndex,
    },
  )
}

function getSkipReason(
  rowValues: ExcelCellValue[],
  columns: ExcelColumn[],
  summaryKeywords: string[],
): string | null {
  const normalizedValues = rowValues.map((value) => String(value ?? '').trim())
  const nonEmptyValues = normalizedValues.filter(Boolean)

  if (!hasPrimaryDataContent(rowValues, columns) && nonEmptyValues.length < 4) {
    return '合并单元格延伸出的空白行'
  }

  if (nonEmptyValues.length === 1) {
    return '独占整行的说明文本'
  }

  const joinedText = nonEmptyValues.join(' ')

  if (
    summaryKeywords.some((keyword) => joinedText.includes(keyword)) &&
    nonEmptyValues.length <= Math.max(4, Math.ceil(rowValues.length * 0.35))
  ) {
    return '合计或说明行'
  }

  return null
}

function countNonEmpty(values: ExcelCellValue[]): number {
  return values.filter((value) => {
    if (typeof value === 'number') {
      return true
    }

    return String(value ?? '').trim() !== ''
  }).length
}

function isEmptyRow(values: ExcelCellValue[]): boolean {
  return countNonEmpty(values) === 0
}

function hasPrimaryDataContent(rowValues: ExcelCellValue[], columns: ExcelColumn[]): boolean {
  return columns.some((column, index) => {
    const value = String(rowValues[index] ?? '').trim()

    if (!value) {
      return false
    }

    return PRIMARY_DATA_COLUMN_KEYWORDS.some((keyword) => {
      return column.title.includes(keyword) || column.key.includes(slugifyKey(keyword))
    })
  })
}

function normalizeCellValue(value: unknown): ExcelCellValue {
  if (value === undefined || value === null) {
    return ''
  }

  if (typeof value === 'string') {
    return value.trim()
  }

  if (typeof value === 'number' || typeof value === 'boolean') {
    return value
  }

  return String(value).trim()
}

function getCellDisplayValue(cell?: XLSX.CellObject): ExcelCellValue {
  if (!cell) {
    return ''
  }

  if (cell.t === 'n' && typeof cell.v === 'number' && looksLikeDateCell(cell)) {
    return formatExcelDate(cell.v)
  }

  if (typeof cell.w === 'string' && cell.w.trim()) {
    return normalizeDateString(cell.w.trim())
  }

  if (cell.t === 'n' && typeof cell.v === 'number') {
    return cell.v
  }

  if (cell.t === 'b') {
    return Boolean(cell.v)
  }

  return normalizeCellValue(cell.v)
}

function looksLikeDateCell(cell: XLSX.CellObject): boolean {
  const formatText = String(cell.z ?? '').toLowerCase()
  const renderedText = String(cell.w ?? '').trim()

  return (
    /[ymdhs]/.test(formatText) ||
    /^\d{1,4}[/-]\d{1,2}[/-]\d{1,4}$/.test(renderedText) ||
    /^\d{1,2}:\d{2}(:\d{2})?$/.test(renderedText)
  )
}

function formatExcelDate(serial: number): string {
  const parsed = XLSX.SSF.parse_date_code(serial)

  if (!parsed) {
    return String(serial)
  }

  const year = String(parsed.y).padStart(4, '0')
  const month = String(parsed.m)
  const day = String(parsed.d)

  if (parsed.H || parsed.M || parsed.S) {
    const hour = String(parsed.H).padStart(2, '0')
    const minute = String(parsed.M).padStart(2, '0')
    const second = String(parsed.S).padStart(2, '0')
    return `${year}/${month}/${day} ${hour}:${minute}:${second}`
  }

  return `${year}/${month}/${day}`
}

function normalizeDateString(value: string): string {
  const normalized = value.replace(/\./g, '/').replace(/-/g, '/')
  const parts = normalized.split('/')

  if (parts.length !== 3 || parts.some((part) => !/^\d+$/.test(part))) {
    return value
  }

  const [first, second, third] = parts.map((part) => Number(part))

  if (String(parts[0]).length === 4) {
    return `${first}/${second}/${third}`
  }

  if (String(parts[2]).length === 2) {
    const year = third >= 70 ? 1900 + third : 2000 + third
    return `${year}/${first}/${second}`
  }

  if (String(parts[2]).length === 4) {
    return `${third}/${first}/${second}`
  }

  return value
}

function slugifyKey(value: string): string {
  return value
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '_')
    .replace(/[^\w\u4e00-\u9fa5]/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_|_$/g, '')
}
