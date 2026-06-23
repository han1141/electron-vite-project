import type { ExcelDataRow, HasuraImportConfig } from '../types/excel'
import { hasuraImportProfile } from '../config/hasuraImportProfile'

const NUMERIC_FIELDS = new Set([
  'project_code',
  'contract_amount_lowercase',
  'confirmed_revenue_amount',
  'implementation_budget_amount',
  'implementation_budget',
  'actual_expenditure_cost',
  'extortion_money',
  'tariff_rate',
  'amount_returned',
  'remain_balance',
  'procurement_amount',
])

const DATE_FIELDS = new Set([
  'acceptance_date',
  'actual_start_date',
  'implementation_end_date',
  'maintenance_end_date',
  'date_performance',
  'warranty_start_date',
  'warranty_end_date',
])

const TIMESTAMP_FIELDS = new Set([
  'project_code_application_time',
])

const NULL_LIKE_VALUES = new Set(['', '-', '—', '--', '待定', '无', '暂无', '/'])
const REQUIRED_FIELDS = ['name', 'project_code']

interface HasuraResponse<TData> {
  data?: TData
  errors?: Array<{ message: string }>
}

interface ExistingProjectRecord {
  id: string
  project_code: number | null
}

export async function testHasuraConnection(config: HasuraImportConfig): Promise<number> {
  if (!config.endpoint || !config.adminSecret) {
    throw new Error('请先填写 Hasura 地址与请求头密码')
  }

  const result = await requestHasura<{ report_project?: Array<{ id: string }> }>(
    config,
    `
      query TestReportProjectConnection {
        report_project(limit: 1) {
          id
        }
      }
    `,
  )

  return result.data?.report_project?.length ?? 0
}

export async function importRowsToHasura(
  config: HasuraImportConfig,
  rows: ExcelDataRow[],
): Promise<number> {
  if (!hasuraImportProfile.mutationName || !hasuraImportProfile.objectTypeName) {
    throw new Error('请先在 src/config/hasuraImportProfile.ts 中设置固定 Hasura mutation 配置')
  }

  const payload = buildMutationPayload(rows, hasuraImportProfile.fieldMappings)
  validatePreparedObjects(rows, payload.variables.objects)
  const mergePreview = await buildMergePreview(config, payload.variables.objects)

  console.group('Hasura Import Payload Preview')
  console.log('endpoint:', config.endpoint)
  console.log('mutation:', hasuraImportProfile.mutationName)
  console.log('payload:', payload)
  console.log('mergePreview:', mergePreview)
  console.groupEnd()

  return payload.variables.objects.length
}

function buildMutationPayload(
  rows: ExcelDataRow[],
  fieldMappings: Record<string, string | string[]>,
) {
  const objects = rows.map((row) => mapRowToHasuraObject(row, fieldMappings))

  return {
    query: `
      mutation ImportRows($objects: [${hasuraImportProfile.objectTypeName}!]!) {
        ${hasuraImportProfile.mutationName}(objects: $objects) {
          affected_rows
        }
      }
    `,
    variables: {
      objects,
    },
  }
}

function mapRowToHasuraObject(
  row: ExcelDataRow,
  fieldMappings: Record<string, string | string[]>,
): Record<string, string | number | boolean | null> {
  const sourceEntries = Object.entries(row).filter(([key]) => !key.startsWith('__'))

  return sourceEntries.reduce<Record<string, string | number | boolean | null>>(
    (result, [sourceKey, value]) => {
      const mappedTargets = fieldMappings[sourceKey]

      if (!mappedTargets) {
        return result
      }

      const targetKeys = Array.isArray(mappedTargets) ? mappedTargets : [mappedTargets]

      targetKeys.forEach((targetKey) => {
        result[targetKey] = sanitizeValue(targetKey, value)
      })

      return result
    },
    {},
  )
}

function sanitizeValue(targetKey: string, value: unknown): string | number | boolean | null {
  if (value === undefined || value === null) {
    return null
  }

  if (typeof value === 'boolean') {
    return value
  }

  const stringValue = String(value).trim()

  if (NULL_LIKE_VALUES.has(stringValue)) {
    return null
  }

  if (NUMERIC_FIELDS.has(targetKey)) {
    return normalizeNumericValue(stringValue)
  }

  if (DATE_FIELDS.has(targetKey)) {
    return normalizeDateValue(stringValue)
  }

  if (TIMESTAMP_FIELDS.has(targetKey)) {
    return normalizeTimestampValue(stringValue)
  }

  return stringValue
}

function normalizeNumericValue(value: string): number | null {
  const cleaned = value.replace(/,/g, '').replace(/%/g, '')

  if (!cleaned) {
    return null
  }

  const parsed = Number(cleaned)
  return Number.isFinite(parsed) ? parsed : null
}

function normalizeDateValue(value: string): string | null {
  const normalized = value.replace(/\./g, '/').replace(/-/g, '/').trim()
  const parts = normalized.split('/')

  if (parts.length !== 3 || parts.some((part) => !/^\d+$/.test(part))) {
    return null
  }

  let year = Number(parts[0])
  let month = Number(parts[1])
  let day = Number(parts[2])

  if (parts[0].length !== 4 && parts[2].length === 4) {
    year = Number(parts[2])
    month = Number(parts[0])
    day = Number(parts[1])
  }

  if (!year || month < 1 || month > 12 || day < 1 || day > 31) {
    return null
  }

  return `${String(year).padStart(4, '0')}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`
}

function normalizeTimestampValue(value: string): string | null {
  const dateValue = normalizeDateValue(value)
  return dateValue ? `${dateValue}T00:00:00` : null
}

function validatePreparedObjects(
  sourceRows: ExcelDataRow[],
  objects: Array<Record<string, string | number | boolean | null>>,
) {
  const invalidRows = objects
    .map((object, index) => {
      const missingFields = REQUIRED_FIELDS.filter((field) => {
        const value = object[field]
        return value === null || value === undefined || String(value).trim() === ''
      })

      if (!missingFields.length) {
        return null
      }

      return {
        sourceRow: sourceRows[index]?.__sourceRow ?? index + 1,
        missingFields,
      }
    })
    .filter((row): row is { sourceRow: number; missingFields: string[] } => row !== null)

  if (!invalidRows.length) {
    return
  }

  const preview = invalidRows
    .slice(0, 5)
    .map((row) => `第 ${row.sourceRow} 行缺少: ${row.missingFields.join(', ')}`)
    .join('; ')

  throw new Error(`提交前校验失败。${preview}`)
}

async function buildMergePreview(
  config: HasuraImportConfig,
  objects: Array<Record<string, string | number | boolean | null>>,
) {
  const projectCodes = Array.from(
    new Set(
      objects
        .map((object) => object.project_code)
        .filter((value): value is number => typeof value === 'number' && Number.isFinite(value)),
    ),
  )

  if (!projectCodes.length) {
    return {
      inserts: objects,
      updates: [],
      matchedRecords: [],
    }
  }

  const existingRecords = await findExistingProjectsByCodes(config, projectCodes)
  const existingMap = new Map(existingRecords.map((record) => [record.project_code, record]))

  const inserts: Array<Record<string, string | number | boolean | null>> = []
  const updatesById = new Map<string, Record<string, string | number | boolean | null>>()

  objects.forEach((object) => {
    const projectCode = object.project_code
    const matched = typeof projectCode === 'number' ? existingMap.get(projectCode) : undefined

    if (!matched) {
      inserts.push(object)
      return
    }

    const previous = updatesById.get(matched.id) ?? {}
    updatesById.set(matched.id, {
      ...previous,
      ...object,
    })
  })

  return {
    inserts,
    updates: Array.from(updatesById.entries()).map(([id, _set]) => ({
      id,
      _set,
    })),
    matchedRecords: existingRecords,
  }
}

async function findExistingProjectsByCodes(
  config: HasuraImportConfig,
  projectCodes: number[],
): Promise<ExistingProjectRecord[]> {
  const result = await requestHasura<{ report_project: ExistingProjectRecord[] }>(
    config,
    `
      query FindExistingProjects($projectCodes: [numeric!]!) {
        report_project(where: { project_code: { _in: $projectCodes } }) {
          id
          project_code
        }
      }
    `,
    {
      projectCodes,
    },
  )

  return result.data?.report_project ?? []
}

async function requestHasura<TData>(
  config: HasuraImportConfig,
  query: string,
  variables?: Record<string, unknown>,
): Promise<HasuraResponse<TData>> {
  const response = await fetch(config.endpoint, {
    method: 'POST',
    headers: {
      'content-type': 'application/json',
      'x-hasura-admin-secret': config.adminSecret,
    },
    body: JSON.stringify({
      query,
      variables,
    }),
  })

  if (!response.ok) {
    throw new Error(`Hasura 请求失败: ${response.status} ${response.statusText}`)
  }

  const result = (await response.json()) as HasuraResponse<TData>

  if (result.errors?.length) {
    throw new Error(result.errors.map((item) => item.message).join('; '))
  }

  return result
}
