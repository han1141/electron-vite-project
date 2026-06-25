import { computed, reactive, ref, watch } from 'vue'
import { hasuraImportProfile } from '../config/hasuraImportProfile'
import type { ExcelDataRow, HasuraImportConfig, ParsedWorksheet } from '../types/excel'
import { fetchDictItems, importRowsToHasura, testHasuraConnection } from '../utils/hasura'

const STORAGE_KEY = 'excel-import-hasura-config'
const PROJECT_STATUS_DICT_CODE = 'PROJECT_STATUS'
const PROJECT_STATUS_SOURCE_KEY = '项目状态'

export function useExcelImport() {
  const parsedWorksheet = ref<ParsedWorksheet | null>(null)
  const editableRows = ref<ExcelDataRow[]>([])
  const importing = ref(false)
  const testingConnection = ref(false)
  const projectStatusMap = ref(new Map<string, string>())

  const hasuraConfig = reactive<HasuraImportConfig>(loadSavedConfig())

  const canSubmit = computed(() => {
    return Boolean(
        parsedWorksheet.value &&
        editableRows.value.length &&
        hasuraConfig.endpoint &&
        hasuraConfig.adminSecret &&
        hasuraImportProfile.tableName,
    )
  })

  async function setParsedWorksheet(worksheet: ParsedWorksheet) {
    const statusMap = await ensureProjectStatusMap()
    const mappedRows = mapWorksheetProjectStatusRows(worksheet.rows, statusMap)

    parsedWorksheet.value = {
      ...worksheet,
      rows: mappedRows,
    }
    editableRows.value = mappedRows.map((row) => ({ ...row }))
  }

  function resetImportedData() {
    parsedWorksheet.value = null
    editableRows.value = []
  }

  function updateCell(rowId: string, field: string, value: string | number | boolean | null) {
    const target = editableRows.value.find((row) => row.__rowId === rowId)

    if (target) {
      target[field] = value
    }
  }

  function removeRow(rowId: string) {
    editableRows.value = editableRows.value.filter((row) => row.__rowId !== rowId)
  }

  function updateHasuraConfig(patch: Partial<HasuraImportConfig>) {
    Object.assign(hasuraConfig, patch)
  }

  function saveHasuraConfig() {
    localStorage.setItem(
      STORAGE_KEY,
      JSON.stringify({
        endpoint: hasuraConfig.endpoint,
        adminSecret: hasuraConfig.adminSecret,
      }),
    )
  }

  async function submitToHasura() {
    if (!canSubmit.value) {
      throw new Error('请先导入 Excel，并补全 Hasura 地址与请求头密码')
    }

    importing.value = true

    try {
      return await importRowsToHasura(hasuraConfig, editableRows.value)
    } finally {
      importing.value = false
    }
  }

  async function verifyHasuraConnection() {
    testingConnection.value = true

    try {
      return await testHasuraConnection(hasuraConfig)
    } finally {
      testingConnection.value = false
    }
  }

  watch(
    () => [hasuraConfig.endpoint, hasuraConfig.adminSecret] as const,
    async ([endpoint, adminSecret]) => {
      projectStatusMap.value = new Map()

      if (!endpoint || !adminSecret) {
        return
      }

      try {
        await ensureProjectStatusMap()
      } catch (error) {
        console.error('[hasura] 启动时预取字典失败', error)
      }
    },
    { immediate: true },
  )

  return {
    parsedWorksheet,
    editableRows,
    hasuraConfig,
    importing,
    testingConnection,
    canSubmit,
    setParsedWorksheet,
    resetImportedData,
    updateCell,
    removeRow,
    updateHasuraConfig,
    saveHasuraConfig,
    verifyHasuraConnection,
    submitToHasura,
  }

  async function ensureProjectStatusMap(): Promise<Map<string, string>> {
    if (projectStatusMap.value.size > 0) {
      return projectStatusMap.value
    }

    if (!hasuraConfig.endpoint || !hasuraConfig.adminSecret) {
      return projectStatusMap.value
    }

    const dictItems = await fetchDictItems(hasuraConfig, PROJECT_STATUS_DICT_CODE)
    projectStatusMap.value = new Map(
      dictItems.map((item) => [item.name.trim(), item.id]),
    )

    return projectStatusMap.value
  }
}

function mapWorksheetProjectStatusRows(
  rows: ExcelDataRow[],
  projectStatusMap: Map<string, string>,
): ExcelDataRow[] {
  return rows.map((row) => {
    const sourceValue = row[PROJECT_STATUS_SOURCE_KEY]

    if (typeof sourceValue !== 'string') {
      return { ...row }
    }

    const mappedId = projectStatusMap.get(sourceValue.trim())

    if (!mappedId) {
      return { ...row }
    }

    return {
      ...row,
      [PROJECT_STATUS_SOURCE_KEY]: mappedId,
    }
  })
}

function loadSavedConfig(): HasuraImportConfig {
  try {
    const rawValue = localStorage.getItem(STORAGE_KEY)

    if (!rawValue) {
      return createEmptyConfig()
    }

    return {
      ...createEmptyConfig(),
      ...JSON.parse(rawValue),
    }
  } catch {
    return createEmptyConfig()
  }
}

function createEmptyConfig(): HasuraImportConfig {
  return {
    endpoint: '',
    adminSecret: '',
  }
}
