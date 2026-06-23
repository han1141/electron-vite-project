import { computed, reactive, ref } from 'vue'
import { hasuraImportProfile } from '../config/hasuraImportProfile'
import type { ExcelDataRow, HasuraImportConfig, ParsedWorksheet } from '../types/excel'
import { importRowsToHasura, testHasuraConnection } from '../utils/hasura'

const STORAGE_KEY = 'excel-import-hasura-config'

export function useExcelImport() {
  const parsedWorksheet = ref<ParsedWorksheet | null>(null)
  const editableRows = ref<ExcelDataRow[]>([])
  const importing = ref(false)
  const testingConnection = ref(false)

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

  function setParsedWorksheet(worksheet: ParsedWorksheet) {
    parsedWorksheet.value = worksheet
    editableRows.value = worksheet.rows.map((row) => ({ ...row }))
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
