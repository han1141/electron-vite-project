<script setup lang="ts">
import { UploadFilled } from '@element-plus/icons-vue'
import { ElMessage } from 'element-plus'
import EditableDataTable from './components/EditableDataTable.vue'
import ExcelUploader from './components/ExcelUploader.vue'
import FieldMappingPanel from './components/FieldMappingPanel.vue'
import { useExcelImport } from './composables/useExcelImport'
import type { ParsedWorksheet } from './types/excel'

const {
  parsedWorksheet,
  editableRows,
  hasuraConfig,
  importing,
  testingConnection,
  canSubmit,
  setParsedWorksheet,
  updateCell,
  removeRow,
  updateHasuraConfig,
  saveHasuraConfig,
  verifyHasuraConnection,
  submitToHasura,
} = useExcelImport()

function handleParsed(worksheet: ParsedWorksheet) {
  setParsedWorksheet(worksheet)
}

async function handleSubmit() {
  try {
    const rowCount = await submitToHasura()
    ElMessage.success(`已打印 ${rowCount} 条待提交数据到控制台，未执行真实提交`)
  } catch (error) {
    const message = error instanceof Error ? error.message : '导入失败'
    ElMessage.error(message)
  }
}

function handleSaveHasuraConfig() {
  saveHasuraConfig()
  ElMessage.success('Hasura 配置已保存到本地缓存')
}

async function handleTestHasuraConfig() {
  try {
    const matchedRows = await verifyHasuraConnection()
    ElMessage.success(`Hasura 连接成功，report_project 测试查询返回 ${matchedRows} 条记录`)
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Hasura 连接测试失败'
    ElMessage.error(message)
  }
}
</script>

<template>
  <main class="app-shell">
    <section class="hero">
      <div>
        <div class="eyebrow">Vue 3 + Element Plus + Hasura</div>
        <h1>Excel 导入与逐行校对</h1>
        <p>
          处理合并单元格、过滤说明/合计行、逐行编辑后生成 Hasura 提交数据。
        </p>
      </div>

      <el-button
        type="primary"
        size="large"
        :icon="UploadFilled"
        :loading="importing"
        :disabled="!canSubmit"
        @click="handleSubmit"
      >
        打印提交数据
      </el-button>
    </section>

    <section class="top-panels">
      <FieldMappingPanel
        :config="hasuraConfig"
        :testing="testingConnection"
        @update-config="updateHasuraConfig"
        @save="handleSaveHasuraConfig"
        @test="handleTestHasuraConfig"
      />

      <section class="summary-panel">
        <div class="panel-header">
          <div>
            <h2>解析摘要</h2>
            <p>这部分用来核对自动过滤结果。</p>
          </div>
        </div>

        <div class="summary-list">
          <div class="summary-item">
            <span>工作表</span>
            <strong>{{ parsedWorksheet?.sheetName || '未导入' }}</strong>
          </div>
          <div class="summary-item">
            <span>表头行</span>
            <strong>{{ parsedWorksheet ? `第 ${parsedWorksheet.headerRowIndex} 行` : '-' }}</strong>
          </div>
          <div class="summary-item">
            <span>可编辑数据</span>
            <strong>{{ editableRows.length }} 行</strong>
          </div>
          <div class="summary-item">
            <span>目标表</span>
            <strong>report_project</strong>
          </div>
          <div class="summary-item">
            <span>已过滤</span>
            <strong>{{ parsedWorksheet?.skippedRows.length ?? 0 }} 行</strong>
          </div>
        </div>

        <el-scrollbar
          v-if="parsedWorksheet?.skippedRows.length"
          max-height="240px"
          class="skipped-list"
        >
          <div
            v-for="row in parsedWorksheet?.skippedRows"
            :key="`${row.rowIndex}-${row.preview}`"
            class="skipped-item"
          >
            <div class="skipped-meta">第 {{ row.rowIndex }} 行 · {{ row.reason }}</div>
            <div class="skipped-preview">{{ row.preview }}</div>
          </div>
        </el-scrollbar>
        <div v-else class="empty-summary">导入 Excel 后会在这里显示解析结果和过滤明细。</div>
      </section>
    </section>

    <div class="main-column">
      <ExcelUploader @parsed="handleParsed" />

      <EditableDataTable
        v-if="parsedWorksheet"
        :columns="parsedWorksheet.columns"
        :rows="editableRows"
        @update-cell="updateCell"
        @remove-row="removeRow"
      />
    </div>
  </main>
</template>
