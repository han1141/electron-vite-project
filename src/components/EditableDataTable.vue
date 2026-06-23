<script setup lang="ts">
import { Delete } from '@element-plus/icons-vue'
import { computed, ref } from 'vue'
import type { ExcelColumn, ExcelDataRow } from '../types/excel'

const props = defineProps<{
  columns: ExcelColumn[]
  rows: ExcelDataRow[]
}>()

const emit = defineEmits<{
  updateCell: [rowId: string, field: string, value: string]
  removeRow: [rowId: string]
}>()

const editingCell = ref<string | null>(null)

const tableColumns = computed(() => props.columns.slice(0, 24))

function getCellKey(rowId: string, field: string) {
  return `${rowId}:${field}`
}

function startEdit(rowId: string, field: string) {
  editingCell.value = getCellKey(rowId, field)
}

function stopEdit() {
  editingCell.value = null
}
</script>

<template>
  <section class="table-panel">
    <div class="panel-header">
      <div>
        <h2>数据预览</h2>
        <p>每一行都可单独编辑，删除后不会参与导入。</p>
      </div>
      <div class="row-count">共 {{ props.rows.length }} 行</div>
    </div>

    <el-alert
      v-if="props.columns.length > tableColumns.length"
      type="info"
      show-icon
      :closable="false"
      class="table-alert"
      title="列数较多，当前先展示前 24 列；如需全量列，可继续扩展横向分页或列筛选。"
    />

    <div class="table-wrap">
      <el-table
        :data="props.rows"
        row-key="__rowId"
        border
        height="560"
        width="100%"
        table-layout="fixed"
      >
      <el-table-column label="#" width="80" fixed="left">
        <template #default="{ row }">
          {{ row.__sourceRow }}
        </template>
      </el-table-column>

      <el-table-column
        v-for="column in tableColumns"
        :key="column.key"
        :prop="column.key"
        :label="column.title"
        width="140"
        show-overflow-tooltip
      >
        <template #default="{ row }">
          <div
            v-if="editingCell !== getCellKey(row.__rowId, column.key)"
            class="cell-display"
            @dblclick="startEdit(row.__rowId, column.key)"
          >
            {{ String(row[column.key] ?? '') || '-' }}
          </div>
          <el-input
            v-else
            :model-value="String(row[column.key] ?? '')"
            autofocus
            @update:model-value="emit('updateCell', row.__rowId, column.key, $event)"
            @blur="stopEdit"
            @keyup.enter="stopEdit"
          />
        </template>
      </el-table-column>

      <el-table-column label="删除行" width="96" fixed="right" align="center">
        <template #default="{ row }">
          <el-button
            :icon="Delete"
            circle
            type="danger"
            plain
            @click="emit('removeRow', row.__rowId)"
          />
        </template>
      </el-table-column>
      </el-table>
    </div>
  </section>
</template>
