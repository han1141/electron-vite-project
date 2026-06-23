<script setup lang="ts">
import { UploadFilled } from '@element-plus/icons-vue'
import { ElMessage } from 'element-plus'
import type { UploadRequestOptions } from 'element-plus'
import { ref } from 'vue'
import type { ParsedWorksheet } from '../types/excel'
import { parseExcelFile } from '../utils/excelParser'

const emit = defineEmits<{
  parsed: [worksheet: ParsedWorksheet]
}>()

const loading = ref(false)
const currentFileName = ref('')

async function handleUploadRequest(options: UploadRequestOptions) {
  const { file } = options

  if (!(file instanceof File)) {
    ElMessage.error('未读取到有效文件')
    return
  }

  loading.value = true
  currentFileName.value = file.name

  try {
    const worksheet = await parseExcelFile(file)
    emit('parsed', worksheet)
    ElMessage.success(`已解析 ${worksheet.rows.length} 行数据`)
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Excel 解析失败'
    ElMessage.error(message)
  } finally {
    loading.value = false
  }
}
</script>

<template>
  <section class="uploader-panel">
    <div class="panel-header">
      <div>
        <h2>Excel 导入</h2>
        <p>支持合并单元格展开，自动跳过合计/说明行。</p>
      </div>
      <div v-if="currentFileName" class="file-name">{{ currentFileName }}</div>
    </div>

    <el-upload
      drag
      :show-file-list="false"
      accept=".xls,.xlsx"
      :http-request="handleUploadRequest"
      :disabled="loading"
      class="upload-box"
    >
      <el-icon class="upload-icon"><UploadFilled /></el-icon>
      <div class="el-upload__text">拖拽 Excel 到这里，或点击选择文件</div>
      <template #tip>
        <div class="el-upload__tip">建议一个工作表只放一类明细数据，首个有效表头会自动识别。</div>
      </template>
    </el-upload>
  </section>
</template>
