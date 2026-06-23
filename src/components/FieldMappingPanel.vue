<script setup lang="ts">
import { Connection, FolderChecked } from '@element-plus/icons-vue'
import { hasuraImportProfile } from '../config/hasuraImportProfile'
import type { HasuraImportConfig } from '../types/excel'

const props = defineProps<{
  config: HasuraImportConfig
  testing: boolean
}>()

const emit = defineEmits<{
  updateConfig: [patch: Partial<HasuraImportConfig>]
  save: []
  test: []
}>()
</script>

<template>
  <section class="mapping-panel">
    <div class="panel-header">
      <div>
        <h2>Hasura 配置</h2>
        <p>这里只保留运行时配置。表名和特殊字段映射放在代码配置里维护。</p>
      </div>
      <div class="panel-actions">
        <el-button
          type="success"
          plain
          :icon="Connection"
          :loading="props.testing"
          @click="emit('test')"
        >
          测试
        </el-button>
        <el-button type="primary" plain :icon="FolderChecked" @click="emit('save')">
          保存
        </el-button>
      </div>
    </div>

    <el-form label-position="top" class="config-form">
      <div class="config-grid">
        <el-form-item label="GraphQL Endpoint">
          <el-input
            :model-value="props.config.endpoint"
            placeholder="https://your-hasura/v1/graphql"
            @update:model-value="emit('updateConfig', { endpoint: $event })"
          />
        </el-form-item>
        <el-form-item label="Admin Secret">
          <el-input
            :model-value="props.config.adminSecret"
            type="password"
            show-password
            placeholder="x-hasura-admin-secret"
            @update:model-value="emit('updateConfig', { adminSecret: $event })"
          />
        </el-form-item>
      </div>
    </el-form>

    <div class="profile-card">
      <div class="profile-row">
        <span>目标表</span>
        <strong>report_project</strong>
      </div>
      <div class="profile-row">
        <span>字段规则</span>
        <strong>{{ Object.keys(hasuraImportProfile.fieldMappings).length }} 条</strong>
      </div>
      <p class="profile-tip">
        需要固定字段映射时，修改
        <code>src/config/hasuraImportProfile.ts</code>
      </p>
    </div>
  </section>
</template>
