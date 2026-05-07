<script lang="ts" setup>
import { onBeforeUnmount, onMounted, ref } from 'vue'
import type { FUniver, Univer } from '@univerjs/presets'
import { createUniver, LocaleType, mergeLocales } from '@univerjs/presets'
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import UniverPresetSheetsCoreZhCN from '@univerjs/preset-sheets-core/locales/zh-CN'
import '@univerjs/preset-sheets-core/lib/index.css'

const container = ref<HTMLElement | null>(null)
const fileInput = ref<HTMLInputElement | null>(null)
const status = ref<string>('就绪 / Ready')
const loading = ref<boolean>(false)

let univerInstance: Univer | null = null
let univerAPIInstance: FUniver | null = null

function initUniver(initialData?: any) {
  if (univerInstance) {
    univerInstance.dispose()
    univerInstance = null
    univerAPIInstance = null
  }
  const { univer, univerAPI } = createUniver({
    locale: LocaleType.ZH_CN,
    locales: {
      [LocaleType.ZH_CN]: mergeLocales(UniverPresetSheetsCoreZhCN),
    },
    presets: [
      UniverSheetsCorePreset({ container: container.value as HTMLElement }),
    ],
  })
  univerInstance = univer
  univerAPIInstance = univerAPI
  if (initialData) {
    univerAPI.createWorkbook(initialData)
  } else {
    univerAPI.createWorkbook({})
  }
}

async function onImport(e: Event) {
  const target = e.target as HTMLInputElement
  const file = target.files?.[0]
  if (!file) return
  loading.value = true
  status.value = `正在上传 ${file.name}... / Uploading ${file.name}...`
  try {
    const form = new FormData()
    form.append('file', file)
    const resp = await fetch('/api/import', { method: 'POST', body: form })
    if (!resp.ok) {
      const errText = await resp.text()
      throw new Error(errText)
    }
    const data = await resp.json()
    initUniver(data)
    status.value = `已导入 ${file.name} / Imported ${file.name}`
  } catch (err: any) {
    status.value = `导入失败 / Import failed: ${err.message}`
    console.error(err)
  } finally {
    loading.value = false
    target.value = ''
  }
}

function triggerImport() {
  fileInput.value?.click()
}

async function onExport() {
  if (!univerAPIInstance) return
  const fWorkbook = univerAPIInstance.getActiveWorkbook()
  if (!fWorkbook) {
    status.value = '没有活动工作簿 / No active workbook'
    return
  }
  loading.value = true
  status.value = '正在导出... / Exporting...'
  try {
    const snapshot = fWorkbook.save()
    const resp = await fetch('/api/export?name=export', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(snapshot),
    })
    if (!resp.ok) throw new Error(await resp.text())
    const blob = await resp.blob()
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = 'export.xlsx'
    document.body.appendChild(a)
    a.click()
    document.body.removeChild(a)
    URL.revokeObjectURL(url)
    status.value = '已导出 export.xlsx / Exported export.xlsx'
  } catch (err: any) {
    status.value = `导出失败 / Export failed: ${err.message}`
    console.error(err)
  } finally {
    loading.value = false
  }
}

onMounted(() => initUniver())
onBeforeUnmount(() => {
  univerInstance?.dispose()
  univerInstance = null
  univerAPIInstance = null
})
</script>

<template>
  <div class="layout">
    <header class="toolbar">
      <h1>Univer xlsx Demo</h1>
      <div class="actions">
        <input ref="fileInput" type="file" accept=".xlsx" hidden @change="onImport" />
        <button :disabled="loading" @click="triggerImport">📤 导入 xlsx / Import</button>
        <button :disabled="loading" @click="onExport">📥 导出 xlsx / Export</button>
      </div>
      <span class="status">{{ status }}</span>
    </header>
    <div ref="container" class="univer-container" />
  </div>
</template>

<style>
html, body, #app { margin: 0; height: 100%; width: 100%; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; }
.layout { display: flex; flex-direction: column; height: 100vh; }
.toolbar { display: flex; align-items: center; gap: 16px; padding: 8px 16px; background: #fafafa; border-bottom: 1px solid #e5e5e5; }
.toolbar h1 { margin: 0; font-size: 16px; font-weight: 600; }
.actions { display: flex; gap: 8px; }
.actions button { padding: 6px 14px; font-size: 14px; background: #fff; border: 1px solid #d0d0d0; border-radius: 4px; cursor: pointer; }
.actions button:hover:not(:disabled) { background: #f0f0f0; }
.actions button:disabled { opacity: 0.5; cursor: not-allowed; }
.status { margin-left: auto; font-size: 13px; color: #666; }
.univer-container { flex: 1; overflow: hidden; }
</style>
