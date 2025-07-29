<!-- src/components/CReportExporter/CReportExporter.vue -->
<template>
  <!-- 导出按钮 -->
  <el-button type="primary" @click="handleExportClick" :loading="isExporting">
    {{ isExporting ? '导出中...' : '导出为Word（C表）' }}
  </el-button>
</template>

<script setup>
import { ref } from 'vue'
import { ElMessage } from 'element-plus'
import { processExcelFile } from '../../utils/CReportExporters/excelProcessor'
import { generateDocxDocument } from '../../utils/CReportExporters/docxGenerator'
import { saveAs } from 'file-saver'

// 获取上传文件
const props = defineProps({
  originalFile: {
    type: File,
    required: true,
    validator: (file) => file instanceof File || file === null
  }
})

// 按钮状态
const isExporting = ref(false)

// 点击导出事件
const handleExportClick = async () => {
  if (!props.originalFile) {
    ElMessage.error('请先上传文件')
    return
  }

  isExporting.value = true

  try {
    // --- Step 1: 处理Excel数据 ---
    const reportData = await processExcelFile(props.originalFile)

    // --- Step 2: 生成C表文档 ---
    const docxBlob = await generateDocxDocument(reportData)

    // --- Step 3: 导出Word文件 ---
    saveAs(docxBlob, 'C表.docx')
    ElMessage.success('已成功导出Word文件')

  } catch (error) {
    console.error('导出失败:', error)
    ElMessage.error(`导出失败: ${error.message || error}`)
  } finally {
    isExporting.value = false
  }
}
</script>
