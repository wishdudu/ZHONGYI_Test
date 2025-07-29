<!-- WordUploadHandler.vue -->
<script setup lang="ts">
import { ref } from 'vue'
import * as mammoth from 'mammoth'
import { ElMessage } from 'element-plus'

const props = defineProps({
  onWordParsed: {
    type: Function,
    required: true
  }
})

const loadingWord = ref(false)
const fileList = ref<any[]>([])
const file = ref<File | null>(null)

const handleFileChange = async (uploadFile: any) => {
  file.value = uploadFile.raw
  fileList.value = [uploadFile]
  loadingWord.value = true

  try {
    const arrayBuffer = await file.value!.arrayBuffer()
    const result = await mammoth.extractRawText({ arrayBuffer })
    const text = result.value

    const startIndex = text.split('\n').findIndex(line =>
      line.includes('废水检测结果') || line.includes('Test Conclusion')
    )

    if (startIndex === -1) {
      ElMessage.error('上传失败，未找到关键字段')
      props.onWordParsed('❌ 未提取到检测结果段')
    } else {
      const lines = text.split('\n').slice(startIndex)
      const clean = lines.map(l => l.trim()).filter(Boolean).join('，')
      props.onWordParsed(clean)
      ElMessage.success('Word 文件解析成功')
    }
  } catch (e) {
    ElMessage.error('Word 解析失败')
    props.onWordParsed('❌ 解析失败')
  } finally {
    loadingWord.value = false
  }
}
</script>

<template>
  <div v-loading="loadingWord">
    <el-upload
    
      :auto-upload="false"
      :file-list="fileList"
      :on-change="handleFileChange"
      :limit="1"
      accept=".docx"
    >
      <el-button type="primary">上传待检测 Word</el-button>
    </el-upload>
  </div>
</template>
