<script setup lang="ts">
import { ref } from 'vue'
import * as ExcelJS from 'exceljs'
import { ElMessage } from 'element-plus'

const props = defineProps({
  onDataParsed: {
    type: Function,
    required: true
  },
  onFileSelected: {
    type: Function,
    required: false // 可选
  }
})

const loading = ref(false)
const fileList = ref<any[]>([])

const handleFileChange = async (uploadFile: any) => {
  const file = uploadFile.raw
  if (!file) return

  loading.value = true
  fileList.value = [uploadFile]

  try {
    const arrayBuffer = await file.arrayBuffer()
    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.load(arrayBuffer)

    const allGroupedItems: string[] = []

    // 仅从第二张表开始处理（跳过索引为0的第一张）
    const sheets = workbook.worksheets.slice(1)

    for (const worksheet of sheets) {
      const sheetTitle = getFirstCellOfFirstRow(worksheet)
      const testItems = parseTestItemsFromSheet(worksheet)

      if (testItems.length > 0) {
        const group = `${sheetTitle}：${testItems.join('，')}`
        allGroupedItems.push(group)
      }
    }

    const finalResult = allGroupedItems.join('，')
    props.onDataParsed(finalResult)

    // ✅ 传出原始上传文件
    props.onFileSelected?.(file)
    ElMessage.success('Excel文件解析成功')
  } catch (error) {
    console.error('Error parsing Excel:', error)
    ElMessage.error('Excel文件解析失败')
    fileList.value = []
  } finally {
    loading.value = false
  }
}

// 获取工作表第一行第一个单元格内容作为前缀
function getFirstCellOfFirstRow(worksheet: ExcelJS.Worksheet): string {
  const firstRow = worksheet.getRow(1)
  const firstCell = firstRow.getCell(1)
  return firstCell?.value?.toString().trim() || '未知表格'
}

// 提取“检测项目名称”列下的非标题内容，排除“检测项目名称”本身
function parseTestItemsFromSheet(worksheet: ExcelJS.Worksheet): string[] {
  const testItems: string[] = []
  let testItemColumnIndex = -1

  worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    if (rowNumber === 2) {
      row.eachCell((cell, colNumber) => {
        const val = cell.value?.toString().trim()
        if (val && val.includes('检测项目名称')) {
          testItemColumnIndex = colNumber
        }
      })
    }
  })

  if (testItemColumnIndex > 0) {
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      const cell = row.getCell(testItemColumnIndex - 1)
      const value = cell.value?.toString().trim()
      if (
        value &&
        rowNumber > 2 &&
        !value.includes('检测项目名称') // 排除误识别
      ) {
        testItems.push(value)
      }
    })
  }

  return testItems
}
</script>


<template>
  <div v-loading="loading">
    <el-upload
      :auto-upload="false"
      :file-list="fileList"
      :on-change="handleFileChange"
      :limit="1"
      accept=".xlsx,.xls"
    >
      <el-button type="success">上传数据汇总Excel</el-button>
    </el-upload>
  </div>
</template>
