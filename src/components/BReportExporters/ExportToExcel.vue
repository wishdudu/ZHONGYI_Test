<script setup lang="ts">
import { ref, watch } from 'vue'
import { ElMessage } from 'element-plus'
import * as ExcelJS from 'exceljs'
import { saveAs } from 'file-saver'

const props = defineProps({
  originalFile: File,
  responseText: String
})

watch(
  () => props.responseText,
  () => {
    console.log("responseText:", props.responseText)
  }
)

const exporting = ref(false)

const exportExcel = async () => {
  if (!props.originalFile || !props.responseText) {
    ElMessage.warning('缺少原始文件或模型结果')
    return
  }

  exporting.value = true

  try {
    const workbook = new ExcelJS.Workbook()
    const arrayBuffer = await props.originalFile.arrayBuffer()
    await workbook.xlsx.load(arrayBuffer)

    const tableData = parseResponseText(props.responseText)

    // ✅ 只从第二张表开始处理
    const sheetsToProcess = workbook.worksheets.slice(1)

    for (const sheet of sheetsToProcess) {
      const titleCell = sheet.getCell('A1').value?.toString().trim() || ''
      if (!titleCell) continue

      const matched = tableData.find(t =>
        titleCell.includes(t.tableName) || t.tableName.includes(titleCell)
      )
      if (!matched) continue

      const { itemMap } = matched

      let startRowIndex = -1
      let startColIndex = -1
      let testItemColumnIndex = -1

      sheet.eachRow((row, rowNumber) => {
        row.eachCell((cell, colNumber) => {
          const value = cell.value?.toString()?.trim()
          if (value === '单位' && startRowIndex === -1) {
            let nextRow = rowNumber + 1
            while (nextRow <= sheet.rowCount) {
              const nextCell = sheet.getRow(nextRow).getCell(colNumber)
              const nextValue = nextCell.value?.toString()?.trim()
              if (nextValue !== '单位') {
                break
              }
              nextRow++
            }
            startRowIndex = nextRow
            startColIndex = colNumber + 1
          }

          if (rowNumber === 2 && value?.includes('检测项目名称')) {
            testItemColumnIndex = colNumber
          }
        })
      })

      if (startRowIndex === -1 || testItemColumnIndex === -1) continue

      const headerRow = sheet.getRow(startRowIndex)
      const writeCol = headerRow.cellCount + 1
      console.log(startRowIndex, startColIndex)

      for (let i = startRowIndex; i <= sheet.rowCount; i++) {
        const row = sheet.getRow(i)
        const itemName = row.getCell(testItemColumnIndex - 1).value?.toString().trim()
        if (itemName && itemMap.has(itemName)) {
          const resultValue = itemMap.get(itemName)
          row.getCell(writeCol).value = resultValue

          const startCol = startColIndex
          const endCol = writeCol - 1

          if (resultValue && typeof resultValue === 'string' && resultValue.trim() !== '') {
            for (let col = startCol; col <= endCol; col++) {
              const cell = row.getCell(col)
              const text = cell.value?.toString() || ''

              if (text === '-') continue

              const [isBelowDetection, num] = parseCellValue(text)
              if (num === null) continue

              const actualValue = isBelowDetection ? 0 : num
              if (!isValidValue(actualValue, resultValue)) {
                cell.style = {} // 清除原样式
                cell.fill = {
                  type: 'pattern',
                  pattern: 'solid',
                  fgColor: { argb: 'FFFF0000' }
                }
              }
            }
          }
        }
      }
    }

    const buffer = await workbook.xlsx.writeBuffer()
    saveAs(new Blob([buffer]), '校验结果.xlsx')
    ElMessage.success('Excel导出成功')
  } catch (err) {
    console.error('导出失败', err)
    ElMessage.error('Excel导出失败')
  } finally {
    exporting.value = false
  }
}


// 解析单元格值，返回 [是否低于检出限, 数值]
function parseCellValue(text: string): [boolean, number | null] {
  if (text.includes('＜') || text.includes('<')) {
    const num = extractNumber(text)
    return [true, num ?? 0]
  }
  const num = extractNumber(text)
  return [false, num]
}

// 提取数值（支持科学计数法）
function extractNumber(text: string): number | null {
  const match = text.match(/[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?/)
  return match ? parseFloat(match[0]) : null
}

// 判断实际值是否满足AI返回的范围或阈值
function isValidValue(value: number, expected: string): boolean {
  expected = expected.trim()
  if (!expected || expected.toLowerCase() === "null") return true

  const ranges = expected.split(/\s*or\s*|\s*或\s*|\s*;\s*/i).map(r => r.trim())

  for (const range of ranges) {
    if (isSingleValue(range)) {
      if (value > parseFloat(range)) return false
      continue
    }

    if (isRangeValue(range)) {
      const [min, max] = getRangeBounds(range)
      if (value < min || value > max) return false
      continue
    }

    if (isInequalityRange(range)) {
      if (!checkInequalityRange(value, range)) return false
    }

    if (isSimpleInequality(range)) {
      if (!checkSimpleInequality(value, range)) return false
      continue
    }
  }

  return true
}

// 判断是否为单一数值
function isSingleValue(val: string): boolean {
  return /^[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$/.test(val)
}

// 判断是否为范围格式
function isRangeValue(val: string): boolean {
  return /^(-?\d*\.?\d+)\s*[-~至~]\s*(-?\d*\.?\d+)$/.test(val)
}

// 提取范围的上下界数值
function getRangeBounds(val: string): [number, number] {
  const match = val.match(/(-?\d*\.?\d+)\s*[-~至~]\s*(-?\d*\.?\d+)/)
  return match ? [parseFloat(match[1]), parseFloat(match[2])] : [0, 0]
}

// 判断是否为不等式格式（例如 5≤x<10）
function isInequalityRange(val: string): boolean {
  // 更宽松的匹配模式，支持变量和多种符号
  return /([\d.]+)\s*([<>≤≥]=?)\s*[^0-9]+\s*([<>≤≥]=?)\s*([\d.]+)/.test(val);
}

// 解析不等式范围并验证
function checkInequalityRange(value: number, val: string): boolean {
  // 标准化符号
  const normalized = val
    .replace(/≤/g, '<=')
    .replace(/≥/g, '>=')
    .replace(/＜/g, '<')
    .replace(/＞/g, '>')
    .replace(/pH值|pH|[\s:：]/g, ''); // 移除变量名和空格

  // 提取关键部分
  const match = normalized.match(/([\d.]+)([<>]=?)([<>]=?)([\d.]+)/);
  if (!match) return true;

  const [_, leftVal, leftOp, rightOp, rightVal] = match;
  const leftNum = parseFloat(leftVal);
  const rightNum = parseFloat(rightVal);

  // 确定边界类型
  const isLowerBound = leftOp.includes('>');
  const isUpperBound = rightOp.includes('<');

  // 验证范围
  if (isLowerBound && isUpperBound) {
    return (
      (leftOp === '>' ? value > leftNum : value >= leftNum) &&
      (rightOp === '<' ? value < rightNum : value <= rightNum)
    );
  }

  // 处理其他组合情况（如 a<x<b 和 a>x>b）
  let lowerBound = isLowerBound ? leftNum : rightNum;
  let upperBound = isUpperBound ? rightNum : leftNum;
  let lowerOp = isLowerBound ? leftOp : rightOp;
  let upperOp = isUpperBound ? rightOp : leftOp;

  return (
    (lowerOp === '>' ? value > lowerBound : value >= lowerBound) &&
    (upperOp === '<' ? value < upperBound : value <= upperBound)
  );
}


// 判断是否为简单不等式（如 <5、≤10、>0.5、≥1e-3）
function isSimpleInequality(val: string): boolean {
  return /^[<>≤≥]=?\s*\d+\.?\d*(e[-+]?\d+)?$/i.test(val)
}

// 校验简单不等式（如 value < 5）
function checkSimpleInequality(value: number, val: string): boolean {
  const match = val.match(/^([<>≤≥]=?)\s*(\d+\.?\d*(?:e[-+]?\d+)?)/i)
  if (!match) return true
  const [_, operator, numberStr] = match
  const threshold = parseFloat(numberStr)

  switch (operator) {
    case '<':
    case '＜': return value < threshold
    case '<=':
    case '≤': return value <= threshold
    case '>':
    case '＞': return value > threshold
    case '>=':
    case '≥': return value >= threshold
    default: return true
  }
}

// 解析大模型返回文本为：表名 + 项目-值映射
function parseResponseText(text: string): { tableName: string, itemMap: Map<string, string> }[] {
  const result: { tableName: string, itemMap: Map<string, string> }[] = []

  const pattern = /([\u4e00-\u9fa5A-Za-z0-9（）()]+表)[:：]([\s\S]*?)(?=(?:[\u4e00-\u9fa5A-Za-z0-9（）()]+表)[:：]|$)/g
  let match

  while ((match = pattern.exec(text)) !== null) {
    const tableName = match[1].trim()
    const content = match[2].trim()
    const itemMap = new Map<string, string>()

    const pairs = content.split(/；|;/).map(s => s.trim()).filter(Boolean)
    for (const pair of pairs) {
      const [key, value] = pair.split(/[:：]/).map(s => s.trim())
      if (key && value !== undefined) {
        itemMap.set(key, value)
      }
    }

    result.push({ tableName, itemMap })
  }

  return result
}
</script>

<template>
  <el-button type="success" :loading="exporting" @click="exportExcel">
    导出为 Excel（B表）
  </el-button>
</template>
