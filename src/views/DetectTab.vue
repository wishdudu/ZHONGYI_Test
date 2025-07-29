<!-- <script setup lang="ts">
import TheWelcome from '../components/TheWelcome.vue'
</script>

<template>
  <main>
    <TheWelcome />
  </main>
</template> -->


<template>
    <div class="container">
      <!-- <h2 style="text-align: center;">中一检测智能体</h2> -->
  
      <el-form label-position="top" class="upload-form">
        <h3 style="margin-top: 30px;">1.上传检测结果:</h3>
        <div style="display: flex; align-items: center; gap: 10px; margin-top: 16px;">
          <ExcelUploadHandler
            :on-data-parsed="handleExcelData"
            :on-file-selected="handleExcelFile"
          />
          <WordUploadHandler :on-word-parsed="handleWordData" />
        </div>
        <br />
        <p
          style="font-size: 12px; color: #888; margin-top: -20px; margin-bottom: 40px;"
        >
          ⚠️ Excel上传：用于AI校验或生成C表的B表文件<br />
          ⚠️ Word上传：后缀名为docx的检测结果文件（暂不可用）
        </p>
        <h3 style="margin-top: 30px;">2.输入评价标准:</h3>
        <el-form-item label="">
          <el-input
            type="textarea"
            autosize
            placeholder="手动输入评价标准"
            v-model="manualDescription"
            :disabled="disableTextInput"
          />
        </el-form-item>
      </el-form>
  
      <el-form-item label="">
        <div v-loading="loadingPDF">
          <el-upload
            :auto-upload="false"
            :file-list="fileList1"
            :on-change="handleFile1Change"
            :on-remove="handleFile1Remove"
            :limit="1"
            accept="application/pdf"
            :disabled="disablePDFUpload"
          >
            <el-button type="danger" >（上传PDF）</el-button>
          </el-upload>
        </div>
      </el-form-item>
  
      <p
        style="font-size: 12px; color: #888; margin-top: -8px; margin-bottom: 12px;"
      >
        ⚠️ “输入评价标准”与“上传PDF”二选一，尽量避免用edge浏览器
      </p>
      <br />
      <el-button
        type="primary"
        @click="submit"
        :disabled="loadingPDF || loadingExcel"
      >
        发送
      </el-button>
  
      <h3 style="margin-top: 30px;">模型回复：</h3>
  
      <!-- 模型回复内容 -->
      <div class="response-box" v-html="response"></div>
  
      <div style="display: flex; align-items: center; gap: 10px; margin-top: 16px;">
      <ExportToExcel :originalFile="file0 || undefined" :responseText="response" />
      <!-- <ExportToWord :originalFile="file0 || undefined"/> -->
      </div>
  
      <!-- 页面底部的说明文字 -->
      <div style="margin-top: 20px; font-size: 12px; color: #888;">
        注：通过上传检测结果和技术附件，点击发送，模型将自动提取信息并分析，
        输出为表格形式的结果并且可以选择导出为excel文件。
      </div>
    </div>
  </template>
  
  
  <script setup lang="ts">
import { ref, computed, onMounted } from 'vue'
import * as pdfjsLib from 'pdfjs-dist'
import { ElMessage } from 'element-plus'
import ExcelUploadHandler from '../components/UploadHandlers/ExcelUploadHandler.vue'
import WordUploadHandler from '../components/UploadHandlers/WordUploadHandler.vue'
import ExportToExcel from '../components/BReportExporters/ExportToExcel.vue'
import { getBaiduAccessToken, recognizeImageByBaiduOCR } from '../api/baiduOcr'
import { sendToModel } from '../api/backend'
  
  const disablePDFUpload = computed(() => manualDescription.value.trim() !== '')
  const disableTextInput = computed(() => fileList1.value.length > 0)
  
  const timer = ref(0) // 计时器变量，单位为秒
  const timerInterval = ref<any | null>(null) // 计时器ID，用来清除定时器
  
  // 启动计时器
  function startTimer() {
    timer.value = 0
    timerInterval.value = setInterval(() => {
      timer.value += 1 // 每秒增加1s
      updateResponseText() // 更新显示的计时
    }, 1000)
  }
  
  // 更新显示的计时
  function updateResponseText() {
    response.value = `⏳ 正在发送内容给模型... (${timer.value}秒)`
  }
  
  // 停止计时器
  function stopTimer() {
    if (timerInterval.value !== null) {
      clearInterval(timerInterval.value)
      timerInterval.value = null
    }
  }
  
  // 设置 PDF.js 的 worker 路径
  pdfjsLib.GlobalWorkerOptions.workerSrc =
    'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js'
  
  // 引用和文件管理
  const fileList0 = ref<any[]>([]) // Word 文件
  const fileList1 = ref<any[]>([]) // PDF 文件
  const file0 = ref<File | null>(null)
  const file1 = ref<File | null>(null)
  
  const manualDescription = ref('')
  const response = ref('')
  const sessionId = ref<string | null>(null)
  const uploadStatus = ref<{ success: boolean, message: string } | null>(null)
  
  // 保存文件上传的处理结果
  const excelResult = ref<string | null>(null)
  const wordResult = ref<string | null>(null)
  const pdfResult = ref<string | null>(null)
  
  const loadingPDF = ref(false)
  const loadingWord = ref(false)
  const loadingExcel = ref(false)
  
  const handleExcelFile = (file: File) => {
    file0.value = file
  }
  
  // 处理excel文件上传
  const handleExcelData = (data: string) => {
    excelResult.value = data
    console.log('Excel检测项目解析结果', data)
  }
  
  // 处理 Word 文件上传
  const handleWordData = (data: string) => {
    // wordResult.value = data
    console.log('📄 Word解析结果:', data)
  }
  
  // 处理 PDF 文件上传
  const handleFile1Change = async (uploadFile: any) => {
    file1.value = uploadFile.raw
    fileList1.value = [uploadFile]
  
    loadingPDF.value = true
    uploadStatus.value = null
    
    if (file1.value) { // 检查 file1.value 是否为 null
      try {
        const pdfText = await extractTextFromScannedPDF(file1.value) // 确保文件已上传并且是有效的
        const result = extractBetweenKeywordsFromTable(pdfText)
  
        if (result && !result.startsWith('❌')) {
          pdfResult.value = result
          ElMessage.success( '上传PDF成功')
          console.log('📄 PDF解析结果:', result)
        } else {
          file1.value = null
          fileList1.value = []
          pdfResult.value = null
          ElMessage.error('上传失败，未找到关键字' )
        }
      } catch (error) {
        ElMessage.error( '上传失败，解析出错' )
        pdfResult.value = null
    } finally {
      loadingPDF.value = false
    }
    } else {
      ElMessage.error( '上传失败，文件为空' )
      pdfResult.value = null
    }
  }
  const handleFile1Remove = () => {
    file1.value = null
    fileList1.value = []
  }
  
  
  // 在页面加载时自动请求模型的打招呼内容
  onMounted(() => {
    sendGreetingRequest()
  })
  
  // 发送请求获取模型的打招呼内容
  async function sendGreetingRequest() {
    try {
      response.value = '⏳ 正在请求模型打招呼内容...'
  
      // 发送一个简单的请求，只是用来获取模型的打招呼消息
      const greetingMessage = 'Hello, could you please greet me?'  // 可以根据模型的设计调整
      const { reply } = await sendToModel(greetingMessage)
      response.value = reply // 显示模型返回的打招呼内容
    } catch (e: any) {
      response.value = `❌ 出错：${e.message}`
    }
  }
  
  
  
  // // 提取 Word 文本
  // async function extractTextFromDocx(file: File): Promise<string> {
  //   const arrayBuffer = await file.arrayBuffer()
  //   const result = await mammoth.extractRawText({ arrayBuffer })
  //   return result.value
  // }
  
  // // 提取 Word 中“检测结果”部分
  // function extractTestResultTable(text: string): string {
  //   const lines = text.split('\n')
  //   const startIndex = lines.findIndex(line =>
  //     line.includes('废水检测结果') || line.includes('Test Conclusion')
  //   )
  //   if (startIndex === -1) return '❌ 未提取到检测结果段'
  
  //   const relevantLines = lines.slice(startIndex).map(line => line.trim()).filter(line => line)
  //   return relevantLines.join('，') // 用顿号连接压缩文本
  // }
  
  
  
  
  
  // ✅ 修改：返回所有页的表格识别结果数组拼接
  async function extractTextFromScannedPDF(file: File): Promise<any[]> {
    const buffer = await file.arrayBuffer()
    const pdf = await pdfjsLib.getDocument({ data: buffer }).promise
    const accessToken = await getBaiduAccessToken()
  
    const allResults: any[] = []
  
    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i)
      const viewport = page.getViewport({ scale: 2.0 })
      const canvas = document.createElement('canvas')
      const context = canvas.getContext('2d')!
      canvas.width = viewport.width
      canvas.height = viewport.height
  
      await page.render({ canvasContext: context, viewport }).promise
      const base64Image = canvas.toDataURL('image/jpeg')
  
      const result = await recognizeImageByBaiduOCR(base64Image, accessToken)
      allResults.push(...result) // 拼接所有单元格
    }
  
    return allResults
  }
  
  
  // ✅ 新增：支持从结构化 OCR 表格中提取关键字段
  function extractBetweenKeywordsFromTable(tableData: any[]): string {
    console.log('📋 表格识别数据:', tableData.map(item => item.words))  // ✅ 打印所有字段
  
    const startKeywords = ['检测方案(可另附)', '检测方案（可另附）', '样品类别']
    const endKeywords = ['报告出具与发送要求', '报告形式']
  
    const indexStart = tableData.findIndex(item =>
      startKeywords.some(keyword => item.words?.includes(keyword))
    )
    if (indexStart === -1) return '❌ 未找到起始字段（如“样品类别”或“检测方案”）'
  
    const indexEnd = tableData.findIndex((item, i) =>
      i > indexStart &&
      endKeywords.some(keyword => item.words?.includes(keyword))
    )
  
    const indexStop = indexEnd !== -1 ? indexEnd : tableData.length
  
    const extracted = tableData
      .slice(indexStart, indexStop + 1)
      .map(item => item.words)
      .filter(Boolean)
  
    return extracted.join('，') // 用顿号压缩
  }
  
  
  // async function exportToExcel() {
  //   if (!response.value) {
  //     ElMessage.warning('当前无内容可导出')
  //     return
  //   }
  
  //   const workbook = new ExcelJS.Workbook()
  //   const worksheet = workbook.addWorksheet('检测结果')
  
  //   // 1. 从 response 中提取模型回复的 ASCII 表格部分
  //   const tableLines = response.value
  //     .split('\n')
  //     .filter(line => line.includes('|') && !/^[-\s|]+$/.test(line)) // 忽略仅含表格线的行
  
  //   tableLines.forEach((line, rowIndex) => {
  //     const rawCells = line.split('|').map(cell => cell.trim()).filter(c => c.length > 0)
  
  //     const row = worksheet.getRow(rowIndex + 1)
  //     rawCells.forEach((text, colIndex) => {
  //       const cell = row.getCell(colIndex + 1)
  
  //       // 处理标红：如果内容包含 span 红字，保留格式
  //       const match = text.match(/<span[^>]*style="[^"]*color:\s*red[^"]*"[^>]*>(.*?)<\/span>/i)
  //       if (match) {
  //         const plain = match[1]
  //         cell.value = {
  //           richText: [{ text: plain, font: { color: { argb: 'FFFF0000' } } }]
  //         }
  //       } else {
  //         // 正常写入
  //         cell.value = text.replace(/<[^>]+>/g, '') // 去除残留标签
  //       }
  //     })
  
  //     row.commit()
  //   })
  
  // 2. 提取非表格的文字说明行（即不含“|”的行）
  // const nonTableLines = response.value
  //   .split('\n')
  //   .filter(line => !line.includes('|') && line.trim() !== '') // 跳过空行和表格线
  
  // if (nonTableLines.length > 0) {
  //   const startRow = worksheet.lastRow ? worksheet.lastRow.number + 2 : tableLines.length + 2
  
  //   nonTableLines.forEach((text, index) => {
  //     const row = worksheet.getRow(startRow + index)
  //     row.getCell(1).value = text
  //     row.getCell(1).alignment = { wrapText: true, vertical:'top' } // 自动换行显示
  //     row.commit()
  //   })
  // }
  //   const buffer = await workbook.xlsx.writeBuffer()
  //   saveAs(new Blob([buffer]), '检测结果导出.xlsx')
  // }
  
  
  // 提交处理逻辑
  async function submit() {
    if (!file0.value) {
      response.value = '请上传 Excel文件'
      return
    }
    if(!file1.value && manualDescription.value.trim()==''){
      response.value = `水质结果汇总表：六价铬：1；碘化物：2；
  土壤和沉积物结果汇总表：甲苯：3；苯：4；`
      return
    }
    
    try {
      response.value = '⏳ 正在发送内容给模型...'
      startTimer()
  
  const standardText = manualDescription.value?.trim()
    ? manualDescription.value
    : pdfResult.value || '（未提供）';
  
  // 构建最终内容
  const finalMessage = `
  📄 评价标准如下：
  ${standardText}
  
  📄 检测项目如下：
  ${excelResult.value}
  `;
  
  
      console.log('📤 发送内容:\n', finalMessage)
      const { reply, sessionId: newSessionId } = await sendToModel(finalMessage, sessionId.value || undefined)
      response.value = reply
      if (newSessionId) {
        sessionId.value = newSessionId
        console.log('✅ 保存会话 ID:', sessionId.value)
      }
      stopTimer()
    } catch (e: any) {
      response.value = `❌ 出错：${e.message}`
      // 停止计时器并更新状态
      stopTimer()
    }
  }
  
  
  
  </script>
  
  <style scoped>
  .container {
    max-width: 600px;
    margin: 0 auto;
    padding: 0px 32px 32px 32px; /* 去掉顶部padding */
  }
  .upload-form {
    margin-bottom: 16px;
  }
  .question-input {
    margin-bottom: 12px;
  }
  
  .response-box {
    border: 1px solid #dcdfe6;
    padding: 12px;
    border-radius: 4px;
    background: #f9f9f9;
    white-space: pre-wrap;
  }
  </style>
