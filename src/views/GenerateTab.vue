<template>
    <div class="generate-container">
      <h2>数据待导出</h2>
  
      <el-upload
        :auto-upload="false"
        :file-list="fileList"
        :on-change="handleFileChange"
        :on-remove="handleFileRemove"
        :limit="1"
        accept=".xlsx,.xls"
      >
        <el-button type="success">上传数据汇总 Excel</el-button>
      </el-upload>
  
      <!-- <el-button
        type="primary"
        style="margin-top: 20px; width: 150px;"
        @click="exportWord"
        :disabled="!excelFile"
      >
        导出为 Word（C表）
      </el-button> -->
      <CReportExporter :originalFile="excelFile || undefined"/>
    </div>
  </template>
  
  <script setup>
  import { ref } from 'vue'
  import CReportExporter from '../components/CReportExporters/CReportExporter.vue'
  
  const fileList = ref([])
  const excelFile = ref(null)
  
  const handleFileChange = (file) => {
    excelFile.value = file.raw
    fileList.value = [file]
  }
  
  const handleFileRemove = () => {
    excelFile.value = null
    fileList.value = []
  }
  
  </script>
  
  <style scoped>
  .generate-container {
    display: flex;
    flex-direction: column;
    gap: 20px;
  }
  </style>
  