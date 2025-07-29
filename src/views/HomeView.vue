<template>
  <div class="app-container">
    <!-- 顶部导航栏 -->
    <header class="header">
      <div class="logo">
        <img src="/logo.png" alt="logo" />
        <span>环评检测生成智能体</span>
      </div>
    </header>

    <!-- 功能 Tabs -->
    <el-tabs v-model="activeTab" type="border-card" class="tabs" stretch>
      <el-tab-pane label="检测" name="detect">
        <DetectTab />
      </el-tab-pane>
      <el-tab-pane label="生成" name="generate">
        <GenerateTab />
      </el-tab-pane>
    </el-tabs>
  </div>
</template>

<script setup>
import { ref } from 'vue'
import DetectTab from './DetectTab.vue'
import GenerateTab from './GenerateTab.vue'

const activeTab = ref('detect')
</script>

<style scoped>
.app-container {
  min-height: 100vh;
  background: #f5f7fa;
}

.header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 16px 40px;
  background: #ffffff;
  box-shadow: 0 2px 8px rgba(0,0,0,0.05);
}

.logo {
  display: flex;
  align-items: center;
  font-size: 20px;
  font-weight: bold;
}

.logo img {
  width: 32px;
  height: 32px;
  margin-right: 12px;
}

.tabs {
  max-width: 800px;
  margin: 0px auto;
  background: #fff;
  padding: 32px 0px 32px 32px;
  border-radius: 12px;
  box-shadow: 0 4px 12px rgba(0,0,0,0.05);
}
</style>
