<template>
    <div class="login-wrapper">
      <div class="login-box">
        <h2 class="login-title">登录</h2>
        <el-form @submit.prevent="handleLogin" label-position="top">
          <el-form-item label="用户名">
            <el-input v-model="username" placeholder="请输入用户名" />
          </el-form-item>
          <el-form-item label="密码">
            <el-input v-model="password" type="password" placeholder="请输入密码" />
          </el-form-item>
          <el-form-item>
            <el-button type="primary" @click="handleLogin" class="login-button" block>
              登录
            </el-button>
          </el-form-item>
        </el-form>
      </div>
    </div>
  </template>
  
  <script setup lang="ts">
  import { ref } from 'vue'
  import { useRouter } from 'vue-router'
  import { ElMessage } from 'element-plus'

  
  const username = ref('')
  const password = ref('')
  const router = useRouter()
  
  const handleLogin = () => {
    if (username.value === 'admin' && password.value === '123456') {
      localStorage.setItem('loggedIn', 'true')
      router.push('/')
    } else {
      ElMessage.error('账号或密码错误')
    }
  }
  </script>
  
  <style scoped>
  .login-wrapper {
    display: flex;
    justify-content: center;
    align-items: center;
    height: 100vh;
    background: #f8f8f8;
  }
  
  .login-box {
    width: 360px;
    padding: 30px 24px;
    background-color: white;
    box-shadow: 0 6px 18px rgba(0, 0, 0, 0.1);
    border-radius: 12px;
  }
  
  .login-title {
    text-align: center;
    margin-bottom: 24px;
    font-weight: bold;
    color: #333;
  }
  
  .login-button {
    width: 100%;
  }
  </style>
  