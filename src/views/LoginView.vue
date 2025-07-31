<template>
  <div class="login-wrapper">
    <div class="login-box">
      <h2 class="login-title">登录</h2>
      <el-form 
        @submit.prevent="handleLogin" 
        label-position="top"
        :model="{ username, password }"
        :rules="rules"
      >
        <el-form-item label="用户名" prop="username">
          <el-input v-model="username" placeholder="请输入用户名" />
        </el-form-item>
        <el-form-item label="密码" prop="password">
          <el-input v-model="password" type="password" placeholder="请输入密码" show-password />
        </el-form-item>
        <el-form-item>
          <el-checkbox v-model="rememberMe">记住我</el-checkbox>
        </el-form-item>
        <el-form-item>
          <el-button 
            type="primary" 
            @click="handleLogin" 
            class="login-button" 
            block
            :loading="loading"
          >
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

const router = useRouter()
const loading = ref(false)

const username = ref('')
const password = ref('')
const rememberMe = ref(false)

const validateUsername = (rule: any, value: string, callback: any) => {
  if (!value) {
    callback(new Error('请输入用户名'))
  } else if (value.length < 3) {
    callback(new Error('用户名至少3个字符'))
  } else {
    callback()
  }
}

const validatePassword = (rule: any, value: string, callback: any) => {
  if (!value) {
    callback(new Error('请输入密码'))
  } else if (value.length < 6) {
    callback(new Error('密码至少6个字符'))
  } else {
    callback()
  }
}

const rules = {
  username: [{ validator: validateUsername, trigger: 'blur' }],
  password: [{ validator: validatePassword, trigger: 'blur' }]
}

const handleLogin = async () => {
  try {
    loading.value = true
    if (username.value === import.meta.env.VITE_ADMIN_USERNAME && 
        password.value === import.meta.env.VITE_ADMIN_PASSWORD) {
      if (rememberMe.value) {
        localStorage.setItem('loggedIn', 'true')
      } else {
        sessionStorage.setItem('loggedIn', 'true')
      }
      router.push('/')
    } else {
      throw new Error('账号或密码错误')
    }
  } catch (error) {
    ElMessage.error(error instanceof Error ? error.message : '登录失败')
  } finally {
    loading.value = false
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
