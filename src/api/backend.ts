// 后端API配置
export const BACKEND_CONFIG = {
  baseUrl: import.meta.env.VITE_BACKEND_BASE_URL,
  agentId: import.meta.env.VITE_BACKEND_AGENT_ID,
  authToken: import.meta.env.VITE_BACKEND_AUTH_TOKEN
}

interface ModelResponse {
  reply: string
  sessionId?: string
}

// 请求大模型接口
export async function sendToModel(message: string, sessionId?: string): Promise<ModelResponse> {
  try {
    const requestBody: any = {
      model: 'model',
      messages: [{ role: 'user', content: message }],
      stream: false
    }

    if (sessionId) {
      requestBody.id = sessionId
    }

    const res = await fetch(
      `${BACKEND_CONFIG.baseUrl}/${BACKEND_CONFIG.agentId}/chat/completions`,
      {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          Authorization: BACKEND_CONFIG.authToken
        },
        body: JSON.stringify(requestBody)
      }
    )

    if (!res.ok) {
      const errorText = await res.text()
      throw new Error(`HTTP ${res.status}: ${errorText}`)
    }

    const data = await res.json()
    const reply = data.choices?.[0]?.message?.content || '无返回'
    return {
      reply,
      sessionId: data.id
    }
  } catch (error: any) {
    console.error('❌ 请求失败:', error)
    throw error
  }
}
