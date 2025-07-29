// 百度OCR API配置
export const BAIDU_OCR_CONFIG = {
  clientId: import.meta.env.VITE_BAIDU_OCR_CLIENT_ID,
  clientSecret: import.meta.env.VITE_BAIDU_OCR_CLIENT_SECRET,
  tokenUrl: import.meta.env.VITE_BAIDU_OCR_TOKEN_URL,
  recognizeUrl: import.meta.env.VITE_BAIDU_OCR_RECOGNIZE_URL
}

// 获取百度OCR Access Token
export async function getBaiduAccessToken(): Promise<string> {
  const res = await fetch(BAIDU_OCR_CONFIG.tokenUrl, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
    },
    body: new URLSearchParams({
      grant_type: 'client_credentials',
      client_id: BAIDU_OCR_CONFIG.clientId,
      client_secret: BAIDU_OCR_CONFIG.clientSecret,
    }),
  })

  const data = await res.json()
  return data.access_token
}

// 识别图像文字
export async function recognizeImageByBaiduOCR(
  base64Image: string, 
  accessToken: string
): Promise<string> {
  const res = await fetch(
    `${BAIDU_OCR_CONFIG.recognizeUrl}?access_token=${accessToken}`, 
    {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      body: new URLSearchParams({
        image: base64Image.replace(/^data:image\/\w+;base64,/, ''),
      }),
    }
  )
  const data = await res.json()
  return data.tables_result?.[0]?.body || []
}
