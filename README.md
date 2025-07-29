# ZHONGYI_Test - 文档处理平台

基于Vue 3 + TypeScript + Vite构建的文档处理前端应用，对接后端ragflow工作流，提供Excel/Word文件处理、导出和OCR功能。

## 功能特性

- Excel文件上传、处理和导出
- Word文档生成和导出  
- PDF文档解析
- OCR文字识别功能
- 基于Element Plus的UI界面

## 演示视频

[![Bilibili演示](https://i0.hdslb.com/bfs/archive/8267ecd45b3cdbe543a18d91b92a62e76c5f1f4d.jpg)](https://www.bilibili.com/video/BV1Aj8DzzEHd/)

## 项目结构

```
src/
├── components/            # 可复用组件
│   ├── BReportExporters/  # Excel导出相关
│   ├── CReportExporters/  # Word导出相关
│   └── UploadHandlers/    # 文件上传处理
├── utils/                 # 工具函数
│   └── CReportExporters/  # 文档处理核心逻辑
├── views/                 # 页面视图
│   ├── DetectTab.vue      # 检测功能页
│   └── GenerateTab.vue    # 生成功能页
```

## 技术栈

- Vue 3 + TypeScript
- Vite构建工具
- Pinia状态管理
- Element Plus UI框架
- Vitest单元测试
- Cypress端到端测试

## 核心依赖

- `docx`: Word文档生成
- `exceljs`: Excel文件处理  
- `xlsx`: Excel文件解析
- `mammoth`: Word文档解析
- `pdfjs-dist`: PDF文档解析
- `tesseract.js`: OCR文字识别

## 开发环境配置

### 推荐IDE

[VSCode](https://code.visualstudio.com/) + [Volar](https://marketplace.visualstudio.com/items?itemName=Vue.volar) (禁用Vetur)

### TypeScript支持

TypeScript默认无法处理`.vue`导入的类型信息，我们使用`vue-tsc`替代`tsc`进行类型检查。

## 项目设置

```sh
npm install
```

### 开发模式

```sh
npm run dev
```

### 生产构建

```sh
npm run build
```

### 运行单元测试

```sh
npm run test:unit
```

### 运行端到端测试

开发模式测试(快速):
```sh
npm run test:e2e:dev
```

生产构建测试(推荐CI使用):
```sh
npm run build
npm run test:e2e
```

### 代码检查

```sh
npm run lint
```

### 代码格式化

```sh
npm run format
```

## 自定义配置

参考[Vite配置文档](https://vite.dev/config/)
