# NewsFlow.ai / AI 每日简报

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

NewsFlow.ai 是一个 AI 驱动的新闻聚合与推送助手。它从十几个中文媒体 RSS 源自动抓取内容，
经过两级过滤（代码规则层 + AI 语义层），压缩成每天 5-8 条高质量摘要，
通过邮件主动推送或对话即时查询送达用户。

核心解决的问题是**信息焦虑**：用户每天面对海量科技/财经资讯，
既怕错过重要信息（FOMO），又没有时间逐条筛选。
产品把“人找信息”变成“信息找人”，用 AI 完成筛选与压缩，
让用户每天花 2 分钟就能掌握关键动态。

## 功能特点 (Features)

- 📨 **每日日报推送**：每天自动输出 5-8 条高质量摘要，主动送达用户邮箱。
- 📧 **按主题即时推送**：支持按 AI / 财经 / 科技主题触发测试简报发送。
- 💬 **AI 对话浮窗**：支持流式输出、连续对话、移动端键盘适配。
- 🧠 **两级智能过滤**：代码规则层做基础筛选，AI 语义层做质量压缩。
- 🎨 **响应式 UI**：桌面端与移动端统一体验。

## 技术栈 (Tech Stack)

- **Frontend**: HTML5, Tailwind CSS, Vanilla JS
- **Backend**: Node.js, Express, Axios
- **Workflow / Data**: Coze Workflow, Supabase (via env)

## 项目结构 (Project Structure)

```text
newsflow-showcase/
├── backend/                # Node/Express 后端
│   ├── .env.example        # 脱敏后的环境变量示例
│   ├── routes/             # API 路由
│   ├── lib/                # 业务辅助库
│   └── server.js           # 服务入口
├── frontend/               # 前端静态资源
│   ├── index.html          # 主页面
│   ├── js/                 # 前端逻辑
│   └── assets/             # 静态资源
├── workflow/               # 流程导出（已脱敏）
├── docs/                   # 文档目录
├── tests/                  # 自动化测试脚本与样例
└── README.md
```

## 快速开始 (Quick Start)

1. 安装并启动后端

```bash
cd backend
npm install
npm start
```

2. 打开页面：`http://localhost:3000/`

## 配置说明 (Environment)

复制 `backend/.env.example` 为 `backend/.env` 并填写你自己的配置。

常用变量：

- `COZE_API_TOKEN`
- `COZE_WORKFLOW_ID`
- `COZE_BASE_URL`
- `COZE_BOT_ID`
- `SUPABASE_URL`
- `SUPABASE_SERVICE_ROLE_KEY`
- `ALLOWED_ORIGINS`

## 安全提示 (Security)

- 本仓库不包含真实 API Key / Token / 私密配置。
- `workflow/` 已做脱敏处理（如 `Authorization`, `Bearer`, token 字段）。
- `docs` 只保留 PRD 文档（文本文件），不公开测试数据二进制文件。
- 严禁提交 `backend/.env`。

## 许可证 (License)

[MIT](LICENSE)
