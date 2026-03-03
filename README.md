# AI 每日简报 / Daily AI News (Public Showcase)

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

精选全球 AI 资讯，每日早晨 8 点准时送达。告别信息过载，只关注真正重要的科技变革。  
Selected global AI news, delivered promptly at 8 AM every morning.

> 这是**公开展示镜像仓库**（Public Showcase Mirror），用于展示产品与工程实现。  
> 生产环境（Zeabur）仍使用原仓库，本仓库不参与生产部署。

## 功能特点 (Features)

- 📧 **简报测试发送**：支持按主题触发新闻简报测试流程。
- 💬 **AI 对话浮窗**：支持流式输出、连续对话、移动端键盘适配。
- 🎯 **主题覆盖**：支持 AI / 财经 / 科技 等主题。
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
├── docs/prd...             # PRD 文档（仅文本文件）
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

## 安全与公开策略 (Security for Public Repo)

- 本仓库不包含真实 API Key / Token / 私密配置。
- `workflow/` 已做脱敏处理（如 `Authorization`, `Bearer`, token 字段）。
- `docs` 只保留 PRD 文档（文本文件），不公开测试数据二进制文件。
- 严禁提交 `backend/.env`。

## 许可证 (License)

[MIT](LICENSE)
