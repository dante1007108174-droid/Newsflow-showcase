const express = require('express');
const cors = require('cors');
const dotenv = require('dotenv');
const path = require('path');
const rateLimit = require('express-rate-limit');

// 加载环境变量（固定读取 backend/.env，避免工作目录差异）
dotenv.config({ path: path.join(__dirname, '.env') });

const app = express();
const PORT = process.env.PORT || 3000;
const IS_PRODUCTION = process.env.NODE_ENV === 'production';

app.set('trust proxy', 1);

function parseAllowedOrigins(rawOrigins) {
  if (!rawOrigins || typeof rawOrigins !== 'string') return [];
  return rawOrigins
    .split(',')
    .map((origin) => origin.trim())
    .filter(Boolean);
}

const allowedOrigins = parseAllowedOrigins(process.env.ALLOWED_ORIGINS);

if (IS_PRODUCTION && allowedOrigins.length === 0) {
  throw new Error('ALLOWED_ORIGINS must be configured in production');
}

const corsOptions = {
  origin(origin, callback) {
    // Allow non-browser clients (curl/health checks) that do not send Origin.
    if (!origin) return callback(null, true);

    // In local dev, allow all if ALLOWED_ORIGINS is not configured.
    if (!IS_PRODUCTION && allowedOrigins.length === 0) {
      return callback(null, true);
    }

    if (allowedOrigins.includes(origin)) {
      return callback(null, true);
    }

    return callback(new Error('Not allowed by CORS'));
  },
  credentials: true,
};

const chatLimiter = rateLimit({
  windowMs: 60 * 1000,
  max: 30,
  standardHeaders: true,
  legacyHeaders: false,
  message: { error: '请求过于频繁，请稍后重试' },
});

const workflowLimiter = rateLimit({
  windowMs: 60 * 1000,
  max: 20,
  standardHeaders: true,
  legacyHeaders: false,
  message: { error: '请求过于频繁，请稍后重试' },
});

const subscriptionLimiter = rateLimit({
  windowMs: 60 * 1000,
  max: 40,
  standardHeaders: true,
  legacyHeaders: false,
  message: { error: '请求过于频繁，请稍后重试' },
});

const feedbackLimiter = rateLimit({
  windowMs: 60 * 1000,
  max: 30,
  standardHeaders: true,
  legacyHeaders: false,
  message: { error: '请求过于频繁，请稍后重试' },
});

// 中间件
app.use(express.json());
app.use(cors(corsOptions));
app.use('/api/chat', chatLimiter);
app.use('/api/workflow', workflowLimiter);
app.use('/api/subscription', subscriptionLimiter);
app.use('/api/feedback', feedbackLimiter);

// 托管前端静态文件
app.use(express.static(path.join(__dirname, '../frontend')));

// 请求日志中间件
app.use((req, res, next) => {
  console.log(`[${new Date().toISOString()}] ${req.method} ${req.path}`);
  next();
});

// 路由
const workflowRoutes = require('./routes/workflow');
app.use('/api/workflow', workflowRoutes);

const chatRoutes = require('./routes/chat');
app.use('/api/chat', chatRoutes);

const subscriptionRoutes = require('./routes/subscription');
app.use('/api/subscription', subscriptionRoutes);

const feedbackRoutes = require('./routes/feedback');
app.use('/api/feedback', feedbackRoutes);

const trackRoutes = require('./routes/track');
app.use('/api/track', trackRoutes);

// 健康检查端点
app.get('/health', (req, res) => {
  res.json({ 
    status: 'ok', 
    timestamp: new Date().toISOString(),
    service: 'daily-ai-news-backend'
  });
});

// 根路径返回前端 index.html
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, '../frontend/index.html'));
});

// 404处理
app.use((req, res) => {
  res.status(404).json({ error: '接口不存在' });
});

// 全局错误处理
app.use((err, req, res, next) => {
  console.error('❌ Error:', err);
  res.status(500).json({ 
    error: '服务器内部错误',
    message: process.env.NODE_ENV === 'development' ? err.message : undefined
  });
});

let shuttingDown = false;

// 启动服务器
const server = app.listen(PORT, () => {
  console.log('');
  console.log('🚀 ================================');
  console.log(`🚀 Server running on http://localhost:${PORT}`);
  console.log(`🌐 CORS Origins: ${allowedOrigins.length ? allowedOrigins.join(', ') : 'not configured'}`);
  console.log('🚀 ================================');
  console.log('');
});

function gracefulShutdown(signal) {
  if (shuttingDown) return;
  shuttingDown = true;
  console.log(`\n🛑 Received ${signal}, shutting down gracefully...`);

  server.close((err) => {
    if (err) {
      console.error('❌ Graceful shutdown failed:', err);
      process.exit(1);
    }
    console.log('✅ Server closed gracefully');
    process.exit(0);
  });

  setTimeout(() => {
    console.error('❌ Force shutdown after timeout');
    process.exit(1);
  }, 10000).unref();
}

process.on('SIGTERM', () => gracefulShutdown('SIGTERM'));
process.on('SIGINT', () => gracefulShutdown('SIGINT'));
