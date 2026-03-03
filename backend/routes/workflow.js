const express = require('express');
const nodemailer = require('nodemailer');
const axios = require('axios');
const { extractIdentity, trackServerEvent } = require('../lib/tracking');
const router = express.Router();

// 有效的主题列表
const VALID_TOPICS = ['AI', '财经', '科技'];

function maskEmail(email) {
  if (typeof email !== 'string') return '***';
  const normalized = email.trim();
  const atIndex = normalized.indexOf('@');
  if (atIndex <= 0) return '***';

  const localPart = normalized.slice(0, atIndex);
  const domain = normalized.slice(atIndex + 1);
  if (!domain) return `${localPart.charAt(0) || '*'}***`;

  return `${localPart.charAt(0) || '*'}***@${domain}`;
}

/**
 * POST /api/workflow/send-test
 * 发送测试邮件 - 触发Coze Workflow
 */
router.post('/send-test', async (req, res) => {
  const startedAt = Date.now();
  const identity = extractIdentity(req);

  try {
    const { email, keyword } = req.body;

    // 1. 验证必填字段
    if (!email || !keyword) {
      return res.status(400).json({ 
        error: '缺少必要参数',
        required: ['email', 'keyword'],
        received: { email: !!email, keyword: !!keyword }
      });
    }

    // 2. 验证邮箱格式
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(email)) {
      return res.status(400).json({ 
        error: '邮箱格式不正确',
        received: email
      });
    }

    // 3. 验证主题
    if (!VALID_TOPICS.includes(keyword)) {
      return res.status(400).json({ 
        error: '无效的主题',
        validTopics: VALID_TOPICS,
        received: keyword
      });
    }

    console.log(`📨 Sending test email to: ${maskEmail(email)}`);
    console.log(`📌 Topic: ${keyword}`);

    // 4. 调用Coze Workflow API
    const response = await axios.post(
      `${process.env.COZE_BASE_URL}/v1/workflow/run`,
      {
        workflow_id: process.env.COZE_WORKFLOW_ID,
        is_async: true, // 开启异步执行
        parameters: {
          keyword: keyword,
          email: email
        }
      },
      {
        headers: {
          'Authorization': `Bearer ${process.env.COZE_API_TOKEN}`,
          'Content-Type': 'application/json'
        },
        timeout: 30000 // 30秒超时
      }
    );

    console.log('✅ Workflow triggered successfully');
    console.log(`📋 Workflow response code: ${response.status}`);

    const durationMs = Date.now() - startedAt;
    const emailDomain = typeof email === 'string' && email.includes('@')
      ? email.split('@').pop().toLowerCase()
      : null;

    trackServerEvent({
      eventName: 'workflow_call',
      userId: identity.userId,
      sessionId: identity.sessionId,
      eventData: {
        channel: 'send_test',
        workflow_id: process.env.COZE_WORKFLOW_ID || null,
        status: 'success',
        duration_ms: durationMs,
      },
    }).catch((trackError) => {
      console.warn('⚠️ workflow_call track failed:', trackError?.message || trackError);
    });

    trackServerEvent({
      eventName: 'email_sent',
      userId: identity.userId,
      sessionId: identity.sessionId,
      eventData: {
        keyword,
        success: true,
        email_domain: emailDomain,
      },
    }).catch((trackError) => {
      console.warn('⚠️ email_sent track failed:', trackError?.message || trackError);
    });

    // 5. 返回成功响应
    res.json({
      success: true,
      message: '测试邮件已发送',
      data: {
        email: email,
        topic: keyword,
        workflow_response: response.data
      }
    });

  } catch (error) {
    console.error('❌ Workflow error:', error.response?.data || error.message);

    const durationMs = Date.now() - startedAt;
    const statusCode = error.response?.status || null;

    trackServerEvent({
      eventName: 'workflow_call',
      userId: identity.userId,
      sessionId: identity.sessionId,
      eventData: {
        channel: 'send_test',
        workflow_id: process.env.COZE_WORKFLOW_ID || null,
        status: 'failed',
        duration_ms: durationMs,
        http_status: statusCode,
        error_code: error.code || null,
      },
    }).catch((trackError) => {
      console.warn('⚠️ workflow_call(track failed case) error:', trackError?.message || trackError);
    });

    trackServerEvent({
      eventName: 'email_sent',
      userId: identity.userId,
      sessionId: identity.sessionId,
      eventData: {
        keyword: req.body?.keyword || null,
        success: false,
        http_status: statusCode,
        error_code: error.code || null,
      },
    }).catch((trackError) => {
      console.warn('⚠️ email_sent(track failed case) error:', trackError?.message || trackError);
    });

    // 处理Coze API错误
    if (error.response) {
      const status = error.response.status;
      const data = error.response.data;
      
      return res.status(status).json({
        error: 'Coze API调用失败',
        status: status,
        details: data
      });
    }

    // 处理超时
    if (error.code === 'ECONNABORTED') {
      return res.status(408).json({ 
        error: '请求超时，请稍后重试' 
      });
    }

    // 处理网络错误
    if (error.code === 'ENOTFOUND' || error.code === 'ECONNREFUSED') {
      return res.status(503).json({ 
        error: '无法连接到Coze服务' 
      });
    }

    // 其他错误
    res.status(500).json({ 
      error: '发送失败，请稍后重试',
      message: process.env.NODE_ENV === 'development' ? error.message : undefined
    });
  }
});

// ── 邮件发送 (供 Coze HTTP 请求节点调用) ──────────────────────────
const mailer = nodemailer.createTransport({
  host: 'smtp.qq.com',
  port: 465,
  secure: true,
  auth: {
    user: process.env.MAIL_USER,
    pass: process.env.MAIL_SMTP_PASS
  }
});

/**
 * POST /api/workflow/send-mail
 * Coze 工作流通过 HTTP 请求节点调用，自定义发件人名称发送邮件
 */
router.post('/send-mail', async (req, res) => {
  try {
    const { email, subject, html_content, token } = req.body;

    if (token !== process.env.MAIL_API_TOKEN) {
      return res.status(403).json({ error: '无权限' });
    }

    if (!email || !subject || !html_content) {
      return res.status(400).json({
        error: '缺少必要参数',
        required: ['email', 'subject', 'html_content']
      });
    }

    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(email)) {
      return res.status(400).json({ error: '邮箱格式不正确' });
    }

    console.log(`📨 [send-mail] Sending to: ${maskEmail(email)}`);

    const info = await mailer.sendMail({
      from: `"newsflow.ai" <${process.env.MAIL_USER}>`,
      to: email,
      subject,
      html: html_content
    });

    console.log(`✅ [send-mail] Sent, messageId: ${info.messageId}`);
    res.json({ success: true, message_id: info.messageId });
  } catch (error) {
    console.error('❌ [send-mail] Error:', error.message);
    res.status(500).json({
      error: '邮件发送失败',
      message: process.env.NODE_ENV === 'development' ? error.message : undefined
    });
  }
});

module.exports = router;
