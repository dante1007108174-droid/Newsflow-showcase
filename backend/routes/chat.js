const express = require('express');
const axios = require('axios');

const router = express.Router();

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function pickAssistantAnswer(messages) {
  if (!Array.isArray(messages)) return null;

  // Prefer the final assistant answer.
  for (let i = messages.length - 1; i >= 0; i -= 1) {
    const m = messages[i] || {};
    if (!isRenderableAssistantMessage(m)) continue;
    const answer = extractRenderableText(m.content, m.content_type);
    if (typeof answer === 'string' && answer.trim()) return answer;
  }

  return null;
}

function sseWrite(res, event, data) {
  res.write(`event: ${event}\n`);
  res.write(`data: ${JSON.stringify(data)}\n\n`);
  if (typeof res.flush === 'function') {
    res.flush();
  }
}

function safeJsonParse(text) {
  if (typeof text !== 'string') return null;
  const s = text.trim();
  if (!s) return null;
  try {
    return JSON.parse(s);
  } catch {
    return null;
  }
}

function isFinishSignalText(text) {
  const s = (text || '').trim();
  if (!s) return false;

  if (s.includes('"msg_type":"generate_answer_finish"')) {
    return true;
  }

  const obj = safeJsonParse(s);
  if (!obj || typeof obj !== 'object' || Array.isArray(obj)) {
    return false;
  }

  if (String(obj.msg_type || '') === 'generate_answer_finish') {
    return true;
  }

  const data = obj.data;
  if (typeof data === 'string') {
    const nested = safeJsonParse(data);
    if (nested && typeof nested === 'object' && !Array.isArray(nested)) {
      if (Object.prototype.hasOwnProperty.call(nested, 'finish_reason')) {
        const finData = nested.FinData;
        if (finData === null || finData === '') {
          return true;
        }
      }
    }
  }

  return false;
}

function extractRenderableText(content, contentType) {
  if (typeof content !== 'string') return '';
  const raw = content;
  const trimmed = raw.trim();
  if (!trimmed) {
    return '';
  }

  const normalizedContentType = typeof contentType === 'string' ? contentType.toLowerCase() : '';
  if (normalizedContentType !== 'object_string') {
    return raw;
  }

  if (isFinishSignalText(trimmed)) {
    return '';
  }

  const obj = safeJsonParse(trimmed);
  if (!obj || typeof obj !== 'object' || Array.isArray(obj)) {
    return raw;
  }

  const topKeys = ['answer', 'content', 'text', 'message'];
  for (const key of topKeys) {
    const value = obj[key];
    if (typeof value === 'string') {
      const candidate = value.trim();
      if (candidate && !isFinishSignalText(candidate)) {
        return value;
      }
    }
  }

  const data = obj.data;
  if (typeof data === 'string') {
    const nested = safeJsonParse(data);
    if (nested && typeof nested === 'object' && !Array.isArray(nested)) {
      const nestedKeys = ['answer', 'content', 'text', 'message', 'FinData'];
      for (const key of nestedKeys) {
        const value = nested[key];
        if (typeof value === 'string') {
          const candidate = value.trim();
          if (candidate && !isFinishSignalText(candidate)) {
            return value;
          }
        }
      }
      return raw;
    }

    if (String(obj.msg_type || '').includes('generate_answer')) {
      const candidate = data.trim();
      if (candidate && !isFinishSignalText(candidate)) {
        return data;
      }
    }
  }

  return raw;
}

function isRenderableAssistantMessage(messageObj) {
  const role = typeof messageObj?.role === 'string' ? messageObj.role.toLowerCase() : '';
  const contentType = typeof messageObj?.content_type === 'string' ? messageObj.content_type.toLowerCase() : '';
  const type = typeof messageObj?.type === 'string' ? messageObj.type.toLowerCase() : '';

  if (role && role !== 'assistant') return false;
  if (type && ['function_call', 'tool_output', 'tool_response', 'verbose', 'follow_up'].includes(type)) {
    return false;
  }
  if (contentType && !['text', 'markdown', 'object_string'].includes(contentType)) {
    return false;
  }

  return true;
}

async function cancelChatOnCoze({ cozeBaseUrl, headers, conversationId, chatId }) {
  const resp = await axios.post(
    `${cozeBaseUrl}/v3/chat/cancel`,
    {
      conversation_id: conversationId,
      chat_id: chatId,
    },
    {
      headers,
      timeout: 15000,
    }
  );

  return resp.data;
}

/**
 * POST /api/chat/send
 * Proxy floating-chat messages to Coze v3 chat API.
 */
router.post('/send', async (req, res) => {
  try {
    const { message, conversation_id: conversationId, user_id: userId } = req.body || {};

    if (!message || typeof message !== 'string' || !message.trim()) {
      return res.status(400).json({
        error: '缺少必要参数',
        required: ['message'],
      });
    }

    const botId = process.env.COZE_BOT_ID;
    if (!botId) {
      return res.status(500).json({
        error: '服务端未配置 COZE_BOT_ID',
      });
    }

    const cozeBaseUrl = process.env.COZE_BASE_URL;
    const token = process.env.COZE_API_TOKEN;
    if (!cozeBaseUrl || !token) {
      return res.status(500).json({
        error: '服务端未配置 Coze API 相关环境变量',
      });
    }

    // user_id is developer-defined; keep it stable per browser.
    const effectiveUserId = typeof userId === 'string' && userId.trim() ? userId.trim() : `web_${Date.now()}`;
    // Keep the conversation_id the caller intended (same as /stream).
    const requestConversationId = typeof conversationId === 'string' && conversationId.trim() ? conversationId.trim() : null;

    const trimmedMessage = message.trim();

    const createPayload = {
      bot_id: botId,
      user_id: effectiveUserId,
      stream: false,
      auto_save_history: true,
      custom_variables: {
        user_id: effectiveUserId,
      },
      additional_messages: [
        {
          role: 'user',
          content: trimmedMessage,
          content_type: 'text',
        },
      ],
    };

    const headers = {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
    };

    // Coze v3 requires conversation_id as a URL query parameter, not in the body.
    let chatUrl = `${cozeBaseUrl}/v3/chat`;
    if (requestConversationId) {
      chatUrl += `?conversation_id=${encodeURIComponent(requestConversationId)}`;
    }

    const createResp = await axios.post(chatUrl, createPayload, {
      headers,
      timeout: 30000,
    });

    const createData = createResp.data?.data;
    const chatId = createData?.id;
    const newConversationId = createData?.conversation_id;
    const status = createData?.status;

    if (!chatId || !newConversationId) {
      return res.status(502).json({
        error: 'Coze 返回数据缺失 (chat_id/conversation_id)',
        details: createResp.data,
      });
    }

    // Poll until completed if needed. Some workflow-heavy prompts can take longer.
    let finalStatus = status;
    const pollIntervalMs = Number(process.env.COZE_SEND_POLL_INTERVAL_MS) || 1000;
    const maxWaitMs = Number(process.env.COZE_SEND_MAX_WAIT_MS) || 60000;
    const maxPolls = Math.max(1, Math.ceil(maxWaitMs / pollIntervalMs));
    const terminalStatuses = new Set(['completed', 'failed', 'cancelled', 'requires_action']);

    for (let i = 0; i < maxPolls && finalStatus && !terminalStatuses.has(finalStatus); i += 1) {
      await sleep(pollIntervalMs);
      try {
        const retrieveUrl = `${cozeBaseUrl}/v3/chat/retrieve?chat_id=${encodeURIComponent(chatId)}&conversation_id=${encodeURIComponent(newConversationId)}`;
        const retrieveResp = await axios.get(retrieveUrl, { headers, timeout: 15000 });
        finalStatus = retrieveResp.data?.data?.status;
      } catch (e) {
        // Some deployments use POST for retrieve; try once.
        const retrieveUrl = `${cozeBaseUrl}/v3/chat/retrieve?chat_id=${encodeURIComponent(chatId)}&conversation_id=${encodeURIComponent(newConversationId)}`;
        const retrieveResp = await axios.post(retrieveUrl, {}, { headers, timeout: 15000 });
        finalStatus = retrieveResp.data?.data?.status;
      }
    }

    if (finalStatus !== 'completed') {
      return res.status(504).json({
        error: 'Coze 对话响应超时',
        chat_id: chatId,
        conversation_id: newConversationId,
        status: finalStatus,
      });
    }

    const listUrl = `${cozeBaseUrl}/v3/chat/message/list?chat_id=${encodeURIComponent(chatId)}&conversation_id=${encodeURIComponent(newConversationId)}`;
    const listResp = await axios.post(listUrl, {}, { headers, timeout: 30000 });
    const messages = listResp.data?.data;
    const answer = pickAssistantAnswer(messages) || '';

    if (!answer) {
      return res.status(502).json({
        error: '未获取到智能体回复',
        chat_id: chatId,
        conversation_id: newConversationId,
        details: listResp.data,
      });
    }

    return res.json({
      success: true,
      conversation_id: requestConversationId || newConversationId,
      chat_id: chatId,
      user_id: effectiveUserId,
      answer,
    });
  } catch (error) {
    console.error('❌ Chat proxy error:', error.response?.data || error.message);

    if (error.response) {
      return res.status(error.response.status).json({
        error: 'Coze API调用失败',
        status: error.response.status,
        details: error.response.data,
      });
    }

    if (error.code === 'ECONNABORTED') {
      return res.status(408).json({ error: '请求超时，请稍后重试' });
    }

    return res.status(500).json({ error: '发送失败，请稍后重试' });
  }
});

/**
 * POST /api/chat/cancel
 * Cancel an in-progress Coze chat.
 */
router.post('/cancel', async (req, res) => {
  try {
    const { conversation_id: conversationId, chat_id: chatId } = req.body || {};

    if (!conversationId || !chatId) {
      return res.status(400).json({
        error: '缺少必要参数',
        required: ['conversation_id', 'chat_id'],
      });
    }

    const cozeBaseUrl = process.env.COZE_BASE_URL;
    const token = process.env.COZE_API_TOKEN;
    if (!cozeBaseUrl || !token) {
      return res.status(500).json({
        error: '服务端未配置 Coze API 相关环境变量',
      });
    }

    const headers = {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
    };

    const data = await cancelChatOnCoze({
      cozeBaseUrl,
      headers,
      conversationId,
      chatId,
    });

    return res.json({ success: true, data });
  } catch (error) {
    console.error('❌ Cancel chat error:', error.response?.data || error.message);
    if (error.response) {
      return res.status(error.response.status).json({
        error: 'Coze API调用失败',
        status: error.response.status,
        details: error.response.data,
      });
    }
    return res.status(500).json({ error: '取消失败，请稍后重试' });
  }
});

/**
 * POST /api/chat/stream
 * Streaming chat proxy using SSE (consumed via fetch streaming on frontend).
 */
router.post('/stream', async (req, res) => {
  const { message, conversation_id: conversationId, user_id: userId } = req.body || {};

  if (!message || typeof message !== 'string' || !message.trim()) {
    return res.status(400).json({
      error: '缺少必要参数',
      required: ['message'],
    });
  }

  const botId = process.env.COZE_BOT_ID;
  const cozeBaseUrl = process.env.COZE_BASE_URL;
  const token = process.env.COZE_API_TOKEN;

  if (!botId) {
    return res.status(500).json({ error: '服务端未配置 COZE_BOT_ID' });
  }
  if (!cozeBaseUrl || !token) {
    return res.status(500).json({ error: '服务端未配置 Coze API 相关环境变量' });
  }

  // SSE headers
  res.status(200);
  res.setHeader('Content-Type', 'text/event-stream; charset=utf-8');
  res.setHeader('Cache-Control', 'no-cache, no-transform');
  res.setHeader('Connection', 'keep-alive');
  res.setHeader('X-Accel-Buffering', 'no');
  res.flushHeaders?.();

  const headers = {
    Authorization: `Bearer ${token}`,
    'Content-Type': 'application/json',
  };

  const effectiveUserId = typeof userId === 'string' && userId.trim() ? userId.trim() : `web_${Date.now()}`;
  // Keep the conversation_id the caller intended so we never accidentally
  // switch conversations when Coze returns a different id (e.g. workflow fork).
  const requestConversationId = typeof conversationId === 'string' && conversationId.trim() ? conversationId.trim() : null;

  const trimmedMessage = message.trim();

  const createPayload = {
    bot_id: botId,
    user_id: effectiveUserId,
    stream: true,
    auto_save_history: true,
    custom_variables: {
      user_id: effectiveUserId,
    },
    additional_messages: [
      {
        role: 'user',
        content: trimmedMessage,
        content_type: 'text',
      },
    ],
  };

  // Coze v3 requires conversation_id as a URL query parameter, not in the body.
  let chatUrl = `${cozeBaseUrl}/v3/chat`;
  if (requestConversationId) {
    chatUrl += `?conversation_id=${encodeURIComponent(requestConversationId)}`;
  }

  const abortController = new AbortController();
  let upstreamConversationId = null;
  let upstreamChatId = null;
  let cancelled = false;

  const onAbort = () => {
    cancelled = true;
    abortController.abort();
  };
  req.on('aborted', onAbort);
  res.on('close', () => {
    if (!res.writableEnded) {
      cancelled = true;
      abortController.abort();
    }
  });

  try {
    const upstream = await axios.post(chatUrl, createPayload, {
      headers,
      responseType: 'stream',
      timeout: 0,
      signal: abortController.signal,
      validateStatus: () => true,
    });

    if (upstream.status < 200 || upstream.status >= 300) {
      sseWrite(res, 'error', {
        error: 'Coze API调用失败',
        status: upstream.status,
        details: upstream.data,
      });
      res.end();
      return;
    }

    const stream = upstream.data;
    stream.setEncoding('utf8');

    let buffer = '';
    let metaSent = false;

    const flushBuffer = () => {
      // Coze streaming is SSE-like: event: xxx\ndata: {...}\n\n
      while (true) {
        const sep = buffer.indexOf('\n\n');
        if (sep === -1) break;

        const rawEvent = buffer.slice(0, sep);
        buffer = buffer.slice(sep + 2);

        const lines = rawEvent.split('\n');
        let eventName = null;
        const dataLines = [];

        for (const line of lines) {
          if (line.startsWith('event:')) {
            eventName = line.slice('event:'.length).trim();
          } else if (line.startsWith('data:')) {
            dataLines.push(line.slice('data:'.length).trim());
          }
        }

        if (!eventName) continue;
        const dataStr = dataLines.join('\n');

        if (eventName === 'done') {
          sseWrite(res, 'done', {});
          res.end();
          return;
        }

        if (eventName === 'error') {
          sseWrite(res, 'error', { error: 'Coze error', details: dataStr });
          continue;
        }

        if (!dataStr) continue;

        let obj;
        try {
          obj = JSON.parse(dataStr);
        } catch {
          // If parsing fails, still pass raw data through.
          sseWrite(res, 'raw', { event: eventName, data: dataStr });
          continue;
        }

        if (!metaSent && (eventName === 'conversation.chat.created' || eventName === 'conversation.chat.in_progress')) {
          upstreamChatId = obj.id || upstreamChatId;
          upstreamConversationId = obj.conversation_id || upstreamConversationId;
          if (upstreamChatId && upstreamConversationId) {
            metaSent = true;
            // Prefer the conversation_id the caller sent so multi-turn context
            // stays on the same thread even if Coze forks a new one for workflows.
            sseWrite(res, 'meta', {
              chat_id: upstreamChatId,
              conversation_id: requestConversationId || upstreamConversationId,
              user_id: effectiveUserId,
            });
          }
          sseWrite(res, 'event', { event: eventName, data: obj });
          continue;
        }

        if (eventName === 'conversation.message.delta') {
          upstreamChatId = obj.chat_id || upstreamChatId;
          upstreamConversationId = obj.conversation_id || upstreamConversationId;

          if (!metaSent && upstreamChatId && upstreamConversationId) {
            metaSent = true;
            sseWrite(res, 'meta', {
              chat_id: upstreamChatId,
              conversation_id: requestConversationId || upstreamConversationId,
              user_id: effectiveUserId,
            });
          }

          if (!isRenderableAssistantMessage(obj)) {
            continue;
          }

          const delta = extractRenderableText(obj.content, obj.content_type);
          if (delta) {
            sseWrite(res, 'delta', { content: delta });
          }
          continue;
        }

        if (eventName === 'conversation.message.completed') {
          if (!isRenderableAssistantMessage(obj)) {
            continue;
          }

          const finalContent = extractRenderableText(obj.content, obj.content_type);
          if (finalContent) {
            sseWrite(res, 'final', { content: finalContent });
          }
          continue;
        }

        if (eventName === 'conversation.chat.completed') {
          sseWrite(res, 'event', { event: eventName, data: obj });
          // Let terminal 'done' event close the stream.
          continue;
        }

        // Forward other events (e.g., requires_action) for future handling.
        sseWrite(res, 'event', { event: eventName, data: obj });
      }
    };

    stream.on('data', (chunk) => {
      buffer += chunk;
      buffer = buffer.replace(/\r\n/g, '\n');
      flushBuffer();
    });

    stream.on('end', () => {
      if (!res.writableEnded) {
        sseWrite(res, 'done', {});
        res.end();
      }
    });

    stream.on('error', (err) => {
      if (!res.writableEnded) {
        sseWrite(res, 'error', { error: err.message || 'Upstream stream error' });
        res.end();
      }
    });

    // Keep-alive ping
    const ping = setInterval(() => {
      if (res.writableEnded) {
        clearInterval(ping);
        return;
      }
      res.write(': ping\n\n');
    }, 15000);

    res.on('close', async () => {
      clearInterval(ping);

      // Best-effort: also tell Coze to cancel if we have ids.
      if (cancelled && upstreamConversationId && upstreamChatId) {
        try {
          await cancelChatOnCoze({
            cozeBaseUrl,
            headers,
            conversationId: upstreamConversationId,
            chatId: upstreamChatId,
          });
        } catch {
          // Ignore.
        }
      }
    });
  } catch (error) {
    if (abortController.signal.aborted) {
      // Client disconnected or cancelled.
      if (!res.writableEnded) {
        sseWrite(res, 'done', { cancelled: true });
        res.end();
      }
      return;
    }

    console.error('❌ Chat stream proxy error:', error.response?.data || error.message);
    if (!res.writableEnded) {
      sseWrite(res, 'error', {
        error: '发送失败，请稍后重试',
        details: error.response?.data || error.message,
      });
      res.end();
    }
  }
});

module.exports = router;
