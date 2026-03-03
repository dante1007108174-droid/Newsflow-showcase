const express = require('express');
const { createClient } = require('@supabase/supabase-js');
const { extractIdentity, trackServerEvent } = require('../lib/tracking');

const router = express.Router();

const KEYWORD_ALIASES = {
  ai: 'ai',
  AI: 'ai',
  '人工智能': 'ai',
  finance: '财经',
  '财经': '财经',
  tech: '科技',
  '科技': '科技',
};

function normalizeKeyword(keyword) {
  if (typeof keyword !== 'string') return null;
  const trimmed = keyword.trim();
  return KEYWORD_ALIASES[trimmed] || null;
}

function normalizeEmail(email) {
  if (typeof email !== 'string') return '';
  return email.trim().toLowerCase();
}

function assertEnv(name) {
  const v = process.env[name];
  if (!v || !String(v).trim()) {
    const err = new Error(`Missing env: ${name}`);
    err.status = 500;
    throw err;
  }
  return String(v).trim();
}

let supabaseClient;
function getSupabaseClient() {
  if (supabaseClient) return supabaseClient;
  const supabaseUrl = assertEnv('SUPABASE_URL');
  const supabaseKey = assertEnv('SUPABASE_SERVICE_ROLE_KEY');
  supabaseClient = createClient(supabaseUrl, supabaseKey, {
    auth: {
      persistSession: false,
      autoRefreshToken: false,
      detectSessionInUrl: false,
    },
  });
  return supabaseClient;
}

async function findSubscriptionByEmail(supabase, email) {
  // Prefer exact match on normalized storage; fall back to ilike for legacy rows.
  const select = 'id, user_id, keyword, email, status, send_time';
  const exact = await supabase
    .from('subscriptions')
    .select(select)
    .eq('email', email)
    .order('updated_at', { ascending: false })
    .limit(1);

  if (exact.error) {
    const err = new Error('Supabase lookup error');
    err.status = 502;
    err.details = exact.error;
    throw err;
  }

  if (Array.isArray(exact.data) && exact.data.length) return exact.data[0];

  const { data, error } = await supabase
    .from('subscriptions')
    .select(select)
    .ilike('email', email)
    .order('updated_at', { ascending: false })
    .limit(1);

  if (error) {
    const err = new Error('Supabase lookup error');
    err.status = 502;
    err.details = error;
    throw err;
  }

  return Array.isArray(data) && data.length ? data[0] : null;
}

async function findSubscriptionByUserId(supabase, userId) {
  const { data, error } = await supabase
    .from('subscriptions')
    .select('id, user_id, keyword, email, status, send_time')
    .eq('user_id', userId)
    .order('updated_at', { ascending: false })
    .limit(1);

  if (error) {
    const err = new Error('Supabase lookup error');
    err.status = 502;
    err.details = error;
    throw err;
  }

  return Array.isArray(data) && data.length ? data[0] : null;
}

/**
 * POST /api/subscription/lookup
 * Lookup a subscription by email.
 */
router.post('/lookup', async (req, res) => {
  try {
    const { email: emailRaw } = req.body || {};
    const email = normalizeEmail(emailRaw);

    if (!email) {
      return res.status(400).json({ error: '缺少 email' });
    }
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(email)) {
      return res.status(400).json({ error: '邮箱格式不正确' });
    }

    const supabase = getSupabaseClient();
    const existing = await findSubscriptionByEmail(supabase, email);

    return res.json({
      found: Boolean(existing),
      subscription: existing || null,
    });
  } catch (error) {
    const details = error.details || error.message;
    const detailMessage =
      error?.details?.message ||
      error?.details?.error ||
      error?.message;

    console.error('❌ Subscription lookup error:', details);
    return res.status(error.status || 500).json({
      error: detailMessage || '查询订阅失败',
      details: process.env.NODE_ENV === 'development' ? details : undefined,
    });
  }
});

/**
 * POST /api/subscription/lookup-user
 * Lookup a subscription by user_id.
 */
router.post('/lookup-user', async (req, res) => {
  try {
    const { user_id: userIdRaw } = req.body || {};
    const userId = typeof userIdRaw === 'string' ? userIdRaw.trim() : '';

    if (!userId) {
      return res.status(400).json({ error: '缺少 user_id' });
    }

    const supabase = getSupabaseClient();
    const existing = await findSubscriptionByUserId(supabase, userId);

    return res.json({
      found: Boolean(existing),
      subscription: existing || null,
    });
  } catch (error) {
    const details = error.details || error.message;
    const detailMessage =
      error?.details?.message ||
      error?.details?.error ||
      error?.message;

    console.error('❌ Subscription lookup-user error:', details);
    return res.status(error.status || 500).json({
      error: detailMessage || '查询订阅失败',
      details: process.env.NODE_ENV === 'development' ? details : undefined,
    });
  }
});

/**
 * POST /api/subscription/upsert
 * Create or update a subscription record in Supabase.
 */
router.post('/upsert', async (req, res) => {
  try {
    const {
      user_id: userIdRaw,
      keyword: keywordRaw,
      email: emailRaw,
      status: statusRaw,
      confirm: confirmRaw,
    } = req.body || {};

    const userId = typeof userIdRaw === 'string' ? userIdRaw.trim() : '';
    const email = normalizeEmail(emailRaw);
    const keyword = normalizeKeyword(keywordRaw);
    const status = typeof statusRaw === 'string' && statusRaw.trim() ? statusRaw.trim() : null;
    const confirm = confirmRaw === true;
    const sendTime = '08';
    const identity = extractIdentity(req);

    const trackSubscribe = (action, replaced, unchanged) => {
      trackServerEvent({
        eventName: 'subscribe',
        userId: userId || identity.userId,
        sessionId: identity.sessionId,
        eventData: {
          action,
          keyword,
          has_email: Boolean(email),
          replaced: Boolean(replaced),
          unchanged: Boolean(unchanged),
        },
      }).catch((trackError) => {
        console.warn('⚠️ subscribe track failed:', trackError?.message || trackError);
      });
    };

    if (!userId) {
      return res.status(400).json({ error: '缺少 user_id' });
    }
    if (!email) {
      return res.status(400).json({ error: '缺少 email' });
    }
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(email)) {
      return res.status(400).json({ error: '邮箱格式不正确' });
    }
    if (!keyword) {
      return res.status(400).json({ error: '无效的 keyword', valid: ['ai', '财经', '科技'] });
    }

    const supabase = getSupabaseClient();
    const payload = {
      user_id: userId,
      keyword,
      // Store normalized email to keep application behavior consistent.
      email,
      send_time: sendTime,
    };
    if (status) payload.status = status;

    // Rule enforcement:
    // - One user_id can have only one subscription row.
    // - One email can be bound to only one user_id (case-insensitive).
    const existingByUser = await findSubscriptionByUserId(supabase, userId);
    const existingByEmail = await findSubscriptionByEmail(supabase, email);

    if (existingByEmail && existingByEmail.user_id && existingByEmail.user_id !== userId) {
      return res.status(409).json({
        error: '该邮箱已被绑定，请更换邮箱',
        code: 'EMAIL_BOUND',
      });
    }

    if (existingByUser) {
      const existingEmail = normalizeEmail(existingByUser.email);
      const unchanged = existingEmail === email && existingByUser.keyword === keyword;
      if (unchanged) {
        trackSubscribe('unchanged', true, true);
        return res.json({
          success: true,
          record_id: existingByUser.id,
          keyword: existingByUser.keyword,
          status: existingByUser.status || status || 'active',
          send_time: existingByUser.send_time || sendTime,
          replaced: true,
          unchanged: true,
        });
      }

      if (!confirm) {
        return res.status(409).json({
          error: '该账号已订阅，是否覆盖为新信息？',
          code: 'CONFIRM_UPDATE',
          current: {
            email: existingByUser.email,
            keyword: existingByUser.keyword,
          },
          next: {
            email,
            keyword,
          },
        });
      }

      const result = await supabase
        .from('subscriptions')
        .update(payload)
        .eq('id', existingByUser.id)
        .select('id, keyword, status, send_time')
        .single();

      if (result.error) {
        const err = new Error('Supabase update error');
        err.status = 502;
        err.details = result.error;
        throw err;
      }

      trackSubscribe('update', true, false);

      return res.json({
        success: true,
        record_id: result.data?.id || null,
        keyword: result.data?.keyword || keyword,
        status: result.data?.status || status || 'active',
        send_time: result.data?.send_time || sendTime,
        replaced: true,
      });
    }

    const insert = await supabase
      .from('subscriptions')
      .insert(payload)
      .select('id, keyword, status, send_time')
      .single();

    if (insert.error) {
      const err = new Error('Supabase insert error');
      err.status = 502;
      err.details = insert.error;
      throw err;
    }

    trackSubscribe('create', false, false);

    return res.json({
      success: true,
      record_id: insert.data?.id || null,
      keyword: insert.data?.keyword || keyword,
      status: insert.data?.status || status || 'active',
      send_time: insert.data?.send_time || sendTime,
      replaced: false,
    });
  } catch (error) {
    const details = error.details || error.message;
    const detailMessage =
      error?.details?.message ||
      error?.details?.error ||
      error?.message;

    console.error('❌ Subscription upsert error:', details);
    return res.status(error.status || 500).json({
      error: detailMessage || '订阅写入失败',
      details: process.env.NODE_ENV === 'development' ? details : undefined,
    });
  }
});

module.exports = router;
