const { createClient } = require('@supabase/supabase-js');

const ALLOWED_TRACK_EVENTS = [
  'page_view',
  'message_sent',
  'bot_reply',
  'workflow_call',
  'subscribe',
  'email_sent',
  'feedback',
  'client_error',
];

const FRONTEND_TRACK_EVENTS = new Set([
  'page_view',
  'message_sent',
  'client_error',
]);

const ALLOWED_TRACK_EVENTS_SET = new Set(ALLOWED_TRACK_EVENTS);
const BLOCKED_FIELD_RE = /^(email|message_content|raw_message|token|authorization|password|secret|api[_-]?key)$/i;
const MAX_EVENT_DATA_BYTES = 4096;

let supabaseClient;
let didWarnMissingEnv = false;

function assertEnv(name) {
  const value = process.env[name];
  if (!value || !String(value).trim()) {
    const err = new Error(`Missing env: ${name}`);
    err.status = 500;
    throw err;
  }
  return String(value).trim();
}

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

function getSupabaseClientSafe() {
  try {
    return getSupabaseClient();
  } catch (error) {
    if (!didWarnMissingEnv) {
      didWarnMissingEnv = true;
      console.warn('⚠️ Tracking disabled:', error.message);
    }
    return null;
  }
}

function isPlainObject(value) {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value);
}

function hasBlockedField(value) {
  if (!isPlainObject(value) && !Array.isArray(value)) return false;

  const queue = [value];
  while (queue.length) {
    const current = queue.pop();

    if (Array.isArray(current)) {
      current.forEach((item) => {
        if (isPlainObject(item) || Array.isArray(item)) queue.push(item);
      });
      continue;
    }

    for (const [key, val] of Object.entries(current)) {
      if (BLOCKED_FIELD_RE.test(key)) {
        return true;
      }
      if (isPlainObject(val) || Array.isArray(val)) {
        queue.push(val);
      }
    }
  }

  return false;
}

function sanitizeEventData(eventDataRaw) {
  if (eventDataRaw == null) {
    return { ok: true, data: {}, size: 2 };
  }

  if (!isPlainObject(eventDataRaw)) {
    return { ok: false, status: 400, error: 'event_data 必须是对象' };
  }

  if (hasBlockedField(eventDataRaw)) {
    return { ok: false, status: 400, error: 'event_data 包含敏感字段' };
  }

  let normalized;
  try {
    normalized = JSON.parse(JSON.stringify(eventDataRaw));
  } catch {
    return { ok: false, status: 400, error: 'event_data 不可序列化' };
  }

  const serialized = JSON.stringify(normalized);
  const size = Buffer.byteLength(serialized, 'utf8');
  if (size > MAX_EVENT_DATA_BYTES) {
    return { ok: false, status: 400, error: `event_data 过大（最大 ${MAX_EVENT_DATA_BYTES} bytes）` };
  }

  return { ok: true, data: normalized, size };
}

function normalizeIdentityValue(value) {
  if (typeof value !== 'string') return null;
  const trimmed = value.trim();
  return trimmed || null;
}

function pickRequestValue(req, bodyKey, headerKey) {
  const body = req && isPlainObject(req.body) ? req.body : null;
  const bodyValue = body ? normalizeIdentityValue(body[bodyKey]) : null;
  if (bodyValue) return bodyValue;

  const headerValue = normalizeIdentityValue(req?.headers?.[headerKey]);
  if (headerValue) return headerValue;

  return null;
}

function extractIdentity(req) {
  return {
    userId: pickRequestValue(req, 'user_id', 'x-user-id'),
    sessionId: pickRequestValue(req, 'session_id', 'x-session-id'),
  };
}

async function trackServerEvent(payload) {
  const eventName = typeof payload?.eventName === 'string' ? payload.eventName.trim() : '';
  if (!eventName || !ALLOWED_TRACK_EVENTS_SET.has(eventName)) {
    return { ok: false, status: 400, error: '不支持的 event_name' };
  }

  const sanitized = sanitizeEventData(payload?.eventData);
  if (!sanitized.ok) {
    return sanitized;
  }

  const supabase = getSupabaseClientSafe();
  if (!supabase) {
    return { ok: false, status: 202, error: 'tracking_unavailable' };
  }

  const row = {
    user_id: normalizeIdentityValue(payload?.userId),
    session_id: normalizeIdentityValue(payload?.sessionId),
    event_name: eventName,
    event_data: sanitized.data,
  };

  const { error } = await supabase.from('events').insert(row);
  if (error) {
    console.warn('⚠️ track event failed:', error.message || error);
    return { ok: false, status: 202, error: 'tracking_write_failed' };
  }

  return { ok: true, status: 200 };
}

module.exports = {
  ALLOWED_TRACK_EVENTS,
  FRONTEND_TRACK_EVENTS,
  MAX_EVENT_DATA_BYTES,
  extractIdentity,
  sanitizeEventData,
  trackServerEvent,
};
