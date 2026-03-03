const express = require('express');
const crypto = require('crypto');
const { createClient } = require('@supabase/supabase-js');
const { extractIdentity, trackServerEvent } = require('../lib/tracking');

const router = express.Router();

// Valid values for validation
const VALID_FEEDBACK_TYPES = ['like', 'dislike'];
const VALID_DISLIKE_REASONS = [
  'irrelevant',        // 回答不相关
  'outdated',          // 新闻过时/不准确
  'too_long',          // 内容太长
  'too_short',         // 内容太短
  'unreliable_source', // 来源不可靠
  'robotic',           // 语气/表达生硬
  'other',             // 其他问题
];
const VALID_LIKE_TAGS = [
  'accurate',  // 信息准确
  'concise',   // 摘要精炼
  'helpful',   // 很有帮助
];

function assertEnv(name) {
  const v = process.env[name];
  if (!v || !String(v).trim()) {
    const err = new Error(`Missing env: ${name}`);
    err.status = 500;
    throw err;
  }
  return String(v).trim();
}

function buildMessageHash(rawHash, messageContent) {
  if (typeof rawHash === 'string' && rawHash.trim()) return rawHash.trim();
  if (typeof messageContent !== 'string' || !messageContent.trim()) return null;
  return crypto.createHash('sha256').update(messageContent).digest('hex');
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

/**
 * POST /api/feedback
 * Submit user feedback (like or dislike) for a bot message.
 * 
 * Request Body:
 * {
 *   user_id: string (required),
 *   feedback_type: 'like' | 'dislike' (required),
 *   message_id?: string,
 *   conversation_id?: string,
 *   reasons?: string[],        // For dislikes
 *   comment?: string,          // For dislikes (especially 'other')
 *   like_tags?: string[],      // For likes
 *   message_hash?: string,     // Optional stable hash for upsert
 *   message_content?: string,  // Snapshot of the message
 *   keyword?: string           // Topic context
 * }
 */
router.post('/', async (req, res) => {
  try {
    const {
      user_id: userIdRaw,
      feedback_type: feedbackTypeRaw,
      message_id: messageIdRaw,
      conversation_id: conversationIdRaw,
      reasons: reasonsRaw,
      comment: commentRaw,
      like_tags: likeTagsRaw,
      message_hash: messageHashRaw,
      message_content: messageContentRaw,
      keyword: keywordRaw,
    } = req.body || {};

    // Validate required fields
    const userId = typeof userIdRaw === 'string' ? userIdRaw.trim() : '';
    const feedbackType = typeof feedbackTypeRaw === 'string' ? feedbackTypeRaw.trim().toLowerCase() : '';
    const identity = extractIdentity(req);

    if (!userId) {
      return res.status(400).json({ error: '缺少 user_id' });
    }

    if (!feedbackType || !VALID_FEEDBACK_TYPES.includes(feedbackType)) {
      return res.status(400).json({
        error: '无效的 feedback_type',
        valid: VALID_FEEDBACK_TYPES,
      });
    }

    // Sanitize optional fields
    const messageId = typeof messageIdRaw === 'string' ? messageIdRaw.trim() : null;
    const conversationId = typeof conversationIdRaw === 'string' ? conversationIdRaw.trim() : null;
    const comment = typeof commentRaw === 'string' ? commentRaw.trim().slice(0, 500) : null; // Limit to 500 chars
    const messageContent = typeof messageContentRaw === 'string' ? messageContentRaw.slice(0, 2000) : null; // Limit snapshot
    const keyword = typeof keywordRaw === 'string' ? keywordRaw.trim() : null;
    const messageHash = buildMessageHash(messageHashRaw, messageContent);

    if (!messageHash) {
      return res.status(400).json({ error: '缺少 message_hash 或 message_content' });
    }

    // Validate and sanitize reasons (for dislikes)
    let reasons = null;
    if (feedbackType === 'dislike' && Array.isArray(reasonsRaw)) {
      reasons = reasonsRaw
        .filter(r => typeof r === 'string' && VALID_DISLIKE_REASONS.includes(r.trim()))
        .map(r => r.trim());
      if (reasons.length === 0) reasons = null;
    }

    // Validate and sanitize like_tags (for likes)
    let likeTags = null;
    if (feedbackType === 'like' && Array.isArray(likeTagsRaw)) {
      likeTags = likeTagsRaw
        .filter(t => typeof t === 'string' && VALID_LIKE_TAGS.includes(t.trim()))
        .map(t => t.trim());
      if (likeTags.length === 0) likeTags = null;
    }

    // Build payload
    const payload = {
      user_id: userId,
      feedback_type: feedbackType,
      message_hash: messageHash,
      reasons,
      comment,
      like_tags: likeTags,
      message_content: messageContent,
      keyword,
    };

    // Upsert into Supabase (one feedback per user per message)
    const supabase = getSupabaseClient();
    const { data, error } = await supabase
      .from('feedback')
      .upsert(payload, { onConflict: 'user_id,message_hash' })
      .select('id, feedback_type, created_at')
      .single();

    if (error) {
      console.error('❌ Feedback insert error:', error);
      return res.status(502).json({
        error: '反馈提交失败',
        details: process.env.NODE_ENV === 'development' ? error : undefined,
      });
    }

    trackServerEvent({
      eventName: 'feedback',
      userId: userId || identity.userId,
      sessionId: identity.sessionId,
      eventData: {
        feedback_type: feedbackType,
        reasons: Array.isArray(reasons) ? reasons : [],
        has_comment: Boolean(comment),
      },
    }).catch((trackError) => {
      console.warn('⚠️ feedback track failed:', trackError?.message || trackError);
    });

    return res.json({
      success: true,
      feedback_id: data?.id || null,
      feedback_type: data?.feedback_type || feedbackType,
      message: feedbackType === 'like' ? '感谢您的认可！' : '感谢您的反馈，我们会持续改进！',
    });
  } catch (error) {
    console.error('❌ Feedback error:', error.message);
    return res.status(error.status || 500).json({
      error: error.message || '反馈提交失败',
    });
  }
});

/**
 * GET /api/feedback/stats
 * Get feedback statistics (for admin/analytics purposes).
 */
router.get('/stats', async (req, res) => {
  try {
    const supabase = getSupabaseClient();

    // Count likes and dislikes
    const { data: likes, error: likesError } = await supabase
      .from('feedback')
      .select('id', { count: 'exact', head: true })
      .eq('feedback_type', 'like');

    const { data: dislikes, error: dislikesError } = await supabase
      .from('feedback')
      .select('id', { count: 'exact', head: true })
      .eq('feedback_type', 'dislike');

    if (likesError || dislikesError) {
      throw new Error('Failed to fetch stats');
    }

    return res.json({
      success: true,
      stats: {
        total_likes: likes?.length || 0,
        total_dislikes: dislikes?.length || 0,
      },
    });
  } catch (error) {
    console.error('❌ Feedback stats error:', error.message);
    return res.status(500).json({ error: '获取统计失败' });
  }
});

module.exports = router;
