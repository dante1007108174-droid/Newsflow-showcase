-- ============================================
-- Feedback Table for Daily AI News
-- ============================================
-- Purpose: Store user feedback (likes/dislikes) on AI responses
-- Created: 2026-02-02
-- ============================================

-- Create the feedback table
CREATE TABLE IF NOT EXISTS feedback (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  
  -- Message identification
  message_id TEXT,                    -- Optional: ID of the message being rated
  conversation_id TEXT,               -- Optional: Conversation context
  message_hash TEXT,                  -- Stable hash for upsert (user_id + message_hash unique)
  
  -- User identification
  user_id TEXT NOT NULL,              -- User identifier (from localStorage)
  
  -- Feedback data
  feedback_type TEXT NOT NULL CHECK (feedback_type IN ('like', 'dislike')),
  
  -- Dislike reasons (array of predefined options)
  -- Options: 'irrelevant', 'outdated', 'too_long', 'too_short', 'unreliable_source', 'robotic', 'other'
  reasons TEXT[],
  
  -- Optional user comment (for 'other' or additional context)
  comment TEXT,
  
  -- Like tags (optional quick feedback for likes)
  -- Options: 'accurate', 'concise', 'helpful'
  like_tags TEXT[],
  
  -- Snapshot data (for debugging and analysis)
  message_content TEXT,               -- The actual bot response being rated
  keyword TEXT,                       -- Topic at the time (AI/财经/科技)
  
  -- Timestamps
  created_at TIMESTAMPTZ DEFAULT NOW()
);

-- Ensure new columns exist when running on existing tables
ALTER TABLE feedback ADD COLUMN IF NOT EXISTS message_hash TEXT;

-- Create indexes for common queries
CREATE INDEX IF NOT EXISTS idx_feedback_user_id ON feedback(user_id);
CREATE UNIQUE INDEX IF NOT EXISTS idx_feedback_user_message ON feedback(user_id, message_hash);
CREATE INDEX IF NOT EXISTS idx_feedback_type ON feedback(feedback_type);
CREATE INDEX IF NOT EXISTS idx_feedback_created_at ON feedback(created_at DESC);
CREATE INDEX IF NOT EXISTS idx_feedback_keyword ON feedback(keyword);

-- Add comments for documentation
COMMENT ON TABLE feedback IS 'Stores user feedback (likes/dislikes) on AI chat responses';
COMMENT ON COLUMN feedback.feedback_type IS 'Type of feedback: like or dislike';
COMMENT ON COLUMN feedback.reasons IS 'Array of predefined dislike reasons';
COMMENT ON COLUMN feedback.like_tags IS 'Array of optional positive feedback tags';
COMMENT ON COLUMN feedback.message_content IS 'Snapshot of the rated message for analysis';
COMMENT ON COLUMN feedback.message_hash IS 'Stable hash to keep one feedback per user and message';

-- ============================================
-- Reason Code Reference:
-- ============================================
-- Dislike Reasons:
--   'irrelevant'        -> 回答不相关
--   'outdated'          -> 新闻过时/不准确
--   'too_long'          -> 内容太长
--   'too_short'         -> 内容太短
--   'unreliable_source' -> 来源不可靠
--   'robotic'           -> 语气/表达生硬
--   'other'             -> 其他问题
--
-- Like Tags:
--   'accurate'  -> 信息准确
--   'concise'   -> 摘要精炼
--   'helpful'   -> 很有帮助
-- ============================================
