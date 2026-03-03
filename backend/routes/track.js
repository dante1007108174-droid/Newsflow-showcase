const express = require('express');
const rateLimit = require('express-rate-limit');
const {
  FRONTEND_TRACK_EVENTS,
  extractIdentity,
  trackServerEvent,
} = require('../lib/tracking');

const router = express.Router();

const trackLimiter = rateLimit({
  windowMs: 60 * 1000,
  max: 60,
  standardHeaders: true,
  legacyHeaders: false,
});

router.use(trackLimiter);

/**
 * POST /api/track
 * Frontend event ingestion endpoint.
 */
router.post('/', async (req, res) => {
  try {
    const { event_name: eventNameRaw, event_data: eventData } = req.body || {};
    const eventName = typeof eventNameRaw === 'string' ? eventNameRaw.trim() : '';

    if (!eventName) {
      return res.status(400).json({ error: '缺少 event_name' });
    }

    if (!FRONTEND_TRACK_EVENTS.has(eventName)) {
      return res.status(400).json({
        error: 'event_name 不允许通过前端上报',
        allowed: Array.from(FRONTEND_TRACK_EVENTS),
      });
    }

    const identity = extractIdentity(req);
    const result = await trackServerEvent({
      eventName,
      eventData,
      userId: identity.userId,
      sessionId: identity.sessionId,
    });

    if (result.ok) {
      return res.json({ success: true });
    }

    if (result.status === 400) {
      return res.status(400).json({ error: result.error || '参数不合法' });
    }

    // Tracking failures should not break user-facing flows.
    return res.status(202).json({
      success: false,
      message: '埋点暂不可用，已降级处理',
    });
  } catch (error) {
    console.warn('⚠️ /api/track unexpected error:', error.message || error);
    return res.status(202).json({
      success: false,
      message: '埋点暂不可用，已降级处理',
    });
  }
});

module.exports = router;
