/**
 * Daily AI News Frontend Logic
 * Designed & Developed by Antigravity
 */

const ANALYTICS_USER_ID_KEY = 'daily_ai_news_user_id';
const ANALYTICS_SESSION_ID_KEY = 'daily_ai_news_session_id';

function createRandomId(prefix) {
    if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
        return `${prefix}_${crypto.randomUUID()}`;
    }
    return `${prefix}_${Date.now()}_${Math.random().toString(16).slice(2, 10)}`;
}

function getOrCreateUserId() {
    try {
        const existing = localStorage.getItem(ANALYTICS_USER_ID_KEY);
        if (existing) return existing;
        const generated = createRandomId('web');
        localStorage.setItem(ANALYTICS_USER_ID_KEY, generated);
        return generated;
    } catch {
        return createRandomId('web');
    }
}

function getOrCreateSessionId() {
    try {
        const existing = sessionStorage.getItem(ANALYTICS_SESSION_ID_KEY);
        if (existing) return existing;
        const generated = createRandomId('session');
        sessionStorage.setItem(ANALYTICS_SESSION_ID_KEY, generated);
        return generated;
    } catch {
        return createRandomId('session');
    }
}

function getIdentityHeaders() {
    return {
        'X-User-Id': getOrCreateUserId(),
        'X-Session-Id': getOrCreateSessionId(),
    };
}

async function trackEvent(eventName, eventData = {}) {
    if (!eventName) return;

    try {
        await fetch('/api/track', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                ...getIdentityHeaders(),
            },
            body: JSON.stringify({
                event_name: eventName,
                user_id: getOrCreateUserId(),
                session_id: getOrCreateSessionId(),
                event_data: eventData,
            }),
            keepalive: true,
        });
    } catch (error) {
        console.warn('trackEvent failed:', error?.message || error);
    }
}

function registerGlobalErrorTracking() {
    window.addEventListener('error', (event) => {
        const message = typeof event?.message === 'string' ? event.message.slice(0, 200) : 'unknown';
        const stack = typeof event?.error?.stack === 'string' ? event.error.stack.slice(0, 500) : null;

        trackEvent('client_error', {
            error_type: 'window_error',
            message,
            source: event?.filename || null,
            line: Number.isFinite(event?.lineno) ? event.lineno : null,
            stack,
        });
    });

    window.addEventListener('unhandledrejection', (event) => {
        const reason = typeof event?.reason === 'string'
            ? event.reason
            : event?.reason?.message || 'unhandled_rejection';

        trackEvent('client_error', {
            error_type: 'unhandled_rejection',
            message: String(reason).slice(0, 200),
        });
    });
}

document.addEventListener('DOMContentLoaded', () => {
    // --- Analytics (must run first, independently) ---
    getOrCreateUserId();
    getOrCreateSessionId();
    registerGlobalErrorTracking();

    trackEvent('page_view', {
        url: window.location.pathname,
        referrer: document.referrer || '',
        ua: navigator.userAgent,
    });
});
