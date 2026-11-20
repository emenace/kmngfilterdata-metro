const crypto = require('node:crypto');

const LOGIN_WINDOW_MS = 15 * 60 * 1000;
const MAX_FAILED_ATTEMPTS = 25;

/**
 * Parse a Cookie header string into a key-value map.
 */
function parseCookies(cookieHeader) {
  const cookies = Object.create(null);
  if (typeof cookieHeader !== 'string' || cookieHeader.length === 0) {
    return cookies;
  }

  const pairs = cookieHeader.split(';');
  for (let i = 0; i < pairs.length; i += 1) {
    const pair = pairs[i];
    const eqIdx = pair.indexOf('=');
    if (eqIdx === -1) {
      continue;
    }
    const name = pair.slice(0, eqIdx).trim();
    const rawValue = pair.slice(eqIdx + 1).trim();
    if (!name) {
      continue;
    }
    try {
      cookies[name] = decodeURIComponent(rawValue);
    } catch {
      cookies[name] = rawValue;
    }
  }
  return cookies;
}

/**
 * Compare candidate password against configured password in constant time via SHA-256 digests.
 */
function verifyPassword(inputPassword, expectedPassword) {
  if (typeof inputPassword !== 'string' || typeof expectedPassword !== 'string') {
    return false;
  }
  const inputDigest = crypto.createHash('sha256').update(inputPassword, 'utf8').digest();
  const expectedDigest = crypto.createHash('sha256').update(expectedPassword, 'utf8').digest();
  return crypto.timingSafeEqual(inputDigest, expectedDigest);
}

/**
 * Build an HttpOnly Set-Cookie header string.
 */
function buildSessionCookie({ name, value, maxAgeSeconds, secure = false }) {
  const safeMaxAge = Math.max(0, Math.floor(maxAgeSeconds));
  const expiresAt =
    safeMaxAge === 0 ? new Date(0) : new Date(Date.now() + safeMaxAge * 1000);
  const parts = [
    `${name}=${encodeURIComponent(value)}`,
    'Path=/',
    'HttpOnly',
    'SameSite=Lax',
    `Max-Age=${safeMaxAge}`,
    `Expires=${expiresAt.toUTCString()}`,
  ];
  if (secure) {
    parts.push('Secure');
  }
  return parts.join('; ');
}

/**
 * Create session auth helpers and lightweight per-IP login rate limiting.
 */
function createAuthHelpers({ config, localDb }) {
  const failedAttemptsByIp = new Map();

  function getClientKey(req) {
    return req.ip || (req.socket && req.socket.remoteAddress) || 'unknown';
  }

  function pruneAttempts(ip, now = Date.now()) {
    const list = failedAttemptsByIp.get(ip);
    if (!list) {
      return [];
    }
    const fresh = list.filter((ts) => now - ts < LOGIN_WINDOW_MS);
    if (fresh.length === 0) {
      failedAttemptsByIp.delete(ip);
    } else {
      failedAttemptsByIp.set(ip, fresh);
    }
    return fresh;
  }

  function isRateLimited(req) {
    const attempts = pruneAttempts(getClientKey(req));
    return attempts.length >= MAX_FAILED_ATTEMPTS;
  }

  function recordFailedAttempt(req) {
    const ip = getClientKey(req);
    const now = Date.now();
    const list = pruneAttempts(ip, now);
    list.push(now);
    failedAttemptsByIp.set(ip, list);
  }

  function clearFailedAttempts(req) {
    failedAttemptsByIp.delete(getClientKey(req));
  }

  function getSessionToken(req) {
    const cookies = parseCookies(req.headers && req.headers.cookie);
    return cookies[config.SESSION_COOKIE_NAME] || '';
  }

  function isAuthenticated(req) {
    const token = getSessionToken(req);
    return localDb.validateSession(token);
  }

  function requirePageAuth(req, res, next) {
    if (isAuthenticated(req)) {
      next();
      return;
    }
    res.redirect(302, '/login');
  }

  function requireApiAuth(req, res, next) {
    if (isAuthenticated(req)) {
      next();
      return;
    }
    res.status(401).json({ error: 'Authentication required. Please log in.' });
  }

  return {
    getSessionToken,
    isAuthenticated,
    requirePageAuth,
    requireApiAuth,
    isRateLimited,
    recordFailedAttempt,
    clearFailedAttempts,
  };
}

module.exports = {
  parseCookies,
  verifyPassword,
  buildSessionCookie,
  createAuthHelpers,
};

