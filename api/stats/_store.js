const { del, list, put } = require("@vercel/blob");

const VISITOR_PREFIX = "stats/visitors/";
const SESSION_PREFIX = "stats/sessions/";
const ACTIVE_WINDOW_MS = 2 * 60 * 1000;
const STALE_SESSION_MS = 24 * 60 * 60 * 1000;
const ID_PATTERN = /^[a-zA-Z0-9-]{10,100}$/;

function getNow() {
  return Date.now();
}

function isValidId(value) {
  return typeof value === "string" && ID_PATTERN.test(value);
}

async function listAll(prefix) {
  let cursor;
  const blobs = [];

  do {
    const page = await list({
      prefix,
      cursor,
      limit: 1000,
      mode: "expanded",
    });

    blobs.push(...page.blobs);
    cursor = page.hasMore ? page.cursor : undefined;
  } while (cursor);

  return blobs;
}

async function ensureVisitor(visitorId, pathname, referrer) {
  const pathnameKey = `${VISITOR_PREFIX}${visitorId}.json`;

  try {
    await put(
      pathnameKey,
      JSON.stringify({
        visitorId,
        firstSeenAt: new Date().toISOString(),
        firstPathname: pathname,
        firstReferrer: referrer || "",
      }),
      {
        access: "private",
        contentType: "application/json",
      },
    );
  } catch (error) {
    const alreadyExists =
      error?.statusCode === 409 ||
      error?.status === 409 ||
      String(error?.message || "").includes("already exists");

    if (!alreadyExists) {
      throw error;
    }

    return false;
  }

  return true;
}

async function upsertSession(sessionId, visitorId, pathname, referrer) {
  const pathnameKey = `${SESSION_PREFIX}${sessionId}.json`;

  await put(
    pathnameKey,
    JSON.stringify({
      sessionId,
      visitorId,
      pathname,
      referrer: referrer || "",
      lastSeenAt: new Date().toISOString(),
    }),
    {
      access: "private",
      allowOverwrite: true,
      contentType: "application/json",
    },
  );
}

async function summarizeStats() {
  const [visitors, sessions] = await Promise.all([
    listAll(VISITOR_PREFIX),
    listAll(SESSION_PREFIX),
  ]);

  const now = getNow();
  const activeThreshold = now - ACTIVE_WINDOW_MS;
  const staleThreshold = now - STALE_SESSION_MS;
  const staleSessions = [];

  let liveVisitors = 0;

  for (const session of sessions) {
    const uploadedAt = new Date(session.uploadedAt).getTime();

    if (uploadedAt >= activeThreshold) {
      liveVisitors += 1;
    }

    if (uploadedAt < staleThreshold) {
      staleSessions.push(session.pathname);
    }
  }

  if (staleSessions.length) {
    await Promise.allSettled(staleSessions.map((pathname) => del(pathname)));
  }

  return {
    totalVisitors: visitors.length,
    liveVisitors,
    measuredAt: new Date(now).toISOString(),
  };
}

async function registerVisit(payload) {
  const { visitorId, sessionId, pathname = "/", referrer = "" } = payload || {};

  if (!isValidId(visitorId) || !isValidId(sessionId)) {
    const error = new Error("Invalid visitor/session id");
    error.statusCode = 400;
    throw error;
  }

  await Promise.all([
    ensureVisitor(visitorId, pathname, referrer),
    upsertSession(sessionId, visitorId, pathname, referrer),
  ]);

  return summarizeStats();
}

module.exports = {
  registerVisit,
  summarizeStats,
};
