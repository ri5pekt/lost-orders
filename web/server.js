import cookieParser from "cookie-parser";
import { config } from "dotenv";
import express from "express";
import jwt from "jsonwebtoken";
import passport from "passport";
import { Strategy as GoogleStrategy } from "passport-google-oauth20";
import { dirname, join } from "path";
import { Readable } from "stream";
import { fileURLToPath } from "url";

config();

const __dirname = dirname(fileURLToPath(import.meta.url));
const PDF_SERVICE = process.env.PDF_SERVICE_URL || "http://pdf-service:5000";
const ALLOWED_DOMAIN = process.env.ALLOWED_DOMAIN || "particleformen.com";

const app = express();
app.use(express.json());
app.use(cookieParser());

// ── Google OAuth ──────────────────────────────────────────────────────────────

passport.use(
  new GoogleStrategy(
    {
      clientID: process.env.GOOGLE_CLIENT_ID,
      clientSecret: process.env.GOOGLE_CLIENT_SECRET,
      callbackURL: process.env.GOOGLE_CALLBACK_URL,
    },
    (_accessToken, _refreshToken, profile, done) => {
      const email = profile.emails?.[0]?.value || "";
      if (!email.endsWith(`@${ALLOWED_DOMAIN}`)) {
        return done(null, false);
      }
      return done(null, {
        email,
        name: profile.displayName,
        picture: profile.photos?.[0]?.value || null,
      });
    }
  )
);

app.use(passport.initialize());

app.get("/auth/google", passport.authenticate("google", { scope: ["profile", "email"] }));

app.get(
  "/auth/google/callback",
  passport.authenticate("google", {
    session: false,
    failureRedirect: "/login?error=unauthorized",
  }),
  (req, res) => {
    const token = jwt.sign(req.user, process.env.JWT_SECRET, { expiresIn: "7d" });
    res.cookie("token", token, {
      httpOnly: true,
      secure: process.env.NODE_ENV === "production",
      sameSite: "lax",
      maxAge: 7 * 24 * 60 * 60 * 1000,
    });
    res.redirect("/");
  }
);

app.post("/auth/logout", (_req, res) => {
  res.clearCookie("token");
  res.json({ ok: true });
});

// ── Auth middleware ───────────────────────────────────────────────────────────

function requireAuth(req, res, next) {
  const token = req.cookies?.token;
  if (!token) return res.status(401).json({ error: "Unauthorized" });
  try {
    req.user = jwt.verify(token, process.env.JWT_SECRET);
    next();
  } catch {
    res.status(401).json({ error: "Unauthorized" });
  }
}

// ── API routes ────────────────────────────────────────────────────────────────

app.get("/api/me", requireAuth, (req, res) => {
  res.json({ email: req.user.email, name: req.user.name, picture: req.user.picture });
});

// Start export job — returns job_id immediately
app.post("/api/export", requireAuth, async (req, res) => {
  const { orderIds, afterDate } = req.body;

  if (!Array.isArray(orderIds) || orderIds.length === 0) {
    return res.status(400).json({ error: "No order IDs provided" });
  }

  try {
    const pdfRes = await fetch(`${PDF_SERVICE}/render`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ order_ids: orderIds, after: afterDate }),
      signal: AbortSignal.timeout(15_000),
    });

    const data = await pdfRes.json();
    if (!pdfRes.ok) return res.status(pdfRes.status).json(data);
    res.json(data); // { job_id: "..." }
  } catch (err) {
    console.error("Export start error:", err);
    res.status(500).json({ error: err.message || "Failed to start export" });
  }
});

// Stream SSE progress from pdf-service to browser
app.get("/api/export/progress/:jobId", requireAuth, async (req, res) => {
  res.set({
    "Content-Type": "text/event-stream",
    "Cache-Control": "no-cache",
    Connection: "keep-alive",
    "X-Accel-Buffering": "no",
  });
  res.flushHeaders();

  try {
    const upstream = await fetch(`${PDF_SERVICE}/progress/${req.params.jobId}`, {
      signal: AbortSignal.timeout(600_000),
    });

    if (!upstream.ok) {
      res.write(`data: ${JSON.stringify({ type: "error", message: "Job not found" })}\n\n`);
      return res.end();
    }

    Readable.fromWeb(upstream.body).pipe(res);
  } catch (err) {
    res.write(`data: ${JSON.stringify({ type: "error", message: err.message })}\n\n`);
    res.end();
  }
});

// Download completed PDF
app.get("/api/export/download/:jobId", requireAuth, async (req, res) => {
  try {
    const pdfRes = await fetch(`${PDF_SERVICE}/download/${req.params.jobId}`, {
      signal: AbortSignal.timeout(60_000),
    });

    if (!pdfRes.ok) {
      const err = await pdfRes.json().catch(() => ({ error: "Download failed" }));
      return res.status(pdfRes.status).json(err);
    }

    const missing = pdfRes.headers.get("X-Missing-Orders") || "";
    const found = pdfRes.headers.get("X-Found-Count") || "0";
    const total = pdfRes.headers.get("X-Total-Count") || "0";

    const buffer = await pdfRes.arrayBuffer();
    res.set("Content-Type", "application/pdf");
    res.set("Content-Disposition", 'attachment; filename="orders-export.pdf"');
    res.set("X-Missing-Orders", missing);
    res.set("X-Found-Count", found);
    res.set("X-Total-Count", total);
    res.set("Access-Control-Expose-Headers", "X-Missing-Orders, X-Found-Count, X-Total-Count");
    res.send(Buffer.from(buffer));
  } catch (err) {
    console.error("Download error:", err);
    res.status(500).json({ error: err.message || "Download failed" });
  }
});

// ── Serve Vue SPA ─────────────────────────────────────────────────────────────

app.use(express.static(join(__dirname, "client/dist")));
app.get("*", (_req, res) => res.sendFile(join(__dirname, "client/dist/index.html")));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on :${PORT}`));
