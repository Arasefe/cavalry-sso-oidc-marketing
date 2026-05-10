require("dotenv").config();
const express = require("express");
const session = require("express-session");
const msal = require("@azure/msal-node");

const app = express();
const PORT = process.env.PORT || 4003;

const msalConfig = {
  auth: {
    clientId: process.env.CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
    clientSecret: process.env.CLIENT_SECRET,
  },
  system: { loggerOptions: { loggerCallback: () => {}, piiLoggingEnabled: false, logLevel: msal.LogLevel.Warning } },
};

const pca = new msal.ConfidentialClientApplication(msalConfig);
const cryptoProvider = new msal.CryptoProvider();

app.use(session({
  secret: process.env.SESSION_SECRET || "marketing-secret",
  resave: false,
  saveUninitialized: false,
  cookie: { secure: false, maxAge: 8 * 60 * 60 * 1000 },
}));

function isAuthenticated(req, res, next) {
  if (req.session.account) return next();
  res.redirect("/");
}

function decodeToken(token) {
  try { return JSON.parse(Buffer.from(token.split(".")[1], "base64url").toString("utf8")); }
  catch { return {}; }
}

const C = { primary: "#6d28d9", accent: "#7c3aed", red: "#dc2626" };

const head = `
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width,initial-scale=1"/>
  <link href="https://cdn.jsdelivr.net/npm/tailwindcss@2.2.19/dist/tailwind.min.css" rel="stylesheet"/>
  <style>
    body { background:#0d0a1a; color:#e2e8f0; font-family:'Segoe UI',sans-serif; }
    .card { background:#130f24; border:1px solid #3b1f6b; border-radius:12px; }
    .claim-row:hover { background:#1a1335; }
    pre { white-space:pre-wrap; word-break:break-all; }
    .badge { display:inline-block; padding:2px 10px; border-radius:9999px; font-size:11px; font-weight:600; }
  </style>`;

const nav = (user) => `
  <nav style="background:#080513;border-bottom:1px solid #3b1f6b;" class="px-8 py-4 flex items-center justify-between">
    <div class="flex items-center gap-3">
      <div style="background:${C.accent};border-radius:8px;" class="p-2">
        <span style="font-size:18px;">📣</span>
      </div>
      <div>
        <div class="font-bold text-white text-sm">Marketing Hub</div>
        <div style="color:#a78bfa;font-size:11px;">SSO · OIDC · Microsoft Entra ID</div>
      </div>
    </div>
    <div class="flex items-center gap-4">
      ${user
        ? `<span style="color:#9ca3af;font-size:13px;">Signed in as <strong style="color:#a78bfa;">${user.name || user.username}</strong></span>
           <a href="/dashboard" style="background:#3b1f6b;" class="px-4 py-2 rounded-lg text-sm text-white">Dashboard</a>
           <a href="/auth/logout" style="background:${C.red};" class="px-4 py-2 rounded-lg text-sm text-white font-semibold">Sign Out</a>`
        : `<a href="/auth/login" style="background:${C.accent};" class="px-5 py-2 rounded-lg text-sm font-semibold text-white">Sign In with Microsoft</a>`
      }
    </div>
  </nav>`;

// ── Home ──────────────────────────────────────────────────────────────────────
app.get("/", (req, res) => {
  if (req.session.account) return res.redirect("/dashboard");
  res.send(`<!DOCTYPE html><html><head>${head}<title>Marketing — SSO OIDC</title></head><body>
    ${nav(null)}
    <div class="max-w-3xl mx-auto px-6 py-16 text-center space-y-8">
      <div>
        <div style="font-size:64px;">📣</div>
        <h1 class="text-3xl font-bold text-white mt-4">Marketing Hub</h1>
        <p style="color:#a78bfa;" class="text-lg font-semibold mt-1">cavalry-sso-oidc-marketing</p>
        <p style="color:#9ca3af;" class="mt-3 text-sm max-w-xl mx-auto">
          Focuses on OAuth 2.0 scopes and the access token. Every OIDC app requests specific
          permission scopes — this app shows what those scopes grant and how the access token differs from the ID token.
        </p>
      </div>
      <div class="card p-6 text-left space-y-3">
        <h3 class="font-bold text-white text-sm">This app teaches:</h3>
        ${[
          ["🎯", "OAuth 2.0 Scopes", "openid · profile · email · User.Read — what each scope unlocks"],
          ["🪙", "Access Token vs ID Token", "ID token = who you are. Access token = what you can do"],
          ["📦", "Token audience (aud)", "Access tokens are issued for a specific API, not the app itself"],
          ["🔒", "Least privilege", "Apps should request only the scopes they need — never broader"],
        ].map(([icon, title, desc]) => `
          <div class="flex gap-3 items-start">
            <span>${icon}</span>
            <div><div class="text-sm font-semibold text-white">${title}</div>
            <div style="color:#9ca3af;font-size:12px;">${desc}</div></div>
          </div>`).join("")}
      </div>
      <a href="/auth/login" style="background:linear-gradient(135deg,${C.accent},#4c1d95);"
         class="inline-flex items-center gap-3 px-8 py-4 rounded-xl text-white font-semibold text-base hover:opacity-90 transition shadow-lg">
        <svg width="20" height="20" viewBox="0 0 21 21" fill="none">
          <rect x="1" y="1" width="9" height="9" fill="#f25022"/>
          <rect x="11" y="1" width="9" height="9" fill="#7fba00"/>
          <rect x="1" y="11" width="9" height="9" fill="#00a4ef"/>
          <rect x="11" y="11" width="9" height="9" fill="#ffb900"/>
        </svg>
        Sign In with Microsoft Entra ID
      </a>
      <p style="color:#1f2937;font-size:12px;">Port 4003 · OIDC · cavalry-sso-oidc-marketing</p>
    </div>
  </body></html>`);
});

// ── Auth ──────────────────────────────────────────────────────────────────────
app.get("/auth/login", async (req, res) => {
  try {
    const { verifier, challenge } = await cryptoProvider.generatePkceCodes();
    req.session.pkceCodes = { verifier, challenge };
    req.session.signedInAt = new Date().toISOString();
    const authUrl = await pca.getAuthCodeUrl({
      scopes: ["openid", "profile", "email", "User.Read"],
      redirectUri: process.env.REDIRECT_URI,
      codeChallenge: challenge,
      codeChallengeMethod: "S256",
    });
    res.redirect(authUrl);
  } catch (err) {
    res.send(errorPage("Login Failed", err.message));
  }
});

app.get("/auth/callback", async (req, res) => {
  try {
    const tokenResponse = await pca.acquireTokenByCode({
      code: req.query.code,
      scopes: ["openid", "profile", "email", "User.Read"],
      redirectUri: process.env.REDIRECT_URI,
      codeVerifier: req.session.pkceCodes?.verifier,
    });
    req.session.account = tokenResponse.account;
    req.session.idToken = tokenResponse.idToken;
    req.session.accessToken = tokenResponse.accessToken;
    res.redirect("/dashboard");
  } catch (err) {
    res.send(errorPage("Authentication Failed", err.message));
  }
});

app.get("/auth/logout", (req, res) => {
  req.session.destroy();
  res.redirect(`https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/logout?post_logout_redirect_uri=${encodeURIComponent(`http://localhost:${PORT}`)}`);
});

// ── Dashboard ─────────────────────────────────────────────────────────────────
app.get("/dashboard", isAuthenticated, (req, res) => {
  const account = req.session.account;
  const idClaims  = req.session.idToken    ? decodeToken(req.session.idToken)    : {};
  const atClaims  = req.session.accessToken? decodeToken(req.session.accessToken): {};
  const signedInAt = req.session.signedInAt
    ? new Date(req.session.signedInAt).toLocaleString("en-GB", { dateStyle: "medium", timeStyle: "medium" })
    : "—";

  const SCOPES = [
    { scope: "openid",    grants: "Authenticate the user and get a unique sub identifier", required: true  },
    { scope: "profile",   grants: "Access name, preferred_username, and other profile claims", required: true  },
    { scope: "email",     grants: "Access the user's email address claim", required: false },
    { scope: "User.Read", grants: "Read the signed-in user's profile from Microsoft Graph API", required: false },
  ];

  const tokenRows = (claims) => Object.entries(claims).map(([k, v]) => `
    <div class="claim-row px-5 py-2" style="display:grid;grid-template-columns:160px 1fr;gap:8px;font-size:12px;">
      <code style="color:#f59e0b;">${k}</code>
      <span style="font-family:monospace;color:#e2e8f0;word-break:break-all;">${
        ["exp","iat","nbf"].includes(k)
          ? `${v} <span style="color:#64748b;">(${new Date(v*1000).toLocaleString()})</span>`
          : String(v)
      }</span>
    </div>`).join("");

  res.send(`<!DOCTYPE html><html><head>${head}<title>Marketing — Dashboard</title></head><body>
    ${nav(account)}

    <!-- Success Banner -->
    <div style="background:linear-gradient(90deg,#080513,#130f24);border-bottom:1px solid #3b1f6b;" class="px-8 py-6 flex items-center gap-5">
      <div style="font-size:40px;">✅</div>
      <div>
        <h2 class="text-xl font-bold text-white">Sign In Successful</h2>
        <p style="color:#a78bfa;" class="text-sm">Welcome, <strong>${account.name || account.username}</strong> — authenticated via Microsoft Entra ID at ${signedInAt}</p>
        <p style="color:#3b1f6b;font-size:12px;" class="mt-1">Scopes granted: openid · profile · email · User.Read</p>
      </div>
      <div class="ml-auto">
        <a href="/auth/logout" style="background:${C.red};" class="px-5 py-2 rounded-lg text-sm text-white font-semibold">Sign Out</a>
      </div>
    </div>

    <div class="max-w-4xl mx-auto px-6 py-8 space-y-8">

      <!-- Scopes panel -->
      <div class="card p-6 space-y-4">
        <h3 class="font-bold text-white">OAuth 2.0 Scopes Requested by This App</h3>
        <p style="color:#9ca3af;font-size:13px;">Scopes are permissions the app requested from Entra ID. The user (or admin) consents to these. Each scope unlocks specific claims or API access.</p>
        <div class="space-y-3">
          ${SCOPES.map(({ scope, grants, required }) => `
            <div style="background:#080513;border:1px solid #3b1f6b;border-radius:8px;" class="p-4 flex gap-4 items-start">
              <code style="color:#a78bfa;font-size:14px;font-weight:700;min-width:100px;">${scope}</code>
              <div class="flex-1">
                <div style="color:#e2e8f0;font-size:13px;">${grants}</div>
              </div>
              <span class="badge" style="background:${required ? "#052e16" : "#1e1b4b"};color:${required ? "#4ade80" : "#818cf8"};white-space:nowrap;">
                ${required ? "Core (OIDC)" : "Optional"}
              </span>
            </div>`).join("")}
        </div>
      </div>

      <!-- ID Token vs Access Token -->
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;">

        <!-- ID Token -->
        <div class="card overflow-hidden">
          <div style="background:#1a0f35;border-bottom:1px solid #3b1f6b;" class="px-5 py-3">
            <h3 class="font-bold text-white text-sm">ID Token</h3>
            <p style="color:#a78bfa;font-size:11px;">Who you are — decoded JWT claims</p>
          </div>
          <div class="divide-y" style="border-color:#3b1f6b;">
            ${tokenRows(idClaims)}
          </div>
        </div>

        <!-- Access Token -->
        <div class="card overflow-hidden">
          <div style="background:#1a0f35;border-bottom:1px solid #3b1f6b;" class="px-5 py-3">
            <h3 class="font-bold text-white text-sm">Access Token</h3>
            <p style="color:#a78bfa;font-size:11px;">What you can do — grants API access</p>
          </div>
          <div class="divide-y" style="border-color:#3b1f6b;">
            ${Object.keys(atClaims).length > 0
              ? tokenRows(atClaims)
              : `<div class="px-5 py-4 text-sm" style="color:#64748b;">Access token is opaque for Microsoft Graph — use ID token claims instead.</div>`}
          </div>
        </div>
      </div>

      <!-- Key distinction -->
      <div class="card p-5" style="border-color:#7c3aed;">
        <h3 class="font-bold text-white text-sm mb-3">ID Token vs Access Token — Key Difference</h3>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;font-size:13px;">
          <div>
            <div style="color:#a78bfa;font-weight:700;margin-bottom:6px;">ID Token</div>
            <ul style="color:#9ca3af;line-height:2;list-style:disc;padding-left:16px;">
              <li>Tells the app <em>who you are</em></li>
              <li>Consumed by the app itself</li>
              <li>aud = this app's Client ID</li>
              <li>Contains name, email, oid, tid</li>
            </ul>
          </div>
          <div>
            <div style="color:#34d399;font-weight:700;margin-bottom:6px;">Access Token</div>
            <ul style="color:#9ca3af;line-height:2;list-style:disc;padding-left:16px;">
              <li>Grants access to an <em>API</em></li>
              <li>Sent to Microsoft Graph or your API</li>
              <li>aud = the target API's identifier</li>
              <li>Contains scp (scopes granted)</li>
            </ul>
          </div>
        </div>
      </div>

      <div class="text-center pb-8">
        <a href="/auth/logout" style="background:#1f2937;" class="inline-flex items-center gap-2 px-6 py-3 rounded-lg text-sm text-white">
          Sign Out &amp; Clear Session
        </a>
      </div>
    </div>
  </body></html>`);
});

function errorPage(title, msg) {
  return `<!DOCTYPE html><html><head>${head}<title>Error</title></head><body>
    <div class="max-w-xl mx-auto px-6 py-20 text-center space-y-4">
      <div style="font-size:48px;">⚠️</div>
      <h1 class="text-2xl font-bold text-white">${title}</h1>
      <div style="background:#130f24;border:1px solid #7f1d1d;border-radius:8px;" class="p-4 text-left">
        <code style="color:#fca5a5;font-size:12px;">${msg}</code>
      </div>
      <a href="/" style="color:#a78bfa;" class="text-sm underline">← Back to Home</a>
    </div></body></html>`;
}

app.listen(PORT, () => {
  console.log(`cavalry-sso-oidc-marketing →  http://localhost:${PORT}`);
  console.log(`Redirect URI:                 http://localhost:${PORT}/auth/callback`);
});
