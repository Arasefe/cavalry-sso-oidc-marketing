# cavalry-sso-oidc-marketing

A Microsoft Entra ID OIDC SSO demo app themed around a Marketing Hub. Part of the Cavalry SSO demo suite.

**Port:** `4003` · **Protocol:** OIDC (OpenID Connect) · **Library:** `@azure/msal-node`

---

## Teaching Focus

This app teaches **OAuth 2.0 scopes** and the difference between the **ID Token and Access Token**.

Students will see:
- The four scopes this app requests — `openid`, `profile`, `email`, `User.Read` — and what each one grants
- The decoded ID Token (who you are) side by side with the Access Token (what you can do)
- Why the `aud` claim differs between the two tokens — ID token audience = Client ID, Access token audience = the target API
- The principle of least privilege — apps should request only the scopes they need

---

## Prerequisites

- Node.js 18+
- A Microsoft Entra ID tenant
- An App Registration in Entra ID (see setup below)

---

## Entra ID App Registration Setup

1. Go to **Entra ID → App Registrations → New registration**
2. Name: `cavalry-sso-oidc-marketing`
3. Supported account types: **Accounts in this organizational directory only**
4. Platform: **Web** — Redirect URI: `http://localhost:4003/auth/callback`
5. Click **Register**
6. Note the **Application (Client) ID** and **Directory (Tenant) ID** from Overview
7. Go to **Certificates & secrets → New client secret** — copy the secret value immediately
8. Go to **API Permissions** — confirm `User.Read` (Microsoft Graph) is listed. Grant admin consent if required.

> Each app in the suite needs its own App Registration with its own Client ID and Client Secret.
> They share the same Tenant ID.

---

## Setup

```bash
git clone https://github.com/Arasefe/cavalry-sso-oidc-marketing.git
cd cavalry-sso-oidc-marketing
npm install
cp .env.example .env
```

Edit `.env`:

```env
TENANT_ID=your-tenant-id
CLIENT_ID=your-client-id
CLIENT_SECRET=your-client-secret
REDIRECT_URI=http://localhost:4003/auth/callback
SESSION_SECRET=any-long-random-string
PORT=4003
```

---

## Run

```bash
npm start
```

Navigate to `http://localhost:4003` and click **Sign in with Microsoft Entra ID**.

---

## Endpoints

| Endpoint | Description |
|---|---|
| `GET /` | Landing page with sign-in button |
| `GET /auth/login` | Initiates OIDC login — redirects to Entra ID |
| `GET /auth/callback` | Receives authorization code, exchanges for tokens |
| `GET /dashboard` | Authenticated view — scopes panel, ID token vs access token comparison |
| `GET /auth/logout` | Destroys local session and redirects to Entra ID logout |

---

## Scopes Requested

| Scope | Type | Grants |
|---|---|---|
| `openid` | Core OIDC | Enables OIDC login, provides `sub` identifier |
| `profile` | Core OIDC | Provides `name`, `preferred_username`, and profile claims |
| `email` | Optional | Provides the user's email address claim |
| `User.Read` | Microsoft Graph | Read the signed-in user's profile from the Graph API |

---

## SSO Demo Sequence

```bash
# Tab 1 — Finance (sign in here first)
cd ~/Desktop/cavalry-sso-oidc-finance && npm start

# Tab 2 — HR
cd ~/Desktop/cavalry-sso-oidc-hr && npm start

# Tab 3 — Marketing (open last — compares tokens and scopes)
cd ~/Desktop/cavalry-sso-oidc-marketing && npm start
```

**Key teaching point:** Open this app last in the demo sequence. Compare its dashboard with Finance and HR — the `sid` claim will match across all three (same Entra ID session), but each app has its own `aud`, its own `iat`, and its own access token scoped to its own Client ID.

---

## Part of the Cavalry Demo Suite

| App | Port | Protocol | Focus |
|---|---|---|---|
| `cavalry-sso-oidc-finance` | 4001 | OIDC | ID Token claims & JWT structure |
| `cavalry-sso-oidc-hr` | 4002 | OIDC | SSO session reuse & `sid` claim |
| `cavalry-sso-oidc-marketing` | 4003 | OIDC | Scopes, access token vs ID token |
| `cavalry-hr-portal` | 3001 | SAML | SAML claims table |
| `cavalry-dev-tools` | 3002 | SAML | Raw SAML assertion payload |
| `cavalry-finance` | 3003 | SAML | User assignment access control |
| `cavalry-saml-app` | 3004 | SAML | SP-initiated SAML flow walkthrough |
| `cavalry-gen-five` | 3005 | SAML | Session metadata & continuity |
