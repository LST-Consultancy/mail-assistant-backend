# Mail Assistant: backend API

The backend for the Mail Assistant Outlook add-in. The add-in sends it trimmed email text, and it returns AI summaries, reply suggestions, digests and drafts. The add-in lives in its own repo (`mail-assistant-addin`).

**It never touches the mailbox.** It never receives Graph tokens, and it contains no Graph calls.

> **AI provider:** OpenAI, using the official `openai` SDK. Set `LLM_API_KEY` to an OpenAI API key.

`REQUIREMENTS.md` is the spec and the source of truth. Its section 16 lists the decisions taken and where the build differs from the spec.

Stack: Node 22 + Express 5 + TypeScript, openai (OpenAI-compatible client), zod, jose. Ships as a Docker image.

## How it works

```
Outlook task pane ── API token ──► Backend /api/*    (trimmed email text only, never Graph tokens)
                                     ├─ validates the Entra JWT (JWKS, aud, iss = token tenant, tenant allow-list, scope)
                                     ├─ rate limit (60/min/user) + daily token budget per user
                                     └─ LLM: JSON mode validated with zod (one retry) for analysis, streaming (SSE) for replies/compose
```

- **Prompt-injection safety:**
  - Email text is wrapped in `<email>` tags.
  - Look-alike tags inside an email are removed, so an email can't close its own tag.
  - The system prompt says to treat email text as data and never follow instructions in it.
  - The model has no tools that act.
- **Privacy:**
  - The backend logs only the request ID, path, status, a user-ID hash, token count, latency and error class.
  - It never logs or stores email bodies or AI output.
  - It has no database.

## 1. Entra ID app registration

The add-in (SPA) and this API share one app registration. Set it up once; the add-in repo points here.

1. In **Entra admin center → App registrations → New registration**:
   - **Name:** `Mail Assistant`
   - **Supported account types:** *Accounts in any organizational directory and personal Microsoft accounts*. This is needed because Outlook.com accounts are supported. The backend still accepts only your tenant plus the consumer tenant; see `ALLOWED_TENANT_IDS`.
2. In **Authentication → Add a platform → Single-page application**, add these redirect URIs:
   - `brk-multihub://localhost:3000`: NAA broker, dev
   - `https://localhost:3000/dialog.html`: Office dialog fallback, dev
   - For production, add the same two URIs with your add-in domain, for example `brk-multihub://addin.example.com` and `https://addin.example.com/dialog.html`. NAA URIs take the origin only, no path.
3. In **API permissions → Microsoft Graph → Delegated**, add `User.Read` and `Mail.ReadWrite`. `Mail.ReadWrite` is needed for `createReply` drafts. Optionally, grant admin consent for your tenant.
4. In **Expose an API**:
   - Set the Application ID URI to `api://<client-id>`.
   - Add the scope `access_as_user`, consentable by admins and users.
   - Under **Authorized client applications**, add the app's own client ID and tick `access_as_user`. This pre-authorizes the add-in.
5. In **Manifest**, set `"requestedAccessTokenVersion": 2`. Personal accounts require it, and the backend checks the v2 issuer.
6. Put the client ID in this repo's `.env` as `ENTRA_CLIENT_ID`, and in the add-in repo's `.env` as `VITE_ENTRA_CLIENT_ID`.
7. Set `ALLOWED_TENANT_IDS=<your-tenant-id>,9188040d-6c67-4c5b-b112-36a304b66dad`. The second ID is the Microsoft consumer tenant, used by Outlook.com accounts.

## 2. Local development

Prerequisites:
- Node 20 or later (22 recommended).
- An OpenAI API key (see section 4).

```bash
cp .env.example .env                       # fill in the values
npm install
npx office-addin-dev-certs install         # trusted localhost cert, shared with the add-in dev server
```

In `.env`, point the backend at that certificate:

```
TLS_CERT_FILE=/home/<you>/.office-addin-dev-certs/localhost.crt
TLS_KEY_FILE=/home/<you>/.office-addin-dev-certs/localhost.key
```

```bash
npm run dev      # https://localhost:8080  (GET /api/health)
```

### Mock mode (no mailbox, no sign-in)

```bash
DEV_AUTH_BYPASS=1 npm run dev    # accepts the bearer token "dev" (never in NODE_ENV=production)
```

Pair it with `npm run mock` in the add-in repo. The AI calls are real.

### Tests and checks

```bash
npm test && npm run typecheck && npm run lint
```

The tests cover the LLM client's JSON parsing and retry, token validation against a mocked JWKS (audience, issuer, tenant, expiry, scope, signature), rate limit and daily budget, zod validation and truncation, prompt delimiters and injection fencing, 401/400/429 over HTTP, and per-item batch failures.

## 3. Deployment

```bash
docker build -t <registry>/mail-assistant-backend:1.0.0 .
docker push <registry>/mail-assistant-backend:1.0.0
```

Run it on any container host with HTTPS ingress. On Azure Container Apps that is `az containerapp create … --ingress external --target-port 8080`. Set these environment variables:
- `LLM_API_KEY`: your OpenAI API key; store it as a secret. Optionally override `LLM_BASE_URL` and `LLM_MODEL`.
- `ENTRA_CLIENT_ID`
- `ALLOWED_TENANT_IDS`
- `ADDIN_ORIGIN`: set it to the add-in's origin, for example `https://addin.example.com`.
- `DAILY_TOKEN_BUDGET_PER_USER`
- `NODE_ENV=production`: the Dockerfile sets it already.

TLS terminates at the ingress.

> shortcut: rate limits and token budgets are in memory, per replica. Run one replica, or move the counters to Redis before scaling out.

## 4. Configuration

See [.env.example](.env.example). Set `LLM_API_KEY` to an OpenAI API key; the defaults use `https://api.openai.com/v1` and `gpt-4o-mini`.

| Variable | Default | Notes |
|---|---|---|
| `LLM_BASE_URL` | `https://api.openai.com/v1` | Override only if you use a proxy or OpenAI-compatible gateway. |
| `LLM_MODEL` | `gpt-4o-mini` | Any chat model that supports JSON mode and streaming. |
| `LLM_API_KEY` | empty | Create one at platform.openai.com/api-keys. |
| `LLM_CONCURRENCY` | `5` | Parallel AI calls per digest batch (the spec allows at most 5). |

If you see "The configured AI model was not found", check that `LLM_MODEL` is available to your OpenAI account. The client retries 429s with backoff.
