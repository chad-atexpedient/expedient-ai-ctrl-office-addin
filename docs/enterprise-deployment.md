# Enterprise Deployment Guide

This guide covers deploying CTRL BYOK Office Add-in in a production Microsoft 365 enterprise environment. It consolidates the information from the production launch checklist, M365 centralized deployment steps, and architecture documentation into a single enterprise reference.

> **Prerequisites:** This guide assumes an Azure/Entra ID tenant, Microsoft 365 admin access, a container hosting platform (Azure App Service, AWS ECS, GCP Cloud Run, or equivalent), and familiarity with Office add-in sideloading.

---

## Table of Contents

1. [Architecture Overview](#architecture-overview)
2. [Infrastructure Requirements](#infrastructure-requirements)
3. [Entra App Registration](#entra-app-registration)
4. [Server Configuration](#server-configuration)
5. [Container Build and Hosting](#container-build-and-hosting)
6. [Production Manifest Generation](#production-manifest-generation)
7. [Microsoft 365 Admin Center Deployment](#microsoft-365-admin-center-deployment)
8. [Tenant Policy and Settings Store](#tenant-policy-and-settings-store)
9. [Security Posture](#security-posture)
10. [Monitoring and Observability](#monitoring-and-observability)
11. [Rollout Strategy](#rollout-strategy)
12. [Updating the Add-in](#updating-the-add-in)
13. [Troubleshooting](#troubleshooting)

---

## Architecture Overview

CTRL BYOK Office Add-in is a single hosted web application that serves as a task pane in Excel, Word, and PowerPoint simultaneously. The runtime has three pieces:

```text
Excel / Word / PowerPoint ribbon
  └─> manifest.xml command
      └─> hosted task pane /index.html (React SPA)
          ├─> Office.js host APIs (read/write active document)
          └─> same-origin Node server APIs
               /api/settings       tenant/user-scoped settings
               /api/proxy          BYOK provider proxy (OpenAI + Anthropic)
               /api/web-search     lightweight web search tool
               /api/web-fetch      public URL fetch tool
               /api/image-asset    public image fetch/validation
               /api/generated/*    generated Office file artifacts
               /api/m365/*         Office SSO / Graph delegated context
               /healthz            liveness probe
               /readyz             production readiness probe
```

The add-in does **not** run an AI model locally. It proxies user-supplied or organization-managed API keys to approved OpenAI-compatible or Anthropic-compatible provider endpoints. The BYOK proxy validates targets against an explicit allowlist before forwarding.

### Key design properties

- **Single manifest, three hosts:** One Office XML manifest registers the add-in for Excel (`Workbook`), Word (`Document`), and PowerPoint (`Presentation`).
- **BYOK (Bring Your Own Key):** End users or tenant admins configure the provider endpoint and API key. The server never stores provider keys at rest in production — they live in the authenticated settings store scoped to tenant/user identity.
- **Fail-closed production boundary:** All API routes reject unauthenticated requests when `REQUIRE_AUTH=true`. Placeholder SSO IDs, missing allowlists, and development fallbacks cause the `/readyz` probe to fail.
- **Provider protocol translation:** The adapter layer supports both OpenAI `/chat/completions` and Anthropic `/messages` wire protocols, with native tool calling for both.

---

## Infrastructure Requirements

| Component | Requirement |
|---|---|
| **Runtime** | Node.js 22+ (Alpine container image provided) |
| **TLS** | HTTPS with a trusted certificate (terminate at load balancer or platform) |
| **DNS** | Public hostname, e.g. `ctrl-ai.example.com` |
| **Identity** | Microsoft Entra ID (Azure AD) tenant |
| **Microsoft 365** | Admin Center access for centralized deployment |
| **Container platform** | Any: Azure App Service, AWS ECS/Fargate, GCP Cloud Run, Kubernetes |
| **Secrets vault** | Azure Key Vault, AWS Secrets Manager, or equivalent |
| **Persistent storage** | Database or object store for tenant/user settings (replaces dev JSON files) |

### Network requirements

The hosted service needs outbound HTTPS access to:
- Approved BYOK provider endpoints (e.g. `api.openai.com`, `api.anthropic.com`, internal gateways)
- Microsoft Graph API (`graph.microsoft.com`) for M365 context features
- Microsoft identity endpoints (`login.microsoftonline.com`) for token exchange

Inbound: HTTPS on the configured port from Office clients (desktop and web).

---

## Entra App Registration

Office SSO and Microsoft Graph delegated access require an Entra (Azure AD) app registration.

### Step-by-step

1. **Create the app registration** in the Azure portal under your tenant.
2. **Record the Application (client) ID** — this becomes `MSAL_CLIENT_ID`.
3. **Set the Application ID URI** to match your hosting domain:
   ```
   api://ctrl-ai.example.com/<client-id>
   ```
4. **Add delegated Microsoft Graph permissions** (start minimal):
   ```
   openid
   profile
   User.Read
   Files.Read
   ```
   Add broader scopes (`Files.ReadWrite.All`, `offline_access`) only when a documented product requirement and tenant consent decision justify them.
5. **Create a client secret or certificate** for backend On-Behalf-Of token exchange. Store it in a vault — never in environment files or source control.
6. **Configure supported account types** per your tenant policy (single tenant recommended for enterprise).
7. **Set the redirect URI** if using interactive auth flows (not required for SSO-only).

### Manifest alignment

The manifest `<WebApplicationInfo>` block must reference the same app registration:

```xml
<WebApplicationInfo>
  <Id>{MSAL_CLIENT_ID}</Id>
  <Resource>api://ctrl-ai.example.com/{MSAL_CLIENT_ID}</Resource>
  <Scopes>
    <Scope>openid</Scope>
    <Scope>profile</Scope>
    <Scope>offline_access</Scope>
    <Scope>Files.Read</Scope>
    <Scope>User.Read</Scope>
  </Scopes>
</WebApplicationInfo>
```

The production manifest generator (`tools/make-production-manifest.mjs`) enforces this — it rejects output containing placeholder IDs or `localhost`.

---

## Server Configuration

Start from `deploy/production/.env.example`. The complete reference:

```dotenv
# Core server
PORT=3000
HOST=0.0.0.0

# Outbound allowlists (comma-separated exact origins)
BYOK_ALLOWED_TARGETS=https://api.openai.com,https://api.anthropic.com,https://your-gateway.example.com
IMAGE_ALLOWED_TARGETS=https://images.unsplash.com,https://upload.wikimedia.org
WEB_FETCH_ALLOWED_TARGETS=https://www.microsoft.com,https://learn.microsoft.com

# Add-in identity
ADDIN_ORIGIN=https://ctrl-ai.example.com

# Authentication (required for production)
REQUIRE_AUTH=true
AUTH_VALIDATOR_MODULE=./server/auth.mjs

# Pluggable production modules (replace fail-closed templates)
TENANT_POLICY_MODULE=./server/tenant-policy.production.mjs
SETTINGS_STORE_MODULE=./server/settings-store.production.mjs
RATE_LIMITER_MODULE=./server/rate-limiter.production.mjs

# Entra / Office SSO
MSAL_CLIENT_ID=<real-client-id>
MSAL_TENANT_ID=<tenant-id>
MSAL_CLIENT_SECRET=<vault-backed-secret>
OFFICE_SSO_RESOURCE=api://ctrl-ai.example.com/<real-client-id>
GRAPH_SCOPES=openid profile User.Read Files.Read
AUTHORIZED_TENANTS=<comma-separated-tenant-ids>

# Request/response limits
REQUEST_MAX_BYTES=2000000
PROXY_MAX_BYTES=2000000
PROXY_MAX_RESPONSE_BYTES=10000000
OUTBOUND_TIMEOUT_MS=15000
RATE_LIMIT_PER_MINUTE=120
RATE_LIMIT_EXPENSIVE_PER_MINUTE=30

# Observability
AUDIT_LOG_STDOUT=false
```

### Critical production rules

- **`AUTHORIZED_TENANTS`** must list specific tenant IDs. If unset, blank, or set to `common`/`organizations`/`*`, every request is rejected.
- **`GRAPH_ACCESS_TOKEN`** and local token-cache paths must **not** be set. The `/readyz` probe rejects their presence.
- The checked-in `*.production.example.mjs` files are fail-closed contract templates. They reject every request by default. Replace them with real implementations connected to your policy service, database, and distributed rate limiter.

---

## Container Build and Hosting

### Build the container

From the project root:

```bash
docker build -f deploy/production/Dockerfile -t ctrl-byok-office-addin:1.0.0 .
```

The multi-stage Dockerfile:
1. **Build stage:** Installs dependencies, runs the Vite build to produce static task-pane assets.
2. **Runtime stage:** Copies only the server code, built assets, and `package.json`. Runs as a non-root `app` user.

### Run locally in production mode

```bash
docker run --rm -p 3000:3000 \
  --env-file deploy/production/.env \
  ctrl-byok-office-addin:1.0.0
```

### Deploy to a container platform

**Azure App Service (recommended for Microsoft 365 environments):**
```bash
az webapp create --resource-group myRG --plan myPlan \
  --name ctrl-ai-addin --deployment-container-image ctrl-byok-office-addin:1.0.0
az webapp config appsettings set --resource-group myRG --name ctrl-ai-addin \
  --settings @deploy/production/.env
```

**AWS ECS / Fargate:**
- Push the image to ECR
- Create a task definition referencing the image
- Configure secrets from AWS Secrets Manager
- Front with an ALB terminating TLS

**GCP Cloud Run:**
```bash
gcloud run deploy ctrl-ai-addin \
  --image gcr.io/PROJECT/ctrl-byok-office-addin:1.0.0 \
  --port 3000 --allow-unauthenticated
```

### Health checks

Configure your platform's health checks against:
- **Liveness:** `GET /healthz` — returns 200 if the server process is alive.
- **Readiness:** `GET /readyz` — returns 200 only when all production configuration is valid. Fails if placeholder IDs, missing allowlists, or development fallbacks are detected.

### Smoke test

After deployment:

```bash
node deploy/production/scripts/smoke-test.mjs https://ctrl-ai.example.com
```

This verifies the task pane HTML loads, required icon assets are served, critical API routes respond, and the readiness probe passes.

---

## Production Manifest Generation

The development `manifest.xml` uses `localhost` URLs and placeholder SSO IDs. Production requires a generated manifest:

```bash
export MSAL_CLIENT_ID="<real-client-id>"
export OFFICE_SSO_RESOURCE="api://ctrl-ai.example.com/<real-client-id>"

node tools/make-production-manifest.mjs https://ctrl-ai.example.com
```

Output: `dist/manifest.production.xml`

The generator:
- Replaces all `localhost:3000` URLs with the production origin
- Substitutes the placeholder SSO app ID with the real `MSAL_CLIENT_ID`
- **Rejects** output that still contains `localhost` or `00000000-0000-0000-0000-000000000000`
- **Rejects** runs where `MSAL_CLIENT_ID` or `OFFICE_SSO_RESOURCE` are missing or placeholder

### Release packaging

For a complete handoff package:

```bash
pnpm package:prod
```

Creates `dist/release/ctrl-byok-office-addin/` containing the built app, server, production manifest (as `manifest.xml`), deployment docs, environment template, Dockerfile, smoke test, and `RELEASE_REPORT.md` with checksums.

---

## Microsoft 365 Admin Center Deployment

Once the hosted service is reachable over HTTPS and the production manifest is generated and validated:

1. Open the **Microsoft 365 Admin Center**.
2. Navigate to **Settings → Integrated apps**.
3. Choose **Upload custom apps**.
4. Select **Office Add-in**.
5. Upload `dist/manifest.production.xml`.
6. Assign users or groups for pilot rollout (see [Rollout Strategy](#rollout-strategy)).
7. Finish the deployment wizard.
8. Ask pilot users to fully restart Excel, Word, and PowerPoint.
9. Confirm the **CTRL AI** button appears on the Home ribbon in each app.

### Pilot validation checklist

For each Office host (Excel, Word, PowerPoint):

- [ ] Open a normal file and click **CTRL AI** from the ribbon.
- [ ] Confirm the task pane loads from the hosted domain (check DevTools network tab for the correct origin).
- [ ] Configure BYOK provider settings and test the connection.
- [ ] Confirm Office SSO / M365 file context works for a signed-in pilot user, or that the app returns clear setup guidance if SSO is intentionally disabled.
- [ ] Read current document context into a prompt.
- [ ] Insert an assistant response back into the document.

---

## Tenant Policy and Settings Store

Production requires three pluggable modules beyond the fail-closed examples shipped in the repo:

| Module | Env var | Contract |
|---|---|---|
| Auth validator | `AUTH_VALIDATOR_MODULE` | Verifies bearer tokens: issuer, audience, tenant, subject, expiry, signing keys |
| Tenant policy | `TENANT_POLICY_MODULE` | `allows({ identity, capability })` — admin-backed authorization decisions |
| Settings store | `SETTINGS_STORE_MODULE` | Tenant/user-scoped `get`, `put`, `delete` for BYOK settings and branding |
| Rate limiter | `RATE_LIMITER_MODULE` | Distributed rate limiting (the shipped example is in-memory/single-instance only) |

Reference contract shapes are documented inline in `server/*.production.example.mjs`. Do not ship those example files as-is to production — they intentionally reject every request.

### Data residency

- Bind all settings, uploaded logos, generated artifacts, and M365 state to authenticated tenant/user identity.
- Do not rely on the development global JSON settings file or local token cache in production.
- Store BYOK API keys either user-entered per session (never persisted server-side) or in a vault-backed managed-credential mode if your organization centrally manages provider keys.

---

## Security Posture

- **Fail-closed by default.** `REQUIRE_AUTH`, tenant authorization, and outbound allowlists all fail closed rather than open when misconfigured.
- **Exact-origin outbound allowlists.** `BYOK_ALLOWED_TARGETS`, `IMAGE_ALLOWED_TARGETS`, `WEB_FETCH_ALLOWED_TARGETS` — no wildcard domains.
- **No credential forwarding to unapproved hosts.** Outbound requests reject credentials passthrough and private/local network destinations (SSRF protection), disable redirects, and enforce timeouts.
- **Byte caps.** `REQUEST_MAX_BYTES`, `PROXY_MAX_BYTES`, `PROXY_MAX_RESPONSE_BYTES` bound request/response sizes.
- **Rate limiting.** Per-minute limits for standard and "expensive" (e.g. long-context or generation) routes.
- **Non-root container.** The runtime image drops to a dedicated `app` user.
- **CI security gates** (see `.github/workflows/ci.yml`):
  - `gitleaks` secret scanning on every push/PR
  - `trivy` container vulnerability scan (fails on CRITICAL/HIGH with available fixes)
  - SBOM generation (SPDX format) uploaded as a build artifact
  - `security:release-gates` script rejects placeholder production packaging
- **No local AI execution.** The service only proxies to approved external providers; it does not run models locally, reducing the on-host attack surface.

---

## Monitoring and Observability

- `GET /healthz` — liveness; wire to your platform's restart policy.
- `GET /readyz` — readiness; wire to load-balancer target health so misconfigured instances never receive traffic.
- `AUDIT_LOG_STDOUT` — when enabled, emits privacy-safe audit events (provider route, host, outcome — not prompt/response content) to stdout for collection by your log pipeline.
- Recommended: ship container stdout/stderr to your centralized logging (Azure Monitor, CloudWatch, Datadog, etc.) and alert on sustained `/readyz` failures or elevated 401/403 rates.

---

## Rollout Strategy

Recommended staged rollout using Microsoft 365 Admin Center group assignment:

1. **IT/admin pilot** — validate manifest trust, domain access, and SSO token exchange with a handful of IT accounts.
2. **Power-user pilot** — validate BYOK provider settings and Office context actions with a small group of engaged users across Excel, Word, and PowerPoint.
3. **Department rollout** — expand group assignment to one department; monitor support tickets and provider proxy logs.
4. **Broad rollout** — expand assignment tenant-wide after the pilot phases show stable error rates and support volume.

At each stage, watch:
- `/readyz` and `/healthz` trends
- Rate-limit rejection counts (may indicate limits are too tight for real usage)
- Provider proxy error rates (upstream outages vs. misconfiguration)
- Office SSO failure rates

---

## Updating the Add-in

**For UI/server code changes only** (manifest unchanged):
1. Rebuild the container image and redeploy.
2. Users get the new task pane automatically on next reload/reopen — no manifest re-upload needed.

**For manifest changes** (URLs, icons, permissions, add-in ID/version):
1. Rebuild and regenerate `dist/manifest.production.xml`.
2. Re-upload via Microsoft 365 Admin Center → Integrated apps.
3. For ribbon icon changes specifically, bump the manifest `<Version>` and use new icon filenames — Office aggressively caches command surface assets.

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| `/readyz` returns non-200 | Missing/placeholder Entra config, unset `AUTHORIZED_TENANTS`, or dev fallback env vars present | Review the readiness response body; it lists the specific failing checks |
| Ribbon button missing after deployment | Office command cache | Have users run `clear-office-addin-cache.cmd`-equivalent cache clear, or bump manifest version + restart Office |
| SSO silently falls back to setup guidance | `MSAL_CLIENT_ID`/`OFFICE_SSO_RESOURCE` unset or still placeholder | Confirm both are set to real Entra values and match the uploaded manifest |
| 401 on all API calls | `REQUIRE_AUTH=true` but `AUTH_VALIDATOR_MODULE` misconfigured, or client not sending a valid bearer token | Check validator logs; confirm Office SSO token acquisition succeeded client-side |
| BYOK proxy call rejected | Target not in `BYOK_ALLOWED_TARGETS` | Add the exact provider origin to the allowlist and redeploy |
| Container fails Trivy scan in CI | New CRITICAL/HIGH vulnerability in a base image or dependency | Update `node:22-alpine` base image or the flagged dependency; re-run CI |

---

## Related Documentation

- [`docs/production-launch-checklist.md`](./production-launch-checklist.md) — granular pre-launch checklist
- [`docs/architecture.md`](./architecture.md) — full runtime and file-structure reference
- [`docs/office-api-hardening.md`](./office-api-hardening.md) — Office API and M365 context hardening notes
- [`deploy/production/README.md`](../deploy/production/README.md) — production build/package commands
- [`deploy/production/m365-centralized-deployment.md`](../deploy/production/m365-centralized-deployment.md) — condensed Admin Center steps
