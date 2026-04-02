# Sentinote Microsoft Add-in

Microsoft Word Online Add-in for ResearchCloud. See workspace `CLAUDE.md` for architecture overview, RBAC model, and API documentation.

## Tech Stack

- **Framework**: React 18 (functional components + hooks)
- **UI**: Fluent UI React v9 (`@fluentui/react-components`)
- **Icons**: `@fluentui/react-icons`
- **Language**: TypeScript 5.x
- **Build**: Webpack 5 + Babel (with React preset)
- **Office API**: `@types/office-js` + Office.js CDN script
- **Auth**: Backend SSO via checkout session token (no separate login required)

## Architecture Decision: React over Angular

The main ResearchCloud frontend uses Angular, but this add-in uses React. This was a deliberate choice documented in `docs/microsoft-addin-analysis.md`. Key reasons:
- Microsoft's Office Add-in tooling (Yeoman generator, docs, samples) is React-first
- Fluent UI v9 has native React components; no official Angular bindings exist
- Angular template was removed from the Office generator in 2023
- The add-in is a small, isolated taskpane — no shared code with the Angular frontend

## Project Structure

```
src/
├── taskpane/
│   ├── index.tsx                  # Entry point (Office.onReady + FluentProvider)
│   ├── App.tsx                    # Root component (auto-detects document state)
│   ├── taskpane.html              # HTML shell (loads office.js)
│   ├── ms-callback.html           # OAuth callback for Microsoft Graph (Office Dialog API)
│   ├── components/
│   │   ├── CheckInView.tsx        # Check-in UI (Save to ResearchCloud / Cancel)
│   │   ├── ReadOnlyView.tsx       # View-only message for preview documents
│   │   └── NotResearchCloud.tsx   # Message for non-ResearchCloud documents
│   └── services/
│       ├── auth.ts                # Token storage (set/get ID token from backend session)
│       ├── api.ts                 # HTTP client + addin-session, checkin, cancel endpoints
│       ├── microsoft-auth.ts      # Microsoft Graph OAuth via Office Dialog API (PKCE)
│       ├── onedrive.ts            # OneDrive upload/download/delete via Graph API
│       └── environment.ts         # Environment config (API URLs)
├── commands/
│   ├── commands.ts                # Ribbon command handlers
│   └── commands.html              # Commands HTML shell
assets/                            # Add-in icons (16/32/64/80/128px)
manifest.xml                       # Office Add-in manifest
webpack.config.js                  # Build config
```

## Commands

```bash
npm run dev-server    # Start webpack dev server at https://localhost:3000
npm run build         # Production build
npm run build:dev     # Development build
npm run validate      # Validate manifest.xml
npm start             # Start with Office debugging (sideload)
npm run watch         # Watch mode
```

## Add-in States

The add-in auto-detects the document type on load (no login required):

| State | Detection | UI |
|-------|-----------|-----|
| **Check-in** | OneDrive path contains `{checkoutId}_{filename}` pattern | File details + "Save to ResearchCloud" / "Cancel and Release" buttons |
| **Read-only preview** | OneDrive path contains `.preview/` | "View-only mode" message |
| **Not ResearchCloud** | No matching pattern found | "Not checked out from ResearchCloud" message |

## Authentication Flow (Backend SSO)

No separate login is required. The add-in bootstraps its session from the backend:

1. Angular frontend checks out a file → backend stores Cognito ID token + MS Graph token in DynamoDB checkout record
2. Add-in loads in Word Online → reads document URL from `Office.context.document`
3. Extracts `checkoutId` from the OneDrive filename (pattern: `{checkoutId}_{filename}`)
4. Calls unauthenticated `GET /files/checkout/addin-session?checkoutId=...`
5. Backend returns: file details, Cognito ID token, MS Graph token
6. Add-in uses returned Cognito token for authenticated API calls (check-in, cancel)

For check-in, the add-in reads the document content directly via `Office.context.document.getFileAsync(Office.FileType.Compressed)` — this captures all edits from the Word Online session, bypassing OneDrive caching.

The MS Graph token (from the session) is used for OneDrive cleanup. If expired, the add-in falls back to interactive auth via Office Dialog API with PKCE.

## Environment Configuration

Defined in `src/taskpane/services/environment.ts`:

| Environment | Cognito Pool | API URL |
|-------------|-------------|---------|
| `development` | `eu-central-1_pko7ADLrF` | Personal account API |
| `freshminds` | `eu-central-1_L064ymPKl` | FreshMinds account API |

## Manifest

`manifest.xml` defines the add-in metadata, icons, and allowed domains. Key points:
- `AppDomains` must include all domains the Dialog API navigates to (Cognito auth domains)
- Icons are served from the dev server or production CDN
- The manifest must be re-uploaded in Word Online when changed (not hot-reloaded)

## Testing / Sideloading

1. Run `npm run dev-server`
2. In Word Online: Insert → Add-ins → Upload My Add-in → select `manifest.xml`
3. The add-in appears in the Home ribbon tab

## Coding Conventions

- Functional components with hooks (no class components)
- `makeStyles` from Fluent UI for styling (CSS-in-JS)
- Fluent UI tokens for colors, spacing, typography
- Keep components small — the taskpane has limited screen space
- Handle the case where `Office.context.ui` is undefined (direct browser access)

## Related Infrastructure

- **API Gateway**: `/files/checkout/addin-session` route has **no Cognito authorizer** (unauthenticated — security relies on unguessable checkoutId + 24h TTL)
- **DynamoDB**: Checkout records in Studies table (`CHECKOUT#` and `ADDIN_SESSION#` SK patterns)
- **CORS**: `allowedOrigins` in `sentinote-infra/bin/sentinote-infra.ts` must include the add-in URL
- **Production deployment**: See `docs/microsoft-addin-production-deployment.md` for multi-tenant Azure AD setup, add-in hosting, and security analysis
