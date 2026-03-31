# Sentinote Microsoft Add-in

Microsoft Word Online Add-in for ResearchCloud. See workspace `CLAUDE.md` for architecture overview, RBAC model, and API documentation.

## Tech Stack

- **Framework**: React 18 (functional components + hooks)
- **UI**: Fluent UI React v9 (`@fluentui/react-components`)
- **Icons**: `@fluentui/react-icons`
- **Language**: TypeScript 5.x
- **Build**: Webpack 5 + Babel (with React preset)
- **Office API**: `@types/office-js` + Office.js CDN script
- **Auth**: Cognito Hosted UI via Office Dialog API (OAuth authorization code flow)

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
│   ├── App.tsx                    # Root component (auth routing)
│   ├── taskpane.html              # HTML shell (loads office.js)
│   ├── callback.html              # OAuth callback page for Office Dialog API
│   ├── components/
│   │   ├── Login.tsx              # Sign-in screen
│   │   ├── StudyBrowser.tsx       # Study list
│   │   └── FileBrowser.tsx        # File/folder browser with breadcrumbs
│   └── services/
│       ├── auth.ts                # Cognito auth via Office Dialog API
│       ├── api.ts                 # HTTP client with auth headers
│       └── environment.ts         # Environment config (Cognito + API URLs)
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

## Authentication Flow

1. User clicks "Sign In" in taskpane
2. `Office.context.ui.displayDialogAsync()` opens Cognito Hosted UI in a popup
3. User authenticates on the Cognito hosted page
4. Cognito redirects to `callback.html` with an authorization code
5. `callback.html` exchanges the code for tokens via Cognito `/oauth2/token`
6. Tokens are sent to the parent taskpane via `Office.context.ui.messageParent()`
7. Taskpane stores tokens in memory, uses ID token for API calls

The ID token is sent in the `Authorization` header without a "Bearer" prefix, matching the existing frontend convention.

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

Changes to `sentinote-infra/bin/sentinote-infra.ts` are needed for:
- `allowedOrigins`: Add-in dev server and production URLs for CORS
- `callbackUrls`: Cognito OAuth callback URL for the add-in
- `logoutUrls`: Cognito logout redirect URL
