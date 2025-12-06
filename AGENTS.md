# Repository Guidelines

## Project Structure & Module Organization
- Root files: `index.js` (MCP server), `package.json`, `README.md`, `LICENSE`.
- Dependencies are in `node_modules/` (ES modules via `type: module`).
- Screenshots are written to `./screenshots/` by default (created at runtime) or to a custom folder provided by the caller.
- No separate `src/` or test directories; all server logic lives in `index.js`.

## Build, Test, and Development Commands
- `npm install`: Install dependencies.
- `npm start` or `node index.js`: Start the MCP server over stdio.
- Local integration: add the server to your Claude MCP config (see `README.md`), then invoke tools like `list_windows` and `take_screenshot` from your client.

## Coding Style & Naming Conventions
- Language: Node.js (ESM). Use `import`/`export`, 2‑space indent, semicolons, and camelCase for functions/variables.
- Structure: keep tools declared in `ListToolsRequestSchema` and implemented in `CallToolRequestSchema` with clear, user‑facing descriptions.
- Error handling: return friendly text; avoid exposing raw PowerShell output. Follow existing patterns in `index.js`.
- Lint/format: no linters configured; match the current style. If adding one, prefer Prettier with default settings.

## Testing Guidelines
- No automated test suite. Validate manually through an MCP client:
  - `list_windows` (optionally `filter`, `format: "detailed"`).
  - `take_screenshot` (`filename`, `monitor`, `windowTitle`, `processName`, `windowNumber`, `folder`).
- Verify screenshots are created and error messages are clear (e.g., minimized window handling, invalid indices).

## Commit & Pull Request Guidelines
- Commits: short, imperative, scoped (e.g., "Add custom folder support"), optionally note version (e.g., `(v1.2.0)`).
- PRs: include purpose, user impact, sample commands or output, and any Windows/WSL considerations. Link issues and add screenshots of window lists or saved images when helpful.

## Security & Configuration Tips
- Environment: WSL + Windows PowerShell. The server base64‑encodes scripts; still validate and sanitize new inputs.
- Paths: support both WSL (`/mnt/c/...`) and Windows (`C:\...`) paths; avoid writing outside intended locations.
- Agent additions: when adding tools, define a strict `inputSchema`, prefer explicit types, and default-safe behavior.

