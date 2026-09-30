# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.1.0] - 2026-10-01

The first public release. Full notes: [docs/release-notes/v0.1.0.md](docs/release-notes/v0.1.0.md).

### Added

**Tools and runtime**
- 17 MCP tools: `r7_inspect`, `r7_read`, `r7_create`, `r7_edit`, `r7_replace`,
  `r7_insert`, `r7_table`, `r7_sheet_read`, `r7_sheet_write`,
  `r7_sheet_formula`, `r7_slide_create`, `r7_slide_edit`, `r7_convert`,
  `r7_validate`, `r7_desktop_status`, `r7_desktop_selection`,
  `r7_desktop_exec`.
- DeepSeek Harness plugin registering the tools on `ctx.tools` plus an agent
  guidance section on `systemPrompt`.
- Standalone MCP server over stdio JSON-RPC 2.0 (`src/mcp/cli.js`).
- R7-Office auto-detection across the standard install locations of Windows,
  Linux and macOS, with version-agnostic discovery.
- PDF/HTML/TXT conversion through the local R7 `x2t` engine.
- Live R7-Office Desktop bridge over a loopback WebSocket: read the selection,
  modify and save the document the user has open.

**Engines**
- Dependency-free ZIP reader/writer built on `node:zlib`, including CRC-32.
- DOCX, XLSX and PPTX engines with style-preserving editing, tables, ranges,
  formulas and slides.
- Integrity validation for every supported package format.

**Security**
- `r7_desktop_exec` runs only an allowlist of safe, argument-driven commands by
  default; arbitrary DocScript requires an explicit `developerMode: true`.
- Desktop bridge binds to `127.0.0.1` only and reports its effective security
  mode through `r7_desktop_status`.

**Testing and tooling**
- 135 tests across 27 suites: unit, OOXML regression, file end-to-end, desktop
  bridge and external MCP client.
- `scripts/live-desktop-e2e.mjs` — 19-check live run against a real R7-Office
  Desktop driven through the CEF DevTools protocol.
- `scripts/harness-smoke.mjs` — boots a fresh DeepSeek Harness and asserts the
  plugin reached the tool registry.
- `scripts/clean-install-check.mjs` — clones the repository into an empty
  directory, installs, runs the suite and inspects `npm pack` output.
- `scripts/cdp.mjs` — dependency-free Chrome DevTools Protocol client.

**Documentation**
- English and Russian READMEs with clean-install instructions.
- `docs/architecture.md`, `docs/development.md`, `docs/roadmap.md` and
  ADR [0001](docs/decisions/0001-r7-office-plugin-architecture.md).
- `SECURITY.md` with the project threat model.
- Explicit unofficial/community disclaimer: no affiliation with АО «Р7» or
  DeepSeek, and no R7 code or binaries redistributed.

### Changed

- Reading a package and writing it back now reproduces the file byte for byte;
  editing one member leaves every other member byte-identical. This required
  preserving original flag words, extra fields, timestamps and compressed
  streams for untouched members.

### Fixed

- The desktop bridge plugin probes a port range and connects to whichever port
  the bridge actually bound, instead of pinning `7888`. Previously a busy
  default port silently broke the bridge.
- R7 Desktop saves a locally-opened document through
  `editor.asc_Save` → `CDocsSaveApi.saveChanges` → `LocalFileSaveChanges`; the
  bridge now uses that path, because the plugin's `executeMethod('Save')`
  reported success without writing to disk.
- Live desktop teardown terminates `editors.exe`, not only the
  `DesktopEditors.exe` launcher, so the document file handle is released.

### Notes

- Requirements: Node.js >= 20 to run, >= 22 for the full test suite. R7-Office
  Desktop is optional and needed only for conversion and the desktop bridge.
- The automated live desktop test is verified on Windows only.

[0.1.0]: https://github.com/OWNER/dsh-r7-office/releases/tag/v0.1.0
