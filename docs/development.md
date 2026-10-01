# Development Guide

## Prerequisites

- Node.js **>= 22** for the full test suite (the bridge and CDP clients use the
  global `WebSocket` class, which exists from Node 22). The plugin itself runs
  on Node 20.
- *Optional:* R7-Office Desktop on the host, for native conversion and the
  live desktop bridge tests. Everything skips cleanly without it.

## Setup

```bash
git clone https://github.com/OWNER/dsh-r7-office.git
cd dsh-r7-office
npm install
npm test
```

## Test layout

| Path | What it covers |
|---|---|
| `tests/unit/` | ZIP reader/writer, XML helpers, R7 detection, plugin contract, security policy |
| `tests/integration/` | DOCX/XLSX/PPTX engines, and the OOXML regression suite (Unicode, large documents, malformed packages, byte-level round-trips) |
| `tests/e2e/` | Full file workflows through the MCP tool surface, the desktop bridge wire protocol, and the external MCP client smoke suite |

```bash
npm test                 # everything
npm run test:unit
npm run test:integration
npm run test:e2e
```

Tests that need R7-Office skip themselves with a clear message, so a machine
without it still gets a green run.

## Verification beyond the test suite

```bash
# Live R7-Office Desktop: installs the bridge plugin, drives the real editor
# through the CEF DevTools protocol, edits, saves, re-opens and validates.
npm run test:live

# DeepSeek Harness: boots a fresh Harness and asserts the plugin reached the
# tool registry with arbitrary code execution disabled.
npm run test:harness -- --profile web

# Clean install: clones the committed repository into an empty directory,
# installs, runs the suite and inspects the npm pack contents.
npm run test:clean-install

# Probe a running R7 Desktop instance over CDP.
npm run inspect:r7
```

The live tests launch R7 with `--remote-debugging-port` and drive it as a
**program**, not by emulating mouse or keyboard input. They install the bridge
plugin into the user's R7 plugin directory and start it through the editor's own
`pluginsManager`.

## Running the MCP server by hand

```bash
echo '{"jsonrpc":"2.0","id":1,"method":"tools/list"}' | node src/mcp/cli.js
```

## Installing into DeepSeek Harness

Install the package directory as a bundle (`plugin_manager` → `install_bundle`),
or add the row to the profile's `cordis.patch.yml`:

```yaml
- insert:
    - id: r7-office
      name: 'dsh-r7-office'
```

Do **not** do both: two rows with the same `r7-office` id make the plugin fail to
activate.

## Code layout

```text
src/plugin/        DeepSeek Harness (Cordis) integration
src/mcp/           MCP server, tool definitions, desktop bridge server
src/r7/            R7 detection, DOCX/XLSX/PPTX engines
src/shared/        Dependency-free ZIP and XML helpers
desktop-bridge/    The plugin installed into R7-Office Desktop
scripts/           Verification and diagnostic scripts
```

Design rules worth keeping:

- **Never rewrite what you did not change.** The ZIP layer replays a member's
  original compressed bytes, flags, extra fields and timestamps unless its
  content was replaced. `tests/integration/zip-ooxml-regression.test.js`
  enforces byte-identical round-trips against real R7 templates; keep it green.
- **No runtime dependencies.** The runtime uses only the Node standard library.
  Adding a dependency needs a strong reason.
- **No machine-specific paths, secrets or fixtures.** Every test authors its own
  documents at run time.

## Publishing a release

**Nothing is published automatically.** The release stays a local candidate
until the maintainer gives the explicit go-ahead; every step below is manual.

### Preconditions

`npm test` and `npm run test:clean-install` both green, and the acceptance
artefacts for every format visually approved by the maintainer
(`R7_MCP_DOCX_Acceptance`, `R7_MCP_XLSX_Acceptance`, `R7_MCP_PPTX_Acceptance`
plus their PDFs).

### Publication procedure (run only on the go-ahead)

Target repository: **`blazar-source/dsh-r7-office`**, public.

1. Replace every `OWNER` placeholder with `blazar-source` in the badge and
   release-link URLs:
   - `README.md`
   - `README.ru.md`
   - `CHANGELOG.md`
   - `docs/release-notes/v0.1.0.md`

   Verify none remain:

   ```bash
   grep -rn "OWNER" --include="*.md" .
   ```

2. Confirm the working tree is clean and the tag points at the release commit:

   ```bash
   git status --porcelain          # must print nothing
   git log --oneline -1
   git tag -l 'v0.1.0'
   ```

3. Create the public repository and push the branch and the tag:

   ```bash
   git remote add origin https://github.com/blazar-source/dsh-r7-office.git
   git push -u origin main
   git push origin v0.1.0
   ```

4. Set the repository topics exactly:

   ```text
   dsh-plugin, deepseek-harness, mcp, r7-office, docx, xlsx, pptx, office
   ```

   ```bash
   gh repo edit blazar-source/dsh-r7-office --add-topic dsh-plugin \
     --add-topic deepseek-harness --add-topic mcp --add-topic r7-office \
     --add-topic docx --add-topic xlsx --add-topic pptx --add-topic office
   ```

5. Create the GitHub Release `v0.1.0` from the tag, using
   `docs/release-notes/v0.1.0.md` as the body:

   ```bash
   gh release create v0.1.0 --title "v0.1.0" --notes-file docs/release-notes/v0.1.0.md
   ```

6. Post-publication verification:
   - `gh run list` — the CI workflow must pass on `main` for both Node
     versions and all three operating systems.
   - Clean clone from the **public** URL and run the suite there, which is the
     only check that proves the published artefact is complete:

     ```bash
     git clone https://github.com/blazar-source/dsh-r7-office.git /tmp/r7-published
     cd /tmp/r7-published && npm install && npm test
     ```

   - Confirm the repository page renders the README, the disclaimer, the
     SECURITY policy and the release notes.

### Notes

- R7-Office is proprietary: the repository must never contain its binaries.
  `npm pack` is checked for this by `scripts/clean-install-check.mjs`.
- The project is an unofficial community effort and must keep saying so in both
  READMEs and in the release notes.
