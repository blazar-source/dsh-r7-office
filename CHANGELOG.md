# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.1.0] - 2026-10-01

### Added
- Initial project scaffolding and architecture design.
- Dynamic R7-Office host detection adapter (Windows Registry & standard directories).
- Core OOXML document processing engine (DOCX, XLSX, PPTX).
- Full suite of MCP tools: `r7_inspect`, `r7_read`, `r7_create`, `r7_edit`, `r7_replace`, `r7_insert`, `r7_format`, `r7_table`, `r7_sheet_read`, `r7_sheet_write`, `r7_sheet_formula`, `r7_slide_create`, `r7_slide_edit`, `r7_convert`, `r7_validate`.
- Native format conversion to PDF/HTML via R7 `x2t` converter engine.
- DeepSeek Harness plugin integration with Cordis `ctx.tools` and `systemPrompt`.
- R7 Desktop bridge plugin foundation.
- Comprehensive test suites (unit, integration, and e2e).
