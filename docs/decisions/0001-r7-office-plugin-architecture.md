# ADR 0001: R7-Office Plugin & MCP Server Architecture

## Status
Accepted

## Context
DeepSeek Harness needs programmatic, safe, and style-preserving document processing for Russian and international office formats (DOCX, XLSX, PPTX, PDF) integrated with R7-Office (Р7-Офис).

Requirements:
1. Work directly on document files without keyboard/mouse emulation.
2. Preserve existing styles, numbering, headers, and document structures.
3. Integrate with DeepSeek Harness via Cordis plugin mechanism.
4. Expose standard MCP (Model Context Protocol) interface.
5. Utilize local R7-Office engine (`x2t.exe` converter and templates) when available.
6. Enable live interaction with opened R7 Desktop documents via a desktop bridge plugin.
7. Maintain clean open-source repository ready for public release.

## Decision
1. **Repository Layout**:
   - Pure ESM Node.js architecture with zero runtime binaries committed to git.
   - Dynamic host discovery of R7-Office installation and `x2t` converter.
   - Native OOXML package manipulation engine with robust run-preserving text substitution.
2. **Dual-Facing Interface**:
   - Host plugin exports Cordis `apply(ctx)` registering `r7_*` tools on `ctx.tools` and providing agent prompt instructions.
   - MCP server exports standard JSON-RPC 2.0 tools for MCP clients.
3. **Safety Model**:
   - Non-destructive by default: validation before write, automatic rollback/backup option.
   - Verification tool `r7_validate` to ensure document integrity before returning success.

## Consequences
- Fast execution with low memory footprint.
- Runs anywhere Node.js >= 20 is present.
- Enhances rendering fidelity to 100% when local R7-Office is installed.
