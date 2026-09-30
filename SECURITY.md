# Security Policy

## Supported versions

| Version | Supported |
|---------|-----------|
| 0.1.x   | Yes       |

## Reporting a vulnerability

Please report security issues **privately** through GitHub's
[Security Advisories](https://docs.github.com/en/code-security/security-advisories/guidance-on-reporting-and-writing-information-about-vulnerabilities/privately-reporting-a-security-vulnerability)
for this repository ("Report a vulnerability" on the *Security* tab).

Do not open a public issue for a vulnerability. Include, where possible:

- the affected version and platform,
- a minimal reproduction,
- the impact you believe it has,
- any suggested remediation.

Expect an initial acknowledgement within a few days. This is a community
project maintained on a best-effort basis; there is no commercial support
commitment.

## Threat model

`dsh-r7-office` runs as a local plugin inside DeepSeek Harness and as a local
MCP server. It reads and writes office documents on the local filesystem and,
optionally, talks to a locally installed R7-Office Desktop over loopback. The
relevant threats and the current mitigations are:

### 1. Arbitrary script execution in the open desktop document

`r7_desktop_exec` can ask R7-Office Desktop to run a DocScript/DocumentBuilder
program inside the user's open document. That is effectively arbitrary code
execution with the editor's privileges.

**Mitigation.** Arbitrary code is **disabled by default**:

- Two explicit paths exist: a fixed allowlist of safe, argument-driven
  commands (`safeCommand`) and a raw-code path (`code`).
- The raw-code path is refused unless the operator explicitly sets
  `developerMode: true` in the plugin configuration, or exports
  `DSH_R7_DEVELOPER_MODE=1`.
- The check happens in `DesktopBridge.execute()` before any request reaches
  the editor, and `r7_desktop_status` reports the effective mode.
- The bridge binds to `127.0.0.1` only.

**Residual risk.** With `developerMode` on, a caller that can reach the plugin
can run arbitrary editor scripts. Leave it off unless you are actively
developing plugins.

### 2. Loopback bridge exposure

The desktop bridge listens on a local TCP port (default `7888`, and it walks
upward if the port is taken).

**Mitigation.** It binds `127.0.0.1`, never `0.0.0.0`, and it is not reachable
from other hosts.

**Residual risk.** Any local process on the machine can connect to the port
and issue commands. This is the same trust level as the local editor itself.
Change `desktopBridgePort` if the default conflicts with another tool.

### 3. Untrusted document input

Office packages are ZIP containers with XML inside. A malicious `.docx` can be
a zip bomb, a path-traversal archive, or malformed XML.

**Mitigation.** The bundled ZIP reader validates each entry against the central
directory, rejects wrapped/truncated payloads, refuses unsupported compression
methods, encryption, Zip64 and multi-volume archives, and surfaces inflate
failures as errors instead of crashing. Entry names are never used as output
paths, so archive entries cannot write outside the target document.

### 4. Credentials and data handling

**Mitigation.** The project stores no credentials and ships no telemetry. It
never uploads document content. All processing is local.

**Residual risk.** Documents you process may contain sensitive data; the plugin
writes modified copies where you ask it to. Use the `outputPath` argument to
avoid overwriting originals.

## Scope

Out of scope for this project:

- vulnerabilities in R7-Office itself — report those to АО «Р7»,
- vulnerabilities in DeepSeek Harness — report those to its maintainers,
- issues that require `developerMode: true` to be enabled and then describe the
  arbitrary-code path as a vulnerability (that path is documented and opt-in).
