# Security Policy & Offline Integrity Guarantee

## 🛡️ Our Security & Privacy Promise

This project processes sensitive personal and business bank statements. For this reason, we enforce a strict **Offline-First Security Architecture**:

1. **100% Local Execution**:
   - The application does not initiate any outbound or inbound network connections.
   - Zero telemetry, zero analytics, zero external API dependencies.
2. **No Data Storage or Telemetry**:
   - Financial statements are never cached remotely, saved to hidden locations, or transmitted over the wire.
3. **Zero Dynamic Code Execution**:
   - `eval()`, `exec()`, and unsafe deserialization are strictly forbidden.

---

## 🚫 Prohibited Code Rules for Contributors

Any Pull Request that contains any of the following will be **immediately rejected and blocked by automated CI**:

- ❌ Importing any network, HTTP, or socket libraries (`socket`, `requests`, `urllib`, `http.client`, `httpx`, `aiohttp`, etc.).
- ❌ Adding third-party tracking, crash reporting, or telemetry SDKs (`sentry`, `mixpanel`, `google-analytics`, etc.).
- ❌ Dynamic code execution (`eval()`, `exec()`, `__import__`).
- ❌ Unvetted external binaries or pre-compiled C extensions without source code.

---

## 🤖 Automated Security Verification

Every commit and Pull Request is automatically verified by:
- **`test_offline_integrity.py`**: AST-level scanner that checks every Python source file to guarantee 0 network imports.
- **Bandit Security Linter**: Scans for common Python vulnerabilities (CWEs, unsafe subprocesses).
- **GitHub CodeQL**: Deep semantic security analysis for injection vulnerabilities.

---

## 📬 Reporting a Vulnerability

If you discover a potential security flaw or unexpected network behavior:

1. **Do NOT open a public issue.**
2. Please report the vulnerability privately via **[GitHub Security Advisory](https://github.com/MaizaJoel/pichincha-excel-cleaner/security/advisories/new)** or by contacting the maintainer directly.
3. Include:
   - Description of the vulnerability.
   - Steps to reproduce.
   - Potential impact.

We will review, patch, and release a fix promptly.
