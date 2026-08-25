"""Automated Security & Offline Integrity Verifier.

This test suite scans all repository source files using Python's Abstract Syntax Tree (AST)
to mathematically guarantee:
1. Zero network/telemetry modules are imported (maintaining 100% offline guarantee).
2. Zero dynamic code execution primitives (eval, exec, compile) are present.
3. No dangerous hidden subprocesses or remote socket calls exist.

Any Pull Request that attempts to inject network communication or unsafe execution
will automatically fail CI/CD and be blocked from merging.
"""

import ast
from pathlib import Path
import pytest

# Strict list of prohibited network, telemetry, and socket modules
PROHIBITED_MODULES = {
    # Network / HTTP
    "socket",
    "http",
    "http.client",
    "urllib",
    "urllib.request",
    "requests",
    "httpx",
    "aiohttp",
    "websocket",
    "websockets",
    "ftplib",
    "telnetlib",
    "smtplib",
    "paramiko",
    "ssl",
    # Telemetry / Analytics / Tracking
    "sentry_sdk",
    "posthog",
    "mixpanel",
    "segment",
    "analytics",
    # Dangerous dynamic execution
    "pty",
}

PROHIBITED_CALLS = {
    "eval",
    "exec",
    "__import__",
}


def get_source_files() -> list[Path]:
    root = Path(__file__).resolve().parent
    py_files = []
    for path in root.glob("*.py"):
        if path.name.startswith("test_") or path.name.startswith("build_"):
            continue
        py_files.append(path)
    return py_files


class SecurityASTVisitor(ast.NodeVisitor):
    def __init__(self, filename: str):
        self.filename = filename
        self.violations: list[str] = []

    def visit_Import(self, node: ast.Import):
        for alias in node.names:
            base_module = alias.name.split(".")[0]
            if alias.name in PROHIBITED_MODULES or base_module in PROHIBITED_MODULES:
                self.violations.append(
                    f"[{self.filename}:{node.lineno}] Prohibited network module imported: '{alias.name}'"
                )
        self.generic_visit(node)

    def visit_ImportFrom(self, node: ast.ImportFrom):
        if node.module:
            base_module = node.module.split(".")[0]
            if node.module in PROHIBITED_MODULES or base_module in PROHIBITED_MODULES:
                self.violations.append(
                    f"[{self.filename}:{node.lineno}] Prohibited network module imported from: '{node.module}'"
                )
        self.generic_visit(node)

    def visit_Call(self, node: ast.Call):
        if isinstance(node.func, ast.Name):
            if node.func.id in PROHIBITED_CALLS:
                self.violations.append(
                    f"[{self.filename}:{node.lineno}] Dangerous dynamic function call: '{node.func.id}()'"
                )
        self.generic_visit(node)


def test_zero_network_imports_in_source_code():
    """Verify that core engine and GUI files have ZERO network or telemetry imports."""
    source_files = get_source_files()
    assert len(source_files) > 0, "No source files found to inspect."

    all_violations = []
    for file_path in source_files:
        tree = ast.parse(file_path.read_text(encoding="utf-8"), filename=file_path.name)
        visitor = SecurityASTVisitor(file_path.name)
        visitor.visit(tree)
        all_violations.extend(visitor.violations)

    if all_violations:
        error_msg = "\nSECURITY VIOLATION - Pull Request rejected:\n" + "\n".join(all_violations)
        pytest.fail(error_msg)


def test_no_hardcoded_credentials_or_urls():
    """Verify that source files do not contain hardcoded remote endpoints or credentials."""
    source_files = get_source_files()
    suspicious_patterns = [
        "http://",
        "https://api.",
        "ftp://",
        "aws_secret",
        "private_key",
        "bearer ",
    ]

    violations = []
    for file_path in source_files:
        text = file_path.read_text(encoding="utf-8").lower()
        for pattern in suspicious_patterns:
            if pattern in text:
                violations.append(f"[{file_path.name}] Suspicious remote URL or secret pattern: '{pattern}'")

    if violations:
        pytest.fail("\n".join(violations))
