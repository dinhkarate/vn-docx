#!/usr/bin/env node
// Thin wrapper: the engine is a stdlib-only Python script; npm only puts it on PATH.
"use strict";
const { spawnSync } = require("child_process");
const path = require("path");

const script = path.join(__dirname, "..", "skills", "vn-docx", "scripts", "vanban.py");
const candidates = process.env.VN_DOCX_PYTHON
  ? [process.env.VN_DOCX_PYTHON]
  : process.platform === "win32" ? ["py", "python", "python3"] : ["python3", "python"];

for (const py of candidates) {
  const args = py === "py" ? ["-3", script, ...process.argv.slice(2)] : [script, ...process.argv.slice(2)];
  const r = spawnSync(py, args, { stdio: "inherit", env: { ...process.env, PYTHONIOENCODING: "utf-8" } });
  if (r.error && r.error.code === "ENOENT") continue;
  if (r.error) { console.error(r.error.message); process.exit(1); }
  if (r.signal) process.kill(process.pid, r.signal);
  process.exit(r.status === null ? 1 : r.status);
}
console.error("vn-docx cần Python 3.9+ (python3). Cài Python tại https://www.python.org/downloads/ hoặc đặt biến VN_DOCX_PYTHON.");
process.exit(1);
