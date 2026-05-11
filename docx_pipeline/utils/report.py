"""
utils/report.py â€” Per-file issue logger and HTML report generator.
"""

import os
from datetime import datetime
from pathlib import Path


class ReportLogger:
    """Collects issues and info messages during pipeline processing."""

    def __init__(self, filename: str):
        self.filename  = filename
        self.messages  = []   # list of {"level", "step", "message"}
        self._step     = "init"

    def set_step(self, step: str):
        self._step = step

    def info(self, msg: str):
        self.messages.append({"level": "info",  "step": self._step, "message": msg})

    def warning(self, msg: str):
        self.messages.append({"level": "warning", "step": self._step, "message": msg})

    def error(self, msg: str):
        self.messages.append({"level": "error", "step": self._step, "message": msg})

    def flag(self, msg: str):
        """Flag = editorial review required (not auto-fixed)."""
        self.messages.append({"level": "flag", "step": self._step, "message": msg})

    def get_issues(self) -> list:
        return self.messages

    def has_errors(self) -> bool:
        return any(m["level"] == "error" for m in self.messages)

    def summary(self) -> dict:
        counts = {"info": 0, "warning": 0, "error": 0, "flag": 0}
        for m in self.messages:
            counts[m["level"]] = counts.get(m["level"], 0) + 1
        return counts

    def to_html(self, output_dir: str) -> str:
        """Write a sidecar HTML report and return its path."""
        fname = Path(self.filename).stem + "_report.html"
        out   = os.path.join(output_dir, fname)

        level_colors = {
            "info":    ("d4edda", "155724"),
            "warning": ("fff3cd", "664d03"),
            "error":   ("f8d7da", "721c24"),
            "flag":    ("cfe2ff", "084298"),
        }
        rows = ""
        for m in self.messages:
            bg, fg = level_colors.get(m["level"], ("fff", "000"))
            level_label = m["level"].upper()
            step = m["step"]
            message = m["message"]
            rows += (
                f'<tr>'
                f'<td><span class="badge level-{m["level"]}">{level_label}</span></td>'
                f'<td><code>{step}</code></td>'
                f'<td>{message}</td></tr>\n'
            )

        s = self.summary()
        has_errors = s["error"] > 0
        status_badge = "error" if has_errors else "success"
        status_text = "âš  Processing completed with errors" if has_errors else "âœ“ Processing completed successfully"

        html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Pipeline Report â€” {self.filename}</title>
<style>
  *{{box-sizing:border-box;margin:0;padding:0}}
  body{{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Oxygen,Ubuntu,Cantarell,sans-serif;font-size:14px;background:linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);color:#2c3e50;line-height:1.6;min-height:100vh}}
  .container{{max-width:1000px;margin:0 auto;padding:32px 20px}}

  header{{background:linear-gradient(135deg, #667eea 0%, #764ba2 100%);color:#fff;padding:28px;border-radius:12px;margin-bottom:28px;box-shadow:0 8px 32px rgba(102,126,234,.2)}}
  header h1{{font-size:26px;font-weight:700;margin-bottom:8px;letter-spacing:-.5px}}
  header .filename{{font-size:13px;opacity:.85;word-break:break-all}}
  header .timestamp{{font-size:12px;opacity:.75;margin-top:8px}}

  .status-card{{background:#fff;border-radius:12px;padding:20px;margin-bottom:28px;box-shadow:0 4px 16px rgba(0,0,0,.08);border-left:4px solid #667eea}}
  .status-card.success{{border-left-color:#155724}}
  .status-card.error{{border-left-color:#721c24}}
  .status-text{{font-size:15px;font-weight:600;margin-bottom:16px}}

  .summary-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:12px;margin-top:16px}}
  .summary-item{{background:linear-gradient(135deg, #f9f9ff 0%, #f0f0ff 100%);padding:14px;border-radius:8px;text-align:center;border:1px solid #e0e0ff}}
  .summary-item .count{{font-size:24px;font-weight:700;color:#667eea;margin-bottom:4px}}
  .summary-item .label{{font-size:11px;color:#999;text-transform:uppercase;letter-spacing:.3px}}

  .content{{background:#fff;border-radius:12px;padding:28px;box-shadow:0 4px 16px rgba(0,0,0,.08)}}
  .content h2{{font-size:16px;font-weight:700;margin-bottom:16px;color:#2c3e50;padding-bottom:12px;border-bottom:1px solid #e0e0e0}}

  table{{width:100%;border-collapse:collapse;margin-top:16px}}
  th{{background:linear-gradient(135deg, #667eea 0%, #764ba2 100%);color:#fff;padding:12px 14px;text-align:left;font-weight:600;font-size:12px;text-transform:uppercase;letter-spacing:.3px}}
  td{{padding:12px 14px;border-bottom:1px solid #f0f0f0;vertical-align:top}}
  tr:last-child td{{border-bottom:none}}
  tr:hover td{{background:#f9f9ff}}

  code{{background:#f5f5f5;padding:4px 8px;border-radius:4px;font-family:Menlo,Monaco,monospace;font-size:12px;color:#e74c3c}}

  .badge{{display:inline-block;padding:4px 10px;border-radius:6px;font-weight:600;font-size:11px;text-transform:uppercase;letter-spacing:.3px}}
  .level-info{{background:#d4edda;color:#155724}}
  .level-warning{{background:#fff3cd;color:#664d03}}
  .level-error{{background:#f8d7da;color:#721c24}}
  .level-flag{{background:#cfe2ff;color:#084298}}

  .empty{{text-align:center;padding:32px;color:#999;font-size:13px}}
</style>
</head>
<body>
<div class="container">
  <header>
    <h1>Pipeline Report</h1>
    <div class="filename">{self.filename}</div>
    <div class="timestamp">Generated: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</div>
  </header>

  <div class="status-card {status_badge}">
    <div class="status-text">{status_text}</div>
    <div class="summary-grid">
      <div class="summary-item">
        <div class="count">{s["info"]}</div>
        <div class="label">Info</div>
      </div>
      <div class="summary-item">
        <div class="count">{s["warning"]}</div>
        <div class="label">Warnings</div>
      </div>
      <div class="summary-item">
        <div class="count">{s["error"]}</div>
        <div class="label">Errors</div>
      </div>
      <div class="summary-item">
        <div class="count">{s["flag"]}</div>
        <div class="label">Flags</div>
      </div>
    </div>
  </div>

  <div class="content">
    <h2>Processing Details</h2>
    {f'<table><thead><tr><th>Level</th><th>Step</th><th>Message</th></tr></thead><tbody>{rows}</tbody></table>' if rows else '<div class="empty">No messages recorded</div>'}
  </div>
</div>
</body></html>"""

        with open(out, "w", encoding="utf-8") as f:
            f.write(html)
        return out

