#!/usr/bin/env python3
"""
Servidor local do dashboard com endpoint para atualizar a base JSON
a partir da planilha Excel original.
"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse


PROJECT_DIR = Path(__file__).resolve().parents[1]
DASH_DIR = PROJECT_DIR / "dashboard"
BASE_FILE = PROJECT_DIR / "Base_Painel_KPI_.xlsx"
EXTRACTOR = PROJECT_DIR / "scripts" / "extract_kpi_data.py"
DATA_FILE = DASH_DIR / "data" / "kpi_data.json"


def run_extractor() -> tuple[bool, str]:
    cmd = [
        sys.executable,
        str(EXTRACTOR),
        "--input",
        str(BASE_FILE),
        "--output",
        str(DATA_FILE),
    ]
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        msg = (result.stderr or result.stdout or "Falha ao atualizar base.").strip()
        return False, msg
    return True, (result.stdout or "Base atualizada com sucesso.").strip()


class DashboardHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(DASH_DIR), **kwargs)

    def _send_json(self, status: int, payload: dict) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_POST(self) -> None:  # noqa: N802
        path = urlparse(self.path).path
        if path != "/api/refresh":
            self._send_json(HTTPStatus.NOT_FOUND, {"ok": False, "message": "Rota não encontrada."})
            return

        ok, message = run_extractor()
        if not ok:
            self._send_json(HTTPStatus.INTERNAL_SERVER_ERROR, {"ok": False, "message": message})
            return

        generated_at = ""
        try:
            payload = json.loads(DATA_FILE.read_text(encoding="utf-8"))
            generated_at = str(payload.get("generatedAt", ""))
        except Exception:
            generated_at = ""

        self._send_json(
            HTTPStatus.OK,
            {
                "ok": True,
                "message": message,
                "generatedAt": generated_at,
            },
        )


def main() -> None:
    parser = argparse.ArgumentParser(description="Servidor local do dashboard KPI com API de refresh")
    parser.add_argument("--port", type=int, default=5500, help="Porta do servidor")
    args = parser.parse_args()

    server = ThreadingHTTPServer(("0.0.0.0", args.port), DashboardHandler)
    print(f"Servidor KPI em http://localhost:{args.port}")
    print("Endpoint de atualização: POST /api/refresh")
    server.serve_forever()


if __name__ == "__main__":
    main()

