from __future__ import annotations

import json
import mimetypes
import time
import uuid
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib import error, request


ROOT = Path(__file__).resolve().parent
DATASET = ROOT / "synthetic_bottle_dataset"
STATUS_PATH = ROOT / "edge_impulse_upload_status.json"
PROJECT_ID = 1021052
HOST = "127.0.0.1"
PORT = 8765


def write_status(**kwargs) -> None:
    payload = {"updated": time.strftime("%Y-%m-%d %H:%M:%S"), **kwargs}
    STATUS_PATH.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")


def http_json(method: str, url: str, api_key: str, body: dict | None = None) -> dict:
    data = None if body is None else json.dumps(body).encode("utf-8")
    headers = {"x-api-key": api_key}
    if data is not None:
        headers["Content-Type"] = "application/json"
    req = request.Request(url, data=data, headers=headers, method=method)
    with request.urlopen(req, timeout=240) as resp:
        raw = resp.read().decode("utf-8", errors="replace")
        try:
            return json.loads(raw) if raw else {"success": True}
        except json.JSONDecodeError:
            return {"success": resp.status < 300, "text": raw}


def files_for_split(split: str) -> list[Path]:
    split_dir = DATASET / split
    labels = split_dir / "bounding_boxes.labels"
    if not labels.exists():
        raise FileNotFoundError(labels)
    return [labels] + sorted(split_dir.glob("*.jpg"))


def multipart_body(paths: list[Path]) -> tuple[bytes, str]:
    boundary = "----ei-upload-" + uuid.uuid4().hex
    chunks: list[bytes] = []
    for path in paths:
        mime = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
        chunks.append(f"--{boundary}\r\n".encode("utf-8"))
        chunks.append(
            (
                f'Content-Disposition: form-data; name="data"; filename="{path.name}"\r\n'
                f"Content-Type: {mime}\r\n\r\n"
            ).encode("utf-8")
        )
        chunks.append(path.read_bytes())
        chunks.append(b"\r\n")
    chunks.append(f"--{boundary}--\r\n".encode("utf-8"))
    return b"".join(chunks), f"multipart/form-data; boundary={boundary}"


def upload_split(api_key: str, split: str) -> str:
    paths = files_for_split(split)
    body, content_type = multipart_body(paths)
    url = f"https://ingestion.edgeimpulse.com/api/{split}/files"
    req = request.Request(
        url,
        data=body,
        headers={
            "x-api-key": api_key,
            "Content-Type": content_type,
            "Content-Length": str(len(body)),
        },
        method="POST",
    )
    with request.urlopen(req, timeout=420) as resp:
        return resp.read().decode("utf-8", errors="replace")[:1600]


def run_upload(api_key: str) -> dict:
    write_status(stage="starting", ok=False)
    delete_url = f"https://studio.edgeimpulse.com/v1/api/{PROJECT_ID}/raw-data/delete-all"
    delete_result = http_json("POST", delete_url, api_key)
    if not delete_result.get("success"):
        raise RuntimeError(f"delete-all failed: {delete_result}")
    write_status(stage="deleted-old-template-data", ok=True)

    training_text = upload_split(api_key, "training")
    write_status(stage="uploaded-training", ok=True, training_response=training_text)

    testing_text = upload_split(api_key, "testing")
    result = {
        "ok": True,
        "stage": "uploaded",
        "training_files": len(files_for_split("training")) - 1,
        "testing_files": len(files_for_split("testing")) - 1,
        "training_response": training_text,
        "testing_response": testing_text,
    }
    write_status(**result)
    return result


class Handler(BaseHTTPRequestHandler):
    def _send(self, code: int, payload: dict) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Access-Control-Allow-Origin", "https://studio.edgeimpulse.com")
        self.send_header("Access-Control-Allow-Headers", "content-type")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Private-Network", "true")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self) -> None:
        self._send(204, {})

    def do_GET(self) -> None:
        if self.path != "/status":
            self._send(404, {"ok": False, "error": "not found"})
            return
        if STATUS_PATH.exists():
            self._send(200, json.loads(STATUS_PATH.read_text(encoding="utf-8")))
        else:
            self._send(200, {"ok": False, "stage": "idle"})

    def do_POST(self) -> None:
        if self.path != "/upload":
            self._send(404, {"ok": False, "error": "not found"})
            return
        length = int(self.headers.get("Content-Length", "0"))
        payload = json.loads(self.rfile.read(length).decode("utf-8"))
        api_key = payload.get("apiKey")
        if not isinstance(api_key, str) or not api_key.startswith("ei_"):
            self._send(400, {"ok": False, "error": "missing api key"})
            return
        try:
            result = run_upload(api_key)
            self._send(200, {k: v for k, v in result.items() if "key" not in k.lower()})
        except error.HTTPError as exc:
            text = exc.read().decode("utf-8", errors="replace")[:1600]
            write_status(stage="failed", ok=False, http_status=exc.code, response=text)
            self._send(500, {"ok": False, "error": f"HTTP {exc.code}", "response": text})
        except Exception as exc:
            write_status(stage="failed", ok=False, error=str(exc))
            self._send(500, {"ok": False, "error": str(exc)})

    def log_message(self, format: str, *args) -> None:
        return


def main() -> None:
    write_status(stage="ready", ok=False)
    server = ThreadingHTTPServer((HOST, PORT), Handler)
    print(f"Edge Impulse upload bridge on http://{HOST}:{PORT}")
    server.serve_forever()


if __name__ == "__main__":
    main()
