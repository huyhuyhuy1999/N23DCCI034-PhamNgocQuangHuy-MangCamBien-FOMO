from __future__ import annotations

import mimetypes
import os
import uuid
from pathlib import Path
from urllib import request, error


ROOT = Path(__file__).resolve().parent
DATASET = ROOT / "synthetic_bottle_dataset"
INGESTION = "https://ingestion.edgeimpulse.com/api/{category}/files"


def files_for_split(split: str) -> list[Path]:
    split_dir = DATASET / split
    labels = split_dir / "bounding_boxes.labels"
    if not labels.exists():
        raise FileNotFoundError(labels)
    return [labels] + sorted(split_dir.glob("*.jpg"))


def multipart_body(paths: list[Path]) -> tuple[bytes, str]:
    boundary = "----ei-dataset-" + uuid.uuid4().hex
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


def upload_split(api_key: str, split: str) -> None:
    paths = files_for_split(split)
    body, content_type = multipart_body(paths)
    url = INGESTION.format(category=split)
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
    try:
        with request.urlopen(req, timeout=240) as resp:
            text = resp.read().decode("utf-8", errors="replace").strip()
            print(f"{split}: HTTP {resp.status}")
            if text:
                print(text[:1000])
    except error.HTTPError as exc:
        text = exc.read().decode("utf-8", errors="replace").strip()
        print(f"{split}: HTTP {exc.code}")
        if text:
            print(text[:1000])
        raise


def main() -> None:
    api_key = os.environ.get("EI_API_KEY")
    if not api_key:
        raise SystemExit("Missing EI_API_KEY environment variable.")
    if not DATASET.exists():
        raise SystemExit(f"Dataset folder not found: {DATASET}")
    upload_split(api_key, "training")
    upload_split(api_key, "testing")


if __name__ == "__main__":
    main()
