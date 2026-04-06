
from __future__ import annotations

import argparse
import base64
import contextlib
import hashlib
import html as htmlmod
import json
import os
import re
import shutil
import socket
import sys
import time
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any
from urllib.parse import unquote, urljoin, urlparse

from selenium import webdriver
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.chrome.options import Options


_CT_EXT = {
    "text/html": ".html",
    "text/css": ".css",
    "application/javascript": ".js",
    "text/javascript": ".js",
    "application/json": ".json",
    "application/manifest+json": ".webmanifest",
    "image/png": ".png",
    "image/jpeg": ".jpg",
    "image/webp": ".webp",
    "image/gif": ".gif",
    "image/svg+xml": ".svg",
    "font/woff2": ".woff2",
    "font/woff": ".woff",
    "application/font-woff": ".woff",
    "application/octet-stream": "",
}


@dataclass
class CaptureConfig:
    url: str
    out_dir: str = "site_dump"
    bundle_zip_name: str | None = None
    runner_python_name: str = "launch_site.py"
    runner_bat_name: str = "run_site.bat"
    runner_sh_name: str = "run_site.sh"
    max_wait_secs: int = 60
    network_idle_secs: float = 2.0
    extra_record_secs: float = 15.0
    disable_cache: bool = True
    browser_binary: str | None = None
    headless: bool | None = None
    open_browser_after_capture: bool = False
    zip_output: bool = True
    capture_screenshot: bool = True
    capture_storage: bool = True
    capture_console: bool = True
    capture_same_origin_iframes: bool = True


@dataclass
class CaptureWarning:
    stage: str
    message: str
    url: str | None = None
    request_id: str | None = None


@dataclass
class CaptureResult:
    root_url: str
    final_url: str
    out_dir: str
    bundle_zip: str | None
    offline_html: str
    final_html: str
    saved_count: int
    missing_refs_count: int
    warnings_count: int
    runner: str


def load_config_file(config_path: str | Path) -> CaptureConfig:
    path = Path(config_path).resolve()
    if not path.exists():
        raise FileNotFoundError(f"No existe el archivo de configuración: {path}")

    payload = json.loads(path.read_text(encoding="utf-8"))

    if not isinstance(payload, dict):
        raise ValueError("La configuración debe ser un objeto JSON.")

    if not payload.get("url"):
        raise ValueError("La configuración debe incluir 'url'.")

    return CaptureConfig(**payload)


def _json_dump(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def _jsonl_dump(path: Path, rows: list[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as fh:
        for row in rows:
            fh.write(json.dumps(row, ensure_ascii=False) + "\n")


def _normalize_url(url: str) -> str:
    try:
        parsed = urlparse(url)
        parsed = parsed._replace(fragment="")
        return parsed.geturl()
    except Exception:
        return url


def _guess_ext(mime: str | None) -> str:
    if not mime:
        return ""
    return _CT_EXT.get(mime.split(";")[0].strip().lower(), "")


def _safe_relpath(url: str, mime: str | None = None) -> Path:
    parsed = urlparse(url)
    host = parsed.netloc or "nohost"

    raw_path = unquote(parsed.path or "/")
    if raw_path.endswith("/"):
        raw_path += "index"

    safe_path = raw_path.lstrip("/").replace("..", "_")
    safe_path = re.sub(r'[<>:"|?*]', "_", safe_path)
    rel = Path(host) / safe_path

    if not rel.suffix:
        ext = _guess_ext(mime)
        if ext:
            rel = rel.with_suffix(ext)

    if parsed.query:
        qhash = hashlib.sha256(parsed.query.encode("utf-8")).hexdigest()[:12]
        rel = rel.with_name(f"{rel.stem}__q{qhash}{rel.suffix}")

    return rel


def _detect_runtime() -> dict[str, Any]:
    is_github_actions = os.getenv("GITHUB_ACTIONS", "").strip().lower() == "true"
    is_ci = is_github_actions or os.getenv("CI", "").strip().lower() == "true"
    is_codespaces = os.getenv("CODESPACES", "").strip().lower() == "true"
    no_display = sys.platform.startswith("linux") and not os.getenv("DISPLAY")
    return {
        "platform": sys.platform,
        "github_actions": is_github_actions,
        "ci": is_ci,
        "codespaces": is_codespaces,
        "no_display": no_display,
    }


def _should_use_headless(config: CaptureConfig) -> bool:
    if config.headless is not None:
        return config.headless
    runtime = _detect_runtime()
    return bool(runtime["ci"] or runtime["no_display"])


def _find_browser_binary(preferred: str | None = None) -> str | None:
    candidates: list[str] = []

    if preferred:
        candidates.append(preferred)

    env_candidate = os.getenv("CHROME_BINARY")
    if env_candidate:
        candidates.append(env_candidate)

    if sys.platform.startswith("win"):
        candidates.extend([
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files\Chromium\Application\chrome.exe",
            r"C:\Program Files (x86)\Chromium\Application\chrome.exe",
        ])
    elif sys.platform == "darwin":
        candidates.extend([
            "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
            "/Applications/Chromium.app/Contents/MacOS/Chromium",
        ])
    else:
        candidates.extend([
            "google-chrome",
            "google-chrome-stable",
            "chromium",
            "chromium-browser",
            "chrome",
        ])

    for candidate in candidates:
        if not candidate:
            continue
        expanded = os.path.expandvars(os.path.expanduser(candidate))
        if Path(expanded).exists():
            return expanded
        resolved = shutil.which(expanded)
        if resolved:
            return resolved

    return None


def _build_driver(config: CaptureConfig) -> webdriver.Chrome:
    browser_binary = _find_browser_binary(config.browser_binary)
    headless = _should_use_headless(config)

    options = Options()
    if browser_binary:
        options.binary_location = browser_binary

    options.page_load_strategy = "normal"
    options.set_capability("goog:loggingPrefs", {"performance": "ALL", "browser": "ALL"})

    if headless:
        options.add_argument("--headless=new")

    options.add_argument("--window-size=1600,2200")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-infobars")
    options.add_argument("--no-first-run")
    options.add_argument("--no-default-browser-check")
    options.add_argument("--disable-extensions")
    options.add_argument("--allow-running-insecure-content")
    options.add_argument("--ignore-certificate-errors")

    runtime = _detect_runtime()
    if runtime["platform"].startswith("linux"):
        options.add_argument("--disable-dev-shm-usage")
        if runtime["ci"] or runtime["no_display"]:
            options.add_argument("--no-sandbox")

    try:
        return webdriver.Chrome(options=options)
    except WebDriverException as exc:
        raise RuntimeError(
            "No se pudo iniciar Chrome/Chromium. "
            "Si estás en un runner o servidor, instala Chrome/Chromium y vuelve a intentar. "
            "También puedes pasar browser_binary en la config o definir CHROME_BINARY."
        ) from exc


def _drain_logs(
    driver: webdriver.Chrome,
    perf_logs: list[dict[str, Any]],
    browser_logs: list[dict[str, Any]],
) -> None:
    with contextlib.suppress(Exception):
        perf_entries = driver.get_log("performance")
        if perf_entries:
            perf_logs.extend(perf_entries)

    with contextlib.suppress(Exception):
        browser_entries = driver.get_log("browser")
        if browser_entries:
            browser_logs.extend(browser_entries)


def _wait_for_network_stable(
    driver: webdriver.Chrome,
    config: CaptureConfig,
    perf_logs: list[dict[str, Any]],
    browser_logs: list[dict[str, Any]],
) -> None:
    last_network_event = time.time()
    started = time.time()

    while True:
        _drain_logs(driver, perf_logs, browser_logs)

        for entry in perf_logs[-800:]:
            try:
                message = json.loads(entry["message"]).get("message", {})
                method = message.get("method", "")
                if method.startswith("Network."):
                    last_network_event = time.time()
            except Exception:
                continue

        try:
            ready_state = driver.execute_script("return document.readyState")
        except Exception:
            ready_state = "loading"

        now = time.time()
        if ready_state == "complete" and (now - last_network_event) >= config.network_idle_secs:
            break
        if (now - started) >= config.max_wait_secs:
            break
        time.sleep(0.25)

    extra_started = time.time()
    while (time.time() - extra_started) < float(config.extra_record_secs):
        _drain_logs(driver, perf_logs, browser_logs)
        time.sleep(0.25)


def _parse_perf_logs(perf_logs: list[dict[str, Any]]) -> list[dict[str, Any]]:
    records: dict[str, dict[str, Any]] = {}
    order: list[str] = []

    def ensure_record(request_id: str) -> dict[str, Any]:
        if request_id not in records:
            records[request_id] = {
                "request_id": request_id,
                "request": {},
                "response": {},
                "redirects": [],
                "loading": {},
                "failures": [],
            }
            order.append(request_id)
        return records[request_id]

    for entry in perf_logs:
        try:
            message = json.loads(entry["message"]).get("message", {})
            method = message.get("method", "")
            params = message.get("params", {})
        except Exception:
            continue

        if method == "Network.requestWillBeSent":
            request_id = params.get("requestId")
            request = params.get("request", {})
            if not request_id:
                continue

            record = ensure_record(request_id)
            redirect = params.get("redirectResponse")
            if redirect:
                record["redirects"].append({
                    "status": redirect.get("status"),
                    "url": _normalize_url(redirect.get("url", "")),
                    "headers": redirect.get("headers", {}),
                    "mime_type": redirect.get("mimeType"),
                })

            record["request"] = {
                "url": _normalize_url(request.get("url", "")),
                "method": request.get("method"),
                "headers": request.get("headers", {}),
                "post_data": request.get("postData"),
                "initial_priority": params.get("initialPriority"),
                "resource_type": params.get("type"),
                "document_url": params.get("documentURL"),
                "timestamp": params.get("timestamp"),
                "wall_time": params.get("wallTime"),
                "initiator": params.get("initiator"),
            }

        elif method == "Network.responseReceived":
            request_id = params.get("requestId")
            response = params.get("response", {})
            if not request_id:
                continue

            record = ensure_record(request_id)
            record["response"] = {
                "url": _normalize_url(response.get("url", "")),
                "status": response.get("status"),
                "status_text": response.get("statusText"),
                "mime_type": response.get("mimeType"),
                "headers": response.get("headers", {}),
                "protocol": response.get("protocol"),
                "remote_ip": response.get("remoteIPAddress"),
                "from_disk_cache": response.get("fromDiskCache"),
                "from_service_worker": response.get("fromServiceWorker"),
                "resource_type": params.get("type"),
                "timestamp": params.get("timestamp"),
            }

        elif method == "Network.loadingFinished":
            request_id = params.get("requestId")
            if not request_id:
                continue

            record = ensure_record(request_id)
            record["loading"] = {
                "finished": True,
                "encoded_data_length": params.get("encodedDataLength"),
                "timestamp": params.get("timestamp"),
            }

        elif method == "Network.loadingFailed":
            request_id = params.get("requestId")
            if not request_id:
                continue

            record = ensure_record(request_id)
            record["failures"].append({
                "error_text": params.get("errorText"),
                "canceled": params.get("canceled"),
                "blocked_reason": params.get("blockedReason"),
                "timestamp": params.get("timestamp"),
            })

    return [records[rid] for rid in order]


def _capture_storage(driver: webdriver.Chrome) -> dict[str, Any]:
    script = """
    function dumpStorage(store) {
      const out = {};
      for (let i = 0; i < store.length; i++) {
        const key = store.key(i);
        out[key] = store.getItem(key);
      }
      return out;
    }
    return {
      localStorage: dumpStorage(window.localStorage),
      sessionStorage: dumpStorage(window.sessionStorage),
      title: document.title,
      readyState: document.readyState,
      location: window.location.href
    };
    """
    with contextlib.suppress(Exception):
        return driver.execute_script(script)
    return {}


def _capture_same_origin_iframes(driver: webdriver.Chrome) -> list[dict[str, Any]]:
    script = """
    const result = [];
    const frames = Array.from(document.querySelectorAll("iframe"));
    for (let i = 0; i < frames.length; i++) {
      const frame = frames[i];
      const row = {
        index: i,
        src: frame.getAttribute("src"),
        name: frame.getAttribute("name"),
        title: frame.getAttribute("title"),
      };
      try {
        row.url = frame.contentWindow.location.href;
        row.html = frame.contentDocument.documentElement.outerHTML;
        row.sameOrigin = true;
      } catch (e) {
        row.sameOrigin = false;
        row.error = String(e);
      }
      result.push(row);
    }
    return result;
    """
    with contextlib.suppress(Exception):
        return driver.execute_script(script)
    return []


def _capture_cookies(driver: webdriver.Chrome) -> list[dict[str, Any]]:
    with contextlib.suppress(Exception):
        return driver.get_cookies()
    return []


def _get_all_cookies_cdp(driver: webdriver.Chrome) -> list[dict[str, Any]]:
    with contextlib.suppress(Exception):
        payload = driver.execute_cdp_cmd("Network.getAllCookies", {})
        return payload.get("cookies", [])
    return []


def _save_response_bodies(
    driver: webdriver.Chrome,
    records: list[dict[str, Any]],
    out_dir: Path,
    warnings: list[CaptureWarning],
) -> tuple[list[dict[str, Any]], dict[str, str]]:
    saved_resources: list[dict[str, Any]] = []
    url_to_local: dict[str, str] = {}
    seen_urls: set[str] = set()

    for record in records:
        request_id = record.get("request_id")
        response = record.get("response", {})
        request = record.get("request", {})
        resource_url = _normalize_url(response.get("url") or request.get("url", ""))
        mime_type = response.get("mime_type")

        if not request_id or not resource_url:
            continue
        if not resource_url.startswith(("http://", "https://")):
            continue
        if resource_url in seen_urls:
            continue

        seen_urls.add(resource_url)

        try:
            body_obj = driver.execute_cdp_cmd("Network.getResponseBody", {"requestId": request_id})
            body = body_obj.get("body", "")
            is_b64 = bool(body_obj.get("base64Encoded", False))
            data = base64.b64decode(body) if is_b64 else body.encode("utf-8", errors="ignore")

            rel = _safe_relpath(resource_url, mime_type)
            full_path = out_dir / rel
            full_path.parent.mkdir(parents=True, exist_ok=True)
            full_path.write_bytes(data)

            local_ref = "/" + str(rel).replace("\\", "/")
            url_to_local[resource_url] = local_ref

            saved_resources.append({
                "request_id": request_id,
                "url": resource_url,
                "status": response.get("status"),
                "mime_type": mime_type,
                "bytes": len(data),
                "relpath": str(rel).replace("\\", "/"),
                "local_ref": local_ref,
                "from_service_worker": response.get("from_service_worker"),
                "from_disk_cache": response.get("from_disk_cache"),
            })
        except Exception as exc:
            warnings.append(CaptureWarning(
                stage="get_response_body",
                message=str(exc),
                url=resource_url,
                request_id=request_id,
            ))

    return saved_resources, url_to_local


def _replace_url_value(raw: str, base_url: str, url_to_local: dict[str, str], missing: set[str]) -> str:
    value = htmlmod.unescape(htmlmod.unescape(raw.strip()))
    lowered = value.lower()

    if lowered.startswith(("#", "mailto:", "tel:", "javascript:", "data:", "blob:", "about:")):
        return raw

    absolute = _normalize_url(urljoin(base_url, value))
    local = url_to_local.get(absolute)
    if not local:
        missing.add(absolute)
        return raw
    return local


def _rewrite_css_text(css_text: str, css_url: str, url_to_local: dict[str, str], missing: set[str]) -> str:
    def repl(match: re.Match[str]) -> str:
        raw = match.group(2).strip().strip('"').strip("'")
        replaced = _replace_url_value(raw, css_url, url_to_local, missing)
        if replaced == raw:
            return match.group(0)
        return f'url("{replaced}")'

    return re.sub(r"url\(\s*([\"']?)(.*?)\1\s*\)", repl, css_text, flags=re.IGNORECASE)


def _rewrite_srcset_value(srcset: str, base_url: str, url_to_local: dict[str, str], missing: set[str]) -> str:
    parts = []
    for item in srcset.split(","):
        piece = item.strip()
        if not piece:
            continue
        tokens = piece.split()
        url_token = tokens[0]
        descriptor = " ".join(tokens[1:])
        replaced = _replace_url_value(url_token, base_url, url_to_local, missing)
        parts.append((replaced + (" " + descriptor if descriptor else "")).strip())
    return ", ".join(parts)


def _rewrite_html_to_local(
    html_text: str,
    base_url: str,
    main_host: str,
    url_to_local: dict[str, str],
) -> tuple[str, list[str]]:
    missing: set[str] = set()

    html_text = re.sub(
        r'<meta[^>]+http-equiv=["\']content-security-policy["\'][^>]*>\s*',
        "",
        html_text,
        flags=re.IGNORECASE,
    )

    base_tag = f'<base href="/{main_host}/">\n'
    inject = f'{base_tag}<script src="/{main_host}/_offline_bootstrap.js"></script>\n'

    if re.search(r"<head[^>]*>", html_text, flags=re.IGNORECASE):
        html_text = re.sub(r"(<head[^>]*>\s*)", r"\1" + inject, html_text, count=1, flags=re.IGNORECASE)
    else:
        html_text = inject + html_text

    attr_names = ["src", "href", "data-src", "poster"]

    def repl_attr(match: re.Match[str]) -> str:
        attr = match.group(1)
        quote = match.group(2)
        raw = match.group(3)
        replaced = _replace_url_value(raw, base_url, url_to_local, missing)
        return f"{attr}={quote}{replaced}{quote}"

    for attr in attr_names:
        pattern = rf'({attr})\s*=\s*(["\'])(.*?)\2'
        html_text = re.sub(pattern, repl_attr, html_text, flags=re.IGNORECASE)

    def repl_srcset(match: re.Match[str]) -> str:
        quote = match.group(1)
        raw = match.group(2)
        replaced = _rewrite_srcset_value(raw, base_url, url_to_local, missing)
        return f'srcset={quote}{replaced}{quote}'

    html_text = re.sub(r'srcset\s*=\s*(["\'])(.*?)\1', repl_srcset, html_text, flags=re.IGNORECASE)

    def repl_style_attr(match: re.Match[str]) -> str:
        quote = match.group(1)
        raw = match.group(2)
        rewritten = _rewrite_css_text(raw, base_url, url_to_local, missing)
        return f'style={quote}{rewritten}{quote}'

    html_text = re.sub(r'style\s*=\s*(["\'])(.*?)\1', repl_style_attr, html_text, flags=re.IGNORECASE | re.DOTALL)

    def repl_style_block(match: re.Match[str]) -> str:
        open_tag = match.group(1)
        css_body = match.group(2)
        close_tag = match.group(3)
        rewritten = _rewrite_css_text(css_body, base_url, url_to_local, missing)
        return f"{open_tag}{rewritten}{close_tag}"

    html_text = re.sub(r"(<style[^>]*>)(.*?)(</style>)", repl_style_block, html_text, flags=re.IGNORECASE | re.DOTALL)

    def repl_meta_refresh(match: re.Match[str]) -> str:
        quote = match.group(1)
        raw = match.group(2)
        meta_match = re.match(r"(\s*\d+\s*;\s*url\s*=\s*)(.*)", raw, flags=re.IGNORECASE)
        if not meta_match:
            return match.group(0)
        prefix, url_part = meta_match.groups()
        replaced = _replace_url_value(url_part, base_url, url_to_local, missing)
        return f'content={quote}{prefix}{replaced}{quote}'

    html_text = re.sub(r'content\s*=\s*(["\'])(.*?)\1', repl_meta_refresh, html_text, flags=re.IGNORECASE)

    return html_text, sorted(missing)


def _build_offline_bootstrap(root_url: str, url_to_local: dict[str, str]) -> str:
    payload = {
        "root_url": root_url,
        "map": url_to_local,
    }
    return (
        "(function() {\n"
        f"  const PAYLOAD = {json.dumps(payload, ensure_ascii=False)};\n"
        "  const ORIGINAL_BASE = PAYLOAD.root_url;\n"
        "  const MAP = PAYLOAD.map;\n"
        "  const normalize = (u) => {\n"
        "    try {\n"
        "      const abs = new URL(u, ORIGINAL_BASE);\n"
        "      abs.hash = '';\n"
        "      return abs.toString();\n"
        "    } catch (e) {\n"
        "      return u;\n"
        "    }\n"
        "  };\n"
        "  const resolve = (u) => MAP[normalize(u)] || null;\n"
        "  if ('serviceWorker' in navigator) {\n"
        "    window.addEventListener('load', () => {\n"
        "      navigator.serviceWorker.getRegistrations()\n"
        "        .then(regs => Promise.all(regs.map(r => r.unregister())))\n"
        "        .catch(() => null);\n"
        "    });\n"
        "  }\n"
        "  const realFetch = window.fetch ? window.fetch.bind(window) : null;\n"
        "  if (realFetch) {\n"
        "    window.fetch = function(input, init) {\n"
        "      const requestUrl = (typeof input === 'string') ? input : (input && input.url) ? input.url : '';\n"
        "      const method = (init && init.method) ? String(init.method).toUpperCase() : 'GET';\n"
        "      const local = resolve(requestUrl);\n"
        "      if ((method === 'GET' || method === 'HEAD') && local) {\n"
        "        return realFetch(local, init);\n"
        "      }\n"
        "      return realFetch(input, init);\n"
        "    };\n"
        "  }\n"
        "  const RealXHR = window.XMLHttpRequest;\n"
        "  if (RealXHR) {\n"
        "    function PatchedXHR() {\n"
        "      const xhr = new RealXHR();\n"
        "      const open0 = xhr.open;\n"
        "      xhr.open = function(method, url, async, user, password) {\n"
        "        try {\n"
        "          const m = String(method || 'GET').toUpperCase();\n"
        "          const local = resolve(url);\n"
        "          if ((m === 'GET' || m === 'HEAD') && local) {\n"
        "            return open0.call(xhr, method, local, async, user, password);\n"
        "          }\n"
        "        } catch (e) {}\n"
        "        return open0.call(xhr, method, url, async, user, password);\n"
        "      };\n"
        "      return xhr;\n"
        "    }\n"
        "    window.XMLHttpRequest = PatchedXHR;\n"
        "  }\n"
        "  const RealEventSource = window.EventSource;\n"
        "  if (RealEventSource) {\n"
        "    window.EventSource = function(url, config) {\n"
        "      const local = resolve(url);\n"
        "      return new RealEventSource(local || url, config);\n"
        "    };\n"
        "  }\n"
        "  window.__SITE_DUMP_MAP__ = MAP;\n"
        "})();\n"
    )


def _write_runner_files(config: CaptureConfig, out_dir: Path, offline_relpath: str) -> None:
    normalized_target = "/" + offline_relpath.replace("\\", "/")
    launch_site_py = """from __future__ import annotations

import http.server
import socket
import socketserver
import webbrowser
from pathlib import Path

ROOT = Path(__file__).resolve().parent
TARGET = "__TARGET__"


class QuietHandler(http.server.SimpleHTTPRequestHandler):
    def log_message(self, format, *args):
        return


def find_port() -> int:
    sock = socket.socket()
    sock.bind(("127.0.0.1", 0))
    port = sock.getsockname()[1]
    sock.close()
    return port


def main() -> None:
    port = find_port()
    handler = lambda *args, **kwargs: QuietHandler(*args, directory=str(ROOT), **kwargs)
    with socketserver.ThreadingTCPServer(("127.0.0.1", port), handler) as httpd:
        url = f"http://127.0.0.1:{port}{TARGET}"
        print("Servidor local:", url)
        try:
            webbrowser.open(url)
        except Exception:
            pass
        print("Presiona Ctrl+C para cerrar.")
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\\nCerrando servidor...")


if __name__ == "__main__":
    main()
""".replace("__TARGET__", normalized_target)

    run_bat = f'@echo off\r\npython "%~dp0{config.runner_python_name}"\r\n'
    run_sh = f'#!/usr/bin/env bash\npython3 "$(dirname "$0")/{config.runner_python_name}"\n'

    (out_dir / config.runner_python_name).write_text(launch_site_py, encoding="utf-8")
    (out_dir / config.runner_bat_name).write_text(run_bat, encoding="utf-8")
    (out_dir / config.runner_sh_name).write_text(run_sh, encoding="utf-8")


def _write_capture_readme(out_dir: Path) -> None:
    text = (
        "Este bundle sirve para inspección y reconstrucción offline parcial.\n\n"
        "Qué sí incluye:\n"
        "- HTML final renderizado\n"
        "- Recursos HTTP(S) cuyo body se pudo recuperar\n"
        "- Reescritura de referencias HTML y CSS\n"
        "- Cookies, storage, screenshot, logs, red y runner local\n\n"
        "Qué no garantiza:\n"
        "- Reproducir POST, login vivo o estados del backend\n"
        "- Reproducir WebSocket frames\n"
        "- Reproducir service workers complejos o apps 100% SPA con lógica remota\n"
        "- Reproducir blob: creados dinámicamente y no materializados como respuestas HTTP(S)\n"
    )
    (out_dir / "README_CAPTURE.txt").write_text(text, encoding="utf-8")


def _zip_directory(source_dir: Path, zip_path: Path) -> None:
    zip_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for path in source_dir.rglob("*"):
            if path.is_file():
                zf.write(path, arcname=str(path.relative_to(source_dir.parent)))


def capture_site(config: CaptureConfig) -> CaptureResult:
    out_dir = Path(config.out_dir).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    warnings: list[CaptureWarning] = []
    perf_logs: list[dict[str, Any]] = []
    browser_logs: list[dict[str, Any]] = []

    driver = _build_driver(config)
    bundle_zip: Path | None = None

    try:
        driver.execute_cdp_cmd("Network.enable", {"maxTotalBufferSize": 150_000_000, "maxResourceBufferSize": 100_000_000})
        driver.execute_cdp_cmd("Page.enable", {})
        if config.disable_cache:
            driver.execute_cdp_cmd("Network.setCacheDisabled", {"cacheDisabled": True})

        driver.get(config.url)
        _wait_for_network_stable(driver, config, perf_logs, browser_logs)

        final_url = driver.current_url
        final_html = driver.page_source
        main_host = urlparse(final_url).netloc or "nohost"
        host_dir = out_dir / main_host
        host_dir.mkdir(parents=True, exist_ok=True)

        (host_dir / "index.final.html").write_text(final_html, encoding="utf-8")

        if config.capture_screenshot:
            with contextlib.suppress(Exception):
                driver.save_screenshot(str(host_dir / "screenshot.png"))

        if config.capture_console:
            _json_dump(host_dir / "console_logs.json", browser_logs)

        records = _parse_perf_logs(perf_logs)
        _jsonl_dump(host_dir / "network_records.jsonl", records)

        saved_resources, url_to_local = _save_response_bodies(driver, records, out_dir, warnings)

        for resource in saved_resources:
            mime = (resource.get("mime_type") or "").split(";")[0].strip().lower()
            relpath = resource["relpath"]
            if mime != "text/css" and not relpath.lower().endswith(".css"):
                continue

            css_path = out_dir / relpath
            try:
                css_text = css_path.read_text(encoding="utf-8", errors="ignore")
                missing_css: set[str] = set()
                css_path.write_text(
                    _rewrite_css_text(css_text, resource["url"], url_to_local, missing_css),
                    encoding="utf-8",
                )
                if missing_css:
                    for missing_url in sorted(missing_css):
                        warnings.append(CaptureWarning(
                            stage="rewrite_css_missing",
                            message="No local mapping found",
                            url=missing_url,
                        ))
            except Exception as exc:
                warnings.append(CaptureWarning(stage="rewrite_css", message=str(exc), url=resource["url"]))

        bootstrap_js = _build_offline_bootstrap(final_url, url_to_local)
        (host_dir / "_offline_bootstrap.js").write_text(bootstrap_js, encoding="utf-8")

        offline_html, missing_refs = _rewrite_html_to_local(final_html, final_url, main_host, url_to_local)
        (host_dir / "index.offline.html").write_text(offline_html, encoding="utf-8")
        (host_dir / "missing_refs.txt").write_text("\n".join(missing_refs) + ("\n" if missing_refs else ""), encoding="utf-8")

        if config.capture_storage:
            _json_dump(host_dir / "storage.json", _capture_storage(driver))
            _json_dump(host_dir / "cookies.json", _capture_cookies(driver))
            _json_dump(host_dir / "cookies_cdp.json", _get_all_cookies_cdp(driver))

        if config.capture_same_origin_iframes:
            iframe_dir = host_dir / "iframes"
            iframe_dir.mkdir(parents=True, exist_ok=True)
            iframe_rows = _capture_same_origin_iframes(driver)
            manifest_rows = []
            for row in iframe_rows:
                item = dict(row)
                html_value = item.pop("html", None)
                if item.get("sameOrigin") and isinstance(html_value, str):
                    iframe_name = f'iframe_{item.get("index", 0):03d}.html'
                    (iframe_dir / iframe_name).write_text(html_value, encoding="utf-8")
                    item["saved_html"] = f"iframes/{iframe_name}"
                manifest_rows.append(item)
            _json_dump(host_dir / "iframes.json", manifest_rows)

        _json_dump(host_dir / "runtime.json", _detect_runtime())
        _jsonl_dump(host_dir / "saved_resources.jsonl", saved_resources)
        _json_dump(host_dir / "warnings.json", [asdict(w) for w in warnings])

        manifest = {
            "root_url": config.url,
            "final_url": final_url,
            "out_dir": str(out_dir),
            "host_dir": str(host_dir),
            "final_html": f"{main_host}/index.final.html",
            "offline_html": f"{main_host}/index.offline.html",
            "saved_count": len(saved_resources),
            "missing_refs_count": len(missing_refs),
            "warnings_count": len(warnings),
            "runner": config.runner_bat_name,
            "notes": [
                "Ejecuta el runner generado desde la raíz del bundle.",
                "La copia offline es parcial. Sirve para análisis y para rearmado local de muchos sitios, no para replicar cualquier backend.",
            ],
        }
        _json_dump(out_dir / "dump_index.json", manifest)
        tree_rows = [item["relpath"] for item in saved_resources]
        tree_rows.extend([f"{main_host}/index.final.html", f"{main_host}/index.offline.html"])
        (out_dir / "tree.txt").write_text("\n".join(sorted(tree_rows)), encoding="utf-8")

        _write_runner_files(config, out_dir, manifest["offline_html"])
        _write_capture_readme(out_dir)

        if config.zip_output:
            zip_name = config.bundle_zip_name or f"{out_dir.name}.zip"
            bundle_zip = out_dir.parent / zip_name
            _zip_directory(out_dir, bundle_zip)

        return CaptureResult(
            root_url=config.url,
            final_url=final_url,
            out_dir=str(out_dir),
            bundle_zip=str(bundle_zip) if bundle_zip else None,
            offline_html=manifest["offline_html"],
            final_html=manifest["final_html"],
            saved_count=len(saved_resources),
            missing_refs_count=len(missing_refs),
            warnings_count=len(warnings),
            runner=config.runner_bat_name,
        )

    finally:
        if not config.open_browser_after_capture:
            with contextlib.suppress(Exception):
                driver.quit()



def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Ejecutable principal para el volcado de sitios."
    )
    parser.add_argument(
        "--config",
        default="config.json",
        help="Ruta al archivo de configuración JSON."
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    config_path = Path(args.config).resolve()
    config = load_config_file(config_path)
    result = capture_site(config)

    print("Config:", config_path)
    print("Final URL:", result.final_url)
    print("Salida:", result.out_dir)
    print("HTML offline:", result.offline_html)
    print("Recursos guardados:", result.saved_count)
    print("Referencias faltantes:", result.missing_refs_count)
    print("Warnings:", result.warnings_count)
    print("Runner:", result.runner)
    if result.bundle_zip:
        print("ZIP:", result.bundle_zip)


if __name__ == "__main__":
    import argparse
    main()
