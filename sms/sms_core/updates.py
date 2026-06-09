from dataclasses import dataclass
import json
import ssl
import time
import urllib.request


@dataclass(frozen=True)
class UpdateCheckPlan:
    kind: str
    current_version: str
    latest_tag: str = ""
    download_url: str = ""


def normalize_proxy_base(value: str) -> str:
    text = (value or "").strip()
    if not text:
        return ""
    if not (text.startswith("http://") or text.startswith("https://")):
        text = "https://" + text
    if not text.endswith("/"):
        text += "/"
    return text


def normalize_config_bases(proxy_base: str, api_proxy_base: str):
    return normalize_proxy_base(proxy_base), normalize_proxy_base(api_proxy_base)


def build_release_api_urls(owner: str, repo: str, api_proxy_base: str, now=None):
    api_path = f"/repos/{owner}/{repo}/releases/latest"
    urls = []
    timestamp = int(time.time() if now is None else now)

    if api_proxy_base:
        for base in (item.strip() for item in api_proxy_base.split("|") if item.strip()):
            base = normalize_proxy_base(base).rstrip("/")
            urls.append(base + api_path + f"?t={timestamp}")

    urls.append("https://api.github.com" + api_path)
    return urls


def pick_zip_asset(release_json: dict):
    assets = release_json.get("assets") or []
    zip_assets = [asset for asset in assets if (asset.get("name", "").lower().endswith(".zip"))]
    if not zip_assets:
        return None
    zip_assets.sort(key=lambda asset: -int(asset.get("size", 0) or 0))
    return zip_assets[0]


def version_tuple(version: str):
    text = (version or "").strip().lstrip("vV")
    parts = []
    for item in text.split("."):
        try:
            parts.append(int(item))
        except Exception:
            parts.append(0)
    while len(parts) < 3:
        parts.append(0)
    return tuple(parts[:3])


def build_download_url(raw_url: str, proxy_base: str):
    url = raw_url or ""
    proxy = normalize_proxy_base(proxy_base)
    if proxy and url.startswith("http"):
        return proxy + url
    return url


def fetch_latest_release(owner: str, repo: str, api_proxy_base: str, get_json=None):
    get_json = get_json or http_get_json
    urls = build_release_api_urls(owner, repo, api_proxy_base)
    last_err = None

    for url in urls:
        try:
            return get_json(url, timeout=8, retries=3)
        except Exception as exc:
            last_err = exc

    raise RuntimeError(f"获取最新版本信息失败：{last_err}")


def plan_update_check(
    release_json: dict,
    current_version: str,
    proxy_base: str,
    *,
    pick_asset=pick_zip_asset,
    parse_version=version_tuple,
):
    tag = (release_json or {}).get("tag_name") or ""
    latest = parse_version(tag)
    current = parse_version(current_version)

    if latest <= current:
        return UpdateCheckPlan("latest", current_version=current_version, latest_tag=tag)

    asset = pick_asset(release_json or {})
    if not asset:
        return UpdateCheckPlan("no_zip", current_version=current_version, latest_tag=tag)

    raw_url = asset.get("browser_download_url") or ""
    return UpdateCheckPlan(
        "update",
        current_version=current_version,
        latest_tag=tag,
        download_url=build_download_url(raw_url, proxy_base),
    )


def check_latest_release(
    owner: str,
    repo: str,
    current_version: str,
    proxy_base: str,
    api_proxy_base: str,
    *,
    get_json=None,
):
    release_json = fetch_latest_release(owner, repo, api_proxy_base, get_json=get_json)
    return plan_update_check(release_json, current_version, proxy_base)


def http_get_json(url: str, timeout=8, retries=3):
    last_err = None
    opener = _build_no_proxy_opener()

    for index in range(retries):
        try:
            req = urllib.request.Request(
                url,
                headers={"User-Agent": "sms-updater", "Accept": "application/vnd.github+json"},
                method="GET",
            )
            with opener.open(req, timeout=timeout) as resp:
                data = resp.read().decode("utf-8", "ignore")
                return json.loads(data)
        except Exception as exc:
            last_err = exc
            _sleep_backoff(0.6, index)

    raise last_err


def http_probe(url: str, timeout=8, retries=2):
    last_err = None
    opener = _build_no_proxy_opener()

    for index in range(retries):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "sms-updater"}, method="HEAD")
            with opener.open(req, timeout=timeout) as resp:
                return True, f"HTTP {getattr(resp, 'status', 200)}"
        except Exception:
            try:
                req = urllib.request.Request(url, headers={"User-Agent": "sms-updater"}, method="GET")
                with opener.open(req, timeout=timeout) as resp:
                    resp.read(64)
                    return True, f"HTTP {getattr(resp, 'status', 200)}"
            except Exception as exc:
                last_err = exc
                _sleep_backoff(0.4, index)

    raise last_err


def classify_update_error(exc: Exception) -> str:
    text = str(exc)
    lower_text = text.lower()
    if "sslv3_alert_handshake_failure" in text or "sslv3 alert handshake failure" in lower_text:
        return "TLS握手失败（代理节点/线路不兼容或被干扰）"
    if "timed out" in lower_text:
        return "连接超时（线路慢/被阻断）"
    if "name or service not known" in lower_text or "getaddrinfo failed" in lower_text:
        return "DNS 解析失败"
    return text


def test_update_proxy_connectivity(
    owner: str,
    repo: str,
    api_proxy_base: str,
    proxy_base: str,
    get_json=http_get_json,
    probe=http_probe,
    pick_asset=pick_zip_asset,
):
    api_path = f"/repos/{owner}/{repo}/releases/latest"
    checks = []
    ok_bases = []
    ok_api = False
    release_json = None

    bases = [item.strip() for item in str(api_proxy_base or "").split("|") if item.strip()]
    for base in bases:
        base_n = normalize_proxy_base(base)
        url = base_n.rstrip("/") + api_path
        try:
            release_json = get_json(url, timeout=6, retries=2)
            checks.append((base_n, True, "OK"))
            ok_bases.append(base_n)
            ok_api = True
            break
        except Exception as exc:
            checks.append((base_n, False, classify_update_error(exc)))

    if not ok_api:
        direct_url = "https://api.github.com" + api_path
        try:
            release_json = get_json(direct_url, timeout=6, retries=2)
            checks.append(("直连 api.github.com", True, "OK"))
            ok_api = True
        except Exception as exc:
            checks.append(("直连 api.github.com", False, classify_update_error(exc)))

    if ok_api and release_json:
        asset = pick_asset(release_json)
        if asset:
            raw_url = asset.get("browser_download_url") or ""
            pb = normalize_proxy_base(proxy_base)
            if pb and raw_url.startswith("http"):
                test_url = pb + raw_url
                try:
                    probe(test_url, timeout=6, retries=2)
                    checks.append((f"下载代理 {pb}", True, "OK"))
                except Exception as exc:
                    checks.append((f"下载代理 {pb}", False, classify_update_error(exc)))
            else:
                checks.append(("下载代理 proxy_base", False, "未填写或获取不到下载链接"))
        else:
            checks.append(("下载代理 proxy_base", False, "Release 无 .zip 附件"))

    return {
        "checks": checks,
        "ok_bases": ok_bases,
        "download_ok": any(
            ok is True and isinstance(name, str) and name.startswith("下载代理 ")
            for name, ok, _info in checks
        ),
    }


def format_proxy_test_result(result: dict):
    checks = result.get("checks") or []
    lines = [("✅ " if ok else "❌ ") + f"{name}：{info}" for name, ok, info in checks]

    ok_bases = result.get("ok_bases") or []
    download_ok = bool(result.get("download_ok"))
    if ok_bases or download_ok:
        lines.append("")
    if ok_bases:
        lines.append("提示：API 代理可用，检测更新将优先使用它。")
    if download_ok:
        lines.append("提示：下载代理可用，下载链接将优先使用它。")
    return "\n".join(lines)


def _build_no_proxy_opener():
    ctx = ssl.create_default_context()
    return urllib.request.build_opener(
        urllib.request.ProxyHandler({}),
        urllib.request.HTTPSHandler(context=ctx),
    )


def _sleep_backoff(base_seconds, index):
    try:
        time.sleep(base_seconds * (2 ** index))
    except Exception:
        pass
