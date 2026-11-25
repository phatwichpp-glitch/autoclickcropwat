"""
tests/test_updater_check.py
============================
v1.1.1 — ล็อกพฤติกรรมเช็คอัปเดต (ไม่ต่อเน็ตจริง: แทน urlopen ด้วยของปลอม)

อาการจริงจากผู้ใช้: บางเครื่องกดเช็คแล้วไม่ขึ้นเวอร์ชันใหม่ — GitHub API แบบไม่ล็อกอิน
จำกัด 60 ครั้ง/ชั่วโมงต่อ IP เครื่องในเน็ตเดียวกันใช้โควตาร่วมกันจนได้ 403
"""

import io
import json
import urllib.error

import updater


class _Resp(io.BytesIO):
    def __init__(self, body=b"", url=""):
        super().__init__(body)
        self._url = url

    def geturl(self):
        return self._url

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False


def _fake_urlopen(api=None, web_url=None):
    def urlopen(req, timeout=None):
        url = req.full_url
        if "api.github.com" in url:
            if isinstance(api, Exception):
                raise api
            return _Resp(json.dumps(api).encode())
        if url.endswith("/releases/latest"):
            assert req.get_method() == "HEAD"
            if isinstance(web_url, Exception):
                raise web_url
            return _Resp(url=web_url)
        raise AssertionError(f"unexpected url {url}")
    return urlopen


def _rate_limited():
    return urllib.error.HTTPError(updater.API_LATEST, 403, "rate limit exceeded", {}, None)


def test_api_ok(monkeypatch):
    monkeypatch.setattr(updater, "APP_VERSION", "1.0.0")
    api = {"tag_name": "v1.1.0", "body": "notes",
           "assets": [{"name": updater.ASSET_NAME, "browser_download_url": "https://x/CropWatAutoRunner.exe"}]}
    monkeypatch.setattr(updater.urllib.request, "urlopen", _fake_urlopen(api=api))
    r = updater.check_for_update()
    assert r["update_available"] and r["latest"] == "1.1.0" and r["notes"] == "notes" and r["error"] is None


def test_rate_limited_api_falls_back_to_releases_page(monkeypatch):
    monkeypatch.setattr(updater, "APP_VERSION", "1.0.0")
    monkeypatch.setattr(updater.urllib.request, "urlopen", _fake_urlopen(
        api=_rate_limited(),
        web_url="https://github.com/phatwichpp-glitch/autoclickcropwat/releases/tag/v1.1.0",
    ))
    r = updater.check_for_update()
    assert r["error"] is None
    assert r["update_available"] and r["latest"] == "1.1.0"
    assert r["asset_url"] == (
        "https://github.com/phatwichpp-glitch/autoclickcropwat/releases/download/v1.1.0/CropWatAutoRunner.exe"
    )


def test_fallback_same_version_is_not_update(monkeypatch):
    monkeypatch.setattr(updater, "APP_VERSION", "1.1.0")
    monkeypatch.setattr(updater.urllib.request, "urlopen", _fake_urlopen(
        api=_rate_limited(),
        web_url="https://github.com/phatwichpp-glitch/autoclickcropwat/releases/tag/v1.1.0",
    ))
    r = updater.check_for_update()
    assert r["error"] is None and not r["update_available"]


def test_both_fail_reports_error_not_latest(monkeypatch):
    monkeypatch.setattr(updater, "APP_VERSION", "1.0.0")
    monkeypatch.setattr(updater.urllib.request, "urlopen", _fake_urlopen(
        api=_rate_limited(), web_url=urllib.error.URLError("offline"),
    ))
    r = updater.check_for_update()
    assert r["error"] and not r["update_available"]


def test_releases_page_without_tag_is_error(monkeypatch):
    monkeypatch.setattr(updater.urllib.request, "urlopen", _fake_urlopen(
        api=_rate_limited(), web_url="https://github.com/phatwichpp-glitch/autoclickcropwat/releases",
    ))
    r = updater.check_for_update()
    assert r["error"] and not r["update_available"]
