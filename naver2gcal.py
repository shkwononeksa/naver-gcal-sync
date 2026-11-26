#!/usr/bin/env python3
"""naver2gcal — 네이버 캘린더(CalDAV) → Google 캘린더 단방향 미러 / 이관 도구.

명령:
  set-password  네이버 애플리케이션 비밀번호를 Windows 자격 증명 관리자(keyring)에 저장
  discover      CalDAV로 보이는 네이버 캘린더 목록 확인 (팀 공유 캘린더 노출 여부 판정)
  auth          Google OAuth 최초 인증 (token.json 생성)
  export        네이버 캘린더 → .ics 파일 (Google '가져오기'용, 개인 일정 이관)
  sync          [[mirror]] 설정대로 네이버 → Google 단방향 미러 (작업 스케줄러로 5분마다 실행)

설계 원칙
  * 단방향·읽기전용 미러: 네이버가 원본(Source of truth), Google은 사본.
  * 이 도구는 자신이 만든 일정(extendedProperties.private.n2gM=<mirror id>)만 수정/삭제한다.
  * 반복 일정은 개별 회차로 펼쳐서(expand) 넣는다 → RRULE/예외/시간대 변환 오류를 원천 차단.
  * 결정적 이벤트 ID(sha1) + 내용 해시 비교 → 바뀐 것만 갱신, 로컬 상태 파일 없음.
  * 대량 삭제 보호: 네이버 조회가 0건이거나 삭제가 과도하면 삭제를 보류한다.
"""
from __future__ import annotations

import argparse
import datetime as dt
import getpass
import hashlib
import json
import logging
import logging.handlers
import os
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable
from urllib.parse import urljoin
from zoneinfo import ZoneInfo

try:
    import tomllib  # Python 3.11+
except ModuleNotFoundError:  # pragma: no cover
    import tomli as tomllib  # type: ignore

APP = "naver2gcal"
BASE_DIR = Path(__file__).resolve().parent
SCOPES = ["https://www.googleapis.com/auth/calendar"]
# 본인 로그인(OAuth)용: Google Tasks 이관 + 기본 캘린더 중복 정리
USER_SCOPES = ["https://www.googleapis.com/auth/tasks", "https://www.googleapis.com/auth/calendar.events"]
DAY = dt.timedelta(days=1)
log = logging.getLogger(APP)


# ─────────────────────────────── 설정 ───────────────────────────────

@dataclass
class MirrorCfg:
    naver: str                      # 네이버 캘린더 표시 이름(또는 CalDAV URL)
    google: str                     # Google 캘린더 이름(없으면 생성) 또는 캘린더 ID
    id: str                         # 안정적인 미러 식별자 (이벤트 ID 생성에 사용)
    color_id: str | None = None
    reminders: str = "none"         # none | default | source
    title_prefix: str = ""
    copy_description: bool = True
    past_days: int | None = None    # 미러별 기간 (없으면 [sync] 값)
    future_days: int | None = None


@dataclass
class Config:
    naver_url: str
    naver_username: str
    timezone: str = "Asia/Seoul"
    past_days: int = 30
    future_days: int = 365
    fetch_mode: str = "range"       # range | all
    delete_guard_ratio: float = 0.5
    google_auth: str = "oauth"      # oauth | service_account
    google_credentials: Path = BASE_DIR / "credentials.json"
    google_token: Path = BASE_DIR / "token.json"
    user_credentials: Path = BASE_DIR / "credentials.json"   # [google_user] 본인 로그인용 OAuth 클라이언트
    user_token: Path = BASE_DIR / "token_user.json"
    log_file: Path | None = None
    mirrors: list[MirrorCfg] = field(default_factory=list)


def _short_hash(text: str, n: int = 8) -> str:
    return hashlib.sha1(text.encode("utf-8")).hexdigest()[:n]


def load_config(path: Path) -> Config:
    if not path.exists():
        raise SystemExit(
            f"설정 파일이 없습니다: {path}\n"
            f"config.example.toml 을 config.toml 로 복사한 뒤 수정하세요."
        )
    raw = tomllib.loads(path.read_text(encoding="utf-8"))
    base = path.resolve().parent

    def p(value: str | None, default: str) -> Path:
        v = Path(value or default)
        return v if v.is_absolute() else base / v

    naver = raw.get("naver", {})
    google = raw.get("google", {})
    sync = raw.get("sync", {})
    if not naver.get("username"):
        raise SystemExit("config.toml 의 [naver] username 을 입력하세요 (네이버 아이디).")

    mirrors = []
    for m in raw.get("mirror", []):
        if not m.get("naver") or not m.get("google"):
            raise SystemExit("[[mirror]] 항목에는 naver 와 google 값이 모두 필요합니다.")
        reminders = str(m.get("reminders", "none")).lower()
        if reminders not in ("none", "default", "source"):
            raise SystemExit(f"reminders 값은 none/default/source 중 하나여야 합니다: {reminders}")
        mirrors.append(MirrorCfg(
            naver=str(m["naver"]),
            google=str(m["google"]),
            id=str(m.get("id") or _short_hash(str(m["naver"]))),
            color_id=str(m["color_id"]) if m.get("color_id") else None,
            reminders=reminders,
            title_prefix=str(m.get("title_prefix", "")),
            copy_description=bool(m.get("copy_description", True)),
            past_days=int(m["past_days"]) if "past_days" in m else None,
            future_days=int(m["future_days"]) if "future_days" in m else None,
        ))
    ids = [m.id for m in mirrors]
    if len(ids) != len(set(ids)):
        raise SystemExit("[[mirror]] 의 id(또는 naver 이름)가 중복됩니다.")

    cfg = Config(
        naver_url=str(naver.get("url", "https://caldav.calendar.naver.com/")),
        naver_username=str(naver["username"]).strip(),
        timezone=str(sync.get("timezone", "Asia/Seoul")),
        past_days=int(sync.get("past_days", 30)),
        future_days=int(sync.get("future_days", 365)),
        fetch_mode=str(sync.get("fetch_mode", "range")).lower(),
        delete_guard_ratio=float(sync.get("delete_guard_ratio", 0.5)),
        google_auth=str(google.get("auth", "oauth")).lower(),
        google_credentials=p(google.get("credentials"), "credentials.json"),
        google_token=p(google.get("token"), "token.json"),
        user_credentials=p(raw.get("google_user", {}).get("credentials"), "credentials.json"),
        user_token=p(raw.get("google_user", {}).get("token"), "token_user.json"),
        log_file=p(str(lf), "logs/naver2gcal.log") if (lf := sync.get("log_file", "logs/naver2gcal.log")) else None,
        mirrors=mirrors,
    )
    ZoneInfo(cfg.timezone)  # 잘못된 시간대면 여기서 예외
    if cfg.fetch_mode not in ("range", "all"):
        raise SystemExit("fetch_mode 는 range 또는 all 이어야 합니다.")
    if cfg.google_auth not in ("oauth", "service_account"):
        raise SystemExit("google.auth 는 oauth 또는 service_account 이어야 합니다.")
    return cfg


# ───────────────────────── 네이버 CalDAV ─────────────────────────

def naver_password(cfg: Config) -> str:
    pw = os.environ.get("NAVER_CALDAV_PASSWORD")
    if pw:
        return pw.strip()
    try:
        import keyring
        pw = keyring.get_password(APP, cfg.naver_username)
    except Exception as e:  # keyring 백엔드 없음 등
        log.debug("keyring 조회 실패: %s", e)
    if not pw:
        raise SystemExit(
            "네이버 애플리케이션 비밀번호가 저장되어 있지 않습니다.\n"
            "  python naver2gcal.py set-password   (또는 환경변수 NAVER_CALDAV_PASSWORD)"
        )
    return pw


def naver_principal(cfg: Config):
    """(client, principal) 반환. 루트 URL → /.well-known/caldav 순으로 탐색."""
    import caldav
    # caldav 가 import 시 로거 레벨을 바꾸므로 여기서 다시 낮춘다 (네이버 VTODO의 DTSTAMP 누락 경고 대량 출력 방지)
    logging.getLogger("caldav").setLevel(logging.DEBUG if log.isEnabledFor(logging.DEBUG) else logging.ERROR)

    pw = naver_password(cfg)
    tried: list[str] = []
    candidates = [cfg.naver_url]
    last_err: Exception | None = None
    while candidates:
        url = candidates.pop(0)
        if url in tried:
            continue
        tried.append(url)
        try:
            client = caldav.DAVClient(url=url, username=cfg.naver_username, password=pw)
            principal = client.principal()
            log.debug("CalDAV principal: %s", principal.url)
            return client, principal
        except Exception as e:
            last_err = e
            log.debug("principal 탐색 실패 (%s): %r", url, e)
            if "Authoriz" in type(e).__name__ or "401" in str(e):
                break
            if not any(u.endswith("/.well-known/caldav") for u in tried):
                wk = urljoin(cfg.naver_url, "/.well-known/caldav")
                try:
                    import requests
                    r = requests.request("PROPFIND", wk, auth=(cfg.naver_username, pw),
                                         headers={"Depth": "0"}, allow_redirects=False, timeout=30)
                    loc = r.headers.get("Location")
                    tried.append(wk)
                    if loc:
                        candidates.append(urljoin(wk, loc))
                except Exception as e2:
                    log.debug("well-known 탐색 실패: %r", e2)
    hint = ""
    if last_err is not None and ("Authoriz" in type(last_err).__name__ or "401" in str(last_err)):
        hint = ("\n  → 인증 실패(401). 네이버 2단계 인증을 켜고 '애플리케이션 비밀번호'를 새로 만들어 "
                "set-password 로 다시 저장하세요. (일반 로그인 비밀번호는 거부될 수 있음)")
    raise SystemExit(f"네이버 CalDAV 연결 실패: {last_err!r}{hint}")


def cal_display_name(cal) -> str:
    try:
        name = cal.get_display_name()
    except Exception:
        name = getattr(cal, "name", None)
    return str(name) if name else str(cal.url).rstrip("/").rsplit("/", 1)[-1]


def find_naver_calendar(cals: list, wanted: str, names: dict | None = None):
    names = names if names is not None else {}
    for c in cals:
        if wanted.startswith("http") and str(c.url).rstrip("/") == wanted.rstrip("/"):
            return c
    for c in cals:
        key = str(c.url)
        if key not in names:
            names[key] = cal_display_name(c)
        if names[key] == wanted:
            return c
    visible = ", ".join(repr(n) for n in names.values())
    raise LookupError(f"네이버 캘린더 '{wanted}' 를 찾지 못했습니다. CalDAV에 보이는 캘린더: {visible}")


def _object_data(obj) -> str | None:
    data = getattr(obj, "data", None)
    if not data:
        try:
            obj.load()
            data = obj.data
        except Exception as e:
            log.warning("일정 데이터 로드 실패 %s: %r", getattr(obj, "url", "?"), e)
            return None
    return data if isinstance(data, str) else data.decode("utf-8", "replace")


def fetch_naver_raw(cal, start: dt.datetime, end: dt.datetime, mode: str) -> list[str]:
    """CalDAV 리소스(한 UID = 마스터 + 예외 회차)를 원본 iCalendar 문자열로 가져온다."""
    objs = None
    if mode == "range":
        try:
            objs = cal.search(start=start, end=end, event=True, expand=False)
        except Exception as e:
            log.warning("기간 검색(REPORT time-range) 실패 → 전체 조회로 전환: %r", e)
    if objs is None:
        objs = cal.events()
    out = []
    for o in objs:
        d = _object_data(o)
        if d:
            out.append(d)
    return out


# ───────────────────────── 변환(순수 로직) ─────────────────────────

def _orig_start_token(value) -> str:
    if isinstance(value, dt.datetime):
        if value.tzinfo is None:
            return value.strftime("%Y%m%dT%H%M%S")
        return value.astimezone(dt.timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    return value.strftime("%Y%m%d")


def expand_instances(raw_items: Iterable[str], start: dt.datetime, end: dt.datetime) -> dict[str, Any]:
    """{안정 키: VEVENT 회차}. 단일 일정 키=UID, 반복 회차 키=UID|원래시작(UTC)."""
    import recurring_ical_events
    from icalendar import Calendar

    result: dict[str, Any] = {}
    for raw in raw_items:
        try:
            vcal = Calendar.from_ical(raw)
        except Exception as e:
            log.warning("iCalendar 파싱 실패: %r", e)
            continue
        vevents = list(vcal.walk("VEVENT"))
        if not vevents:
            continue  # VTODO(할 일) 등은 무시
        recurring = any(("RRULE" in c) or ("RDATE" in c) or ("RECURRENCE-ID" in c) for c in vevents)
        try:
            try:
                occs = recurring_ical_events.of(vcal, skip_bad_series=True).between(start, end)
            except TypeError:
                occs = recurring_ical_events.of(vcal).between(start, end)
        except Exception as e:
            log.warning("반복 일정 전개 실패 (UID=%s): %r", vevents[0].get("UID"), e)
            continue
        for occ in occs:
            if str(occ.get("STATUS", "")).upper() == "CANCELLED":
                continue
            uid = str(occ.get("UID", "") or "").strip()
            if not uid:
                uid = "nouid-" + _short_hash(f"{occ.get('DTSTART').to_ical()!r}{occ.get('SUMMARY')}", 16)
            if recurring:
                rid = occ.get("RECURRENCE-ID")
                orig = rid.dt if rid is not None else occ["DTSTART"].dt
                key = f"{uid}|{_orig_start_token(orig)}"
            else:
                key = uid
            result[key] = occ
    return result


def _tz_name(tzinfo) -> str | None:
    name = getattr(tzinfo, "key", None) or getattr(tzinfo, "zone", None)
    if not name:
        return None
    try:
        ZoneInfo(str(name))
        return str(name)
    except Exception:
        return None


def _aware(value: dt.datetime, tz: str) -> dt.datetime:
    return value.replace(tzinfo=ZoneInfo(tz)) if value.tzinfo is None else value


def _g_time(value, tz: str) -> dict:
    if isinstance(value, dt.datetime):
        value = _aware(value, tz).replace(microsecond=0)
        out = {"dateTime": value.isoformat()}
        name = _tz_name(value.tzinfo)
        if name:
            out["timeZone"] = name
        return out
    return {"date": value.isoformat()}


def _reminders(ev, mode: str) -> dict:
    if mode == "default":
        return {"useDefault": True}
    if mode == "source":
        minutes: list[int] = []
        for alarm in ev.walk("VALARM"):
            trig = alarm.get("TRIGGER")
            if trig is None or not isinstance(getattr(trig, "dt", None), dt.timedelta):
                continue
            if str(trig.params.get("RELATED", "START")).upper() != "START":
                continue
            m = int(-trig.dt.total_seconds() // 60)
            if 0 <= m <= 40320 and m not in minutes:
                minutes.append(m)
        if minutes:
            return {"useDefault": False,
                    "overrides": [{"method": "popup", "minutes": m} for m in sorted(minutes)[:5]]}
    return {"useDefault": False}


def event_id(mirror_id: str, key: str) -> str:
    """Google 이벤트 ID 규칙(base32hex 소문자 a-v,0-9 / 5~1024자)을 만족하는 결정적 ID."""
    return hashlib.sha1(f"{mirror_id}|{key}".encode("utf-8")).hexdigest()


def body_hash(body: dict) -> str:
    core = {k: v for k, v in body.items() if k != "extendedProperties"}
    return hashlib.sha1(json.dumps(core, sort_keys=True, ensure_ascii=False).encode("utf-8")).hexdigest()[:16]


def to_google_event(ev, mirror: MirrorCfg, tz: str) -> dict:
    start = ev["DTSTART"].dt
    if "DTEND" in ev:
        end = ev["DTEND"].dt
    elif "DURATION" in ev:
        end = start + ev["DURATION"].dt
    else:
        end = start if isinstance(start, dt.datetime) else start + DAY

    if isinstance(start, dt.datetime):          # 시간 지정 일정
        start = _aware(start, tz)
        if not isinstance(end, dt.datetime):
            end = dt.datetime.combine(end, dt.time(), tzinfo=start.tzinfo)
        end = _aware(end, tz)
        if end < start:
            end = start
    else:                                        # 종일 일정
        if isinstance(end, dt.datetime):
            end = end.date()
        if end <= start:
            end = start + DAY

    summary = str(ev.get("SUMMARY", "") or "").strip() or "(제목 없음)"
    parts = []
    desc = str(ev.get("DESCRIPTION", "") or "").strip()
    if desc and mirror.copy_description:
        parts.append(desc)
    url = str(ev.get("URL", "") or "").strip()
    if url:
        parts.append(url)
    parts.append(f"[네이버 캘린더 '{mirror.naver}'에서 자동 동기화된 일정 — 수정은 네이버 캘린더에서 하세요]")
    klass = str(ev.get("CLASS", "") or "").upper()

    body: dict[str, Any] = {
        "summary": (mirror.title_prefix + summary)[:1000],
        "description": "\n\n".join(parts)[:8000],
        "start": _g_time(start, tz),
        "end": _g_time(end, tz),
        "status": "confirmed",
        "transparency": "transparent" if str(ev.get("TRANSP", "")).upper() == "TRANSPARENT" else "opaque",
        "visibility": "private" if klass in ("PRIVATE", "CONFIDENTIAL") else "default",
        "reminders": _reminders(ev, mirror.reminders),
        "source": {"title": "Naver Calendar", "url": "https://calendar.naver.com/"},
    }
    location = str(ev.get("LOCATION", "") or "").strip()
    if location:
        body["location"] = location[:1000]
    if mirror.color_id:
        body["colorId"] = mirror.color_id
    body["extendedProperties"] = {"private": {"n2g": "1", "n2gM": mirror.id, "n2gH": body_hash(body)}}
    return body


# ───────────────────────── 동기화 계획(순수 로직) ─────────────────────────

@dataclass
class Plan:
    inserts: list[tuple[str, dict]]
    updates: list[tuple[str, dict]]
    deletes: list[str]
    unchanged: int
    guard_msg: str | None = None

    def summary(self) -> str:
        s = f"추가 {len(self.inserts)} · 변경 {len(self.updates)} · 삭제 {len(self.deletes)} · 동일 {self.unchanged}"
        return s + (f" · ⚠ {self.guard_msg}" if self.guard_msg else "")


def plan_changes(desired: dict[str, dict], existing: dict[str, dict],
                 del_min: dt.datetime, del_max: dt.datetime,
                 guard_ratio: float = 0.5, force: bool = False,
                 protected: set[str] | None = None) -> Plan:
    protected = protected or set()
    inserts, updates, unchanged = [], [], 0
    for gid, body in desired.items():
        ex = existing.get(gid)
        if ex is None:
            inserts.append((gid, body))
        elif ex.get("hash") != body["extendedProperties"]["private"]["n2gH"]:
            updates.append((gid, body))
        else:
            unchanged += 1
    deletes = [gid for gid, ex in existing.items()
               if gid not in desired and gid not in protected
               and ex.get("start") is not None and del_min <= ex["start"] <= del_max]
    guard = None
    if deletes and not force:
        if not desired and len(existing) >= 5:
            guard = "네이버 조회 결과가 0건이라 삭제를 보류했습니다(일시 오류 가능). 실제로 비운 것이면 --force"
        elif len(deletes) > max(10, guard_ratio * len(existing)):
            guard = f"삭제 {len(deletes)}건이 기존 {len(existing)}건 대비 과다하여 보류했습니다. 확인 후 --force"
        if guard:
            deletes = []
    return Plan(inserts, updates, deletes, unchanged, guard)


# ───────────────────────── Google Calendar ─────────────────────────

def _save_token(cfg: Config, creds) -> None:
    cfg.google_token.parent.mkdir(parents=True, exist_ok=True)
    cfg.google_token.write_text(creds.to_json(), encoding="utf-8")


def google_credentials(cfg: Config, interactive: bool = False):
    if cfg.google_auth == "service_account":
        from google.oauth2 import service_account
        sa_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")  # 클라우드 실행용 (파일 대신)
        if sa_json:
            import json as _json
            return service_account.Credentials.from_service_account_info(_json.loads(sa_json), scopes=SCOPES)
        return service_account.Credentials.from_service_account_file(str(cfg.google_credentials), scopes=SCOPES)

    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials

    creds = None
    if cfg.google_token.exists():
        creds = Credentials.from_authorized_user_file(str(cfg.google_token), SCOPES)
    if creds and creds.valid:
        return creds
    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
            _save_token(cfg, creds)
            return creds
        except Exception as e:
            log.warning("Google 토큰 갱신 실패: %r", e)
    if not interactive:
        raise SystemExit("Google 인증이 필요합니다:  python naver2gcal.py auth\n"
                         "(OAuth 동의화면이 '테스트' 상태면 7일마다 만료 → '프로덕션'으로 게시하세요)")
    if not cfg.google_credentials.exists():
        raise SystemExit(f"OAuth 클라이언트 파일이 없습니다: {cfg.google_credentials}")
    from google_auth_oauthlib.flow import InstalledAppFlow
    flow = InstalledAppFlow.from_client_secrets_file(str(cfg.google_credentials), SCOPES)
    creds = flow.run_local_server(port=0, open_browser=True)
    _save_token(cfg, creds)
    return creds


def google_service(cfg: Config, interactive: bool = False):
    from googleapiclient.discovery import build
    return build("calendar", "v3", credentials=google_credentials(cfg, interactive), cache_discovery=False)


def _http_status(err) -> int | None:
    resp = getattr(err, "resp", None)
    try:
        return int(getattr(resp, "status", None))
    except (TypeError, ValueError):
        return None


def _parse_gstart(start: dict, tz: str) -> dt.datetime | None:
    if "dateTime" in start:
        return dt.datetime.fromisoformat(start["dateTime"].replace("Z", "+00:00"))
    if "date" in start:
        return dt.datetime.combine(dt.date.fromisoformat(start["date"]), dt.time(), tzinfo=ZoneInfo(tz))
    return None


class GoogleOps:
    """Google Calendar API 호출 래퍼 (테스트에서 가짜 구현으로 대체 가능)."""

    def __init__(self, svc):
        self.svc = svc

    def resolve_calendar(self, target: str, tz: str, allow_create: bool) -> str | None:
        if "@" in target or target == "primary":
            return target
        token = None
        while True:
            resp = self.svc.calendarList().list(pageToken=token, minAccessRole="writer").execute(num_retries=3)
            for item in resp.get("items", []):
                if target in (item.get("summary"), item.get("summaryOverride")):
                    return item["id"]
            token = resp.get("nextPageToken")
            if not token:
                break
        if not allow_create:
            return None
        cal = self.svc.calendars().insert(body={
            "summary": target, "timeZone": tz,
            "description": "네이버 캘린더 단방향 미러 (naver2gcal). 이 캘린더의 일정은 자동으로 덮어써집니다.",
        }).execute(num_retries=3)
        log.info("Google 캘린더 생성: %s (%s)", target, cal["id"])
        return cal["id"]

    def list_mirror_events(self, cal_id: str, mirror_id: str, tmin: dt.datetime, tmax: dt.datetime,
                           tz: str) -> dict[str, dict]:
        out: dict[str, dict] = {}
        token = None
        while True:
            resp = self.svc.events().list(
                calendarId=cal_id, privateExtendedProperty=f"n2gM={mirror_id}",
                timeMin=tmin.isoformat(), timeMax=tmax.isoformat(),
                singleEvents=True, showDeleted=False, maxResults=2500, pageToken=token,
                fields="items(id,status,start,extendedProperties),nextPageToken",
            ).execute(num_retries=5)
            for it in resp.get("items", []):
                if it.get("status") == "cancelled":
                    continue
                priv = (it.get("extendedProperties") or {}).get("private") or {}
                out[it["id"]] = {"hash": priv.get("n2gH"), "start": _parse_gstart(it.get("start") or {}, tz)}
            token = resp.get("nextPageToken")
            if not token:
                return out

    def insert(self, cal_id: str, gid: str, body: dict) -> None:
        from googleapiclient.errors import HttpError
        try:
            self.svc.events().insert(calendarId=cal_id, body=dict(body, id=gid),
                                     sendUpdates="none").execute(num_retries=5)
        except HttpError as e:
            if _http_status(e) == 409:  # 같은 ID가 (삭제 상태로) 이미 존재 → 갱신으로 되살림
                self.update(cal_id, gid, body)
            else:
                raise

    def update(self, cal_id: str, gid: str, body: dict) -> None:
        self.svc.events().update(calendarId=cal_id, eventId=gid, body=body,
                                 sendUpdates="none").execute(num_retries=5)

    def delete(self, cal_id: str, gid: str) -> None:
        from googleapiclient.errors import HttpError
        try:
            self.svc.events().delete(calendarId=cal_id, eventId=gid, sendUpdates="none").execute(num_retries=5)
        except HttpError as e:
            if _http_status(e) not in (404, 410):
                raise


def apply_plan(gops, cal_id: str, plan: Plan) -> int:
    errors = 0
    for label, items, fn in (
        ("추가", plan.inserts, lambda g, b: gops.insert(cal_id, g, b)),
        ("변경", plan.updates, lambda g, b: gops.update(cal_id, g, b)),
        ("삭제", [(g, None) for g in plan.deletes], lambda g, b: gops.delete(cal_id, g)),
    ):
        for gid, body in items:
            try:
                fn(gid, body)
            except Exception as e:
                errors += 1
                title = body.get("summary") if body else gid
                log.error("%s 실패 [%s]: %r", label, title, e)
    return errors


def mirror_window(cfg: Config, mirror: MirrorCfg, now: dt.datetime) -> tuple[dt.datetime, dt.datetime]:
    past = cfg.past_days if mirror.past_days is None else mirror.past_days
    future = cfg.future_days if mirror.future_days is None else mirror.future_days
    return now - past * DAY, now + future * DAY


def sync_mirror(cfg: Config, mirror: MirrorCfg, raw_items: list[str], gops, cal_id: str | None,
                now: dt.datetime, dry_run: bool = False, force: bool = False) -> tuple[Plan, int]:
    tz = cfg.timezone
    start, end = mirror_window(cfg, mirror, now)
    desired: dict[str, dict] = {}
    protected: set[str] = set()
    for key, ev in expand_instances(raw_items, start, end).items():
        gid = event_id(mirror.id, key)
        try:
            desired[gid] = to_google_event(ev, mirror, tz)
        except Exception as e:
            protected.add(gid)  # 변환 실패한 일정은 기존 사본을 지우지 않는다
            log.warning("변환 실패 [%s]: %r", ev.get("SUMMARY"), e)
    existing = gops.list_mirror_events(cal_id, mirror.id, start, end, tz) if cal_id else {}
    plan = plan_changes(desired, existing, start + DAY, end - DAY,
                        cfg.delete_guard_ratio, force, protected)
    if dry_run or not cal_id:
        return plan, 0
    return plan, apply_plan(gops, cal_id, plan)


# ───────────────────────── ICS 내보내기(이관) ─────────────────────────

def _strip_people(ev) -> None:
    """ORGANIZER/ATTENDEE 제거 → Google에 '초대받은 일정'이 아닌 '내 일정'(편집 가능)으로 들어가게 함.
    참석자 이름은 설명 끝에 텍스트로 남긴다."""
    attendees = ev.get("ATTENDEE")
    if attendees is not None:
        if not isinstance(attendees, list):
            attendees = [attendees]
        names = []
        for a in attendees:
            cn = getattr(a, "params", {}).get("CN")
            names.append(str(cn) if cn else re.sub(r"(?i)^mailto:", "", str(a)))
        desc = str(ev.get("DESCRIPTION", "") or "").strip()
        ev.pop("DESCRIPTION", None)
        ev.add("DESCRIPTION", (desc + "\n\n" if desc else "") + "참석자: " + ", ".join(names))
    ev.pop("ATTENDEE", None)
    ev.pop("ORGANIZER", None)


def build_ics_chunks(cal_name: str, raw_items: Iterable[str], max_bytes: int = 900_000,
                     keep_people: bool = False) -> list[bytes]:
    """VEVENT를 UID 단위(마스터+예외 회차)로 묶어 max_bytes 이하 .ics 들로 분할."""
    from icalendar import Calendar

    tzs: dict[str, Any] = {}
    groups: dict[str, list] = {}
    for raw in raw_items:
        try:
            c = Calendar.from_ical(raw)
        except Exception as e:
            log.warning("iCalendar 파싱 실패: %r", e)
            continue
        for sub in c.subcomponents:
            if sub.name == "VTIMEZONE":
                tzs.setdefault(str(sub.get("TZID")), sub)
            elif sub.name == "VEVENT":
                if not keep_people:
                    _strip_people(sub)
                uid = str(sub.get("UID", "") or "") or f"nouid-{len(groups)}"
                groups.setdefault(uid, []).append(sub)

    def new_cal():
        cal = Calendar()
        cal.add("prodid", "-//naver2gcal//KO")
        cal.add("version", "2.0")
        cal.add("calscale", "GREGORIAN")
        cal.add("x-wr-calname", cal_name)
        for t in tzs.values():
            cal.add_component(t)
        return cal

    base = len(new_cal().to_ical())
    chunks: list[bytes] = []
    cur, size, count = new_cal(), base, 0
    for evs in groups.values():
        gsize = sum(len(e.to_ical()) for e in evs)
        if count and size + gsize > max_bytes:
            chunks.append(cur.to_ical())
            cur, size, count = new_cal(), base, 0
        for e in evs:
            cur.add_component(e)
        size += gsize
        count += 1
    if count:
        chunks.append(cur.to_ical())
    return chunks


# ───────────────────── 본인 로그인(OAuth): Tasks 이관 · 중복 정리 ─────────────────────

def user_credentials(cfg: Config, interactive: bool = False):
    """Google Tasks 와 기본 캘린더는 서비스 계정으로 접근할 수 없으므로 본인 OAuth 로그인을 쓴다.
    (앱이 '테스트' 상태라 refresh token 은 7일 뒤 만료 → 일회성 작업용)"""
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials

    creds = None
    if cfg.user_token.exists():
        creds = Credentials.from_authorized_user_file(str(cfg.user_token), USER_SCOPES)
    if creds and creds.valid and creds.has_scopes(USER_SCOPES):
        return creds
    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
            cfg.user_token.write_text(creds.to_json(), encoding="utf-8")
            return creds
        except Exception as e:
            log.warning("본인 로그인 토큰 갱신 실패: %r", e)
    if not interactive:
        raise SystemExit("본인 구글 로그인이 필요합니다:  python naver2gcal.py user-auth")
    if not cfg.user_credentials.exists():
        raise SystemExit(f"OAuth 클라이언트 파일이 없습니다: {cfg.user_credentials}")
    from google_auth_oauthlib.flow import InstalledAppFlow
    flow = InstalledAppFlow.from_client_secrets_file(str(cfg.user_credentials), USER_SCOPES)
    creds = flow.run_local_server(port=0, open_browser=True,
                                  success_message="naver2gcal 인증이 끝났습니다. 이 창을 닫아도 됩니다.")
    cfg.user_token.write_text(creds.to_json(), encoding="utf-8")
    return creds


_FREQ_KO = {"DAILY": ("매일", "일"), "WEEKLY": ("매주", "주"), "MONTHLY": ("매월", "개월"), "YEARLY": ("매년", "년")}
_DAY_KO = {"MO": "월", "TU": "화", "WE": "수", "TH": "목", "FR": "금", "SA": "토", "SU": "일"}


def repeat_text(rule) -> str:
    """RRULE → '매월 4일', '2주마다 월,수' 같은 한국어 설명."""
    freq = str((rule.get("FREQ") or ["?"])[0]).upper()
    interval = int((rule.get("INTERVAL") or [1])[0])
    every, unit = _FREQ_KO.get(freq, (freq, ""))
    text = every if interval == 1 else f"{interval}{unit}마다"
    if rule.get("BYDAY"):
        text += " " + ",".join(_DAY_KO.get(str(d)[-2:].upper(), str(d)) for d in rule["BYDAY"])
    if rule.get("BYMONTHDAY"):
        text += " " + ",".join(f"{d}일" for d in rule["BYMONTHDAY"])
    if rule.get("BYMONTH") and freq == "YEARLY":
        text += " (" + ",".join(f"{m}월" for m in rule["BYMONTH"]) + ")"
    return text


def _as_date(value, tz: str) -> dt.date:
    if isinstance(value, dt.datetime):
        return (_aware(value, tz).astimezone(ZoneInfo(tz))).date()
    return value


def next_occurrence(todo, today: dt.date, tz: str) -> dt.date | None:
    """반복 할 일의 '오늘 이후 첫 회차' 마감일. 반복이 끝났으면 마지막 회차."""
    from dateutil.rrule import rrulestr
    due = todo.get("DUE")
    if due is None or "RRULE" not in todo:
        return None
    start = dt.datetime.combine(_as_date(due.dt, tz), dt.time())
    rule = todo["RRULE"].to_ical().decode()
    rule = re.sub(r"(UNTIL=\d{8}(T\d{6})?)Z", r"\1", rule)   # 시작값이 시간대 없는 값이라 UNTIL 도 맞춘다
    try:
        rr = rrulestr(rule, dtstart=start)
        nxt = rr.after(dt.datetime.combine(today, dt.time()), inc=True)
        if nxt is None:
            prev = rr.before(dt.datetime.combine(today, dt.time()), inc=True)
            return prev.date() if prev else start.date()
        return nxt.date()
    except Exception as e:
        log.debug("반복 계산 실패 %s: %r", todo.get("SUMMARY"), e)
        return start.date()


def todo_to_task(todo, tz: str, today: dt.date) -> dict:
    """네이버 VTODO → Google Tasks task 본문."""
    title = str(todo.get("SUMMARY", "") or "").strip() or "(제목 없음)"
    notes = []
    desc = str(todo.get("DESCRIPTION", "") or "").strip()
    if desc:
        notes.append(desc)
    completed = str(todo.get("STATUS", "")).upper() == "COMPLETED"
    due_date = _as_date(todo["DUE"].dt, tz) if todo.get("DUE") is not None else None
    if "RRULE" in todo:
        notes.append(f"🔁 반복(네이버): {repeat_text(todo['RRULE'])} — 구글 Tasks 앱에서 반복을 다시 설정할 수 있습니다")
        if not completed:
            due_date = next_occurrence(todo, today, tz) or due_date
    try:
        if 1 <= int(str(todo.get("PRIORITY", 0))) <= 4:
            notes.append("★ 중요(네이버 우선순위)")
    except ValueError:
        pass
    body: dict[str, Any] = {"title": title[:1024], "status": "completed" if completed else "needsAction"}
    if notes:
        body["notes"] = "\n\n".join(notes)[:8000]
    if due_date:
        body["due"] = f"{due_date.isoformat()}T00:00:00.000Z"
    if completed:
        when = todo.get("COMPLETED") or todo.get("LAST-MODIFIED")
        if when is not None and isinstance(when.dt, dt.datetime):
            body["completed"] = _aware(when.dt, tz).astimezone(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.000Z")
    return body


def collect_todos(raw_items: Iterable[str]) -> list:
    from icalendar import Calendar
    todos = []
    for raw in raw_items:
        try:
            todos.extend(Calendar.from_ical(raw).walk("VTODO"))
        except Exception as e:
            log.warning("할 일 파싱 실패: %r", e)
    return todos


def cmd_user_auth(cfg: Config, args) -> int:
    from googleapiclient.discovery import build
    creds = user_credentials(cfg, interactive=True)
    lists = build("tasks", "v1", credentials=creds, cache_discovery=False).tasklists().list().execute()
    print("본인 구글 로그인 성공. 현재 Tasks 목록:", ", ".join(t["title"] for t in lists.get("items", [])) or "(없음)")
    print("※ 이 로그인은 앱이 '테스트' 상태라 7일 뒤 만료됩니다(일회성 작업용).")
    return 0


def cmd_tasks_import(cfg: Config, args) -> int:
    """네이버 '내 할 일' → Google Tasks 목록으로 한 번 옮기기."""
    _, principal = naver_principal(cfg)
    cal = find_naver_calendar(principal.calendars(), args.naver)
    raw = [d for d in (_object_data(o) for o in cal.todos(include_completed=True)) if d]
    todos = collect_todos(raw)
    if not args.include_completed:
        todos = [t for t in todos if str(t.get("STATUS", "")).upper() != "COMPLETED"]
    today = now_in(cfg.timezone).date()
    bodies = [todo_to_task(t, cfg.timezone, today) for t in todos]
    # 새 항목은 목록 맨 위에 들어가므로, 마감일이 늦은 것부터 넣어 '가장 급한 일'이 맨 위에 오게 한다
    bodies.sort(key=lambda b: b.get("due", "9999"), reverse=True)
    open_n = sum(b["status"] == "needsAction" for b in bodies)
    rec_n = sum("🔁" in b.get("notes", "") for b in bodies)
    print(f"네이버 '{args.naver}': 옮길 할 일 {len(bodies)}건 (미완료 {open_n} · 완료 {len(bodies) - open_n} · 반복 {rec_n})")
    for b in [b for b in bodies if b["status"] == "needsAction"][-10:][::-1]:
        print(f"   {b.get('due', '')[:10] or '마감 없음':<10}  {b['title'][:50]}")
    if args.dry_run:
        print("[DRY-RUN] Google Tasks 에는 아무것도 쓰지 않았습니다.")
        return 0

    from googleapiclient.discovery import build
    svc = build("tasks", "v1", credentials=user_credentials(cfg), cache_discovery=False)
    lists = svc.tasklists().list(maxResults=100).execute().get("items", [])
    if args.list == "@default":
        target = svc.tasklists().get(tasklist="@default").execute()
        args.force = True   # 기본 목록에는 원래 항목이 있을 수 있으므로 비어 있지 않아도 진행
    else:
        target = next((l for l in lists if l["title"] == args.list), None)
    if target is None:
        target = svc.tasklists().insert(body={"title": args.list}).execute()
        print(f"Google Tasks 목록 생성: {args.list}")
    else:
        existing = svc.tasks().list(tasklist=target["id"], maxResults=1, showCompleted=True, showHidden=True).execute()
        if existing.get("items") and not args.force:
            raise SystemExit(f"Google Tasks 목록 '{args.list}' 에 이미 항목이 있습니다. 중복을 막기 위해 중단합니다 (--force 로 무시).")
    errors = 0
    for i, b in enumerate(bodies, 1):
        try:
            svc.tasks().insert(tasklist=target["id"], body=b).execute(num_retries=5)
        except Exception as e:
            errors += 1
            log.error("할 일 추가 실패 [%s]: %r", b["title"], e)
        if i % 100 == 0:
            print(f"   … {i}/{len(bodies)}")
    print(f"완료: {len(bodies) - errors}건 추가, 실패 {errors}건 → Google Tasks '{args.list}'")
    return 1 if errors else 0


def cmd_cleanup_imported(cfg: Config, args) -> int:
    """앞서 .ics 로 기본 캘린더에 가져온 일정 중, 그 파일의 UID 와 정확히 일치하는 것만 삭제(휴지통으로)."""
    from icalendar import Calendar
    from googleapiclient.discovery import build
    uids: set[str] = set()
    for f in args.ics:
        path = Path(f) if Path(f).is_absolute() else BASE_DIR / f
        uids |= {str(e["UID"]) for e in Calendar.from_ical(path.read_bytes()).walk("VEVENT") if e.get("UID")}
    svc = build("calendar", "v3", credentials=user_credentials(cfg), cache_discovery=False)
    targets, token = [], None
    while True:
        resp = svc.events().list(calendarId=args.calendar, singleEvents=False, showDeleted=False, maxResults=2500,
                                 pageToken=token, fields="items(id,iCalUID,summary,start),nextPageToken"
                                 ).execute(num_retries=5)
        targets += [e for e in resp.get("items", []) if e.get("iCalUID") in uids]
        token = resp.get("nextPageToken")
        if not token:
            break
    print(f"파일의 일정 UID {len(uids)}개 중, '{args.calendar}' 캘린더에서 일치하는 일정 {len(targets)}건")
    for e in targets[:5]:
        s = e.get("start", {})
        print(f"   {(s.get('dateTime') or s.get('date') or '')[:10]}  {e.get('summary', '')[:50]}")
    if len(targets) > len(uids):
        raise SystemExit("일치 건수가 파일 일정 수보다 많아 중단합니다(예상 밖 상황).")
    if args.dry_run:
        print("[DRY-RUN] 삭제하지 않았습니다.")
        return 0
    errors = 0
    for e in targets:
        try:
            svc.events().delete(calendarId=args.calendar, eventId=e["id"], sendUpdates="none").execute(num_retries=5)
        except Exception as ex:
            if _http_status(ex) not in (404, 410):
                errors += 1
                log.error("삭제 실패 [%s]: %r", e.get("summary"), ex)
    print(f"삭제 완료: {len(targets) - errors}건 (구글 캘린더 휴지통에서 30일간 복구 가능), 실패 {errors}건")
    return 1 if errors else 0


def _safe_filename(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|\s]+', "_", name).strip("_") or "calendar"


# ─────────────────────────────── 명령 ───────────────────────────────

def now_in(tz: str) -> dt.datetime:
    return dt.datetime.now(ZoneInfo(tz)).replace(microsecond=0)


def cmd_set_password(cfg: Config, args) -> int:
    import keyring
    pw = getpass.getpass(f"네이버 '{cfg.naver_username}' 애플리케이션 비밀번호 (입력 내용은 표시되지 않음): ").strip()
    if not pw:
        raise SystemExit("빈 값은 저장하지 않습니다.")
    keyring.set_password(APP, cfg.naver_username, pw)
    print("저장 완료 (Windows 자격 증명 관리자 → 일반 자격 증명 → naver2gcal).")
    return 0


def cmd_discover(cfg: Config, args) -> int:
    _, principal = naver_principal(cfg)
    cals = principal.calendars()
    now = now_in(cfg.timezone)
    start, end = now - 30 * DAY, now + 90 * DAY
    print(f"CalDAV principal: {principal.url}")
    print(f"보이는 캘린더 {len(cals)}개 (일정 수 = 최근 30일 ~ 향후 90일, 반복 일정은 1건으로 셈)\n")
    for c in cals:
        name = cal_display_name(c)
        try:
            comps = ",".join(c.get_supported_components())
        except Exception:
            comps = "?"
        try:
            n: Any = len(c.search(start=start, end=end, event=True, expand=False))
        except Exception as e:
            n = f"조회 실패({type(e).__name__})"
        print(f"• {name}\n    일정 {n}건 · 구성요소 {comps}\n    {c.url}")
    print("\n팀 공유 캘린더가 위 목록에 있으면 → config.toml 의 [[mirror]] naver = \"이름\" 으로 등록해 미러할 수 있습니다.")
    print("목록에 없으면 → 네이버가 CalDAV로 공유 캘린더를 내주지 않는 것이므로 README의 'NO-GO 대안'을 보세요.")
    return 0


def cmd_auth(cfg: Config, args) -> int:
    svc = google_service(cfg, interactive=True)
    items = svc.calendarList().list(minAccessRole="writer").execute().get("items", [])
    print("Google 인증 성공. 쓰기 가능한 캘린더:")
    for it in items:
        print(f"  - {it.get('summaryOverride') or it.get('summary')}  ({it['id']})")
    return 0


def cmd_export(cfg: Config, args) -> int:
    _, principal = naver_principal(cfg)
    cals = principal.calendars()
    names: dict = {}
    selected = [find_naver_calendar(cals, n, names) for n in args.calendar] if args.calendar else cals
    out = Path(args.out)
    out = out if out.is_absolute() else BASE_DIR / out
    out.mkdir(parents=True, exist_ok=True)
    for cal in selected:
        name = cal_display_name(cal)
        raw = fetch_naver_raw(cal, None, None, "all")
        chunks = build_ics_chunks(name, raw, args.max_kb * 1000, args.keep_attendees)
        for i, data in enumerate(chunks, 1):
            f = out / f"{_safe_filename(name)}_{i:02d}.ics"
            f.write_bytes(data)
            print(f"  {f}  ({len(data) // 1024} KB)")
        print(f"• {name}: 원본 {len(raw)}건 → 파일 {len(chunks)}개")
    print("\nGoogle 캘린더 → 설정 → 가져오기/내보내기 → 가져오기 에서 파일별로 대상 캘린더를 골라 올리세요.")
    return 0


def cmd_sync(cfg: Config, args) -> int:
    if not cfg.mirrors:
        raise SystemExit("config.toml 에 [[mirror]] 항목이 없습니다.")
    _, principal = naver_principal(cfg)
    cals = principal.calendars()
    names: dict = {}
    gops = GoogleOps(google_service(cfg))
    now = now_in(cfg.timezone)
    rc = 0
    for m in cfg.mirrors:
        if args.only and m.naver not in args.only and m.google not in args.only and m.id not in args.only:
            continue
        try:
            ncal = find_naver_calendar(cals, m.naver, names)
            allow_create = cfg.google_auth == "oauth" and not args.dry_run
            cal_id = gops.resolve_calendar(m.google, cfg.timezone, allow_create)
            if cal_id is None and not args.dry_run:
                raise LookupError(f"Google 캘린더 '{m.google}' 없음 (서비스 계정은 캘린더 ID를 지정하세요)")
            raw = fetch_naver_raw(ncal, *mirror_window(cfg, m, now), cfg.fetch_mode)
            plan, errors = sync_mirror(cfg, m, raw, gops, cal_id, now, args.dry_run, args.force)
            tag = "[DRY-RUN] " if args.dry_run else ""
            log.info("%s%s → %s : %s%s", tag, m.naver, m.google, plan.summary(),
                     f" · 오류 {errors}" if errors else "")
            if args.dry_run:
                for gid, b in (plan.inserts + plan.updates)[:15]:
                    log.info("    %s %s", b["start"].get("dateTime", b["start"].get("date")), b["summary"])
            if errors or plan.guard_msg:
                rc = 1
        except Exception as e:
            rc = 1
            log.error("미러 실패 [%s → %s]: %s", m.naver, m.google, e, exc_info=args.verbose)
    return rc


# ─────────────────────────────── main ───────────────────────────────

def _fix_std_streams() -> None:
    # pythonw.exe(창 없이 실행)에서는 stdout/stderr 가 None
    for name in ("stdout", "stderr"):
        if getattr(sys, name) is None:
            setattr(sys, name, open(os.devnull, "w", encoding="utf-8"))
        else:
            try:
                getattr(sys, name).reconfigure(errors="replace")
            except Exception:
                pass


def setup_logging(cfg: Config | None, verbose: bool) -> None:
    log.setLevel(logging.DEBUG if verbose else logging.INFO)
    log.propagate = False
    # 네이버 VTODO(할 일)에 DTSTAMP가 빠져 caldav 라이브러리가 경고를 대량 출력함 — 동작에는 무해
    logging.getLogger("caldav").setLevel(logging.DEBUG if verbose else logging.ERROR)
    for h in list(log.handlers):  # 재호출 시 핸들러 중복 방지
        log.removeHandler(h)
        h.close()
    fmt =logging.Formatter("%(asctime)s %(levelname)s %(message)s", "%Y-%m-%d %H:%M:%S")
    sh = logging.StreamHandler(sys.stderr)
    sh.setFormatter(fmt)
    log.addHandler(sh)
    if cfg and cfg.log_file:
        cfg.log_file.parent.mkdir(parents=True, exist_ok=True)
        fh = logging.handlers.RotatingFileHandler(cfg.log_file, maxBytes=1_000_000, backupCount=3, encoding="utf-8")
        fh.setFormatter(fmt)
        log.addHandler(fh)


def main(argv: list[str] | None = None) -> int:
    _fix_std_streams()
    ap = argparse.ArgumentParser(prog=APP, description="네이버 캘린더 → Google 캘린더 미러/이관")
    ap.add_argument("--config", default=str(BASE_DIR / "config.toml"))
    ap.add_argument("-v", "--verbose", action="store_true")
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("set-password", help="네이버 애플리케이션 비밀번호 저장")
    sub.add_parser("discover", help="CalDAV로 보이는 네이버 캘린더 목록")
    sub.add_parser("auth", help="Google OAuth 최초 인증")
    pe = sub.add_parser("export", help="네이버 캘린더를 .ics 로 내보내기(이관용)")
    pe.add_argument("--calendar", action="append", help="캘린더 이름(여러 번 지정 가능, 생략 시 전체)")
    pe.add_argument("--out", default="export")
    pe.add_argument("--max-kb", type=int, default=900, help="파일당 최대 크기(KB)")
    pe.add_argument("--keep-attendees", action="store_true",
                    help="주최자/참석자 정보 유지 (기본: 제거해서 Google에서 내 일정으로 편집 가능하게)")
    ps = sub.add_parser("sync", help="네이버 → Google 단방향 미러")
    ps.add_argument("--dry-run", action="store_true", help="Google에 쓰지 않고 변경 예정만 표시")
    ps.add_argument("--force", action="store_true", help="대량 삭제 보호 해제")
    ps.add_argument("--only", action="append", help="특정 미러만 (naver/google 이름 또는 id)")
    sub.add_parser("user-auth", help="본인 구글 로그인 (Tasks 이관·중복 정리용, 7일 유효)")
    pt = sub.add_parser("tasks-import", help="네이버 할 일 → Google Tasks (한 번 옮기기)")
    pt.add_argument("--naver", default="내 할 일", help="네이버 할 일 캘린더 이름")
    pt.add_argument("--list", default="@default", help="Google Tasks 목록 이름 (기본: 기본 목록 — Gmail·캘린더 옆 패널에 보이는 목록)")
    pt.add_argument("--include-completed", action="store_true", help="완료된 할 일도 옮기기")
    pt.add_argument("--dry-run", action="store_true")
    pt.add_argument("--force", action="store_true", help="대상 목록에 항목이 있어도 진행")
    pc = sub.add_parser("cleanup-imported", help=".ics 로 가져온 일정을 기본 캘린더에서 삭제(중복 정리)")
    pc.add_argument("--ics", action="append", default=None, help="가져왔던 .ics 파일 (기본: export/내_캘린더_01.ics)")
    pc.add_argument("--calendar", default="primary", help="정리할 구글 캘린더 (기본: 기본 캘린더)")
    pc.add_argument("--dry-run", action="store_true")
    args = ap.parse_args(argv)
    if args.cmd == "cleanup-imported" and not args.ics:
        args.ics = ["export/내_캘린더_01.ics"]

    cfg = load_config(Path(args.config))
    setup_logging(cfg, args.verbose)
    handlers = {"set-password": cmd_set_password, "discover": cmd_discover, "auth": cmd_auth,
                "export": cmd_export, "sync": cmd_sync, "user-auth": cmd_user_auth,
                "tasks-import": cmd_tasks_import, "cleanup-imported": cmd_cleanup_imported}
    try:
        return handlers[args.cmd](cfg, args) or 0
    except (SystemExit, KeyboardInterrupt):
        raise
    except Exception as e:
        log.error("실패: %r", e, exc_info=True)
        return 1


if __name__ == "__main__":
    sys.exit(main())
