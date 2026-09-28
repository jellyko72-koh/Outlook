#!/usr/bin/env python3
"""
지멘스 공문 자동 등록·요약 스크립트

모바일(Android Termux / iOS Pyto 등)에서 실행하면, 모바일 Outlook 앱과 같은
Microsoft 365/Outlook.com 사서함을 Microsoft Graph API로 조회하여

  1) 최근 N일(기본 3일)간 수신된 지멘스(Siemens) 공문 메일을 찾고
  2) 공문 첨부파일을 우선 추출(다운로드)한 뒤
  3) '지멘스 대응' 프로젝트(Microsoft To Do 목록 / Planner 플랜 + OneDrive 폴더)에 등록하고
  4) Claude로 첨부 공문 내용을 한국어로 요약하여 프로젝트 작업에 기록합니다.

사용법:
    python siemens_letters.py               # 최근 3일
    python siemens_letters.py --days 7      # 최근 7일
    python siemens_letters.py --dry-run     # 등록 없이 조회·요약만
"""

from __future__ import annotations

import argparse
import base64
import datetime as dt
import html
import io
import json
import re
import sys
import time
import uuid
import urllib.parse
import webbrowser
import zipfile
from pathlib import Path
from xml.etree import ElementTree

import anthropic
import msal
import requests

BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "config.json"
TOKEN_CACHE_PATH = BASE_DIR / ".token_cache.json"
GRAPH = "https://graph.microsoft.com/v1.0"
SCOPES = ["User.Read", "Mail.Read", "Tasks.ReadWrite", "Files.ReadWrite"]

DEFAULT_CONFIG = {
    "azure_client_id": "",
    "azure_authority": "https://login.microsoftonline.com/common",
    "anthropic_api_key": "",
    "claude_model": "claude-opus-5",
    "project_name": "지멘스 대응",
    # todo | planner | local
    "project_backend": "todo",
    "planner_plan_id": "",
    "planner_bucket_name": "공문 접수",
    "upload_to_onedrive": True,
    "output_dir": "output",
    "days": 3,
    "sender_domain_keywords": ["siemens"],
    "text_keywords": ["지멘스", "siemens"],
    "letter_keywords": ["공문", "letter", "notice", "notification", "claim", "통지", "요청", "회신"],
    "only_with_attachments": False,
}

# 첨부 중 이 크기를 넘으면 Claude 요약에서는 제외(파일 저장·등록은 그대로 진행)
MAX_SUMMARY_ATTACHMENT_BYTES = 25 * 1024 * 1024
TODO_ATTACHMENT_LIMIT = 3 * 1024 * 1024
ONEDRIVE_SIMPLE_UPLOAD_LIMIT = 4 * 1024 * 1024
IMAGE_TYPES = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
               ".gif": "image/gif", ".webp": "image/webp"}


# ---------------------------------------------------------------------------
# 설정 / 인증
# ---------------------------------------------------------------------------

def load_config() -> dict:
    cfg = dict(DEFAULT_CONFIG)
    if CONFIG_PATH.exists():
        cfg.update(json.loads(CONFIG_PATH.read_text(encoding="utf-8")))
    if not cfg["azure_client_id"]:
        sys.exit("config.json 에 azure_client_id 를 입력하세요. (README.md 의 'Azure 앱 등록' 참고)")
    return cfg


def get_access_token(cfg: dict) -> str:
    """디바이스 코드 방식 로그인. 최초 1회만 브라우저 로그인, 이후 토큰 캐시로 자동 갱신."""
    cache = msal.SerializableTokenCache()
    if TOKEN_CACHE_PATH.exists():
        cache.deserialize(TOKEN_CACHE_PATH.read_text(encoding="utf-8"))
    app = msal.PublicClientApplication(
        cfg["azure_client_id"], authority=cfg["azure_authority"], token_cache=cache)

    result = None
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
    if not result:
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            sys.exit(f"로그인 시작 실패: {flow.get('error_description')}")
        print("\n[Outlook 로그인] 아래 주소에서 코드를 입력하고, 모바일 Outlook 앱과 같은 계정으로 로그인하세요.")
        print(f"  주소: {flow['verification_uri']}\n  코드: {flow['user_code']}\n")
        try:
            webbrowser.open(flow["verification_uri"])
        except Exception:
            pass
        result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        sys.exit(f"로그인 실패: {result.get('error_description')}")
    if cache.has_state_changed:
        TOKEN_CACHE_PATH.write_text(cache.serialize(), encoding="utf-8")
    return result["access_token"]


class Graph:
    """Microsoft Graph 호출 래퍼 (429/503 재시도 포함)."""

    def __init__(self, token: str):
        self.s = requests.Session()
        self.s.headers["Authorization"] = f"Bearer {token}"

    def request(self, method: str, url: str, **kw) -> requests.Response:
        if not url.startswith("http"):
            url = GRAPH + url
        for attempt in range(5):
            r = self.s.request(method, url, timeout=120, **kw)
            if r.status_code in (429, 503, 504):
                time.sleep(int(r.headers.get("Retry-After", 2 ** attempt)))
                continue
            if r.status_code >= 400:
                raise RuntimeError(f"Graph {method} {url} -> {r.status_code}: {r.text[:500]}")
            return r
        raise RuntimeError(f"Graph {method} {url}: 재시도 한도 초과")

    def get(self, url, **kw):
        return self.request("GET", url, **kw).json()

    def post(self, url, body, **kw):
        return self.request("POST", url, json=body, **kw).json()

    def paged(self, url, params=None):
        data = self.get(url, params=params)
        yield from data.get("value", [])
        while "@odata.nextLink" in data:
            data = self.get(data["@odata.nextLink"])
            yield from data.get("value", [])


# ---------------------------------------------------------------------------
# 1) 지멘스 공문 메일 조회
# ---------------------------------------------------------------------------

def is_siemens_letter(msg: dict, cfg: dict) -> bool:
    sender = (msg.get("from") or {}).get("emailAddress") or {}
    address = (sender.get("address") or "").lower()
    domain = address.split("@")[-1]
    name = (sender.get("name") or "").lower()
    subject = (msg.get("subject") or "").lower()
    preview = (msg.get("bodyPreview") or "").lower()

    from_siemens = any(k in domain for k in cfg["sender_domain_keywords"])
    mentions_siemens = any(k.lower() in f"{name} {subject}" for k in cfg["text_keywords"])
    if not (from_siemens or mentions_siemens):
        return False
    if cfg["only_with_attachments"] and not msg.get("hasAttachments"):
        return False
    # 지멘스 도메인 발신은 모두 공문 후보, 그 외(전달 메일 등)는 공문 키워드가 있어야 함
    return from_siemens or any(k.lower() in f"{subject} {preview}" for k in cfg["letter_keywords"])


def find_siemens_letters(graph: Graph, cfg: dict, days: int) -> list[dict]:
    me = graph.get("/me", params={"$select": "mail,userPrincipalName"})
    my_addresses = {(me.get("mail") or "").lower(), (me.get("userPrincipalName") or "").lower()}
    since = (dt.datetime.now(dt.timezone.utc) - dt.timedelta(days=days)).strftime("%Y-%m-%dT%H:%M:%SZ")
    params = {
        "$filter": f"receivedDateTime ge {since}",
        "$orderby": "receivedDateTime desc",
        "$select": "id,subject,from,receivedDateTime,hasAttachments,bodyPreview,body,"
                   "webLink,internetMessageId,isDraft,importance",
        "$top": "50",
    }
    letters = []
    for msg in graph.paged("/me/messages", params):
        sender = (((msg.get("from") or {}).get("emailAddress") or {}).get("address") or "").lower()
        if msg.get("isDraft") or sender in my_addresses:
            continue
        if is_siemens_letter(msg, cfg):
            letters.append(msg)
    # 첨부가 있는 공문을 우선 처리
    letters.sort(key=lambda m: (not m.get("hasAttachments"), m["receivedDateTime"]))
    return letters


# ---------------------------------------------------------------------------
# 2) 첨부파일 추출
# ---------------------------------------------------------------------------

def safe_name(text: str, limit: int = 60) -> str:
    text = re.sub(r'[\\/:*?"<>|\r\n\t]+', "_", text or "").strip(" ._")
    return text[:limit] or "untitled"


def download_attachments(graph: Graph, msg: dict, folder: Path) -> list[dict]:
    """메일의 파일 첨부를 모두 저장하고 메타데이터 목록을 반환."""
    saved = []
    if not msg.get("hasAttachments"):
        return saved
    folder.mkdir(parents=True, exist_ok=True)
    for att in graph.paged(f"/me/messages/{msg['id']}/attachments"):
        kind = att.get("@odata.type", "")
        if kind.endswith("referenceAttachment"):
            # 클라우드 링크 첨부 (OneDrive/SharePoint 공유 링크)
            saved.append({"name": att.get("name"), "link": att.get("sourceUrl"), "path": None})
            continue
        if not kind.endswith("fileAttachment") or att.get("isInline"):
            continue  # 서명 이미지 등 본문 삽입 이미지, 전달된 메일 항목은 제외
        if att.get("contentBytes"):
            content = base64.b64decode(att["contentBytes"])
        else:
            content = graph.request(
                "GET", f"/me/messages/{msg['id']}/attachments/{att['id']}/$value").content
        path = folder / safe_name(att["name"], 120)
        path.write_bytes(content)
        saved.append({"name": att["name"], "path": path, "size": len(content),
                      "content_type": att.get("contentType") or "application/octet-stream"})
    return saved


# ---------------------------------------------------------------------------
# 첨부 → Claude 입력 블록 변환
# ---------------------------------------------------------------------------

def _xml_text(data: bytes, para_tags: tuple[str, ...]) -> str:
    root = ElementTree.fromstring(data)
    lines = []
    for el in root.iter():
        tag = el.tag.rsplit("}", 1)[-1]
        if tag in para_tags:
            line = "".join(t.text or "" for t in el.iter() if t.tag.rsplit("}", 1)[-1] in ("t", "T"))
            if line.strip():
                lines.append(line)
    return "\n".join(lines)


def office_text(path: Path) -> str | None:
    """docx / pptx / xlsx / hwpx 에서 텍스트 추출 (외부 라이브러리 없이 zip+XML 파싱)."""
    ext = path.suffix.lower()
    try:
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
            if ext == ".docx":
                return _xml_text(z.read("word/document.xml"), ("p",))
            if ext == ".pptx":
                slides = sorted((n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n)),
                                key=lambda n: int(re.findall(r"\d+", n)[-1]))
                return "\n\n".join(f"[슬라이드 {i}]\n" + _xml_text(z.read(n), ("p",))
                                   for i, n in enumerate(slides, 1))
            if ext == ".xlsx":
                shared = []
                if "xl/sharedStrings.xml" in names:
                    root = ElementTree.fromstring(z.read("xl/sharedStrings.xml"))
                    shared = ["".join(t.text or "" for t in si.iter() if t.tag.endswith("}t"))
                              for si in root if si.tag.endswith("}si")]
                out = []
                for n in sorted(n for n in names if re.match(r"xl/worksheets/sheet\d+\.xml$", n)):
                    out.append(f"[{n.rsplit('/', 1)[-1]}]")
                    for row in ElementTree.fromstring(z.read(n)).iter():
                        if not row.tag.endswith("}row"):
                            continue
                        cells = []
                        for c in row:
                            v = next((x.text for x in c.iter() if x.tag.endswith("}v") or x.tag.endswith("}t")), "")
                            if c.get("t") == "s" and v and v.isdigit() and int(v) < len(shared):
                                v = shared[int(v)]
                            cells.append(v or "")
                        if any(cells):
                            out.append("\t".join(cells))
                return "\n".join(out)
            if ext == ".hwpx":
                sections = sorted(n for n in names if re.match(r"Contents/section\d+\.xml$", n))
                return "\n".join(_xml_text(z.read(n), ("p",)) for n in sections)
    except (zipfile.BadZipFile, KeyError, ElementTree.ParseError):
        return None
    return None


def attachment_blocks(attachments: list[dict]) -> tuple[list[dict], list[str]]:
    """첨부파일을 Claude content block 으로 변환. (블록 목록, 요약 불가 파일 메모)"""
    blocks, notes = [], []
    for a in attachments:
        path: Path | None = a.get("path")
        if path is None:
            notes.append(f"- {a['name']}: 링크 첨부 ({a.get('link')}) — 내용 미포함")
            continue
        if a["size"] > MAX_SUMMARY_ATTACHMENT_BYTES:
            notes.append(f"- {a['name']}: 파일이 커서({a['size'] // 1024 // 1024}MB) 요약에서 제외")
            continue
        ext = path.suffix.lower()
        data = path.read_bytes()
        if ext == ".pdf":
            blocks.append({"type": "text", "text": f"[첨부파일: {a['name']}]"})
            blocks.append({"type": "document", "title": a["name"], "source": {
                "type": "base64", "media_type": "application/pdf",
                "data": base64.standard_b64encode(data).decode()}})
        elif ext in IMAGE_TYPES:
            blocks.append({"type": "text", "text": f"[첨부 이미지: {a['name']}]"})
            blocks.append({"type": "image", "source": {
                "type": "base64", "media_type": IMAGE_TYPES[ext],
                "data": base64.standard_b64encode(data).decode()}})
        elif ext in (".docx", ".pptx", ".xlsx", ".hwpx"):
            text = office_text(path)
            if text:
                blocks.append({"type": "text", "text": f"[첨부파일: {a['name']}]\n{text}"})
            else:
                notes.append(f"- {a['name']}: 텍스트 추출 실패")
        elif ext in (".txt", ".csv", ".md", ".xml", ".json", ".html", ".htm"):
            blocks.append({"type": "text",
                           "text": f"[첨부파일: {a['name']}]\n{data.decode('utf-8', errors='replace')}"})
        else:
            notes.append(f"- {a['name']}: 지원하지 않는 형식({ext or '확장자 없음'}) — PDF로 변환 후 재실행 권장")
    return blocks, notes


def html_to_text(body: dict) -> str:
    content = (body or {}).get("content") or ""
    if (body or {}).get("contentType") == "html":
        content = re.sub(r"(?is)<(script|style).*?</\1>", "", content)
        content = re.sub(r"(?i)<br\s*/?>|</p>|</div>|</tr>", "\n", content)
        content = html.unescape(re.sub(r"<[^>]+>", "", content))
    return re.sub(r"\n\s*\n+", "\n\n", content).strip()


# ---------------------------------------------------------------------------
# 3) Claude 요약
# ---------------------------------------------------------------------------

SUMMARY_SCHEMA = {
    "type": "object",
    "properties": {
        "document_title": {"type": "string", "description": "공문 제목"},
        "document_number": {"type": "string", "description": "공문/레터 번호 (없으면 빈 문자열)"},
        "sender_org": {"type": "string", "description": "발신 조직/담당자"},
        "document_date": {"type": "string", "description": "공문 일자 YYYY-MM-DD (없으면 빈 문자열)"},
        "category": {"type": "string",
                     "enum": ["계약", "클레임", "기술", "납기·일정", "품질", "대금·정산", "회의·협의", "일반", "기타"]},
        "summary": {"type": "string", "description": "핵심 내용 3~5문장 요약"},
        "key_points": {"type": "array", "items": {"type": "string"}},
        "requested_actions": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "action": {"type": "string"},
                    "due_date": {"type": "string", "description": "YYYY-MM-DD, 없으면 빈 문자열"},
                },
                "required": ["action", "due_date"],
                "additionalProperties": False,
            },
        },
        "response_deadline": {"type": "string", "description": "회신/조치 기한 YYYY-MM-DD, 없으면 빈 문자열"},
        "priority": {"type": "string", "enum": ["높음", "보통", "낮음"]},
        "risks": {"type": "string", "description": "당사에 미치는 영향·리스크 (없으면 빈 문자열)"},
    },
    "required": ["document_title", "document_number", "sender_org", "document_date", "category",
                 "summary", "key_points", "requested_actions", "response_deadline", "priority", "risks"],
    "additionalProperties": False,
}

SYSTEM_PROMPT = (
    "당신은 지멘스(Siemens)와의 계약·프로젝트 대응을 담당하는 실무자를 돕는 문서 분석가입니다. "
    "메일에 첨부된 공문을 최우선 근거로 삼고, 메일 본문은 보조 정보로만 사용하세요. "
    "공문에 적힌 사실만 요약하고, 문서에 없는 날짜·번호·금액은 추측하지 말고 빈 문자열로 두세요. "
    "상대방이 요구하는 조치와 기한, 당사에 불리할 수 있는 조항(클레임, 지체상금, 계약 변경 등)은 빠짐없이 짚어 주세요. "
    "모든 출력은 한국어로 작성하되, 고유명사·문서번호는 원문 표기를 유지하세요."
)


def summarize(client: anthropic.Anthropic, cfg: dict, msg: dict,
              attachments: list[dict]) -> dict:
    blocks, notes = attachment_blocks(attachments)
    sender = ((msg.get("from") or {}).get("emailAddress") or {})
    header = (
        f"메일 제목: {msg.get('subject')}\n"
        f"발신자: {sender.get('name')} <{sender.get('address')}>\n"
        f"수신일시: {msg.get('receivedDateTime')}\n"
        f"첨부파일: {', '.join(a['name'] for a in attachments) or '없음'}\n"
        + (f"요약에 포함되지 않은 첨부:\n" + "\n".join(notes) + "\n" if notes else "")
        + f"\n[메일 본문]\n{html_to_text(msg.get('body'))}"
    )
    content = blocks + [{"type": "text", "text": header + "\n\n위 지멘스 공문을 분석해 요약하세요."}]

    response = client.beta.messages.create(
        model=cfg["claude_model"],
        max_tokens=16000,
        system=SYSTEM_PROMPT,
        thinking={"type": "adaptive"},
        output_config={"effort": "medium", "format": {"type": "json_schema", "schema": SUMMARY_SCHEMA}},
        # 안전 분류기가 거절할 경우 서버 측에서 권장 모델로 자동 재시도
        betas=["server-side-fallback-2026-07-01"],
        fallbacks="default",
        messages=[{"role": "user", "content": content}],
    )
    if response.stop_reason == "refusal":
        raise RuntimeError("Claude가 요약 요청을 거절했습니다.")
    if response.stop_reason == "max_tokens":
        raise RuntimeError("요약 출력이 max_tokens 에서 잘렸습니다.")
    text = next(b.text for b in response.content if b.type == "text")
    result = json.loads(text)
    result["unsummarized_attachments"] = notes
    return result


def summary_markdown(msg: dict, s: dict, attachments: list[dict], onedrive_url: str | None) -> str:
    sender = ((msg.get("from") or {}).get("emailAddress") or {})
    lines = [
        f"# {s['document_title'] or msg.get('subject')}",
        "",
        f"- 공문번호: {s['document_number'] or '-'}",
        f"- 발신: {s['sender_org'] or sender.get('name')} ({sender.get('address')})",
        f"- 공문일자: {s['document_date'] or '-'} / 메일 수신: {msg.get('receivedDateTime')}",
        f"- 분류: {s['category']} / 중요도: {s['priority']}",
        f"- 회신·조치 기한: {s['response_deadline'] or '-'}",
        "",
        "## 요약",
        s["summary"],
        "",
        "## 주요 내용",
        *[f"- {p}" for p in s["key_points"]],
        "",
        "## 요청 조치사항",
        *([f"- [ ] {a['action']}" + (f" (기한: {a['due_date']})" if a["due_date"] else "")
           for a in s["requested_actions"]] or ["- 없음"]),
    ]
    if s["risks"]:
        lines += ["", "## 영향·리스크", s["risks"]]
    lines += ["", "## 첨부파일", *([f"- {a['name']}" for a in attachments] or ["- 없음"])]
    if s.get("unsummarized_attachments"):
        lines += ["", "요약에서 제외된 첨부:", *s["unsummarized_attachments"]]
    lines += ["", f"- Outlook 원문: {msg.get('webLink')}"]
    if onedrive_url:
        lines.append(f"- OneDrive 폴더: {onedrive_url}")
    return "\n".join(lines) + "\n"


# ---------------------------------------------------------------------------
# 4) '지멘스 대응' 프로젝트 등록
# ---------------------------------------------------------------------------

def onedrive_upload(graph: Graph, remote_path: str, data: bytes) -> dict:
    quoted = urllib.parse.quote(remote_path)
    if len(data) <= ONEDRIVE_SIMPLE_UPLOAD_LIMIT:
        return graph.request("PUT", f"/me/drive/root:/{quoted}:/content", data=data).json()
    session = graph.post(f"/me/drive/root:/{quoted}:/createUploadSession",
                         {"item": {"@microsoft.graph.conflictBehavior": "replace"}})
    chunk = 5 * 1024 * 1024  # 320KiB 배수
    result = {}
    for start in range(0, len(data), chunk):
        part = data[start:start + chunk]
        r = requests.put(session["uploadUrl"], data=part, timeout=300, headers={
            "Content-Range": f"bytes {start}-{start + len(part) - 1}/{len(data)}"})
        r.raise_for_status()
        result = r.json()
    return result


def register_onedrive(graph: Graph, cfg: dict, folder_name: str,
                      attachments: list[dict], md: str) -> str:
    base = f"{cfg['project_name']}/{folder_name}"
    for a in attachments:
        if a.get("path"):
            onedrive_upload(graph, f"{base}/{a['path'].name}", a["path"].read_bytes())
    onedrive_upload(graph, f"{base}/요약.md", md.encode("utf-8"))
    return graph.get(f"/me/drive/root:/{urllib.parse.quote(base)}")["webUrl"]


def due_datetime(date_str: str) -> dict | None:
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", date_str or ""):
        return {"dateTime": f"{date_str}T09:00:00", "timeZone": "Asia/Seoul"}
    return None


def task_title(msg: dict, s: dict) -> str:
    number = f"[{s['document_number']}] " if s["document_number"] else ""
    return f"[지멘스 공문] {number}{s['document_title'] or msg.get('subject')}"[:255]


def register_todo(graph: Graph, cfg: dict, msg: dict, s: dict, md: str,
                  attachments: list[dict], onedrive_url: str | None) -> str:
    lists = [l for l in graph.paged("/me/todo/lists") if l["displayName"] == cfg["project_name"]]
    todo_list = lists[0] if lists else graph.post("/me/todo/lists", {"displayName": cfg["project_name"]})
    links = [{"webUrl": msg["webLink"], "applicationName": "Outlook", "displayName": "공문 메일 원문"}]
    if onedrive_url:
        links.append({"webUrl": onedrive_url, "applicationName": "OneDrive", "displayName": "첨부 폴더"})
    task = {
        "title": task_title(msg, s),
        "body": {"contentType": "text", "content": md},
        "importance": {"높음": "high", "보통": "normal", "낮음": "low"}[s["priority"]],
        "categories": [s["category"]],
        "linkedResources": links,
    }
    if due := due_datetime(s["response_deadline"]):
        task["dueDateTime"] = due
    created = graph.post(f"/me/todo/lists/{todo_list['id']}/tasks", task)

    # 요청 조치사항은 체크리스트 단계로 등록
    for a in s["requested_actions"]:
        step = a["action"] + (f" (기한 {a['due_date']})" if a["due_date"] else "")
        graph.post(f"/me/todo/lists/{todo_list['id']}/tasks/{created['id']}/checklistItems",
                   {"displayName": step[:255]})
    # 3MB 이하 첨부는 작업에 직접 첨부 (큰 파일은 OneDrive 링크로 확인)
    for a in attachments:
        if a.get("path") and a["size"] <= TODO_ATTACHMENT_LIMIT:
            graph.post(f"/me/todo/lists/{todo_list['id']}/tasks/{created['id']}/attachments", {
                "@odata.type": "#microsoft.graph.taskFileAttachment",
                "name": a["name"], "contentType": a["content_type"],
                "contentBytes": base64.b64encode(a["path"].read_bytes()).decode()})
    return f"To Do 목록 '{cfg['project_name']}'"


def register_planner(graph: Graph, cfg: dict, msg: dict, s: dict, md: str,
                     attachments: list[dict], onedrive_url: str | None) -> str:
    plan_id = cfg["planner_plan_id"]
    if not plan_id:
        sys.exit("project_backend 가 planner 이면 config.json 에 planner_plan_id 가 필요합니다.")
    buckets = [b for b in graph.paged(f"/planner/plans/{plan_id}/buckets")
               if b["name"] == cfg["planner_bucket_name"]]
    bucket = buckets[0] if buckets else graph.post(
        "/planner/buckets", {"name": cfg["planner_bucket_name"], "planId": plan_id, "orderHint": " !"})
    task = {"planId": plan_id, "bucketId": bucket["id"], "title": task_title(msg, s),
            "priority": {"높음": 1, "보통": 5, "낮음": 9}[s["priority"]]}
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", s["response_deadline"] or ""):
        task["dueDateTime"] = f"{s['response_deadline']}T00:00:00Z"
    created = graph.post("/planner/tasks", task)

    def ref_key(url: str) -> str:  # Planner 참조 키는 % . : @ # 를 인코딩해야 함
        return (url.replace("%", "%25").replace(".", "%2E").replace(":", "%3A")
                .replace("@", "%40").replace("#", "%23"))

    references = {ref_key(msg["webLink"]): {"@odata.type": "microsoft.graph.plannerExternalReference",
                                            "alias": "공문 메일 원문", "type": "Other"}}
    if onedrive_url:
        references[ref_key(onedrive_url)] = {"@odata.type": "microsoft.graph.plannerExternalReference",
                                             "alias": "첨부 폴더(OneDrive)", "type": "Other"}
    checklist = {str(uuid.uuid4()): {
        "@odata.type": "microsoft.graph.plannerChecklistItem", "title": a["action"][:100], "isChecked": False}
        for a in s["requested_actions"][:20]}
    time.sleep(1)  # 작업 생성 직후 details 가 준비될 때까지 대기
    details = graph.get(f"/planner/tasks/{created['id']}/details")
    graph.request("PATCH", f"/planner/tasks/{created['id']}/details",
                  headers={"If-Match": details["@odata.etag"]},
                  json={"description": md, "references": references, "checklist": checklist})
    return f"Planner 버킷 '{cfg['planner_bucket_name']}'"


# ---------------------------------------------------------------------------
# 실행
# ---------------------------------------------------------------------------

def main() -> None:
    ap = argparse.ArgumentParser(description="지멘스 공문 → '지멘스 대응' 프로젝트 등록 및 요약")
    ap.add_argument("--days", type=int, help="조회 기간(일), 기본 3")
    ap.add_argument("--dry-run", action="store_true", help="프로젝트 등록 없이 추출·요약만")
    ap.add_argument("--force", action="store_true", help="이미 등록한 공문도 다시 처리")
    args = ap.parse_args()

    cfg = load_config()
    days = args.days or cfg["days"]
    graph = Graph(get_access_token(cfg))
    client = anthropic.Anthropic(api_key=cfg["anthropic_api_key"] or None)  # 비어 있으면 ANTHROPIC_API_KEY 사용

    out_root = (BASE_DIR / cfg["output_dir"] / cfg["project_name"])
    out_root.mkdir(parents=True, exist_ok=True)
    registry_path = out_root / "registry.json"
    registry = json.loads(registry_path.read_text(encoding="utf-8")) if registry_path.exists() else {}

    letters = find_siemens_letters(graph, cfg, days)
    print(f"최근 {days}일 지멘스 공문 {len(letters)}건 발견")

    results = []
    for msg in letters:
        key = msg.get("internetMessageId") or msg["id"]
        if key in registry and not args.force:
            print(f"- 건너뜀(이미 등록): {msg['subject']}")
            continue
        received = msg["receivedDateTime"][:10]
        folder_name = f"{received}_{safe_name(msg.get('subject'))}"
        folder = out_root / folder_name
        print(f"\n▶ {msg['subject']} ({received})")

        # 1. 첨부 우선 추출
        attachments = download_attachments(graph, msg, folder)
        print(f"  첨부 {len(attachments)}개 저장 → {folder}")

        # 2. 요약
        try:
            s = summarize(client, cfg, msg, attachments)
        except (anthropic.APIError, RuntimeError) as e:
            print(f"  ! 요약 실패: {e}")
            continue

        # 3. 프로젝트 등록
        onedrive_url, target = None, "로컬 폴더"
        if not args.dry_run and cfg["upload_to_onedrive"]:
            onedrive_url = register_onedrive(graph, cfg, folder_name, attachments,
                                             summary_markdown(msg, s, attachments, None))
        md = summary_markdown(msg, s, attachments, onedrive_url)
        folder.mkdir(parents=True, exist_ok=True)
        (folder / "요약.md").write_text(md, encoding="utf-8")
        if not args.dry_run:
            if cfg["project_backend"] == "todo":
                target = register_todo(graph, cfg, msg, s, md, attachments, onedrive_url)
            elif cfg["project_backend"] == "planner":
                target = register_planner(graph, cfg, msg, s, md, attachments, onedrive_url)
            registry[key] = {"subject": msg["subject"], "received": msg["receivedDateTime"],
                             "registered_at": dt.datetime.now().isoformat(timespec="seconds"),
                             "folder": folder_name}
            registry_path.write_text(json.dumps(registry, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"  등록 완료: {target}")
        results.append((msg, s))

    # 전체 요약 보고
    if results:
        print("\n" + "=" * 60 + f"\n지멘스 공문 요약 ({len(results)}건)\n" + "=" * 60)
        digest = [f"# 지멘스 공문 요약 ({dt.date.today()} 기준 최근 {days}일)\n"]
        for msg, s in results:
            item = (f"## {s['document_title'] or msg['subject']}\n"
                    f"- 번호: {s['document_number'] or '-'} | 분류: {s['category']} | 중요도: {s['priority']} "
                    f"| 기한: {s['response_deadline'] or '-'}\n\n{s['summary']}\n")
            print("\n" + item)
            digest.append(item)
        (out_root / f"요약보고_{dt.date.today()}.md").write_text("\n".join(digest), encoding="utf-8")


if __name__ == "__main__":
    main()
