from __future__ import annotations

import csv
import datetime as dt
import email.utils
import html as html_lib
import json
import re
import socket
import ssl
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Iterable

from lxml import etree, html
from openpyxl import Workbook

import paths


sys.dont_write_bytecode = True
paths.ensure_app_dirs()

RUNS_DIR = paths.RUNS_DIR
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
)

SECTION_LABELS = {
    "CAMBODIA_ECONOMY": "캄보디아 금융/경제",
    "CAMBODIA_POLITICS": "캄보디아 정치/사회",
    "VIETNAM_ECONOMY": "베트남 금융/경제",
    "VIETNAM_POLITICS": "베트남 정치/사회",
}

SECTION_NO = {
    "CAMBODIA_ECONOMY": "1",
    "CAMBODIA_POLITICS": "2",
    "VIETNAM_ECONOMY": "3",
    "VIETNAM_POLITICS": "4",
}


@dataclass
class Source:
    name: str
    url: str
    country: str
    fixed_section: str | None = None
    max_items: int = 5
    provider: str = "html"
    list_params: dict = field(default_factory=dict)


# provider = 목록을 몇 페이지든 넘겨받는 방식. 사이트마다 페이징 수단이 다르다.
#   caminsight : 1페이지는 HTML, 2페이지부터 「더보기」가 쓰는 POST /bbs/ajax_<섹션>.php (JSON)
#   camnews    : index.html?section=..&category=..&page=N — category 없이는 page가 먹지 않는다
#   m2         : /news/ajaxArticlePaging.php?page=N (인사이드비나·시티타임즈·베트남코리아타임즈 공통 CMS)
#   wprss      : WordPress RSS ?paged=N — 목록·본문 HTML이 403이라 RSS로만 접근된다
SOURCES: list[Source] = [
    Source("캄보디아 인사이트 경제", "https://thecaminsight.com/economy", "cambodia", "CAMBODIA_ECONOMY", provider="caminsight"),
    Source("캄보디아 인사이트 금융", "https://thecaminsight.com/finance", "cambodia", "CAMBODIA_ECONOMY", provider="caminsight"),
    Source("캄보디아 인사이트 정치", "https://thecaminsight.com/politics", "cambodia", "CAMBODIA_POLITICS", provider="caminsight"),
    Source("캄보디아 인사이트 사회", "https://thecaminsight.com/society", "cambodia", "CAMBODIA_POLITICS", provider="caminsight"),
    Source("캄푸치아 신문 경제", "http://camnews.kr/news/index.html?section=1", "cambodia", "CAMBODIA_ECONOMY", provider="camnews", list_params={"category": "95"}),
    Source("캄푸치아 신문 정치", "http://camnews.kr/news/index.html?section=2", "cambodia", "CAMBODIA_POLITICS", provider="camnews", list_params={"category": "88"}),
    Source("캄푸치아 신문 사회", "http://camnews.kr/news/index.html?section=9", "cambodia", "CAMBODIA_POLITICS", provider="camnews", list_params={"category": "108"}),
    Source("Khmer Times National", "https://www.khmertimeskh.com/category/national/", "cambodia", provider="wprss", list_params={"feed": "https://www.khmertimeskh.com/category/national/feed/"}),
    Source("Phnom Penh Post National", "https://phnompenhpost.com/category/national/", "cambodia", provider="wprss", list_params={"feed": "https://phnompenhpost.com/category/national/feed/"}),
    Source("베트남 코리아 타임즈 세계공급망", "https://www.vietnamkoreatimes.com/news/articleList.html?sc_section_code=S1N2&view_type=sm", "vietnam", provider="m2"),
    Source("베트남 코리아 타임즈 베트남 한걸음 더", "https://www.vietnamkoreatimes.com/news/articleList.html?sc_section_code=S1N7&view_type=sm", "vietnam", provider="m2"),
    Source("베트남 코리아 타임즈 현지속살", "https://www.vietnamkoreatimes.com/news/articleList.html?sc_section_code=S1N3&view_type=sm", "vietnam", provider="m2"),
    Source("시티타임즈 베트남 최신", "https://www.citytimes.co.kr/news/articleList.html?sc_section_code=S1N8&view_type=sm", "vietnam", provider="m2"),
    Source("인사이드비나 경제", "https://www.insidevina.com/news/articleList.html?sc_section_code=S1N6&view_type=sm", "vietnam", "VIETNAM_ECONOMY", provider="m2"),
    Source("인사이드비나 금융·부동산", "https://www.insidevina.com/news/articleList.html?sc_section_code=S1N12&view_type=sm", "vietnam", "VIETNAM_ECONOMY", provider="m2"),
    Source("인사이드비나 정치", "https://www.insidevina.com/news/articleList.html?sc_section_code=S1N5&view_type=sm", "vietnam", "VIETNAM_POLITICS", provider="m2"),
    Source("인사이드비나 사회·문화", "https://www.insidevina.com/news/articleList.html?sc_section_code=S1N9&view_type=sm", "vietnam", "VIETNAM_POLITICS", provider="m2"),
]

# 목록 페이지를 무한정 넘기지 않기 위한 사이트별 안전 상한
LIST_PAGE_LIMIT = 40
# 목록 '행'에서 날짜를 인정할 최대 길이. 이보다 큰 덩어리는 행이 아니라 목록 전체(또는
# 좌측 메뉴)라서, 거기서 뽑은 날짜는 옆 기사 것이거나 본문 속 숫자다.
MAX_ROW_TEXT = 400


@dataclass
class Item:
    category: str
    source_name: str
    title: str
    url: str
    published_date: str = ""
    file_type: str = "article"
    local_path: str = ""
    notes: str = ""
    original_url: str = ""
    extra: dict = field(default_factory=dict)

    @property
    def published_mm_dd(self) -> str:
        parsed = parse_date(self.published_date)
        return f"{parsed.month}.{parsed.day:02d}." if parsed else ""

    def row(self) -> dict:
        return {
            "category": self.category,
            "source_name": self.source_name,
            "title": self.title,
            "url": self.url,
            "original_url": self.original_url,
            "published_date": normalize_date(self.published_date) or self.published_date,
            "published_mm_dd": self.published_mm_dd,
            "file_type": self.file_type,
            "local_path": self.local_path,
            "notes": self.notes,
            "extra_json": json.dumps(self.extra, ensure_ascii=False),
        }


class Http:
    def __init__(self) -> None:
        ctx = ssl.create_default_context()
        self.opener = urllib.request.build_opener(
            urllib.request.ProxyHandler({}),
            urllib.request.HTTPCookieProcessor(),
            urllib.request.HTTPSHandler(context=ctx),
        )

    def get(self, url: str, referer: str | None = None, timeout: int = 25, attempts: int = 2) -> bytes:
        headers = {"User-Agent": USER_AGENT}
        if referer:
            headers["Referer"] = referer
        req = urllib.request.Request(url, headers=headers)
        last_exc: Exception | None = None
        for attempt in range(attempts):
            try:
                with self.opener.open(req, timeout=timeout) as resp:
                    return resp.read()
            except (urllib.error.URLError, TimeoutError, socket.timeout, ConnectionResetError) as exc:
                last_exc = exc
                time.sleep(0.6 * (attempt + 1))
        assert last_exc is not None
        raise last_exc

    def post(self, url: str, body: str, referer: str | None = None, timeout: int = 25, attempts: int = 2) -> bytes:
        headers = {
            "User-Agent": USER_AGENT,
            "Content-Type": "application/x-www-form-urlencoded",
        }
        if referer:
            headers["Referer"] = referer
        req = urllib.request.Request(url, data=body.encode("utf-8"), headers=headers)
        last_exc: Exception | None = None
        for attempt in range(attempts):
            try:
                with self.opener.open(req, timeout=timeout) as resp:
                    return resp.read()
            except (urllib.error.URLError, TimeoutError, socket.timeout, ConnectionResetError) as exc:
                last_exc = exc
                time.sleep(0.6 * (attempt + 1))
        assert last_exc is not None
        raise last_exc


def decode_html(data: bytes) -> str:
    head = data[:3000].decode("ascii", "ignore")
    match = re.search(r"charset=['\"]?([-\w]+)", head, re.I)
    encodings = [match.group(1)] if match else []
    encodings += ["utf-8", "cp949", "euc-kr"]
    best = ""
    best_score = -10**9
    for enc in encodings:
        try:
            text = data.decode(enc, "replace")
        except LookupError:
            continue
        score = sum("\uac00" <= ch <= "\ud7a3" for ch in text) - text.count("\ufffd") * 20
        if score > best_score:
            best = text
            best_score = score
    return best or data.decode("utf-8", "replace")


def parse_doc(data: bytes):
    return html.fromstring(decode_html(data))


def clean_text(value: str) -> str:
    return re.sub(r"\s+", " ", html_lib.unescape(value or "")).strip()


def text_content(node) -> str:
    return clean_text(node.text_content() if node is not None else "")


def absolutize(base: str, href: str) -> str:
    return urllib.parse.urljoin(base, html_lib.unescape(href or ""))


def parse_date(value: str | None) -> dt.date | None:
    if not value:
        return None
    value = clean_text(value)
    short_year_match = re.search(r"\b(\d{2})-(\d{1,2})-(\d{1,2})(?:\s+\d{1,2}:\d{2})?\b", value)
    if short_year_match:
        year, month, day = (int(x) for x in short_year_match.groups())
        year += 2000 if year < 80 else 1900
        try:
            return dt.date(year, month, day)
        except ValueError:
            pass
    iso_match = re.search(r"(20\d{2})-(\d{1,2})-(\d{1,2})", value)
    if iso_match:
        try:
            return dt.date(*(int(x) for x in iso_match.groups()))
        except ValueError:
            pass
    match = re.search(r"(20\d{2})[.\-/년\s]+(\d{1,2})[.\-/월\s]+(\d{1,2})", value)
    if match:
        try:
            return dt.date(*(int(x) for x in match.groups()))
        except ValueError:
            pass
    month_names = {
        "jan": 1, "january": 1,
        "feb": 2, "february": 2,
        "mar": 3, "march": 3,
        "apr": 4, "april": 4,
        "may": 5,
        "jun": 6, "june": 6,
        "jul": 7, "july": 7,
        "aug": 8, "august": 8,
        "sep": 9, "sept": 9, "september": 9,
        "oct": 10, "october": 10,
        "nov": 11, "november": 11,
        "dec": 12, "december": 12,
    }
    english_patterns = [
        r"\b([A-Za-z]{3,9})\.?\s+(\d{1,2}),?\s+(20\d{2})\b",
        r"\b(\d{1,2})\s+([A-Za-z]{3,9})\.?,?\s+(20\d{2})\b",
    ]
    for pattern in english_patterns:
        match = re.search(pattern, value)
        if not match:
            continue
        groups = match.groups()
        if groups[0].isdigit():
            day = int(groups[0])
            month = month_names.get(groups[1].lower())
            year = int(groups[2])
        else:
            month = month_names.get(groups[0].lower())
            day = int(groups[1])
            year = int(groups[2])
        if month:
            try:
                return dt.date(year, month, day)
            except ValueError:
                pass
    try:
        parsed = email.utils.parsedate_to_datetime(value)
        return parsed.date()
    except Exception:
        pass
    # 연도가 없는 "9.02." / "08-31" 류. 오탐이 많은 형태라 맨 마지막에만 시도한다.
    match = re.search(r"(\d{1,2})[.\-/월\s]+(\d{1,2})[.\-/일]?", value)
    if match:
        month, day = (int(x) for x in match.groups())
        try:
            return dt.date(dt.date.today().year, month, day)
        except ValueError:
            return None
    return None


def normalize_date(value: str | None) -> str:
    parsed = parse_date(value)
    return parsed.isoformat() if parsed else ""


def within_range(value: str, start: dt.date, end: dt.date) -> bool:
    parsed = parse_date(value)
    if not parsed:
        return True
    return start <= parsed <= end


# 기사 주소에 딸려오는 '목록 상태' 파라미터. 기사 자체와는 무관한데 페이지마다 값이
# 달라서, 떼어내지 않으면 같은 기사가 페이지 수만큼 중복 수집된다(캄푸치아 신문).
LISTING_PARAMS = {
    "utm_source", "utm_medium", "utm_campaign", "utm_term", "utm_content",
    "page", "paged", "pageno", "page_no", "curpage", "cp", "total", "box_idxno",
}


def clean_article_url(url: str) -> str:
    parsed = urllib.parse.urlsplit(url)
    query = [
        (key, value)
        for key, value in urllib.parse.parse_qsl(parsed.query, keep_blank_values=True)
        if key.lower() not in LISTING_PARAMS
    ]
    return urllib.parse.urlunsplit((parsed.scheme, parsed.netloc, parsed.path, urllib.parse.urlencode(query), ""))


def article_key(url: str) -> str:
    parsed = urllib.parse.urlsplit(clean_article_url(url))
    return urllib.parse.urlunsplit((parsed.scheme, parsed.netloc, parsed.path.rstrip("/"), parsed.query, ""))


def dedupe(items: Iterable[Item]) -> list[Item]:
    seen: set[str] = set()
    out: list[Item] = []
    for item in items:
        key = article_key(item.url) if item.url else f"{item.source_name}:{item.title}"
        if key in seen:
            continue
        seen.add(key)
        out.append(item)
    return out


def extract_listing_candidates(doc, base_url: str) -> list[dict]:
    # 스크립트·스타일을 먼저 걷어낸다. 남겨두면 행 텍스트에 인라인 JS가 섞여
    # 거기 든 숫자가 날짜로 잡힌다(캄푸치아 신문의 좌측 메뉴 스크립트).
    for bad in doc.xpath("//script|//style|//noscript"):
        parent = bad.getparent()
        if parent is not None:
            parent.remove(bad)
    candidates: list[dict] = []
    for a in doc.xpath("//a[@href]"):
        href = a.get("href") or ""
        url = clean_article_url(absolutize(base_url, href))
        if not is_article_url(url):
            continue
        title = clean_text(a.get("title") or text_content(a))
        if len(title) < 8:
            continue
        if title.lower() in GENERIC_TITLES:
            continue
        row_text, date = listing_row(a)
        candidates.append({"title": title, "url": url, "date": date, "row_text": row_text, "content": ""})
    return compact_candidates(candidates)


def is_article_url(url: str) -> bool:
    parsed = urllib.parse.urlparse(url)
    host = parsed.netloc.lower()
    path = parsed.path.lower()
    query = parsed.query.lower()
    if any(ext in path for ext in [".jpg", ".png", ".gif", ".pdf", ".zip"]):
        return False
    if "thecaminsight.com" in host:
        return bool(re.search(r"/(economy|finance|politics|society)/\d+", path))
    if "camnews.kr" in host:
        return "view.html" in path and "no=" in query
    if "insidevina.com" in host or "citytimes.co.kr" in host or "vietnamkoreatimes.com" in host:
        return "articleview" in path.lower() or "articleView" in url
    if "khmertimeskh.com" in host:
        return bool(re.search(r"/\d{6,}/", path))
    if "phnompenhpost.com" in host:
        return "/category/" not in path and "/national/" in path and len(path.strip("/").split("/")) >= 2
    return False


# 행 안에서 '날짜만 담긴 칸'을 먼저 찾기 위한 경로. 목록 행은 대개 요약문을 함께
# 싣기 때문에, 요약문에서 날짜를 긁으면 사진 설명이나 본문 속 연도를 발행일로 오인한다.
DATE_NODE_XPATH = (
    ".//time | .//*[@datetime]"
    " | .//*[contains(@class,'date')] | .//*[contains(@class,'time')]"
    " | .//*[contains(@class,'byline')] | .//*[contains(@class,'info')]"
)


def row_date_node(element) -> str:
    for node in element.xpath(DATE_NODE_XPATH):
        text = text_content(node)
        if not text or len(text) > 120:
            continue
        found = find_date(text)
        if found:
            return found
    return ""


def listing_row(node) -> tuple[str, str]:
    """목록 링크에서 위로 올라가며 (행 텍스트, 날짜)를 찾는다.

    행이라기엔 너무 큰 조상(MAX_ROW_TEXT 초과)에서 멈춘다. 거기서 뽑은 날짜는
    옆 기사 것이거나 좌측 메뉴·본문 숫자라 신뢰할 수 없기 때문이다.

    날짜는 두 번에 나눠 찾는다. 먼저 행 전체를 훑어 '날짜 칸'(span.date, time 등)을
    찾고, 그런 칸이 하나도 없을 때만 행 텍스트에서 날짜 모양을 찾는다. 순서를 섞으면
    요약문 속 연도가 날짜 칸을 이긴다(캄푸치아 신문의 "2026년 2.7%" 사례).
    """
    chain = []
    current = node
    for _ in range(8):
        if current is None:
            break
        text = text_content(current)
        if len(text) > MAX_ROW_TEXT:
            break
        chain.append((current, text))
        current = current.getparent()
    if not chain:
        return text_content(node), ""
    row_text = chain[-1][1]
    for element, _ in chain:
        found = row_date_node(element)
        if found:
            return row_text, found
    for element, text in chain:
        found = find_date(text)
        if found:
            return row_text, found
    return row_text, ""


# 연도가 없는 "M.D" 형태는 본문 숫자("2.7%", "1.4억 달러")와 구별이 안 된다.
# 앞뒤에 숫자·소수점·퍼센트·단위가 붙어 있으면 날짜로 보지 않는다.
_UNIT_AFTER = r"(?!\s*[%조억만천]|\s*달러|\s*동\b|\s*포인트|\s*배\b|\d)"
DATE_PATTERNS = [
    r"20\d{2}-\d{1,2}-\d{1,2}",
    r"\b\d{2}-\d{1,2}-\d{1,2}(?:\s+\d{1,2}:\d{2})?\b",
    r"20\d{2}[.\-/년\s]+\d{1,2}[.\-/월\s]+\d{1,2}\s*일?",
    r"\b[A-Za-z]{3,9}\.?\s+\d{1,2},?\s+20\d{2}\b",
    r"\b\d{1,2}\s+[A-Za-z]{3,9}\.?,?\s+20\d{2}\b",
    r"\d{1,2}월\s*\d{1,2}일",
    r"(?<![\d.,])\d{1,2}\.\s?\d{1,2}\.(?!\d)",
    r"(?<![\d.,\-])\d{1,2}-\d{1,2}(?:\s+\d{1,2}:\d{2})?" + _UNIT_AFTER,
]


def find_date(text: str) -> str:
    text = text or ""
    for pattern in DATE_PATTERNS:
        for match in re.finditer(pattern, text):
            value = clean_text(match.group(0))
            if is_plausible_date(value):
                return value
    return ""


def is_plausible_date(value: str) -> bool:
    parsed = parse_date(value)
    if not parsed:
        return False
    # 발행일이 내일보다 미래이거나 2000년 이전이면 날짜가 아니라 본문 숫자다.
    today = dt.date.today()
    return dt.date(2000, 1, 1) <= parsed <= today + dt.timedelta(days=1)


def compact_candidates(candidates: list[dict]) -> list[dict]:
    seen: set[str] = set()
    out: list[dict] = []
    for candidate in candidates:
        key = article_key(candidate["url"])
        if key in seen:
            continue
        seen.add(key)
        out.append(candidate)
    return out


# ---------------------------------------------------------------- 목록 provider
# 사이트마다 "2페이지 이후"를 받는 방법이 다르다. 아래 함수들은 모두 페이지 번호를
# 받아 후보 목록(title/url/date/row_text/content)을 돌려준다. content 는 목록 응답이
# 본문 요약까지 같이 줄 때만 채워지고, 상세 페이지 접속이 막힌 사이트에서 본문 대신 쓴다.

def fetch_list_page(http: Http, source: Source, page: int) -> list[dict]:
    if source.provider == "m2":
        return list_page_m2(http, source, page)
    if source.provider == "caminsight":
        return list_page_caminsight(http, source, page)
    if source.provider == "camnews":
        return list_page_camnews(http, source, page)
    if source.provider == "wprss":
        return list_page_wprss(http, source, page)
    if page > 1:
        return []
    return list_page_html(http, source.url)


def list_page_html(http: Http, url: str) -> list[dict]:
    return extract_listing_candidates(parse_doc(http.get(url, timeout=30, attempts=2)), url)


def split_host(url: str) -> str:
    parsed = urllib.parse.urlsplit(url)
    return f"{parsed.scheme}://{parsed.netloc}"


def list_page_m2(http: Http, source: Source, page: int) -> list[dict]:
    """인사이드비나·시티타임즈·베트남코리아타임즈가 쓰는 CMS의 「더보기」 엔드포인트.

    JSON이라 발행일(pub_date)이 그대로 나온다. HTML 목록을 긁을 때처럼 행에서
    날짜를 추측할 필요가 없다.
    """
    host = split_host(source.url)
    query = dict(urllib.parse.parse_qsl(urllib.parse.urlsplit(source.url).query))
    params = {
        "total": "0",
        "list_per_page": "20",
        "page_per_page": "10",
        "page": str(page),
        "sc_section_code": query.get("sc_section_code", ""),
        "sc_sub_section_code": query.get("sc_sub_section_code", ""),
        "view_type": query.get("view_type", "sm"),
    }
    url = f"{host}/news/ajaxArticlePaging.php?{urllib.parse.urlencode(params)}"
    payload = json.loads(http.get(url, referer=source.url, timeout=25).decode("utf-8", "replace"))
    out: list[dict] = []
    for row in payload.get("data") or []:
        # 같은 CMS라도 link_url 을 비워 보내는 사이트가 있다(베트남코리아타임즈·시티타임즈).
        link = row.get("link_url") or ""
        if not link and row.get("idxno"):
            link = f"/news/articleView.html?idxno={row['idxno']}"
        if not link:
            continue
        summary = clean_text(row.get("summary") or "")
        out.append({
            "title": clean_text(row.get("title") or ""),
            "url": clean_article_url(absolutize(host, link)),
            "date": clean_text(row.get("pub_date") or row.get("reg_date") or ""),
            "row_text": summary,
            "content": summary,
        })
    return out


def list_page_caminsight(http: Http, source: Source, page: int) -> list[dict]:
    """캄보디아 인사이트. 1페이지는 HTML, 2페이지부터는 「더보기」 버튼이 쓰는 POST."""
    if page <= 1:
        return list_page_html(http, source.url)
    parsed = urllib.parse.urlsplit(source.url)
    section = (parsed.path.strip("/").split("/") or [""])[0]
    if not section:
        return []
    url = f"{split_host(source.url)}/bbs/ajax_{section}.php"
    payload = json.loads(http.post(url, f"page={page}", referer=source.url, timeout=25).decode("utf-8", "replace"))
    out: list[dict] = []
    for row in payload if isinstance(payload, list) else []:
        href = row.get("href") or ""
        if not href:
            continue
        content = clean_text(row.get("wr_content") or "")
        out.append({
            "title": clean_text(row.get("wr_subject") or ""),
            "url": clean_article_url(absolutize(source.url, href)),
            "date": clean_text(row.get("wr_datetime") or ""),
            "row_text": content,
            "content": content,
        })
    return out


def list_page_camnews(http: Http, source: Source, page: int) -> list[dict]:
    """캄푸치아 신문. 섹션 index 는 페이징이 없고, category 를 붙인 목록만 page 가 먹는다."""
    parsed = urllib.parse.urlsplit(source.url)
    query = dict(urllib.parse.parse_qsl(parsed.query))
    category = source.list_params.get("category", "")
    if category:
        query["category"] = category
    query["page"] = str(page)
    url = urllib.parse.urlunsplit((parsed.scheme, parsed.netloc, parsed.path, urllib.parse.urlencode(query), ""))
    return list_page_html(http, url)


# 목록 응답이 발행일을 필드로 주는 provider — 상세 페이지에서 날짜를 추측하지 않는다.
TRUSTED_DATE_PROVIDERS = {"m2", "caminsight", "wprss"}
# 상세 페이지 접속이 막혀(403) 목록 응답의 본문을 그대로 쓰는 provider.
SKIP_DETAIL_PROVIDERS = {"wprss"}

RSS_CONTENT_NS = {"content": "http://purl.org/rss/1.0/modules/content/"}


def list_page_wprss(http: Http, source: Source, page: int) -> list[dict]:
    """Khmer Times / Phnom Penh Post. 목록·본문 HTML이 403이라 RSS 로만 접근된다.

    RSS 는 제목·링크·발행일에 본문(content:encoded 또는 description)까지 실어주므로
    상세 페이지 접속 없이 기사 한 건이 완성된다.
    """
    feed = source.list_params.get("feed") or source.url
    url = feed if page <= 1 else f"{feed}{'&' if '?' in feed else '?'}paged={page}"
    root = etree.fromstring(http.get(url, timeout=25, attempts=2))
    out: list[dict] = []
    for node in root.xpath("//item"):
        link = clean_text(node.findtext("link") or "")
        if not link:
            continue
        encoded = node.xpath("./content:encoded/text()", namespaces=RSS_CONTENT_NS)
        body = ""
        for raw in (encoded[0] if encoded else "", node.findtext("description") or ""):
            if not raw:
                continue
            try:
                body = clean_text(html.fromstring(raw).text_content())
            except Exception:
                body = clean_text(raw)
            if body:
                break
        out.append({
            "title": clean_text(node.findtext("title") or ""),
            "url": clean_article_url(link),
            "date": clean_text(node.findtext("pubDate") or ""),
            "row_text": body,
            "content": body[:7000],
        })
    return out


def fetch_article(http: Http, url: str, referer: str = "") -> tuple[str, str, str]:
    try:
        data = http.get(url, referer=referer or None, timeout=20, attempts=2)
    except urllib.error.HTTPError as exc:
        if exc.code in {403, 429, 500, 502, 503, 508}:
            data = Http().get(url, referer=referer or None, timeout=20, attempts=1)
        else:
            raise
    doc = parse_doc(data)
    title = meta_content(doc, "og:title", "twitter:title") or first_text(doc, ["//h1", "//*[contains(@class,'title')]", "//title"])
    date = article_date(doc) or first_text(doc, [
        "//time",
        "//*[contains(@class,'date')]",
        "//*[contains(@class,'profile_info')]",
        "//*[contains(@class,'profile_info_ct')]",
        "//*[contains(@class,'me-2')]",
        "//*[contains(@class,'byline')]",
        "//*[contains(@class,'info')]",
    ])
    text = extract_article_text(doc)
    return clean_text(title), find_date(date + " " + text[:300]), text


def is_timeout_error(exc: Exception) -> bool:
    if isinstance(exc, (TimeoutError, socket.timeout)):
        return True
    if isinstance(exc, urllib.error.URLError):
        reason = getattr(exc, "reason", None)
        if isinstance(reason, (TimeoutError, socket.timeout)):
            return True
        return "timed out" in str(reason).lower()
    return "timed out" in str(exc).lower()


def meta_content(doc, *names: str) -> str:
    for name in names:
        nodes = doc.xpath(f"//meta[@property='{name}' or @name='{name}']/@content")
        for value in nodes:
            value = clean_text(value)
            if value:
                return value
    return ""


def article_date(doc) -> str:
    candidates: list[str] = []
    for name in [
        "article:published_time",
        "article:modified_time",
        "og:updated_time",
        "datePublished",
        "date",
        "pubdate",
        "publishdate",
        "sailthru.date",
    ]:
        value = meta_content(doc, name)
        if value:
            candidates.append(value)
    candidates.extend(doc.xpath("//time/@datetime"))
    candidates.extend(doc.xpath("//*[@datetime]/@datetime"))
    for value in candidates:
        if parse_date(value):
            return clean_text(value)
    return ""


def first_text(doc, xpaths: list[str]) -> str:
    for xpath in xpaths:
        nodes = doc.xpath(xpath)
        for node in nodes:
            value = text_content(node) if hasattr(node, "text_content") else clean_text(str(node))
            if value:
                return value
    return ""


def extract_article_text(doc) -> str:
    for bad in doc.xpath("//script|//style|//noscript|//iframe|//nav|//header|//footer|//aside"):
        parent = bad.getparent()
        if parent is not None:
            parent.remove(bad)
    candidates = []
    for xpath in [
        "//article",
        "//*[contains(@class,'article')]",
        "//*[contains(@class,'content')]",
        "//*[contains(@class,'view')]",
        "//*[contains(@id,'article')]",
        "//*[contains(@id,'content')]",
    ]:
        for node in doc.xpath(xpath):
            text = text_content(node)
            if len(text) > 120:
                candidates.append(text)
    if candidates:
        return max(candidates, key=len)[:7000]
    body = first_text(doc, ["//body"])
    return body[:7000]


GENERIC_TITLES = {
    "경제", "금융", "정치", "사회", "national", "what's hot", "browsing: politics", "ranking news", "랭킹 뉴스"
}


def is_better_detail_title(detail_title: str, listing_title: str) -> bool:
    detail = clean_text(detail_title)
    listing = clean_text(listing_title)
    if not detail:
        return False
    if detail.lower() in GENERIC_TITLES:
        return False
    if len(detail) < 8 and len(listing) >= 8:
        return False
    return True


def clean_article_title(title: str, source_name: str = "") -> str:
    text = clean_text(title)
    if not text:
        return ""
    suffix_patterns = [
        r"\s*>\s*캄보디아인사이트\s*$",
        r"\s*[-–—|]\s*캄보디아인사이트\s*$",
        r"\s*[-–—|]\s*캄보디아\s*인사이트\s*$",
        r"\s*[-–—|]\s*캄푸치아\s*신문\s*$",
        r"\s*[-–—|]\s*Khmer\s*Times\s*$",
        r"\s*[-–—|]\s*Phnom\s*Penh\s*Post\s*$",
        r"\s*[-–—|]\s*phnompenhpost\.com\s*$",
        r"\s*[-–—|]\s*Vietnam\s*Korea\s*Times\s*$",
        r"\s*[-–—|]\s*베트남\s*코리아\s*타임즈\s*$",
        r"\s*[-–—|]\s*시티타임즈\s*$",
        r"\s*[-–—|]\s*인사이드비나\s*$",
    ]
    source = display_source_name(source_name)
    if source:
        suffix_patterns.append(rf"\s*[-–—|]\s*{re.escape(source)}\s*$")
    changed = True
    while changed:
        changed = False
        for pattern in suffix_patterns:
            new_text = re.sub(pattern, "", text, flags=re.I).strip()
            if new_text != text:
                text = new_text
                changed = True
    return text.strip(" \t\r\n-–—|>")


def display_source_name(source_name: str) -> str:
    if source_name.startswith("캄보디아 인사이트"):
        return "캄보디아 인사이트"
    if source_name.startswith("캄푸치아 신문"):
        return "캄푸치아 신문"
    if source_name.startswith("Khmer Times"):
        return "Khmer Times"
    if source_name.startswith("Phnom Penh Post"):
        return "Phnom Penh Post"
    if source_name.startswith("베트남 코리아 타임즈"):
        return "베트남 코리아 타임즈"
    if source_name.startswith("시티타임즈"):
        return "시티타임즈"
    if source_name.startswith("인사이드비나"):
        return "인사이드비나"
    return source_name


ECONOMY_KEYWORDS = {
    "bank", "banks", "banking", "finance", "financial", "economy", "economic", "investment", "investor",
    "market", "trade", "export", "import", "loan", "credit", "debt", "inflation", "gdp", "tax", "tariff",
    "real estate", "property", "currency", "riel", "dong", "fdi", "fund", "stock", "securities",
    "은행", "금융", "경제", "투자", "시장", "무역", "수출", "수입", "대출", "신용", "부채", "물가",
    "부동산", "환율", "증권", "기업", "산업", "관세", "성장", "FDI",
}
POLITICS_KEYWORDS = {
    "government", "minister", "prime minister", "parliament", "senate", "election", "law", "police",
    "court", "crime", "border", "diplomatic", "society", "education", "health", "tourism", "culture",
    "정부", "총리", "장관", "국회", "선거", "법", "경찰", "법원", "범죄", "국경", "외교", "사회",
    "교육", "보건", "관광", "문화", "공휴일", "마약", "사기",
}


def classify_by_rules(source: Source, title: str, content: str) -> tuple[str, float, str]:
    if source.fixed_section:
        return source.fixed_section, 1.0, "fixed_source"
    text = f"{title} {content[:1500]}".lower()
    economy = sum(1 for keyword in ECONOMY_KEYWORDS if keyword.lower() in text)
    politics = sum(1 for keyword in POLITICS_KEYWORDS if keyword.lower() in text)
    if source.country == "cambodia":
        econ_section, pol_section = "CAMBODIA_ECONOMY", "CAMBODIA_POLITICS"
    else:
        econ_section, pol_section = "VIETNAM_ECONOMY", "VIETNAM_POLITICS"
    if economy >= politics + 2:
        return econ_section, 0.85, f"rule_economy:{economy}/{politics}"
    if politics >= economy + 2:
        return pol_section, 0.85, f"rule_politics:{economy}/{politics}"
    if economy > politics:
        return econ_section, 0.55, f"weak_rule_economy:{economy}/{politics}"
    if politics > economy:
        return pol_section, 0.55, f"weak_rule_politics:{economy}/{politics}"
    return econ_section if source.country == "vietnam" else pol_section, 0.3, f"ambiguous:{economy}/{politics}"


def collect_source(
    http: Http,
    source: Source,
    start_date: dt.date,
    end_date: dt.date,
    limit: int,
    classifier: Callable[[dict], str] | None = None,
    progress: Callable[[str, dict], None] | None = None,
) -> list[Item]:
    """limit=0 이면 기간 안의 기사를 모두 가져온다(목록을 끝까지 넘긴다)."""
    out: list[Item] = []
    seen: set[str] = set()
    for page in range(1, LIST_PAGE_LIMIT + 1):
        if progress and page > 1:
            progress("source_page", {"source": source.name, "page": page, "count": len(out)})
        if page == LIST_PAGE_LIMIT and progress:
            # 여기서 끊긴 사이트는 '모두'를 다 가져온 게 아니다. 조용히 자르지 않는다.
            progress("source_truncated", {"source": source.name, "page": page, "count": len(out)})
        try:
            candidates = fetch_list_page(http, source, page)
        except Exception as exc:
            if page == 1:
                return [Item("error", source.name, source.name, source.url, notes=list_failure_note(exc))]
            break
        fresh: list[dict] = []
        for candidate in candidates:
            key = article_key(candidate["url"])
            if key in seen:
                continue
            seen.add(key)
            fresh.append(candidate)
        if not fresh:
            break
        reached_limit = False
        for candidate in fresh:
            if limit and len(out) >= limit:
                reached_limit = True
                break
            if progress:
                progress("item_start", {"source": source.name, "count": len(out)})
            item = build_item(http, source, candidate, start_date, end_date, classifier)
            if item is not None:
                out.append(item)
        if reached_limit:
            break
        # 목록은 최신순이다. 이 페이지의 기사가 모두 시작일보다 과거면 뒤 페이지는 볼 필요가 없다.
        # 날짜를 못 읽은 행은 판단에서 뺀다. 한 행만 날짜가 없어도 멈추지 못하면
        # 건수 '모두'일 때 목록 끝(LIST_PAGE_LIMIT)까지 헛돌게 된다.
        page_dates = [parse_date(candidate["date"]) for candidate in fresh]
        page_dates = [value for value in page_dates if value]
        if page_dates and max(page_dates) < start_date:
            break
    return out


def build_item(
    http: Http,
    source: Source,
    candidate: dict,
    start_date: dt.date,
    end_date: dt.date,
    classifier: Callable[[dict], str] | None = None,
) -> Item | None:
    if candidate["date"] and not within_range(candidate["date"], start_date, end_date):
        return None
    title = clean_article_title(candidate["title"], source.name)
    date = candidate["date"]
    content = candidate.get("content") or ""
    notes = ""
    if source.provider not in SKIP_DETAIL_PROVIDERS:
        try:
            detail_title, detail_date, detail_content = fetch_article(http, candidate["url"], source.url)
            if is_better_detail_title(detail_title, title):
                title = clean_article_title(detail_title, source.name)
            # 목록이 발행일을 필드로 준 사이트는 그 값이 상세 페이지 추출보다 정확하다.
            if not (date and source.provider in TRUSTED_DATE_PROVIDERS):
                date = detail_date or date
            content = detail_content or content
        except Exception as exc:
            notes = detail_failure_note(exc)
            if content:
                notes = append_note(notes, "목록 요약으로 대체")
    if date and not within_range(date, start_date, end_date):
        return None
    section, confidence, reason = classify_by_rules(source, title, content or candidate["row_text"])
    if confidence < 0.7 and classifier:
        try:
            llm_section = classifier({
                "source": source.name,
                "country": source.country,
                "title": title,
                "url": candidate["url"],
                "content": content[:3500],
                "rule_section": section,
                "rule_reason": reason,
            })
            if llm_section in SECTION_LABELS:
                section = llm_section
                reason = f"{reason}; llm:{llm_section}"
                confidence = 0.9
        except Exception as exc:
            notes = append_note(notes, classification_failure_note(exc))
            reason = f"{reason}; llm_failed"
    return Item(
        category=SECTION_LABELS.get(section, section),
        source_name=display_source_name(source.name),
        title=title,
        url=candidate["url"],
        published_date=normalize_date(date) or date,
        notes=notes,
        extra={
            "section_key": section,
            "classification_confidence": confidence,
            "classification_reason": reason,
            "article_text": content,
        },
    )


def detail_failure_note(exc: Exception) -> str:
    if is_timeout_error(exc):
        return "상세 페이지 접속 실패: 응답 시간 초과"
    if isinstance(exc, urllib.error.HTTPError):
        return f"상세 페이지 접속 실패: HTTP {exc.code}"
    return f"상세 페이지 접속 실패: {exc}"


def list_failure_note(exc: Exception) -> str:
    if is_timeout_error(exc):
        return "목록 페이지 접속 실패: 응답 시간 초과"
    if isinstance(exc, urllib.error.HTTPError):
        return f"목록 페이지 접속 실패: HTTP {exc.code}"
    return f"목록 페이지 접속 실패: {exc}"


def classification_failure_note(exc: Exception) -> str:
    if is_timeout_error(exc):
        return "LLM 분류 보류: 응답 시간 초과"
    return f"LLM 분류 보류: {exc}"


def append_note(current: str, note: str) -> str:
    current = clean_text(current)
    note = clean_text(note)
    if not current:
        return note
    if not note:
        return current
    return f"{current}; {note}"


def collect_all(
    start_date: dt.date,
    end_date: dt.date,
    max_per_source: int = 5,
    include_cambodia: bool = True,
    include_vietnam: bool = True,
    classifier: Callable[[dict], str] | None = None,
    progress: Callable[[str, dict], None] | None = None,
) -> list[Item]:
    http = Http()
    items: list[Item] = []
    for source in SOURCES:
        if source.country == "cambodia" and not include_cambodia:
            continue
        if source.country == "vietnam" and not include_vietnam:
            continue
        if progress:
            progress("source_start", {"source": source.name})
        source_items = collect_source(
            http, source, start_date, end_date, max_per_source,
            classifier=classifier, progress=progress,
        )
        items.extend(source_items)
        if progress:
            progress("source_done", {"source": source.name, "count": len(source_items)})
    return dedupe(items)


def make_run_dir(when: dt.datetime | None = None) -> Path:
    when = when or dt.datetime.now()
    base = RUNS_DIR / when.strftime("%y%m%d_%H%M")
    path = base
    suffix = 1
    while path.exists():
        suffix += 1
        path = base.with_name(f"{base.name}_{suffix}")
    path.mkdir(parents=True, exist_ok=True)
    return path


def write_outputs(items: list[Item], output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    fieldnames = list(Item("", "", "", "").row().keys())
    with (output_dir / "metadata.csv").open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for item in items:
            writer.writerow(item.row())
    wb = Workbook()
    ws = wb.active
    ws.title = "metadata"
    ws.append(fieldnames)
    for item in items:
        ws.append([item.row()[key] for key in fieldnames])
    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col[:100])
        ws.column_dimensions[col[0].column_letter].width = min(max(max_len + 2, 12), 50)
    wb.save(output_dir / "metadata.xlsx")


def item_from_row(row: dict) -> Item:
    extra = {}
    try:
        extra = json.loads(row.get("extra_json") or "{}")
    except Exception:
        extra = {}
    return Item(
        category=row.get("category", ""),
        source_name=row.get("source_name", ""),
        title=row.get("title", ""),
        url=row.get("url", ""),
        original_url=row.get("original_url", ""),
        published_date=row.get("published_date", ""),
        file_type=row.get("file_type", "article"),
        local_path=row.get("local_path", ""),
        notes=row.get("notes", ""),
        extra=extra,
    )


def read_items(run_dir: Path) -> list[Item]:
    path = run_dir / "metadata.csv"
    if not path.exists():
        return []
    with path.open("r", newline="", encoding="utf-8-sig") as f:
        return [item_from_row(row) for row in csv.DictReader(f)]


if __name__ == "__main__":
    end = dt.date.today()
    start = end - dt.timedelta(days=14)
    run_dir = make_run_dir()
    result = collect_all(start, end)
    write_outputs(result, run_dir)
    print(run_dir)
