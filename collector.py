from __future__ import annotations

import csv
import datetime as dt
import email.utils
import html as html_lib
import json
import re
import ssl
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Iterable

from lxml import html
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


SOURCES: list[Source] = [
    Source("캄보디아 인사이트 경제", "https://thecaminsight.com/economy", "cambodia", "CAMBODIA_ECONOMY"),
    Source("캄보디아 인사이트 금융", "https://thecaminsight.com/finance", "cambodia", "CAMBODIA_ECONOMY"),
    Source("캄보디아 인사이트 정치", "https://thecaminsight.com/politics", "cambodia", "CAMBODIA_POLITICS"),
    Source("캄보디아 인사이트 사회", "https://thecaminsight.com/society", "cambodia", "CAMBODIA_POLITICS"),
    Source("캄푸치아 신문 경제", "http://camnews.kr/news/index.html?section=1", "cambodia", "CAMBODIA_ECONOMY"),
    Source("캄푸치아 신문 정치", "http://camnews.kr/news/index.html?section=2", "cambodia", "CAMBODIA_POLITICS"),
    Source("캄푸치아 신문 사회", "http://camnews.kr/news/index.html?section=9", "cambodia", "CAMBODIA_POLITICS"),
    Source("Khmer Times National", "https://www.khmertimeskh.com/category/national/", "cambodia"),
    Source("Phnom Penh Post National", "https://phnompenhpost.com/category/national/", "cambodia"),
    Source("베트남 코리아 타임즈 세계공급망", "https://www.vietnamkoreatimes.com/news/articleList.html?sc_section_code=S1N2&view_type=sm", "vietnam"),
    Source("베트남 코리아 타임즈 베트남 한걸음 더", "https://www.vietnamkoreatimes.com/news/articleList.html?sc_section_code=S1N7&view_type=sm", "vietnam"),
    Source("베트남 코리아 타임즈 현지속살", "https://www.vietnamkoreatimes.com/news/articleList.html?sc_section_code=S1N3&view_type=sm", "vietnam"),
    Source("시티타임즈 베트남 최신", "https://www.citytimes.co.kr/news/articleList.html?sc_section_code=S1N8&view_type=sm", "vietnam"),
    Source("인사이드비나 경제", "https://www.insidevina.com/news/articleList.html?sc_section_code=S1N6&view_type=sm", "vietnam", "VIETNAM_ECONOMY"),
    Source("인사이드비나 금융·부동산", "https://www.insidevina.com/news/articleList.html?sc_section_code=S1N12&view_type=sm", "vietnam", "VIETNAM_ECONOMY"),
    Source("인사이드비나 정치", "https://www.insidevina.com/news/articleList.html?sc_section_code=S1N5&view_type=sm", "vietnam", "VIETNAM_POLITICS"),
    Source("인사이드비나 사회·문화", "https://www.insidevina.com/news/articleList.html?sc_section_code=S1N9&view_type=sm", "vietnam", "VIETNAM_POLITICS"),
]


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
            except (urllib.error.URLError, TimeoutError, ConnectionResetError) as exc:
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
    patterns = [
        r"(20\d{2})[.\-/년\s]+(\d{1,2})[.\-/월\s]+(\d{1,2})",
        r"(\d{1,2})[.\-/월\s]+(\d{1,2})[.\-/일]?",
    ]
    for pattern in patterns:
        match = re.search(pattern, value)
        if not match:
            continue
        parts = [int(x) for x in match.groups()]
        if len(parts) == 2:
            year = dt.date.today().year
            month, day = parts
        else:
            year, month, day = parts
        try:
            return dt.date(year, month, day)
        except ValueError:
            return None
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
        return None


def normalize_date(value: str | None) -> str:
    parsed = parse_date(value)
    return parsed.isoformat() if parsed else ""


def within_range(value: str, start: dt.date, end: dt.date) -> bool:
    parsed = parse_date(value)
    if not parsed:
        return True
    return start <= parsed <= end


def article_key(url: str) -> str:
    parsed = urllib.parse.urlsplit(url)
    query = urllib.parse.parse_qsl(parsed.query, keep_blank_values=True)
    query = [(k, v) for k, v in query if k.lower() not in {"utm_source", "utm_medium", "utm_campaign"}]
    return urllib.parse.urlunsplit((parsed.scheme, parsed.netloc, parsed.path.rstrip("/"), urllib.parse.urlencode(query), ""))


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
    candidates: list[dict] = []
    for a in doc.xpath("//a[@href]"):
        href = a.get("href") or ""
        url = absolutize(base_url, href)
        if not is_article_url(url):
            continue
        title = clean_text(a.get("title") or text_content(a))
        if len(title) < 8:
            continue
        if title.lower() in GENERIC_TITLES:
            continue
        container = nearest_container(a)
        row_text = text_content(container)
        date = find_date(row_text)
        candidates.append({"title": title, "url": url, "date": date, "row_text": row_text})
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


def nearest_container(node):
    current = node
    while current is not None:
        tag = getattr(current, "tag", "")
        if tag in {"li", "tr", "article"}:
            return current
        cls = (current.get("class") or "").lower()
        if any(token in cls for token in ["list", "article", "news", "post", "item"]):
            return current
        current = current.getparent()
    return node


def find_date(text: str) -> str:
    text = text or ""
    patterns = [
        r"20\d{2}-\d{1,2}-\d{1,2}",
        r"\b\d{2}-\d{1,2}-\d{1,2}(?:\s+\d{1,2}:\d{2})?\b",
        r"20\d{2}[.\-/년\s]+\d{1,2}[.\-/월\s]+\d{1,2}",
        r"\d{1,2}[.월]\s*\d{1,2}[.일]?",
        r"\b[A-Za-z]{3,9}\.?\s+\d{1,2},?\s+20\d{2}\b",
        r"\b\d{1,2}\s+[A-Za-z]{3,9}\.?,?\s+20\d{2}\b",
    ]
    for pattern in patterns:
        match = re.search(pattern, text)
        if match and parse_date(match.group(0)):
            return clean_text(match.group(0))
    return ""


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
) -> list[Item]:
    try:
        doc = parse_doc(http.get(source.url))
    except Exception as exc:
        return [Item("error", source.name, source.name, source.url, notes=f"list failed: {exc}")]
    out: list[Item] = []
    for candidate in extract_listing_candidates(doc, source.url):
        if len(out) >= limit:
            break
        if candidate["date"] and not within_range(candidate["date"], start_date, end_date):
            continue
        title = candidate["title"]
        date = candidate["date"]
        content = ""
        notes = ""
        try:
            detail_title, detail_date, content = fetch_article(http, candidate["url"], source.url)
            if is_better_detail_title(detail_title, title):
                title = detail_title
            date = detail_date or date
        except Exception as exc:
            notes = detail_failure_note(exc)
        if date and not within_range(date, start_date, end_date):
            continue
        section, confidence, reason = classify_by_rules(source, title, content or candidate["row_text"])
        if confidence < 0.7 and classifier:
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
        out.append(Item(
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
        ))
    return out


def detail_failure_note(exc: Exception) -> str:
    if isinstance(exc, urllib.error.HTTPError):
        return f"상세 페이지 접속 실패: HTTP {exc.code}"
    return f"상세 페이지 접속 실패: {exc}"


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
        source_items = collect_source(http, source, start_date, end_date, max_per_source, classifier=classifier)
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
