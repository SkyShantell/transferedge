import csv
import json
import shutil
import subprocess
import tempfile
import zipfile
import hashlib
import re
import time
import difflib
import os
from collections import deque
from datetime import datetime
from io import BytesIO
from dataclasses import dataclass
from pathlib import Path
from urllib.request import Request, urlopen
import xml.etree.ElementTree as ET
from PIL import Image, ImageDraw, ImageFont, ImageFilter
from dotenv import load_dotenv
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

_THIS_DIR = Path(__file__).resolve().parent
load_dotenv(dotenv_path=_THIS_DIR / ".env")

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}

BASE_DIR = _THIS_DIR
TEMPLATE_DIR = BASE_DIR / "PngTemplate"
OUT_DIR = Path(os.getenv("PROPSEDGE_PNG_OUT_DIR", str(BASE_DIR / "generated_png"))).expanduser()
PNG_PARENT_FOLDER_ID = os.getenv(
    "PROPS_EDGE_PNG_PARENT_FOLDER_ID",
    "1FbGWTi45fFsFrgjQvoalA2dBEwUnt_4G",
).strip()

MAX_ROWS = 10
PLAYER_CACHE_DIR = OUT_DIR / "player_cache"
BADGE_CACHE_DIR = OUT_DIR / "badge_cache"
PLAYER_LOOKUP_CSV = BASE_DIR / "nba_player_lookup.csv"
LOOKUP_SHEET_ID = os.getenv("LOOKUP_SHEET_ID", "1sM1aMiHdiTzj22PKvVhtFZzVlbxkCB2HNMYgQf16-wg")
LOOKUP_SHEET_TAB = os.getenv("LOOKUP_SHEET_TAB", "Sheet1")
GOOGLE_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive",
]
USE_PPTX_RENDER = os.getenv("USE_PPTX_RENDER", "").strip().lower() in {"1", "true", "yes", "on"}
RENDER_PNG_DEBUG = os.getenv("RENDER_PNG_DEBUG", "").strip().lower() in {"1", "true", "yes", "on"}
CUSTOM_FONT_FILES = [
    "Zing Rust Base.ttf",
    "ZingRustBase.ttf",
    "Zing Rust Base.otf",
    "ZingRustBase.otf",
]
BADGE_CROP_BOX = {
    "goblins": (33, 205, 91, 266),  # x1, y1, x2, y2 from sample png
    "demons": (33, 216, 86, 246),
}
BADGE_ICON_FILES = {
    "goblins": "goblins_icon.png",
    "demons": "demons_icon.png",
}
BADGE_ICON_LEFT_FILES = {
    "goblins": "goblins_icon_left.png",
    "demons": "demons_icon_left.png",
}
BADGE_ROTATE_LEFT = {
    "goblins": 60,
    "demons": 18,
}
BADGE_ROTATE_RIGHT = {
    "goblins": -20,
    "demons": -55,
}
PLAYER_Y_OFFSET = 6
LEFT_BADGE_SCALE = 1.55
LEFT_BADGE_X_SHIFT = 0.42
LEFT_BADGE_Y = {
    "goblins": -14,
    "demons": -18,
}
RIGHT_BADGE_Y = 22
PLAYER_IMAGE_MAP: dict[str, str] = {}
PLAYER_IMAGE_KEYS: list[str] = []
PLAYER_LOCAL_FACE_MAP: dict[str, Path] = {}
IMAGE_FAILURE_LOG = OUT_DIR / "image_failures.log"
MISSING_LOOKUP_LOG = OUT_DIR / "missing_in_lookup.log"


def _debug(msg: str) -> None:
    if RENDER_PNG_DEBUG:
        print(msg)


@dataclass
class ModeCfg:
    mode: str
    csv_path: Path
    template_pptx: Path
    layout_json: Path
    out_stem: str


def _tpl(mode: str, ext: str) -> Path:
    clear = TEMPLATE_DIR / f"{mode}_blank_clear.{ext}"
    if clear.exists():
        return clear
    return TEMPLATE_DIR / f"{mode}_blank.{ext}"


GOBLINS = ModeCfg(
    mode="goblins",
    csv_path=BASE_DIR / "projections_extract_all_props_goblins.csv",
    template_pptx=_tpl("goblins", "pptx"),
    layout_json=OUT_DIR / "template_layout_goblins.json",
    out_stem="best_goblins_local",
)

DEMONS = ModeCfg(
    mode="demons",
    csv_path=BASE_DIR / "projections_extract_all_props_demons.csv",
    template_pptx=_tpl("demons", "pptx"),
    layout_json=OUT_DIR / "template_layout_demons.json",
    out_stem="best_demons_local",
)

# (output_slug, selected_prop_in_csv)
CATEGORY_SPECS_BY_SPORT: dict[str, list[tuple[str, str]]] = {
    "NBA": [
        ("points", "Player Points"),
        ("assists", "Player Assists"),
        ("rebounds", "Player Rebounds"),
        ("blocks", "Player Blocks"),
        ("steals", "Player Steals"),
        ("threes", "Player Threes"),
        ("threes_attempted", "Player Threes Attempts"),
        ("points_rebounds", "Player Points + Rebounds"),
        ("points_assists", "Player Points + Assists"),
        ("rebounds_assists", "Player Rebounds + Assists"),
        ("points_rebounds_assists", "Player Points + Rebounds + Assists"),
    ],
    "MLB": [
        ("total_bases", "Batter Total Bases"),
        ("hits", "Batter Hits"),
        ("singles", "Batter Singles"),
        ("home_runs", "Batter Home Runs"),
        ("rbis", "Batter RBIs"),
        ("runs", "Batter Runs"),
        ("strikeouts", "Pitcher Strikeouts"),
        ("outs", "Pitcher Outs"),
        ("walks", "Pitcher Walks"),
        ("hits_allowed", "Pitcher Hits Allowed"),
        ("earned_runs", "Pitcher Earned Runs"),
    ],
}


def active_category_specs() -> list[tuple[str, str]]:
    sport = (os.getenv("PROPSEDGE_SPORT") or os.getenv("SPORT") or "MLB").strip().upper()
    specs = CATEGORY_SPECS_BY_SPORT.get(sport)
    if not specs:
        raise RuntimeError(f"Unsupported PROPSEDGE_SPORT={sport}. Use one of: {', '.join(CATEGORY_SPECS_BY_SPORT)}")
    override = (os.getenv("PROPSEDGE_TARGET_PROPS") or "").strip()
    if override:
        return [(_slug(item), item) for item in [p.strip() for p in override.split(",") if p.strip()]]
    return specs


def parse_pct(value: str) -> float | None:
    v = (value or "").strip()
    if not v:
        return None
    if v.endswith("%"):
        v = v[:-1]
    try:
        return float(v)
    except ValueError:
        return None


def pct_text(v: float | None) -> str:
    if v is None:
        return ""
    if abs(v - round(v)) < 1e-9:
        return f"{int(round(v))}%"
    return f"{v:.1f}%"


def tendency_label(row: dict[str, str]) -> str:
    vals = {
        "H2H": parse_pct(row.get("H2H", "")),
        "L10": parse_pct(row.get("L10", "")),
        "SZN": parse_pct(row.get("SZN", "")),
    }
    valid = {k: v for k, v in vals.items() if v is not None}
    if not valid:
        return "N/A"
    best = max(valid.values())
    labels = [k for k, v in valid.items() if abs(v - best) < 1e-9]
    return f"{pct_text(best)} - {'+'.join(labels)}"


def model_avg(row: dict[str, str]) -> str:
    vals = [parse_pct(row.get("SZN", "")), parse_pct(row.get("L10", "")), parse_pct(row.get("H2H", ""))]
    vals = [v for v in vals if v is not None]
    if not vals:
        return ""
    return pct_text(sum(vals) / len(vals))


def two_lines(name: str, max_chars: int = 15) -> tuple[str, str]:
    parts = (name or "").strip().split()
    if not parts:
        return "", ""
    if len(parts) == 1:
        return parts[0].upper(), ""
    line1 = []
    line2 = []
    cur = 0
    for p in parts:
        add = len(p) + (1 if cur else 0)
        if not line2 and cur + add <= max_chars:
            line1.append(p)
            cur += add
        else:
            line2.append(p)
    if not line2:
        line1 = parts[:-1]
        line2 = [parts[-1]]
    return " ".join(line1).upper(), " ".join(line2).upper()


def display_prop_text(prop: str) -> str:
    raw = (prop or "").strip()
    if not raw:
        return ""
    key = raw.lower()
    mapping = {
        "threes attempts": "THREES FGA",
        "threes attempted": "THREES FGA",
        "rebounds + assists": "REB+AST",
        "rebounds + assist": "REB+AST",
        "points + rebounds": "PTS+REB",
        "points + assists": "PTS+AST",
        "points + assist": "PTS+AST",
        "points + rebounds + assists": "PTS+REB+AST",
        "points + rebounds + assist": "PTS+REB+AST",
        "batter total bases": "TOTAL BASES",
        "batter hits": "HITS",
        "batter singles": "SINGLES",
        "batter home runs": "HOME RUNS",
        "batter rbis": "RBIS",
        "batter runs": "RUNS",
        "pitcher strikeouts": "STRIKEOUTS",
        "pitcher outs": "OUTS",
        "pitcher walks": "WALKS",
        "pitcher hits allowed": "HITS ALLOWED",
        "pitcher earned runs": "EARNED RUNS",
    }
    if key in mapping:
        return mapping[key]
    return raw.upper()


def read_rows(csv_path: Path) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    with csv_path.open("r", newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append(row)
    return rows


def _slug(text: str) -> str:
    s = re.sub(r"[^a-zA-Z0-9]+", "_", (text or "").strip().lower()).strip("_")
    return s or "unknown"


def clean_name(name: str) -> str:
    # Match PropsCash-style normalization for stable lookup.
    text = str(name or "")
    text = text.split("|", 1)[0]
    text = text.replace(".", "").replace("'", "").replace("-", "")
    parts: list[str] = []
    for token in text.split():
        if token.lower() in {"jr", "sr", "ii", "iii", "iv"}:
            continue
        parts.append(token)
    return " ".join(parts).strip().lower()


def _load_lookup_image_map(csv_path: Path) -> dict[str, str]:
    if not csv_path.exists():
        return {}
    out: dict[str, str] = {}
    try:
        with csv_path.open("r", newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            if not reader.fieldnames:
                return {}
            name_col = None
            img_col = None
            for h in reader.fieldnames:
                k = (h or "").strip().lower()
                if k in {"player name", "player_name", "name"}:
                    name_col = h
                if k in {"img", "image", "image url", "image_url", "player_image_url", "player img", "img url"}:
                    img_col = h
            if not name_col or not img_col:
                return {}
            for row in reader:
                nm = clean_name(row.get(name_col, ""))
                url = (row.get(img_col, "") or "").strip()
                if nm and url and nm not in out:
                    out[nm] = url
    except Exception:
        return {}
    return out


def _resolve_service_account_path() -> Path | None:
    env_path = (os.getenv("GOOGLE_SERVICE_ACCOUNT_PATH") or "").strip()
    if env_path:
        p = Path(env_path).expanduser()
        if p.exists():
            return p
    candidates = [
        BASE_DIR / "service_account.json",
        BASE_DIR / "booming-argon-480004-e3-fe3eae6966c1.json",
        BASE_DIR.parent / "PropsCash" / "booming-argon-480004-e3-fe3eae6966c1.json",
    ]
    for p in candidates:
        if p.exists():
            return p
    return None


def _load_lookup_image_map_from_sheet() -> dict[str, str]:
    sa_path = _resolve_service_account_path()
    if not sa_path:
        return {}
    try:
        creds = service_account.Credentials.from_service_account_file(
            str(sa_path), scopes=GOOGLE_SCOPES
        )
        sheets = build("sheets", "v4", credentials=creds)
        vals = (
            sheets.spreadsheets()
            .values()
            .get(spreadsheetId=LOOKUP_SHEET_ID, range=LOOKUP_SHEET_TAB)
            .execute()
            .get("values", [])
        )
    except Exception as e:
        print(f"[render_png] sheet lookup read failed: {e}")
        return {}
    if not vals:
        return {}
    header = vals[0]
    rows = vals[1:]
    name_idx = None
    img_idx = None
    for i, h in enumerate(header):
        k = str(h or "").strip().lower()
        if k in {"player name", "player_name", "name"}:
            name_idx = i
        if k in {"player img link", "player_image_url", "image_url", "img", "image", "player img"}:
            img_idx = i
    if name_idx is None or img_idx is None:
        return {}
    out: dict[str, str] = {}
    for r in rows:
        if len(r) <= max(name_idx, img_idx):
            continue
        nm = clean_name(r[name_idx])
        url = str(r[img_idx]).strip()
        if nm and url and nm not in out:
            out[nm] = url
    return out


def _get_google_creds():
    sa_path = _resolve_service_account_path()
    if not sa_path:
        return None
    return service_account.Credentials.from_service_account_file(
        str(sa_path), scopes=GOOGLE_SCOPES
    )


def _drive_find_folder(drive_service, name: str, parent_id: str) -> str | None:
    safe_name = name.replace("'", "\\'")
    q = (
        f"mimeType = 'application/vnd.google-apps.folder' and "
        f"name = '{safe_name}' and "
        f"'{parent_id}' in parents and trashed = false"
    )
    resp = drive_service.files().list(
        q=q,
        fields="files(id,name)",
        pageSize=10,
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()
    files = resp.get("files", [])
    return files[0]["id"] if files else None


def _drive_ensure_folder(drive_service, name: str, parent_id: str) -> str:
    existing = _drive_find_folder(drive_service, name, parent_id)
    if existing:
        return existing
    created = drive_service.files().create(
        body={
            "name": name,
            "mimeType": "application/vnd.google-apps.folder",
            "parents": [parent_id],
        },
        fields="id",
        supportsAllDrives=True,
    ).execute()
    return created["id"]


def _drive_upsert_png(drive_service, folder_id: str, local_png: Path) -> None:
    safe_name = local_png.name.replace("'", "\\'")
    q = (
        f"name = '{safe_name}' and "
        f"'{folder_id}' in parents and trashed = false"
    )
    resp = drive_service.files().list(
        q=q,
        fields="files(id,name)",
        pageSize=10,
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()
    files = resp.get("files", [])
    media = MediaFileUpload(str(local_png), mimetype="image/png", resumable=False)
    if files:
        drive_service.files().update(
            fileId=files[0]["id"],
            media_body=media,
            supportsAllDrives=True,
        ).execute()
    else:
        drive_service.files().create(
            body={"name": local_png.name, "parents": [folder_id]},
            media_body=media,
            fields="id",
            supportsAllDrives=True,
        ).execute()


def _upload_pngs_to_drive(goblins: list[Path], demons: list[Path]) -> None:
    if not PNG_PARENT_FOLDER_ID:
        print("[render_png] skip Drive upload: PROPS_EDGE_PNG_PARENT_FOLDER_ID is empty")
        return

    creds = _get_google_creds()
    if not creds:
        print("[render_png] skip Drive upload: service account JSON not found")
        return

    try:
        drive = build("drive", "v3", credentials=creds)
        date_folder_name = datetime.now().strftime("%Y-%m-%d")
        date_folder_id = _drive_ensure_folder(drive, date_folder_name, PNG_PARENT_FOLDER_ID)
        goblins_folder_id = _drive_ensure_folder(drive, "goblins", date_folder_id)
        demons_folder_id = _drive_ensure_folder(drive, "demons", date_folder_id)

        for p in goblins:
            _drive_upsert_png(drive, goblins_folder_id, p)
        for p in demons:
            _drive_upsert_png(drive, demons_folder_id, p)

        print(
            "[render_png] uploaded PNGs to Drive folder "
            f"{date_folder_name}/goblins and {date_folder_name}/demons"
        )
    except Exception as e:
        print(f"[render_png] Drive upload failed: {e}")


def _chunks(rows: list[dict[str, str]], size: int) -> list[list[dict[str, str]]]:
    return [rows[i : i + size] for i in range(0, len(rows), size)]


def _score_row(row: dict[str, str]) -> float:
    szn = parse_pct(row.get("SZN", ""))
    l10 = parse_pct(row.get("L10", ""))
    h2h = parse_pct(row.get("H2H", ""))
    vals = [v for v in (szn, l10, h2h) if v is not None]
    if vals:
        return sum(vals) / len(vals)
    return -1.0


def _dedupe_players(rows: list[dict[str, str]]) -> list[dict[str, str]]:
    out: list[dict[str, str]] = []
    seen: set[str] = set()
    for r in rows:
        key = (r.get("player_name", "") or "").strip().lower()
        if not key:
            continue
        if key in seen:
            continue
        seen.add(key)
        out.append(r)
    return out


def _player_image_url(row: dict[str, str]) -> str:
    # PropsCash style: lookup map first.
    name_lookup = clean_name(row.get("player_name", ""))
    if name_lookup and PLAYER_IMAGE_MAP:
        if name_lookup in PLAYER_IMAGE_MAP:
            return PLAYER_IMAGE_MAP[name_lookup]
        near = difflib.get_close_matches(name_lookup, PLAYER_IMAGE_KEYS, n=1, cutoff=0.88)
        if near:
            return PLAYER_IMAGE_MAP.get(near[0], "")
    return (row.get("player_img_url") or row.get("player_image_url") or "").strip()


def _normalize_name(name: str) -> str:
    return re.sub(r"[^a-z0-9 ]+", "", (name or "").strip().lower()).strip()


def _last_name(name: str) -> str:
    parts = _normalize_name(name).split()
    return parts[-1] if parts else ""


def _levenshtein(a: str, b: str) -> int:
    if not a:
        return len(b)
    if not b:
        return len(a)
    m, n = len(a), len(b)
    dp = [[0] * (n + 1) for _ in range(m + 1)]
    for i in range(m + 1):
        dp[i][0] = i
    for j in range(n + 1):
        dp[0][j] = j
    for i in range(1, m + 1):
        for j in range(1, n + 1):
            cost = 0 if a[i - 1] == b[j - 1] else 1
            dp[i][j] = min(
                dp[i - 1][j] + 1,
                dp[i][j - 1] + 1,
                dp[i - 1][j - 1] + cost,
            )
    return dp[m][n]


def _best_fuzzy_name(target: str, candidates: list[str], min_score: float = 0.84) -> str | None:
    t = _normalize_name(target)
    if not t:
        return None
    t_last = _last_name(t)
    best = None
    best_score = 0.0
    for cand in candidates:
        c = _normalize_name(cand)
        if not c:
            continue
        if t_last and _last_name(c) != t_last:
            continue
        dist = _levenshtein(t, c)
        score = 1 - dist / max(len(t), len(c))
        if score > best_score:
            best_score = score
            best = cand
    return best if best and best_score >= min_score else None


def _download_image_cached(url: str, player_name: str = "") -> Path | None:
    if not url:
        return None
    PLAYER_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    ext = ".png"
    key = hashlib.sha1(url.encode("utf-8")).hexdigest()
    out = PLAYER_CACHE_DIR / f"{key}.png"
    if out.exists() and out.stat().st_size > 0:
        try:
            Image.open(out).verify()
            return out
        except Exception:
            try:
                out.unlink(missing_ok=True)
            except Exception:
                pass

    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept": "image/avif,image/webp,image/apng,image/*,*/*;q=0.8",
        "Referer": "https://propsedge.io/",
        "Accept-Language": "en-US,en;q=0.9",
        "Connection": "keep-alive",
    }
    last_err: Exception | None = None
    for attempt in range(1, 5):
        try:
            req = Request(url, headers=headers)
            with urlopen(req, timeout=12) as resp:
                data = resp.read()
            if not data:
                raise RuntimeError("empty image response")
            img = Image.open(BytesIO(data)).convert("RGBA")
            img.save(out, format="PNG")
            return out
        except Exception as e:
            last_err = e
            time.sleep(0.30 * attempt)
    if last_err:
        msg = f"[render_png] image download failed for '{player_name}': {url} ({last_err})"
        print(msg)
        try:
            OUT_DIR.mkdir(parents=True, exist_ok=True)
            with IMAGE_FAILURE_LOG.open("a", encoding="utf-8") as f:
                f.write(msg + "\n")
        except Exception:
            pass
    return None


def _load_badge_icon(mode: str, side: str = "right") -> Image.Image | None:
    BADGE_CACHE_DIR.mkdir(parents=True, exist_ok=True)

    # Prefer explicit icon files provided by user.
    # left side can have a dedicated pre-rotated asset for exact look.
    candidates: list[Path] = []
    if side == "left":
        left_name = BADGE_ICON_LEFT_FILES.get(mode)
        if left_name:
            candidates.append(TEMPLATE_DIR / left_name)
    right_name = BADGE_ICON_FILES.get(mode)
    if right_name:
        candidates.append(TEMPLATE_DIR / right_name)

    for icon_path in candidates:
        if not icon_path.exists():
            continue
        try:
            ic = Image.open(icon_path).convert("RGBA")
            bb = ic.getchannel("A").getbbox()
            if bb:
                ic = ic.crop(bb)
            return ic
        except Exception:
            continue

    src = TEMPLATE_DIR / f"{mode}_sample.png"
    if not src.exists():
        return None

    box = BADGE_CROP_BOX.get(mode)
    if not box:
        return None

    try:
        im = Image.open(src).convert("RGBA")
        crop = im.crop(box)
        px = crop.load()
        w, h = crop.size
        # Remove dark row background around icon.
        for yy in range(h):
            for xx in range(w):
                r, g, b, a = px[xx, yy]
                if r < 95 and g < 105 and b < 105:
                    px[xx, yy] = (r, g, b, 0)
        bb = crop.getchannel("A").getbbox()
        if bb:
            crop = crop.crop(bb)
        return crop
    except Exception:
        return None


def _soft_remove_edge_background(image: Image.Image) -> Image.Image:
    """Softly remove only a neutral matte connected to the image edge."""
    im = image.convert("RGBA")
    width, height = im.size
    if width < 3 or height < 3:
        return im

    px = im.load()
    samples = [
        px[0, 0][:3], px[width - 1, 0][:3],
        px[0, height - 1][:3], px[width - 1, height - 1][:3],
    ]
    bg = tuple(sorted(c[i] for c in samples)[len(samples) // 2] for i in range(3))
    if max(bg) - min(bg) > 28:
        return im

    hard, soft = 16.0, 38.0
    queue: deque[tuple[int, int]] = deque()
    seen: set[tuple[int, int]] = set()
    for x in range(width):
        queue.extend(((x, 0), (x, height - 1)))
    for y in range(height):
        queue.extend(((0, y), (width - 1, y)))

    while queue:
        x, y = queue.popleft()
        if (x, y) in seen:
            continue
        seen.add((x, y))
        r, g, b, a = px[x, y]
        distance = ((r - bg[0]) ** 2 + (g - bg[1]) ** 2 + (b - bg[2]) ** 2) ** 0.5
        neutral = max(r, g, b) - min(r, g, b) <= 34
        if not neutral or distance > soft:
            continue
        new_alpha = 0 if distance <= hard else int(a * (distance - hard) / (soft - hard))
        px[x, y] = (r, g, b, new_alpha)
        if x:
            queue.append((x - 1, y))
        if x + 1 < width:
            queue.append((x + 1, y))
        if y:
            queue.append((x, y - 1))
        if y + 1 < height:
            queue.append((x, y + 1))
    return im


def _paste_player_avatar(base: Image.Image, row: dict[str, str], pbox: dict[str, float], mode: str) -> None:
    name_key = (row.get("player_name", "") or "").strip()
    _debug(f"[render_png][{mode}] player='{name_key}' start")
    url = _player_image_url(row)
    if not url:
        msg = f"[render_png] missing lookup image for '{name_key}'"
        _debug(msg)
        try:
            OUT_DIR.mkdir(parents=True, exist_ok=True)
            with MISSING_LOOKUP_LOG.open("a", encoding="utf-8") as f:
                f.write(msg + "\n")
        except Exception:
            pass
    else:
        _debug(f"[render_png][{mode}] player='{name_key}' url='{url[:90]}'")
    img_path = _download_image_cached(url, player_name=name_key)
    if img_path:
        _debug(f"[render_png][{mode}] player='{name_key}' cache='{img_path.name}'")
    if not img_path and name_key and PLAYER_LOCAL_FACE_MAP:
        match = _best_fuzzy_name(name_key, list(PLAYER_LOCAL_FACE_MAP.keys()))
        if match:
            img_path = PLAYER_LOCAL_FACE_MAP.get(match)
            if img_path:
                _debug(
                    f"[render_png][{mode}] player='{name_key}' reused_local_from='{match}' file='{img_path.name}'"
                )
    if not img_path:
        _debug(f"[render_png][{mode}] player='{name_key}' skip=no_image")
        return
    try:
        avatar = Image.open(img_path).convert("RGBA")
    except Exception:
        _debug(f"[render_png][{mode}] player='{name_key}' skip=bad_image_file")
        return

    # Keep original ESPN canvas framing for consistent row-to-row alignment.

    resample_lanczos = getattr(getattr(Image, "Resampling", Image), "LANCZOS", Image.LANCZOS)

    # Sized for the compact 10-row Sky Knows Bets card layout.
    target_h = int(max(78, min(92, pbox["h"] * 1.10)))
    scale = target_h / max(1, avatar.height)
    target_w = int(avatar.width * scale)
    avatar = avatar.resize((target_w, target_h), resample_lanczos)
    avatar = _soft_remove_edge_background(avatar)

    # Use a fixed anchor in the left player-image column.
    # Text box X from templates is not reliable for image placement.
    # Nudge face right/down so the left template badge is more visible
    # (appears further left and slightly higher relative to the player).
    center_x = 93
    row_top = int(pbox["y"])

    x = center_x - target_w // 2
    y = int(pbox["y"] + pbox["h"] - target_h + PLAYER_Y_OFFSET)
    # Keep avatar on-canvas even if template text anchors are too far left.
    x = max(6, min(140, x))
    y = max(0, min(base.height - target_h, y))

    # Add two badges in code:
    # - left badge: larger and behind player
    # - right badge: smaller and in front of player
    left_badge_src = _load_badge_icon(mode, side="left")
    right_badge_src = _load_badge_icon(mode, side="right")
    if left_badge_src is not None or right_badge_src is not None:
        resample_lanczos = getattr(getattr(Image, "Resampling", Image), "LANCZOS", Image.LANCZOS)
        right_h = int(max(36, min(62, target_h * 0.54)))
        if mode == "demons":
            right_h = int(right_h * 0.90)
        right_w = right_h
        right_badge = None
        if right_badge_src is not None:
            right_src = right_badge_src
            # Always apply right-side tilt from constants for consistent look.
            right_src = right_src.rotate(
                BADGE_ROTATE_RIGHT.get(mode, 0),
                expand=True,
                resample=Image.BICUBIC,
            )
            rb = right_src.getchannel("A").getbbox()
            if rb:
                right_src = right_src.crop(rb)
            right_w = int(right_src.width * (right_h / max(1, right_src.height)))
            right_badge = right_src.resize((right_w, right_h), resample_lanczos)

        left_h = int(right_h * LEFT_BADGE_SCALE)
        left_w = left_h
        left_badge = None
        if left_badge_src is not None:
            left_src = left_badge_src
            # Always apply runtime rotation so orientation stays consistent.
            left_src = left_src.rotate(
                BADGE_ROTATE_LEFT.get(mode, 0),
                expand=True,
                resample=Image.BICUBIC,
            )
            lb = left_src.getchannel("A").getbbox()
            if lb:
                left_src = left_src.crop(lb)
            left_w = int(left_src.width * (left_h / max(1, left_src.height)))
            left_badge = left_src.resize((left_w, left_h), resample_lanczos)

        # Left badge (behind player)
        if left_badge is not None:
            left_x = x - int(left_w * LEFT_BADGE_X_SHIFT)
            left_y = row_top + LEFT_BADGE_Y.get(mode, 0)
            left_x = max(0, min(base.width - left_w, left_x))
            left_y = max(0, min(base.height - left_h, left_y))
            base.alpha_composite(left_badge, dest=(left_x, left_y))

    # Player in front of left badge.
    base.alpha_composite(avatar, dest=(x, y))
    if name_key:
        PLAYER_LOCAL_FACE_MAP[name_key] = img_path
    _debug(f"[render_png][{mode}] player='{name_key}' pasted_face")

    # Right badge (in front of player)
    if right_badge is not None:
        bx = x + int(target_w * 0.62)
        by = row_top + RIGHT_BADGE_Y
        bx = max(0, min(base.width - right_w, bx))
        by = max(0, min(base.height - right_h, by))
        base.alpha_composite(right_badge, dest=(bx, by))
    _debug(f"[render_png][{mode}] player='{name_key}' pasted_badges")


def _shape_text(sp: ET.Element) -> str:
    vals = [t.text for t in sp.findall(".//a:t", NS) if t.text]
    return " ".join(vals).strip()


def _set_shape_lines(sp: ET.Element, lines: list[str]) -> None:
    tx = sp.find("./p:txBody", NS)
    if tx is None:
        return
    ps = tx.findall("./a:p", NS)
    if not ps:
        return
    while len(lines) > len(ps):
        tx.append(ET.Element(f"{{{NS['a']}}}p"))
        ps = tx.findall("./a:p", NS)
    for i, p in enumerate(ps):
        texts = p.findall(".//a:t", NS)
        if not texts:
            r = ET.SubElement(p, f"{{{NS['a']}}}r")
            ET.SubElement(r, f"{{{NS['a']}}}t")
            texts = p.findall(".//a:t", NS)
        value = lines[i] if i < len(lines) else ""
        texts[0].text = value
        for t in texts[1:]:
            t.text = ""


def _collect_placeholder_shapes(slide_xml: bytes) -> dict[str, list[ET.Element]]:
    root = ET.fromstring(slide_xml)
    shapes = root.findall(".//p:sp", NS)
    out: dict[str, list[ET.Element]] = {"player": [], "prop": [], "tendency": [], "model": []}
    for sp in shapes:
        txt = _shape_text(sp).lower()
        if txt == "vj edgecombe":
            out["player"].append(sp)
        elif txt == "7.5 points":
            out["prop"].append(sp)
        elif "100 %" in txt and "h2h" in txt:
            out["tendency"].append(sp)
        elif txt == "70 %":
            out["model"].append(sp)

    def by_y(sp: ET.Element) -> float:
        xfrm = sp.find("./p:spPr/a:xfrm", NS)
        off = xfrm.find("./a:off", NS) if xfrm is not None else None
        if off is None:
            return 0.0
        return float(off.attrib.get("y", 0))

    for k in out:
        out[k].sort(key=by_y)
        out[k] = out[k][:MAX_ROWS]
    return out


def _patch_slide(slide_xml: bytes, rows: list[dict[str, str]]) -> tuple[bytes, int]:
    root = ET.fromstring(slide_xml)
    shapes = _collect_placeholder_shapes(slide_xml)
    base = min(len(rows), len(shapes["player"]), len(shapes["prop"]))
    count = min(base, len(shapes["tendency"])) if shapes["tendency"] else base

    for i in range(count):
        row = rows[i]
        p1, p2 = two_lines(row.get("player_name", ""))
        prop_line = (row.get("prop_line", "") or "").strip()
        prop_name = (row.get("prop", "") or "").strip().upper()
        tend = tendency_label(row)
        model = model_avg(row)

        _set_shape_lines(shapes["player"][i], [p1, p2])
        _set_shape_lines(shapes["prop"][i], [prop_line, prop_name])
        if i < len(shapes["tendency"]):
            _set_shape_lines(shapes["tendency"][i], [tend])
        if i < len(shapes["model"]):
            _set_shape_lines(shapes["model"][i], [model])

    return ET.tostring(root, encoding="utf-8", xml_declaration=True), count


def _render_pptx_to_png(pptx: Path, out_png: Path) -> None:
    with tempfile.TemporaryDirectory(prefix="propsedge_png_") as td:
        tdir = Path(td)
        cmd = [
            "soffice",
            "--headless",
            "--convert-to",
            "png",
            "--outdir",
            str(tdir),
            str(pptx),
        ]
        proc = subprocess.run(cmd, check=False, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        generated = tdir / f"{pptx.stem}.png"
        if not generated.exists():
            raise RuntimeError(
                f"Expected PNG not generated for {pptx}\n"
                f"LibreOffice rc={proc.returncode}\nstdout:\n{proc.stdout}\nstderr:\n{proc.stderr}"
            )
        shutil.copy2(generated, out_png)


def _font(size: int, bold_italic: bool = False) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    # Prefer exact template font if present locally.
    for f in CUSTOM_FONT_FILES:
        p = TEMPLATE_DIR / f
        if p.exists():
            try:
                return ImageFont.truetype(str(p), size=size)
            except Exception:
                pass
    if bold_italic:
        cands = [
            "/usr/share/fonts/truetype/dejavu/DejaVuSansCondensed-BoldOblique.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-BoldOblique.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        ]
    else:
        cands = [
            "/usr/share/fonts/truetype/dejavu/DejaVuSansCondensed-Bold.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        ]
    for c in cands:
        if Path(c).exists():
            return ImageFont.truetype(c, size=size)
    return ImageFont.load_default()


def _load_layout_boxes(layout_json: Path) -> dict[str, list[dict[str, float]]]:
    payload = json.loads(layout_json.read_text(encoding="utf-8"))
    blank = payload.get("blank", {})
    d = blank.get("boxes", [])
    summary = payload.get("blank_summary", {})

    def pick(key: str) -> list[dict[str, float]]:
        arr = [b for b in d if key.lower() in (b.get("text", "").lower())]
        arr.sort(key=lambda b: (b["y"], b["x"]))
        return arr[:MAX_ROWS]

    def first_x(keys: list[str], fallback: float) -> float:
        for k in keys:
            arr = pick(k)
            if arr:
                return float(arr[0]["x"])
        return fallback

    def first_w(keys: list[str], fallback: float) -> float:
        for k in keys:
            arr = pick(k)
            if arr:
                return float(arr[0]["w"])
        return fallback

    player_boxes = pick("VJ EDGECOMBE")
    prop_boxes = pick("7.5 POINTS")
    tendency_boxes = pick("100 %")
    model_boxes = pick("70 %")

    row_ys = [float(y) for y in summary.get("row_y_starts_px", [])][:MAX_ROWS]
    if not row_ys and player_boxes:
        row_ys = [float(b["y"]) for b in player_boxes[:MAX_ROWS]]
    if not row_ys:
        row_ys = [225.0 + i * 110.0 for i in range(MAX_ROWS)]

    cols = summary.get("columns", {})
    player_x = float(cols.get("player_x") or first_x(["PLAYER", "VJ EDGECOMBE"], 205.0))
    prop_x = float(cols.get("prop_x") or first_x(["Prop", "7.5 POINTS"], 470.0))
    tendency_x = float(cols.get("tendency_x") or first_x(["TENDENCY", "100 %"], 684.0))
    model_x = float(cols.get("model_x") or first_x(["MODEL", "70 %"], 942.0))

    player_w = first_w(["VJ EDGECOMBE", "PLAYER"], 130.0)
    prop_w = first_w(["7.5 POINTS", "Prop"], 90.0)
    tendency_w = first_w(["100 %", "TENDENCY"], 226.0)
    model_w = first_w(["70 %", "MODEL"], 80.0)
    # Some extracted layout JSONs may contain corrupted/oversized text-box heights.
    # Prefer sane player placeholder heights; otherwise infer from row spacing.
    def _median(vals: list[float]) -> float:
        if not vals:
            return 0.0
        vals = sorted(vals)
        n = len(vals)
        if n % 2 == 1:
            return vals[n // 2]
        return (vals[n // 2 - 1] + vals[n // 2]) / 2.0

    sane_h = [
        float(b.get("h", 0.0))
        for b in player_boxes
        if 18.0 <= float(b.get("h", 0.0)) <= 120.0
    ]
    if sane_h:
        box_h = _median(sane_h)
    elif len(row_ys) >= 2:
        diffs = [row_ys[i + 1] - row_ys[i] for i in range(len(row_ys) - 1)]
        diffs = [d for d in diffs if d > 20]
        if diffs:
            # Empirical ratio from working templates (row spacing ~110 => text box ~63).
            box_h = max(52.0, min(78.0, _median(diffs) * 0.58))
        else:
            box_h = 63.0
    else:
        box_h = 63.0

    def mk(x: float, w: float) -> list[dict[str, float]]:
        return [{"x": x, "y": y, "w": w, "h": box_h} for y in row_ys[:MAX_ROWS]]

    return {
        "player": mk(player_x, player_w),
        "prop": mk(prop_x, prop_w),
        "tendency": mk(tendency_x, tendency_w),
        "model": mk(model_x, model_w),
    }


def _render_png_fallback(cfg: ModeCfg, rows: list[dict[str, str]], out_png: Path) -> None:
    """Render the Sky Knows Bets 10-row charcoal card design directly with Pillow."""
    width, height = 1080, 1350
    accent = (83, 244, 0, 255) if cfg.mode == "goblins" else (241, 31, 48, 255)
    accent_dark = (45, 153, 0, 255) if cfg.mode == "goblins" else (150, 9, 24, 255)
    graphite = (32, 34, 37, 255)
    charcoal = (28, 29, 31, 255)
    white = (255, 255, 255, 255)
    muted = (210, 213, 217, 255)

    img = Image.new("RGBA", (width, height), (247, 248, 249, 255))
    draw = ImageDraw.Draw(img)

    # Subtle neutral geometry keeps the sheet branded without competing with the mode color.
    draw.polygon([(0, 0), (470, 0), (0, 480)], fill=(239, 241, 243, 255))
    draw.polygon([(width, 0), (720, 0), (width, 420)], fill=(235, 238, 241, 255))
    draw.polygon([(0, height), (0, 1080), (350, height)], fill=(239, 241, 243, 255))
    draw.polygon([(width, height), (730, height), (width, 1010)], fill=(235, 238, 241, 255))

    def rounded_shadow(box: tuple[int, int, int, int], radius: int = 22, blur: int = 8) -> None:
        layer = Image.new("RGBA", img.size, (0, 0, 0, 0))
        ld = ImageDraw.Draw(layer)
        x1, y1, x2, y2 = box
        ld.rounded_rectangle((x1 + 2, y1 + 6, x2 + 2, y2 + 8), radius, fill=(0, 0, 0, 62))
        img.alpha_composite(layer.filter(ImageFilter.GaussianBlur(blur)))

    def contain(asset: Image.Image, box: tuple[int, int]) -> Image.Image:
        asset = asset.convert("RGBA").copy()
        asset.thumbnail(box, Image.Resampling.LANCZOS)
        return asset

    f_top = _font(23)
    f_date = _font(22)
    f_title = _font(66, bold_italic=True)
    f_subtitle = _font(22)
    f_col = _font(21)
    f_name = _font(25, bold_italic=True)
    f_meta = _font(15)
    f_line = _font(27, bold_italic=True)
    f_prop = _font(18)
    f_tendency = _font(25, bold_italic=True)
    f_tendency_sub = _font(17)
    f_model = _font(29, bold_italic=True)

    sport = (os.getenv("PROPSEDGE_SPORT") or os.getenv("SPORT") or "MLB").strip().upper()
    date_text = datetime.now().strftime("%m/%d/%Y").lstrip("0").replace("/0", "/")
    draw.rounded_rectangle((38, 35, 330, 92), 25, fill=charcoal)
    prizepicks_path = TEMPLATE_DIR / "PrizePicks_Logo.png"
    if prizepicks_path.exists():
        prizepicks_logo = contain(Image.open(prizepicks_path), (48, 48))
        img.alpha_composite(prizepicks_logo, (55, 39 + (48 - prizepicks_logo.height) // 2))
        draw.text((116, 49), sport, font=f_top, fill=white)
    else:
        draw.text((68, 49), sport, font=f_top, fill=white)
    draw.text((357, 53), date_text, font=f_date, fill=(63, 67, 72, 255))

    # Header and mode icon.
    rounded_shadow((30, 118, 805, 280), 28, 10)
    draw.rounded_rectangle((30, 118, 805, 280), 28, fill=charcoal)
    draw.rectangle((30, 252, 805, 280), fill=accent)
    draw.text((58, 142), "BEST", font=f_title, fill=white)
    best_w = draw.textlength("BEST ", font=f_title)
    mode_title = "GOBLINS" if cfg.mode == "goblins" else "DEMONS"
    draw.text((58 + best_w, 142), mode_title, font=f_title, fill=accent)
    draw.text((61, 220), f"TOP MODEL EDGES  •  {sport}", font=f_subtitle, fill=muted)

    icon_name = BADGE_ICON_FILES.get(cfg.mode, f"{cfg.mode}_icon.png")
    icon_path = TEMPLATE_DIR / icon_name
    mode_icon = None
    if icon_path.exists():
        icon_source = Image.open(icon_path).convert("RGBA")
        icon_bbox = icon_source.getchannel("A").getbbox()
        if icon_bbox:
            icon_source = icon_source.crop(icon_bbox)
        mode_icon = contain(icon_source, (145, 125))
        img.alpha_composite(mode_icon, (720 - mode_icon.width // 2, 128))

    # Sky Knows Bets logo and PrizePicks promo code.
    logo_path = TEMPLATE_DIR / "SKB_LOGO_Transparent.PNG"
    if logo_path.exists():
        logo = contain(Image.open(logo_path), (230, 215))
        img.alpha_composite(logo, (825 + (230 - logo.width) // 2, 34))
        promo_box = (816, 230, 1064, 282)
        draw.rounded_rectangle(promo_box, 16, fill=(119, 0, 255, 255))
        promo_main_font = _font(20)
        promo_sub_font = _font(12)
        draw.text((940, 238), "USE CODE: SCRAP", font=promo_main_font, fill=white, anchor="ma")
        draw.text((940, 264), "ON PRIZEPICKS", font=promo_sub_font, fill=white, anchor="ma")

    for label, x in (("PLAYER", 185), ("PROP", 470), ("TENDENCY", 710), ("MODEL", 932)):
        draw.text((x, 294), label, font=f_col, fill=(72, 76, 80, 255), anchor="mm")

    y0, row_h, gap = 320, 94, 7
    for i, row in enumerate(rows[:MAX_ROWS]):
        y = y0 + i * (row_h + gap)
        rounded_shadow((32, y, 1048, y + row_h), 20, 6)
        draw.rounded_rectangle((32, y, 1048, y + row_h), 20, fill=white)
        draw.rounded_rectangle((143, y, 1048, y + row_h), 20, fill=graphite)
        draw.rectangle((143, y, 1020, y + row_h), fill=graphite)
        draw.rounded_rectangle((914, y, 1048, y + row_h), 20, fill=accent)
        draw.rectangle((914, y, 1028, y + row_h), fill=accent)
        draw.rectangle((143, y, 149, y + row_h), fill=accent_dark)

        # Existing image lookup/cache behavior is retained; this also layers the mode badges.
        player_box = {"x": 170.0, "y": float(y + 5), "w": 140.0, "h": 82.0}
        _paste_player_avatar(img, row, player_box, cfg.mode)

        first, last = two_lines(row.get("player_name", ""), max_chars=18)
        name_y = y + 17 if last else y + 28
        draw.text((170, name_y), first, font=f_name, fill=white)
        if last:
            draw.text((170, name_y + 28), last, font=f_name, fill=white)

        prop_line = (row.get("prop_line", "") or "").strip()
        prop_name = display_prop_text(row.get("prop", ""))
        draw.text((436, y + 17), prop_line, font=f_line, fill=white)
        draw.text((436, y + 52), prop_name, font=f_prop, fill=accent)

        tendency = tendency_label(row)
        if " - " in tendency:
            tend_pct, tend_basis = tendency.split(" - ", 1)
        else:
            tend_pct, tend_basis = tendency, ""
        draw.text((660, y + 18), tend_pct, font=f_tendency, fill=accent)
        if tend_basis:
            draw.text((660, y + 54), tend_basis, font=f_tendency_sub, fill=muted)

        model = model_avg(row)
        if model:
            draw.text(
                (981, y + row_h / 2),
                model,
                font=f_model,
                fill=(18, 20, 22, 255),
                anchor="mm",
            )

    img.convert("RGB").save(out_png, format="PNG")


def render_mode(cfg: ModeCfg, rows_override: list[dict[str, str]] | None = None, out_name: str | None = None) -> Path:
    if not cfg.csv_path.exists():
        raise FileNotFoundError(f"Missing CSV: {cfg.csv_path}")
    if not cfg.template_pptx.exists():
        raise FileNotFoundError(f"Missing template PPTX: {cfg.template_pptx}")

    rows = rows_override if rows_override is not None else read_rows(cfg.csv_path)[:MAX_ROWS]
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    stem = out_name or cfg.out_stem
    out_png = OUT_DIR / f"{stem}.png"
    if USE_PPTX_RENDER:
        try:
            with tempfile.TemporaryDirectory(prefix=f"propsedge_{cfg.mode}_") as td:
                tmp_pptx = Path(td) / f"{cfg.out_stem}.pptx"
                with zipfile.ZipFile(cfg.template_pptx, "r") as zin:
                    slide_xml = zin.read("ppt/slides/slide1.xml")
                    patched, patched_count = _patch_slide(slide_xml, rows)
                    with zipfile.ZipFile(tmp_pptx, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                        for item in zin.infolist():
                            data = patched if item.filename == "ppt/slides/slide1.xml" else zin.read(item.filename)
                            zout.writestr(item, data)

                if patched_count > 0:
                    _render_pptx_to_png(tmp_pptx, out_png)
                else:
                    raise RuntimeError("No editable placeholders found; using direct PNG render")
        except Exception:
            _render_png_fallback(cfg, rows, out_png)
    else:
        _render_png_fallback(cfg, rows, out_png)
    return out_png


def render_all_categories(cfg: ModeCfg) -> list[Path]:
    all_rows = read_rows(cfg.csv_path)
    # Build image lookup map (PropsCash-style) from local lookup first.
    global PLAYER_IMAGE_MAP, PLAYER_IMAGE_KEYS, PLAYER_LOCAL_FACE_MAP
    PLAYER_LOCAL_FACE_MAP = {}
    sheet_map = _load_lookup_image_map_from_sheet()
    PLAYER_IMAGE_MAP = sheet_map
    if sheet_map:
        print(
            f"[render_png] using lookup sheet: {LOOKUP_SHEET_ID} / {LOOKUP_SHEET_TAB} "
            f"({len(sheet_map)} players)"
        )
    else:
        PLAYER_IMAGE_MAP = _load_lookup_image_map(PLAYER_LOOKUP_CSV)
    if PLAYER_IMAGE_MAP and not sheet_map:
        print(f"[render_png] using lookup file: {PLAYER_LOOKUP_CSV.name} ({len(PLAYER_IMAGE_MAP)} players)")
    if not PLAYER_IMAGE_MAP:
        print("[render_png] WARNING: no lookup map loaded (sheet/file). Player faces may be missing.")
    PLAYER_IMAGE_KEYS = list(PLAYER_IMAGE_MAP.keys())

    outputs: list[Path] = []
    for slug, selected_prop in active_category_specs():
        rows = [r for r in all_rows if (r.get("selected_prop", "") or "").strip() == selected_prop]
        rows.sort(
            key=lambda r: (
                _score_row(r),
                parse_pct(r.get("SZN", "")) or -1.0,
            ),
            reverse=True,
        )
        rows = _dedupe_players(rows)[:MAX_ROWS]
        if not rows:
            print(f"[render_png] skipping {cfg.mode}_{slug}: no rows for {selected_prop}")
            continue
        out_name = f"{cfg.mode}_{slug}"
        out_png = render_mode(cfg, rows_override=rows, out_name=out_name)
        outputs.append(out_png)
    return outputs


def main() -> None:
    goblins = render_all_categories(GOBLINS)
    demons = render_all_categories(DEMONS)
    print(f"Generated {len(goblins)} goblins PNG(s)")
    print(f"Generated {len(demons)} demons PNG(s)")
    for p in goblins + demons:
        print(f"Saved {p}")
    skip_upload = os.getenv("PROPSEDGE_SKIP_UPLOAD", "").strip().lower() in {"1", "true", "yes", "on"}
    if skip_upload:
        print("[render_png] Upload skipped by PROPSEDGE_SKIP_UPLOAD")
    else:
        _upload_pngs_to_drive(goblins, demons)


if __name__ == "__main__":
    main()
