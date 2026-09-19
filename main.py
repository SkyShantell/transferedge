import os
import re
import csv
import json
import mimetypes
import smtplib
import time
import uuid
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError, sync_playwright
from dotenv import load_dotenv
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
try:
    import psycopg2
except Exception:
    psycopg2 = None

from login import perform_login
import render_png

BASE_DIR = Path(__file__).resolve().parent
DRIVE_FOLDER_ID = "12zmrFvmtsAAixClGaY7PPeQ5mcAtDevW"
LOGS_FOLDER_ID = "1ffT9nyVUQNoY3u7lcEQO3coE159gbOUu"
SHEET_ID = "1cfVgt3Ooe8y2sgJQyfY7cX9dcFQnwYGtuYdLqAQlrCo"
GOOGLE_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
PNG_OUT_DIR = BASE_DIR / "generated_png"

SPORT_CONFIGS = {
    "NBA": {
        "label": "NBA",
        "key": "nba",
        "href_slug": "nba",
        "image_slug": "nba",
        "target_props": [
            "Player Points",
            "Player Assists",
            "Player Rebounds",
            "Player Blocks",
            "Player Steals",
            "Player Threes",
            "Player Threes Attempts",
            "Player Points + Rebounds",
            "Player Points + Assists",
            "Player Rebounds + Assists",
            "Player Points + Rebounds + Assists",
        ],
    },
    "MLB": {
        "label": "MLB",
        "key": "mlb",
        "href_slug": "mlb",
        "image_slug": "mlb",
        "target_props": [
            "Batter Total Bases",
            "Batter Hits",
            "Batter Singles",
            "Batter Home Runs",
            "Batter RBIs",
            "Batter Runs",
            "Pitcher Strikeouts",
            "Pitcher Outs",
            "Pitcher Walks",
            "Pitcher Hits Allowed",
            "Pitcher Earned Runs",
        ],
    },
}


def active_sport_config() -> dict:
    sport = (os.getenv("PROPSEDGE_SPORT") or os.getenv("SPORT") or "MLB").strip().upper()
    if sport not in SPORT_CONFIGS:
        raise RuntimeError(f"Unsupported PROPSEDGE_SPORT={sport}. Use one of: {', '.join(SPORT_CONFIGS)}")
    cfg = dict(SPORT_CONFIGS[sport])
    override = (os.getenv("PROPSEDGE_TARGET_PROPS") or "").strip()
    if override:
        cfg["target_props"] = [item.strip() for item in override.split(",") if item.strip()]
    return cfg


def _env_enabled(name: str, default: bool = False) -> bool:
    value = (os.getenv(name) or "").strip().lower()
    if not value:
        return default
    return value in {"1", "true", "yes", "on"}


def collect_generated_pngs() -> list[Path]:
    if not PNG_OUT_DIR.exists():
        return []
    files = []
    for p in sorted(PNG_OUT_DIR.glob("*.png")):
        n = p.name.lower()
        if n.startswith("_"):
            continue
        if n.startswith("goblins_") or n.startswith("demons_"):
            files.append(p)
    return files


def clear_generated_pngs(log) -> None:
    if not PNG_OUT_DIR.exists():
        return
    removed = 0
    for p in PNG_OUT_DIR.glob("*.png"):
        n = p.name.lower()
        if not (n.startswith("goblins_") or n.startswith("demons_")):
            continue
        try:
            p.unlink()
            removed += 1
        except OSError as exc:
            log(f"[Render] Failed to remove old PNG {p.name}: {exc}")
    if removed:
        log(f"[Render] Removed {removed} old generated PNG(s)")


def _encode_multipart(fields: dict[str, str], file_field: str, file_path: Path) -> tuple[bytes, str]:
    boundary = f"----PropsEdgeBoundary{uuid.uuid4().hex}"
    lines: list[bytes] = []
    for key, value in fields.items():
        lines.append(f"--{boundary}".encode("utf-8"))
        lines.append(f'Content-Disposition: form-data; name="{key}"'.encode("utf-8"))
        lines.append(b"")
        lines.append(str(value).encode("utf-8"))

    ctype = mimetypes.guess_type(str(file_path))[0] or "application/octet-stream"
    lines.append(f"--{boundary}".encode("utf-8"))
    lines.append(
        f'Content-Disposition: form-data; name="{file_field}"; filename="{file_path.name}"'.encode("utf-8")
    )
    lines.append(f"Content-Type: {ctype}".encode("utf-8"))
    lines.append(b"")
    lines.append(file_path.read_bytes())
    lines.append(f"--{boundary}--".encode("utf-8"))
    lines.append(b"")
    body = b"\r\n".join(lines)
    return body, boundary


def send_pngs_to_telegram(log, png_files: list[Path]) -> None:
    if not png_files:
        log("[Notify] No PNG files to send to Telegram")
        return

    if not _env_enabled("TELEGRAM_ENABLE", default=True):
        log("[Notify] Telegram disabled by TELEGRAM_ENABLE")
        return
    # PropsCash-compatible toggle
    if not _env_enabled("TELEGRAM_SEND_PNGS", default=True):
        log("[Notify] Telegram disabled by TELEGRAM_SEND_PNGS")
        return

    token = (os.getenv("TELEGRAM_BOT_TOKEN") or "").strip()
    chat_id = (os.getenv("TELEGRAM_CHAT_ID") or "").strip()
    if not token or not chat_id:
        log("[Notify] Telegram skipped (missing TELEGRAM_BOT_TOKEN or TELEGRAM_CHAT_ID)")
        return

    base_url = f"https://api.telegram.org/bot{token}"
    caption = f"PropsEdge PNGs - {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    sent = 0

    # Send header message first.
    try:
        qs = urlencode({"chat_id": chat_id, "text": caption})
        req = Request(f"{base_url}/sendMessage?{qs}", method="GET")
        with urlopen(req, timeout=20):
            pass
    except Exception as exc:
        log(f"[Notify] Telegram header message failed: {exc}")

    def send_one(png_path: Path) -> tuple[bool, float | None, str]:
        try:
            body, boundary = _encode_multipart(
                fields={"chat_id": chat_id},
                file_field="photo",
                file_path=png_path,
            )
            req = Request(
                f"{base_url}/sendPhoto",
                data=body,
                method="POST",
                headers={"Content-Type": f"multipart/form-data; boundary={boundary}"},
            )
            with urlopen(req, timeout=60) as resp:
                payload = resp.read().decode("utf-8", errors="replace")
                if getattr(resp, "status", 200) == 200:
                    return True, None, payload
                return False, None, payload
        except HTTPError as exc:
            body_txt = exc.read().decode("utf-8", errors="replace")
            retry_after = None
            try:
                data = json.loads(body_txt)
                retry_after = data.get("parameters", {}).get("retry_after")
            except Exception:
                pass
            return False, float(retry_after) if retry_after else None, body_txt
        except (URLError, TimeoutError, OSError) as exc:
            return False, None, str(exc)

    for png in png_files:
        delivered = False
        for attempt in range(3):
            ok, retry_after, detail = send_one(png)
            if ok:
                sent += 1
                delivered = True
                log(f"[Notify][Telegram] sent {png.name}")
                time.sleep(1.0)
                break
            if retry_after:
                delay = retry_after + 1.0
                log(f"[Notify][Telegram] rate limited for {png.name}, retry in {delay:.1f}s")
                time.sleep(delay)
                continue
            if attempt < 2:
                time.sleep(1.0)
        if not delivered:
            log(f"[Notify][Telegram] failed {png.name}: {detail}")

    log(f"[Notify] Telegram sent {sent}/{len(png_files)} PNG files")


def send_pngs_to_discord(log, png_files: list[Path]) -> None:
    if not png_files:
        log("[Notify] No PNG files to send to Discord")
        return
    if not _env_enabled("DISCORD_ENABLE", default=False):
        log("[Notify] Discord disabled by DISCORD_ENABLE")
        return

    goblins_webhook = (os.getenv("DISCORD_WEBHOOK_GOBLINS") or "").strip()
    demons_webhook = (os.getenv("DISCORD_WEBHOOK_DEMONS") or "").strip()
    if not goblins_webhook and not demons_webhook:
        log("[Notify] Discord skipped (missing DISCORD_WEBHOOK_GOBLINS / DISCORD_WEBHOOK_DEMONS)")
        return

    goblins_pngs = [p for p in png_files if p.name.lower().startswith("goblins_")]
    demons_pngs = [p for p in png_files if p.name.lower().startswith("demons_")]
    discord_headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept": "application/json",
    }

    def send_header(webhook_url: str, text: str) -> None:
        payload = json.dumps({"content": text}).encode("utf-8")
        req = Request(
            webhook_url + "?wait=true",
            data=payload,
            method="POST",
            headers={"Content-Type": "application/json", **discord_headers},
        )
        try:
            with urlopen(req, timeout=20):
                pass
        except Exception as exc:
            log(f"[Notify][Discord] header failed: {exc}")

    def send_one(webhook_url: str, png_path: Path) -> tuple[bool, float | None, str, int | None]:
        try:
            body, boundary = _encode_multipart(
                fields={"content": ""},
                file_field="file",
                file_path=png_path,
            )
            req = Request(
                webhook_url + "?wait=true",
                data=body,
                method="POST",
                headers={"Content-Type": f"multipart/form-data; boundary={boundary}", **discord_headers},
            )
            with urlopen(req, timeout=60) as resp:
                payload = resp.read().decode("utf-8", errors="replace")
                status = getattr(resp, "status", 200)
                if 200 <= status < 300:
                    return True, None, payload, status
                return False, None, payload, status
        except HTTPError as exc:
            status = getattr(exc, "code", None)
            body_txt = exc.read().decode("utf-8", errors="replace")
            retry_after = None
            if status == 429:
                try:
                    data = json.loads(body_txt)
                    retry_after = data.get("retry_after")
                except Exception:
                    retry_after = None
            return False, float(retry_after) if retry_after else None, body_txt, status
        except (URLError, TimeoutError, OSError) as exc:
            return False, None, str(exc), None

    def send_batch(label: str, webhook_url: str, files: list[Path]) -> None:
        if not webhook_url:
            log(f"[Notify][Discord] {label} skipped (missing webhook)")
            return
        if not files:
            log(f"[Notify][Discord] {label} has no PNG files")
            return

        header_text = f"Best {label} @everyone"
        send_header(webhook_url, header_text)
        sent = 0
        for png in files:
            delivered = False
            detail = ""
            status_code = None
            for attempt in range(5):
                ok, retry_after, detail, status_code = send_one(webhook_url, png)
                if ok:
                    sent += 1
                    delivered = True
                    log(f"[Notify][Discord][{label}] sent {png.name}")
                    time.sleep(0.8)
                    break

                if retry_after is not None:
                    delay = max(0.5, float(retry_after)) + 0.5
                    log(f"[Notify][Discord][{label}] 429 rate limit for {png.name}, retry in {delay:.1f}s")
                    time.sleep(delay)
                    continue

                transient = status_code in {500, 502, 503, 504} or status_code is None
                if transient and attempt < 4:
                    delay = min(20.0, 2.0 * (2 ** attempt))
                    log(
                        f"[Notify][Discord][{label}] transient failure for {png.name} "
                        f"(status={status_code}), retry in {delay:.1f}s"
                    )
                    time.sleep(delay)
                    continue
                break

            if not delivered:
                log(f"[Notify][Discord][{label}] failed {png.name} (status={status_code}): {detail}")

        log(f"[Notify] Discord {label} sent {sent}/{len(files)} PNG files")

    send_batch("Goblins", goblins_webhook, goblins_pngs)
    send_batch("Demons", demons_webhook, demons_pngs)


def maybe_send_email(log, subject: str, body: str) -> None:
    if not _env_enabled("EMAIL_ENABLE", default=True):
        log("[Notify] Email disabled by EMAIL_ENABLE")
        return

    # PropsCash-compatible env names first.
    sender = (os.getenv("GMAIL_USER") or "").strip()
    app_password = (os.getenv("GMAIL_APP_PASSWORD") or "").strip()
    notify_env = (os.getenv("NOTIFY_EMAILS") or "").strip()
    recipients = [addr.strip() for addr in notify_env.split(",") if addr.strip()] if notify_env else []

    # Fallback support (keeps current PropsEdge env working too).
    smtp_host = (os.getenv("SMTP_HOST") or "smtp.gmail.com").strip()
    smtp_port = int((os.getenv("SMTP_PORT") or "465").strip())
    smtp_user = (os.getenv("SMTP_USER") or "").strip()
    smtp_pass = (os.getenv("SMTP_PASS") or "").strip()
    if not sender:
        sender = smtp_user
    if not app_password:
        app_password = smtp_pass
    if not recipients:
        fallback_to = (os.getenv("NOTIFY_EMAIL_TO") or "").strip()
        if fallback_to:
            recipients = [addr.strip() for addr in fallback_to.split(",") if addr.strip()]

    if not sender or not app_password or not recipients:
        log("[Notify] Email skipped (missing sender/app-password/recipients)")
        return

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = sender
    msg["To"] = ", ".join(recipients)
    msg.set_content(body)

    try:
        # PropsCash behavior: SMTP SSL Gmail by default.
        with smtplib.SMTP_SSL(smtp_host, smtp_port, timeout=30) as smtp:
            smtp.login(sender, app_password)
            smtp.send_message(msg)
        log(f"[Notify] Email sent: {subject}")
    except Exception as exc:
        log(f"[Notify] Email failed ({subject}): {exc}")


def resolve_service_account_path() -> Path:
    """
    Deployment-safe credential discovery.
    Priority:
    1) GOOGLE_SERVICE_ACCOUNT_PATH env var
    2) PropsEdge/service_account.json
    3) PropsEdge/booming-argon-480004-e3-fe3eae6966c1.json
    4) sibling PropsCash JSON (legacy local layout)
    """
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

    raise RuntimeError(
        "Google service account JSON not found. "
        "Set GOOGLE_SERVICE_ACCOUNT_PATH or place service_account.json in PropsEdge."
    )


def dismiss_home_popups(page, log) -> None:
    # Ensure projections UI has settled before handling popups.
    try:
        page.wait_for_load_state("networkidle", timeout=10000)
    except PlaywrightTimeoutError:
        log("Network idle wait timed out; continuing popup handling")

    # 1) Wait for popup 1, then close it.
    welcome_modal = page.locator("div").filter(has=page.locator("#welcome-title")).first
    try:
        welcome_modal.wait_for(state="visible", timeout=7000)
        welcome_modal.get_by_label("Close", exact=True).first.click(timeout=2000, force=True)
        log("Closed popup 1 (Welcome to PropsEdge)")
    except PlaywrightTimeoutError:
        log("Popup 1 did not appear")

    # 2) After popup 1 attempt, wait for popup 2, then close it.
    announcement_dialog = page.locator("[role='dialog']").filter(
        has=page.locator("#announcement-title")
    ).first
    try:
        announcement_dialog.wait_for(state="visible", timeout=7000)
        popup2_closed = False
        for _ in range(4):
            got_it = announcement_dialog.get_by_role("button", name="Got it!", exact=True).first
            close_btn = announcement_dialog.get_by_role("button", name="Close", exact=True).first

            try:
                if got_it.is_visible(timeout=500):
                    got_it.click(timeout=1500, force=True)
                elif close_btn.is_visible(timeout=500):
                    close_btn.click(timeout=1500, force=True)
                else:
                    page.keyboard.press("Escape")
            except PlaywrightTimeoutError:
                page.keyboard.press("Escape")

            try:
                announcement_dialog.wait_for(state="hidden", timeout=1500)
                popup2_closed = True
                break
            except PlaywrightTimeoutError:
                page.wait_for_timeout(250)

        if popup2_closed:
            log("Closed popup 2 (New Tutorial Video)")
        else:
            log("Popup 2 detected but did not close")
    except PlaywrightTimeoutError:
        log("Popup 2 did not appear")

    # 3) Premium Discord prompt can block every control behind its backdrop.
    discord_community_card = page.locator("[data-testid='discord-community-card']").first
    try:
        discord_community_card.wait_for(state="visible", timeout=3000)
        skip_button = discord_community_card.get_by_role(
            "button", name="Skip for now", exact=True
        ).first
        skip_button.click(timeout=2000, force=True)
        discord_community_card.wait_for(state="hidden", timeout=3000)
        log("Closed premium Discord prompt")
    except PlaywrightTimeoutError:
        if discord_community_card.count() > 0:
            log("Premium Discord prompt detected but did not close")
        else:
            log("Premium Discord prompt did not appear")

    # 4) PropsEdge product tour can appear after dismissing the Discord prompt.
    main_lines_tour = page.locator("div").filter(
        has=page.get_by_text("Check out Main Lines", exact=True)
    ).filter(
        has=page.get_by_text("Go to Main Lines", exact=True)
    ).first
    try:
        main_lines_tour.wait_for(state="visible", timeout=3000)
        closed = False
        maybe_later = main_lines_tour.get_by_text("Maybe later", exact=True).first
        close_btn = main_lines_tour.get_by_label("Close", exact=True).first

        for target in (maybe_later, close_btn):
            try:
                if target.count() > 0 and target.is_visible(timeout=500):
                    target.click(timeout=1500, force=True)
                    closed = True
                    break
            except PlaywrightTimeoutError:
                continue

        if not closed:
            page.keyboard.press("Escape")

        try:
            main_lines_tour.wait_for(state="hidden", timeout=2500)
            log("Closed Main Lines product tour")
        except PlaywrightTimeoutError:
            log("Main Lines product tour detected but did not close")
    except PlaywrightTimeoutError:
        log("Main Lines product tour did not appear")

    # 5) Older linked-account prompt can block Settings/filters.
    discord_modal = page.locator("div").filter(
        has=page.get_by_text("Link Your Discord to PropsEdge!", exact=True)
    ).filter(
        has=page.get_by_text("Link Discord Now", exact=True)
    ).first
    try:
        discord_modal.wait_for(state="visible", timeout=3000)
        closed = False
        for target in (
            discord_modal.get_by_role("button", name="Got it!", exact=True).first,
            discord_modal.get_by_label("Close", exact=True).first,
        ):
            try:
                if target.count() > 0 and target.is_visible(timeout=500):
                    target.click(timeout=1500, force=True)
                    closed = True
                    break
            except PlaywrightTimeoutError:
                continue

        if not closed:
            page.keyboard.press("Escape")

        try:
            discord_modal.wait_for(state="hidden", timeout=2500)
            log("Closed Discord link prompt")
        except PlaywrightTimeoutError:
            log("Discord link prompt detected but did not close")
    except PlaywrightTimeoutError:
        log("Discord link prompt did not appear")

    # Wait for backdrop to clear so next click is not intercepted.
    backdrop = page.locator("div.fixed.inset-0.z-\\[100\\]").first
    try:
        backdrop.wait_for(state="hidden", timeout=5000)
    except PlaywrightTimeoutError:
        page.wait_for_timeout(800)


def get_visible_results_count(page, log, timeout_ms: int = 15000) -> int:
    count_locator = page.locator("div.mt-8.mb-4.text-sm.text-center.text-gray-400").first
    count_locator.wait_for(state="visible", timeout=timeout_ms)
    text = (count_locator.inner_text() or "").strip()
    match = re.search(r"Showing\s+(\d+)\s+of\s+(\d+)\s+projections", text, re.IGNORECASE)
    if not match:
        log(f"Could not parse projections count from: {text}")
        return 0
    visible_count = int(match.group(1))
    total_count = int(match.group(2))
    log(f"Count for current prop: showing {visible_count} of {total_count}")
    return visible_count


def parse_pct_value(value: str) -> float | None:
    text = str(value or "").strip()
    if not text:
        return None
    if text.endswith("%"):
        text = text[:-1].strip()
    try:
        return float(text)
    except ValueError:
        return None


def parse_decimal_value(value: str | float | int | None) -> float | None:
    text = str(value or "").strip()
    if not text:
        return None
    if text.endswith("%"):
        text = text[:-1].strip()
    text = text.replace(",", "")
    try:
        return float(text)
    except ValueError:
        return None


def average_pct_text(values: list[str]) -> str:
    parsed = [v for v in (parse_pct_value(x) for x in values) if v is not None]
    if not parsed:
        return ""
    avg = sum(parsed) / len(parsed)
    if abs(avg - round(avg)) < 1e-9:
        return f"{int(round(avg))}%"
    return f"{avg:.2f}%"


def _db_enabled() -> bool:
    return _env_enabled("DB_ENABLE", default=False)


def _db_connect():
    if not _db_enabled():
        return None
    if psycopg2 is None:
        raise RuntimeError("psycopg2 not installed but DB_ENABLE=true")
    host = (os.getenv("DB_HOST") or "").strip()
    port = int((os.getenv("DB_PORT") or "5432").strip())
    name = (os.getenv("DB_NAME") or "").strip()
    user = (os.getenv("DB_USER") or "").strip()
    password = (os.getenv("DB_PASSWORD") or "").strip().strip("'").strip('"')
    sslmode = (os.getenv("DB_SSLMODE") or "require").strip()
    if not host or not name or not user or not password:
        raise RuntimeError("Missing DB_HOST/DB_NAME/DB_USER/DB_PASSWORD")
    conn = psycopg2.connect(
        host=host,
        port=port,
        dbname=name,
        user=user,
        password=password,
        sslmode=sslmode,
        connect_timeout=10,
    )
    conn.autocommit = False
    return conn


def db_insert_run_start(log, mode: str, sport_cfg: dict) -> int | None:
    if not _db_enabled():
        return None
    conn = None
    try:
        conn = _db_connect()
        with conn.cursor() as cur:
            cur.execute(
                """
                INSERT INTO runs (source_bot, mode, sport, status)
                VALUES (%s, %s, %s, %s)
                RETURNING id
                """,
                ("propsedge", mode, sport_cfg["key"], "running"),
            )
            run_id = cur.fetchone()[0]
        conn.commit()
        log(f"DB run started (id={run_id}, mode={mode})")
        return run_id
    except Exception as exc:
        if conn:
            conn.rollback()
        log(f"DB run start failed: {exc}")
        return None
    finally:
        if conn:
            conn.close()


def db_update_run_finish(log, run_id: int | None, status: str, error_message: str | None = None) -> None:
    if not _db_enabled() or not run_id:
        return
    conn = None
    try:
        conn = _db_connect()
        with conn.cursor() as cur:
            cur.execute(
                """
                UPDATE runs
                SET status = %s, finished_at = NOW(), error_message = %s
                WHERE id = %s
                """,
                (status, error_message, run_id),
            )
        conn.commit()
        log(f"DB run finished (id={run_id}, status={status})")
    except Exception as exc:
        if conn:
            conn.rollback()
        log(f"DB run finish update failed (id={run_id}): {exc}")
    finally:
        if conn:
            conn.close()


def db_insert_props_rows(log, run_id: int | None, output_suffix: str, rows: list[dict], sport_cfg: dict) -> None:
    if not _db_enabled() or not rows:
        return
    conn = None
    mode = "demons" if output_suffix == "demons" else "goblins"
    try:
        conn = _db_connect()
        payload = []
        for row in rows:
            selected_prop = (row.get("selected_prop") or "").strip()
            category = selected_prop
            for prefix in ("Player ", "Batter ", "Pitcher "):
                if category.startswith(prefix):
                    category = category.replace(prefix, "", 1).strip()
                    break
            payload.append(
                (
                    run_id,
                    "propsedge",
                    sport_cfg["key"],
                    sport_cfg["label"],
                    mode,
                    category,
                    selected_prop,
                    row.get("player_name") or "",
                    row.get("player_id") or None,
                    row.get("prop") or "",
                    parse_decimal_value(row.get("prop_line")),
                    parse_pct_value(row.get("L5")),
                    parse_pct_value(row.get("L10")),
                    parse_pct_value(row.get("L20")),
                    parse_pct_value(row.get("H2H")),
                    parse_pct_value(row.get("SZN")),
                    parse_pct_value(row.get("AVG_3")),
                )
            )

        with conn.cursor() as cur:
            cur.executemany(
                """
                INSERT INTO props_rows (
                    run_id, source_bot, sport, league, mode, category, selected_prop,
                    player_name, player_id, prop, prop_line, l5, l10, l20, h2h, szn, model_avg
                )
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """,
                payload,
            )
        conn.commit()
        log(f"DB inserted {len(payload)} props rows for mode={mode}")
    except Exception as exc:
        if conn:
            conn.rollback()
        log(f"DB props row insert failed: {exc}")
    finally:
        if conn:
            conn.close()


def get_google_services():
    service_account_path = resolve_service_account_path()
    creds = service_account.Credentials.from_service_account_file(
        str(service_account_path), scopes=GOOGLE_SCOPES
    )
    sheets_service = build("sheets", "v4", credentials=creds)
    drive_service = build("drive", "v3", credentials=creds)
    return sheets_service, drive_service


def ensure_sheet_tab(sheets_service, sheet_id: str, tab_name: str) -> None:
    meta = sheets_service.spreadsheets().get(spreadsheetId=sheet_id).execute()
    titles = {s["properties"]["title"] for s in meta.get("sheets", [])}
    if tab_name in titles:
        return
    sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=sheet_id,
        body={"requests": [{"addSheet": {"properties": {"title": tab_name}}}]},
    ).execute()


def update_sheet_from_csv(sheets_service, sheet_id: str, tab_name: str, csv_path: Path) -> None:
    with csv_path.open("r", newline="", encoding="utf-8") as f:
        rows = list(csv.reader(f))
    if not rows:
        rows = [["No data"]]
    ensure_sheet_tab(sheets_service, sheet_id, tab_name)
    sheets_service.spreadsheets().values().clear(
        spreadsheetId=sheet_id, range=f"{tab_name}!A:Z"
    ).execute()
    sheets_service.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range=f"{tab_name}!A1",
        valueInputOption="RAW",
        body={"values": rows},
    ).execute()


def upsert_file_to_drive_folder(drive_service, folder_id: str, local_file: Path) -> None:
    filename = local_file.name
    q = (
        f"name = '{filename}' and '{folder_id}' in parents and trashed = false"
    )
    resp = drive_service.files().list(
        q=q,
        fields="files(id,name)",
        pageSize=10,
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()
    files = resp.get("files", [])
    media = MediaFileUpload(str(local_file), mimetype="text/csv", resumable=False)
    if files:
        file_id = files[0]["id"]
        drive_service.files().update(
            fileId=file_id,
            media_body=media,
            supportsAllDrives=True,
        ).execute()
    else:
        drive_service.files().create(
            body={"name": filename, "parents": [folder_id]},
            media_body=media,
            fields="id",
            supportsAllDrives=True,
        ).execute()


def sync_outputs_to_google(log, output_suffix: str) -> None:
    csv_path = BASE_DIR / f"projections_extract_all_props_{output_suffix}.csv"
    if not csv_path.exists():
        log(f"Skip Google sync, file not found: {csv_path.name}")
        return
    tab_name = "Demons" if output_suffix == "demons" else "Goblins"
    sheets_service, drive_service = get_google_services()
    update_sheet_from_csv(sheets_service, SHEET_ID, tab_name, csv_path)
    log(f"Updated Google Sheet tab '{tab_name}' from {csv_path.name}")
    upsert_file_to_drive_folder(drive_service, DRIVE_FOLDER_ID, csv_path)
    log(f"Uploaded/updated Drive CSV in folder for {csv_path.name}")


def control_projection_scroller(page, sport_cfg: dict, action: str, delta: int = 0) -> dict:
    """
    PropsEdge scrolls the projections page through #scroll-container.

    The projections rows themselves are virtualized. On desktop, the row wrapper is
    overflow-y-hidden and rows are absolutely positioned; scrolling #scroll-container
    causes the mounted data-index range to change.

    Playwright page.evaluate() accepts ONE optional argument, so all JS inputs are
    passed in a single dict.
    """
    try:
        return page.evaluate(
            """
            ({action, delta}) => {
                const scroller = document.querySelector("#scroll-container");
                const table = document.querySelector("#tour-projections-table");

                if (!scroller) {
                    return {
                        found: false,
                        source: "#scroll-container-missing",
                        scrollTop: -1,
                        scrollHeight: -1,
                        clientHeight: -1,
                        atBottom: false
                    };
                }

                const safeDelta = Number.isFinite(Number(delta))
                    ? Math.max(0, Number(delta))
                    : 0;

                const beforeTop = Number(scroller.scrollTop || 0);
                const scrollHeight = Number(scroller.scrollHeight || 0);
                const clientHeight = Number(scroller.clientHeight || 0);
                const maxTop = Math.max(0, scrollHeight - clientHeight);

                let targetTop = beforeTop;

                if (action === "top") {
                    if (table) {
                        const scrollerRect = scroller.getBoundingClientRect();
                        const tableRect = table.getBoundingClientRect();

                        // Absolute content offset of the projections table inside
                        // #scroll-container. This resets virtualized rows to their
                        // first range without scrolling the outer document/body.
                        const tableTop =
                            beforeTop + (tableRect.top - scrollerRect.top);
                        targetTop = Math.max(0, Math.min(maxTop, tableTop));
                    } else {
                        targetTop = 0;
                    }
                    scroller.scrollTo({top: targetTop, behavior: "auto"});
                    scroller.scrollTop = targetTop;
                } else if (action === "scroll" && safeDelta > 0) {
                    targetTop = Math.min(maxTop, beforeTop + safeDelta);
                    scroller.scrollTo({top: targetTop, behavior: "auto"});
                    scroller.scrollTop = targetTop;
                }

                // Read the real state after the write.
                const afterTop = Number(scroller.scrollTop || 0);
                const afterHeight = Number(scroller.scrollHeight || 0);
                const afterView = Number(scroller.clientHeight || 0);

                // Keep diagnostics selector simple. A previous selector used the
                // Tailwind class `md:block`; escaping that through Python -> JS -> CSS
                // caused querySelectorAll() to throw AFTER the scroll had already happened.
                const desktopIndexEls = Array.from(
                    document.querySelectorAll("#tour-projections-table a[data-index]")
                );
                const fallbackIndexEls = desktopIndexEls.length
                    ? desktopIndexEls
                    : Array.from(document.querySelectorAll("#tour-projections-table div[data-index]"));
                const allIndexes = fallbackIndexEls
                    .map(el => Number(el.getAttribute("data-index")))
                    .filter(Number.isFinite);

                return {
                    found: true,
                    source: "#scroll-container",
                    moved: afterTop - beforeTop,
                    scrollTop: afterTop,
                    scrollHeight: afterHeight,
                    clientHeight: afterView,
                    targetTop,
                    firstDataIndex: allIndexes.length ? Math.min(...allIndexes) : null,
                    lastDataIndex: allIndexes.length ? Math.max(...allIndexes) : null,
                    atBottom: afterTop + afterView >= afterHeight - 8
                };
            }
            """,
            {"action": action, "delta": int(delta)},
        )
    except Exception as exc:
        return {
            "found": False,
            "source": "evaluate-error",
            "error": str(exc),
            "scrollTop": -1,
            "scrollHeight": -1,
            "clientHeight": -1,
            "atBottom": False,
        }


def extract_projections_table(page, log, target_prop: str, expected_count: int, sport_cfg: dict) -> list[dict]:
    """
    Extract every virtualized desktop projection row without silently skipping indexes.

    PropsEdge's desktop table uses stable a[data-index] virtual rows. The expected
    count is therefore treated as indexes 0..expected_count-1. We collect rows by
    data-index, scroll with overlap, then revisit any missing index ranges directly.
    If rows are still missing after recovery, fail instead of saving partial output.
    """
    table = page.locator("#tour-projections-table").first
    table.wait_for(state="visible", timeout=15000)

    rows_selector = (
        f"#tour-projections-table a[data-index][href*='/home/{sport_cfg['href_slug']}/']"
    )
    log(f"Using indexed desktop row selector: {rows_selector}")

    seen_by_index: dict[int, dict] = {}

    def read_mounted_rows() -> list[dict]:
        """Take one atomic DOM snapshot so virtualization cannot detach locators mid-row."""
        try:
            return page.evaluate(
                """
                ({selector}) => Array.from(document.querySelectorAll(selector)).map((row) => {
                    const textOf = (el) => (el && el.textContent ? el.textContent : '').replace(/\\s+/g, ' ').trim();
                    const percentOf = (el) => {
                        const match = textOf(el).match(/\\b\\d+(?:\\.\\d+)?%/);
                        return match ? match[0] : '';
                    };

                    const href = row.getAttribute('href') || '';
                    const projectionRow = row.querySelector('#tour-projection-row');
                    const cells = projectionRow ? Array.from(projectionRow.children) : [];
                    const playerCell = row.querySelector('#tour-cell-player');

                    // PropsEdge changed the inner typography/classes while keeping the
                    // outer virtualized row structure. Prefer current selectors and keep
                    // old selectors as compatibility fallbacks.
                    const player =
                        (playerCell && playerCell.querySelector('span.text-sm.font-semibold.text-white.truncate')) ||
                        row.querySelector('span.text-base.font-medium.text-gray-100');
                    const playerImage = playerCell
                        ? Array.from(playerCell.querySelectorAll('img[alt]')).find(
                            (img) => img.getAttribute('alt') !== 'Team logo'
                        )
                        : null;
                    const lineSpan =
                        (playerCell && playerCell.querySelector('span.shrink-0.whitespace-nowrap.text-sm.font-semibold.text-gray-100')) ||
                        Array.from(row.querySelectorAll('span')).find(
                            (s) => /^(Over|Under)\\s+-?\\d+(?:\\.\\d+)?$/i.test(textOf(s))
                        );
                    const playerSpans = playerCell
                        ? Array.from(playerCell.querySelectorAll('span'))
                        : [];
                    const prop = lineSpan
                        ? playerSpans.slice(playerSpans.indexOf(lineSpan) + 1).find((s) => {
                            const value = textOf(s);
                            return value && value !== 'PE Score' && value !== 'Park';
                        })
                        : null;

                    let marketFromHref = '';
                    try {
                        marketFromHref = new URL(href, document.baseURI).searchParams.get('market') || '';
                    } catch (_) {}

                    // Read the trend cells directly. The row now contains a separate
                    // Chance percentage before L5, so scraping every % from rowText would
                    // shift the L5/L10/L20/H2H/SZN values.
                    const l5Cell = row.querySelector('#tour-cell-all-odds') || cells[3] || null;
                    const l10Cell = row.querySelector('#tour-cell-trends') || cells[4] || null;
                    const l20Cell = cells[5] || null;
                    const h2hCell = cells[6] || null;
                    const sznCell = cells[7] || null;

                    return {
                        dataIndex: Number(row.getAttribute('data-index')),
                        href,
                        playerName: textOf(player) || (playerImage ? playerImage.getAttribute('alt') || '' : ''),
                        propText: textOf(prop) || marketFromHref,
                        lineText: textOf(lineSpan),
                        L5: percentOf(l5Cell),
                        L10: percentOf(l10Cell),
                        L20: percentOf(l20Cell),
                        H2H: percentOf(h2hCell),
                        SZN: percentOf(sznCell),
                        rowText: (row.innerText || row.textContent || '').replace(/\\s+/g, ' ').trim(),
                    };
                }).filter((r) => Number.isFinite(r.dataIndex) && r.href)
                """,
                {"selector": rows_selector},
            ) or []
        except Exception as exc:
            log(f"Mounted-row snapshot failed for '{target_prop}': {exc}")
            return []

    def absorb_mounted_rows() -> tuple[int | None, int | None, int]:
        mounted = read_mounted_rows()
        indexes = []
        added = 0
        for item in mounted:
            try:
                data_index = int(item.get("dataIndex"))
            except (TypeError, ValueError):
                continue
            if data_index < 0:
                continue
            if expected_count > 0 and data_index >= expected_count:
                continue

            indexes.append(data_index)
            if data_index in seen_by_index:
                continue

            href = str(item.get("href") or "")
            player_name = str(item.get("playerName") or "").strip()
            prop_text = str(item.get("propText") or "").strip()
            line_text = str(item.get("lineText") or "").strip()
            row_text = str(item.get("rowText") or "")

            # A mounted row can briefly exist before React finishes hydrating it.
            # Do not mark it collected until the fields we actually need are present.
            if not href or not player_name or not prop_text:
                continue

            trends = [
                str(item.get("L5") or "").strip(),
                str(item.get("L10") or "").strip(),
                str(item.get("L20") or "").strip(),
                str(item.get("H2H") or "").strip(),
                str(item.get("SZN") or "").strip(),
            ]
            # Compatibility fallback for the older row layout. Only use this when
            # none of the explicit trend cells were available.
            if not any(trends):
                percent_values = re.findall(r"\b\d+%", row_text)
                trends = (percent_values + ["", "", "", "", ""])[:5]
            line_match = re.search(
                r"\b(?:Over|Under)\s+([0-9]+(?:\.[0-9]+)?)\b",
                line_text,
                re.IGNORECASE,
            )
            if not line_match:
                # href contains the exact projection line too; use it as a stable fallback.
                line_match = re.search(r"[?&]line=([0-9]+(?:\.[0-9]+)?)", href, re.IGNORECASE)
            prop_line = line_match.group(1) if line_match else ""

            player_id_match = re.search(
                rf"/home/{re.escape(sport_cfg['href_slug'])}/([^/?#]+)",
                href,
                re.IGNORECASE,
            )
            player_id = player_id_match.group(1) if player_id_match else ""
            player_img_url = (
                f"https://player-images.propsedge.io/{sport_cfg['image_slug']}/{player_id}/high.png"
                if player_id else ""
            )

            avg_3 = average_pct_text([trends[4], trends[1], trends[3]])
            seen_by_index[data_index] = {
                "selected_prop": target_prop,
                "player_name": player_name,
                "player_id": player_id,
                "player_img_url": player_img_url,
                "prop": prop_text,
                "prop_line": prop_line,
                "L5": trends[0],
                "L10": trends[1],
                "L20": trends[2],
                "H2H": trends[3],
                "SZN": trends[4],
                "AVG_3": avg_3,
            }
            added += 1

        return (
            min(indexes) if indexes else None,
            max(indexes) if indexes else None,
            added,
        )

    def scroll_to_index(data_index: int) -> dict:
        """Position #scroll-container near an exact virtual row index."""
        try:
            return page.evaluate(
                """
                ({index}) => {
                    const scroller = document.querySelector('#scroll-container');
                    const table = document.querySelector('#tour-projections-table');
                    if (!scroller || !table) return {found:false};

                    const rows = Array.from(table.querySelectorAll('a[data-index]'))
                        .map((el) => ({
                            el,
                            index: Number(el.getAttribute('data-index')),
                            transform: el.style.transform || ''
                        }))
                        .filter((x) => Number.isFinite(x.index))
                        .sort((a,b) => a.index - b.index);

                    let pitch = 64; // exact current desktop row pitch; derive when possible.
                    if (rows.length >= 2) {
                        const y = (s) => {
                            const m = String(s || '').match(/translateY\\(([-0-9.]+)px\\)/i);
                            return m ? Number(m[1]) : NaN;
                        };
                        for (let i = 1; i < rows.length; i++) {
                            const y0 = y(rows[i-1].transform);
                            const y1 = y(rows[i].transform);
                            const di = rows[i].index - rows[i-1].index;
                            if (Number.isFinite(y0) && Number.isFinite(y1) && di > 0) {
                                const candidate = (y1 - y0) / di;
                                if (candidate > 20 && candidate < 300) {
                                    pitch = candidate;
                                    break;
                                }
                            }
                        }
                    }

                    const scrollerRect = scroller.getBoundingClientRect();
                    const tableRect = table.getBoundingClientRect();
                    const tableTop = Number(scroller.scrollTop || 0) + (tableRect.top - scrollerRect.top);
                    const header = table.querySelector('[data-sticky-header="projections-table"]');
                    const headerHeight = header ? Number(header.offsetHeight || 0) : 54;
                    const maxTop = Math.max(0, Number(scroller.scrollHeight || 0) - Number(scroller.clientHeight || 0));
                    const viewportLead = Math.max(120, Math.floor(Number(scroller.clientHeight || 0) * 0.28));
                    const target = Math.max(0, Math.min(
                        maxTop,
                        tableTop + headerHeight + (Number(index) * pitch) - viewportLead
                    ));
                    const before = Number(scroller.scrollTop || 0);
                    scroller.scrollTo({top: target, behavior:'auto'});
                    scroller.scrollTop = target;
                    scroller.dispatchEvent(new Event('scroll', {bubbles:true}));
                    return {
                        found:true,
                        source:'#scroll-container-index',
                        index:Number(index),
                        pitch,
                        before,
                        scrollTop:Number(scroller.scrollTop || 0),
                        target,
                        maxTop,
                    };
                }
                """,
                {"index": int(data_index)},
            )
        except Exception as exc:
            return {"found": False, "source": "index-scroll-error", "error": str(exc)}

    def missing_indexes() -> list[int]:
        if expected_count <= 0:
            return []
        return [i for i in range(expected_count) if i not in seen_by_index]

    # Start at the first projection row and collect with heavy overlap.
    control_projection_scroller(page, sport_cfg, "top")
    page.wait_for_timeout(250)

    max_rounds = 500
    last_count = -1
    no_progress_rounds = 0

    for round_idx in range(max_rounds):
        first_idx, last_idx, added = absorb_mounted_rows()

        if expected_count > 0 and len(seen_by_index) >= expected_count:
            break

        if len(seen_by_index) == last_count:
            no_progress_rounds += 1
        else:
            no_progress_rounds = 0
        last_count = len(seen_by_index)

        scroll_state = control_projection_scroller(page, sport_cfg, "check")
        at_bottom = bool(scroll_state.get("found")) and bool(scroll_state.get("atBottom", False))

        if round_idx % 10 == 0:
            log(
                f"Extraction progress '{target_prop}': collected={len(seen_by_index)}/{expected_count} "
                f"mounted={first_idx}-{last_idx} added={added} "
                f"scroll_top={scroll_state.get('scrollTop', '?')}"
            )

        if at_bottom and no_progress_rounds >= 2:
            break

        # 640px is ~10 desktop rows at the observed 64px pitch while roughly 21 rows
        # are mounted, so successive snapshots overlap by about half a viewport.
        control_projection_scroller(page, sport_cfg, "scroll", 640)
        page.wait_for_timeout(180)

    # Recovery pass: completeness is based on the actual virtual indexes, not on
    # whether we happened to reach the bottom. Revisit every missing range.
    if expected_count > 0 and len(seen_by_index) < expected_count:
        for recovery_pass in range(1, 5):
            missing = missing_indexes()
            if not missing:
                break
            log(
                f"Recovery pass {recovery_pass} for '{target_prop}': "
                f"missing {len(missing)} indexes"
            )

            # Visit starts throughout the missing set. Spacing by 8 indexes keeps
            # strong overlap with the ~21 mounted desktop rows.
            targets = []
            last_target = -999
            for idx in missing:
                if idx - last_target >= 8:
                    targets.append(idx)
                    last_target = idx
            if missing[-1] not in targets:
                targets.append(missing[-1])

            before_recovery = len(seen_by_index)
            for target_idx in targets:
                state = scroll_to_index(target_idx)
                if not state.get("found"):
                    log(
                        f"Index recovery scroll failed at {target_idx}: "
                        f"{state.get('error', state.get('source', 'unknown'))}"
                    )
                    continue
                page.wait_for_timeout(220)
                absorb_mounted_rows()
                if len(seen_by_index) >= expected_count:
                    break

            if len(seen_by_index) == before_recovery:
                break

    if expected_count > 0:
        missing = missing_indexes()
        if missing:
            preview = ",".join(str(i) for i in missing[:30])
            if len(missing) > 30:
                preview += ",..."
            raise RuntimeError(
                f"Incomplete extraction for '{target_prop}': collected "
                f"{len(seen_by_index)}/{expected_count}; missing data-index(es): {preview}"
            )

    # data-index is the canonical table order.
    data = [seen_by_index[i] for i in sorted(seen_by_index)]
    log(f"Extracted all {len(data)} rows from projections table for '{target_prop}'")
    return data


def select_props_filter(page, log, prop_name: str, all_target_props: list[str]) -> bool:
    target_text = prop_name.strip()

    def panel_candidates():
        return [
            page.locator("div#filter-panel-animation-wrapper"),
            page.locator("div[data-sticky-filters='projections-filters']"),
            page.locator("div.fixed.right-0.top-0.bottom-0"),
            page.locator("div.absolute.right-0.top-0.bottom-0"),
            page.locator("section"),
            page.locator("body"),
        ]

    def collect_candidates(container):
        selectors = [
            "button.grid.w-full.outline-0",
            "button.grid.w-full",
        ]
        found = []
        for sel in selectors:
            try:
                btns = container.locator(sel).all()
            except Exception:
                btns = []
            if btns:
                found = btns
                break
        return found

    def wait_for_open_panel() -> tuple[list, any]:
        for _ in range(24):
            for cont in panel_candidates():
                try:
                    if not cont.count():
                        continue
                    cands = collect_candidates(cont)
                    if cands:
                        # keep only rows that at least contain the prop suffix pattern or a count suffix
                        filtered = [
                            c for c in cands
                            if re.sub(r"\s+\d[\d,]*$", "", (c.inner_text() or "").strip()).strip()
                        ]
                        if filtered:
                            return filtered, cont
                except Exception:
                    continue
            page.wait_for_timeout(120)
        return [], page.locator("body").first

    def open_props_menu() -> bool:
        log(f"Opening props menu for '{target_text}'")
        for _ in range(12):
            candidates = [
                page.get_by_role("button", name=re.compile(r"^Props filter", re.IGNORECASE)).first,
                page.locator("button").filter(has_text=re.compile(r"^Props$", re.IGNORECASE)).first,
                page.locator("button").filter(has_text=re.compile(r"^Props filter,", re.IGNORECASE)).first,
            ]
            for button in candidates:
                try:
                    if button.count() and button.is_visible():
                        button.scroll_into_view_if_needed()
                        button.click(timeout=3000)
                        page.wait_for_timeout(180)
                        log("Clicked props dropdown")
                        cands, _ = wait_for_open_panel()
                        if cands:
                            return True
                except Exception:
                    continue
            page.wait_for_timeout(200)
        return False

    def close_menu() -> None:
        close_buttons = [
            page.get_by_role("button", name=re.compile(r"^Props filter,\s*\d+\s*active$", re.IGNORECASE)).first,
            page.get_by_role("button", name=re.compile(r"^Props filter", re.IGNORECASE)).first,
        ]
        for btn in close_buttons:
            try:
                if btn.count() and btn.is_visible():
                    btn.click(timeout=1200)
                    page.wait_for_timeout(120)
                    log("Closed props menu")
                    return
            except Exception:
                continue
        try:
            page.keyboard.press("Escape")
        except Exception:
            pass
        page.wait_for_timeout(120)
        try:
            page.mouse.click(20, 20)
            page.wait_for_timeout(120)
            log("Closed props menu (fallback click)")
        except Exception:
            pass

    def find_panel():
        # try to return the first container that actually contains prop option rows
        for container in panel_candidates():
            try:
                cands, found = wait_for_open_panel()
                if cands:
                    for c in panel_candidates():
                        try:
                            if c.count() and c.locator("button.grid.w-full.outline-0").count() > 0:
                                return c
                        except Exception:
                            continue
                    return found
            except Exception:
                pass
        return page.locator("body").first

    def option_has_name(option, target: str) -> bool:
        try:
            name = option.locator(".grow").first.inner_text().replace("\xa0", " ")
            if not name.strip():
                name = option.inner_text()
            name = name.replace("\xa0", " ")
            name = re.sub(r"\s+", " ", name).strip()
            # UI often appends row counts like "Batter Total Bases 692"
            name = re.sub(r"\s+\d[\d,]*$", "", name).strip()
            return re.sub(r"\s+", " ", name).strip().lower() == target.lower()
        except Exception:
            return False

    def click_option(option) -> bool:
        # Click with JS to avoid viewport-limited locator clicks on virtualized rows.
        # The menu is a long virtualized panel and some rows are not considered
        # "in viewport" by Playwright even after page scroll.
        script = """
        (el) => {
            const isScrollable = (node) => {
                if (!node || !node.ownerDocument) return false;
                const style = window.getComputedStyle(node);
                const overflowY = (style.overflowY || '').toLowerCase();
                const overflow = (style.overflow || '').toLowerCase();
                return /(auto|scroll|overlay)/.test(overflowY) || /(auto|scroll|overlay)/.test(overflow);
            };
            let node = el.parentElement;
            while (node && node !== document.body && node !== document.documentElement) {
                if (isScrollable(node) && node.scrollHeight > node.clientHeight) {
                    const targetTop = Math.max(0, el.offsetTop - Math.max(40, Math.floor(node.clientHeight * 0.45)));
                    node.scrollTop = targetTop;
                }
                node = node.parentElement;
            }
            el.scrollIntoView({ block: 'center', inline: 'nearest' });
            el.dispatchEvent(new MouseEvent('pointerdown', { bubbles: true, cancelable: true, button: 0 }));
            el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, button: 0 }));
            el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, button: 0 }));
            el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, button: 0 }));
            return true;
        }
        """
        for _ in range(6):
            try:
                option.evaluate(script)
                return True
            except Exception:
                pass
            try:
                box = option.bounding_box()
                if not box:
                    break
                x = box["x"] + box["width"] / 2
                y = box["y"] + box["height"] / 2
                page.mouse.move(x, y)
                page.mouse.down()
                page.mouse.up()
                return True
            except Exception:
                page.wait_for_timeout(120)
                continue
        return False

    def option_is_selected(option) -> bool:
        try:
            aria_pressed = (option.get_attribute("aria-pressed") or "").strip().lower()
            if aria_pressed in {"true", "1", "yes"}:
                return True
        except Exception:
            pass
        try:
            aria_selected = (option.get_attribute("aria-selected") or "").strip().lower()
            if aria_selected in {"true", "1", "yes"}:
                return True
        except Exception:
            pass
        try:
            marker = option.locator("div.w-4").first
            marker_class = (marker.get_attribute("class") or "").lower()
            if "bg-purple-600" in marker_class or "border-purple-600" in marker_class:
                return True
        except Exception:
            pass
        try:
            icon = option.locator("svg").first
            if icon.count() and icon.is_visible():
                return True
        except Exception:
            pass
        return False

    def clear_other_selected_props(panel, keep_text: str) -> None:
        try:
            current_options = panel.locator("button.grid.w-full.outline-0").all()
        except Exception:
            return
        for option in current_options:
            if not option.count():
                continue
            if option_is_selected(option) and not option_has_name(option, keep_text):
                try:
                    clicked_raw = option.inner_text().replace("\xa0", " ")
                    clicked_raw = re.sub(r"\s+", " ", clicked_raw).strip()
                    if click_option(option):
                        log(f"Cleared prior selected prop '{clicked_raw}'")
                        page.wait_for_timeout(100)
                except Exception:
                    pass

    def locate_menu_scroller(container):
        # Prefer closest explicit scroller from button ancestors; fallback to explicit wrappers.
        candidates = [
            "div.overflow-y-auto",
            "div[style*='overflow: auto']",
            "div[style*='overflow-y: auto']",
            "section",
            "div.fixed.right-0.top-0.bottom-0",
            "div#filter-panel-animation-wrapper",
            "div[data-sticky-filters='projections-filters']",
        ]
        for c in candidates:
            try:
                loc = container.locator(c).first
                if loc.count() and loc.evaluate("el => (el.scrollHeight - el.clientHeight) > 2"):
                    return loc
            except Exception:
                continue

        # Last resort: first scrollable-ish ancestor via JS
        try:
            return container.locator("xpath=ancestor::*").first
        except Exception:
            return None

    def pick_in_panel(panel, target: str):
        # Props rows are plain buttons in current UI, with trailing row counts in the label.
        candidates = panel.locator("button.grid.w-full.outline-0").all()
        if not candidates:
            # Fallback: direct text match on buttons in the whole page when the panel
            # container selectors miss the virtualized/animated menu rows.
            try:
                direct = page.locator("button").filter(
                    has_text=re.compile(
                        rf"^\s*{re.escape(target)}(?:\s+\d[\d,]*)?\s*$",
                        re.IGNORECASE,
                    )
                ).all()
                if direct:
                    candidates = direct
                    log(f"Props menu direct text-match candidates: {len(candidates)}")
            except Exception:
                pass
        try:
            log(f"Props menu options visible: {len(candidates)}")
        except Exception:
            pass

        for idx, c in enumerate(candidates):
            try:
                raw = c.inner_text().replace("\xa0", " ")
                raw = re.sub(r"\s+", " ", raw).strip()
                # UI often appends row counts like "Batter Total Bases 692"
                cleaned = re.sub(r"\s+\d[\d,]*$", "", raw).strip()
                log(f"Option[{idx}] raw='{raw}' cleaned='{cleaned}'")
            except Exception:
                continue

        for c in candidates:
            if option_has_name(c, target):
                log(f"Matched prop option '{target}'")
                if not click_option(c):
                    continue
                try:
                    clicked_text = c.inner_text().replace("\xa0", " ")
                    clicked_text = re.sub(r"\s+", " ", clicked_text).strip()
                except Exception:
                    clicked_text = target
                log(f"Clicked prop option raw='{clicked_text}' cleaned='{target}'")
                return c

        # fallback: global scan to cover slow render paths
        for cont in panel_candidates():
            try:
                all_buttons = collect_candidates(cont)
                for c in all_buttons:
                    if option_has_name(c, target):
                        log(f"Matched prop option '{target}'")
                        if not click_option(c):
                            continue
                        try:
                            clicked_text = c.inner_text().replace("\xa0", " ")
                            clicked_text = re.sub(r"\s+", " ", clicked_text).strip()
                        except Exception:
                            clicked_text = target
                        log(f"Clicked prop option raw='{clicked_text}' cleaned='{target}'")
                        return c
            except Exception:
                continue

        # fallback: scroll inside discovered panel scroller and retry
        for c in candidates:
            try:
                break
            except Exception:
                continue

        # detect first matched-row-like button (for scroll context)
        sample = panel.locator("button.grid.w-full.outline-0").first
        if sample.count() == 0:
            return None

        scroller = locate_menu_scroller(sample)
        if scroller is None or not scroller.count():
            return None

        for _ in range(40):
            try:
                scroller.evaluate("el => { if (el.scrollTop !== undefined) el.scrollTop += 260; }")
            except Exception:
                try:
                    page.mouse.wheel(0, 260)
                except Exception:
                    pass
            page.wait_for_timeout(200)
            candidates = panel.locator("button.grid.w-full.outline-0").all()
            for c in candidates:
                if option_has_name(c, target):
                    log(f"Matched prop option '{target}'")
                    if not click_option(c):
                        continue
                    try:
                        clicked_text = c.inner_text().replace("\xa0", " ")
                        clicked_text = re.sub(r"\s+", " ", clicked_text).strip()
                    except Exception:
                        clicked_text = target
                    log(f"Clicked prop option raw='{clicked_text}' cleaned='{target}'")
                    return c

        return None

    if not open_props_menu():
        log("Failed to open props selector")
        return False

    panel = find_panel()
    if not all_target_props:
        return False

    clear_other_selected_props(panel, target_text)

    option = pick_in_panel(panel, target_text)
    if option is None:
        # one retry: close/open menu once in case of stale render
        if not open_props_menu():
            log(f"Prop '{prop_name}' not available; skipping")
            return False
        panel = find_panel()
        option = pick_in_panel(panel, target_text)

    if option is None:
        log(f"Prop '{prop_name}' not available; skipping")
        return False

    close_menu()
    page.wait_for_timeout(120)
    log(f"Selected '{prop_name}'")
    return True


def reset_table_to_top_and_sort(page, log, sport_cfg: dict) -> None:
    projections_table = page.locator("#tour-projections-table").first

    # Fast path: do not reload unless table is actually missing.
    if projections_table.count() > 0:
        try:
            projections_table.wait_for(state="visible", timeout=5000)
            control_projection_scroller(page, sport_cfg, "top")
            page.wait_for_timeout(250)
            szn_button = page.get_by_role("button", name="SZN").first
            szn_label = szn_button.locator("span").first
            szn_class = szn_label.get_attribute("class") or ""
            if "text-purple-400" not in szn_class:
                szn_button.click(timeout=5000)
                log("Set SZN sort after reset")
            else:
                log("SZN sort already active after reset")
            try:
                page.wait_for_load_state("networkidle", timeout=4000)
            except PlaywrightTimeoutError:
                page.wait_for_timeout(400)
            return
        except PlaywrightTimeoutError:
            pass

    # Slow recovery path: reload then rehydrate projections table.
    page.reload(wait_until="domcontentloaded")
    try:
        page.wait_for_load_state("networkidle", timeout=10000)
    except PlaywrightTimeoutError:
        page.wait_for_timeout(800)

    projections_table = page.locator("#tour-projections-table").first
    def ensure_table_visible() -> bool:
        nonlocal projections_table
        if projections_table.count() > 0:
            try:
                projections_table.wait_for(state="visible", timeout=8000)
                return True
            except PlaywrightTimeoutError:
                pass

        show_results_btn = page.get_by_role(
            "button", name=re.compile(r"^Show [\d,]+ results$", re.IGNORECASE)
        ).first
        if show_results_btn.count() > 0:
            try:
                show_results_btn.click(timeout=5000, force=True)
            except Exception:
                return False
            try:
                page.wait_for_load_state("networkidle", timeout=10000)
            except PlaywrightTimeoutError:
                page.wait_for_timeout(900)
            projections_table = page.locator("#tour-projections-table").first
            if projections_table.count() > 0:
                try:
                    projections_table.wait_for(state="visible", timeout=12000)
                    return True
                except PlaywrightTimeoutError:
                    return False
        return False

    if not ensure_table_visible():
        # Recovery path: explicit projections navigation, then retry table hydration.
        log("Table not visible after reload; recovering via projections page")
        page.goto("https://www.propsedge.io/home/projections", wait_until="domcontentloaded")
        try:
            page.wait_for_load_state("networkidle", timeout=10000)
        except PlaywrightTimeoutError:
            page.wait_for_timeout(900)
        if not ensure_table_visible():
            raise PlaywrightTimeoutError("Projections table not visible after reload recovery")

    control_projection_scroller(page, sport_cfg, "top")
    page.wait_for_timeout(300)
    szn_button = page.get_by_role("button", name="SZN").first
    szn_label = szn_button.locator("span").first
    szn_class = szn_label.get_attribute("class") or ""
    szn_selected = "text-purple-400" in szn_class
    if not szn_selected:
        szn_button.click(timeout=5000)
        log("Set SZN sort after reload")
    else:
        log("SZN sort already active after reload")
    try:
        page.wait_for_load_state("networkidle", timeout=6000)
    except PlaywrightTimeoutError:
        page.wait_for_timeout(500)


def click_sport_and_scrape(
    page,
    log,
    creature_toggle: str,
    output_suffix: str,
    sport_cfg: dict,
    run_id: int | None = None,
) -> None:
    sport_selector = page.locator("#tour-sport-selector").first
    sport_selector.wait_for(state="visible", timeout=15000)

    sport_tab = sport_selector.locator(".sport-selector-item").filter(
        has_text=re.compile(rf"\b{re.escape(sport_cfg['label'])}\b", re.IGNORECASE)
    ).first
    sport_tab.wait_for(state="attached", timeout=10000)

    sport_class = sport_tab.get_attribute("class") or ""
    if "disabled" in sport_class:
        raise RuntimeError(f"{sport_cfg['label']} tab is disabled/not in season on PropsEdge.")

    if "selected" in sport_class:
        log(f"{sport_cfg['label']} already selected on projections")
    else:
        sport_tab.scroll_into_view_if_needed(timeout=5000)
        sport_tab.click(timeout=8000, force=True)
        log(f"Clicked {sport_cfg['label']} on projections")

    try:
        page.wait_for_load_state("domcontentloaded", timeout=10000)
        page.wait_for_load_state("networkidle", timeout=10000)
    except PlaywrightTimeoutError:
        log(f"Load wait after {sport_cfg['label']} click timed out; continuing")

    settings_btn = page.get_by_role("button", name="Settings").first
    settings_btn.wait_for(state="visible", timeout=15000)
    try:
        settings_btn.click(timeout=15000)
    except PlaywrightTimeoutError:
        settings_btn.scroll_into_view_if_needed(timeout=5000)
        settings_btn.click(timeout=15000, force=True)
    log("Clicked Settings")

    settings_panel = page.locator("#tour-settings-panel").first
    settings_panel.wait_for(state="visible", timeout=10000)

    def enable_toggle(toggle_name: str, required: bool = True) -> bool:
        def _make_candidates() -> list:
            stripped = toggle_name.replace(" only", "").strip()
            alias_terms = [
                toggle_name,
                stripped,
                stripped.rstrip("s") if stripped.endswith("s") else stripped,
            ]
            terms = list(dict.fromkeys(t for t in alias_terms if t and t.strip()))
            for term in list(dict.fromkeys(t.strip() for t in terms if t.strip())):
                escaped = re.escape(term)
                terms.append(term.title())
                terms.append(term.lower())
                # Also try a compact pattern for cases where UI omits `only`.
                yield_candidates = [
                    settings_panel.get_by_role("checkbox", name=re.compile(escaped, re.IGNORECASE)).first,
                    settings_panel.get_by_role("option", name=re.compile(escaped, re.IGNORECASE)).first,
                    settings_panel.get_by_role("button", name=re.compile(escaped, re.IGNORECASE)).first,
                    settings_panel.get_by_role("switch", name=re.compile(escaped, re.IGNORECASE)).first,
                    settings_panel.locator("label").filter(has_text=re.compile(escaped, re.IGNORECASE)).first,
                    settings_panel.locator("button").filter(has_text=re.compile(escaped, re.IGNORECASE)).first,
                    settings_panel.locator("div").filter(has_text=re.compile(escaped, re.IGNORECASE)).filter(
                        has=settings_panel.locator("input[type='checkbox']")
                    ).first,
                ]
                for candidate in yield_candidates:
                    yield candidate

        def _coalesce_locator(*locators):
            for loc in locators:
                try:
                    if loc.count() > 0:
                        return loc
                except Exception:
                    continue
            return None

        seen_terms = {""}
        row = None
        for candidate in _make_candidates():
            if candidate is None:
                continue
            try:
                if candidate.count() == 0:
                    continue
            except Exception:
                continue
            try:
                text = (candidate.inner_text(timeout=500) or "").strip()
            except Exception:
                continue
            if not text:
                continue
            candidate_key = text.lower()
            if candidate_key in seen_terms:
                continue
            seen_terms.add(candidate_key)
            if candidate.count() > 0:
                row = candidate
                break

        if row is None:
            # broad fallback: any visible control row containing the term and a checkbox.
            fallback_term = re.compile(toggle_name.split()[0], re.IGNORECASE)
            rows = settings_panel.locator("label,div,button").filter(has_text=fallback_term)
            if rows.count() > 0:
                # Modifiers and toggle controls are sometimes wrapped in non-semantic div/button
                # rows with a checkbox in the same container.
                filtered_rows = rows.filter(
                    has=settings_panel.locator("input[type='checkbox']")
                )
                if filtered_rows.count() > 0:
                    row = filtered_rows.first

        if row is None:
            if required:
                raise RuntimeError(f"{toggle_name} toggle control not found")
            log(f"{toggle_name} not available; continuing")
            return False

        checkbox = row.locator("input[type='checkbox']").first
        control_role = (row.get_attribute("role") or "").strip().lower()
        is_option = control_role == "option"
        has_checkbox = checkbox.count() > 0
        if not has_checkbox and is_option:
            checkbox = row

        try:
            row.wait_for(state="visible", timeout=5000)
        except PlaywrightTimeoutError:
            if required:
                raise
            log(f"{toggle_name} not available; continuing")
            return False

        enabled = False
        for _ in range(4):
            if has_checkbox:
                enabled = bool(checkbox.evaluate("el => el.checked"))
            else:
                if is_option:
                    enabled = (str(row.get_attribute("aria-selected") or "").lower() == "true")
                else:
                    enabled = bool(checkbox.evaluate("el => el.checked"))
            if enabled:
                break

            try:
                checkbox.check(timeout=1500, force=True) if has_checkbox else row.click(timeout=1500, force=True)
            except PlaywrightTimeoutError:
                pass
            page.wait_for_timeout(250)

            if has_checkbox:
                enabled = bool(checkbox.evaluate("el => el.checked"))
            elif is_option:
                enabled = (str(row.get_attribute("aria-selected") or "").lower() == "true")
            if enabled:
                break

            try:
                row.locator("label").first.click(timeout=1500, force=True)
            except PlaywrightTimeoutError:
                pass
            page.wait_for_timeout(200)

            if has_checkbox:
                enabled = bool(checkbox.evaluate("el => el.checked"))
            elif is_option:
                enabled = (str(row.get_attribute("aria-selected") or "").lower() == "true")
            if enabled:
                break

            checkbox.evaluate(
                """el => {
                    el.click();
                    el.dispatchEvent(new Event('input', { bubbles: true }));
                    el.dispatchEvent(new Event('change', { bubbles: true }));
                }"""
            )
            page.wait_for_timeout(200)

        if enabled:
            log(f"Enabled {toggle_name}")
            return True
        else:
            log(f"Failed to enable {toggle_name}")
            return False

    def open_modifiers_panel(log) -> None:
        for label in ["Modifiers", "Modifier"]:
            candidates = [
                page.get_by_role("button", name=re.compile(rf"^{re.escape(label)}$", re.IGNORECASE)).first,
                page.locator("button").filter(has_text=re.compile(rf"^{re.escape(label)}$", re.IGNORECASE)).first,
                page.locator("div").filter(has_text=re.compile(rf"^{re.escape(label)}$", re.IGNORECASE)).first,
                page.get_by_text(re.compile(rf"^{re.escape(label)}$", re.IGNORECASE)).first,
            ]
            for candidate in candidates:
                try:
                    if candidate.count() > 0:
                        candidate.click(timeout=1500, force=True)
                        log(f"Clicked {label}")
                        page.wait_for_timeout(500)
                        return
                except Exception:
                    continue

    enable_toggle("Show alt lines", required=False)
    if not enable_toggle(creature_toggle, required=False):
        open_modifiers_panel(log)
        enable_toggle(creature_toggle, required=False)

    show_results_btn = page.get_by_role("button", name=re.compile(r"^Show [\d,]+ results$", re.IGNORECASE)).first
    show_results_btn.wait_for(state="visible", timeout=15000)
    results_value = 0
    results_text = ""
    for _ in range(80):
        try:
            page.wait_for_load_state("networkidle", timeout=2000)
        except PlaywrightTimeoutError:
            pass

        results_text = (show_results_btn.inner_text() or "").strip()
        match = re.search(r"Show\s+([\d,]+)\s+results", results_text, re.IGNORECASE)
        if match:
            results_value = int(match.group(1).replace(",", ""))
            if results_value > 0:
                break

        page.wait_for_timeout(750)

    if results_value > 0:
        log(f"Results count loaded: {results_value:,}")
        show_results_btn.click(timeout=5000, force=True)
        log("Clicked Show results")
        try:
            page.wait_for_load_state("domcontentloaded", timeout=10000)
            page.wait_for_load_state("networkidle", timeout=10000)
        except PlaywrightTimeoutError:
            log("Load wait after Show results click timed out; continuing")

        projections_table = page.locator("#tour-projections-table").first
        projections_table.wait_for(state="visible", timeout=20000)
        log("Projections table loaded")

        page.get_by_role("button", name="SZN").click(timeout=5000)
        log("Sorted by SZN")
        try:
            page.wait_for_load_state("networkidle", timeout=7000)
        except PlaywrightTimeoutError:
            page.wait_for_timeout(800)

        target_props = sport_cfg["target_props"]

        all_rows: list[dict] = []
        for idx, prop_name in enumerate(target_props):
            if idx > 0:
                reset_table_to_top_and_sort(page, log, sport_cfg)

            if not select_props_filter(page, log, prop_name, target_props):
                continue
            try:
                page.wait_for_load_state("networkidle", timeout=6000)
            except PlaywrightTimeoutError:
                page.wait_for_timeout(800)

            control_projection_scroller(page, sport_cfg, "top")
            page.wait_for_timeout(300)
            expected_rows = get_visible_results_count(page, log)

            prop_rows = extract_projections_table(page, log, prop_name, expected_rows, sport_cfg)
            all_rows.extend(prop_rows)

            per_prop_path = BASE_DIR / (
                f"projections_extract_{output_suffix}_{prop_name.replace(' + ', '_').replace(' ', '_')}.json"
            )
            per_prop_path.write_text(json.dumps(prop_rows, indent=2))
            log(f"Saved per-prop JSON to {per_prop_path}")

        combined_json = BASE_DIR / f"projections_extract_all_props_{output_suffix}.json"
        combined_csv = BASE_DIR / f"projections_extract_all_props_{output_suffix}.csv"
        combined_json.write_text(json.dumps(all_rows, indent=2))

        fieldnames = [
            "selected_prop",
            "player_name",
            "player_id",
            "player_img_url",
            "prop",
            "prop_line",
            "L5",
            "L10",
            "L20",
            "H2H",
            "SZN",
            "AVG_3",
        ]
        with combined_csv.open("w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(all_rows)

        log(f"Saved combined JSON to {combined_json}")
        log(f"Saved combined CSV to {combined_csv}")
        log(f"Total extracted rows across all props: {len(all_rows)}")
        db_insert_props_rows(log, run_id, output_suffix, all_rows, sport_cfg)
    else:
        log(f"Results count still zero/unparsed: {results_text}")
        log("Skipping Show results click because count is not ready")


def main() -> None:
    load_dotenv()
    email = os.getenv("PROPSEDGE_EMAIL")
    password = os.getenv("PROPSEDGE_PSWD")
    if not email or not password:
        raise RuntimeError("Missing PROPSEDGE_EMAIL or PROPSEDGE_PSWD in .env")
    sport_cfg = active_sport_config()

    log_path = BASE_DIR / "run.log"
    def log(message: str) -> None:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = f"[{timestamp}] {message}"
        print(line)
        log_path.write_text(log_path.read_text() + line + "\n" if log_path.exists() else line + "\n")

    log(f"Active sport: {sport_cfg['label']} ({len(sport_cfg['target_props'])} target props)")

    with sync_playwright() as p:
        storage_path = Path("storage_state.json")
        login_notified = {"sent": False}
        failure_email_sent = {"sent": False}
        def run_pass(mode_name: str, creature_toggle: str, output_suffix: str) -> None:
            mode_prefix = f"[{mode_name}] "

            def mode_log(message: str) -> None:
                log(mode_prefix + message)

            run_id = db_insert_run_start(mode_log, output_suffix, sport_cfg)
            browser = None
            context = None
            try:
                headless = os.getenv("PROPSEDGE_HEADLESS", "1").strip().lower() not in {"0", "false", "f", "off", "no", "disable", "disabled"}
                browser = p.chromium.launch(headless=headless)
                if storage_path.exists():
                    context = browser.new_context(storage_state=str(storage_path))
                    mode_log("Loaded existing session from storage_state.json")
                else:
                    context = browser.new_context()
                    mode_log("No saved session found; starting fresh session")

                page = context.new_page()
                page.goto("https://www.propsedge.io/", wait_until="domcontentloaded")

                account_menu = page.locator("[data-test=\"account-dropdown-trigger\"]")
                sign_in_link = page.get_by_role("link", name="Sign In")
                logged_in = not (account_menu.count() == 0 and sign_in_link.count() > 0)

                if not logged_in:
                    mode_log("Not logged in; starting login flow")
                    try:
                        perform_login(page, email, password)
                    except Exception as exc:
                        if not login_notified["sent"]:
                            maybe_send_email(
                                log,
                                "PropsEdge: login failed",
                                f"Login failed with error: {exc}",
                            )
                            login_notified["sent"] = True
                        raise
                    page.wait_for_timeout(3000)
                    dismiss_home_popups(page, mode_log)
                    context.storage_state(path=str(storage_path))
                    mode_log("Login complete; session saved to storage_state.json")
                    if not login_notified["sent"]:
                        maybe_send_email(
                            log,
                            "PropsEdge: login success",
                            "Login completed successfully; scraping will proceed.",
                        )
                        login_notified["sent"] = True
                else:
                    mode_log("Already logged in; skipping login flow")
                    if not login_notified["sent"]:
                        maybe_send_email(
                            log,
                            "PropsEdge: login already active",
                            "Session is already logged in; scraping will proceed.",
                        )
                        login_notified["sent"] = True

                page.goto("https://www.propsedge.io/home/projections", wait_until="domcontentloaded")
                dismiss_home_popups(page, mode_log)
                click_sport_and_scrape(page, mode_log, creature_toggle, output_suffix, sport_cfg, run_id=run_id)
                try:
                    sync_outputs_to_google(mode_log, output_suffix)
                except Exception as exc:
                    mode_log(f"Google sync failed: {exc}")
                mode_log("Navigated to /home/projections")
                page.wait_for_timeout(5000)
                db_update_run_finish(mode_log, run_id, "success")
            except Exception as exc:
                db_update_run_finish(mode_log, run_id, "failed", str(exc))
                raise
            finally:
                if context:
                    context.close()
                if browser:
                    browser.close()

        try:
            run_pass("Demons", "Demons", "demons")
            run_pass("Goblins", "Goblins", "goblins")
            maybe_send_email(
                log,
                "PropsEdge: scraper complete",
                "Scrape finished and outputs were generated. PNG export starting.",
            )
            try:
                log("[Render] Starting PNG generation")
                clear_generated_pngs(log)
                render_png.main()
                log("[Render] PNG generation completed")
                png_files = collect_generated_pngs()
                log(f"[Render] Collected {len(png_files)} generated PNG files")
                send_pngs_to_telegram(log, png_files)
                send_pngs_to_discord(log, png_files)
                goblins_count = len([p for p in png_files if p.name.lower().startswith("goblins_")])
                demons_count = len([p for p in png_files if p.name.lower().startswith("demons_")])
                maybe_send_email(
                    log,
                    "PropsEdge: PNG export complete",
                    (
                        "PNG export completed.\n\n"
                        f"PNGs created: {len(png_files)}\n"
                        f"- Goblins: {goblins_count}\n"
                        f"- Demons: {demons_count}"
                    ),
                )
            except Exception as exc:
                log(f"[Render] PNG generation failed: {exc}")
                maybe_send_email(
                    log,
                    "PropsEdge: scraper failed",
                    f"Scraper failed with error: {exc}",
                )
                failure_email_sent["sent"] = True
                raise
        except Exception as exc:
            log(f"[Run] Failed: {exc}")
            if not failure_email_sent["sent"]:
                maybe_send_email(
                    log,
                    "PropsEdge: scraper failed",
                    f"Scraper failed with error: {exc}",
                )
                failure_email_sent["sent"] = True
            raise


if __name__ == "__main__":
    main()
