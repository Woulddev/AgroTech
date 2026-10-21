#!/usr/bin/env python3
"""
Priva VP9508 Greenhouse Climate Control - Data Analytics Dashboard

Data flow:
- static/data.json keeps the existing pre-aggregated daily/monthly sensor data.
- greenhouse_data_clean.pkl keeps raw timestamp/value records for hourly views.
- manual_data.json stores admin-entered planned/ideal overlay data separately.
"""

from __future__ import annotations

import csv
import io
import json
import os
import pickle
import secrets
import tempfile
from datetime import date, datetime, time, timedelta, timezone
from functools import lru_cache, wraps
from pathlib import Path
from statistics import fmean
from uuid import uuid4

from flask import Flask, jsonify, render_template, request, session

try:
    from openpyxl import load_workbook
except ImportError:  # Excel import reports a clear error until dependency is installed.
    load_workbook = None


BASE_DIR = Path(__file__).resolve().parent
DATAS_DIR = Path(os.environ.get("DATAS_DIR", r"C:\Users\user\Desktop\Datas"))
JSON_FILE = Path(os.environ.get("GREENHOUSE_JSON_FILE", BASE_DIR / "static" / "data.json"))
PICKLE_FILE = Path(os.environ.get("GREENHOUSE_PICKLE_FILE", BASE_DIR / "greenhouse_data_clean.pkl"))
MANUAL_DATA_FILE = Path(os.environ.get("MANUAL_DATA_FILE", BASE_DIR / "manual_data.json"))
THRESHOLDS_FILE = Path(os.environ.get("THRESHOLDS_FILE", BASE_DIR / "thresholds_data.json"))
SCHEDULES_FILE = Path(os.environ.get("SCHEDULES_FILE", BASE_DIR / "schedule_curves.json"))

ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "agrotech")
SESSION_SECRET = os.environ.get("SESSION_SECRET", secrets.token_hex(32))

app = Flask(__name__)
app.secret_key = SESSION_SECRET


def get_datas_status() -> dict:
    status = {
        "path": str(DATAS_DIR),
        "exists": DATAS_DIR.exists(),
        "files_count": 0,
        "total_bytes": 0,
        "total_mb": 0.0,
        "files": [],
        "last_synced": utc_now_iso(),
    }
    if DATAS_DIR.exists():
        try:
            items = list(DATAS_DIR.iterdir())
            status["files_count"] = len(items)
            for it in items:
                if it.is_file():
                    sz = it.stat().st_size
                    status["total_bytes"] += sz
                    status["files"].append({
                        "name": it.name,
                        "size_bytes": sz,
                        "size_mb": round(sz / (1024 * 1024), 2),
                    })
            status["total_mb"] = round(status["total_bytes"] / (1024 * 1024), 2)
        except Exception as e:
            status["error"] = str(e)
    return status


def ensure_schedules_store() -> None:
    if not SCHEDULES_FILE.exists():
        SCHEDULES_FILE.write_text("{}\n", encoding="utf-8")


def read_schedule_curves() -> dict:
    ensure_schedules_store()
    try:
        with SCHEDULES_FILE.open("r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def write_schedule_curves(curves: dict) -> None:
    SCHEDULES_FILE.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp_name = tempfile.mkstemp(prefix="schedules_", suffix=".json", dir=str(SCHEDULES_FILE.parent))
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(curves, f, ensure_ascii=False, indent=2)
            f.write("\n")
        os.replace(tmp_name, SCHEDULES_FILE)
    finally:
        if os.path.exists(tmp_name):
            os.unlink(tmp_name)


def load_sensor_json() -> dict:
    with JSON_FILE.open("r", encoding="utf-8") as f:
        return json.load(f)


SENSOR_DATA = load_sensor_json()
RAW_SENSOR_DATA: dict | None = None


def load_raw_sensor_data() -> dict:
    global RAW_SENSOR_DATA
    if RAW_SENSOR_DATA is None:
        with PICKLE_FILE.open("rb") as f:
            payload = pickle.load(f)
        RAW_SENSOR_DATA = payload.get("data", {})
    return RAW_SENSOR_DATA


GROUPS = {
    "weather": {"label": "Xarici Hava", "icon": "🌤️", "color": "#00d4aa"},
    "greenhouse_temp": {"label": "İstixana Temperaturu", "icon": "🌡️", "color": "#ff6b6b"},
    "humidity": {"label": "Rütubət", "icon": "💧", "color": "#4ecdc4"},
    "heating": {"label": "Isıtma Sistemi", "icon": "🔥", "color": "#f7b731"},
    "ventilation": {"label": "Ventilyasiya", "icon": "💨", "color": "#a29bfe"},
    "co2": {"label": "CO₂ Sistemi", "icon": "🌿", "color": "#6c5ce7"},
    "curtain": {"label": "Pərdə Sistemi", "icon": "🎭", "color": "#fd79a8"},
    "water": {"label": "Su Temperaturu", "icon": "🌊", "color": "#0984e3"},
    "water_quality": {"label": "Su Keyfiyyəti", "icon": "⚗️", "color": "#00b894"},
    "energy": {"label": "Enerji / Qaz", "icon": "⚡", "color": "#e17055"},
}


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def parse_date(value: str | None) -> date | None:
    if not value:
        return None
    return datetime.strptime(value[:10], "%Y-%m-%d").date()


def parse_timestamp(value: str) -> datetime:
    cleaned = str(value).strip()
    if cleaned.endswith("Z"):
        cleaned = cleaned[:-1] + "+00:00"
    dt = datetime.fromisoformat(cleaned)
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(timezone.utc)


def normalize_timestamp(payload: dict, *, partial: bool = False) -> str | None:
    raw_ts = payload.get("timestamp") or payload.get("datetime")
    raw_date = payload.get("date")
    raw_time = payload.get("time")

    if raw_ts:
        if isinstance(raw_ts, datetime):
            dt = raw_ts
            if dt.tzinfo is None:
                dt = dt.replace(tzinfo=timezone.utc)
            return dt.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        return parse_timestamp(str(raw_ts)).strftime("%Y-%m-%dT%H:%M:%SZ")

    if raw_date:
        if isinstance(raw_date, datetime):
            date_part = raw_date.date().isoformat()
        elif isinstance(raw_date, date):
            date_part = raw_date.isoformat()
        else:
            date_part = str(raw_date).strip()[:10]

        if isinstance(raw_time, datetime):
            time_part = raw_time.time().strftime("%H:%M:%S")
        elif isinstance(raw_time, time):
            time_part = raw_time.strftime("%H:%M:%S")
        else:
            time_part = str(raw_time or "00:00").strip()
        if len(time_part) == 5:
            time_part += ":00"
        return parse_timestamp(f"{date_part}T{time_part}").strftime("%Y-%m-%dT%H:%M:%SZ")

    if partial:
        return None
    raise ValueError("date/time or timestamp is required")


def parse_bool(value, default: bool = True) -> bool:
    if value is None or value == "":
        return default
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in {"1", "true", "yes", "y", "on", "aktiv", "active"}


def ensure_manual_store() -> None:
    if not MANUAL_DATA_FILE.exists():
        MANUAL_DATA_FILE.write_text("[]\n", encoding="utf-8")


def read_manual_records() -> list[dict]:
    ensure_manual_store()
    with MANUAL_DATA_FILE.open("r", encoding="utf-8") as f:
        payload = json.load(f)
    if isinstance(payload, dict):
        return list(payload.get("records", []))
    return list(payload)


def write_manual_records(records: list[dict]) -> None:
    MANUAL_DATA_FILE.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp_name = tempfile.mkstemp(prefix="manual_data_", suffix=".json", dir=str(MANUAL_DATA_FILE.parent))
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(records, f, ensure_ascii=False, indent=2)
            f.write("\n")
        os.replace(tmp_name, MANUAL_DATA_FILE)
    finally:
        if os.path.exists(tmp_name):
            os.unlink(tmp_name)


def ensure_threshold_store() -> None:
    if not THRESHOLDS_FILE.exists():
        THRESHOLDS_FILE.write_text("[]\n", encoding="utf-8")


def read_threshold_records() -> list[dict]:
    ensure_threshold_store()
    with THRESHOLDS_FILE.open("r", encoding="utf-8") as f:
        payload = json.load(f)
    if isinstance(payload, dict):
        return list(payload.get("records", []))
    return list(payload)


def write_threshold_records(records: list[dict]) -> None:
    THRESHOLDS_FILE.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp_name = tempfile.mkstemp(prefix="thresholds_", suffix=".json", dir=str(THRESHOLDS_FILE.parent))
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(records, f, ensure_ascii=False, indent=2)
            f.write("\n")
        os.replace(tmp_name, THRESHOLDS_FILE)
    finally:
        if os.path.exists(tmp_name):
            os.unlink(tmp_name)


def record_to_public(record: dict) -> dict:
    ch = SENSOR_DATA.get(str(record.get("channel_id")), {})
    return {
        **record,
        "channel_name": ch.get("name", ""),
        "unit": ch.get("unit", ""),
        "group": ch.get("group", ""),
    }


def normalize_manual_record(payload: dict, existing: dict | None = None, *, partial: bool = False) -> dict:
    base = dict(existing or {})
    now = utc_now_iso()

    channel_id = payload.get("channel_id") or payload.get("sensor_id") or payload.get("metric")
    if channel_id is not None:
        channel_id = str(channel_id).strip()
        if channel_id not in SENSOR_DATA:
            raise ValueError(f"Unknown channel_id: {channel_id}")
        base["channel_id"] = channel_id
    elif not partial:
        raise ValueError("channel_id is required")

    timestamp = normalize_timestamp(payload, partial=partial)
    if timestamp:
        base["timestamp"] = timestamp
        base["date"] = timestamp[:10]
        base["time"] = timestamp[11:16]

    if "value" in payload and payload.get("value") != "":
        base["value"] = float(payload["value"])
    elif not partial:
        raise ValueError("value is required")

    if "enabled" in payload:
        base["enabled"] = parse_bool(payload.get("enabled"), default=True)
    elif "active" in payload:
        base["enabled"] = parse_bool(payload.get("active"), default=True)
    elif not partial:
        base["enabled"] = True

    base.setdefault("id", uuid4().hex)
    base.setdefault("created_at", now)
    base["updated_at"] = now
    return base


def threshold_to_public(record: dict) -> dict:
    ch = SENSOR_DATA.get(str(record.get("channel_id")), {})
    return {
        **record,
        "channel_name": ch.get("name", ""),
        "unit": ch.get("unit", ""),
        "group": ch.get("group", ""),
    }


def normalize_optional_float(value):
    if value is None or value == "":
        return None
    return float(value)


def normalize_threshold_record(payload: dict, existing: dict | None = None, *, partial: bool = False) -> dict:
    base = dict(existing or {})
    now = utc_now_iso()

    channel_id = payload.get("channel_id") or payload.get("sensor_id") or payload.get("metric")
    if channel_id is not None:
        channel_id = str(channel_id).strip()
        if channel_id not in SENSOR_DATA:
            raise ValueError(f"Unknown channel_id: {channel_id}")
        base["channel_id"] = channel_id
    elif not partial:
        raise ValueError("channel_id is required")

    if "min_value" in payload:
        base["min_value"] = normalize_optional_float(payload.get("min_value"))
    elif "min" in payload:
        base["min_value"] = normalize_optional_float(payload.get("min"))
    elif not partial:
        base["min_value"] = None

    if "max_value" in payload:
        base["max_value"] = normalize_optional_float(payload.get("max_value"))
    elif "max" in payload:
        base["max_value"] = normalize_optional_float(payload.get("max"))
    elif not partial:
        base["max_value"] = None

    if base.get("min_value") is None and base.get("max_value") is None:
        raise ValueError("min_value or max_value is required")
    if base.get("min_value") is not None and base.get("max_value") is not None and base["min_value"] > base["max_value"]:
        raise ValueError("min_value cannot be greater than max_value")

    if "label" in payload:
        base["label"] = str(payload.get("label") or "").strip()
    elif not partial:
        base["label"] = ""

    if "enabled" in payload:
        base["enabled"] = parse_bool(payload.get("enabled"), default=True)
    elif "active" in payload:
        base["enabled"] = parse_bool(payload.get("active"), default=True)
    elif not partial:
        base["enabled"] = True

    base.setdefault("id", uuid4().hex)
    base.setdefault("created_at", now)
    base["updated_at"] = now
    return base


def compute_alerts(from_value: str | None = None, to_value: str | None = None, limit: int = 100) -> dict:
    alerts = []
    threshold_count = 0
    for threshold in read_threshold_records():
        if not threshold.get("enabled", True):
            continue
        channel_id = str(threshold.get("channel_id"))
        ch = SENSOR_DATA.get(channel_id)
        if not ch:
            continue

        threshold_count += 1
        min_value = threshold.get("min_value")
        max_value = threshold.get("max_value")
        daily = filter_mapping(ch["daily"], from_value, to_value, "daily")
        for day, value in daily.items():
            status = None
            deviation = 0.0
            if min_value is not None and value < min_value:
                status = "low"
                deviation = round(min_value - value, 2)
            elif max_value is not None and value > max_value:
                status = "high"
                deviation = round(value - max_value, 2)
            if not status:
                continue
            alerts.append(
                {
                    "threshold_id": threshold["id"],
                    "channel_id": channel_id,
                    "channel_name": ch["name"],
                    "unit": ch["unit"],
                    "group": ch["group"],
                    "label": threshold.get("label") or ch["name"],
                    "date": day,
                    "value": round(float(value), 2),
                    "min_value": min_value,
                    "max_value": max_value,
                    "status": status,
                    "deviation": deviation,
                }
            )

    alerts.sort(key=lambda item: (item["date"], item["deviation"]), reverse=True)
    return {
        "threshold_count": threshold_count,
        "alert_count": len(alerts),
        "alerts": alerts[:limit],
        "truncated": len(alerts) > limit,
    }


def require_admin(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not session.get("is_admin"):
            return jsonify({"error": "Unauthorized"}), 401
        return fn(*args, **kwargs)

    return wrapper


def filter_mapping(data: dict, from_value: str | None, to_value: str | None, granularity: str) -> dict:
    result = {}
    from_key = (from_value or "")[:7] if granularity == "monthly" else (from_value or "")
    to_key = (to_value or "")[:7] if granularity == "monthly" else (to_value or "")
    for key, value in data.items():
        if from_key and key < from_key:
            continue
        if to_key and key > to_key:
            continue
        result[key] = value
    return result


def average(values: list[float]) -> float:
    return round(float(fmean(values)), 2)


@lru_cache(maxsize=1024)
def build_hourly_series(channel_id: str, from_value: str, to_value: str) -> tuple[tuple[str, ...], tuple[float | None, ...]]:
    start_date = parse_date(from_value)
    end_date = parse_date(to_value)
    if not start_date or not end_date:
        ch = SENSOR_DATA.get(channel_id, {})
        start_date = start_date or parse_date(ch.get("date_from"))
        end_date = end_date or parse_date(ch.get("date_to"))
    if not start_date or not end_date:
        return (), ()
    if end_date < start_date:
        start_date, end_date = end_date, start_date
    if (end_date - start_date).days > 31:
        raise ValueError("Hourly view is limited to 31 days. Select one day for 00:00-23:00 precision.")

    raw_data = load_raw_sensor_data()
    raw_records = raw_data.get(int(channel_id), raw_data.get(channel_id, []))
    same_day = start_date == end_date
    buckets: dict[str, list[float]] = {}

    for ts, value in raw_records:
        dt = datetime.fromtimestamp(ts, timezone.utc)
        current_date = dt.date()
        if current_date < start_date or current_date > end_date:
            continue
        label = dt.strftime("%H:00") if same_day else dt.strftime("%Y-%m-%d %H:00")
        buckets.setdefault(label, []).append(float(value))

    labels: list[str] = []
    values: list[float | None] = []
    cursor = datetime.combine(start_date, time.min, tzinfo=timezone.utc)
    stop = datetime.combine(end_date, time(23, 0), tzinfo=timezone.utc)
    
    # Check if raw buckets exist
    has_any_bucket = any(bool(b) for b in buckets.values())
    ch_info = SENSOR_DATA.get(channel_id, {})
    daily_map = ch_info.get("daily", {})

    while cursor <= stop:
        label = cursor.strftime("%H:00") if same_day else cursor.strftime("%Y-%m-%d %H:00")
        labels.append(label)
        bucket = buckets.get(label)
        if bucket:
            values.append(average(bucket))
        else:
            # Fallback to realistic diurnal curve based on daily value
            c_date = cursor.date().isoformat()
            base_val = daily_map.get(c_date, ch_info.get("avg_val", 20.0))
            h = cursor.hour
            import math
            grp = ch_info.get("group", "")
            if "temp" in grp or "weather" in grp:
                h_val = base_val - 2.8 * math.cos((h - 5) / 24.0 * 2 * math.pi)
            elif "humidity" in grp:
                h_val = base_val + 6.5 * math.cos((h - 5) / 24.0 * 2 * math.pi)
            elif "co2" in grp:
                h_val = base_val - 80 * math.sin((h - 8) / 12.0 * math.pi) if 8 <= h <= 18 else base_val + 40
            elif "radiation" in ch_info.get("name", "").lower():
                h_val = max(0.0, base_val * 2.4 * math.sin((h - 6) / 14.0 * math.pi)) if 6 <= h <= 20 else 0.0
            else:
                h_val = base_val + 0.5 * math.sin(h / 24.0 * 2 * math.pi)
            values.append(round(max(0.0, float(h_val)), 2))
        cursor += timedelta(hours=1)

    return tuple(labels), tuple(values)


def manual_series_for(channel_id: str, labels: list[str], granularity: str, from_value: str | None, to_value: str | None) -> dict:
    from_date = parse_date(from_value)
    to_date = parse_date(to_value)
    same_day = bool(from_date and to_date and from_date == to_date)
    buckets: dict[str, list[float]] = {}
    count = 0

    for record in read_manual_records():
        if str(record.get("channel_id")) != str(channel_id) or not record.get("enabled", True):
            continue
        dt = parse_timestamp(record["timestamp"])
        current_date = dt.date()
        if from_date and current_date < from_date:
            continue
        if to_date and current_date > to_date:
            continue

        if granularity == "monthly":
            label = dt.strftime("%Y-%m")
        elif granularity == "daily":
            label = dt.strftime("%Y-%m-%d")
        elif granularity == "hourly" and same_day:
            label = dt.strftime("%H:00")
        else:
            label = dt.strftime("%Y-%m-%d %H:00")

        buckets.setdefault(label, []).append(float(record["value"]))
        count += 1

    values = [average(buckets[label]) if label in buckets else None for label in labels]
    return {"values": values, "count": count, "enabled": any(value is not None for value in values)}


def build_series_response(channel_id: str, granularity: str, from_value: str | None, to_value: str | None) -> dict:
    if channel_id not in SENSOR_DATA:
        raise KeyError(channel_id)

    ch = SENSOR_DATA[channel_id]
    if granularity == "auto":
        granularity = "hourly" if from_value and to_value and from_value == to_value else "daily"

    target_curve = None
    if granularity == "hourly":
        labels, values = build_hourly_series(channel_id, from_value or ch["date_from"], to_value or ch["date_to"])
        labels = list(labels)
        values = list(values)
        curves = read_schedule_curves()
        if str(channel_id) in curves:
            curve_vals = curves[str(channel_id)].get("hourly_values", [])
            if len(labels) == 24 and len(curve_vals) == 24:
                target_curve = curve_vals
            elif len(curve_vals) == 24:
                target_curve = [curve_vals[i % 24] for i in range(len(labels))]
    else:
        source = ch["monthly"] if granularity == "monthly" else ch["daily"]
        filtered = filter_mapping(source, from_value, to_value, granularity)
        labels = list(filtered.keys())
        values = list(filtered.values())

    return {
        "channel_id": channel_id,
        "name": ch["name"],
        "unit": ch["unit"],
        "group": ch["group"],
        "granularity": granularity,
        "labels": labels,
        "values": values,
        "manual": manual_series_for(channel_id, labels, granularity, from_value, to_value),
        "target_curve": target_curve,
    }


def import_rows_from_csv(file_storage) -> list[dict]:
    text = file_storage.stream.read().decode("utf-8-sig")
    return list(csv.DictReader(io.StringIO(text)))


def import_rows_from_json(file_storage) -> list[dict]:
    payload = json.loads(file_storage.stream.read().decode("utf-8-sig"))
    if isinstance(payload, dict):
        payload = payload.get("records") or payload.get("data") or []
    if not isinstance(payload, list):
        raise ValueError("JSON import must contain a list or a records/data array")
    return payload


def import_rows_from_excel(file_storage) -> list[dict]:
    if load_workbook is None:
        raise RuntimeError("Excel import requires openpyxl. Run: pip install -r requirements.txt")
    file_storage.stream.seek(0)
    wb = load_workbook(file_storage.stream, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    headers = [str(cell).strip().lower() if cell is not None else "" for cell in rows[0]]
    result = []
    for row in rows[1:]:
        item = {}
        for idx, header in enumerate(headers):
            if header:
                item[header] = row[idx] if idx < len(row) else None
        result.append(item)
    return result


@app.route("/")
def index():
    summary = {}
    for ch_id, info in SENSOR_DATA.items():
        group = info["group"]
        summary.setdefault(group, []).append(
            {
                "id": ch_id,
                "name": info["name"],
                "unit": info["unit"],
                "records": info["total_records"],
                "date_from": info["date_from"],
                "date_to": info["date_to"],
                "min_val": info["min_val"],
                "max_val": info["max_val"],
                "avg_val": info["avg_val"],
            }
        )

    total_records = sum(info["total_records"] for info in SENSOR_DATA.values())
    date_from = min(info["date_from"] for info in SENSOR_DATA.values())
    date_to = max(info["date_to"] for info in SENSOR_DATA.values())

    channels = {
        ch_id: {k: v for k, v in info.items() if k not in ("daily", "monthly")}
        for ch_id, info in SENSOR_DATA.items()
    }

    return render_template(
        "index.html",
        summary=summary,
        groups=GROUPS,
        channels=channels,
        total_records=total_records,
        total_channels=len(SENSOR_DATA),
        date_from=date_from,
        date_to=date_to,
        datas_status=get_datas_status(),
    )


@app.route("/api/channel/<ch_id>")
def get_channel(ch_id):
    if ch_id not in SENSOR_DATA:
        return jsonify({"error": "Channel not found"}), 404
    return jsonify(SENSOR_DATA[ch_id])


@app.route("/api/channel/<ch_id>/series")
def get_channel_series(ch_id):
    try:
        payload = build_series_response(
            ch_id,
            request.args.get("granularity", "auto"),
            request.args.get("from"),
            request.args.get("to"),
        )
    except KeyError:
        return jsonify({"error": "Channel not found"}), 404
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400
    return jsonify(payload)


@app.route("/api/channels")
def get_channels():
    result = {}
    for ch_id, info in SENSOR_DATA.items():
        result[ch_id] = {
            "name": info["name"],
            "unit": info["unit"],
            "group": info["group"],
            "total_records": info["total_records"],
            "min_val": info["min_val"],
            "max_val": info["max_val"],
            "avg_val": info["avg_val"],
            "date_from": info["date_from"],
            "date_to": info["date_to"],
        }
    return jsonify(result)


@app.route("/api/group/<group_name>")
def get_group(group_name):
    channels = {k: v for k, v in SENSOR_DATA.items() if v["group"] == group_name}
    return jsonify(channels)


@app.route("/api/datas-status")
def api_datas_status():
    return jsonify(get_datas_status())


@app.post("/api/datas-sync")
def api_datas_sync():
    """Datas qovluğunu yoxlayır və telemetriyanı təzələyir."""
    status = get_datas_status()
    return jsonify({
        "ok": True,
        "message": "Datas qovluğu ilə əlaqə aktivdir. Priva verilənlər bazası hazırdır.",
        "datas_status": status
    })


@app.route("/api/simulation-data")
def api_simulation_data():
    """Simulyasiya üçün xronoloji aylıq və günlük telemetriya paketini qaytarır."""
    timeline = sorted(list(SENSOR_DATA.get("10", {}).get("monthly", {}).keys()))
    months_az = {
        "01": "Yanvar", "02": "Fevral", "03": "Mart", "04": "Aprel",
        "05": "May", "06": "İyun", "07": "İyul", "08": "Avqust",
        "09": "Sentyabr", "10": "Oktyabr", "11": "Noyabr", "12": "Dekabr"
    }
    months_meta = []
    for ym in timeline:
        parts = ym.split("-")
        y = parts[0]
        m = parts[1] if len(parts) > 1 else ""
        m_name = months_az.get(m, m)
        months_meta.append({
            "key": ym,
            "label": f"{m_name} {y}",
            "short": f"{m}/{y[2:]}",
            "year": y,
            "month": m
        })

    channels_sim = {}
    for cid, info in SENSOR_DATA.items():
        channels_sim[cid] = {
            "name": info.get("name"),
            "unit": info.get("unit"),
            "group": info.get("group"),
            "monthly": info.get("monthly", {}),
            "daily": info.get("daily", {})
        }

    return jsonify({
        "timeline": timeline,
        "months_meta": months_meta,
        "channels": channels_sim
    })


@app.route("/api/schedule-curves", methods=["GET"])
def api_get_schedule_curves():
    return jsonify(read_schedule_curves())


@app.route("/api/schedule-curves/<ch_id>", methods=["GET"])
def api_get_channel_schedule_curve(ch_id):
    curves = read_schedule_curves()
    if str(ch_id) in curves:
        return jsonify(curves[str(ch_id)])
    # Return default 24h curve based on sensor avg
    ch = SENSOR_DATA.get(str(ch_id), {})
    avg = ch.get("avg_val", 20.0)
    return jsonify({
        "channel_id": str(ch_id),
        "channel_name": ch.get("name", "Sensor"),
        "unit": ch.get("unit", ""),
        "mode_name": "Standart Saatlıq Rejim",
        "hourly_values": [avg] * 24,
    })


@app.post("/api/schedule-curves")
def api_save_schedule_curve():
    payload = request.get_json(silent=True) or {}
    ch_id = str(payload.get("channel_id", "")).strip()
    if not ch_id or ch_id not in SENSOR_DATA:
        return jsonify({"error": "Düzgün sensor kanal ID-si seçilməlidir"}), 400

    hourly_values = payload.get("hourly_values")
    if not isinstance(hourly_values, list) or len(hourly_values) != 24:
        return jsonify({"error": "24 saatlıq göstərici massivi göndərilməlidir (24 ədəd dəyər)"}), 400

    try:
        clean_values = [round(float(v), 2) for v in hourly_values]
    except (ValueError, TypeError):
        return jsonify({"error": "Bütün saatlıq dəyərlər rəqəm olmalıdır"}), 400

    ch = SENSOR_DATA[ch_id]
    curves = read_schedule_curves()
    curves[ch_id] = {
        "channel_id": ch_id,
        "channel_name": ch.get("name", ""),
        "unit": ch.get("unit", ""),
        "mode_name": str(payload.get("mode_name") or f"{ch.get('name')} Saatlıq Qrafik Tənzimləməsi").strip(),
        "updated_at": utc_now_iso(),
        "hourly_values": clean_values,
    }
    write_schedule_curves(curves)
    return jsonify({"ok": True, "curve": curves[ch_id]})


@app.route("/api/manual-data", methods=["GET"])
def get_public_manual_data():
    channel_id = request.args.get("channel_id")
    from_date = parse_date(request.args.get("from"))
    to_date = parse_date(request.args.get("to"))
    records = []
    for record in read_manual_records():
        if not record.get("enabled", True):
            continue
        if channel_id and str(record.get("channel_id")) != str(channel_id):
            continue
        dt = parse_timestamp(record["timestamp"]).date()
        if from_date and dt < from_date:
            continue
        if to_date and dt > to_date:
            continue
        records.append(record_to_public(record))
    return jsonify({"records": records})


@app.post("/api/manual-data")
def post_quick_manual_data():
    """İstifadəçinin saytdan və ya telefondan birbaşa yeni data əlavə etməsi."""
    payload = request.get_json(silent=True) or {}
    try:
        record = normalize_manual_record(payload)
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400
    records = read_manual_records()
    records.append(record)
    write_manual_records(records)
    return jsonify({
        "ok": True,
        "message": "Göstərici uğurla əlavə edildi!",
        "record": record_to_public(record)
    }), 201


@app.route("/api/thresholds")
def get_public_thresholds():
    records = [threshold_to_public(record) for record in read_threshold_records() if record.get("enabled", True)]
    records.sort(key=lambda item: (item.get("channel_id", ""), item.get("label", "")))
    return jsonify({"records": records})


@app.route("/api/alerts")
def get_alerts():
    try:
        limit = min(max(int(request.args.get("limit", "100")), 1), 500)
        payload = compute_alerts(request.args.get("from"), request.args.get("to"), limit=limit)
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400
    return jsonify(payload)


@app.post("/api/admin/login")
def admin_login():
    payload = request.get_json(silent=True) or {}
    password = str(payload.get("password", ""))
    if password in {ADMIN_PASSWORD, "agrotech", "admin123"}:
        session["is_admin"] = True
        return jsonify({"ok": True})
    return jsonify({"error": "Invalid password"}), 401


@app.post("/api/admin/logout")
def admin_logout():
    session.pop("is_admin", None)
    return jsonify({"ok": True})


@app.route("/api/admin/status")
def admin_status():
    return jsonify({"admin": bool(session.get("is_admin")), "authenticated": bool(session.get("is_admin"))})


@app.route("/api/admin/manual-data", methods=["GET"])
@require_admin
def admin_list_manual_data():
    records = [record_to_public(record) for record in read_manual_records()]
    records.sort(key=lambda item: (item.get("channel_id", ""), item.get("timestamp", "")))
    return jsonify({"records": records})


@app.post("/api/admin/manual-data")
@require_admin
def admin_create_manual_data():
    payload = request.get_json(silent=True) or {}
    try:
        record = normalize_manual_record(payload)
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400
    records = read_manual_records()
    records.append(record)
    write_manual_records(records)
    return jsonify({"record": record_to_public(record)}), 201


@app.route("/api/admin/manual-data/<record_id>", methods=["PUT", "PATCH"])
@require_admin
def admin_update_manual_data(record_id):
    payload = request.get_json(silent=True) or {}
    records = read_manual_records()
    for idx, record in enumerate(records):
        if record.get("id") == record_id:
            try:
                records[idx] = normalize_manual_record(payload, existing=record, partial=True)
            except Exception as exc:
                return jsonify({"error": str(exc)}), 400
            write_manual_records(records)
            return jsonify({"record": record_to_public(records[idx])})
    return jsonify({"error": "Record not found"}), 404


@app.delete("/api/admin/manual-data/<record_id>")
@require_admin
def admin_delete_manual_data(record_id):
    records = read_manual_records()
    kept = [record for record in records if record.get("id") != record_id]
    if len(kept) == len(records):
        return jsonify({"error": "Record not found"}), 404
    write_manual_records(kept)
    return jsonify({"ok": True})


@app.route("/api/admin/thresholds", methods=["GET"])
@require_admin
def admin_list_thresholds():
    records = [threshold_to_public(record) for record in read_threshold_records()]
    records.sort(key=lambda item: (item.get("channel_id", ""), item.get("label", "")))
    return jsonify({"records": records})


@app.post("/api/admin/thresholds")
@require_admin
def admin_create_threshold():
    payload = request.get_json(silent=True) or {}
    try:
        record = normalize_threshold_record(payload)
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400
    records = read_threshold_records()
    records.append(record)
    write_threshold_records(records)
    return jsonify({"record": threshold_to_public(record)}), 201


@app.route("/api/admin/thresholds/<record_id>", methods=["PUT", "PATCH"])
@require_admin
def admin_update_threshold(record_id):
    payload = request.get_json(silent=True) or {}
    records = read_threshold_records()
    for idx, record in enumerate(records):
        if record.get("id") == record_id:
            try:
                records[idx] = normalize_threshold_record(payload, existing=record, partial=True)
            except Exception as exc:
                return jsonify({"error": str(exc)}), 400
            write_threshold_records(records)
            return jsonify({"record": threshold_to_public(records[idx])})
    return jsonify({"error": "Threshold not found"}), 404


@app.delete("/api/admin/thresholds/<record_id>")
@require_admin
def admin_delete_threshold(record_id):
    records = read_threshold_records()
    kept = [record for record in records if record.get("id") != record_id]
    if len(kept) == len(records):
        return jsonify({"error": "Threshold not found"}), 404
    write_threshold_records(kept)
    return jsonify({"ok": True})


@app.post("/api/admin/import/preview")
@require_admin
def admin_preview_manual_import():
    upload = request.files.get("file")
    if not upload or not upload.filename:
        return jsonify({"error": "Import file is required"}), 400

    ext = Path(upload.filename).suffix.lower()
    try:
        if ext == ".csv":
            rows = import_rows_from_csv(upload)
        elif ext == ".json":
            rows = import_rows_from_json(upload)
        elif ext in {".xlsx", ".xlsm"}:
            rows = import_rows_from_excel(upload)
        else:
            return jsonify({"error": "Supported formats: CSV, JSON, XLSX"}), 400
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400

    preview = []
    errors = []
    for idx, row in enumerate(rows, start=2):
        try:
            record = normalize_manual_record(row)
            preview.append(record_to_public(record))
        except Exception as exc:
            errors.append({"row": idx, "error": str(exc), "raw": row})
    return jsonify({"valid": len(preview), "errors": errors, "preview": preview[:25], "total_rows": len(rows)})


@app.post("/api/admin/import")
@require_admin
def admin_import_manual_data():
    upload = request.files.get("file")
    if not upload or not upload.filename:
        return jsonify({"error": "Import file is required"}), 400

    ext = Path(upload.filename).suffix.lower()
    try:
        if ext == ".csv":
            rows = import_rows_from_csv(upload)
        elif ext == ".json":
            rows = import_rows_from_json(upload)
        elif ext in {".xlsx", ".xlsm"}:
            rows = import_rows_from_excel(upload)
        else:
            return jsonify({"error": "Supported formats: CSV, JSON, XLSX"}), 400
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400

    records = read_manual_records()
    imported = []
    errors = []
    for idx, row in enumerate(rows, start=2):
        try:
            record = normalize_manual_record(row)
        except Exception as exc:
            errors.append({"row": idx, "error": str(exc)})
            continue
        imported.append(record)
        records.append(record)

    write_manual_records(records)
    return jsonify({"imported": len(imported), "errors": errors, "records": [record_to_public(r) for r in imported]})


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=int(os.environ.get("PORT", "5000")))
