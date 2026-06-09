#!/usr/bin/env python3
from __future__ import annotations

import json
import math
import os
import xml.etree.ElementTree as ET
from datetime import datetime, timezone
from pathlib import Path
from statistics import mean

import matplotlib

matplotlib.use("Agg")

import matplotlib.pyplot as plt
import pandas as pd


REPO_ROOT = Path(__file__).resolve().parents[1]
TRAININGS_ROOT = REPO_ROOT / "RunGPSData" / "Trainings"
OUT_ROOT = REPO_ROOT / "RunGPSData" / "Stats"
CHARTS_DIR = OUT_ROOT / "charts"
GEOJSON_DIR = OUT_ROOT / "geojson"

ELEVATION_GAIN_THRESHOLD_M = 3.0

import re

def parse_rungps_trainings_xml_metadata(path: Path) -> pd.DataFrame:
    """
    Liest RunGPSData/Trainings.xml.
    Die Datei ist faktisch keine normale XML-Struktur, sondern eine PowerShell-nahe Textdarstellung.
    Relevante Zeilen sehen z.B. so aus:

    DateTime 2 ~~Autofahren~~ 2868449 ~~(unbenannt)~~ 835.62 PT10H48M25S ...
    """

    if not path.exists():
        print(f"RunGPS metadata file not found: {path}")
        return pd.DataFrame()

    rows = []

    pattern = re.compile(
        r"^DateTime\s+"
        r"(?P<display_hint>\d+)\s+"
        r"~~(?P<sportart>.*?)~~\s+"
        r"(?P<id>\d+)\s+"
        r"~~(?P<titel>.*?)~~\s+"
        r"(?P<distanz>[-+]?\d+(?:\.\d+)?)\s+"
        r"(?P<dauer>PT\S+)\s+"
        r"(?P<kalorien>[-+]?\d+(?:\.\d+)?)\s+"
        r"(?P<herzfrequenz_d>[-+]?\d+(?:\.\d+)?)\s+"
        r"(?P<trittfrequenz_d>[-+]?\d+(?:\.\d+)?)\s+"
        r"(?P<geschwindigkeit_d>[-+]?\d+(?:\.\d+)?)\s+"
        r"(?P<geschwindigkeit_da>[-+]?\d+(?:\.\d+)?)\s+"
        r"(?P<hoehe_min>[-+]?\d+(?:\.\d+)?)\s+"
        r"(?P<hoehe_max>[-+]?\d+(?:\.\d+)?)\s+"
        r"(?P<abstieg>[-+]?\d+(?:\.\d+)?)\s+"
        r"(?P<aufstieg>[-+]?\d+(?:\.\d+)?)\s+"
        r"(?P<gewicht>[-+]?\d+(?:\.\d+)?)\s+"
        r"~~(?P<distanz_bereich>.*?)~~\s+"
        r"(?P<export_date>\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2})"
    )

    for line in path.read_text(encoding="utf-8", errors="replace").splitlines():
        line = line.strip()
        m = pattern.match(line)
        if not m:
            continue

        d = m.groupdict()

        rows.append(
            {
                "id": d["id"],
                "rungps_sportart": d["sportart"],
                "rungps_titel": d["titel"],
                "rungps_distanz_km": float(d["distanz"]),
                "rungps_dauer": d["dauer"],
                "rungps_kalorien": float(d["kalorien"]),
                "rungps_geschwindigkeit_kmh": float(d["geschwindigkeit_d"]),
                "rungps_geschwindigkeit_da_kmh": float(d["geschwindigkeit_da"]),
                "rungps_hoehe_min_m": float(d["hoehe_min"]),
                "rungps_hoehe_max_m": float(d["hoehe_max"]),
                "rungps_abstieg_m": float(d["abstieg"]),
                "rungps_aufstieg_m": float(d["aufstieg"]),
                "rungps_distanz_bereich": d["distanz_bereich"],
                "rungps_export_date": d["export_date"],
            }
        )

    result = pd.DataFrame(rows)

    if not result.empty:
        result["id"] = result["id"].astype(str)

    print(f"RunGPS metadata rows parsed: {len(result)}")
    return result
    
def local_name(tag: str) -> str:
    if "}" in tag:
        return tag.rsplit("}", 1)[1]
    return tag


def child(elem: ET.Element, name: str) -> ET.Element | None:
    for c in list(elem):
        if local_name(c.tag) == name:
            return c
    return None


def child_text(elem: ET.Element, name: str) -> str | None:
    c = child(elem, name)
    if c is None or c.text is None:
        return None
    return c.text.strip()


def first_desc_text(root: ET.Element, name: str) -> str | None:
    for e in root.iter():
        if local_name(e.tag) == name and e.text:
            return e.text.strip()
    return None


def safe_float(text: str | None) -> float | None:
    if text is None:
        return None
    text = str(text).strip()
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def safe_int(text: str | None) -> int | None:
    value = safe_float(text)
    if value is None:
        return None
    return int(round(value))


def parse_tcx_time(text: str | None) -> datetime | None:
    if not text:
        return None
    value = text.strip()
    try:
        return datetime.fromisoformat(value.replace("Z", "+00:00"))
    except ValueError:
        return None


def haversine_m(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    radius_m = 6_371_000.0
    phi1 = math.radians(lat1)
    phi2 = math.radians(lat2)
    d_phi = math.radians(lat2 - lat1)
    d_lambda = math.radians(lon2 - lon1)

    a = (
        math.sin(d_phi / 2) ** 2
        + math.cos(phi1) * math.cos(phi2) * math.sin(d_lambda / 2) ** 2
    )
    return 2 * radius_m * math.atan2(math.sqrt(a), math.sqrt(1 - a))


def format_duration(seconds: float | int | None) -> str:
    if seconds is None or pd.isna(seconds):
        return ""
    seconds = int(round(float(seconds)))
    hours = seconds // 3600
    minutes = (seconds % 3600) // 60
    secs = seconds % 60
    if hours:
        return f"{hours:d}:{minutes:02d}:{secs:02d}"
    return f"{minutes:d}:{secs:02d}"


def parse_tcx(path: Path) -> dict:
    tree = ET.parse(path)
    root = tree.getroot()

    activity = None
    for e in root.iter():
        if local_name(e.tag) == "Activity":
            activity = e
            break

    sport = activity.attrib.get("Sport", "") if activity is not None else ""

    activity_id = None
    if activity is not None:
        activity_id = child_text(activity, "Id")
    if not activity_id:
        activity_id = first_desc_text(root, "Id")

    laps = [e for e in root.iter() if local_name(e.tag) == "Lap"]

    lap_distance_m = 0.0
    lap_duration_s = 0.0
    calories = 0
    lap_start_times: list[datetime] = []

    for lap in laps:
        d = safe_float(child_text(lap, "DistanceMeters"))
        if d:
            lap_distance_m += d

        t = safe_float(child_text(lap, "TotalTimeSeconds"))
        if t:
            lap_duration_s += t

        c = safe_int(child_text(lap, "Calories"))
        if c:
            calories += c

        start_time = parse_tcx_time(lap.attrib.get("StartTime"))
        if start_time:
            lap_start_times.append(start_time)

    trackpoints = [e for e in root.iter() if local_name(e.tag) == "Trackpoint"]

    points: list[dict] = []
    distances_m: list[float] = []
    altitudes_m: list[float] = []
    hr_values: list[int] = []
    cadence_values: list[int] = []
    times: list[datetime] = []

    for tp in trackpoints:
        time_value = parse_tcx_time(child_text(tp, "Time"))
        if time_value:
            times.append(time_value)

        distance_m = safe_float(child_text(tp, "DistanceMeters"))
        if distance_m is not None:
            distances_m.append(distance_m)

        altitude_m = safe_float(child_text(tp, "AltitudeMeters"))
        if altitude_m is not None:
            altitudes_m.append(altitude_m)

        cadence = safe_int(child_text(tp, "Cadence"))
        if cadence is not None:
            cadence_values.append(cadence)

        hr_node = child(tp, "HeartRateBpm")
        if hr_node is not None:
            hr = safe_int(child_text(hr_node, "Value"))
            if hr is not None:
                hr_values.append(hr)

        pos = child(tp, "Position")
        lat = lon = None
        if pos is not None:
            lat = safe_float(child_text(pos, "LatitudeDegrees"))
            lon = safe_float(child_text(pos, "LongitudeDegrees"))

        if lat is not None and lon is not None:
            points.append(
                {
                    "lat": lat,
                    "lon": lon,
                    "altitude_m": altitude_m,
                    "time": time_value.isoformat() if time_value else None,
                }
            )

    fallback_distance_m = 0.0
    if len(points) >= 2:
        for a, b in zip(points, points[1:]):
            fallback_distance_m += haversine_m(a["lat"], a["lon"], b["lat"], b["lon"])

    distance_m = 0.0
    distance_source = "unknown"

    if lap_distance_m > 0:
        distance_m = lap_distance_m
        distance_source = "laps"
    elif distances_m:
        distance_m = max(distances_m)
        distance_source = "trackpoint-distance"
    elif fallback_distance_m > 0:
        distance_m = fallback_distance_m
        distance_source = "haversine"

    duration_s = 0.0
    duration_source = "unknown"

    if lap_duration_s > 0:
        duration_s = lap_duration_s
        duration_source = "laps"
    elif len(times) >= 2:
        duration_s = (max(times) - min(times)).total_seconds()
        duration_source = "trackpoint-time"

    elevation_gain_m = 0.0
    elevation_loss_m = 0.0
    if len(altitudes_m) >= 2:
        for a, b in zip(altitudes_m, altitudes_m[1:]):
            delta = b - a
            if delta > ELEVATION_GAIN_THRESHOLD_M:
                elevation_gain_m += delta
            elif delta < -ELEVATION_GAIN_THRESHOLD_M:
                elevation_loss_m += abs(delta)

    start_time = None
    if activity_id:
        start_time = parse_tcx_time(activity_id)
    if start_time is None and lap_start_times:
        start_time = min(lap_start_times)
    if start_time is None and times:
        start_time = min(times)

    start_lat = points[0]["lat"] if points else None
    start_lon = points[0]["lon"] if points else None

    avg_speed_kmh = None
    if duration_s > 0 and distance_m > 0:
        avg_speed_kmh = (distance_m / 1000.0) / (duration_s / 3600.0)

    rel_path = path.relative_to(REPO_ROOT).as_posix()
    rel_from_stats = os.path.relpath(path, OUT_ROOT).replace(os.sep, "/")

    return {
        "id": path.stem,
        "file": path.name,
        "path": rel_path,
        "link_from_stats": rel_from_stats,
        "activity_id": activity_id,
        "sport": sport,
        "start_time": start_time.isoformat() if start_time else None,
        "date": start_time.date().isoformat() if start_time else None,
        "distance_km": distance_m / 1000.0,
        "distance_source": distance_source,
        "duration_s": duration_s,
        "duration": format_duration(duration_s),
        "duration_source": duration_source,
        "avg_speed_kmh": avg_speed_kmh,
        "calories": calories,
        "elevation_gain_m": elevation_gain_m,
        "elevation_loss_m": elevation_loss_m,
        "min_altitude_m": min(altitudes_m) if altitudes_m else None,
        "max_altitude_m": max(altitudes_m) if altitudes_m else None,
        "avg_hr": mean(hr_values) if hr_values else None,
        "max_hr": max(hr_values) if hr_values else None,
        "avg_cadence": mean(cadence_values) if cadence_values else None,
        "max_cadence": max(cadence_values) if cadence_values else None,
        "trackpoints": len(trackpoints),
        "geo_points": len(points),
        "start_lat": start_lat,
        "start_lon": start_lon,
    }


def fmt_num(value, decimals=1, suffix=""):
    if value is None or pd.isna(value):
        return ""
    return f"{float(value):,.{decimals}f}{suffix}".replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_int(value):
    if value is None or pd.isna(value):
        return ""
    return f"{int(round(float(value))):,}".replace(",", ".")


def md_table(df: pd.DataFrame, columns: list[tuple[str, str]], max_rows: int = 20) -> str:
    if df.empty:
        return "_Keine Daten._\n"

    df = df.head(max_rows).copy()

    lines = []
    headers = [title for _, title in columns]
    lines.append("| " + " | ".join(headers) + " |")
    lines.append("| " + " | ".join(["---"] * len(headers)) + " |")

    for _, row in df.iterrows():
        values = []
        for col, _ in columns:
            value = row.get(col, "")
            if pd.isna(value):
                value = ""
            values.append(str(value))
        lines.append("| " + " | ".join(values) + " |")

    return "\n".join(lines) + "\n"


def best_row(df: pd.DataFrame, column: str) -> pd.Series | None:
    valid = df.dropna(subset=[column])
    if valid.empty:
        return None
    return valid.loc[valid[column].idxmax()]


def longest_streak(dates: list[pd.Timestamp]) -> tuple[pd.Timestamp | None, pd.Timestamp | None, int]:
    if not dates:
        return None, None, 0

    unique_days = sorted({d.normalize() for d in dates if pd.notna(d)})
    if not unique_days:
        return None, None, 0

    best_start = current_start = unique_days[0]
    best_end = current_end = unique_days[0]
    best_len = current_len = 1

    for day in unique_days[1:]:
        if (day - current_end).days == 1:
            current_end = day
            current_len += 1
        else:
            if current_len > best_len:
                best_start, best_end, best_len = current_start, current_end, current_len
            current_start = current_end = day
            current_len = 1

    if current_len > best_len:
        best_start, best_end, best_len = current_start, current_end, current_len

    return best_start, best_end, best_len


def save_bar(series: pd.Series, title: str, ylabel: str, out_file: Path, figsize=(11, 5), rotation=45):
    if series.empty:
        return

    fig, ax = plt.subplots(figsize=figsize)
    series.plot(kind="bar", ax=ax)
    ax.set_title(title)
    ax.set_ylabel(ylabel)
    ax.set_xlabel("")
    ax.tick_params(axis="x", rotation=rotation)
    fig.tight_layout()
    fig.savefig(out_file, dpi=150)
    plt.close(fig)


def save_line(series: pd.Series, title: str, ylabel: str, out_file: Path, figsize=(11, 5), rotation=45):
    if series.empty:
        return

    fig, ax = plt.subplots(figsize=figsize)
    series.plot(kind="line", marker="o", ax=ax)
    ax.set_title(title)
    ax.set_ylabel(ylabel)
    ax.set_xlabel("")
    ax.tick_params(axis="x", rotation=rotation)
    fig.tight_layout()
    fig.savefig(out_file, dpi=150)
    plt.close(fig)


def write_geojson(df: pd.DataFrame):
    features = []

    for _, row in df.dropna(subset=["start_lat", "start_lon"]).iterrows():
        features.append(
            {
                "type": "Feature",
                "properties": {
                    "id": str(row["id"]),
                    "date": row.get("date"),
                    "sport": row.get("sport_display"),
                    "title": row.get("title_display"),
                    "distance_km": round(float(row["distance_km"]), 3)
                    if pd.notna(row["distance_km"])
                    else None,
                    "duration": row.get("duration"),
                    "tcx": row.get("path"),
                },
                "geometry": {
                    "type": "Point",
                    "coordinates": [float(row["start_lon"]), float(row["start_lat"])],
                },
            }
        )

    collection = {"type": "FeatureCollection", "features": features}
    out_file = GEOJSON_DIR / "trainings-overview.geojson"
    out_file.write_text(json.dumps(collection, ensure_ascii=False, indent=2), encoding="utf-8")


def generate_report(df: pd.DataFrame, yearly: pd.DataFrame, monthly: pd.DataFrame):
    report = []

    total_trainings = len(df)
    total_km = df["distance_km"].sum()
    total_duration_h = df["duration_s"].sum() / 3600.0
    total_elevation_m = df["elevation_gain_m"].sum()
    first_date = df["dt"].min()
    last_date = df["dt"].max()

    report.append("# RunGPS Trainingsstatistik\n")
    report.append("_Automatisch erzeugt aus den TCX-Dateien in `RunGPSData/Trainings`._\n")

    report.append("## Überblick\n")
    report.append(f"- Trainings: **{fmt_int(total_trainings)}**")
    report.append(f"- Gesamtdistanz: **{fmt_num(total_km, 1, ' km')}**")
    report.append(f"- Gesamtdauer: **{fmt_num(total_duration_h, 1, ' h')}**")
    report.append(f"- Positive Höhenmeter, aus Trackpunkten geschätzt: **{fmt_num(total_elevation_m, 0, ' m')}**")
    if pd.notna(first_date) and pd.notna(last_date):
        report.append(f"- Zeitraum: **{first_date.date()} bis {last_date.date()}**")
    report.append(f"- Trainings mit GPS-Startpunkt: **{fmt_int(df['start_lat'].notna().sum())}**")
    report.append(f"- Trainings mit Herzfrequenzdaten: **{fmt_int(df['avg_hr'].notna().sum())}**\n")

    report.append("## Highlights\n")

    highlight_rows = []

    candidates = [
        ("Längste Strecke", best_row(df, "distance_km"), "distance_km", " km", 1),
        ("Längste Dauer", best_row(df, "duration_s"), "duration_s", "", 0),
        ("Meiste Höhenmeter", best_row(df, "elevation_gain_m"), "elevation_gain_m", " m", 0),
        ("Höchster Punkt", best_row(df, "max_altitude_m"), "max_altitude_m", " m", 0),
        ("Schnellste Ø-Geschwindigkeit", best_row(df[df["duration_s"] > 0], "avg_speed_kmh"), "avg_speed_kmh", " km/h", 1),
        ("Meiste Trackpoints", best_row(df, "trackpoints"), "trackpoints", "", 0),
    ]

    for label, row, col, suffix, decimals in candidates:
        if row is None:
            continue
        if col == "duration_s":
            value = format_duration(row[col])
        elif suffix == "":
            value = fmt_int(row[col])
        else:
            value = fmt_num(row[col], decimals, suffix)

        link = f"[{row['id']}]({row['link_from_stats']})"
        highlight_rows.append(
            {
                "Highlight": label,
                "Training": link,
                "Datum": row.get("date", ""),
                "Wert": value,
            }
        )

    if not yearly.empty:
        best_year = yearly.sort_values("distance_km", ascending=False).iloc[0]
        highlight_rows.append(
            {
                "Highlight": "Stärkstes Jahr nach Kilometern",
                "Training": str(int(best_year["year"])),
                "Datum": "",
                "Wert": fmt_num(best_year["distance_km"], 1, " km"),
            }
        )

    if not monthly.empty:
        best_month = monthly.sort_values("distance_km", ascending=False).iloc[0]
        highlight_rows.append(
            {
                "Highlight": "Stärkster Monat nach Kilometern",
                "Training": best_month["month"],
                "Datum": "",
                "Wert": fmt_num(best_month["distance_km"], 1, " km"),
            }
        )

    streak_start, streak_end, streak_len = longest_streak(list(df["dt"].dropna()))
    if streak_len:
        highlight_rows.append(
            {
                "Highlight": "Längste Serie mit Trainingstagen",
                "Training": "",
                "Datum": f"{streak_start.date()} bis {streak_end.date()}",
                "Wert": f"{streak_len} Tage",
            }
        )

    report.append(
        md_table(
            pd.DataFrame(highlight_rows),
            [
                ("Highlight", "Highlight"),
                ("Training", "Training"),
                ("Datum", "Datum"),
                ("Wert", "Wert"),
            ],
            max_rows=50,
        )
    )

    report.append("## Grafiken\n")
    report.append("![Kilometer pro Jahr](charts/km_per_year.png)\n")
    report.append("![Kilometer pro Monat](charts/km_per_month.png)\n")
    report.append("![Kumulative Kilometer](charts/cumulative_km.png)\n")
    report.append("![Trainings pro Wochentag](charts/trainings_by_weekday.png)\n")
    report.append("![Top 20 längste Trainings](charts/top20_distance.png)\n")

    report.append("## Karten\n")
    report.append("- [GitHub-Karte: Trainings-Startpunkte](geojson/trainings-overview.geojson)\n")

    report.append("## Jahresübersicht\n")
    yearly_md = yearly.copy()
    if not yearly_md.empty:
        yearly_md["distance_km"] = yearly_md["distance_km"].map(lambda x: fmt_num(x, 1))
        yearly_md["duration_h"] = yearly_md["duration_h"].map(lambda x: fmt_num(x, 1))
        yearly_md["elevation_gain_m"] = yearly_md["elevation_gain_m"].map(lambda x: fmt_num(x, 0))
    report.append(
        md_table(
            yearly_md,
            [
                ("year", "Jahr"),
                ("trainings", "Trainings"),
                ("distance_km", "km"),
                ("duration_h", "Stunden"),
                ("elevation_gain_m", "Hm+"),
            ],
            max_rows=100,
        )
    )

    report.append("## Top 20 längste Trainings\n")
    top = df.sort_values("distance_km", ascending=False).head(20).copy()
    top["Training"] = top.apply(lambda r: f"[{r['id']}]({r['link_from_stats']})", axis=1)
    top["distance_km_fmt"] = top["distance_km"].map(lambda x: fmt_num(x, 1))
    top["elevation_gain_fmt"] = top["elevation_gain_m"].map(lambda x: fmt_num(x, 0))
    top["avg_speed_fmt"] = top["avg_speed_kmh"].map(lambda x: fmt_num(x, 1))
    report.append(
        md_table(
            top,
            [
                ("Training", "Training"),
                ("date", "Datum"),
                ("sport_display", "Sport"),
                ("distance_km_fmt", "km"),
                ("duration", "Dauer"),
                ("elevation_gain_fmt", "Hm+"),
                ("avg_speed_fmt", "Ø km/h"),
            ],
            max_rows=20,
        )
    )

    report.append("## Dateien\n")
    report.append("- `trainings.csv`: alle erkannten Trainings")
    report.append("- `yearly-summary.csv`: Jahreswerte")
    report.append("- `monthly-summary.csv`: Monatswerte")
    report.append("- `highlights.json`: maschinenlesbare Highlights")
    report.append("- `geojson/trainings-overview.geojson`: Startpunktkarte\n")

    (OUT_ROOT / "README.md").write_text("\n".join(report), encoding="utf-8")


def main():
    OUT_ROOT.mkdir(parents=True, exist_ok=True)
    CHARTS_DIR.mkdir(parents=True, exist_ok=True)
    GEOJSON_DIR.mkdir(parents=True, exist_ok=True)

    tcx_files = sorted(TRAININGS_ROOT.rglob("*.tcx")) + sorted(TRAININGS_ROOT.rglob("*.TCX"))
    tcx_files = sorted(set(tcx_files))

    print(f"TCX files found: {len(tcx_files)}")

    if not tcx_files:
        (OUT_ROOT / "README.md").write_text(
            "# RunGPS Trainingsstatistik\n\nKeine TCX-Dateien unter `RunGPSData/Trainings` gefunden.\n",
            encoding="utf-8",
        )
        return

    records = []
    errors = []

    for path in tcx_files:
        try:
            records.append(parse_tcx(path))
        except Exception as exc:
            errors.append({"file": str(path.relative_to(REPO_ROOT)), "error": str(exc)})
            print(f"WARNING: could not parse {path}: {exc}")

    if not records:
        raise RuntimeError("Keine TCX-Datei konnte erfolgreich ausgewertet werden.")

    df = pd.DataFrame(records)

    metadata_path = REPO_ROOT / "RunGPSData" / "Trainings.xml"
    metadata_df = parse_rungps_trainings_xml_metadata(metadata_path)
    
    df["id"] = df["id"].astype(str)
    
    if not metadata_df.empty:
        df = df.merge(metadata_df, on="id", how="left")
    
        df["sport_display"] = df["rungps_sportart"].fillna(df["sport"])
        df["title_display"] = df["rungps_titel"].fillna("")
    else:
        df["sport_display"] = df["sport"]
        df["title_display"] = ""
        

    df["dt"] = pd.to_datetime(df["start_time"], utc=True, errors="coerce")
    df["year"] = df["dt"].dt.year
    df["month"] = df["dt"].dt.strftime("%Y-%m")
    df["weekday_num"] = df["dt"].dt.weekday
    df["weekday"] = df["weekday_num"].map(
        {
            0: "Mo",
            1: "Di",
            2: "Mi",
            3: "Do",
            4: "Fr",
            5: "Sa",
            6: "So",
        }
    )

    numeric_cols = [
        "distance_km",
        "duration_s",
        "avg_speed_kmh",
        "elevation_gain_m",
        "elevation_loss_m",
        "min_altitude_m",
        "max_altitude_m",
        "avg_hr",
        "max_hr",
        "avg_cadence",
        "max_cadence",
        "trackpoints",
        "geo_points",
        "calories",
    ]

    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    df = df.sort_values(["dt", "id"], na_position="last").reset_index(drop=True)
    df["duration_h"] = df["duration_s"] / 3600.0

    df_csv = df.drop(columns=["dt"])
    df_csv.to_csv(OUT_ROOT / "trainings.csv", index=False)

    dated = df.dropna(subset=["dt"]).copy()

    yearly = (
        dated.groupby("year", as_index=False)
        .agg(
            trainings=("id", "count"),
            distance_km=("distance_km", "sum"),
            duration_h=("duration_h", "sum"),
            elevation_gain_m=("elevation_gain_m", "sum"),
            calories=("calories", "sum"),
            avg_speed_kmh=("avg_speed_kmh", "mean"),
        )
        .sort_values("year")
    )

    monthly = (
        dated.groupby("month", as_index=False)
        .agg(
            trainings=("id", "count"),
            distance_km=("distance_km", "sum"),
            duration_h=("duration_h", "sum"),
            elevation_gain_m=("elevation_gain_m", "sum"),
            calories=("calories", "sum"),
        )
        .sort_values("month")
    )

    yearly.to_csv(OUT_ROOT / "yearly-summary.csv", index=False)
    monthly.to_csv(OUT_ROOT / "monthly-summary.csv", index=False)

    # Grafiken
    if not yearly.empty:
        yearly_plot = yearly.set_index("year")
        save_bar(yearly_plot["distance_km"], "Kilometer pro Jahr", "km", CHARTS_DIR / "km_per_year.png", rotation=0)
        save_bar(yearly_plot["trainings"], "Trainings pro Jahr", "Anzahl", CHARTS_DIR / "trainings_per_year.png", rotation=0)
        save_bar(yearly_plot["elevation_gain_m"], "Höhenmeter pro Jahr", "Hm+", CHARTS_DIR / "elevation_per_year.png", rotation=0)

    if not monthly.empty:
        monthly_plot = monthly.set_index("month")
        save_bar(monthly_plot["distance_km"].tail(36), "Kilometer pro Monat, letzte 36 Monate", "km", CHARTS_DIR / "km_per_month.png")
        cumulative = monthly_plot["distance_km"].cumsum()
        save_line(cumulative, "Kumulative Kilometer", "km", CHARTS_DIR / "cumulative_km.png")

    weekday_order = ["Mo", "Di", "Mi", "Do", "Fr", "Sa", "So"]
    weekday_counts = dated["weekday"].value_counts().reindex(weekday_order).fillna(0)
    save_bar(weekday_counts, "Trainings pro Wochentag", "Anzahl", CHARTS_DIR / "trainings_by_weekday.png", rotation=0)

    top20 = df.sort_values("distance_km", ascending=False).head(20)
    if not top20.empty:
        series = pd.Series(top20["distance_km"].values, index=top20["id"].astype(str))
        save_bar(series, "Top 20 längste Trainings", "km", CHARTS_DIR / "top20_distance.png", figsize=(12, 6))

    write_geojson(df)

    highlights = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "training_count": int(len(df)),
        "total_distance_km": float(df["distance_km"].sum()),
        "total_duration_h": float(df["duration_h"].sum()),
        "total_elevation_gain_m": float(df["elevation_gain_m"].sum()),
        "parse_errors": errors,
    }

    (OUT_ROOT / "highlights.json").write_text(
        json.dumps(highlights, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    if errors:
        pd.DataFrame(errors).to_csv(OUT_ROOT / "parse-errors.csv", index=False)

    generate_report(df, yearly, monthly)

    print(f"Stats written to: {OUT_ROOT}")
    print(f"Trainings parsed: {len(df)}")
    if errors:
        print(f"Parse warnings: {len(errors)}")


if __name__ == "__main__":
    main()
