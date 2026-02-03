#!/usr/bin/env python3
"""Generate a burndown chart from a Microsoft Planner .xlsx export and a sprints CSV.

Usage:
    python burndown.py              # all tasks, gaps compressed
    python burndown.py --sprints    # only tasks within sprint ranges
    python burndown.py --sprint     # only the most recent sprint
"""

import argparse
import sys
import tomllib
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
import matplotlib.patches as mpatches
import plotly.graph_objects as go
from datetime import datetime, timedelta
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent
SPRINTS_PATH = SCRIPT_DIR / "sprints.csv"
CONFIG_PATH = SCRIPT_DIR / "config.toml"

LINE_STYLES = {
    "solid": "-",
    "dashed": "--",
    "dotted": ":",
    "dash-dot": "-.",
}

PLOTLY_LINE_STYLES = {
    "solid": "solid",
    "dashed": "dash",
    "dotted": "dot",
    "dash-dot": "dashdot",
}


def load_config():
    with open(CONFIG_PATH, "rb") as f:
        return tomllib.load(f)


def normalise_colour(name):
    """Turn human-friendly names like 'light blue' into matplotlib's 'lightblue'."""
    return name.strip().lower().replace(" ", "")


def colour_to_rgba(name, opacity=1.0):
    """Convert a named colour to an rgba() string for plotly."""
    try:
        rgb = mcolors.to_rgb(normalise_colour(name))
    except ValueError:
        rgb = mcolors.to_rgb("grey")
    r, g, b = (int(c * 255) for c in rgb)
    return f"rgba({r},{g},{b},{opacity})"


def find_xlsx():
    files = sorted(SCRIPT_DIR.glob("*.xlsx"))
    if not files:
        sys.exit("Error: no .xlsx file found in script directory.")
    if len(files) == 1:
        return files[0]
    print("Multiple .xlsx files found:")
    for i, f in enumerate(files, 1):
        print(f"  {i}) {f.name}")
    while True:
        choice = input(f"Select file [1-{len(files)}] (default 1): ").strip()
        if choice == "":
            return files[0]
        if choice.isdigit() and 1 <= int(choice) <= len(files):
            return files[int(choice) - 1]
        print(f"Invalid choice. Enter a number between 1 and {len(files)}.")


def load_plan_meta(path):
    """Return (plan_name, export_date) from the 'Plan name' sheet."""
    df = pd.read_excel(path, sheet_name="Plan name")
    plan_name = df.columns[1]
    export_date = None
    date_row = df[df["Plan name"] == "Date of export"]
    if not date_row.empty:
        export_date = pd.to_datetime(
            date_row.iloc[0, 1], format="%m/%d/%Y", errors="coerce"
        )
    return plan_name, export_date


def load_tasks(path, export_date=None):
    df = pd.read_excel(path, sheet_name="Tasks")
    for col in ("Created Date", "Due date", "Completed Date"):
        df[col] = pd.to_datetime(df[col], format="%m/%d/%Y", errors="coerce")

    # Fallback: if Progress or Bucket Name says completed but no date, use export date
    completed_mask = (
        (df["Progress"].str.lower() == "completed")
        | (df["Bucket Name"].str.lower() == "completed")
    )
    missing_date = df["Completed Date"].isna()
    fallback_date = export_date if export_date else pd.Timestamp(datetime.now().date())
    df.loc[completed_mask & missing_date, "Completed Date"] = fallback_date

    return df


def load_sprints():
    df = pd.read_csv(SPRINTS_PATH)
    df["start_date"] = pd.to_datetime(df["start_date"])
    df["end_date"] = pd.to_datetime(df["end_date"])
    return df.sort_values("start_date").reset_index(drop=True)


def assign_sprint(due_date, sprints):
    if pd.isna(due_date):
        return None
    for _, sprint in sprints.iterrows():
        if sprint["start_date"] <= due_date <= sprint["end_date"]:
            return sprint["id"]
    return None


def filter_to_sprints(df, sprints):
    """Keep only tasks whose due date falls within any sprint range."""
    mask = df["Due date"].apply(lambda d: assign_sprint(d, sprints)).notna()
    return df[mask].copy()


def data_date_range(df):
    candidates = [df["Created Date"].min().normalize()]
    if df["Completed Date"].dropna().shape[0]:
        candidates.append(df["Completed Date"].max().normalize())
    if df["Due date"].dropna().shape[0]:
        candidates.append(df["Due date"].max().normalize())
    candidates.append(pd.Timestamp(datetime.now().date()))
    return min(candidates), max(candidates)


def find_gaps(df, threshold):
    """Find date ranges where nothing happens, longer than threshold days."""
    event_dates = set()
    for col in ("Created Date", "Due date", "Completed Date"):
        event_dates.update(df[col].dropna().dt.normalize().tolist())
    if not event_dates:
        return []
    event_dates.add(pd.Timestamp(datetime.now().date()))

    sorted_dates = sorted(event_dates)
    gaps = []
    for i in range(len(sorted_dates) - 1):
        diff = (sorted_dates[i + 1] - sorted_dates[i]).days
        if diff > threshold:
            gap_start = sorted_dates[i] + timedelta(days=1)
            gap_end = sorted_dates[i + 1] - timedelta(days=1)
            gaps.append((gap_start, gap_end))
    return gaps


def build_burndown(df, start, end, gaps=None):
    dates = pd.date_range(start, end, freq="D")

    if gaps:
        gap_set = set()
        for gs, ge in gaps:
            gap_set.update(pd.date_range(gs, ge, freq="D").tolist())
        dates = dates[~dates.isin(gap_set)]

    total = len(df)
    remaining = []
    for day in dates:
        completed = df[df["Completed Date"].notna() & (df["Completed Date"] <= day)]
        remaining.append(total - len(completed))

    return pd.DataFrame({"date": dates, "remaining": remaining})


def build_ideal(df, sprints, total_tasks, start, end):
    points = []
    tasks_remaining = total_tasks

    unsprinted = len(df[df["sprint"].isna()])
    if unsprinted > 0 and not sprints.empty:
        first_sprint_start = max(sprints["start_date"].min(), start)
        points.append((start, tasks_remaining))
        tasks_remaining -= unsprinted
        points.append((first_sprint_start, tasks_remaining))

    for _, sprint in sprints.iterrows():
        sprint_tasks = len(df[df["sprint"] == sprint["id"]])
        if sprint_tasks == 0:
            continue
        s = max(sprint["start_date"], start)
        e = min(sprint["end_date"], end)
        points.append((s, tasks_remaining))
        tasks_remaining -= sprint_tasks
        points.append((e, tasks_remaining))

    if not points:
        points = [(start, total_tasks), (end, 0)]

    return pd.DataFrame({"date": [p[0] for p in points], "remaining": [p[1] for p in points]})


def visible_sprints(sprints, start, end):
    mask = (sprints["end_date"] >= start) & (sprints["start_date"] <= end)
    return sprints[mask].copy()


def _date_to_idx(date, date_list):
    if date <= date_list[0]:
        return 0
    if date >= date_list[-1]:
        return len(date_list) - 1
    for i in range(len(date_list) - 1):
        if date_list[i] <= date <= date_list[i + 1]:
            span = (date_list[i + 1] - date_list[i]).days
            if span == 0:
                return i
            frac = (date - date_list[i]).days / span
            return i + frac
    return len(date_list) - 1


HATCH_PATTERNS = {
    "diagonal": "//",
    "cross": "xx",
    "dots": "..",
}


def _find_gap_positions(date_list, gaps):
    """Return list of (index, days_skipped) for each gap in the compressed date list."""
    if not gaps:
        return []
    positions = []
    for i in range(len(date_list) - 1):
        day_jump = (date_list[i + 1] - date_list[i]).days
        if day_jump > 1:
            # Find which gap this corresponds to
            for gs, ge in gaps:
                if date_list[i] < gs and date_list[i + 1] > ge:
                    positions.append((i + 0.5, day_jump))
                    break
    return positions


def plot_matplotlib(burndown, ideal, total_tasks, sprints, title, start, end, cfg, gaps=None):
    chart = cfg["chart"]
    fig, ax = plt.subplots(figsize=(chart["width"], chart["height"]))

    date_list = burndown["date"].tolist()
    x_actual = list(range(len(date_list)))

    sprint_colours = cfg["sprints"]["colours"]
    sprint_opacity = cfg["sprints"]["opacity"]

    for i, (_, sprint) in enumerate(sprints.iterrows()):
        s = max(sprint["start_date"], start)
        e = min(sprint["end_date"], end)
        if s >= end or e <= start:
            continue
        si = _date_to_idx(s, date_list)
        ei = _date_to_idx(e, date_list)
        ax.axvspan(
            si, ei, alpha=sprint_opacity,
            color=normalise_colour(sprint_colours[i % len(sprint_colours)]),
            label=f"Sprint {sprint['id']}",
        )

    ideal_cfg = cfg["ideal_line"]
    ideal_x = [_date_to_idx(d, date_list) for d in ideal["date"]]
    ax.plot(
        ideal_x, ideal["remaining"],
        linestyle=LINE_STYLES.get(ideal_cfg["style"], "--"),
        color=normalise_colour(ideal_cfg["colour"]),
        linewidth=ideal_cfg["thickness"],
        label="Ideal",
    )

    actual_cfg = cfg["actual_line"]
    ax.step(
        x_actual, burndown["remaining"],
        where="post",
        color=normalise_colour(actual_cfg["colour"]),
        linewidth=actual_cfg["thickness"],
        label="Actual",
    )

    ax.set_title(title, fontsize=14, fontweight="bold")
    ax.set_xlabel("Date")
    ax.set_ylabel("Tasks Remaining")
    ax.set_ylim(bottom=0, top=total_tasks + 1)

    date_fmt = cfg["dates"]["format"]
    angle = cfg["dates"]["label_angle"]
    if len(date_list) <= 15:
        ax.set_xticks(x_actual)
        ax.set_xticklabels([d.strftime(date_fmt) for d in date_list], rotation=angle, ha="right")
    else:
        step = max(1, len(date_list) // 10)
        ticks = list(range(0, len(date_list), step))
        ax.set_xticks(ticks)
        ax.set_xticklabels([date_list[t].strftime(date_fmt) for t in ticks], rotation=angle, ha="right")

    # Gap break markers
    if gaps:
        gap_cfg = cfg["gaps"]
        marker_colour = normalise_colour(gap_cfg["marker_colour"])
        label_colour = normalise_colour(gap_cfg["marker_label_colour"])
        hatch = HATCH_PATTERNS.get(gap_cfg.get("marker_pattern", "diagonal"), "//")
        gap_positions = _find_gap_positions(date_list, gaps)
        for idx, days in gap_positions:
            w = 0.4
            rect = mpatches.FancyBboxPatch(
                (idx - w / 2, 0), w, total_tasks + 1,
                boxstyle="square,pad=0",
                facecolor=marker_colour, edgecolor="grey",
                alpha=0.7, hatch=hatch, linewidth=0.5,
            )
            ax.add_patch(rect)
            ax.annotate(
                f" {days}d ", (idx, total_tasks * 0.5),
                ha="center", va="center", fontsize=8, fontweight="bold",
                color=label_colour,
                bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="grey", alpha=0.9),
            )

    ax.legend(loc=cfg["legend"]["position"])

    grid_cfg = cfg["grid"]
    if grid_cfg["show"]:
        ax.grid(axis="y", alpha=grid_cfg["opacity"])

    out = SCRIPT_DIR / "burndown.png"
    fig.savefig(out, dpi=chart["dpi"], bbox_inches="tight")
    plt.close(fig)
    print(f"PNG saved: {out}")


def plot_plotly(burndown, ideal, total_tasks, sprints, title, start, end, cfg, gaps=None):
    fig = go.Figure()

    sprint_colours = cfg["sprints"]["colours"]
    sprint_opacity = cfg["sprints"]["opacity"]
    label_colour = cfg["sprints"]["label_colour"]
    label_size = cfg["sprints"]["label_size"]

    for i, (_, sprint) in enumerate(sprints.iterrows()):
        s = max(sprint["start_date"], start)
        e = min(sprint["end_date"], end)
        if s >= end or e <= start:
            continue
        colour_name = sprint_colours[i % len(sprint_colours)]
        fig.add_vrect(
            x0=s, x1=e,
            fillcolor=colour_to_rgba(colour_name, sprint_opacity),
            line_width=0,
            layer="below",
            annotation_text=f"Sprint {sprint['id']}",
            annotation_position="top left",
            annotation_font_size=label_size,
            annotation_font_color=label_colour,
        )

    ideal_cfg = cfg["ideal_line"]
    fig.add_trace(go.Scatter(
        x=ideal["date"], y=ideal["remaining"],
        mode="lines",
        line=dict(
            dash=PLOTLY_LINE_STYLES.get(ideal_cfg["style"], "dash"),
            color=normalise_colour(ideal_cfg["colour"]),
            width=ideal_cfg["thickness"],
        ),
        name="Ideal",
    ))

    actual_cfg = cfg["actual_line"]
    fig.add_trace(go.Scatter(
        x=burndown["date"], y=burndown["remaining"],
        mode="lines",
        line=dict(shape="hv", color=normalise_colour(actual_cfg["colour"]), width=actual_cfg["thickness"]),
        name="Actual",
        hovertemplate="Date: %{x|" + cfg["dates"]["format"] + "}<br>Remaining: %{y}<extra></extra>",
    ))

    rangebreaks = []
    if gaps:
        gap_cfg = cfg["gaps"]
        marker_colour = gap_cfg["marker_colour"]
        label_colour = normalise_colour(gap_cfg["marker_label_colour"])
        for gs, ge in gaps:
            rangebreaks.append(dict(bounds=[gs.isoformat(), (ge + timedelta(days=1)).isoformat()]))
            days = (ge - gs).days + 1
            # Marker line at the boundary just before the gap
            boundary = gs - timedelta(days=1)
            fig.add_vline(
                x=boundary.isoformat(),
                line=dict(color=normalise_colour(marker_colour), width=2, dash="dot"),
                layer="below",
            )
            fig.add_annotation(
                x=boundary.isoformat(),
                y=1.0, yref="paper",
                text=f"<b>{days}d</b>",
                showarrow=False,
                font=dict(size=10, color=label_colour),
                bgcolor=colour_to_rgba(marker_colour, 0.6),
                bordercolor="grey",
                borderwidth=1,
                borderpad=3,
                yanchor="top",
            )

    grid_cfg = cfg["grid"]
    fig.update_layout(
        title=title,
        xaxis_title="Date",
        yaxis_title="Tasks Remaining",
        yaxis=dict(
            range=[0, total_tasks + 1],
            showgrid=grid_cfg["show"],
            gridcolor=colour_to_rgba("grey", grid_cfg["opacity"]) if grid_cfg["show"] else None,
        ),
        xaxis=dict(
            tickformat=cfg["dates"]["format"],
            rangebreaks=rangebreaks if rangebreaks else None,
        ),
        template="plotly_white",
        hovermode="x unified",
    )

    out = SCRIPT_DIR / "burndown.html"
    fig.write_html(str(out))
    print(f"HTML saved: {out}")


def main():
    parser = argparse.ArgumentParser(description="Burndown chart generator")
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--sprints", action="store_true",
                       help="Only tasks within sprint ranges")
    group.add_argument("--sprint", action="store_true",
                       help="Only the most recent sprint")
    parser.add_argument("--compress", nargs="?", type=int, const=-1, default=None,
                        metavar="DAYS",
                        help="Compress gaps of inactivity (optionally specify minimum days, default from config)")
    args = parser.parse_args()

    cfg = load_config()

    xlsx_path = find_xlsx()
    plan_name, export_date = load_plan_meta(xlsx_path)
    df = load_tasks(xlsx_path, export_date)
    sprints = load_sprints()

    df["sprint"] = df["Due date"].apply(lambda d: assign_sprint(d, sprints))

    if args.sprint:
        today = pd.Timestamp(datetime.now().date())
        active = sprints[(sprints["start_date"] <= today) & (sprints["end_date"] >= today)]
        if not active.empty:
            latest = active.iloc[0]
        else:
            upcoming = sprints[sprints["start_date"] > today]
            if not upcoming.empty:
                latest = upcoming.iloc[0]
            else:
                latest = sprints.iloc[-1]
        sprints = sprints[sprints["id"] == latest["id"]].reset_index(drop=True)
        df = df[df["sprint"] == latest["id"]].copy()
        print(f"Mode: single sprint (Sprint {latest['id']})")
    elif args.sprints:
        df = filter_to_sprints(df, sprints)
        active_ids = df["sprint"].dropna().unique()
        sprints = sprints[sprints["id"].isin(active_ids)].reset_index(drop=True)
        print("Mode: all sprints")
    else:
        print("Mode: all tasks")

    if df.empty:
        sys.exit("No tasks matched the filter.")

    title = cfg["chart"]["title"].format(plan=plan_name)

    print(f"Plan: {plan_name}")
    print(f"Found {len(df)} tasks:")
    for _, row in df.iterrows():
        status = "Done" if pd.notna(row["Completed Date"]) else "Open"
        sp = f"Sprint {int(row['sprint'])}" if pd.notna(row["sprint"]) else "No sprint"
        print(f"  [{status}] {row['Task Name']}  ({sp})")

    start, end = data_date_range(df)
    if args.compress is not None:
        gap_threshold = args.compress if args.compress > 0 else cfg["gaps"]["compress_after_days"]
        gaps = find_gaps(df, gap_threshold)
    else:
        gaps = []

    burndown = build_burndown(df, start, end, gaps=gaps)
    total_tasks = len(df)
    ideal = build_ideal(df, sprints, total_tasks, start, end)
    vis_sprints = visible_sprints(sprints, start, end)

    print(f"\nRange: {start.date()} -> {end.date()}")
    if gaps:
        for gs, ge in gaps:
            print(f"  Gap compressed: {gs.date()} -> {ge.date()}")
    print(f"Total: {total_tasks} | Remaining: {burndown['remaining'].iloc[-1]}")

    plot_matplotlib(burndown, ideal, total_tasks, vis_sprints, title, start, end, cfg, gaps=gaps)
    plot_plotly(burndown, ideal, total_tasks, vis_sprints, title, start, end, cfg, gaps=gaps)


if __name__ == "__main__":
    main()
