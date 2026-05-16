from pathlib import Path
import math
import re

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


EXPORT_DIR = Path("MassExports")
TARGET_WORKBOOK = Path("MassProps.xlsm")
KG_TO_LBM = 2.2046226218
M_TO_FT = 3.280839895
KG_M_TO_LBM_FT = KG_TO_LBM * M_TO_FT
FIRST_IMPERIAL_EXPORT_VERSION = (0, 1, 6)
FUEL_ATA_LABEL = "28"
P95_Z_SCORE = 1.6448536269514722
REQUIRED_COLUMNS = {
    "ata",
    "component",
    "total_mass_kg",
    "mx_kgm",
    "my_kgm",
    "mz_kgm",
}
NUMERIC_COLUMNS = [
    "ata",
    "qty",
    "unit_mass_kg",
    "total_mass_kg",
    "x_m",
    "y_m",
    "z_m",
    "mx_kgm",
    "my_kgm",
    "mz_kgm",
]
DATE_COLUMNS = ["last_updated"]
DISPLAY_COLUMNS = [
    "ata_display",
    "component",
    "description",
    "qty",
    "unit_mass_kg",
    "total_mass_kg",
    "share_pct",
    "x_m",
    "y_m",
    "z_m",
    "notes",
    "export_version",
    "status",
]
ATA_NAME_FALLBACK = {
    "21": "Air Conditioning",
    "22": "Auto Flight",
    "23": "Communications",
    "24": "Electrical Power",
    "25": "Equipment / Furnishings",
    "26": "Fire Protection",
    "27": "Flight Controls",
    "28": "Fuel",
    "29": "Hydraulic Power",
    "30": "Ice & Rain Protection",
    "31": "Indicating / Recording",
    "32": "Landing Gear",
    "33": "Lights",
    "34": "Navigation",
    "35": "Oxygen",
    "36": "Pneumatic",
    "38": "Water / Waste",
    "49": "APU",
    "52": "Doors",
    "53": "Fuselage",
    "54": "Nacelles / Pylons",
    "55": "Stabilizers",
    "56": "Windows",
    "57": "Wings",
    "71": "Power Plant",
    "73": "Engine Fuel & Control",
    "75": "Engine Air",
    "78": "Exhaust",
    "79": "Engine Oil",
    "80": "Starting",
}
RISK_VALUE_ALIASES = {
    "target_lbm": ("target", "target oew", "oew target", "target mass"),
    "mean_lbm": ("adjusted mean", "mean", "status mean", "oew mean"),
    "p95_lbm": ("p95", "p95 oew", "95th percentile", "95 percentile"),
    "sigma_lbm": ("sigma", "std dev", "standard deviation", "stdev"),
}


def export_version_from_path(path: Path) -> str:
    version_match = re.search(r"v\d+(?:\.\d+)+", path.stem)
    return version_match.group(0) if version_match else path.stem


def version_sort_key(version: str) -> tuple[int, ...]:
    return tuple(int(part) for part in re.findall(r"\d+", str(version)))


def is_excel_error(value: object) -> bool:
    return bool(re.fullmatch(r"Error\s+\d+", str(value).strip(), flags=re.IGNORECASE))


def ata_display_label(ata_label: str, ata_names: dict[str, str]) -> str:
    ata_name = ata_names.get(str(ata_label), "")
    return f"ATA {ata_label} {ata_name}".strip()


def normalize_legacy_metric_export(frame: pd.DataFrame) -> pd.DataFrame:
    version = frame["export_version"].iloc[0]
    if version_sort_key(version) >= FIRST_IMPERIAL_EXPORT_VERSION:
        return frame

    frame = frame.copy()
    for column in ["unit_mass_kg", "total_mass_kg"]:
        if column in frame.columns:
            frame[column] = frame[column] * KG_TO_LBM
    for column in ["x_m", "y_m", "z_m"]:
        if column in frame.columns:
            frame[column] = frame[column] * M_TO_FT
    for column in ["mx_kgm", "my_kgm", "mz_kgm"]:
        if column in frame.columns:
            frame[column] = frame[column] * KG_M_TO_LBM_FT
    return frame


@st.cache_data(show_spinner=False)
def load_exports(export_dir: str) -> tuple[pd.DataFrame, list[str]]:
    rows = []
    warnings = []
    for path in sorted(Path(export_dir).glob("*.csv")):
        try:
            frame = pd.read_csv(path)
        except Exception as exc:
            warnings.append(f"Skipping {path.name}; could not read CSV: {exc}")
            continue

        missing_columns = REQUIRED_COLUMNS.difference(frame.columns)
        if missing_columns:
            warnings.append(
                f"Skipping {path.name}; missing columns: {', '.join(sorted(missing_columns))}"
            )
            continue

        frame = frame.copy()
        frame["export_file"] = path.name
        frame["iteration"] = path.stem
        frame["export_version"] = export_version_from_path(path)
        for column in NUMERIC_COLUMNS:
            if column in frame.columns:
                frame[column] = pd.to_numeric(frame[column], errors="coerce")
        for column in DATE_COLUMNS:
            if column in frame.columns:
                frame[column] = pd.to_datetime(frame[column], errors="coerce")
        frame = normalize_legacy_metric_export(frame)
        rows.append(frame)

    if not rows:
        return pd.DataFrame(), warnings

    exports = pd.concat(rows, ignore_index=True)
    for column in DISPLAY_COLUMNS:
        if column not in exports.columns:
            exports[column] = ""

    exports["ata_label"] = exports["ata"].fillna(-1).astype(int).astype(str)
    exports["component"] = exports["component"].fillna("Unspecified")
    exports["description"] = exports["description"].fillna("").map(
        lambda value: "" if is_excel_error(value) else value
    )
    exports["notes"] = exports["notes"].fillna("")
    exports["status"] = exports["status"].fillna("")
    exports["total_mass_kg"] = exports["total_mass_kg"].fillna(0.0)
    if "last_updated" not in exports.columns:
        exports["last_updated"] = pd.NaT
    return exports, warnings


@st.cache_data(show_spinner=False)
def load_targets(workbook_path: str) -> tuple[pd.DataFrame, pd.DataFrame, list[str]]:
    warnings = []
    path = Path(workbook_path)
    if not path.exists():
        return pd.DataFrame(), pd.DataFrame(), [f"{path.name} not found; target charts hidden."]

    try:
        raw = pd.read_excel(path, sheet_name="Target", header=None)
    except Exception as exc:
        return pd.DataFrame(), pd.DataFrame(), [f"Could not read Target sheet: {exc}"]

    ata_rows = []
    condition_rows = []
    section = None
    for row in raw.itertuples(index=False, name=None):
        first = row[0]
        if first == "ATA":
            section = "ata"
            continue
        if first == "Code":
            section = "condition"
            continue
        if pd.isna(first):
            continue

        if section == "ata" and isinstance(first, (int, float)):
            ata_rows.append(
                {
                    "ata_label": str(int(first)),
                    "ata_name": str(row[1]),
                    "target_mass_lbm": pd.to_numeric(row[2], errors="coerce") * KG_TO_LBM,
                    "workbook_current_lbm": pd.to_numeric(row[3], errors="coerce") * KG_TO_LBM,
                }
            )
        elif section == "condition" and isinstance(first, str):
            condition_rows.append(
                {
                    "code": first,
                    "description": str(row[1]),
                    "target_mass_lbm": pd.to_numeric(row[2], errors="coerce") * KG_TO_LBM,
                    "workbook_current_lbm": pd.to_numeric(row[3], errors="coerce") * KG_TO_LBM,
                }
            )

    ata_targets = pd.DataFrame(ata_rows)
    aircraft_targets = pd.DataFrame(condition_rows)
    if ata_targets.empty:
        warnings.append("No ATA targets found in the Target sheet.")
    if aircraft_targets.empty:
        warnings.append("No aircraft-level targets found in the Target sheet.")
    return ata_targets, aircraft_targets, warnings


def normalized_label(value: object) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(value).lower()).strip()


def risk_key_for_label(label: str) -> str | None:
    normalized = normalized_label(label)
    for key, aliases in RISK_VALUE_ALIASES.items():
        if any(alias in normalized for alias in aliases):
            return key
    return None


def first_numeric_after(row: tuple[object, ...], start_index: int) -> float | None:
    for value in row[start_index + 1 :]:
        numeric = pd.to_numeric(value, errors="coerce")
        if pd.notna(numeric):
            return float(numeric)
    return None


@st.cache_data(show_spinner=False)
def load_risk(workbook_path: str) -> tuple[dict[str, float], list[str]]:
    path = Path(workbook_path)
    if not path.exists():
        return {}, [f"{path.name} not found; risk plot hidden."]

    try:
        raw = pd.read_excel(path, sheet_name="Risk", header=None)
    except ValueError:
        return {}, [f"{path.name} has no Risk sheet; risk plot hidden."]
    except Exception as exc:
        return {}, [f"Could not read Risk sheet: {exc}"]

    values = {}
    for row in raw.itertuples(index=False, name=None):
        for index, cell in enumerate(row):
            key = risk_key_for_label(cell)
            if key and key not in values:
                numeric = first_numeric_after(row, index)
                if numeric is not None:
                    values[key] = numeric

    if "sigma_lbm" not in values and {"mean_lbm", "p95_lbm"}.issubset(values):
        values["sigma_lbm"] = abs(values["p95_lbm"] - values["mean_lbm"]) / P95_Z_SCORE

    required = {"target_lbm", "mean_lbm", "p95_lbm", "sigma_lbm"}
    missing = sorted(required.difference(values))
    if missing:
        return {}, [f"Risk sheet is missing: {', '.join(missing)}."]
    if values["sigma_lbm"] <= 0:
        return {}, ["Risk sheet sigma must be greater than zero."]

    return values, []


def mass_properties(frame: pd.DataFrame) -> dict[str, float]:
    total_mass = frame["total_mass_kg"].sum()
    if total_mass == 0:
        return {"mass": 0.0, "x": 0.0, "y": 0.0, "z": 0.0}

    return {
        "mass": total_mass,
        "x": frame["mx_kgm"].sum() / total_mass,
        "y": frame["my_kgm"].sum() / total_mass,
        "z": frame["mz_kgm"].sum() / total_mass,
    }


def oew_frame(frame: pd.DataFrame) -> pd.DataFrame:
    return frame[frame["ata_label"] != FUEL_ATA_LABEL]


def target_condition_value(aircraft_targets: pd.DataFrame, code: str) -> float | None:
    if aircraft_targets.empty:
        return None
    rows = aircraft_targets[aircraft_targets["code"].eq(code)]
    if rows.empty:
        return None
    value = rows["workbook_current_lbm"].iloc[0]
    return None if pd.isna(value) else float(value)


def summarize_by(frame: pd.DataFrame, group_column: str) -> pd.DataFrame:
    summary = (
        frame.groupby(group_column, dropna=False)
        .agg(
            total_mass_kg=("total_mass_kg", "sum"),
            item_count=("component", "count"),
            mx_kgm=("mx_kgm", "sum"),
            my_kgm=("my_kgm", "sum"),
            mz_kgm=("mz_kgm", "sum"),
        )
        .reset_index()
        .sort_values("total_mass_kg", ascending=False)
    )
    total_mass = summary["total_mass_kg"].sum()
    summary["share_pct"] = summary["total_mass_kg"] / total_mass * 100 if total_mass else 0
    return summary


def render_header_rows(
    latest: pd.DataFrame,
    latest_iteration: str,
    properties: dict[str, float],
    aircraft_targets: pd.DataFrame,
) -> None:
    versions = sorted(
        str(value) for value in latest["export_version"].dropna().unique() if str(value)
    )
    statuses = sorted(str(value) for value in latest["status"].dropna().unique() if str(value))

    columns = st.columns(2)
    columns[0].metric("Version", ", ".join(versions) if versions else latest_iteration)
    columns[1].metric("Status", ", ".join(statuses) if statuses else "Unspecified")

    oew = properties["mass"]
    mtow = target_condition_value(aircraft_targets, "MTOW")
    zfw = target_condition_value(aircraft_targets, "ZFW")
    columns = st.columns(3)
    columns[0].metric("OEW", f"{oew:,.1f} lbm")
    columns[1].metric("MTOW", f"{mtow:,.1f} lbm" if mtow is not None else "N/A")
    columns[2].metric("ZFW", f"{zfw:,.1f} lbm" if zfw is not None else "N/A")

    columns = st.columns(3)
    columns[0].metric("CG X", f"{properties['x']:.2f} ft")
    columns[1].metric("CG Y", f"{properties['y']:.2f} ft")
    columns[2].metric("CG Z", f"{properties['z']:.2f} ft")


def render_ata_pie(summary: pd.DataFrame) -> None:
    chart = px.pie(
        summary,
        names="ata_display",
        values="total_mass_kg",
        hole=0.34,
        labels={
            "ata_display": "ATA",
            "total_mass_kg": "Mass (lbm)",
            "share_pct": "Share (%)",
        },
        custom_data=["share_pct", "item_count"],
    )
    chart.update_traces(
        textposition="inside",
        textinfo="label+percent",
        hovertemplate="<b>%{label}</b><br>Mass: %{value:,.1f} lbm"
        "<br>Share: %{customdata[0]:.1f}%<br>Components: %{customdata[1]:,}<extra></extra>",
    )
    chart.update_layout(height=500, margin=dict(l=10, r=10, t=20, b=10), showlegend=True)
    st.plotly_chart(chart, use_container_width=True)


def render_component_bar(summary: pd.DataFrame, top_n: int) -> None:
    visible_summary = summary.head(top_n)
    chart = px.bar(
        visible_summary.sort_values("total_mass_kg"),
        x="total_mass_kg",
        y="component_display",
        orientation="h",
        text="total_mass_kg",
        labels={
            "total_mass_kg": "Mass (lbm)",
            "component_display": "Component",
            "share_pct": "Share (%)",
        },
        custom_data=["share_pct", "item_count"],
    )
    chart.update_traces(
        texttemplate="%{text:,.0f} lbm",
        textposition="outside",
        hovertemplate="<b>%{y}</b><br>Mass: %{x:,.1f} lbm"
        "<br>Share: %{customdata[0]:.1f}%<br>Rows: %{customdata[1]:,}<extra></extra>",
    )
    chart.update_layout(
        height=max(420, 26 * len(visible_summary)),
        margin=dict(l=10, r=10, t=20, b=10),
    )
    st.plotly_chart(chart, use_container_width=True)


def render_ata_tables(latest: pd.DataFrame, ata_summary: pd.DataFrame) -> None:
    total_mass = latest["total_mass_kg"].sum()
    st.subheader("ATA Component Tables")

    for row in ata_summary.itertuples(index=False):
        ata_frame = latest[latest["ata_label"] == row.ata_label].copy()
        ata_frame = ata_frame.sort_values(["total_mass_kg", "component"], ascending=[False, True])
        ata_frame["share_pct"] = (
            ata_frame["total_mass_kg"] / total_mass * 100 if total_mass else 0.0
        )
        table_columns = [column for column in DISPLAY_COLUMNS if column in ata_frame.columns]
        table = ata_frame[table_columns].rename(
            columns={
                "ata_display": "ATA",
                "component": "Component",
                "description": "Description",
                "qty": "Qty",
                "unit_mass_kg": "Unit mass (lbm)",
                "total_mass_kg": "Mass (lbm)",
                "share_pct": "Share (%)",
                "x_m": "X (ft)",
                "y_m": "Y (ft)",
                "z_m": "Z (ft)",
                "notes": "Notes",
                "export_version": "Version",
                "status": "Status",
            }
        )

        with st.expander(
            f"{row.ata_display} | {row.total_mass_kg:,.1f} lbm "
            f"({row.share_pct:.1f}%, {row.item_count:,} components)",
            expanded=True,
        ):
            st.dataframe(
                table.style.format(
                    {
                        "Qty": "{:,.0f}",
                        "Unit mass (lbm)": "{:,.1f}",
                        "Mass (lbm)": "{:,.1f}",
                        "Share (%)": "{:.1f}",
                        "X (ft)": "{:.2f}",
                        "Y (ft)": "{:.2f}",
                        "Z (ft)": "{:.2f}",
                    },
                    na_rep="",
                ),
                use_container_width=True,
                hide_index=True,
            )


def render_iteration_trend(exports: pd.DataFrame, aircraft_targets: pd.DataFrame) -> None:
    trend = (
        exports.groupby(["export_version", "iteration"], sort=False)
        .apply(
            lambda frame: pd.Series(
                {
                    **mass_properties(oew_frame(frame)),
                    "export_date": frame["last_updated"].dropna().min(),
                }
            ),
            include_groups=False,
        )
        .reset_index()
    )
    trend = trend.sort_values(
        "export_version",
        key=lambda series: series.map(version_sort_key),
    )

    oew_target = pd.NA
    if not aircraft_targets.empty:
        oew_rows = aircraft_targets[aircraft_targets["code"].eq("OEW")]
        if not oew_rows.empty:
            oew_target = oew_rows["target_mass_lbm"].iloc[0]

    figure = go.Figure()
    figure.add_trace(
        go.Scatter(
            x=trend["export_date"],
            y=trend["mass"],
            mode="lines+markers+text",
            name="Status OEW",
            text=trend["export_version"],
            textposition="top center",
            customdata=trend[["iteration", "export_version"]],
            hovertemplate="<b>%{customdata[1]}</b><br>Date: %{x|%Y-%m-%d}"
            "<br>Status OEW: %{y:,.1f} lbm<extra></extra>",
        )
    )
    if pd.notna(oew_target):
        figure.add_trace(
            go.Scatter(
                x=trend["export_date"],
                y=[oew_target] * len(trend),
                mode="lines+markers",
                name="Target OEW",
                hovertemplate="Date: %{x|%Y-%m-%d}<br>Target OEW: %{y:,.1f} lbm"
                "<extra></extra>",
            )
        )
    figure.update_layout(
        xaxis_title="Date",
        yaxis_title="OEW (lbm)",
        height=460,
        margin=dict(l=10, r=10, t=20, b=10),
    )
    st.subheader("Version History")
    st.plotly_chart(figure, use_container_width=True)


def normal_density_percent_per_100_lbm(
    x_values: np.ndarray,
    mean: float,
    sigma: float,
) -> np.ndarray:
    coefficient = 1 / (sigma * math.sqrt(2 * math.pi))
    density = coefficient * np.exp(-0.5 * ((x_values - mean) / sigma) ** 2)
    return density * 100 * 100


def render_risk_plot(risk_values: dict[str, float]) -> None:
    if not risk_values:
        return

    target = risk_values["target_lbm"]
    mean = risk_values["mean_lbm"]
    p95 = risk_values["p95_lbm"]
    sigma = risk_values["sigma_lbm"]
    x_min = min(target, mean, p95) - 4 * sigma
    x_max = max(target, mean, p95) + 4 * sigma
    x_values = np.linspace(x_min, x_max, 500)
    y_values = normal_density_percent_per_100_lbm(x_values, mean, sigma)

    figure = go.Figure()
    figure.add_trace(
        go.Scatter(
            x=x_values,
            y=y_values,
            mode="lines",
            line=dict(color="#5eaa3a", width=2),
            name="Risk density",
            hovertemplate="OEW: %{x:,.1f} lbm<br>Density: %{y:.1f}% per 100 lbm"
            "<extra></extra>",
        )
    )
    tail_mask = x_values >= target
    figure.add_trace(
        go.Scatter(
            x=x_values[tail_mask],
            y=y_values[tail_mask],
            mode="lines",
            line=dict(color="#d62728", width=2),
            name="Over target tail",
            hovertemplate="OEW: %{x:,.1f} lbm<br>Density: %{y:.1f}% per 100 lbm"
            "<extra></extra>",
        )
    )
    line_specs = [
        ("Target", target, "#000000"),
        ("Adjusted Mean", mean, "#1f77b4"),
        ("P95", p95, "#6f2dbd"),
    ]
    y_top = float(y_values.max()) * 1.15
    for name, value, color in line_specs:
        figure.add_trace(
            go.Scatter(
                x=[value, value],
                y=[0, y_top],
                mode="lines",
                line=dict(color=color, width=2),
                name=name,
                hovertemplate=f"{name}: {value:,.1f} lbm<extra></extra>",
            )
        )

    figure.update_layout(
        height=520,
        xaxis_title="OEW (lbm)",
        yaxis_title="Probability density (% per 100 lbm)",
        margin=dict(l=10, r=10, t=20, b=10),
        legend=dict(orientation="h", yanchor="bottom", y=-0.25, xanchor="center", x=0.5),
    )
    st.subheader("OEW Risk")
    st.plotly_chart(figure, use_container_width=True)


def render_target_comparison(
    latest: pd.DataFrame,
    ata_targets: pd.DataFrame,
    aircraft_targets: pd.DataFrame,
) -> None:
    if ata_targets.empty and aircraft_targets.empty:
        return

    st.subheader("Target Comparison")
    if not ata_targets.empty:
        target_names = dict(zip(ata_targets["ata_label"], ata_targets["ata_name"]))
        actual = (
            latest.groupby(["ata_label", "ata_display"], dropna=False)
            .agg(actual_mass_lbm=("total_mass_kg", "sum"))
            .reset_index()
        )
        comparison = ata_targets.merge(actual, on="ata_label", how="left")
        comparison["ata_display"] = comparison.apply(
            lambda row: row["ata_display"]
            if isinstance(row.get("ata_display"), str)
            else ata_display_label(row["ata_label"], target_names),
            axis=1,
        )
        comparison["status_delta_lbm"] = (
            comparison["actual_mass_lbm"] - comparison["target_mass_lbm"]
        )
        comparison["status_color"] = comparison["status_delta_lbm"].map(
            lambda value: "#d62728" if value > 0 else "#2ca02c"
        )
        comparison = comparison.sort_values("ata_label", key=lambda series: series.astype(int))
        figure = go.Figure()
        figure.add_trace(
            go.Bar(
                x=comparison["ata_display"],
                y=comparison["status_delta_lbm"],
                marker_color=comparison["status_color"],
                name="Status",
                customdata=comparison[["target_mass_lbm", "actual_mass_lbm"]],
                hovertemplate="<b>%{x}</b><br>Status: %{y:+,.1f} lbm"
                "<br>Target: %{customdata[0]:,.1f} lbm"
                "<br>Status: %{customdata[1]:,.1f} lbm<extra></extra>",
            )
        )
        figure.update_layout(
            yaxis_title="Status - Target (lbm)",
            xaxis_title="ATA",
            height=520,
            margin=dict(l=10, r=10, t=20, b=10),
        )
        figure.add_hline(y=0, line_width=1, line_color="#666")
        st.plotly_chart(figure, use_container_width=True)

    if not aircraft_targets.empty:
        aircraft = aircraft_targets.dropna(subset=["target_mass_lbm"])
        aircraft = aircraft.copy()
        aircraft["status_delta_lbm"] = (
            aircraft["workbook_current_lbm"] - aircraft["target_mass_lbm"]
        )
        aircraft["status_color"] = aircraft["status_delta_lbm"].map(
            lambda value: "#d62728" if value > 0 else "#2ca02c"
        )
        figure = go.Figure()
        figure.add_trace(
            go.Bar(
                x=aircraft["code"],
                y=aircraft["status_delta_lbm"],
                marker_color=aircraft["status_color"],
                name="Status",
                text=aircraft["description"],
                customdata=aircraft[["target_mass_lbm", "workbook_current_lbm"]],
                hovertemplate="<b>%{x}</b><br>%{text}<br>Status: %{y:+,.1f} lbm"
                "<br>Target: %{customdata[0]:,.1f} lbm"
                "<br>Status: %{customdata[1]:,.1f} lbm<extra></extra>",
            )
        )
        figure.update_layout(
            yaxis_title="Status - Target (lbm)",
            xaxis_title="Aircraft condition",
            height=420,
            margin=dict(l=10, r=10, t=20, b=10),
        )
        figure.add_hline(y=0, line_width=1, line_color="#666")
        st.plotly_chart(figure, use_container_width=True)


def render_app() -> None:
    st.title("AstroM Mass Breakdown")

    exports, load_warnings = load_exports(str(EXPORT_DIR))
    ata_targets, aircraft_targets, target_warnings = load_targets(str(TARGET_WORKBOOK))
    risk_values, risk_warnings = load_risk(str(TARGET_WORKBOOK))
    for warning in load_warnings:
        st.warning(warning)
    for warning in target_warnings:
        st.warning(warning)
    for warning in risk_warnings:
        st.warning(warning)

    if exports.empty:
        st.error(f"No valid CSV exports found in {EXPORT_DIR.resolve()}.")
        st.stop()

    ata_names = ATA_NAME_FALLBACK.copy()
    if not ata_targets.empty:
        ata_names.update(dict(zip(ata_targets["ata_label"], ata_targets["ata_name"])))
    exports["ata_display"] = exports["ata_label"].map(
        lambda value: ata_display_label(value, ata_names)
    )
    exports["component_display"] = exports["ata_display"] + " | " + exports["component"]

    iterations = sorted(
        exports["iteration"].unique(),
        key=lambda value: version_sort_key(export_version_from_path(Path(value))),
    )
    selected_iterations = st.sidebar.multiselect(
        "Export iterations",
        iterations,
        default=iterations[-1:],
    )
    if not selected_iterations:
        st.info("Select at least one export iteration.")
        st.stop()

    top_n = st.sidebar.slider("Visible categories", 5, 40, 20)
    selected_ata = st.sidebar.multiselect(
        "ATA filter",
        sorted(
            exports["ata_label"].unique(),
            key=lambda value: int(value) if value.lstrip("-").isdigit() else 999,
        ),
    )

    filtered = exports[exports["iteration"].isin(selected_iterations)].copy()
    if selected_ata:
        filtered = filtered[filtered["ata_label"].isin(selected_ata)]

    latest_iteration = selected_iterations[-1]
    latest = filtered[filtered["iteration"] == latest_iteration]
    properties = mass_properties(oew_frame(latest))

    st.caption(
        f"Showing {len(filtered):,} rows from {len(selected_iterations)} export iteration(s)."
    )
    render_header_rows(latest, latest_iteration, properties, aircraft_targets)

    ata_summary = summarize_by(latest, "ata_label")
    ata_summary["ata_display"] = ata_summary["ata_label"].map(
        lambda value: ata_display_label(value, ata_names)
    )
    component_summary = summarize_by(latest, "component_display")

    st.subheader("ATA Breakdown")
    render_ata_pie(ata_summary)

    st.subheader("Component Breakdown")
    render_component_bar(component_summary, top_n)

    render_iteration_trend(exports, aircraft_targets)
    render_risk_plot(risk_values)
    render_target_comparison(latest, ata_targets, aircraft_targets)

    render_ata_tables(latest, ata_summary)


def main() -> None:
    st.set_page_config(page_title="AstroM Mass Breakdown", layout="wide")
    try:
        render_app()
    except Exception as exc:
        st.title("AstroM Mass Breakdown")
        st.error("The dashboard could not finish rendering, but the app process is still online.")
        st.caption("Check the details below, fix the data or code path, and redeploy.")
        st.exception(exc)


if __name__ == "__main__":
    main()
