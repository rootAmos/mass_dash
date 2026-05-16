from pathlib import Path
import re

import pandas as pd
import plotly.express as px
import streamlit as st


EXPORT_DIR = Path("MassExports")
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
DISPLAY_COLUMNS = [
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


def export_version_from_path(path: Path) -> str:
    version_match = re.search(r"v\d+(?:\.\d+)+", path.stem)
    return version_match.group(0) if version_match else path.stem


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
        rows.append(frame)

    if not rows:
        return pd.DataFrame(), warnings

    exports = pd.concat(rows, ignore_index=True)
    for column in DISPLAY_COLUMNS:
        if column not in exports.columns:
            exports[column] = ""

    exports["ata_label"] = exports["ata"].fillna(-1).astype(int).astype(str)
    exports["component"] = exports["component"].fillna("Unspecified")
    exports["description"] = exports["description"].fillna("")
    exports["notes"] = exports["notes"].fillna("")
    exports["status"] = exports["status"].fillna("")
    exports["total_mass_kg"] = exports["total_mass_kg"].fillna(0.0)
    return exports, warnings


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


def render_metric_row(properties: dict[str, float]) -> None:
    columns = st.columns(4)
    columns[0].metric("Total mass", f"{properties['mass']:,.1f} kg")
    columns[1].metric("CG X", f"{properties['x']:.2f} m")
    columns[2].metric("CG Y", f"{properties['y']:.2f} m")
    columns[3].metric("CG Z", f"{properties['z']:.2f} m")


def render_version_row(latest: pd.DataFrame, latest_iteration: str) -> None:
    versions = sorted(
        str(value) for value in latest["export_version"].dropna().unique() if str(value)
    )
    statuses = sorted(str(value) for value in latest["status"].dropna().unique() if str(value))
    files = sorted(str(value) for value in latest["export_file"].dropna().unique() if str(value))

    columns = st.columns(3)
    columns[0].metric("Version", ", ".join(versions) if versions else latest_iteration)
    columns[1].metric("Status", ", ".join(statuses) if statuses else "Unspecified")
    columns[2].metric("Source file", files[0] if len(files) == 1 else f"{len(files)} files")


def render_ata_pie(summary: pd.DataFrame) -> None:
    chart = px.pie(
        summary,
        names="ata_label",
        values="total_mass_kg",
        hole=0.34,
        labels={
            "ata_label": "ATA",
            "total_mass_kg": "Mass (kg)",
            "share_pct": "Share (%)",
        },
        hover_data={"share_pct": ":.1f", "item_count": True},
    )
    chart.update_traces(textposition="inside", textinfo="label+percent")
    chart.update_layout(height=500, margin=dict(l=10, r=10, t=20, b=10), showlegend=True)
    st.plotly_chart(chart, use_container_width=True)


def render_component_bar(summary: pd.DataFrame, top_n: int) -> None:
    visible_summary = summary.head(top_n)
    chart = px.bar(
        visible_summary.sort_values("total_mass_kg"),
        x="total_mass_kg",
        y="component",
        orientation="h",
        text="total_mass_kg",
        labels={
            "total_mass_kg": "Mass (kg)",
            "component": "Component",
            "share_pct": "Share (%)",
        },
        hover_data={"share_pct": ":.1f", "item_count": True},
    )
    chart.update_traces(texttemplate="%{text:,.0f} kg", textposition="outside")
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
                "component": "Component",
                "description": "Description",
                "qty": "Qty",
                "unit_mass_kg": "Unit mass (kg)",
                "total_mass_kg": "Mass (kg)",
                "share_pct": "Share (%)",
                "x_m": "X (m)",
                "y_m": "Y (m)",
                "z_m": "Z (m)",
                "notes": "Notes",
                "export_version": "Version",
                "status": "Status",
            }
        )

        with st.expander(
            f"ATA {row.ata_label} - {row.total_mass_kg:,.1f} kg "
            f"({row.share_pct:.1f}%, {row.item_count:,} components)",
            expanded=True,
        ):
            st.dataframe(
                table.style.format(
                    {
                        "Qty": "{:,.0f}",
                        "Unit mass (kg)": "{:,.1f}",
                        "Mass (kg)": "{:,.1f}",
                        "Share (%)": "{:.1f}",
                        "X (m)": "{:.2f}",
                        "Y (m)": "{:.2f}",
                        "Z (m)": "{:.2f}",
                    },
                    na_rep="",
                ),
                use_container_width=True,
                hide_index=True,
            )


def render_app() -> None:
    st.title("AstroM Mass Breakdown")

    exports, load_warnings = load_exports(str(EXPORT_DIR))
    for warning in load_warnings:
        st.warning(warning)

    if exports.empty:
        st.error(f"No valid CSV exports found in {EXPORT_DIR.resolve()}.")
        st.stop()

    iterations = sorted(exports["iteration"].unique())
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
    properties = mass_properties(latest)

    st.caption(
        f"Showing {len(filtered):,} rows from {len(selected_iterations)} export iteration(s)."
    )
    render_version_row(latest, latest_iteration)
    render_metric_row(properties)

    ata_summary = summarize_by(latest, "ata_label")
    component_summary = summarize_by(latest, "component")

    st.subheader("ATA Breakdown")
    render_ata_pie(ata_summary)

    st.subheader("Component Breakdown")
    render_component_bar(component_summary, top_n)

    trend = (
        filtered.groupby("iteration", sort=False)
        .apply(lambda frame: pd.Series(mass_properties(frame)), include_groups=False)
        .reset_index()
    )
    trend_chart = px.line(
        trend,
        x="iteration",
        y="mass",
        markers=True,
        labels={"iteration": "Export iteration", "mass": "Total mass (kg)"},
    )
    st.subheader("Iteration Trend")
    st.plotly_chart(trend_chart, use_container_width=True)

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
