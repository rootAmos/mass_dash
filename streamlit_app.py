from pathlib import Path

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


@st.cache_data(show_spinner=False)
def load_exports(export_dir: str) -> pd.DataFrame:
    rows = []
    for path in sorted(Path(export_dir).glob("*.csv")):
        frame = pd.read_csv(path)
        missing_columns = REQUIRED_COLUMNS.difference(frame.columns)
        if missing_columns:
            st.warning(f"Skipping {path.name}; missing columns: {', '.join(sorted(missing_columns))}")
            continue

        frame = frame.copy()
        frame["export_file"] = path.name
        frame["iteration"] = path.stem
        for column in NUMERIC_COLUMNS:
            if column in frame.columns:
                frame[column] = pd.to_numeric(frame[column], errors="coerce")
        rows.append(frame)

    if not rows:
        return pd.DataFrame()

    exports = pd.concat(rows, ignore_index=True)
    exports["ata_label"] = exports["ata"].fillna(-1).astype(int).astype(str)
    exports["component"] = exports["component"].fillna("Unspecified")
    exports["total_mass_kg"] = exports["total_mass_kg"].fillna(0.0)
    return exports


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


def main() -> None:
    st.set_page_config(page_title="AstroM Mass Breakdown", layout="wide")
    st.title("AstroM Mass Breakdown")

    exports = load_exports(str(EXPORT_DIR))
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

    group_mode = st.sidebar.radio("Breakdown", ["ATA", "Component"], horizontal=True)
    top_n = st.sidebar.slider("Visible categories", 5, 40, 20)
    selected_ata = st.sidebar.multiselect(
        "ATA filter",
        sorted(exports["ata_label"].unique(), key=lambda value: int(value) if value.lstrip("-").isdigit() else 999),
    )

    filtered = exports[exports["iteration"].isin(selected_iterations)].copy()
    if selected_ata:
        filtered = filtered[filtered["ata_label"].isin(selected_ata)]

    latest_iteration = selected_iterations[-1]
    latest = filtered[filtered["iteration"] == latest_iteration]
    properties = mass_properties(latest)

    st.caption(f"Showing {len(filtered):,} rows from {len(selected_iterations)} export iteration(s).")
    render_metric_row(properties)

    group_column = "ata_label" if group_mode == "ATA" else "component"
    summary = summarize_by(latest, group_column)
    visible_summary = summary.head(top_n)

    bar = px.bar(
        visible_summary.sort_values("total_mass_kg"),
        x="total_mass_kg",
        y=group_column,
        orientation="h",
        text="total_mass_kg",
        labels={
            "total_mass_kg": "Mass (kg)",
            group_column: group_mode,
            "share_pct": "Share (%)",
        },
        hover_data={"share_pct": ":.1f", "item_count": True},
    )
    bar.update_traces(texttemplate="%{text:,.0f} kg", textposition="outside")
    bar.update_layout(height=max(420, 26 * len(visible_summary)), margin=dict(l=10, r=10, t=20, b=10))
    st.plotly_chart(bar, use_container_width=True)

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

    st.subheader("Breakdown Table")
    st.dataframe(
        summary[[group_column, "total_mass_kg", "share_pct", "item_count"]]
        .rename(
            columns={
                group_column: group_mode,
                "total_mass_kg": "Mass (kg)",
                "share_pct": "Share (%)",
                "item_count": "Items",
            }
        )
        .style.format({"Mass (kg)": "{:,.1f}", "Share (%)": "{:.1f}"}),
        use_container_width=True,
        hide_index=True,
    )

    st.subheader("Source Rows")
    st.dataframe(
        latest.sort_values(["ata", "component"])[
            [
                "ata_label",
                "component",
                "qty",
                "unit_mass_kg",
                "total_mass_kg",
                "x_m",
                "y_m",
                "z_m",
                "notes",
            ]
        ],
        use_container_width=True,
        hide_index=True,
    )


if __name__ == "__main__":
    main()
