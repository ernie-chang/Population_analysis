import os
import re
import glob
from typing import List, Optional

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib

# Font settings for Chinese (prefer macOS fonts)
matplotlib.rcParams["font.sans-serif"] = ["Heiti TC", "PingFang TC", "STHeiti", "Noto Sans CJK TC", "Arial Unicode MS"]
matplotlib.rcParams["axes.unicode_minus"] = False


NUMERIC_COLUMNS_CANDIDATES: List[str] = [
    "主日",
    "兒童主日",
    "小排",
    "禱告",
    "晨興",
    "福音出訪",
    "家聚會出訪",
    "家聚會受訪",
    "召會生活",
    "新人主日",
    "新人家聚會受訪",
]


DATE_PATTERNS = [
    r"～(\d{4})年(\d{1,2})月(\d{1,2})日",
    r"-(\d{4})年(\d{1,2})月(\d{1,2})日",
    r"至(\d{4})年(\d{1,2})月(\d{1,2})日",
    r"到(\d{4})年(\d{1,2})月(\d{1,2})日",
    r"(\d{4})年(\d{1,2})月(\d{1,2})日",
]


def parse_week_end_date_from_filename(file_path: str) -> Optional[pd.Timestamp]:
    filename = os.path.basename(file_path)
    for pattern in DATE_PATTERNS:
        m = re.search(pattern, filename)
        if m:
            groups = m.groups()
            year, month, day = map(int, groups[-3:])
            return pd.Timestamp(year=year, month=month, day=day)
    return None


def _clean_table_headers(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [str(col).strip() for col in df.columns]
    if "大區" in df.columns:
        df = df[df["大區"] != "大區"]
    return df


def _coerce_numeric_columns(df: pd.DataFrame) -> pd.DataFrame:
    for column_name in NUMERIC_COLUMNS_CANDIDATES:
        if column_name in df.columns:
            df[column_name] = pd.to_numeric(df[column_name], errors="coerce").fillna(0)
    return df


def read_single_report(file_path: str) -> Optional[pd.DataFrame]:
    week_end_date = parse_week_end_date_from_filename(file_path)
    if week_end_date is None:
        print(f"⚠ 無法從檔名解析日期: {os.path.basename(file_path)}，已略過")
        return None

    dataframe: Optional[pd.DataFrame] = None
    ext = os.path.splitext(file_path)[1].lower()

    try:
        dataframe = pd.read_excel(file_path)
    except Exception:
        try:
            dataframe = pd.read_html(file_path, header=0)[0]
        except Exception:
            try:
                if ext == ".xlsx":
                    dataframe = pd.read_excel(file_path, engine="openpyxl")
                else:
                    dataframe = pd.read_excel(file_path, engine="xlrd")
            except Exception:
                dataframe = None

    if dataframe is None:
        print(f"⚠ 無法讀取報表: {file_path}")
        return None

    dataframe = _clean_table_headers(dataframe)

    if "大區" not in dataframe.columns:
        print(f"⚠ 報表缺少必要欄位 '大區': {file_path}，已略過")
        return None
    if "小區" not in dataframe.columns:
        dataframe["小區"] = "未分小區"
    if "會所" not in dataframe.columns:
        dataframe["會所"] = ""

    dataframe = _coerce_numeric_columns(dataframe)
    dataframe["週末日"] = week_end_date

    keep_columns = ["會所", "大區", "小區", "週末日"] + [
        col for col in NUMERIC_COLUMNS_CANDIDATES if col in dataframe.columns
    ]
    return dataframe[keep_columns]


def _is_summary_text(value: object) -> bool:
    if not isinstance(value, str):
        return False
    return any(keyword in value for keyword in ["總計", "合計", "小計", "總數", "合共", "總和"])


def _remove_summary_rows(df: pd.DataFrame) -> pd.DataFrame:
    id_cols = [col for col in ["會所", "大區", "小區"] if col in df.columns]
    if not id_cols:
        return df
    mask_summary = False
    for col in id_cols:
        mask_summary = mask_summary | df[col].apply(_is_summary_text)
    return df[~mask_summary].copy()


def aggregate_reports(reports_dir: str) -> pd.DataFrame:
    pattern = os.path.join(reports_dir, "*.xls*")
    file_paths = sorted(glob.glob(pattern))
    if not file_paths:
        raise RuntimeError(f"在資料夾中找不到報表檔案: {reports_dir}")

    combined: List[pd.DataFrame] = []
    processed_count = 0
    for path in file_paths:
        report_df = read_single_report(path)
        if report_df is not None:
            combined.append(report_df)
            processed_count += 1
    if not combined:
        raise RuntimeError("沒有任何可用的報表資料。")

    all_data = pd.concat(combined, ignore_index=True)

    # Remove summary rows like 總計/合計/小計 to avoid double counting
    all_data = _remove_summary_rows(all_data)

    # Normalize and sort
    all_data.sort_values("週末日", inplace=True)

    # Drop potential duplicates per week and id columns
    all_data = all_data.drop_duplicates(subset=[col for col in ["會所", "大區", "小區", "週末日"] if col in all_data.columns])

    unique_weeks = all_data["週末日"].dropna().unique()
    print(f"📦 已讀取 {processed_count}/{len(file_paths)} 份報表；週數: {len(unique_weeks)} ({', '.join(pd.Series(unique_weeks).dt.strftime('%Y/%m/%d'))})")

    return all_data


def build_region_timeseries(all_reports: pd.DataFrame, region_name: str) -> pd.DataFrame:
    if region_name == "總計":
        region_df = all_reports.copy()
    else:
        region_df = all_reports[all_reports["大區"] == region_name].copy()
    if region_df.empty:
        return pd.DataFrame()

    aggregation_columns = [col for col in NUMERIC_COLUMNS_CANDIDATES if col in region_df.columns]
    ts = region_df.groupby("週末日")[aggregation_columns].sum().sort_index()

    if "福音出訪" in ts.columns or "家聚會出訪" in ts.columns:
        gospel = ts["福音出訪"] if "福音出訪" in ts.columns else 0
        home = ts["家聚會出訪"] if "家聚會出訪" in ts.columns else 0
        ts["總出訪"] = gospel + home
    return ts


def _format_date_axis(ax, dates=None):
    if dates is not None:
        ax.set_xticks(pd.Index(dates))
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y/%m/%d"))
    plt.setp(ax.get_xticklabels(), rotation=45, ha="right", fontsize=11)
    ax.tick_params(axis="y", labelsize=11)
    ax.margins(y=0.15)
    ax.grid(True, alpha=0.3)


def _annotate_series(ax, x_index: pd.Index, y_series: pd.Series, fontsize: int = 12):
    for x, y in zip(x_index, y_series):
        ax.annotate(
            f"{int(y)}",
            (x, y),
            textcoords="offset points",
            xytext=(0, 10),
            ha="center",
            fontsize=fontsize,
            bbox=dict(boxstyle="round,pad=0.2", fc="white", ec="none", alpha=0.8),
            zorder=3,
            clip_on=False,
        )


def plot_attendance(region_name: str, ts: pd.DataFrame, output_dir: str) -> None:
    plt.figure(figsize=(10, 6))
    ax = plt.gca()

    columns_to_plot = [
        ("主日", "當周主日人數", "red", "-"),
        ("小排", "小排人數", "gold", "-"),
        ("晨興", "晨興人數", "green", "-"),
        ("召會生活", "召會生活", "#8e44ad", "-"),
    ]

    plotted_any = False
    for column_key, label_text, color, linestyle in columns_to_plot:
        if column_key in ts.columns:
            ax.plot(ts.index, ts[column_key], label=label_text, color=color, linestyle=linestyle, marker="o", markersize=5, linewidth=2)
            _annotate_series(ax, ts.index, ts[column_key], fontsize=12)
            plotted_any = True

    if not plotted_any:
        print(f"⚠ {region_name} 沒有可繪製的出席相關欄位")
        plt.close()
        return

    ax.set_title(f"{region_name} - 召會生活人數")
    ax.set_xlabel("日期")
    ax.set_ylabel("人數")
    ax.legend(loc="upper left")
    _format_date_axis(ax, dates=ts.index)

    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, f"{region_name}_attendance.png")
    plt.tight_layout()
    plt.savefig(output_path)
    plt.close()
    print(f"✅ 已輸出 {output_path}")


def plot_burden(region_name: str, ts: pd.DataFrame, output_dir: str) -> None:
    plt.figure(figsize=(10, 6))
    ax = plt.gca()

    plotted_any = False
    if "禱告" in ts.columns:
        ax.plot(ts.index, ts["禱告"], label="禱告人數", color="#00aaff", marker="o", markersize=5, linewidth=2)
        _annotate_series(ax, ts.index, ts["禱告"], fontsize=12)
        plotted_any = True
    if "總出訪" in ts.columns:
        ax.plot(ts.index, ts["總出訪"], label="總出訪人數", color="#0044aa", marker="o", markersize=5, linewidth=2)
        _annotate_series(ax, ts.index, ts["總出訪"], fontsize=12)
        plotted_any = True
    if "家聚會受訪" in ts.columns:
        ax.plot(ts.index, ts["家聚會受訪"], label="受訪人數", color="#66ccff", marker="o", markersize=5, linewidth=2)
        _annotate_series(ax, ts.index, ts["家聚會受訪"], fontsize=12)
        plotted_any = True

    if not plotted_any:
        print(f"⚠ {region_name} 沒有可繪製的負擔相關欄位")
        plt.close()
        return

    ax.set_title(f"{region_name} - 負擔領受程度")
    ax.set_xlabel("日期")
    ax.set_ylabel("人數")
    ax.legend(loc="upper left")
    _format_date_axis(ax, dates=ts.index)

    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, f"{region_name}_burden.png")
    plt.tight_layout()
    plt.savefig(output_path)
    plt.close()
    print(f"✅ 已輸出 {output_path}")


def plot_subdistrict_attendance(region_name: str, subdistrict_name: str, ts: pd.DataFrame, output_dir: str) -> None:
    plt.figure(figsize=(10, 6))
    ax = plt.gca()

    columns_to_plot = [
        ("主日", "當周主日人數", "red", "-"),
        ("小排", "小排人數", "gold", "-"),
        ("晨興", "晨興人數", "green", "-"),
        ("召會生活", "召會生活", "#8e44ad", "-"),
    ]

    plotted_any = False
    for column_key, label_text, color, linestyle in columns_to_plot:
        if column_key in ts.columns:
            ax.plot(ts.index, ts[column_key], label=label_text, color=color, linestyle=linestyle, marker="o", markersize=5, linewidth=2)
            _annotate_series(ax, ts.index, ts[column_key], fontsize=12)
            plotted_any = True

    if not plotted_any:
        print(f"⚠ {region_name} - {subdistrict_name} 沒有可繪製的出席相關欄位")
        plt.close()
        return

    ax.set_title(f"{region_name} - {subdistrict_name} - 召會生活人數")
    ax.set_xlabel("日期")
    ax.set_ylabel("人數")
    ax.legend(loc="upper left")
    _format_date_axis(ax, dates=ts.index)

    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, f"{region_name}_{subdistrict_name}_attendance.png")
    plt.tight_layout()
    plt.savefig(output_path)
    plt.close()
    print(f"✅ 已輸出 {output_path}")


def plot_subdistrict_burden(region_name: str, subdistrict_name: str, ts: pd.DataFrame, output_dir: str) -> None:
    plt.figure(figsize=(10, 6))
    ax = plt.gca()

    plotted_any = False
    if "禱告" in ts.columns:
        ax.plot(ts.index, ts["禱告"], label="禱告人數", color="#00aaff", marker="o", markersize=5, linewidth=2)
        _annotate_series(ax, ts.index, ts["禱告"], fontsize=12)
        plotted_any = True
    if "總出訪" in ts.columns:
        ax.plot(ts.index, ts["總出訪"], label="總出訪人數", color="#0044aa", marker="o", markersize=5, linewidth=2)
        _annotate_series(ax, ts.index, ts["總出訪"], fontsize=12)
        plotted_any = True
    if "家聚會受訪" in ts.columns:
        ax.plot(ts.index, ts["家聚會受訪"], label="受訪人數", color="#66ccff", marker="o", markersize=5, linewidth=2)
        _annotate_series(ax, ts.index, ts["家聚會受訪"], fontsize=12)
        plotted_any = True

    if not plotted_any:
        print(f"⚠ {region_name} - {subdistrict_name} 沒有可繪製的負擔相關欄位")
        plt.close()
        return

    ax.set_title(f"{region_name} - {subdistrict_name} - 負擔領受程度")
    ax.set_xlabel("日期")
    ax.set_ylabel("人數")
    ax.legend(loc="upper left")
    _format_date_axis(ax, dates=ts.index)

    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, f"{region_name}_{subdistrict_name}_burden.png")
    plt.tight_layout()
    plt.savefig(output_path)
    plt.close()
    print(f"✅ 已輸出 {output_path}")


def generate_region_charts(all_reports: pd.DataFrame, region_name: str, output_dir: str) -> None:
    ts = build_region_timeseries(all_reports, region_name)
    if ts.empty:
        print(f"⚠ 找不到 {region_name} 的資料，無法繪圖")
        return
    plot_attendance(region_name, ts, output_dir)
    plot_burden(region_name, ts, output_dir)

    region_df = all_reports if region_name == "總計" else all_reports[all_reports["大區"] == region_name]
    if region_name != "總計" and not region_df.empty and "小區" in region_df.columns:
        for subdistrict in sorted(region_df["小區"].dropna().unique()):
            sub_df = region_df[region_df["小區"] == subdistrict]
            sub_ts = sub_df.groupby("週末日")[NUMERIC_COLUMNS_CANDIDATES].sum().sort_index()
            if "福音出訪" in sub_ts.columns or "家聚會出訪" in sub_ts.columns:
                gospel = sub_ts["福音出訪"] if "福音出訪" in sub_ts.columns else 0
                home = sub_ts["家聚會出訪"] if "家聚會出訪" in sub_ts.columns else 0
                sub_ts["總出訪"] = gospel + home
            if not sub_ts.empty:
                plot_subdistrict_attendance(region_name, str(subdistrict), sub_ts, output_dir)
                plot_subdistrict_burden(region_name, str(subdistrict), sub_ts, output_dir)


if __name__ == "__main__":
    base_dir = os.path.dirname(__file__)
    reports_dir = os.path.join(base_dir, "date")
    static_charts_dir = os.path.abspath(os.path.join(base_dir, "..", "static", "charts"))

    df_reports = aggregate_reports(reports_dir)

    # Ensure '總計' charts are generated correctly
    generate_region_charts(df_reports, "總計", static_charts_dir)
    for region in df_reports["大區"].dropna().unique():
        generate_region_charts(df_reports, str(region), static_charts_dir)
