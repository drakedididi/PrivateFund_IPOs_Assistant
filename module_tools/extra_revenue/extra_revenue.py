from __future__ import annotations

import argparse
import os
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Optional

import matplotlib

matplotlib.use("Agg")

import matplotlib.dates as mdates
import matplotlib.font_manager as fm
import matplotlib.pyplot as plt
import pandas as pd
from matplotlib.gridspec import GridSpec
from matplotlib.ticker import PercentFormatter


INVALID_FILENAME_CHARS = r'<>:"/\|?*'
BUNDLED_FONT = Path(__file__).resolve().parent / "fonts" / "NotoSansSC-Regular.otf"


@dataclass
class DrawdownCycle:
    peak_date: pd.Timestamp
    trough_date: pd.Timestamp
    recovery_date: pd.Timestamp
    peak_value: float
    trough_value: float
    recovery_value: float
    cycle_weeks: float
    is_recovered: bool


@dataclass
class DrawdownResult:
    name: str
    wealth: pd.Series
    drawdown: pd.Series
    max_drawdown: float
    longest_cycle: Optional[DrawdownCycle]


@dataclass
class NavData:
    df: pd.DataFrame
    product_name: str
    benchmark_name: str
    frequency: str


def configure_matplotlib() -> None:
    plt.style.use("seaborn-v0_8-whitegrid")
    if BUNDLED_FONT.exists():
        fm.fontManager.addfont(str(BUNDLED_FONT))
        bundled_name = fm.FontProperties(fname=str(BUNDLED_FONT)).get_name()
    else:
        bundled_name = ""

    font_candidates = [
        bundled_name,
        "Noto Sans SC",
        "Microsoft YaHei",
        "SimHei",
        "Noto Sans CJK SC",
        "Source Han Sans CN",
        "Arial Unicode MS",
    ]
    available_fonts = {font.name for font in fm.fontManager.ttflist}
    for font_name in font_candidates:
        if font_name and font_name in available_fonts:
            plt.rcParams["font.sans-serif"] = [font_name]
            break
    plt.rcParams["axes.unicode_minus"] = False
    plt.rcParams["figure.dpi"] = 140
    plt.rcParams["savefig.dpi"] = 220


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text in {"nan", "None"}:
        return ""
    return text


def safe_filename_part(value: Any) -> str:
    text = clean_text(value)
    cleaned = "".join(ch for ch in text if ch not in INVALID_FILENAME_CHARS)
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" .")
    return cleaned or "未命名"


def has_non_ascii(value: Any) -> bool:
    return any(ord(char) > 127 for char in clean_text(value))


def chart_name(value: Any, fallback: str) -> str:
    text = clean_text(value)
    return text if text and not has_non_ascii(text) else fallback


def chart_frequency_label(value: str) -> str:
    return {"周频": "Weekly", "日频": "Daily"}.get(value, value or "Unknown")


def column_has_data(series: pd.Series) -> bool:
    for value in series:
        if value is None or pd.isna(value):
            continue
        if str(value).strip():
            return True
    return False


def validate_three_column_excel(raw: pd.DataFrame) -> tuple[str, str, str]:
    data_columns = [column for column in raw.columns if column_has_data(raw[column])]
    if len(data_columns) != 3:
        columns_text = "、".join(clean_text(column) or "<空列名>" for column in data_columns)
        raise ValueError(
            "仅支持三列标准数据类型：日期列、产品列、指数列。"
            f"当前检测到 {len(data_columns)} 列存在数据：{columns_text}"
        )

    date_col, product_col, benchmark_col = data_columns
    if not clean_text(product_col) or str(product_col).startswith("Unnamed"):
        raise ValueError("产品列必须有明确列名，并作为图片标题和 ZIP 文件名的一部分。")
    if not clean_text(benchmark_col) or str(benchmark_col).startswith("Unnamed"):
        raise ValueError("指数列必须有明确列名，并作为图片标题和 ZIP 文件名的一部分。")

    return date_col, product_col, benchmark_col


def detect_frequency(dates: pd.DatetimeIndex) -> str:
    if len(dates) < 2:
        raise ValueError("至少需要两行有效日期数据，才能判断日频或周频。")

    diffs = pd.Series(dates.sort_values()).diff().dropna().dt.days
    median_days = float(diffs.median())
    return "周频" if median_days >= 5 else "日频"


def load_nav_data(excel_path: Path, sheet_name: int | str = 0) -> NavData:
    raw = pd.read_excel(excel_path, sheet_name=sheet_name)
    date_col, product_col, benchmark_col = validate_three_column_excel(raw)

    df = raw[[date_col, product_col, benchmark_col]].copy()
    df.columns = ["date", "product_nav", "benchmark_nav"]
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df["product_nav"] = pd.to_numeric(df["product_nav"], errors="coerce")
    df["benchmark_nav"] = pd.to_numeric(df["benchmark_nav"], errors="coerce")
    df = df.dropna().sort_values("date").drop_duplicates(subset="date")

    if len(df) < 2:
        raise ValueError("清洗后有效数据少于两行，请检查 Excel 内容。")
    if (df["product_nav"] <= 0).any() or (df["benchmark_nav"] <= 0).any():
        raise ValueError("产品列和指数列仅支持正数净值或指数点位。")

    indexed = df.set_index("date")
    frequency = detect_frequency(pd.DatetimeIndex(indexed.index))
    return NavData(
        df=indexed,
        product_name=clean_text(product_col),
        benchmark_name=clean_text(benchmark_col),
        frequency=frequency,
    )


def compute_excess_navs(df: pd.DataFrame) -> pd.DataFrame:
    result = df.copy()
    result["product_nav"] = result["product_nav"] / result["product_nav"].iloc[0]
    result["benchmark_cum_nav"] = result["benchmark_nav"] / result["benchmark_nav"].iloc[0]

    result["product_return"] = result["product_nav"].pct_change().fillna(0)
    result["benchmark_return"] = result["benchmark_cum_nav"].pct_change().fillna(0)
    result["arithmetic_excess"] = result["product_return"] - result["benchmark_return"]

    result["excess_nav"] = 1 + result["arithmetic_excess"].cumsum()
    result["excess_nav2"] = result["product_nav"] / result["benchmark_cum_nav"]
    result["excess_nav3"] = (1 + result["arithmetic_excess"]).cumprod()
    return result


def weeks_between(start: pd.Timestamp, end: pd.Timestamp) -> float:
    return round(int((end - start).days) / 7, 2)


def format_week_count(weeks: Optional[float]) -> str:
    if weeks is None:
        return "暂无"
    if float(weeks).is_integer():
        return str(int(weeks))
    return f"{weeks:.2f}"


def find_drawdown_cycles(
    wealth: pd.Series,
    include_unrecovered: bool = True,
) -> list[DrawdownCycle]:
    wealth = wealth.dropna()
    if wealth.empty:
        return []

    cycles: list[DrawdownCycle] = []
    peak_date = wealth.index[0]
    peak_value = float(wealth.iloc[0])
    trough_date = peak_date
    trough_value = peak_value
    in_drawdown = False

    for date_value, value in wealth.iloc[1:].items():
        value = float(value)
        if not in_drawdown:
            if value < peak_value:
                in_drawdown = True
                trough_date = date_value
                trough_value = value
            else:
                peak_date = date_value
                peak_value = value
            continue

        if value < trough_value:
            trough_date = date_value
            trough_value = value

        if value >= peak_value:
            cycles.append(
                DrawdownCycle(
                    peak_date=peak_date,
                    trough_date=trough_date,
                    recovery_date=date_value,
                    peak_value=peak_value,
                    trough_value=trough_value,
                    recovery_value=value,
                    cycle_weeks=weeks_between(peak_date, date_value),
                    is_recovered=True,
                )
            )
            peak_date = date_value
            peak_value = value
            trough_date = date_value
            trough_value = value
            in_drawdown = False

    if include_unrecovered and in_drawdown:
        end_date = wealth.index[-1]
        cycles.append(
            DrawdownCycle(
                peak_date=peak_date,
                trough_date=trough_date,
                recovery_date=end_date,
                peak_value=peak_value,
                trough_value=trough_value,
                recovery_value=float(wealth.iloc[-1]),
                cycle_weeks=weeks_between(peak_date, end_date),
                is_recovered=False,
            )
        )

    return cycles


def analyze_drawdown(name: str, wealth: pd.Series) -> DrawdownResult:
    wealth = wealth.dropna()
    running_max = wealth.cummax()
    drawdown = wealth / running_max - 1
    cycles = find_drawdown_cycles(wealth, include_unrecovered=True)
    longest_cycle = max(cycles, key=lambda cycle: cycle.cycle_weeks) if cycles else None

    return DrawdownResult(
        name=name,
        wealth=wealth,
        drawdown=drawdown,
        max_drawdown=float(drawdown.min()),
        longest_cycle=longest_cycle,
    )


def style_axis(ax: plt.Axes, percent: bool = False) -> None:
    ax.grid(True, linestyle="--", alpha=0.35, linewidth=0.8)
    ax.set_facecolor("#fbfbfb")
    ax.xaxis.set_major_locator(mdates.AutoDateLocator(minticks=6, maxticks=10))
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y-%m"))
    ax.tick_params(axis="x", rotation=20)
    if percent:
        ax.yaxis.set_major_formatter(PercentFormatter(1))
    for spine in ax.spines.values():
        spine.set_alpha(0.3)


def save_figure_with_fallback(fig: plt.Figure, output_path: Path) -> Path:
    try:
        fig.savefig(output_path, bbox_inches="tight")
        return output_path
    except PermissionError:
        counter = 1
        while True:
            fallback_path = output_path.with_name(
                f"{output_path.stem}_{counter}{output_path.suffix}"
            )
            if not fallback_path.exists():
                fig.savefig(fallback_path, bbox_inches="tight")
                return fallback_path
            counter += 1


def resolve_annotation_position(
    ax: plt.Axes,
    date_value: pd.Timestamp,
    y_value: float,
    base_dx: int,
    base_dy: int,
    base_ha: str,
) -> tuple[tuple[int, int], str]:
    x_min, x_max = ax.get_xlim()
    x_value = mdates.date2num(date_value)
    dx = base_dx
    text_ha = base_ha
    if x_max > x_min:
        left_threshold = x_min + 0.08 * (x_max - x_min)
        right_threshold = x_min + 0.92 * (x_max - x_min)
        if x_value <= left_threshold and dx < 0:
            dx = 12
            text_ha = "left"
        elif x_value >= right_threshold and dx > 0:
            dx = -70
            text_ha = "right"

    y_min, y_max = ax.get_ylim()
    dy = base_dy
    if y_max > y_min:
        upper_threshold = y_min + 0.82 * (y_max - y_min)
        lower_threshold = y_min + 0.18 * (y_max - y_min)
        if y_value >= upper_threshold and dy > 0:
            dy = -abs(dy) - 8
        elif y_value <= lower_threshold and dy < 0:
            dy = abs(dy) + 8
    return (dx, dy), text_ha


def annotate_drawdown_event(ax: plt.Axes, result: DrawdownResult, color: str) -> None:
    cycle = result.longest_cycle
    if cycle is None:
        return

    ax.axvspan(cycle.peak_date, cycle.recovery_date, color=color, alpha=0.08)
    end_label = "New high" if cycle.is_recovered else "Period end"
    points = [
        ("Cycle start", cycle.peak_date, cycle.peak_value, -70, 18, "right"),
        ("Trough", cycle.trough_date, cycle.trough_value, 12, -38, "left"),
        (end_label, cycle.recovery_date, cycle.recovery_value, 12, 18, "left"),
    ]

    for label, date_value, y_value, dx, dy, ha in points:
        offset, text_ha = resolve_annotation_position(ax, date_value, y_value, dx, dy, ha)
        ax.scatter(date_value, y_value, color=color, s=36, zorder=5)
        ax.annotate(
            f"{label}\n{date_value:%Y-%m-%d}\n{y_value:.4f}",
            xy=(date_value, y_value),
            xytext=offset,
            textcoords="offset points",
            fontsize=8.5,
            color=color,
            ha=text_ha,
            va="bottom" if offset[1] >= 0 else "top",
            bbox=dict(boxstyle="round,pad=0.25", fc="white", ec=color, alpha=0.95),
            arrowprops=dict(arrowstyle="-", color=color, alpha=0.7, lw=1),
        )


def plot_nav_panel(ax: plt.Axes, result: DrawdownResult, color: str) -> None:
    cycle_weeks = (
        format_week_count(result.longest_cycle.cycle_weeks)
        if result.longest_cycle is not None
        else "N/A"
    )
    ax.plot(result.wealth.index, result.wealth, color=color, linewidth=2.1)
    ax.axhline(1, color="#777777", linewidth=0.9, alpha=0.75)
    ax.set_title(
        f"{result.name} | Max drawdown {result.max_drawdown:.2%} | Longest recovery {cycle_weeks} weeks",
        fontsize=11.5,
        pad=12,
    )
    ax.set_ylabel("Excess NAV")
    style_axis(ax)
    annotate_drawdown_event(ax, result, color)


def plot_drawdown_panel(ax: plt.Axes, result: DrawdownResult, color: str) -> None:
    ax.plot(result.drawdown.index, result.drawdown, color=color, linewidth=1.7)
    ax.axhline(0, color="#777777", linewidth=0.9, alpha=0.75)
    cycle = result.longest_cycle
    if cycle is not None:
        mask = (result.drawdown.index >= cycle.peak_date) & (
            result.drawdown.index <= cycle.recovery_date
        )
        ax.fill_between(
            result.drawdown.index,
            result.drawdown,
            0,
            where=mask,
            color=color,
            alpha=0.22,
        )
        ax.axvspan(cycle.peak_date, cycle.recovery_date, color=color, alpha=0.07)

    ax.set_title("Drawdown window", fontsize=12, pad=10)
    ax.set_ylabel("Drawdown")
    style_axis(ax, percent=True)


def build_drawdown_results(df: pd.DataFrame) -> list[tuple[str, DrawdownResult, str]]:
    specs = [
        ("算数超额", "Arithmetic excess NAV", "excess_nav", "#c47a00"),
        ("几何超额", "Geometric excess NAV", "excess_nav2", "#1f5fbf"),
        ("累加超额", "Compounded arithmetic excess NAV", "excess_nav3", "#0f8b6d"),
    ]
    return [
        (label, analyze_drawdown(name, df[column]), color)
        for label, name, column, color in specs
    ]


def summarize_max_recovery_period(results: list[tuple[str, DrawdownResult, str]]) -> str:
    candidates = [
        result
        for _, result, _ in results
        if result.longest_cycle is not None
    ]
    if not candidates:
        return "暂无"

    longest_label, longest_result, _ = max(
        results,
        key=lambda item: item[1].longest_cycle.cycle_weeks if item[1].longest_cycle is not None else -1,
    )
    cycle = longest_result.longest_cycle
    status = "已修复" if cycle.is_recovered else "未修复"
    return f"{format_week_count(cycle.cycle_weeks)} 周（{longest_label}，{status}）"


def summarize_recovery_periods(results: list[tuple[str, DrawdownResult, str]]) -> str:
    lines: list[str] = []
    for index, (label, result, _) in enumerate(results, start=1):
        if result.longest_cycle is None:
            weeks = "暂无"
        else:
            weeks = f"{format_week_count(result.longest_cycle.cycle_weeks)}周"
        lines.append(f"{index} {label}：{weeks}")
    return "\n".join(lines)


def plot_three_drawdown_analysis(
    df: pd.DataFrame,
    period_label: str,
    product_name: str,
    benchmark_name: str,
    output_path: Path,
) -> tuple[Path, str, str]:
    results = build_drawdown_results(df)

    fig = plt.figure(figsize=(18, 13.5))
    grid = GridSpec(
        nrows=len(results),
        ncols=2,
        figure=fig,
        width_ratios=[2.05, 1.1],
        hspace=0.44,
        wspace=0.18,
    )

    for row, (_, result, color) in enumerate(results):
        nav_ax = fig.add_subplot(grid[row, 0])
        dd_ax = fig.add_subplot(grid[row, 1], sharex=nav_ax)
        plot_nav_panel(nav_ax, result, color)
        plot_drawdown_panel(dd_ax, result, color)

    fig.suptitle(
        f"{chart_name(product_name, 'Product')} vs "
        f"{chart_name(benchmark_name, 'Benchmark')} "
        f"({chart_frequency_label(period_label)})",
        fontsize=20,
        y=0.975,
    )
    fig.subplots_adjust(left=0.06, right=0.98, bottom=0.06, top=0.93)
    saved_path = save_figure_with_fallback(fig, output_path)
    plt.close(fig)
    return saved_path, summarize_max_recovery_period(results), summarize_recovery_periods(results)


def analyze_extra_revenue_excel(
    excel_path: str | os.PathLike[str],
    output_dir: str | os.PathLike[str],
    sheet_name: int | str = 0,
) -> dict[str, Any]:
    configure_matplotlib()

    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)
    nav_data = load_nav_data(Path(excel_path), sheet_name=sheet_name)
    enriched = compute_excess_navs(nav_data.df)
    title = f"{nav_data.product_name} vs {nav_data.benchmark_name}"
    safe_title = safe_filename_part(title)
    chart_path, max_recovery_period, recovery_periods_text = plot_three_drawdown_analysis(
        enriched,
        nav_data.frequency,
        nav_data.product_name,
        nav_data.benchmark_name,
        output_path / f"{safe_title}.png",
    )

    return {
        "title": title,
        "download_stem": safe_title,
        "frequency": nav_data.frequency,
        "max_recovery_period": max_recovery_period,
        "recovery_periods_text": recovery_periods_text,
        "chart_path": str(chart_path),
    }


def find_excel_files(root_dir: str | os.PathLike[str]) -> list[Path]:
    excel_files: list[Path] = []
    for root, _, files in os.walk(root_dir):
        for file_name in files:
            if file_name.startswith("~$"):
                continue
            if file_name.lower().endswith((".xlsx", ".xls")):
                excel_files.append(Path(root) / file_name)
    return excel_files


def process_excel_files(
    input_dir: str | os.PathLike[str],
    output_dir: str | os.PathLike[str],
) -> dict[str, Any]:
    excel_files = find_excel_files(input_dir)
    if len(excel_files) != 1:
        raise ValueError(f"仅支持上传一个 Excel 文件，当前检测到 {len(excel_files)} 个。")
    return analyze_extra_revenue_excel(excel_files[0], output_dir)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="生成产品 vs 指数超额收益回撤图。")
    parser.add_argument("excel_file", help="包含日期列、产品列、指数列的 Excel 文件。")
    parser.add_argument("--sheet", default=0, help="Excel sheet 名称或索引，默认读取第一个 sheet。")
    parser.add_argument("--output-dir", default="output", help="结果输出目录，默认 output。")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    sheet = int(args.sheet) if str(args.sheet).isdigit() else args.sheet
    result = analyze_extra_revenue_excel(args.excel_file, args.output_dir, sheet_name=sheet)
    print(f"数据频率: {result['frequency']}")
    print(f"回撤修复最大周期: {result['max_recovery_period']}")
    print(f"输出图片: {result['chart_path']}")


if __name__ == "__main__":
    main()
