import os
import sys
import math
import argparse
import uuid
import tempfile
import traceback
import datetime
import typing
from typing import List, Dict, Tuple
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
import pandas as pd
import numpy as np
if getattr(sys, "frozen", False):
    os.environ.setdefault("MPLCONFIGDIR", os.path.join(tempfile.gettempdir(), "mplconfig"))
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import docx
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.shared import OxmlElement, qn
from docx.oxml import ns
import scipy
try:
    from scipy import stats as scipy_stats
except Exception:
    scipy_stats = None
"""
本脚本用于统计分析表格文件并生成结果文档。
分析包括: 每一列的分布, 可视化。
表格格式为xlsx,结果文档保存为word。
"""

plt.rcParams["font.sans-serif"] = ["SimHei"]
plt.rcParams["axes.unicode_minus"] = False

"""定义所需要的列的名称，要求文档必须有这几个表头，如果有其他需要的，可以在这里增加"""
CATEGORICAL_COLUMNS = [
    "学历",
    "学制",
    "师范生类别",
    "培养方式",
    "生源地",
    "学习形式",
    "困难群体",
    "毕业去向",
    "单位名称",
    "单位性质",
    "单位行业",
    "单位所在地",
    "工作职位类别",
]

def _set_cell_border(cell, **kwargs):
    """设置单元格边框"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement("w:tcBorders")
        tcPr.append(tcBorders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        edge_data = kwargs.get(edge)
        tag = "w:{}".format(edge)
        element = tcBorders.find(qn(tag))
        if edge_data is None:
            if element is not None:
                tcBorders.remove(element)
            continue
        if element is None:
            element = OxmlElement(tag)
            tcBorders.append(element)
        element.set(qn("w:val"), edge_data.get("val", "single"))
        if "sz" in edge_data:
            element.set(qn("w:sz"), str(edge_data.get("sz", 8)))
        if "space" in edge_data:
            element.set(qn("w:space"), str(edge_data.get("space", 0)))
        element.set(qn("w:color"), edge_data.get("color", "000000"))

def _apply_three_line_table(table):
    """应用三线表样式并自动调整"""
    # 先清除所有边框
    for row in table.rows:
        for cell in row.cells:
            _set_cell_border(
                cell,
                top={"val": "none"},
                bottom={"val": "none"},
                left={"val": "none"},
                right={"val": "none"},
                insideH={"val": "none"},
                insideV={"val": "none"},
            )
    
    # 设置顶线、中线和底线
    top_border = {"val": "single", "sz": 12, "color": "000000"}
    mid_border = {"val": "single", "sz": 8, "color": "000000"}
    bottom_border = {"val": "single", "sz": 12, "color": "000000"}
    
    # 设置表头行（第一行）的上下边框
    header_row = table.rows[0]
    for cell in header_row.cells:
        _set_cell_border(cell, top=top_border, bottom=mid_border)
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                r.bold = True
    
    # 设置最后一行的下边框 - 关键修复：确保每个单元格都有下边框
    last_row = table.rows[len(table.rows) - 1]
    for cell in last_row.cells:
        _set_cell_border(cell, bottom=bottom_border)
    
    # 额外处理：如果是多列，确保底线的连续性
    if len(last_row.cells) > 1:
        # 合并设置整个行的底边框
        tbl = table._tbl
        tblPr = tbl.tblPr
        tblBorders = tblPr.tblBorders if hasattr(tblPr, 'tblBorders') else None
        if tblBorders is not None:
            # 移除可能存在的 tblBorders 设置，改用单元格级别的边框
            tblPr.remove(tblBorders)
    
    # 自动调整表格：先按内容调整，再按窗口调整
    try:
        table.autofit = False
        # 先根据内容调整列宽
        for column in table.columns:
            for cell in column.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        if run.text.strip():
                            # 设置自动调整
                            cell.width = None
        # 再根据窗口调整
        table.autofit = True
        table.style = 'Table Grid'
        # 最后设置为无内部边框（保持三线表样式）
        for row in table.rows:
            for cell in row.cells:
                if row != header_row and row != last_row:
                    _set_cell_border(
                        cell,
                        top={"val": "none"},
                        bottom={"val": "none"},
                        left={"val": "none"},
                        right={"val": "none"},
                    )
    except Exception:
        # 如果自动调整失败，至少保证三线表样式已应用
        pass

def _default_log_dir() -> str:
    """默认的日志目录"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def _write_log(title: str, content: str) -> str:
    """写入日志文件"""
    log_dir = _default_log_dir()
    try:
        os.makedirs(log_dir, exist_ok=True)
    except Exception:
        log_dir = tempfile.gettempdir()
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    path = os.path.join(log_dir, f"stat_analysis_{ts}.log")
    try:
        with open(path, "w", encoding="utf-8") as f:
            f.write(title + "\n")
            f.write(content + "\n")
    except Exception:
        return ""
    return path

def _install_excepthook() -> None:
    """全局异常处理函数，用于记录未处理的异常"""
    def _hook(exc_type, exc, tb):
        content = "".join(traceback.format_exception(exc_type, exc, tb))
        _write_log("Unhandled exception", content)
    sys.excepthook = _hook

_install_excepthook()

def is_numeric_series(s: pd.Series) -> bool:
    """确认是数字信息还是文本"""
    try:
        return pd.api.types.is_numeric_dtype(s.dropna())
    except Exception:
        return False

def sanitize_filename(name: str) -> str:
    """确认文件名是否符合规范"""
    invalid = '<>:"/\\|?*'
    for ch in invalid:
        name = name.replace(ch, "_")
    name = name.strip()
    return name if name else f"untitled_{uuid.uuid4().hex[:8]}"

def ensure_dir(path: str) -> None:
    """确认目录是否存在，用于保存过程"""
    if not os.path.exists(path):
        os.makedirs(path, exist_ok=True)

def read_excel(path: str) -> pd.DataFrame:
    """读取excel文件，转换为DataFrame，数字列转换为数值类型"""
    """日期会造成数据过大无法绘图的问题，这里先不转换，全部按照字符串来读取"""
    df = pd.read_excel(path, dtype=object)
    df.columns = [str(c).replace("\u3000", "").strip() for c in df.columns]
    for col in df.columns:
        s = df[col]
        df[col] = s.map(lambda v: "" if pd.isna(v) else str(v))
    return df

def _pick_col(columns: List[str], *candidates: str) -> str:
    """从候选列中选择一个存在的列"""
    cols = [str(c).replace("\u3000", "").strip() for c in columns]
    for cand in candidates:
        if cand in cols:
            return cand
    for cand in candidates:
        for c in cols:
            if cand in c:
                return c
    return ""

def _normalize_text_series(s: pd.Series) -> pd.Series:
    """将文本序列转换为标准化格式"""
    out = s.astype(str).str.strip()
    out = out.replace({"": np.nan, "nan": np.nan, "None": np.nan})
    return out


def _find_education_column(columns: List[str]) -> str:
    """确认学历列"""
    for k in ["学历分布", "学历"]:
        if k in columns:
            return k
    for c in columns:
        if "学历" in str(c):
            return c
    return ""

def _classify_education(v: str) -> str:
    """将学历分类为博士、学术硕士、专业硕士、硕士、本科"""
    s = str(v).strip()
    if not s or s.lower() == "nan":
        return ""
    if "博士" in s or "博" in s:
        return "博士"
    if "硕士" in s or "硕" in s:
        if "学术" in s:
            return "学术硕士"
        if "专业" in s:
            return "专业硕士"
        return "硕士"
    if "本科" in s or "本" in s:
        return "本科"
    return ""

def _degree_order() -> List[str]:
    """学历分类的顺序"""
    return ["博士", "学术硕士", "专业硕士", "硕士", "本科"]

def _build_stats_by_education(df: pd.DataFrame, col: str, edu_col: str) -> pd.DataFrame:
    """根据_classify_education定义的学历分类统计"""
    """输出学历，类别，标签，人数，占比的dataframe，学历顺序按照_degree_order定义"""
    edu = df[edu_col].map(_classify_education)
    edu_order = _degree_order()
    rows = []
    for e in edu_order:
        sub = df.loc[edu == e, col]
        vc = _normalize_text_series(sub).dropna()
        total = int(len(vc))
        if total == 0:
            continue
        counts = vc.value_counts().sort_index()
        perc = (counts / total * 100).round(2)
        for k in counts.index.astype(str).tolist():
            rows.append(
                {
                    "学历": e,
                    "类别": k,
                    "标签": f"{e}|{k}",
                    "人数": int(counts.loc[k]),
                    "占比(%)": float(perc.loc[k]),
                }
            )
    if len(rows) == 0:
        return pd.DataFrame(columns=["学历", "类别", "标签", "人数", "占比(%)"])
    out = pd.DataFrame(rows)
    out["学历"] = pd.Categorical(out["学历"], categories=edu_order, ordered=True)
    out = out.sort_values(["学历", "类别"]).reset_index(drop=True)
    return out


def plot_counts_by_education(df_stats: pd.DataFrame, title: str, out_path: str) -> None:
    """根据_build_stats_by_education 输出的 dataframe，绘制学历分类的人数柱状图"""
    labels = df_stats["标签"].astype(str).tolist()
    counts = df_stats["人数"].astype(int).tolist()
    # 使用列名列表访问，避免编码问题
    percent_col = [c for c in df_stats.columns if '占' in c and '%' in c][0]
    percents = df_stats[percent_col].tolist()
    edu_values = df_stats["学历"].astype(str).tolist()
    n = len(labels)
    fig_w = max(10, 0.35 * n + 6)
    fig_h = 6
    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    color_map = {"博士": "#4C78A8", "学术硕士": "#F58518", "专业硕士": "#E45756", "硕士": "#FF9DA6", "本科": "#54A24B"}
    x = np.arange(len(labels))
    colors = [color_map.get(e, "#999999") for e in edu_values]
    bars = ax.bar(x, counts, width=0.5, color=colors)
    rotate = 45 if len(labels) >= 12 else 30
    tick_fs = 12
    value_fs = 12
    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=rotate, ha="right", fontsize=tick_fs)
    ax.set_ylabel("人数", fontsize=12)
    ax.set_title(title, fontsize=14, pad=20)
    ax.grid(axis="y", linestyle="--", alpha=0.25)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    for i, b in enumerate(bars):
        v = counts[i]
        p = percents[i]
        ax.text(b.get_x() + b.get_width() / 2, b.get_height(), f"{v}人\n{p}%", ha="center", va="bottom", fontsize=value_fs)
    handles = [
        plt.Line2D([0], [0], marker="s", color="w", markerfacecolor=color_map[k], markersize=9, label=k)
        for k in ["博士", "硕士", "本科", "学术硕士", "专业硕士"]
        if k in set(edu_values)
    ]
    if len(handles) > 0:
        ax.legend(handles=handles, ncol=min(3, len(handles)), fontsize=12, loc='upper right', bbox_to_anchor=(1, 1))
    fig.tight_layout(rect=[0, 0, 1, 0.95])
    fig.savefig(out_path, dpi=200, bbox_inches='tight')
    plt.close(fig)

def plot_source_employment_heatmap_by_level(df: pd.DataFrame, title: str, out_path: str) -> None:
    """绘制生源地和就业单位所在地的网格热力图（本科/硕士/博士分开）"""
    if len(df) == 0:
        return
    df0 = df.copy()
    df0["层次"] = df0["学历"].astype(str).map(_degree_to_level)
    df0 = df0.loc[df0["层次"].isin(["本科", "硕士", "博士"])]
    if len(df0) == 0:
        return
    
    # 按层次分组绘制
    for level in ["本科", "硕士", "博士"]:
        df_level = df0.loc[df0["层次"] == level]
        if len(df_level) == 0:
            continue
        
        # 构建交叉表，确保行列顺序一致
        pivot = pd.crosstab(df_level["单位所在地"], df_level["生源地"])
        if pivot.shape[0] == 0 or pivot.shape[1] == 0:
            continue
        
        # 获取行列标签并排序，确保顺序一致
        all_labels = sorted(set(list(pivot.index) + list(pivot.columns)))
        # 重新索引，使行列顺序一致
        pivot = pivot.reindex(index=all_labels, columns=all_labels, fill_value=0)
        
        fig_w = max(10, 0.5 * pivot.shape[1] + 2)
        fig_h = max(6, 0.5 * pivot.shape[0] + 2)
        fig, ax = plt.subplots(figsize=(fig_w, fig_h))
        
        # 绘制热力图
        im = ax.imshow(pivot.values, cmap='YlOrRd')
        
        # 设置坐标轴 - 所有字体统一为 12pt
        ax.set_xticks(np.arange(len(all_labels)))
        ax.set_yticks(np.arange(len(all_labels)))
        ax.set_xticklabels(all_labels, fontsize=12)
        ax.set_yticklabels(all_labels, fontsize=12)
        
        # 旋转 x 轴标签
        rotate = 45 if len(all_labels) >= 8 else 30
        plt.setp(ax.get_xticklabels(), rotation=rotate, ha="right", rotation_mode="anchor")
        
        # 设置标题 - 向上移动位置
        ax.set_title(f"{title}（{level}）", fontsize=14, pad=20)
        ax.set_xlabel("生源地", fontsize=12)
        ax.set_ylabel("就业单位所在地", fontsize=12)
        
        # 在每个格子中显示数值 - 字体 12pt
        for i in range(len(all_labels)):
            for j in range(len(all_labels)):
                val = pivot.values[i, j]
                text_color = "white" if val > pivot.values.max() * 0.6 else "black"
                ax.text(j, i, str(val), ha="center", va="center", color=text_color, fontsize=12)
        
        # 添加颜色条 - 调小
        cbar = ax.figure.colorbar(im, ax=ax, shrink=0.7)
        cbar.ax.set_ylabel('人数', rotation=-90, va="center", labelpad=20, fontsize=12)
        cbar.ax.tick_params(labelsize=12)
        
        # 调整布局，避免标题遮挡
        fig.tight_layout(rect=[0, 0, 1, 0.95])
        fig.savefig(out_path.replace('.png', f'_{level}.png'), dpi=200, bbox_inches='tight')
        plt.close(fig)

def plot_source_employment_heatmap_single(df: pd.DataFrame, title: str, out_path: str) -> None:
    """绘制单个层次的生源地和就业单位所在地网格热力图"""
    if len(df) == 0:
        return
    
    # 构建交叉表，确保行列顺序一致
    pivot = pd.crosstab(df["单位所在地"], df["生源地"])
    if pivot.shape[0] == 0 or pivot.shape[1] == 0:
        return
    
    # 获取行列标签并排序，确保顺序一致
    all_labels = sorted(set(list(pivot.index) + list(pivot.columns)))
    # 重新索引，使行列顺序一致
    pivot = pivot.reindex(index=all_labels, columns=all_labels, fill_value=0)
    
    fig_w = max(10, 0.5 * pivot.shape[1] + 2)
    fig_h = max(6, 0.5 * pivot.shape[0] + 2)
    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    
    # 创建自定义颜色映射：0 为灰色，其他用 colormap
    from matplotlib.colors import LinearSegmentedColormap
    colors = [(0.7, 0.7, 0.7)] + list(plt.cm.YlOrRd(np.linspace(0, 1, 256)))
    custom_cmap = LinearSegmentedColormap.from_list('custom_gray', colors, N=257)
    
    # 绘制热力图 - 使用自定义颜色映射
    im = ax.imshow(pivot.values, cmap=custom_cmap, vmin=0, vmax=max(pivot.values.max(), 1))
    
    # 设置坐标轴 - 所有字体统一为 15pt（调大 3 号）
    ax.set_xticks(np.arange(len(all_labels)))
    ax.set_yticks(np.arange(len(all_labels)))
    ax.set_xticklabels(all_labels, fontsize=15)
    ax.set_yticklabels(all_labels, fontsize=15)
    
    # 旋转 x 轴标签
    rotate = 45 if len(all_labels) >= 8 else 30
    plt.setp(ax.get_xticklabels(), rotation=rotate, ha="right", rotation_mode="anchor")
    
    # 设置标题 - 向上移动位置
    ax.set_title(title, fontsize=17, pad=20)
    ax.set_xlabel("生源地", fontsize=15)
    ax.set_ylabel("就业单位所在地", fontsize=15)
    
    # 在每个格子中显示数值 - 字体 15pt
    for i in range(len(all_labels)):
        for j in range(len(all_labels)):
            val = pivot.values[i, j]
            if val == 0:
                text_color = "black"
            else:
                text_color = "white" if val > pivot.values.max() * 0.6 else "black"
            ax.text(j, i, str(val), ha="center", va="center", color=text_color, fontsize=15)
    
    # 添加颜色条 - 调窄
    cbar = ax.figure.colorbar(im, ax=ax, shrink=0.6, aspect=20)
    cbar.ax.set_ylabel('人数', rotation=-90, va="center", labelpad=20, fontsize=15)
    cbar.ax.tick_params(labelsize=15)
    
    # 调整布局，避免标题遮挡
    fig.tight_layout(rect=[0, 0, 1, 0.95])
    fig.savefig(out_path, dpi=200, bbox_inches='tight')
    plt.close(fig)

def add_source_employment_heatmap(doc: Document, df: pd.DataFrame, img_dir: str, edu_col: str, source_col: str, loc_col: str) -> None:
    """添加生源地 - 就业单位所在地网格图到文档"""
    if (edu_col not in df.columns) or (source_col not in df.columns) or (loc_col not in df.columns):
        return
    
    # 准备数据
    df_temp = df.copy()
    df_temp[source_col] = df_temp[source_col].map(_map_location)
    df_temp[loc_col] = df_temp[loc_col].map(_map_location)
    edu = df_temp[edu_col].map(_classify_education)
    
    df_plot = pd.DataFrame({
        "学历": edu,
        "生源地": _normalize_text_series(df_temp[source_col]),
        "单位所在地": _normalize_text_series(df_temp[loc_col])
    })
    df_plot = df_plot.loc[(df_plot["生源地"] != "") & (df_plot["单位所在地"] != "")]
    
    if len(df_plot) == 0:
        return
    
    # 按层次分别添加标题和图片，与其他函数格式一致
    for level in ["本科", "硕士", "博士"]:
        df_level = df_plot.loc[df_plot["学历"].astype(str) == level]
        if len(df_level) == 0:
            continue
        
        title = f"生源地与就业单位所在地分布网格图（{level}）"
        add_heading(doc, title)
        
        # 保存临时路径
        temp_path = os.path.join(img_dir, sanitize_filename(f"生源地_就业单位_网格图_{level}"))
        plot_source_employment_heatmap_single(df_level, title, temp_path + '.png')
        
        # 添加到文档
        img_path = temp_path + '.png'
        if os.path.exists(img_path):
            add_picture(doc, img_path)

def plot_location_stacked_by_level(df_stats: pd.DataFrame, title: str, out_path: str) -> None:
    """把本硕博用不同花纹表示，绘制就业地点分布的柱状堆叠图"""
    if len(df_stats) == 0:
        return
    df0 = df_stats.copy()
    df0["层次"] = df0["学历"].astype(str).map(_degree_to_level)
    df0 = df0.loc[df0["层次"].isin(["本科", "硕士", "博士"])]
    if len(df0) == 0:
        return
    df0 = df0.groupby(["类别", "层次"], as_index=False)["人数"].sum()
    totals = df0.groupby("类别")["人数"].sum().sort_values(ascending=False)
    cats = totals.index.astype(str).tolist()
    if len(cats) == 0:
        return
    levels = ["本科", "硕士", "博士"]
    pivot = df0.pivot_table(index="类别", columns="层次", values="人数", aggfunc="sum", fill_value=0)
    palette = plt.cm.get_cmap("tab20")
    cat_colors: Dict[str, typing.Any] = {c: palette(i % 20) for i, c in enumerate(cats)}
    hatches = {"本科": "", "硕士": "///", "博士": "xx"}
    n = len(cats)
    fig_w = max(10, 0.45 * n + 6)
    fig_h = 6
    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    x = np.arange(n)
    totals_arr = np.zeros(n, dtype=float)
    for level in levels:
        heights = np.array([float(pivot.loc[c, level]) if (c in pivot.index and level in pivot.columns) else 0.0 for c in cats], dtype=float)
        bottoms = totals_arr.copy()
        colors = [cat_colors.get(c, "#999999") for c in cats]
        ax.bar(
            x,
            heights,
            bottom=bottoms,
            width=0.6,
            color=colors,
            edgecolor="black",
            linewidth=0.4,
            hatch=hatches.get(level, ""),
        )
        totals_arr = totals_arr + heights
    rotate = 45 if len(cats) >= 12 else 30
    tick_fs = 12
    value_fs = 12
    ax.set_xticks(x)
    ax.set_xticklabels(cats, rotation=rotate, ha="right", fontsize=tick_fs)
    ax.set_ylabel("人数", fontsize=12)
    ax.set_title(title, fontsize=14, pad=20)
    ax.grid(axis="y", linestyle="--", alpha=0.25)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    for i in range(n):
        v = int(totals_arr[i])
        ax.text(x[i], totals_arr[i], f"{v}人", ha="center", va="bottom", fontsize=value_fs)
    hatch_handles = []
    for level in levels:
        if (level in pivot.columns) and (int(pivot[level].sum()) > 0):
            hatch_handles.append(mpatches.Patch(facecolor="white", edgecolor="black", hatch=hatches.get(level, ""), label=level))
    if len(hatch_handles) > 0:
        ax.legend(handles=hatch_handles, ncol=min(3, len(hatch_handles)), fontsize=12, loc='upper right', bbox_to_anchor=(1, 1))
    fig.tight_layout(rect=[0, 0, 1, 0.95])
    fig.savefig(out_path, dpi=200, bbox_inches='tight')
    plt.close(fig)

def plot_grouped_stacked_binary_by_level(df_long: pd.DataFrame, title: str, out_path: str, level_order: List[str], group_order: List[str]) -> None:
    """"""
    if len(df_long) == 0:
        return
    need_cols = {"层次", "类别", "细分", "人数"}
    if not need_cols.issubset(set(df_long.columns)):
        return
    segs_all = sorted(df_long["细分"].astype(str).replace({"": np.nan, "nan": np.nan}).dropna().unique().tolist())
    if len(segs_all) == 0:
        return
    palette = plt.cm.get_cmap("tab20")
    seg_colors: Dict[str, typing.Any] = {seg: palette(i % 20) for i, seg in enumerate(segs_all)}
    within_gap = 0.75
    level_gap = 2.4
    bar_width = 0.55
    x_positions: List[float] = []
    x_labels: List[str] = []
    for i, lv in enumerate(level_order):
        base = float(i) * level_gap
        for j, g in enumerate(group_order):
            x_positions.append(base + float(j) * within_gap)
            x_labels.append(f"{lv}\n{g}")
    x = np.array(x_positions, dtype=float)
    bottoms = np.zeros(len(x_positions), dtype=float)
    fig_w = max(10, 0.9 * len(x_positions) + 6)
    fig_h = 6
    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    for seg in segs_all:
        heights: List[float] = []
        for lv in level_order:
            for g in group_order:
                v = df_long.loc[
                    (df_long["层次"].astype(str) == lv)
                    & (df_long["类别"].astype(str) == g)
                    & (df_long["细分"].astype(str) == seg),
                    "人数",
                ]
                heights.append(float(v.sum()) if len(v) else 0.0)
        h = np.array(heights, dtype=float)
        if h.sum() == 0:
            continue
        ax.bar(x, h, bottom=bottoms, width=bar_width, label=seg, color=seg_colors.get(seg, "#999999"), edgecolor="white", linewidth=0.6)
        bottoms = bottoms + h
    for i, lv in enumerate(level_order):
        idx0 = i * len(group_order)
        idx1 = idx0 + len(group_order)
        total_lv = float(bottoms[idx0:idx1].sum())
        if total_lv <= 0:
            continue
        for k in range(idx0, idx1):
            n = int(round(bottoms[k]))
            p = float(n / total_lv * 100)
            ax.text(x[k], bottoms[k], f"{n}人\n{p:.2f}%", ha="center", va="bottom", fontsize=12)
    ax.set_xticks(x)
    ax.set_xticklabels(x_labels, fontsize=12)
    ax.set_ylabel("人数", fontsize=12)
    ax.set_title(title, fontsize=14, pad=20)
    ax.grid(axis="y", linestyle="--", alpha=0.25)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.legend(ncol=min(3, len(segs_all)), fontsize=12, loc='upper right', bbox_to_anchor=(1, 1))
    fig.tight_layout(rect=[0, 0, 1, 0.95])
    fig.savefig(out_path, dpi=200, bbox_inches='tight')
    plt.close(fig)

def add_binary_stats_table_by_level(doc: Document, df_long: pd.DataFrame, level_order: List[str], group_order: List[str]) -> bool:
    """添加统计表格到文档"""
    if len(df_long) == 0:
        return False
    need_cols = {"层次", "类别", "细分", "人数", "占比(%)"}
    if not need_cols.issubset(set(df_long.columns)):
        return False
    levels_present = set(df_long["层次"].astype(str).unique().tolist())
    if not all(lv in levels_present for lv in level_order):
        return False
    segs_all = sorted(df_long["细分"].astype(str).replace({"": np.nan, "nan": np.nan}).dropna().unique().tolist())
    if len(segs_all) == 0:
        return False
    rows_list: List[Tuple[str, str]] = []
    for g in group_order:
        for seg in segs_all:
            rows_list.append((str(g), str(seg)))
    table = doc.add_table(rows=len(rows_list) + 1, cols=len(level_order) + 2)
    table.cell(0, 0).text = "类别"
    table.cell(0, 1).text = "细分"
    for j, lv in enumerate(level_order):
        table.cell(0, j + 2).text = lv
    for i, (g, seg) in enumerate(rows_list):
        table.cell(i + 1, 0).text = g
        table.cell(i + 1, 1).text = seg
        for j, lv in enumerate(level_order):
            sub = df_long.loc[
                (df_long["层次"].astype(str) == lv)
                & (df_long["类别"].astype(str) == g)
                & (df_long["细分"].astype(str) == seg)
            ]
            if len(sub) == 0:
                table.cell(i + 1, j + 2).text = "0人, 0%"
            else:
                n = int(sub["人数"].astype(float).sum())
                denom = float(max(1.0, float(sub["人数"].astype(float).sum())))
                p = float((sub["占比(%)"].astype(float) * sub["人数"].astype(float)).sum() / denom)
                table.cell(i + 1, j + 2).text = f"{n}人, {p:.2f}%"
    i = 0
    while i < len(rows_list):
        gname = rows_list[i][0]
        j = i
        while j + 1 < len(rows_list) and rows_list[j + 1][0] == gname:
            j += 1
        if j > i:
            table.cell(i + 1, 0).merge(table.cell(j + 1, 0))
        i = j + 1
    _apply_three_line_table(table)
    return True

def add_stats_text_by_education_custom(doc: Document, df_stats: pd.DataFrame, first_col_title: str) -> None:
    edu_order = [d for d in _degree_order() if d in set(df_stats["学历"].astype(str).unique().tolist())]
    row_labels = [str(row["类别"]) for _, row in df_stats[["类别"]].drop_duplicates().iterrows()]
    table = doc.add_table(rows=len(row_labels) + 1, cols=len(edu_order) + 1)
    table.cell(0, 0).text = first_col_title
    for j, e in enumerate(edu_order):
        table.cell(0, j + 1).text = e
    for i, cat in enumerate(row_labels):
        table.cell(i + 1, 0).text = cat
        for j, e in enumerate(edu_order):
            sub = df_stats.loc[(df_stats["学历"].astype(str) == e) & (df_stats["类别"].astype(str) == cat)]
            if len(sub) == 0:
                table.cell(i + 1, j + 1).text = "0人, 0%"
            else:
                r = sub.iloc[0]
                table.cell(i + 1, j + 1).text = f"{int(r['人数'])}人, {float(r['占比(%)']):.2f}%"
    _apply_three_line_table(table)

def _map_outcome(v: str) -> Tuple[str, str]:
    """将就业状态映射为求职中，升学和就业三个大类，每个大类里面包含小类别，将每个就业状态记录为(大类, 小类别)，例如（"升学", "留学"）"""
    s = str(v).strip()
    if not s or s.lower() == "nan":
        return "", ""
    if "求职中" in s:
        return "求职中", "求职中"
    if ("留学" in s) or ("研究生" in s):
        if "留学" in s:
            return "升学", "留学"
        return "升学", "研究生"
    return "就业", s


def _map_unit_nature(v: str) -> Tuple[str, str]:
    """将单位性质映射为是否科研，每个大类里面包含小类别，将每个单位性质记录为(大类, 小类别)，例如（"科研", "高等教育单位"）"""
    s = str(v).strip()
    if not s or s.lower() == "nan":
        return "", ""
    if "求职中" in s:
        return "求职中", "求职中"
    if ("高等教育单位" in s) or ("科研设计单位" in s):
        return "科研", s
    return "非科研", s

def _map_unit_nature_master_edu(v: str) -> Tuple[str, str]:
    """将硕士的就业映射为是否为教育单位，每个大类里面包含小类别，将每个单位性质记录为(大类, 小类别)，例如（"教育单位", "高等教育单位"）"""
    s = str(v).strip()
    if not s or s.lower() == "nan":
        return "", ""
    if ("高等教育单位" in s) or ("中初教育单位" in s):
        return "教育单位", s
    return "非教育单位", s


def _map_unit_nature_phd_tendency(v: str) -> Tuple[str, str]:
    """将单位性质映射为是否有科研倾向，每个大类里面包含小类别，将每个单位性质记录为(大类, 小类别)，例如（"有博士倾向", "高等教育单位"）"""
    s = str(v).strip()
    if not s or s.lower() == "nan":
        return "", ""
    if "求职中" in s:
        return "求职中", "求职中"
    if "高等教育单位" in s:
        return "科研", "高等教育单位"
    if "科研设计单位" in s:
        return "科研", "科研设计单位"
    return "其他行业", s

def _map_location(v: str) -> str:
    """将省份映射为省份前2个字符，例如（"北京市海淀区" -> "北京"），避免省份名称过长"""
    s = str(v).strip()
    if not s or s.lower() == "nan":
        return ""
    if "国外" in s or "境外" in s:
        return "国外"
    return s[:2]

def _collapse_top_provs(v: str) -> str:
    """将省份映射为北京、上海、广东、国外和其他五个大类"""
    s = _map_location(v)
    if s.startswith("北京"):
        return "北京"
    if s.startswith("上海"):
        return "上海"
    if s.startswith("广东"):
        return "广东"
    if s == "国外":
        return "国外"
    if s == "":
        return ""
    return "其他"

def _map_source_bsgs(v: str) -> str:
    """将生源映射为是否上北上广深，其他"""
    s = str(v).strip()
    if not s or s.lower() == "nan":
        return ""
    for k in ["北京", "上海", "深圳", "广州"]:
        if k in s:
            return "北上广深"
    return "其他"

def _chi_square_table(df: pd.DataFrame, y_col: str, x_col: str) -> Tuple[float, float, int, int]:
    """执行卡方检验，返回卡方统计量、p值、自由度和样本量"""
    if scipy_stats is None:
        raise RuntimeError("缺少scipy，无法执行卡方检验")
    ct = pd.crosstab(df[y_col], df[x_col])
    if ct.shape[0] < 2 or ct.shape[1] < 2:
        return float("nan"), float("nan"), 0, int(ct.values.sum())
    chi2, p, dof, _ = scipy_stats.chi2_contingency(ct.values)
    return float(chi2), float(p), int(dof), int(ct.values.sum())

def process_chi_square_research_influences(df: pd.DataFrame, doc: Document, img_dir: str) -> None:
    """对所选出的列执行卡方检验，绘制柱状图，表格，并写入word文档中"""
    edu_col = _pick_col(list(df.columns), "学历")
    unit_col = _pick_col(list(df.columns), "单位性质")
    if not edu_col or not unit_col:
        return
    sex_col = _pick_col(list(df.columns), "性别")
    train_col = _pick_col(list(df.columns), "培养方式")
    normal_col = _pick_col(list(df.columns), "师范生类别")
    unit_loc_col = _pick_col(list(df.columns), "单位所在地")
    source_col = _pick_col(list(df.columns), "生源所在地", "生源地")
    if scipy_stats is None:
        add_heading(doc, "卡方检验（科研/非科研 与 影响因素）")
        doc.add_paragraph("缺少scipy，无法执行卡方检验。")
        return
    base = df.copy()
    edu = base[edu_col].map(_classify_education)
    base["_层次"] = edu.astype(str).map(_degree_to_level)
    nature_big = base[unit_col].map(lambda v: _map_unit_nature(v)[0])
    base["_科研大类"] = nature_big
    base = base.loc[base["_科研大类"].isin(["科研", "非科研"])]
    if unit_loc_col and source_col:
        base["_单位前三"] = base[unit_loc_col].map(_map_location)
        base["_生源前三"] = base[source_col].map(_map_location)
        valid_loc = (base["_单位前三"] != "") & (base["_生源前三"] != "")
        base["_同源"] = np.where(valid_loc & (base["_单位前三"] == base["_生源前三"]), "同源", np.where(valid_loc, "非同源", ""))
    else:
        base["_同源"] = ""
    if source_col:
        base["_生源大类"] = base[source_col].map(_map_source_bsgs)
    else:
        base["_生源大类"] = ""
    factors = []
    if sex_col:
        factors.append(("性别", sex_col))
    if train_col:
        factors.append(("培养方式", train_col))
    if normal_col:
        factors.append(("师范生类别", normal_col))
    if unit_loc_col and source_col:
        factors.append(("同源/非同源", "_同源"))
    if source_col:
        factors.append(("生源地大类(北上广深/其他)", "_生源大类"))
    if len(factors) == 0:
        return
    add_heading(doc, "卡方检验（科研/非科研 × 影响因素，本/硕/博分开，按影响因素分别绘图）")
    levels = ["本科", "硕士", "博士"]
    for display, colname in factors:
        add_heading(doc, f"影响因素：{display}")
        res_rows = []
        level_cat_counts: Dict[str, Dict[str, Tuple[int, int, int]]] = {}
        for level in levels:
            part = base.loc[base["_层次"] == level].copy()
            if len(part) == 0:
                res_rows.append({"影响因素": display, "层次": level, "chi2": float("nan"), "dof": 0, "p": float("nan"), "显著": "否", "n": 0})
                level_cat_counts[level] = {}
                continue
            sub = part[["_科研大类", colname]].copy()
            sub[colname] = _normalize_text_series(sub[colname]).fillna("")
            sub = sub.loc[sub[colname] != ""]
            if len(sub) == 0:
                res_rows.append({"影响因素": display, "层次": level, "chi2": float("nan"), "dof": 0, "p": float("nan"), "显著": "否", "n": 0})
                level_cat_counts[level] = {}
                continue
            chi2, p, dof, n = _chi_square_table(sub, "_科研大类", colname)
            sig = "是" if (not math.isnan(p) and p < 0.05) else "否"
            res_rows.append({"影响因素": display, "层次": level, "chi2": chi2, "dof": dof, "p": p, "显著": sig, "n": n})

            ct = pd.crosstab(sub[colname], sub["_科研大类"])
            if "科研" not in ct.columns:
                ct["科研"] = 0
            if "非科研" not in ct.columns:
                ct["非科研"] = 0
            ct["合计"] = ct["科研"] + ct["非科研"]
            ct = ct.loc[ct["合计"] > 0]
            cat_map: Dict[str, Tuple[int, int, int]] = {}
            for idx, row in ct.iterrows():
                cat = str(idx)
                r = int(row.get("科研", 0))
                nr = int(row.get("非科研", 0))
                tot = int(row.get("合计", 0))
                cat_map[cat] = (r, nr, tot)
            level_cat_counts[level] = cat_map

        cats_all: Dict[str, int] = {}
        for _, cat_map in level_cat_counts.items():
            for cat, (_, _, tot) in cat_map.items():
                cats_all[cat] = cats_all.get(cat, 0) + int(tot)
        cats = [k for k, _ in sorted(cats_all.items(), key=lambda kv: (-kv[1], kv[0]))]
        if len(cats) > 0:
            positions: List[float] = []
            labels: List[str] = []
            research_pct_all: List[float] = []
            non_pct_all: List[float] = []
            totals_all: List[int] = []
            intra = 0.75
            inter = 1.6
            bar_width = 0.55
            base_x = 0.0
            for level in levels:
                cat_map = level_cat_counts.get(level, {})
                for i, cat in enumerate(cats):
                    r, nr, tot = cat_map.get(cat, (0, 0, 0))
                    positions.append(base_x + i * intra)
                    labels.append(f"{level}\n{cat}")
                    totals_all.append(int(tot))
                    if tot <= 0:
                        research_pct_all.append(0.0)
                        non_pct_all.append(0.0)
                    else:
                        research_pct_all.append(r / tot * 100)
                        non_pct_all.append(nr / tot * 100)
                base_x = base_x + len(cats) * intra + inter

            fig_w = max(10, min(32, 0.55 * len(positions) + 4))
            fig, ax = plt.subplots(figsize=(fig_w, 5.5))
            x = np.array(positions, dtype=float)
            ax.bar(x, research_pct_all, width=bar_width, label="科研", color="#4C78A8", edgecolor="white", linewidth=0.6)
            ax.bar(x, non_pct_all, width=bar_width, bottom=research_pct_all, label="非科研", color="#54A24B", edgecolor="white", linewidth=0.6)
            ax.set_ylim(0, 110)
            ax.set_ylabel("占比 (%)", fontsize=12)
            ax.set_title(f"科研/非科研占比（{display}，本/硕/博同图）", fontsize=14, pad=20)
            ax.grid(axis="y", linestyle="--", alpha=0.25)
            ax.spines["top"].set_visible(False)
            ax.spines["right"].set_visible(False)
            ax.legend(ncol=2, fontsize=12, loc='upper right', bbox_to_anchor=(1, 1))
            ax.set_xticks(x)
            rotate = 45 if len(labels) >= 12 else 0
            ax.set_xticklabels(labels, rotation=rotate, ha="right" if rotate else "center", fontsize=12)
            for i, n_cat in enumerate(totals_all):
                if n_cat <= 0:
                    continue
                ax.text(x[i], 102, f"n={n_cat}", ha="center", va="bottom", fontsize=12)
                rp = float(research_pct_all[i])
                npct = float(non_pct_all[i])
                if rp >= 6:
                    ax.text(x[i], rp / 2, f"{rp:.1f}%", ha="center", va="center", fontsize=12, color="white")
                if npct >= 6:
                    ax.text(x[i], rp + npct / 2, f"{npct:.1f}%", ha="center", va="center", fontsize=12, color="white")
            fig.tight_layout(rect=[0, 0, 1, 0.95])
            img_path = os.path.join(img_dir, sanitize_filename(f"卡方检验_{display}_本硕博同图.png"))
            fig.savefig(img_path, dpi=200, bbox_inches='tight')
            plt.close(fig)
            add_picture(doc, img_path)

        df_res = pd.DataFrame(res_rows)
        table = doc.add_table(rows=len(df_res) + 1, cols=7)
        table.cell(0, 0).text = "影响因素"
        table.cell(0, 1).text = "层次"
        table.cell(0, 2).text = "chi2"
        table.cell(0, 3).text = "dof"
        table.cell(0, 4).text = "p"
        table.cell(0, 5).text = "显著(p<0.05)"
        table.cell(0, 6).text = "n"
        for i, r in enumerate(df_res.to_dict("records")):
            table.cell(i + 1, 0).text = str(r["影响因素"])
            table.cell(i + 1, 1).text = str(r["层次"])
            table.cell(i + 1, 2).text = "" if math.isnan(r["chi2"]) else f"{r['chi2']:.4f}"
            table.cell(i + 1, 3).text = str(int(r["dof"]))
            table.cell(i + 1, 4).text = "" if math.isnan(r["p"]) else f"{r['p']:.6f}"
            table.cell(i + 1, 5).text = str(r["显著"])
            table.cell(i + 1, 6).text = str(int(r["n"]))
        if len(df_res) >= 2:
            table.cell(1, 0).merge(table.cell(len(df_res), 0))
        _apply_three_line_table(table)

def _build_stats_by_degree_and_mapper(df: pd.DataFrame, src_col: str, edu_col: str, mapper) -> pd.DataFrame:
    """根据学历和单位性质（或其他分类）构建统计数据,这里写一个统一的函数，便于调用（包括学历的分类和地区的分类）"""
    edu = df[edu_col].map(_classify_education)
    edu_order = _degree_order()
    rows = []
    for e in edu_order:
        sub_df = df.loc[edu == e, [src_col]]
        mapped = sub_df[src_col].map(lambda x: mapper(x))
        big = []
        small = []
        for x in mapped:
            if isinstance(x, tuple) and len(x) == 2:
                big.append(x[0])
                small.append(x[1])
            else:
                big.append("")
                small.append("")
        tmp = pd.DataFrame({"大类": big, "小类": small})
        tmp = tmp.loc[(tmp["大类"] != "") & (tmp["小类"] != "")]
        if len(tmp) == 0:
            continue
        tmp["组合"] = tmp["大类"].astype(str) + "|" + tmp["小类"].astype(str)
        counts = tmp["组合"].value_counts().sort_index()
        total = int(counts.sum())
        perc = (counts / total * 100).round(2)
        for k in counts.index.astype(str).tolist():
            bb, ss = k.split("|", 1)
            rows.append({"学历": e, "大类": bb, "小类": ss, "标签": f"{e}|{bb}|{ss}", "人数": int(counts.loc[k]), "占比(%)": float(perc.loc[k])})
    if len(rows) == 0:
        return pd.DataFrame(columns=["学历", "大类", "小类", "标签", "人数", "占比(%)"])
    out = pd.DataFrame(rows)
    out["学历"] = pd.Categorical(out["学历"], categories=edu_order, ordered=True)
    out = out.sort_values(["学历", "大类", "小类"]).reset_index(drop=True)
    return out

def add_table_by_degree_groups(doc: Document, df_stats: pd.DataFrame, row_prefix: str) -> None:
    """绘表函数"""
    edu_order = [d for d in _degree_order() if d in set(df_stats["学历"].astype(str).unique().tolist())]
    if "大类" in df_stats.columns and "小类" in df_stats.columns:
        big_col, small_col = "大类", "小类"
    else:
        big_col, small_col = "是否科研", "具体行业"
    pairs = df_stats[[big_col, small_col]].drop_duplicates()
    pairs = pairs.sort_values([big_col, small_col], kind="stable")
    rows_list = [(str(r[big_col]), str(r[small_col])) for _, r in pairs.iterrows()]
    table = doc.add_table(rows=len(rows_list) + 1, cols=len(edu_order) + 2)
    table.cell(0, 0).text = "类别"
    table.cell(0, 1).text = "细分"
    for j, e in enumerate(edu_order):
        table.cell(0, j + 2).text = e
    for i, (bb, ss) in enumerate(rows_list):
        table.cell(i + 1, 0).text = bb
        table.cell(i + 1, 1).text = ss
        for j, e in enumerate(edu_order):
            sub = df_stats.loc[
                (df_stats["学历"].astype(str) == e)
                & (df_stats[big_col].astype(str) == bb)
                & (df_stats[small_col].astype(str) == ss)
            ]
            if len(sub) == 0:
                table.cell(i + 1, j + 2).text = "0人, 0%"
            else:
                r = sub.iloc[0]
                table.cell(i + 1, j + 2).text = f"{int(r['人数'])}人, {float(r['占比(%)']):.2f}%"
    i = 0
    while i < len(rows_list):
        bb = rows_list[i][0]
        j = i
        while j + 1 < len(rows_list) and rows_list[j + 1][0] == bb:
            j += 1
        if j > i:
            table.cell(i + 1, 0).merge(table.cell(j + 1, 0))
        i = j + 1
    _apply_three_line_table(table)

def _degree_to_level(deg: str) -> str:
    """将学历映射为博士、硕士、本科"""
    s = str(deg)
    if "博士" in s:
        return "博士"
    if "本科" in s:
        return "本科"
    if "硕士" in s:
        return "硕士"
    return ""

def plot_stacked_by_level(df_stats: pd.DataFrame, title: str, img_dir: str, file_prefix: str) -> List[str]:
    """根据层次（博士、硕士、本科）绘制堆积柱状图"""
    if len(df_stats) == 0:
        return []
    out_paths: List[str] = []
    df0 = df_stats.copy()
    df0["层次"] = df0["学历"].astype(str).map(_degree_to_level)
    df0 = df0.loc[df0["层次"].isin(["博士", "硕士", "本科"])]
    palette = plt.cm.get_cmap("tab20")
    for level in ["博士", "硕士", "本科"]:
        dfl = df0.loc[df0["层次"] == level]
        if len(dfl) == 0:
            continue
        bigs = sorted(dfl["大类"].astype(str).dropna().unique().tolist())
        segs_all = sorted(dfl["小类"].astype(str).dropna().unique().tolist())
        seg_color: Dict[str, str] = {seg: palette(i % 20) for i, seg in enumerate(segs_all)}
        fig, ax = plt.subplots(figsize=(10, 6))
        x = np.arange(len(bigs))
        bottoms = np.zeros(len(bigs), dtype=float)
        for seg in segs_all:
            heights = []
            for bb in bigs:
                v = dfl.loc[(dfl["大类"].astype(str) == bb) & (dfl["小类"].astype(str) == seg), "人数"]
                heights.append(int(v.iloc[0]) if len(v) else 0)
            heights = np.array(heights, dtype=float)
            if heights.sum() == 0:
                continue
            ax.bar(x, heights, bottom=bottoms, width=0.55, label=seg, color=seg_color.get(seg, "#999999"), edgecolor="white", linewidth=0.6)
            bottoms = bottoms + heights
        total_all = max(1, int(bottoms.sum()))
        for i, bb in enumerate(bigs):
            n = int(bottoms[i])
            p = n / total_all * 100
            ax.text(x[i], n, f"{n}人\n{p:.2f}%", ha="center", va="bottom", fontsize=12)
        ax.set_xticks(x)
        ax.set_xticklabels(bigs, rotation=0, fontsize=12)
        ax.set_ylabel("人数", fontsize=12)
        ax.set_title(f"{title}（{level}）", fontsize=14, pad=20)
        ax.grid(axis="y", linestyle="--", alpha=0.25)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.legend(ncol=2, fontsize=12, loc='upper right', bbox_to_anchor=(1, 1))
        fig.tight_layout(rect=[0, 0, 1, 0.95])
        out_path = os.path.join(img_dir, sanitize_filename(f"{file_prefix}_{level}.png"))
        fig.savefig(out_path, dpi=200, bbox_inches='tight')
        plt.close(fig)
        out_paths.append(out_path)
    return out_paths

def process_phd_research_tendency(df: pd.DataFrame, doc: Document, img_dir: str) -> None:
    """处理博士是否倾向于科研的柱状图绘制"""
    if "层次" not in df.columns or "单位性质" not in df.columns:
        return
    level = _normalize_text_series(df["层次"])
    phd_mask = level.fillna("").astype(str).str.contains("博士研究生", na=False)
    sub = df.loc[phd_mask, "单位性质"]
    mapped = sub.map(_map_unit_nature_phd_tendency)
    big = []
    detail = []
    for x in mapped.tolist():
        if isinstance(x, tuple) and len(x) == 2:
            big.append(x[0])
            detail.append(x[1])
        else:
            big.append("")
            detail.append("")
    tmp = pd.DataFrame({"大类": big, "细类": detail})
    tmp = tmp.loc[(tmp["大类"] != "") & (tmp["细类"] != "")]
    if len(tmp) == 0:
        return
    total = int(len(tmp))
    group_order = ["科研", "求职中", "其他行业"]
    segment_order: Dict[str, List[str]] = {
        "科研": ["高等教育单位", "科研设计单位"],
        "求职中": ["求职中"],
        "其他行业": [],
    }
    other = tmp.loc[tmp["大类"] == "其他行业", "细类"].astype(str)
    if len(other) > 0:
        vc = other.value_counts()
        top_n = 10
        top_details = vc.head(top_n).index.astype(str).tolist()
        other_details = top_details
        if int(vc.iloc[top_n:].sum()) > 0:
            other_details = top_details + ["其他(合并)"]
            tmp.loc[(tmp["大类"] == "其他行业") & (~tmp["细类"].isin(top_details)), "细类"] = "其他(合并)"
        segment_order["其他行业"] = other_details
    colors: Dict[str, str] = {
        "高等教育单位": "#4C78A8",
        "科研设计单位": "#72B7B2",
        "求职中": "#F58518",
        "其他(合并)": "#54A24B",
    }
    palette = plt.cm.get_cmap("tab20")
    for i, seg in enumerate(segment_order["其他行业"]):
        if seg not in colors:
            colors[seg] = palette((i + 4) % 20)
    stats_rows = []
    for g in group_order:
        for s in segment_order[g]:
            n = int(((tmp["大类"] == g) & (tmp["细类"] == s)).sum())
            p = float(n / total * 100)
            stats_rows.append({"分组": g, "细节": s, "人数": n, "占比(%)": p})
    stats_df = pd.DataFrame(stats_rows)
    title = "博士毕业生继续科研倾向统计"
    add_heading(doc, title)
    fig, ax = plt.subplots(figsize=(10, 6))
    x = np.arange(len(group_order))
    bottoms = np.zeros(len(group_order), dtype=float)
    seg_sequence = segment_order["科研"] + segment_order["求职中"] + segment_order["其他行业"]
    for seg in seg_sequence:
        heights = []
        for g in group_order:
            v = stats_df.loc[(stats_df["分组"] == g) & (stats_df["细节"] == seg), "人数"]
            heights.append(int(v.iloc[0]) if len(v) else 0)
        heights = np.array(heights, dtype=float)
        if heights.sum() == 0:
            continue
        ax.bar(x, heights, bottom=bottoms, width=0.5, label=seg, color=colors.get(seg, "#999999"), edgecolor="white", linewidth=0.6)
        bottoms = bottoms + heights
    ax.set_xticks(x)
    ax.set_xticklabels(group_order, fontsize=12)
    ax.set_ylabel("人数", fontsize=12)
    ax.set_title(title, fontsize=14, pad=20)
    ax.grid(axis="y", linestyle="--", alpha=0.25)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    for i, g in enumerate(group_order):
        n = int(stats_df.loc[stats_df["分组"] == g, "人数"].sum())
        p = float(n / total * 100)
        ax.text(x[i], n, f"{n}人\n{p:.2f}%", ha="center", va="bottom", fontsize=12)
    ax.legend(ncol=2, fontsize=12, loc='upper right', bbox_to_anchor=(1, 1))
    fig.tight_layout(rect=[0, 0, 1, 0.95])
    img_path = os.path.join(img_dir, sanitize_filename("博士毕业生继续科研倾向统计.png"))
    fig.savefig(img_path, dpi=200, bbox_inches='tight')
    plt.close(fig)
    add_picture(doc, img_path)
    table = doc.add_table(rows=len(stats_df) + 1, cols=4)
    table.cell(0, 0).text = "类别"
    table.cell(0, 1).text = "细分"
    table.cell(0, 2).text = "人数"
    table.cell(0, 3).text = "占比"
    row_idx = 1
    for g in group_order:
        sub_rows = stats_df.loc[stats_df["分组"] == g].reset_index(drop=True)
        start = row_idx
        for _, r in sub_rows.iterrows():
            table.cell(row_idx, 0).text = g
            table.cell(row_idx, 1).text = str(r["细节"])
            table.cell(row_idx, 2).text = str(int(r["人数"]))
            table.cell(row_idx, 3).text = f"{float(r['占比(%)']):.2f}%"
            row_idx += 1
        end = row_idx - 1
        if end > start:
            table.cell(start, 0).merge(table.cell(end, 0))
    _apply_three_line_table(table)

def process_outcome(df: pd.DataFrame, doc: Document, img_dir: str, edu_col: str, outcome_col: str) -> None:
    """处理毕业去向（按学历与大类分小类）的柱状图绘制"""
    if (edu_col not in df.columns) or (outcome_col not in df.columns):
        return
    df_stats = _build_stats_by_degree_and_mapper(df, outcome_col, edu_col, _map_outcome)
    if len(df_stats) == 0:
        return
    title = "毕业去向（按学历与大类分小类）"
    add_heading(doc, title)
    img_paths = plot_stacked_by_level(df_stats, title, img_dir, "毕业去向_堆叠")
    for p in img_paths:
        add_picture(doc, p)
    add_table_by_degree_groups(doc, df_stats, "毕业去向|")

def process_unit_nature(df: pd.DataFrame, doc: Document, img_dir: str, edu_col: str, unit_col: str) -> None:
    """处理单位性质（科研/非科研/求职中）的柱状图绘制"""
    if (edu_col not in df.columns) or (unit_col not in df.columns):
        return
    df_stats = _build_stats_by_degree_and_mapper(df, unit_col, edu_col, _map_unit_nature)
    if len(df_stats) == 0:
        return
    title = "单位性质（科研/非科研，本科/硕士/博士同图）"
    add_heading(doc, title)
    df0 = df_stats.copy()
    df0["层次"] = df0["学历"].astype(str).map(_degree_to_level)
    df0 = df0.loc[df0["层次"].isin(["本科", "硕士", "博士"])]
    df0 = df0.loc[df0["大类"].astype(str).isin(["科研", "非科研"])]
    if len(df0) > 0:
        df0 = df0.groupby(["层次", "大类", "小类"], as_index=False)["人数"].sum()
        df_long = df0.rename(columns={"大类": "类别", "小类": "细分"})[["层次", "类别", "细分", "人数"]]
        img_path = os.path.join(img_dir, sanitize_filename("单位性质_科研非科研_堆叠_本硕博同图.png"))
        plot_grouped_stacked_binary_by_level(df_long, title, img_path, ["本科", "硕士", "博士"], ["科研", "非科研"])
        add_picture(doc, img_path)
    add_table_by_degree_groups(doc, df_stats, "单位性质|")

def process_master_education_unit_nature(df: pd.DataFrame, doc: Document, img_dir: str, edu_col: str, unit_col: str) -> None:
    """将就业分为是否为教育单位，计算比例并绘图（本科/硕士/博士同图）"""
    if (edu_col not in df.columns) or (unit_col not in df.columns):
        return
    edu = df[edu_col].map(_classify_education)
    level = edu.astype(str).map(_degree_to_level)
    mask = level.isin(["本科", "硕士", "博士"])
    if int(mask.sum()) == 0:
        return
    unit = _normalize_text_series(df.loc[mask, unit_col])
    mapped = unit.map(_map_unit_nature_master_edu)
    tmp = pd.DataFrame(mapped.tolist(), columns=["类别", "细分"])
    tmp["层次"] = level.loc[mask].astype(str).tolist()
    tmp = tmp.loc[(tmp["类别"] != "") & (tmp["细分"] != "")]
    if len(tmp) == 0:
        return
    group_order = ["教育单位", "非教育单位"]
    segs_all = sorted(tmp["细分"].astype(str).replace({"": np.nan, "nan": np.nan}).dropna().unique().tolist())
    if len(segs_all) == 0:
        return
    rows_long = []
    for lv in ["本科", "硕士", "博士"]:
        part_lv = tmp.loc[tmp["层次"].astype(str) == lv]
        for gname in group_order:
            subg = part_lv.loc[part_lv["类别"] == gname, "细分"].astype(str).replace({"": np.nan, "nan": np.nan}).dropna()
            total_g = int(len(subg))
            if total_g == 0:
                for seg in segs_all:
                    rows_long.append({"层次": lv, "类别": gname, "细分": seg, "人数": 0, "占比(%)": 0.0})
                continue
            for seg in segs_all:
                n = int((subg == seg).sum())
                p = float(n / total_g * 100)
                rows_long.append({"层次": lv, "类别": gname, "细分": seg, "人数": n, "占比(%)": p})
    df_long = pd.DataFrame(rows_long)
    title = "毕业生教育单位倾向统计（单位性质，本科/硕士/博士同图）"
    add_heading(doc, title)
    img_path = os.path.join(img_dir, sanitize_filename("教育单位倾向_单位性质_堆叠_本硕博同图.png"))
    plot_grouped_stacked_binary_by_level(df_long, title, img_path, ["本科", "硕士", "博士"], group_order)
    add_picture(doc, img_path)
    level_counts = tmp["层次"].astype(str).value_counts()
    has_all = all(int(level_counts.get(lv, 0)) > 0 for lv in ["本科", "硕士", "博士"])
    if has_all and add_binary_stats_table_by_level(doc, df_long, ["本科", "硕士", "博士"], group_order):
        pass
    else:
        for lv in ["本科", "硕士", "博士"]:
            if int(level_counts.get(lv, 0)) <= 0:
                continue
            part = df_long.loc[df_long["层次"].astype(str) == lv]
            if len(part) == 0:
                continue
            doc.add_paragraph(lv)
            table = doc.add_table(rows=len(part) + 1, cols=4)
            table.cell(0, 0).text = "类别"
            table.cell(0, 1).text = "细分"
            table.cell(0, 2).text = "人数"
            table.cell(0, 3).text = "占比(%)"
            row_idx = 1
            for gname in group_order:
                sub_rows = part.loc[part["类别"] == gname].reset_index(drop=True)
                start = row_idx
                for _, r in sub_rows.iterrows():
                    table.cell(row_idx, 0).text = gname
                    table.cell(row_idx, 1).text = str(r["细分"])
                    table.cell(row_idx, 2).text = str(int(r["人数"]))
                    table.cell(row_idx, 3).text = f"{float(r['占比(%)']):.2f}%"
                    row_idx += 1
                end = row_idx - 1
                if end > start:
                    table.cell(start, 0).merge(table.cell(end, 0))
            _apply_three_line_table(table)

def process_location(df: pd.DataFrame, doc: Document, img_dir: str, edu_col: str, loc_col: str, source_col: str) -> None:
    """处理单位所在地（省/市/国外，按前两汉字归并），处理就业单位是否和出生地同源，非同源是否更多选择北京和上海"""
    if (edu_col not in df.columns) or (loc_col not in df.columns) or (source_col not in df.columns):
        return
    df_loc = df.copy()
    df_loc[loc_col] = df_loc[loc_col].map(_map_location)
    df_loc[source_col] = df_loc[source_col].map(_map_location)
    df_stats_all = _build_stats_by_education(df_loc, loc_col, edu_col)
    if len(df_stats_all) == 0:
        return
    valid = (df_loc[loc_col] != "") & (df_loc[source_col] != "")
    same_source = (df_loc[loc_col] == df_loc[source_col]) & valid
    group = same_source.map(lambda x: "同源" if x else "非同源")
    edu = df_loc[edu_col].map(_classify_education)
    df_group = pd.DataFrame({"学历": edu, "类别": group}).loc[valid]
    rows = []
    for e in _degree_order():
        sub = df_group.loc[df_group["学历"] == e, "类别"]
        if len(sub) == 0:
            continue
        vc = sub.value_counts()
        total = int(vc.sum())
        for k, v in vc.items():
            rows.append({"学历": e, "类别": k, "人数": int(v), "占比(%)": float((v/total*100))})
    df_group_stats = pd.DataFrame(rows)
    add_heading(doc, "同源/非同源占比")
    add_stats_text_by_education(doc, df_group_stats)
    if "单位性质" in df_loc.columns:
        add_heading(doc, "同源/非同源单位性质分布（堆叠图，本科/硕士/博士同图）")
        nature_map = df_loc["单位性质"].map(_map_unit_nature)
        nature_df = pd.DataFrame(nature_map.tolist(), columns=["大类", "细分"])
        df_same = pd.DataFrame({"学历": edu, "大类": nature_df["大类"], "细分": nature_df["细分"], "同源": same_source}).loc[valid]
        df_same["层次"] = df_same["学历"].astype(str).map(_degree_to_level)
        df_same = df_same.loc[df_same["层次"].isin(["本科", "硕士", "博士"])]
        group_order = ["同源", "非同源"]

        if len(df_same) == 0:
            table = doc.add_table(rows=2, cols=2)
            table.cell(0, 0).text = "类别"
            table.cell(0, 1).text = "人数"
            table.cell(1, 0).text = "有效样本数(同时有生源所在地与单位所在地)"
            table.cell(1, 1).text = "0"
            _apply_three_line_table(table)
        df_same2 = df_same.copy()
        df_same2["类别"] = df_same2["同源"].map(lambda x: "同源" if bool(x) else "非同源")
        df_same2 = df_same2.loc[(df_same2["细分"].astype(str) != "") & (df_same2["大类"].astype(str) != "")]
        segs_all = sorted(df_same2["细分"].astype(str).replace({"": np.nan, "nan": np.nan}).dropna().unique().tolist())
        if len(segs_all) > 0:
            rows_long = []
            for lv in ["本科", "硕士", "博士"]:
                part_lv = df_same2.loc[df_same2["层次"].astype(str) == lv]
                for gname in group_order:
                    subg = part_lv.loc[part_lv["类别"] == gname, "细分"].astype(str).replace({"": np.nan, "nan": np.nan}).dropna()
                    total_g = int(len(subg))
                    if total_g == 0:
                        for seg in segs_all:
                            rows_long.append({"层次": lv, "类别": gname, "细分": seg, "人数": 0, "占比(%)": 0.0})
                        continue
                    for seg in segs_all:
                        n = int((subg == seg).sum())
                        p = float(n / total_g * 100)
                        rows_long.append({"层次": lv, "类别": gname, "细分": seg, "人数": n, "占比(%)": p})
            df_long = pd.DataFrame(rows_long)
            title_lv = "同源/非同源单位性质分布（本科/硕士/博士同图）"
            img_path = os.path.join(img_dir, sanitize_filename("同源非同源_单位性质_堆叠_本硕博同图.png"))
            plot_grouped_stacked_binary_by_level(df_long, title_lv, img_path, ["本科", "硕士", "博士"], group_order)
            add_picture(doc, img_path)
            level_counts = df_same2["层次"].astype(str).value_counts()
            has_all = all(int(level_counts.get(lv, 0)) > 0 for lv in ["本科", "硕士", "博士"])
            if has_all and add_binary_stats_table_by_level(doc, df_long, ["本科", "硕士", "博士"], group_order):
                pass
            else:
                for lv in ["本科", "硕士", "博士"]:
                    if int(level_counts.get(lv, 0)) <= 0:
                        continue
                    part = df_long.loc[df_long["层次"].astype(str) == lv]
                    if len(part) == 0:
                        continue
                    doc.add_paragraph(lv)
                    table = doc.add_table(rows=len(part) + 1, cols=4)
                    table.cell(0, 0).text = "类别"
                    table.cell(0, 1).text = "细分"
                    table.cell(0, 2).text = "人数"
                    table.cell(0, 3).text = "占比(%)"
                    row_idx = 1
                    for gname in group_order:
                        sub_rows = part.loc[part["类别"] == gname].reset_index(drop=True)
                        start = row_idx
                        for _, r in sub_rows.iterrows():
                            table.cell(row_idx, 0).text = gname
                            table.cell(row_idx, 1).text = str(r["细分"])
                            table.cell(row_idx, 2).text = str(int(r["人数"]))
                            table.cell(row_idx, 3).text = f"{float(r['占比(%)']):.2f}%"
                            row_idx += 1
                        end = row_idx - 1
                        if end > start:
                            table.cell(start, 0).merge(table.cell(end, 0))
                    _apply_three_line_table(table)

        for flag, tname in [(True, "同源"), (False, "非同源")]:
            part = df_same.loc[df_same["同源"] == flag]
            if len(part) == 0:
                continue
            rows = []
            for e in _degree_order():
                sub = part.loc[part["学历"] == e, "细分"]
                if len(sub) == 0:
                    continue
                vc = sub.value_counts()
                total = int(vc.sum())
                for k, v in vc.items():
                    rows.append({"学历": e, "类别": k, "人数": int(v), "占比(%)": float((v/total*100))})
            df_stat = pd.DataFrame(rows)
            add_heading(doc, f"{tname}组内单位性质占比")
            add_stats_text_by_education_custom(doc, df_stat, "单位性质")
    # 新增：非同源的"单位所在地"详细分布
    df_ns_detail = df_loc.loc[valid & (~same_source)].copy()
    if len(df_ns_detail) > 0:
        df_stats_ns_detail = _build_stats_by_education(df_ns_detail, loc_col, edu_col)
        if len(df_stats_ns_detail) > 0:
            add_heading(doc, "非同源毕业生单位所在地分布")
            title_ns_detail = "非同源毕业生单位所在地分布（本科/硕士/博士堆叠）"
            img_path_ns_detail = os.path.join(img_dir, sanitize_filename("非同源_单位所在地_分布_堆叠.png"))
            plot_location_stacked_by_level(df_stats_ns_detail, title_ns_detail, img_path_ns_detail)
            add_picture(doc, img_path_ns_detail)
            # 配表依然保留总表（横向包含所有学历类别）
            add_stats_text_by_education_custom(doc, df_stats_ns_detail, "单位所在地")
        
    # 在非同源分布表格后添加生源地 - 就业单位所在地网格图
    add_source_employment_heatmap(doc, df_loc, img_dir, edu_col, source_col, loc_col)

    collapsed = df_loc[loc_col].map(_collapse_top_provs)
    df_non_same = pd.DataFrame({"学历": edu, "类别": collapsed, "同源": same_source}).loc[valid]
    df_non_same = df_non_same.loc[df_non_same["同源"] == False]
    if len(df_non_same) > 0:
        rows = []
        for e in _degree_order():
            sub = df_non_same.loc[df_non_same["学历"] == e, "类别"]
            if len(sub) == 0:
                continue
            vc = sub.value_counts()
            total = int(vc.sum())
            for k, v in vc.items():
                rows.append({"学历": e, "类别": k, "人数": int(v), "占比(%)": float((v/total*100))})
        df_stat = pd.DataFrame(rows)
        add_heading(doc, "非同源地区分布（北京/上海/广东/其他/国外）")
        add_stats_text_by_education(doc, df_stat)
    mask_ntnf = (collapsed == "其他") & valid
    subset = pd.DataFrame({"学历": edu, "所在地": df_loc[loc_col], "生源": df_loc[source_col]}).loc[mask_ntnf]
    if len(subset) > 0:
        match_home = (subset["所在地"] == subset["生源"])
        rows = []
        for e in _degree_order():
            sub = match_home.loc[subset["学历"] == e]
            if len(sub) == 0:
                continue
            yes = int(sub.sum())
            no = int(len(sub) - yes)
            total = max(1, yes + no)
            rows.append({"学历": e, "类别": "回生源地", "人数": yes, "占比(%)": float((yes/total*100))})
            rows.append({"学历": e, "类别": "非生源地", "人数": no, "占比(%)": float((no/total*100))})
        df_stat = pd.DataFrame(rows)
        add_heading(doc, "非京沪粤且非国外就业：回生源地 vs 非生源地")
        add_stats_text_by_education(doc, df_stat)

def process_unit_name_for_study(df: pd.DataFrame, doc: Document, img_dir: str, edu_col: str, outcome_col: str, unit_name_col: str) -> None:
    """处理升学单位名称（按博士/硕士分类）的柱状图绘制"""
    if (edu_col not in df.columns) or (outcome_col not in df.columns) or (unit_name_col not in df.columns):
        return
    edu = df[edu_col].map(_classify_education)
    outcome_big = df[outcome_col].map(lambda v: _map_outcome(v)[0])
    df2 = pd.DataFrame({"学历": edu, "单位名称": _normalize_text_series(df[unit_name_col]), "毕业去向大类": outcome_big})
    df2 = df2.loc[(df2["毕业去向大类"] == "升学") & (df2["单位名称"].notna()) & (df2["单位名称"] != "")]
    if len(df2) == 0:
        return
    parts = {
        "本科": df2.loc[df2["学历"] == "本科", "单位名称"].dropna(),
        "硕士": df2.loc[df2["学历"].isin(["学术硕士", "专业硕士", "硕士"]), "单位名称"].dropna(),
    }
    for k, series in parts.items():
        if len(series) == 0:
            continue
        vc = series.value_counts().sort_index()
        total = int(vc.sum())
        stats_df = pd.DataFrame({"人数": vc, "占比(%)": (vc / total * 100).round(2)})
        title = f"升学单位名称分布（{k}）"
        add_heading(doc, title)
        img_path = os.path.join(img_dir, sanitize_filename(f"升学单位名称_{k}.png"))
        plot_counts(stats_df, title, img_path)
        add_picture(doc, img_path)
        table = doc.add_table(rows=len(stats_df) + 1, cols=3)
        table.cell(0, 0).text = "单位名称"
        table.cell(0, 1).text = "人数"
        table.cell(0, 2).text = "占比(%)"
        for i, (name, row) in enumerate(stats_df.iterrows()):
            table.cell(i + 1, 0).text = str(name)
            table.cell(i + 1, 1).text = str(int(row["人数"]))
            table.cell(i + 1, 2).text = f"{float(row['占比(%)']):.2f}%"
        _apply_three_line_table(table)

def count_stats(df: pd.DataFrame, col: str) -> Tuple[pd.DataFrame, int]:
    """统计一列的分布，返回人数和占比"""
    vc = _normalize_text_series(df[col]).dropna()
    total = len(vc)
    counts = vc.value_counts().sort_index()
    perc = (counts / total * 100).round(2)
    out = pd.DataFrame({"人数": counts, "占比(%)": perc})
    return out, total

def plot_counts(df_stats: pd.DataFrame, title: str, out_path: str) -> None:
    """绘制一列的分布柱状图并保存，图中 包含人数和占比"""
    labels = df_stats.index.astype(str).tolist()
    counts = df_stats["人数"].tolist()
    # 使用列名列表访问，避免编码问题
    percent_col = [c for c in df_stats.columns if '占' in c and '%' in c][0]
    percents = df_stats[percent_col].tolist()
    n = len(labels)
    fig_w = max(10, 0.35 * n + 6)
    fig_h = 6
    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    palette = plt.cm.get_cmap("tab20")
    x = np.arange(len(labels))
    colors = [palette(i % 20) for i in range(len(labels))]
    bars = ax.bar(x, counts, width=0.5, color=colors)
    rotate = 45 if len(labels) >= 12 else 30
    tick_fs = 12
    value_fs = 12
    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=rotate, ha="right", fontsize=tick_fs)
    ax.set_ylabel("人数", fontsize=12)
    ax.set_title(title, fontsize=14, pad=20)
    ax.grid(axis="y", linestyle="--", alpha=0.25)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    for i, b in enumerate(bars):
        v = counts[i]
        p = percents[i]
        ax.text(b.get_x() + b.get_width() / 2, b.get_height(), f"{v}人\n{p}%", ha="center", va="bottom", fontsize=value_fs)
    fig.tight_layout(rect=[0, 0, 1, 0.95])
    fig.savefig(out_path, dpi=200, bbox_inches='tight')
    plt.close(fig)

def add_heading(doc: Document, text: str) -> None:
    """设置word中文本的字体格式和大小"""
    p = doc.add_paragraph(text)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.style = doc.styles["Normal"]
    if len(p.runs) == 0:
        p.add_run(text)
    for r in p.runs:
        r.font.size = Pt(12)
        r.font.name = "宋体"

def add_picture(doc: Document, img_path: str, width_inch: float = 6.0) -> None:
    """添加图片到word文档"""
    doc.add_picture(img_path, width=Inches(width_inch))

def add_stats_text(doc: Document, df_stats: pd.DataFrame) -> None:
    """添加统计文本到word文档中的表格里"""
    table = doc.add_table(rows=len(df_stats) + 1, cols=3)
    table.cell(0, 0).text = "类别"
    table.cell(0, 1).text = "人数"
    table.cell(0, 2).text = "占比(%)"
    for i, (idx, row) in enumerate(df_stats.iterrows()):
        table.cell(i + 1, 0).text = str(idx)
        table.cell(i + 1, 1).text = str(int(row["人数"]))
        table.cell(i + 1, 2).text = f"{float(row['占比(%)']):.2f}%"
    _apply_three_line_table(table)

def add_stats_text_by_education(doc: Document, df_stats: pd.DataFrame) -> None:
    """添加统计文本到word文档中的表格里，按学历分类"""
    edu_order = [d for d in _degree_order() if d in set(df_stats["学历"].astype(str).unique().tolist())]
    row_labels = [str(row["类别"]) for _, row in df_stats[["类别"]].drop_duplicates().iterrows()]
    table = doc.add_table(rows=len(row_labels) + 1, cols=len(edu_order) + 1)
    table.cell(0, 0).text = "类别"
    for j, e in enumerate(edu_order):
        table.cell(0, j + 1).text = e
    for i, cat in enumerate(row_labels):
        table.cell(i + 1, 0).text = cat
        for j, e in enumerate(edu_order):
            sub = df_stats.loc[(df_stats["学历"].astype(str) == e) & (df_stats["类别"].astype(str) == cat)]
            if len(sub) == 0:
                table.cell(i + 1, j + 1).text = "0人, 0%"
            else:
                r = sub.iloc[0]
                table.cell(i + 1, j + 1).text = f"{int(r['人数'])}人, {float(r['占比(%)']):.2f}%"
    _apply_three_line_table(table)

def process_categorical(df: pd.DataFrame, col: str, img_dir: str, doc: Document) -> None:
    """处理分类变量，绘制柱状图并添加统计文本"""
    if col not in df.columns:
        return
    edu_col = _find_education_column(list(df.columns)) # 确认学历列
    if edu_col and edu_col in df.columns:
        if col == edu_col:
            s = df[edu_col].map(_classify_education)
            base = s.replace({"": np.nan})
            counts = base.dropna().value_counts().reindex(_degree_order()).dropna()
            total = int(counts.sum())
            if total == 0:
                return
            perc = (counts / total * 100).round(2)
            stats_df = pd.DataFrame({"人数": counts, "占比(%)": perc})
            title = "毕业生学历分布"
            add_heading(doc, title)
            img_name = sanitize_filename(f"{col}_分布.png")
            img_path = os.path.join(img_dir, img_name)
            plot_counts(stats_df, title, img_path)
            add_picture(doc, img_path)
            df_tmp = pd.DataFrame({"类别": stats_df.index.astype(str), "人数": stats_df["人数"].astype(int), "占比(%)": stats_df["占比(%)"].astype(float), "学历": stats_df.index.astype(str)})
            add_stats_text_by_education(doc, df_tmp)
            return
        df_for_stats = df
        if col in {"单位所在地", "生源地", "生源所在地"}:
            df_for_stats = df.copy()
            df_for_stats[col] = df_for_stats[col].map(_map_location)
        df_stats = _build_stats_by_education(df_for_stats, col, edu_col)
        if len(df_stats) == 0:
            return
        title = f"毕业生{col}分布"
        add_heading(doc, title)
        if col == "单位所在地":
            img_name = sanitize_filename(f"{col}_分布.png")
            img_path = os.path.join(img_dir, img_name)
            plot_location_stacked_by_level(df_stats, f"{title}（本科/硕士/博士堆叠）", img_path)
            add_picture(doc, img_path)
            add_stats_text_by_education(doc, df_stats)
        elif col == "生源地" or col == "生源所在地":
            # 生源地也使用堆叠柱状图，本硕博用纹样区分
            img_name = sanitize_filename(f"{col}_分布.png")
            img_path = os.path.join(img_dir, img_name)
            plot_location_stacked_by_level(df_stats, f"{title}（本科/硕士/博士堆叠）", img_path)
            add_picture(doc, img_path)
            add_stats_text_by_education(doc, df_stats)
        elif len(df_stats) > 20:
            edu_order = _degree_order()
            for e in edu_order:
                sub = df_stats.loc[df_stats["学历"].astype(str) == e]
                if len(sub) == 0:
                    continue
                stats_df_sub = sub.set_index("类别")[["人数", "占比(%)"]]
                doc.add_paragraph(e)
                img_name = sanitize_filename(f"{col}_分布_{e}.png")
                img_path = os.path.join(img_dir, img_name)
                plot_counts(stats_df_sub, f"{title}-{e}", img_path)
                add_picture(doc, img_path)
            add_stats_text_by_education(doc, df_stats)
        else:
            img_name = sanitize_filename(f"{col}_分布.png")
            img_path = os.path.join(img_dir, img_name)
            plot_counts_by_education(df_stats, title, img_path)
            add_picture(doc, img_path)
            add_stats_text_by_education(doc, df_stats)
        return
    stats_df, total = count_stats(df, col)
    if total == 0 or len(stats_df) == 0:
        return
    title = f"毕业生{col}分布"
    add_heading(doc, title)
    img_name = sanitize_filename(f"{col}_分布.png")
    img_path = os.path.join(img_dir, img_name)
    plot_counts(stats_df, title, img_path)
    add_picture(doc, img_path)
    add_stats_text(doc, stats_df)

def select_file_and_dir(default_doc_name: str = "毕业生统计报告") -> Tuple[str, str, str]:
    """弹窗选择xlsx文件和保存目录"""
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo("提示", "请选择xlsx格式的文件")
    xlsx_path = filedialog.askopenfilename(title="选择xlsx文件", filetypes=[("Excel 文件", "*.xlsx")])
    if not xlsx_path:
        raise RuntimeError("未选择xlsx文件")
    messagebox.showinfo("提示", "请选择保存Word文件的目录")
    out_dir = filedialog.askdirectory(title="选择保存目录")
    if not out_dir:
        raise RuntimeError("未选择保存目录")
    doc_name = simpledialog.askstring("文件名", "请输入Word文件名（无需扩展名）", initialvalue=default_doc_name)
    if not doc_name:
        doc_name = default_doc_name
    doc_name = sanitize_filename(doc_name) + ".docx"
    root.destroy()
    return xlsx_path, out_dir, doc_name


def generate_report(xlsx_path: str, out_dir: str, doc_name: str) -> str:
    """生成统计分析报告"""
    ensure_dir(out_dir)
    # 读取excel文件
    df = read_excel(xlsx_path)
    doc = Document()
    # 绘图目录
    img_dir = os.path.join(out_dir, "images")
    ensure_dir(img_dir)
    edu_col = _find_education_column(list(df.columns)) #确认学历列存在
    # 在选出的每个列绘制分布图
    for col in CATEGORICAL_COLUMNS:
        if col == "单位名称":
            continue
        process_categorical(df, col, img_dir, doc)
    if edu_col:
        outcome_col = _pick_col(list(df.columns), "毕业去向")
        unit_nature_col = _pick_col(list(df.columns), "单位性质")
        unit_loc_col = _pick_col(list(df.columns), "单位所在地")
        source_col = _pick_col(list(df.columns), "生源地", "生源所在地")
        unit_name_col = _pick_col(list(df.columns), "单位名称")

        if outcome_col:
            process_outcome(df, doc, img_dir, edu_col, outcome_col) # 处理毕业去向
        if unit_nature_col:
            process_unit_nature(df, doc, img_dir, edu_col, unit_nature_col) # 处理单位性质
            process_master_education_unit_nature(df, doc, img_dir, edu_col, unit_nature_col)
        if unit_loc_col and source_col:
            process_location(df, doc, img_dir, edu_col, unit_loc_col, source_col) # 处理单位所在地和生源地
        if unit_name_col and outcome_col:
            process_unit_name_for_study(df, doc, img_dir, edu_col, outcome_col, unit_name_col) # 处理单位名称和毕业去向
    process_chi_square_research_influences(df, doc, img_dir)
    # 绘制博士选择科研倾向工作的柱状图
    process_phd_research_tendency(df, doc, img_dir)
    # 保存报告
    doc_path = os.path.join(out_dir, doc_name)
    doc.save(doc_path)
    return doc_path

def main() -> None:
    """主函数"""
    """弹窗选择目录"""
    parser = argparse.ArgumentParser()
    parser.add_argument("--xlsx", type=str, default=None)
    parser.add_argument("--outdir", type=str, default=None)
    parser.add_argument("--docname", type=str, default=None)
    args = parser.parse_args()
    """确认文件后缀，确认文件名是否规范，生成报告"""
    if args.xlsx and args.outdir and args.docname:
        xlsx_path = args.xlsx
        out_dir = args.outdir
        doc_name = sanitize_filename(args.docname)
        if not doc_name.lower().endswith(".docx"):
            doc_name += ".docx"
        # 生成报告
        result = generate_report(xlsx_path, out_dir, doc_name)
        print(result)
        return
    try:
        xlsx_path, out_dir, doc_name = select_file_and_dir()
        result = generate_report(xlsx_path, out_dir, doc_name)
        # 弹窗提示完成
        messagebox.showinfo("完成", f"已生成Word文件：{result}")
    except Exception as e:
        log_path = _write_log("运行失败", "".join(traceback.format_exception(type(e), e, e.__traceback__)))
        try:
            if log_path:
                messagebox.showerror("错误", f"{e}\n\n已写入日志：{log_path}")
            else:
                messagebox.showerror("错误", str(e))
        except Exception:
            pass

if __name__ == "__main__":
    main()
