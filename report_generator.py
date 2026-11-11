import os
import sys
import math
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
try:
    import geopandas as gpd
except Exception:
    gpd = None
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
try:
    from scipy import stats as scipy_stats
except Exception:
    scipy_stats = None

# 设置中文字体：按顺序取系统中第一个可用的字体，避免非 Windows 环境出现方框
plt.rcParams["font.sans-serif"] = [
    "SimHei",
    "Microsoft YaHei",
    "Noto Sans CJK SC",
    "WenQuanYi Zen Hei",
    "Arial Unicode MS",
]
plt.rcParams["axes.unicode_minus"] = False

# 定义必须存在的列名
REQUIRED_COLUMNS = [
    "学历", "层次", "师范生类别", "培养方式", "生源所在地", "毕业生去向",
    "单位名称", "单位性质", "单位所在地", "工作职位类别", "定向或委培单位"
]

# --- 基础工具函数 ---

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
    """应用三线表样式"""
    for row in table.rows:
        for cell in row.cells:
            _set_cell_border(
                cell,
                top={"val": "none"}, bottom={"val": "none"},
                left={"val": "none"}, right={"val": "none"},
                insideH={"val": "none"}, insideV={"val": "none"},
            )
    
    top_border = {"val": "single", "sz": 12, "color": "000000"}
    mid_border = {"val": "single", "sz": 8, "color": "000000"}
    bottom_border = {"val": "single", "sz": 12, "color": "000000"}
    
    header_row = table.rows[0]
    for cell in header_row.cells:
        _set_cell_border(cell, top=top_border, bottom=mid_border)
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                r.bold = True
    
    last_row = table.rows[-1]
    for cell in last_row.cells:
        _set_cell_border(cell, bottom=bottom_border)
    
    table.autofit = True

def sanitize_filename(name: str) -> str:
    """文件名规范化"""
    invalid = '<>:"/\\|?*'
    for ch in invalid:
        name = name.replace(ch, "_")
    return name.strip() or f"untitled_{uuid.uuid4().hex[:8]}"

def add_table_of_contents(doc: Document):
    """在文档中添加自动目录域，并设置打开时自动提示更新"""
    paragraph = doc.add_paragraph()
    run = paragraph.add_run()
    
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = 'TOC \\o "1-3" \\h \\z \\u'
    
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)
    run._r.append(fldChar3)
    
    # 设置打开文档时自动提示更新域（包括目录）
    element_update_fields = OxmlElement('w:updateFields')
    element_update_fields.set(qn('w:val'), 'true')
    doc.settings.element.append(element_update_fields)

def add_heading(doc: Document, text: str, level: int = 1):
    """添加标题"""
    p = doc.add_heading(text, level=level)
    return p

def add_paragraph(doc: Document, text: str = ""):
    """添加正文段落"""
    p = doc.add_paragraph(text)
    return p

def add_picture(doc: Document, path: str, width_inches: float = 6.0):
    """添加图片并居中"""
    if not os.path.exists(path):
        return
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    run.add_picture(path, width=Inches(width_inches))

# --- 数据预处理函数 ---

def read_excel(path: str) -> pd.DataFrame:
    """读取Excel并清理列名"""
    df = pd.read_excel(path, dtype=object)
    df.columns = [str(c).replace("\u3000", "").strip() for c in df.columns]
    for col in df.columns:
        df[col] = df[col].map(lambda v: "" if pd.isna(v) else str(v).strip())
    return df

def _pick_col(columns: List[str], *candidates: str) -> str:
    """选择存在的列"""
    for cand in candidates:
        if cand in columns:
            return cand
    for cand in candidates:
        for c in columns:
            if cand in c:
                return c
    return ""


# 程序实际会读取的列：缺失时给出可读的提示，而不是抛出 KeyError
USED_COLUMNS = {
    "学历": ("学历",),
    "培养方式": ("培养方式", "定向"),
    "生源所在地": ("生源所在地", "生源地"),
    "毕业生去向": ("毕业生去向", "毕业去向"),
    "单位性质": ("单位性质",),
    "单位所在地": ("单位所在地",),
    "单位名称": ("单位名称",),
}


def validate_columns(columns) -> None:
    """校验表头是否齐全，缺失时抛出中文错误提示"""
    columns = list(columns)
    missing = [label for label, cands in USED_COLUMNS.items() if not _pick_col(columns, *cands)]
    if missing:
        raise ValueError(
            "数据表缺少必需的列：{}。\n当前表头为：{}\n请对照《使用说明》核对表头（注意全角空格与错别字）。".format(
                "、".join(missing), "、".join(str(c) for c in columns)
            )
        )

def _map_location(v: str) -> str:
    """省份前2个汉字"""
    s = str(v).strip()
    if not s or s.lower() == "nan": return ""
    if "国外" in s or "境外" in s or "出国" in s or "留学" in s: return "国外"
    return s[:2]

def _classify_education(v: str) -> str:
    """学历归类"""
    s = str(v).strip()
    if "博士" in s or "博" in s: return "博士"
    if "硕士" in s or "硕" in s: return "硕士"
    if "本科" in s or "本" in s: return "本科"
    return ""

def _degree_to_level(v: str) -> str:
    """层次归类"""
    s = str(v).strip()
    if "博士" in s or "博" in s: return "博士"
    if "硕士" in s or "硕" in s: return "硕士"
    if "本科" in s or "本" in s: return "本科"
    return ""

def _is_stable_work(nature: str) -> bool:
    """定义稳定工作"""
    stable_types = ["机关", "国有企业", "部队", "其他事业单位", "高等教育单位", "中初教育单位", "医疗卫生单位", "科研设计单位"]
    for t in stable_types:
        if t in nature:
            return True
    return False

def _is_research_unit(nature: str) -> bool:
    """博士：是否科研单位"""
    return "高等教育单位" in nature or "科研设计单位" in nature

def _is_education_unit(nature: str) -> bool:
    """定义教育行业偏向：仅包含中初教育单位，不包含高等教育单位"""
    return "中初教育单位" in nature

def _map_outcome_simple(v: str) -> str:
    """去向映射：落实 / 未落实"""
    s = str(v).strip()
    if not s: return "未落实"
    unsettled = ["拟升学", "待就业", "不就业拟升学", "暂不就业", "求职中"]
    for u in unsettled:
        if u in s:
            return "未落实"
    return "落实"

def _map_further_study(v: str) -> str:
    """升学映射：国内 / 境外"""
    s = str(v).strip()
    if "境外" in s or "出国" in s or "留学" in s:
        return "境外升学"
    if "研究生" in s or "升学" in s:
        return "国内升学"
    return ""

# --- 绘图函数 ---

def plot_counts_bar(counts_dict: Dict[str, int], title: str, out_path: str):
    """绘制简单柱状图"""
    if not counts_dict: return
    labels = list(counts_dict.keys())
    values = list(counts_dict.values())
    
    fig, ax = plt.subplots(figsize=(max(10, len(labels) * 1.0), 7))
    bars = ax.bar(labels, values, color="#4C78A8", width=0.5)
    ax.set_title(title, fontsize=18, pad=20)
    ax.set_ylabel("人数", fontsize=16)
    
    # 添加数值标签
    for bar in bars:
        height = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2., height, f"{int(height)}",
                ha='center', va='bottom', fontsize=16)
    
    plt.xticks(rotation=45 if len(labels) > 5 else 0, fontsize=16)
    plt.yticks(fontsize=16)
    fig.tight_layout()
    fig.savefig(out_path, dpi=200, bbox_inches='tight')
    plt.close(fig)

def plot_education_orient_grouped_bar(df_all: pd.DataFrame, orient_col: str, title: str, out_path: str):
    """绘制本硕博定向/非定向的分组柱状图"""
    levels = ["本科", "硕士", "博士"]
    
    non_dir_vals = []
    dir_vals = []
    
    for lv in levels:
        # 获取该学历下的所有人
        df_level = df_all[df_all["_学历分类"] == lv]
        # 统计非定向人数
        non_dir_count = len(df_level[df_level[orient_col].astype(str).str.contains("非定向", na=False)])
        # 定向人数 = 总人数 - 非定向人数
        dir_count = len(df_level) - non_dir_count
        
        non_dir_vals.append(non_dir_count)
        dir_vals.append(dir_count)
    
    x = np.arange(len(levels))
    width = 0.35
    
    fig, ax = plt.subplots(figsize=(12, 7))
    rects1 = ax.bar(x - width/2, non_dir_vals, width, label='非定向', color='#4C78A8')
    rects2 = ax.bar(x + width/2, dir_vals, width, label='定向', color='#F58518')
    
    ax.set_ylabel('人数', fontsize=16)
    ax.set_title(title, fontsize=18, pad=25)
    ax.set_xticks(x)
    ax.set_xticklabels(levels, fontsize=16)
    ax.legend(fontsize=16)
    
    def autolabel(rects):
        for rect in rects:
            height = rect.get_height()
            ax.annotate(f'{int(height)}',
                        xy=(rect.get_x() + rect.get_width() / 2, height),
                        xytext=(0, 3),
                        textcoords="offset points",
                        ha='center', va='bottom', fontsize=16)

    autolabel(rects1)
    autolabel(rects2)
    
    plt.yticks(fontsize=16)
    fig.tight_layout()
    fig.savefig(out_path, dpi=200, bbox_inches='tight')
    plt.close(fig)

def plot_stacked_bar(df_filtered: pd.DataFrame, col: str, title: str, out_path: str, top_n: int = 20):
    """通用的本硕博堆叠柱状图"""
    levels = ["本科", "硕士", "博士"]
    # 统计每个类别的总人数并排序
    # 修复：先对数据进行清洗，去除空值
    df_plot = df_filtered[df_filtered[col] != ""].copy()
    totals = df_plot[col].value_counts()
    top_items = totals.head(top_n).index.tolist()
    
    if not top_items: return
    
    # 准备数据
    level_data = {lv: [] for lv in levels}
    for item in top_items:
        for lv in levels:
            count = len(df_plot[(df_plot[col] == item) & (df_plot["_学历分类"] == lv)])
            level_data[lv].append(count)
            
    x = np.arange(len(top_items))
    width = 0.6
    
    fig, ax = plt.subplots(figsize=(max(12, len(top_items) * 1.0), 8))
    
    bottoms = np.zeros(len(top_items))
    colors = ['#4C78A8', '#F58518', '#E45756']
    
    for i, lv in enumerate(levels):
        ax.bar(x, level_data[lv], width, bottom=bottoms, label=lv, color=colors[i])
        bottoms += np.array(level_data[lv])
        
    ax.set_ylabel('人数', fontsize=16)
    ax.set_title(title, fontsize=18, pad=25)
    ax.set_xticks(x)
    ax.set_xticklabels(top_items, fontsize=16, rotation=45, ha='right')
    ax.legend(fontsize=16)
    
    for i, total in enumerate(bottoms):
        if total > 0:
            ax.text(i, total, f'{int(total)}', ha='center', va='bottom', fontsize=16)
            
    plt.yticks(fontsize=16)
    fig.tight_layout()
    fig.savefig(out_path, dpi=200, bbox_inches='tight')
    plt.close(fig)

def add_stats_table_by_education(doc: Document, df_filtered: pd.DataFrame, col: str, header_text: str):
    """添加本硕博分布统计表"""
    levels = ["本科", "硕士", "博士"]
    # 修复：去除空值，避免第一列出现空白行
    df_table = df_filtered[df_filtered[col] != ""].copy()
    categories = df_table[col].value_counts().index.tolist()
    if not categories: return
    
    table = doc.add_table(rows=len(categories) + 1, cols=len(levels) + 1)
    table.cell(0, 0).text = header_text
    for j, lv in enumerate(levels):
        table.cell(0, j + 1).text = lv
        
    for i, cat in enumerate(categories):
        table.cell(i + 1, 0).text = str(cat).strip()
        for j, lv in enumerate(levels):
            count = len(df_table[(df_table[col] == cat) & (df_table["_学历分类"] == lv)])
            table.cell(i + 1, j + 1).text = str(count)
            
    _apply_three_line_table(table)

def get_top_category_text(df_filtered: pd.DataFrame, col: str):
    """获取各学历人数最多的类别名称"""
    results = {}
    for lv in ["本科", "硕士", "博士"]:
        sub = df_filtered[df_filtered["_学历分类"] == lv]
        if not sub.empty:
            # 清洗数据，去除空值和无效值
            cleaned_data = sub[col][sub[col] != ""].copy()
            cleaned_data = cleaned_data[cleaned_data.notna()]
            
            if not cleaned_data.empty:
                top = cleaned_data.value_counts().idxmax()
                results[lv] = str(top)
            else:
                results[lv] = "无"
        else:
            results[lv] = "无"
    return results

def plot_economic_zone_pie_charts(df_filtered: pd.DataFrame, title: str, out_path: str):
    """
    为本硕博各绘制一个饼图，展示核心经济区分布。
    三个饼图横向排成一排。
    """
    levels = ["本科", "硕士", "博士"]
    zones = ["京津冀", "珠三角", "长三角", "成渝", "其他"]
    fig, axes = plt.subplots(1, 3, figsize=(20, 8))
    
    # 颜色方案：使用鲜明的配色
    colors = ['#4C78A8', '#F58518', '#E45756', '#72B7B2', '#E1E1E1'] 
    
    has_data = False
    for i, lv in enumerate(levels):
        ax = axes[i]
        df_lv = df_filtered[df_filtered["_学历分类"] == lv]
        
        if df_lv.empty:
            ax.text(0.5, 0.5, f"{lv}\n(暂无数据)", ha='center', va='center', fontsize=16)
            ax.axis('off')
            continue
        
        has_data = True
        # 统计各区人数
        counts = [len(df_lv[df_lv["_经济区"] == z]) for z in zones]
        labels = [f"{z}\n({c}人)" for z, c in zip(zones, counts)]
        
        # 过滤 0
        plot_counts = [c for c in counts if c > 0]
        plot_labels = [l for c, l in zip(counts, labels) if c > 0]
        plot_colors = [col for c, col in zip(counts, colors) if c > 0]

        if sum(plot_counts) > 0:
            ax.pie(
                plot_counts, 
                labels=plot_labels, 
                autopct='%1.1f%%',
                startangle=140, 
                colors=plot_colors,
                textprops={'fontsize': 13, 'fontweight': 'bold'},
                pctdistance=0.75,
                explode=[0.05] * len(plot_counts)
            )
        else:
            ax.text(0.5, 0.5, f"{lv}\n(人数为0)", ha='center', va='center', fontsize=16)
        
        ax.set_title(lv, fontsize=20, pad=15, fontweight='bold')

    fig.suptitle(title, fontsize=24, fontweight='bold', y=1.05)
    fig.tight_layout()
    
    if has_data:
        fig.savefig(out_path, dpi=200, bbox_inches='tight')
    plt.close(fig)

def plot_preference_pie_charts(df_filtered: pd.DataFrame, nature_col: str, preference_func, title: str, out_path: str, pref_label: str = "倾向"):
    """
    为本硕博各绘制一个饼图，展示偏向分布。
    三个饼图横向排成一排。
    """
    levels = ["本科", "硕士", "博士"]
    fig, axes = plt.subplots(1, 3, figsize=(18, 7))
    
    # 统一配色：偏向类用蓝色系，非偏向类用浅灰色
    colors = ['#4C78A8', '#E1E1E1'] 
    
    has_data = False
    for i, lv in enumerate(levels):
        ax = axes[i]
        df_lv = df_filtered[df_filtered["_学历分类"] == lv]
        
        if df_lv.empty:
            ax.text(0.5, 0.5, f"{lv}\n(暂无数据)", ha='center', va='center', fontsize=16)
            ax.axis('off')
            continue
        
        has_data = True
        pref_count = len(df_lv[df_lv[nature_col].apply(preference_func)])
        non_pref_count = len(df_lv) - pref_count
        
        counts = [pref_count, non_pref_count]
        labels = [f"{pref_label}\n({pref_count}人)", f"其他\n({non_pref_count}人)"]
        
        # 仅当有数据时绘制饼图
        if sum(counts) > 0:
            # 过滤掉 0 的项，防止饼图显示空白标签
            plot_counts = [c for c in counts if c > 0]
            plot_labels = [l for c, l in zip(counts, labels) if c > 0]
            plot_colors = [colors[0] if c == pref_count and c > 0 else colors[1] for c in plot_counts]
            # 如果两个都有，确保颜色顺序正确
            if len(plot_counts) == 2:
                plot_colors = [colors[0], colors[1]]

            ax.pie(
                plot_counts, 
                labels=plot_labels, 
                autopct='%1.1f%%',
                startangle=90, 
                colors=plot_colors,
                textprops={'fontsize': 14, 'fontweight': 'bold'},
                pctdistance=0.75,
                explode=[0.05] * len(plot_counts)
            )
        else:
            ax.text(0.5, 0.5, f"{lv}\n(人数为0)", ha='center', va='center', fontsize=16)
        
        ax.set_title(lv, fontsize=18, pad=10, fontweight='bold')

    fig.suptitle(title, fontsize=22, fontweight='bold', y=1.05)
    fig.tight_layout()
    
    if has_data:
        fig.savefig(out_path, dpi=200, bbox_inches='tight')
    plt.close(fig)

def add_preference_stats_table(doc: Document, df_filtered: pd.DataFrame, nature_col: str, preference_func, title_prefix: str):
    """添加偏向统计表"""
    levels = ["本科", "硕士", "博士"]
    rows = []
    
    for lv in levels:
        df_lv = df_filtered[df_filtered["_学历分类"] == lv]
        total = len(df_lv)
        if total == 0:
            rows.append([lv, "0", "0", "0.00%"])
            continue
            
        pref_count = len(df_lv[df_lv[nature_col].apply(preference_func)])
        non_pref_count = total - pref_count
        ratio = (pref_count / total) * 100
        rows.append([lv, str(pref_count), str(non_pref_count), f"{ratio:.2f}%"])
        
    table = doc.add_table(rows=len(rows) + 1, cols=4)
    table.cell(0, 0).text = "学历层次"
    table.cell(0, 1).text = f"选择{title_prefix}人数"
    table.cell(0, 2).text = f"未选择{title_prefix}人数"
    table.cell(0, 3).text = "占比(%)"
    
    for i, row_data in enumerate(rows):
        for j, val in enumerate(row_data):
            table.cell(i + 1, j).text = val
            
    _apply_three_line_table(table)

def plot_chi_square_stacked_bar(df_filtered: pd.DataFrame, title: str, out_path: str):
    """
    绘制回乡与稳定工作偏好的分组堆叠柱状图。
    X轴为学历（本、硕、博），每个学历下有两根柱子（回乡 vs 不回乡）。
    每根柱子由“稳定”和“非稳定”堆叠而成。
    """
    levels = ["本科", "硕士", "博士"]
    groups = ["回乡", "不回乡"]
    
    fig, ax = plt.subplots(figsize=(14, 8))
    
    width = 0.35
    x = np.arange(len(levels))
    
    # 颜色：稳定用深绿色，非稳定用浅灰色
    colors = ['#54A24B', '#E1E1E1']
    
    for i, lv in enumerate(levels):
        df_lv = df_filtered[df_filtered["_学历分类"] == lv]
        
        for j, group in enumerate(groups):
            sub = df_lv[df_lv["_是否回乡"] == group]
            offset = -width/2 - 0.02 if j == 0 else width/2 + 0.02
            
            if sub.empty:
                continue
                
            stable_count = len(sub[sub["_是否稳定"] == "稳定"])
            non_stable_count = len(sub) - stable_count
            
            # 堆叠绘图
            ax.bar(x[i] + offset, stable_count, width, label='稳定' if i==0 and j==0 else "", color=colors[0], edgecolor='white')
            ax.bar(x[i] + offset, non_stable_count, width, bottom=stable_count, label='非稳定' if i==0 and j==0 else "", color=colors[1], edgecolor='white')
            
            # 标注总数
            total = stable_count + non_stable_count
            if total > 0:
                ax.text(x[i] + offset, total, f'{int(total)}', ha='center', va='bottom', fontsize=14)

    ax.set_ylabel('人数', fontsize=16)
    ax.set_title(title, fontsize=20, pad=30, fontweight='bold')
    ax.set_xticks(x)
    ax.set_xticklabels(levels, fontsize=16)
    ax.legend(fontsize=14)
    
    plt.yticks(fontsize=16)
    fig.tight_layout()
    fig.savefig(out_path, dpi=200, bbox_inches='tight')
    plt.close(fig)

def plot_outcome_rate_grouped_bar(df: pd.DataFrame, title: str, out_path: str):
    """绘制本硕博去向落实/未落实的分组柱状图"""
    levels = ["本科", "硕士", "博士"]
    
    settled_vals = []
    unsettled_vals = []
    
    for lv in levels:
        sub = df[df["_学历分类"] == lv]
        settled_count = len(sub[sub["_去向状态"] == "落实"])
        unsettled_count = len(sub) - settled_count
        
        settled_vals.append(settled_count)
        unsettled_vals.append(unsettled_count)
    
    x = np.arange(len(levels))
    width = 0.35
    
    fig, ax = plt.subplots(figsize=(12, 7))
    rects1 = ax.bar(x - width/2, settled_vals, width, label='落实', color='#54A24B')
    rects2 = ax.bar(x + width/2, unsettled_vals, width, label='未落实', color='#E45756')
    
    ax.set_ylabel('人数', fontsize=16)
    ax.set_title(title, fontsize=18, pad=25)
    ax.set_xticks(x)
    ax.set_xticklabels(levels, fontsize=16)
    ax.legend(fontsize=16)
    
    def autolabel(rects):
        for rect in rects:
            height = rect.get_height()
            ax.annotate(f'{int(height)}',
                        xy=(rect.get_x() + rect.get_width() / 2, height),
                        xytext=(0, 3),
                        textcoords="offset points",
                        ha='center', va='bottom', fontsize=16)

    autolabel(rects1)
    autolabel(rects2)
    
    plt.yticks(fontsize=16)
    fig.tight_layout()
    fig.savefig(out_path, dpi=200, bbox_inches='tight')
    plt.close(fig)

def plot_heatmap_grid(df: pd.DataFrame, source_col: str, dest_col: str, title: str, out_path: str):
    """绘制生源地-就业地网格热力图"""
    if df.empty: return
    
    # 构建交叉表
    pivot = pd.crosstab(df[dest_col], df[source_col])
    if pivot.empty: return
    
    # 统一行列标签并排序
    all_labels = sorted(set(pivot.index) | set(pivot.columns))
    pivot = pivot.reindex(index=all_labels, columns=all_labels, fill_value=0)
    
    fig_size = max(10, len(all_labels) * 0.7)
    fig, ax = plt.subplots(figsize=(fig_size + 2.5, fig_size))
    
    # 创建自定义颜色映射：0 为灰色，其他用 YlOrRd
    from matplotlib.colors import LinearSegmentedColormap
    base_cmap = plt.cm.YlOrRd
    colors = [(0.8, 0.8, 0.8)] + [base_cmap(i) for i in range(base_cmap.N)]
    custom_cmap = LinearSegmentedColormap.from_list('gray_zero', colors, N=base_cmap.N + 1)
    
    # 绘制热力图 - vmin 设置为 0，让 0 映射到列表第一个颜色（灰色）
    im = ax.imshow(pivot.values, cmap=custom_cmap, vmin=0, vmax=max(pivot.values.max(), 1))
    
    ax.set_xticks(np.arange(len(all_labels)))
    ax.set_yticks(np.arange(len(all_labels)))
    ax.set_xticklabels(all_labels, fontsize=16, rotation=45, ha="right")
    ax.set_yticklabels(all_labels, fontsize=16)
    
    # 标题调大更多
    ax.set_title(title, fontsize=24, pad=30, fontweight='bold')
    ax.set_xlabel("生源地", fontsize=18)
    ax.set_ylabel("就业单位所在地", fontsize=18)
    
    # 标注数值 - 0 不显示，数字调大
    for i in range(len(all_labels)):
        for j in range(len(all_labels)):
            val = pivot.values[i, j]
            if val > 0:
                color = "white" if val > pivot.values.max() * 0.6 else "black"
                ax.text(j, i, str(int(val)), ha="center", va="center", color=color, fontsize=16, fontweight='bold')
            
    cbar = fig.colorbar(im, ax=ax, shrink=0.8)
    cbar.ax.tick_params(labelsize=16)
    fig.tight_layout()
    fig.savefig(out_path, dpi=200, bbox_inches='tight')
    plt.close(fig)

def plot_china_map_geojson(data_dict: Dict[str, int], title: str, out_path: str):
    """使用 GeoJSON 绘制中国地图热力图"""
    if gpd is None:
        return
    
    try:
        # 使用更可靠的 GeoJSON 数据源
        url = "https://geo.datav.aliyun.com/areas_v3/bound/100000_full.json"
        
        # 预处理数据字典：统一去掉“省”、“市”等后缀，只留核心地名
        clean_data = {}
        for k, v in data_dict.items():
            name = k.replace("省", "").replace("市", "").replace("自治区", "").replace("维吾尔", "").replace("壮族", "").replace("回族", "").replace("特别行政区", "").strip()
            if name:
                clean_data[name] = v

        # 读取地理数据
        gdf = gpd.read_file(url)
        
        # 处理地图中的名称：同样提取核心地名进行匹配
        def get_core_name(name):
            return name.replace("省", "").replace("市", "").replace("自治区", "").replace("维吾尔", "").replace("壮族", "").replace("回族", "").replace("特别行政区", "").strip()
        
        gdf['core_name'] = gdf['name'].apply(get_core_name)
        
        # 将统计数据合并到 GeoDataFrame
        stats_df = pd.DataFrame(list(clean_data.items()), columns=['core_name', '人数'])
        gdf = gdf.merge(stats_df, on='core_name', how='left')
        gdf['人数'] = gdf['人数'].fillna(0)
        
        # 开始绘图
        fig, ax = plt.subplots(figsize=(14, 12))
        
        # 绘制背景 (淡灰色)
        gdf.plot(ax=ax, color='#f5f5f5', edgecolor='#999999', linewidth=0.8)
        
        # 绘制热力图
        if gdf['人数'].max() > 0:
            import matplotlib.colors as colors
            # 使用 PowerNorm 让低数值区域的颜色变化更明显 (gamma < 1)
            # 或者使用 LogNorm，但 LogNorm 不支持 0，所以这里用 PowerNorm
            norm = colors.PowerNorm(gamma=0.5, vmin=1, vmax=gdf['人数'].max())
            
            gdf.plot(
                column='人数',
                ax=ax,
                cmap='YlOrRd',
                norm=norm,
                legend=True,
                legend_kwds={'label': "就业人数", 'shrink': 0.4, 'aspect': 30},
                edgecolor='#666666',
                linewidth=0.8
            )
        
        # 在省份中心标注人数 (仅标注人数 > 0 的)
        # 针对京津冀等拥挤地区进行偏移处理
        offsets = {
            "北京": (1.5, 0.5),   # 向右上方偏移
            "天津": (1.5, -0.5),  # 向右下方偏移
            "河北": (-1.5, 0),    # 向左偏移
            "上海": (1.5, 0),     # 向右偏移
            "香港": (1.0, -1.0),
            "澳门": (-1.0, -1.0)
        }

        for _, row in gdf.iterrows():
            if row['人数'] > 0:
                name = row['core_name']
                center = row.geometry.representative_point().coords[0]
                
                if name in offsets:
                    off_x, off_y = offsets[name]
                    text_pos = (center[0] + off_x, center[1] + off_y)
                    # 画指引线
                    ax.annotate(
                        f"{name}\n{int(row['人数'])}",
                        xy=center,
                        xytext=text_pos,
                        fontsize=12,
                        ha='center',
                        va='center',
                        fontweight='bold',
                        arrowprops=dict(arrowstyle='-', color='#666666', linewidth=0.8),
                        bbox=dict(facecolor='white', alpha=0.7, edgecolor='none', pad=1)
                    )
                else:
                    # 普通标注
                    ax.text(center[0], center[1], f"{name}\n{int(row['人数'])}", 
                            fontsize=12, ha='center', va='center', fontweight='bold',
                            bbox=dict(facecolor='white', alpha=0.5, edgecolor='none', pad=1))

        ax.set_title(title, fontsize=20, pad=30)
        
        # 设置显示范围，聚焦中国陆地部分（排除南海诸岛过大导致的缩放问题）
        # 如果需要显示南海，可以去掉这两行
        ax.set_xlim(73, 135)
        ax.set_ylim(18, 54)
        
        ax.axis('off')
        
        fig.savefig(out_path, dpi=200, bbox_inches='tight')
        plt.close(fig)
    except Exception as e:
        # 如果地图绘制失败，记录错误但不中断主程序
        print(f"地图绘制详细错误: {str(e)}")
        traceback.print_exc()

# --- 统计逻辑函数 ---

def perform_chi_square(df: pd.DataFrame, x_col: str, y_col: str) -> Tuple[float, float, int]:
    """执行卡方检验"""
    if scipy_stats is None: return np.nan, np.nan, 0
    ct = pd.crosstab(df[x_col], df[y_col])
    if ct.shape[0] < 2 or ct.shape[1] < 2:
        return np.nan, np.nan, 0
    chi2, p, dof, _ = scipy_stats.chi2_contingency(ct.values)
    return chi2, p, dof

def get_stats_text(df: pd.DataFrame, edu_col: str):
    """获取基本学历分布统计"""
    stats = {}
    levels = ["本科", "硕士", "博士"]
    for lv in levels:
        stats[lv] = len(df[df[edu_col] == lv])
    stats["总计"] = len(df)
    return stats

def main_process(file_path: str, report_name: str, output_path: str):
    """主处理流程"""
    df_raw = read_excel(file_path)
    validate_columns(df_raw.columns)
    
    # 识别列
    edu_col = _pick_col(df_raw.columns, "学历")
    level_col = _pick_col(df_raw.columns, "层次")
    orient_col = _pick_col(df_raw.columns, "培养方式", "定向")
    source_col = _pick_col(df_raw.columns, "生源所在地", "生源地")
    outcome_col = _pick_col(df_raw.columns, "毕业生去向", "毕业去向")
    nature_col = _pick_col(df_raw.columns, "单位性质")
    loc_col = _pick_col(df_raw.columns, "单位所在地")
    unit_name_col = _pick_col(df_raw.columns, "单位名称")
    
    # 预处理
    df_all = df_raw.copy()
    
    # 统一学历分类 (对全量数据进行)
    df_all["_学历分类"] = df_all[edu_col].map(_classify_education)
    if (df_all["_学历分类"] == "").all() and level_col:
        df_all["_学历分类"] = df_all[level_col].map(_degree_to_level)
    
    # 创建临时图片目录
    temp_img_dir = tempfile.mkdtemp()
    
    # 在报告保存目录下创建 image 文件夹
    save_dir = os.path.dirname(output_path)
    image_save_dir = os.path.join(save_dir, "image")
    if not os.path.exists(image_save_dir):
        os.makedirs(image_save_dir)
    
    doc = Document()
    
    # 标题 - 使用用户输入的文件名作为报告标题
    title_text = f"{report_name}"
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run(title_text)
    run_title.bold = True
    run_title.font.size = Pt(22)
    
    add_heading(doc, "目录", level=1)
    add_table_of_contents(doc)
    # 插入分页符，让目录占据独立一页
    doc.add_page_break()
    
    # 一、毕业生构成
    add_heading(doc, "一、毕业生构成", level=1)
    add_heading(doc, "1. 毕业生学历", level=2)
    
    # 确保文字统计和图表统计的对象一致：仅统计能识别为 本/硕/博 的学生
    df_valid = df_all[df_all["_学历分类"].isin(["本科", "硕士", "博士"])].copy()
    stats_valid = get_stats_text(df_valid, "_学历分类")
    
    # 区分定向和非定向 (基于有效学历数据)
    non_dir_count = 0
    dir_count = 0
    if orient_col:
        non_dir_mask = df_valid[orient_col].astype(str).str.contains("非定向", na=False)
        non_dir_count = len(df_valid[non_dir_mask])
        dir_count = len(df_valid) - non_dir_count
    
    # 尝试从文件名提取年份
    import re
    year_match = re.search(r"\d{4}", report_name)
    display_year = year_match.group() if year_match else str(datetime.datetime.now().year)
    
    text1 = (f"心理学部 {display_year} 届毕业生共 {stats_valid['总计']} 人，"
             f"其中本科生 {stats_valid['本科']} 人，硕士研究生 {stats_valid['硕士']} 人，博士研究生 {stats_valid['博士']} 人。"
             f"其中非定向生 {non_dir_count} 人，定向生 {dir_count} 人。")
    add_paragraph(doc, text1)
    
    # 绘制本硕博定向/非定向柱状图 (传入过滤后的有效数据)
    if orient_col:
        edu_orient_img_path = os.path.join(image_save_dir, "毕业生学历构成柱状图.png")
        plot_education_orient_grouped_bar(df_valid, orient_col, "毕业生学历构成", edu_orient_img_path)
        add_picture(doc, edu_orient_img_path)
    
    # --- 后续分析仅针对非定向生 ---
    if orient_col:
        df = df_all[df_all[orient_col].astype(str).str.contains("非定向", na=False)].copy()
    else:
        df = df_all.copy()

    # 省份简写 (仅对非定向生)
    df["_生源省份"] = df[source_col].map(_map_location)
    df["_单位省份"] = df[loc_col].map(_map_location)
    
    # 二、毕业生就业结构
    add_heading(doc, "二、毕业生就业结构", level=1)
    
    # 1. 毕业生去向落实率
    add_heading(doc, "1. 毕业生去向落实率", level=2)
    add_paragraph(doc, "毕业生去向落实率统计如下：")
    
    # 计算落实率
    df["_去向状态"] = df[outcome_col].map(_map_outcome_simple)
    
    outcome_rows = []
    for lv in ["本科", "硕士", "博士"]:
        sub = df[df["_学历分类"] == lv]
        total = len(sub)
        if total == 0: continue
        settled = len(sub[sub["_去向状态"] == "落实"])
        rate = (settled / total * 100) if total > 0 else 0
        outcome_rows.append({"学历": lv, "人数": total, "已落实": settled, "落实率(%)": round(rate, 2)})
        
    table_outcome = doc.add_table(rows=len(outcome_rows)+1, cols=4)
    table_outcome.cell(0,0).text = "学历"
    table_outcome.cell(0,1).text = "人数"
    table_outcome.cell(0,2).text = "已落实"
    table_outcome.cell(0,3).text = "落实率(%)"
    for i, row in enumerate(outcome_rows):
        table_outcome.cell(i+1, 0).text = row["学历"]
        table_outcome.cell(i+1, 1).text = str(row["人数"])
        table_outcome.cell(i+1, 2).text = str(row["已落实"])
        table_outcome.cell(i+1, 3).text = str(row["落实率(%)"])
    _apply_three_line_table(table_outcome)

    # 绘制去向落实率柱状图
    outcome_img_path = os.path.join(image_save_dir, "毕业生去向落实率柱状图.png")
    plot_outcome_rate_grouped_bar(df, "毕业生去向落实情况（按落实/未落实）", outcome_img_path)
    add_picture(doc, outcome_img_path)
    
    # 2. 升学统计
    add_heading(doc, "2. 升学统计", level=2)
    add_paragraph(doc, "统计了本科毕业生国内升学和出国留学的具体单位（人数大于1），结果见下表。")
    df["_升学类型"] = df[outcome_col].map(_map_further_study)
    
    # 仅统计本科毕业生的升学数据
    study_df = df[(df["_学历分类"] == "本科") & (df["_升学类型"] != "")]
    
    if not study_df.empty:
        # 统计人数
        study_list = study_df[[unit_name_col, "_升学类型"]].value_counts().reset_index()
        study_list.columns = ["单位名称", "升学类型", "人数"]
        
        # 仅保留人数 > 0 的记录
        study_list = study_list[study_list["人数"] > 0]
        
        # 排序：升学类型（国内先，境外后），然后按人数降序
        # 定义排序权重：国内升学=0, 境外升学=1
        study_list["_type_sort"] = study_list["升学类型"].map({"国内升学": 0, "境外升学": 1})
        study_list = study_list.sort_values(["_type_sort", "人数"], ascending=[True, False])
        
        if not study_list.empty:
            table_study = doc.add_table(rows=len(study_list)+1, cols=3)
            # 设置表头
            table_study.cell(0,0).text = "升学类型"
            table_study.cell(0,1).text = "单位名称"
            table_study.cell(0,2).text = "人数"
            
            for i, (_, r) in enumerate(study_list.iterrows()):
                table_study.cell(i+1, 0).text = str(r["升学类型"])
                table_study.cell(i+1, 1).text = str(r["单位名称"])
                table_study.cell(i+1, 2).text = str(r["人数"])
            _apply_three_line_table(table_study)
        else:
            add_paragraph(doc, "暂无人数大于0的本科升学记录。")
    else:
        add_paragraph(doc, "暂无本科升学记录。")

    # 3. 就业单位性质统计
    add_heading(doc, "3. 就业单位性质统计", level=2)
    # 剔除升学数据（包括国内升学、境外升学、出国、留学等）
    df_nature = df[df[outcome_col].map(_map_further_study) == ""].copy()
    
    # 添加汇总文字
    tops_nature = get_top_category_text(df_nature, nature_col)
    text_nature = (f"心理学部 {display_year} 届毕业生的就业单位性质中，"
                   f"本科生最多的就业单位类型是 {tops_nature['本科']}，"
                   f"硕士研究生是 {tops_nature['硕士']}，"
                   f"博士研究生是 {tops_nature['博士']}。")
    add_paragraph(doc, text_nature)
    
    # 绘图
    nature_path = os.path.join(image_save_dir, "就业单位性质分布.png")
    plot_stacked_bar(df_nature, nature_col, "就业单位性质分布（按学历堆叠）", nature_path)
    add_picture(doc, nature_path)
    
    # 制表
    add_stats_table_by_education(doc, df_nature, nature_col, "单位性质")

    # 4. 就业单位统计
    add_heading(doc, "4. 就业单位统计", level=2)
    # 剔除升学数据（包括国内升学、境外升学、出国、留学等）进行单位统计
    df_unit = df[df[outcome_col].map(_map_further_study) == ""].copy()
    
    # 过滤掉单位名称为空的行
    df_unit = df_unit[df_unit[unit_name_col] != ""]
    
    # 统计就业人数大于1的单位名称
    unit_counts = df_unit[unit_name_col].value_counts().reset_index()
    unit_counts.columns = [unit_name_col, "人数"]
    unit_counts = unit_counts[unit_counts["人数"] > 1]
    
    if not unit_counts.empty:
        # 添加汇总文字
        add_paragraph(doc, f"心理学部 {display_year} 届毕业生中，共有 {len(unit_counts)} 个单位就业人数大于1人。")
        
        # 筛选出人数大于1的单位数据
        valid_units = unit_counts[unit_name_col].tolist()
        df_unit_filtered = df_unit[df_unit[unit_name_col].isin(valid_units)]
        
        # 制表：本硕博合并展示
        add_paragraph(doc, "就业单位详细统计（按学历分布）见下表：")
        add_stats_table_by_education(doc, df_unit_filtered, unit_name_col, "单位名称")
    else:
        add_paragraph(doc, "暂无就业人数大于1的单位记录。")

    # 5. 就业单位所在地
    add_heading(doc, "5. 就业单位所在地", level=2)
    # 剔除升学数据（包括国内升学、境外升学、出国、留学等）进行所在地统计
    df_loc = df[df[outcome_col].map(_map_further_study) == ""].copy()
    
    # 添加汇总文字
    tops_loc = get_top_category_text(df_loc, "_单位省份")
    text_loc = (f"心理学部 {display_year} 届毕业生的就业单位所在地中，"
                f"本科生最多的就业单位所在地是 {tops_loc['本科']}，"
                f"硕士研究生是 {tops_loc['硕士']}，"
                f"博士研究生是 {tops_loc['博士']}。")
    add_paragraph(doc, text_loc)
    
    # 绘图
    loc_path = os.path.join(image_save_dir, "就业单位所在地分布.png")
    plot_stacked_bar(df_loc, "_单位省份", "就业单位所在地分布（按学历堆叠）", loc_path)
    add_picture(doc, loc_path)
    
    # 新增：中国地图热力图
    if gpd is not None:
        map_path = os.path.join(image_save_dir, "就业单位所在地_中国地图.png")
        # 统计所有人的所在地分布
        loc_counts_all = df_loc["_单位省份"].value_counts().to_dict()
        plot_china_map_geojson(loc_counts_all, "毕业生就业单位所在地分布地图", map_path)
        if os.path.exists(map_path):
            add_paragraph(doc, "毕业生就业单位所在地分布地图如下：")
            add_picture(doc, map_path)
    
    # 制表
    add_stats_table_by_education(doc, df_loc, "_单位省份", "所在地")
    
    # 网格图 - 同样剔除升学数据
    add_heading(doc, f"5.1 生源地与就业单位所在地分布网格图", level=3)
    add_paragraph(doc, "统计了毕业生生源地和就业单位所在地（已剔除升学人员）。")
    for lv in ["本科", "硕士", "博士"]:
        sub = df_loc[(df_loc["_学历分类"] == lv) & (df_loc["_去向状态"] == "落实")]
        if sub.empty: continue
        heatmap_path = os.path.join(image_save_dir, f"生源地与就业单位所在地分布网格图_{lv}.png")
        plot_heatmap_grid(sub, "_生源省份", "_单位省份", f"生源地与就业单位所在地分布网格图（{lv}）", heatmap_path)
        if os.path.exists(heatmap_path):
            add_paragraph(doc, f"生源地与就业单位所在地分布网格图（{lv}）")
            add_picture(doc, heatmap_path)

    # 新增：核心经济区人数占比统计
    add_heading(doc, "5.2 核心经济区就业分布", level=3)
    add_paragraph(doc, "统计了核心经济区的就业分布，核心经济区的定义包括：北京、天津、河北属于“京津冀”，广东、香港、澳门属于“珠三角”，上海、江苏、浙江、安徽属于“长三角”，四川和重庆属于“成渝”。")
    
    economic_zones = {
        "京津冀": ["北京", "天津", "河北"],
        "珠三角": ["广东", "香港", "澳门"],
        "长三角": ["上海", "江苏", "浙江", "安徽"],
        "成渝": ["四川", "重庆"]
    }
    
    # 定义映射函数
    def map_to_zone(loc):
        for zone, provs in economic_zones.items():
            if loc in provs:
                return zone
        return "其他"

    df_loc["_经济区"] = df_loc["_单位省份"].map(map_to_zone)
    
    # 统计各学历在各经济区的人数
    zones_list = ["京津冀", "珠三角", "长三角", "成渝", "其他"]
    levels = ["本科", "硕士", "博士"]
    
    # 绘制经济区分布饼图
    zone_img_path = os.path.join(image_save_dir, "核心经济区就业分布饼图.png")
    plot_economic_zone_pie_charts(df_loc, "核心经济区就业分布（本硕博）", zone_img_path)
    add_picture(doc, zone_img_path)
    
    table_zone = doc.add_table(rows=len(zones_list) + 1, cols=len(levels) + 1)
    table_zone.cell(0, 0).text = "核心经济区"
    for j, lv in enumerate(levels):
        table_zone.cell(0, j + 1).text = lv
        
    for i, zone in enumerate(zones_list):
        table_zone.cell(i + 1, 0).text = zone
        for j, lv in enumerate(levels):
            # 该学历总人数 (已剔除升学)
            total_lv = len(df_loc[df_loc["_学历分类"] == lv])
            # 该学历在某经济区人数
            count_zone = len(df_loc[(df_loc["_学历分类"] == lv) & (df_loc["_经济区"] == zone)])
            
            if total_lv > 0:
                percent = (count_zone / total_lv) * 100
                table_zone.cell(i + 1, j + 1).text = f"{count_zone}人 ({percent:.2f}%)"
            else:
                table_zone.cell(i + 1, j + 1).text = "0人 (0.00%)"
                
    _apply_three_line_table(table_zone)

    # 5. 毕业生的就业偏向
    add_heading(doc, "6. 毕业生的就业偏向", level=2)
    
    # 使用剔除升学的数据进行偏向分析
    df_pref = df[~df[outcome_col].str.contains("研究生", na=False)].copy()
    
    # 5.1 科研倾向统计
    add_heading(doc, "6.1 科研倾向统计", level=3)
    add_paragraph(doc, "统计了毕业生的科研倾向分布（本硕博），科研行业的定义为高等教育单位和科研设计单位。")
    res_img_path = os.path.join(image_save_dir, "毕业生科研倾向饼图.png")
    plot_preference_pie_charts(df_pref, nature_col, _is_research_unit, 
                               "毕业生科研倾向分布（本硕博）", res_img_path, 
                               pref_label="科研类单位")
    add_picture(doc, res_img_path)
    add_preference_stats_table(doc, df_pref, nature_col, _is_research_unit, "科研类单位")

    # 5.2 教育行业倾向统计
    add_heading(doc, "6.2 教育行业倾向统计", level=3)
    add_paragraph(doc, "统计了毕业生的教育行业倾向分布（本硕博），教育行业的定义为中初教育单位。")
    edu_img_path = os.path.join(image_save_dir, "毕业生教育行业倾向饼图.png")
    plot_preference_pie_charts(df_pref, nature_col, _is_education_unit, 
                               "毕业生教育行业倾向分布（本硕博）", edu_img_path, 
                               pref_label="教育行业")
    add_picture(doc, edu_img_path)
    add_preference_stats_table(doc, df_pref, nature_col, _is_education_unit, "教育行业单位")

    # 三、毕业生就业影响因素
    add_heading(doc, "三、毕业生就业影响因素", level=1)
    add_heading(doc, "1. 是否回乡与选择“稳定工作”的偏好", level=2)
    
    add_paragraph(doc, "将教育单位（含高等教育单位），机关，国有企业，部队和其他事业单位划分为“稳定工作”，对选择回乡和不回乡毕业生进行卡方检验以判断是否存在显著差异。")

    df["_是否回乡"] = np.where(df["_生源省份"] == df["_单位省份"], "回乡", "不回乡")
    df["_是否稳定"] = df[nature_col].map(lambda x: "稳定" if _is_stable_work(x) else "非稳定")
    
    chi_df_all = df[df["_去向状态"] == "落实"]
    
    # 绘图：本硕博分开绘制分组堆叠柱状图 (展示回乡/不回乡组的稳定偏好)
    chi_img_path = os.path.join(image_save_dir, "回乡与稳定工作偏好分布图.png")
    plot_chi_square_stacked_bar(chi_df_all, "各学历回乡与稳定工作偏好分布", chi_img_path)
    add_picture(doc, chi_img_path)

    # 分本硕博进行卡方检验并制表
    for lv in ["本科", "硕士", "博士"]:
        sub_lv = chi_df_all[chi_df_all["_学历分类"] == lv]
        if sub_lv.empty: continue
        
        add_heading(doc, f"1.{['1','2','3'][['本科','硕士','博士'].index(lv)]} {lv}生卡方检验结果", level=3)
        
        chi2, p, dof = perform_chi_square(sub_lv, "_是否回乡", "_是否稳定")
        significance = "显著" if p < 0.05 else "不显著"
        
        # 统计表
        ct = pd.crosstab(sub_lv["_是否回乡"], sub_lv["_是否稳定"])
        table_chi = doc.add_table(rows=ct.shape[0]+2, cols=ct.shape[1]+2)
        
        # 填充表头
        table_chi.cell(0,0).text = f"{lv} (P={p:.4f})"
        for j, col in enumerate(ct.columns):
            table_chi.cell(0, j+1).text = str(col)
        table_chi.cell(0, ct.shape[1]+1).text = "结果"
        
        # 填充内容
        for i, idx in enumerate(ct.index):
            table_chi.cell(i+1, 0).text = str(idx)
            for j, col in enumerate(ct.columns):
                table_chi.cell(i+1, j+1).text = str(ct.loc[idx, col])
        
        # 合并最后一列显示结果
        res_cell = table_chi.cell(1, ct.shape[1]+1)
        res_cell.merge(table_chi.cell(ct.shape[0], ct.shape[1]+1))
        res_cell.text = significance
        
        _apply_three_line_table(table_chi)
        add_paragraph(doc, f"注：卡方值为 {chi2:.4f}，P 值为 {p:.4f}，差异{significance}。")

    # 保存文档
    doc.save(output_path)
    return output_path

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    
    messagebox.showinfo("提示", "欢迎使用就业质量报告生成器。")
    
    # 1. 选择输入文件
    file_path = filedialog.askopenfilename(title="第一步：选择 Excel 数据文件", filetypes=[("Excel files", "*.xlsx")])
    if not file_path:
        sys.exit()

    # 2. 选择保存位置并输入文件名
    default_year = datetime.datetime.now().year
    default_filename = f"心理学部 {default_year} 年就业质量报告"
    
    output_path = filedialog.asksaveasfilename(
        title="第二步：选择保存位置并输入文件名",
        initialfile=default_filename,
        defaultextension=".docx",
        filetypes=[("Word documents", "*.docx")]
    )
    
    if not output_path:
        sys.exit()
    
    # 从路径中提取文件名（不含扩展名）作为报告标题
    report_name = os.path.splitext(os.path.basename(output_path))[0]
        
    try:
        out_file = main_process(file_path, report_name, output_path)
        messagebox.showinfo("成功", f"报告生成完毕！\n保存路径：{os.path.abspath(out_file)}")
    except Exception as e:
        error_msg = traceback.format_exc()
        messagebox.showerror("错误", f"程序运行出错：\n{error_msg}")
