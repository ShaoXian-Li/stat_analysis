import os
import re
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from typing import List, Dict, Tuple
import tkinter as tk
from tkinter import filedialog, messagebox
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.shared import OxmlElement, qn

# 设置中文字体：按顺序取系统中第一个可用的字体，避免非 Windows 环境出现方框
plt.rcParams["font.sans-serif"] = [
    "SimHei",
    "Microsoft YaHei",
    "Noto Sans CJK SC",
    "WenQuanYi Zen Hei",
    "Arial Unicode MS",
]
plt.rcParams["axes.unicode_minus"] = False

# --- 常量定义 ---

ECONOMIC_ZONES = {
    "京津冀": ["北京", "天津", "河北"],
    "珠三角": ["广东", "香港", "澳门"],
    "长三角": ["上海", "江苏", "浙江", "安徽"],
    "成渝": ["四川", "重庆"]
}

# 985/211/双一流 列表简版 (仅供演示，实际可扩充或加载外部库)
UNIVERSITIES_985 = {
    "清华大学", "北京大学", "复旦大学", "上海交通大学", "浙江大学", "南京大学", "中国科学技术大学",
    "武汉大学", "中山大学", "华中科技大学", "哈尔滨工业大学", "四川大学", "西安交通大学", "山东大学",
    "中南大学", "同济大学", "东南大学", "北京航空航天大学", "北京理工大学", "南开大学", "天津大学",
    "华南理工大学", "厦门大学", "大连理工大学", "吉林大学", "湖南大学", "重庆大学", "西北工业大学",
    "华东师范大学", "中国农业大学", "电子科技大学", "兰州大学", "中国海洋大学", "东北大学", "大连理工大学",
    "北京师范大学", "中央民族大学", "国防科技大学", "西北农林科技大学"
}

# 这里简化处理，实际可以通过外部数据加载更完整的名单
UNIVERSITIES_211 = UNIVERSITIES_985 | {
    "北京交通大学", "北京工业大学", "北京科技大学", "北京化工大学", "北京邮电大学", "中国农业大学",
    "北京林业大学", "北京中医药大学", "北京外国语大学", "中国传媒大学", "对外经济贸易大学",
    "中国地质大学", "中国政法大学", "中国石油大学", "中国矿业大学", "中国地质大学", "中国人民公安大学",
    "河北工业大学", "太原理工大学", "内蒙古大学", "辽宁大学", "延边大学", "哈尔滨工程大学",
    "东北师范大学", "东北农业大学", "东北林业大学", "华东理工大学", "东华大学", "上海海洋大学",
    "上海大学", "苏州大学", "南京航空航天大学", "南京师范大学", "中国药科大学", "河海大学",
    "江南大学", "合肥工业大学", "安徽大学", "福州大学", "南昌大学", "中国石油大学", "郑州大学",
    "河南大学", "武汉理工大学", "中国地质大学", "华中师范大学", "华中农业大学", "湖南师范大学",
    "暨南大学", "华南师范大学", "广西大学", "海南大学", "四川农业大学", "西南交通大学", "西南大学",
    "中南财经政法大学", "贵州大学", "云南大学", "西藏大学", "西北大学", "西安电子科技大学",
    "长安大学", "宁夏大学", "青海大学", "新疆大学", "石河子大学"
}

# --- 基础工具函数 ---

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
    "毕业生去向": ("毕业生去向", "毕业去向"),
    "单位性质": ("单位性质",),
    "单位所在地": ("单位所在地",),
    "单位名称": ("单位名称",),
    "专业": ("专业", "学科专业"),
}


def validate_columns(columns, file_name: str = "") -> None:
    """校验表头是否齐全，缺失时抛出中文错误提示"""
    columns = list(columns)
    missing = [label for label, cands in USED_COLUMNS.items() if not _pick_col(columns, *cands)]
    if missing:
        prefix = "文件「{}」".format(file_name) if file_name else "数据表"
        raise ValueError(
            "{}缺少必需的列：{}。\n当前表头为：{}\n请对照《使用说明》核对表头（注意全角空格与错别字）。".format(
                prefix, "、".join(missing), "、".join(str(c) for c in columns)
            )
        )

def _classify_education(v: str) -> str:
    """学历归类"""
    s = str(v).strip()
    if "博士" in s or "博" in s: return "博士"
    if "硕士" in s or "硕" in s:
        if "学术" in s: return "学术硕士"
        if "专业" in s: return "专业硕士"
        return "硕士"
    if "本科" in s or "本" in s: return "本科"
    return ""

def _map_location(v: str) -> str:
    """省份前2个汉字"""
    s = str(v).strip()
    if not s or s.lower() == "nan": return ""
    if "国外" in s or "境外" in s or "出国" in s or "留学" in s: return "国外"
    return s[:2]

def _is_research_unit(nature: str) -> bool:
    """是否科研单位"""
    return "高等教育单位" in nature or "科研设计单位" in nature

def _map_further_study(v: str) -> str:
    """升学映射"""
    s = str(v).strip()
    if "境外" in s or "出国" in s or "留学" in s:
        return "境外升学"
    if "研究生" in s or "升学" in s:
        return "国内升学"
    return ""

def _classify_master_type(major: str, edu: str) -> str:
    """根据专业判断硕士类型：应用心理是专硕，其他是学硕"""
    major = str(major).strip()
    edu = str(edu).strip()
    if "硕士" in edu:
        if "应用心理" in major:
            return "专硕"
        return "学硕"
    return ""

def map_to_zone(loc):
    """映射到核心经济区"""
    for zone, provs in ECONOMIC_ZONES.items():
        if loc in provs:
            return zone
    return "其他"

def extract_year(filename: str) -> int:
    """从文件名提取年份"""
    match = re.search(r"(\d{4})", filename)
    return int(match.group(1)) if match else 0

# --- 统计逻辑 ---

def is_in_list(name, uni_set):
    """判断单位名称是否在集合中（支持包含关系）"""
    name = str(name)
    for uni in uni_set:
        if uni in name or name in uni:
            return True
    return False

def analyze_trends(file_paths: List[str]):
    """分析多年数据趋势"""
    results = []
    
    # 按年份排序
    file_info = []
    for p in file_paths:
        year = extract_year(os.path.basename(p))
        if year:
            file_info.append((year, p))
    file_info.sort()
    
    years = [info[0] for info in file_info]
    df_list = []
    for _, path in file_info:
        df_item = read_excel(path)
        validate_columns(df_item.columns, os.path.basename(path))
        df_list.append(df_item)
    
    # 预处理每一年的数据
    for i, df in enumerate(df_list):
        edu_col = _pick_col(df.columns, "学历")
        outcome_col = _pick_col(df.columns, "毕业生去向", "毕业去向")
        nature_col = _pick_col(df.columns, "单位性质")
        loc_col = _pick_col(df.columns, "单位所在地")
        name_col = _pick_col(df.columns, "单位名称")
        job_col = _pick_col(df.columns, "工作职位类别")
        major_col = _pick_col(df.columns, "专业", "学科专业")
        
        df["_学历分类"] = df[edu_col].map(_classify_education)
        df["_升学类型"] = df[outcome_col].map(_map_further_study)
        df["_单位省份"] = df[loc_col].map(_map_location)
        df["_经济区"] = df["_单位省份"].map(map_to_zone)
        df["_硕士类型"] = df.apply(lambda row: _classify_master_type(row[major_col], row[edu_col]), axis=1)
        
        # 博士去向分类
        def classify_phd_dest(row):
            name = str(row[name_col])
            nature = str(row[nature_col])
            study = str(row["_升学类型"])
            
            if study == "境外升学": return "国外深造"
            if _is_research_unit(nature):
                if name in UNIVERSITIES_985: return "985"
                if name in UNIVERSITIES_211: return "211"
                # 双一流判定简化：211基本覆盖了大部分双一流，此处作为演示
                return "其他科研/学术"
            return "非学术类就业"

        df["_博士去向"] = df.apply(classify_phd_dest, axis=1)
    
    # 1. 博士趋势（占比）
    phd_trends = {
        "科研占比": [],
        "企业占比": [],
        "其他占比": []
    }
    
    for df in df_list:
        phd_df = df[df["_学历分类"] == "博士"]
        total_phd = len(phd_df)
        if total_phd == 0:
            for k in phd_trends: phd_trends[k].append(0)
            continue
            
        nature_col = _pick_col(df.columns, "单位性质")
        
        research_ratio = len(phd_df[phd_df[nature_col].apply(_is_research_unit)]) / total_phd * 100
        enterprise_ratio = len(phd_df[phd_df[nature_col].str.contains("企业", na=False)]) / total_phd * 100
        
        phd_trends["科研占比"].append(research_ratio)
        phd_trends["企业占比"].append(enterprise_ratio)
        phd_trends["其他占比"].append(100 - research_ratio - enterprise_ratio)

    # 2. 经济区趋势 (按学历分层，占比) - 排除科研助理、研究生、境外留学
    zone_trends = {
        "本科": {zone: [] for zone in ["京津冀", "珠三角", "长三角", "成渝", "其他"]},
        "硕士": {zone: [] for zone in ["京津冀", "珠三角", "长三角", "成渝", "其他"]},
        "博士": {zone: [] for zone in ["京津冀", "珠三角", "长三角", "成渝", "其他"]}
    }
    for df in df_list:
        outcome_col = _pick_col(df.columns, "毕业生去向", "毕业去向")
        
        for edu in ["本科", "硕士", "博士"]:
            edu_df = df[df["_学历分类"] == edu]
            valid_df = edu_df[
                (edu_df["_经济区"] != "") & 
                (~edu_df[outcome_col].str.contains("拟升学|待就业", na=False)) &
                (~edu_df[outcome_col].str.contains("科研助理|研究生|境外留学", na=False))
            ]
            total = len(valid_df)
            for zone in zone_trends[edu]:
                count = len(valid_df[valid_df["_经济区"] == zone])
                zone_trends[edu][zone].append(count / total * 100 if total > 0 else 0)

    # 3. 本科生趋势（占比）- 排除公费师范生
    ug_trends = {"国内升学占比": [], "出国占比": [], "直接就业占比": []}
    for df in df_list:
        ug_df = df[df["_学历分类"] == "本科"]
        
        teacher_col = _pick_col(df.columns, "师范生类别")
        if teacher_col:
            ug_df = ug_df[~ug_df[teacher_col].str.contains("公费师范生", na=False)]
        
        total = len(ug_df)
        if total == 0:
            for k in ug_trends: ug_trends[k].append(0)
            continue
        ug_trends["国内升学占比"].append(len(ug_df[ug_df["_升学类型"] == "国内升学"]) / total * 100)
        ug_trends["出国占比"].append(len(ug_df[ug_df["_升学类型"] == "境外升学"]) / total * 100)
        # 直接就业：落实且非升学
        outcome_col = _pick_col(df.columns, "毕业生去向", "毕业去向")
        direct = ug_df[(ug_df["_升学类型"] == "") & (~ug_df[outcome_col].str.contains("拟升学|待就业|暂不就业", na=False))]
        ug_trends["直接就业占比"].append(len(direct) / total * 100)

    # 4. 毕业生去向趋势（本科、硕士-应用心理、硕士-心理学、博士，占比）
    grad_trends = {
        "本科-国内升学占比": [], "本科-出国占比": [], "本科-直接就业占比": [],
        "硕士-应用心理-国内升学占比": [], "硕士-应用心理-出国占比": [], "硕士-应用心理-直接就业占比": [],
        "硕士-心理学-国内升学占比": [], "硕士-心理学-出国占比": [], "硕士-心理学-直接就业占比": [],
        "博士-国内升学占比": [], "博士-出国占比": [], "博士-直接就业占比": []
    }
    
    # 5. 科研助理占比趋势（本科、硕士、博士）
    ra_trends = {"本科-科研助理占比": [], "硕士-科研助理占比": [], "博士-科研助理占比": []}
    
    # 6. 学硕升学比例趋势（应用心理是专硕，其他是学硕）
    ms_trends = {"学硕升学比例": []}
    
    for df in df_list:
        outcome_col = _pick_col(df.columns, "毕业生去向", "毕业去向")
        major_col = _pick_col(df.columns, "专业", "学科专业")
        loc_col = _pick_col(df.columns, "单位所在地")
        
        # 本科 - 排除公费师范生
        ug_df = df[df["_学历分类"] == "本科"]
        teacher_col = _pick_col(df.columns, "师范生类别")
        if teacher_col:
            ug_df = ug_df[~ug_df[teacher_col].str.contains("公费师范生", na=False)]
        ug_total = len(ug_df)
        
        ra_count = len(ug_df[ug_df[outcome_col].str.contains("科研助理", na=False)])
        ra_trends["本科-科研助理占比"].append(ra_count / ug_total * 100 if ug_total > 0 else 0)
        grad_trends["本科-国内升学占比"].append(len(ug_df[ug_df["_升学类型"] == "国内升学"]) / ug_total * 100 if ug_total > 0 else 0)
        grad_trends["本科-出国占比"].append(len(ug_df[ug_df["_升学类型"] == "境外升学"]) / ug_total * 100 if ug_total > 0 else 0)
        direct = ug_df[(ug_df["_升学类型"] == "") & (~ug_df[outcome_col].str.contains("拟升学|待就业|暂不就业", na=False))]
        grad_trends["本科-直接就业占比"].append(len(direct) / ug_total * 100 if ug_total > 0 else 0)
        
        # 硕士 - 按专业分开统计
        master_df = df[df["_学历分类"] == "硕士"]
        master_total = len(master_df)
        
        ra_count = len(master_df[master_df[outcome_col].str.contains("科研助理", na=False)])
        ra_trends["硕士-科研助理占比"].append(ra_count / master_total * 100 if master_total > 0 else 0)
        
        # 应用心理
        app_psy_df = master_df[master_df[major_col].str.contains("应用心理", na=False)]
        app_psy_total = len(app_psy_df)
        grad_trends["硕士-应用心理-国内升学占比"].append(len(app_psy_df[app_psy_df["_升学类型"] == "国内升学"]) / app_psy_total * 100 if app_psy_total > 0 else 0)
        grad_trends["硕士-应用心理-出国占比"].append(len(app_psy_df[app_psy_df["_升学类型"] == "境外升学"]) / app_psy_total * 100 if app_psy_total > 0 else 0)
        direct = app_psy_df[(app_psy_df["_升学类型"] == "") & (~app_psy_df[outcome_col].str.contains("拟升学|待就业|暂不就业", na=False))]
        grad_trends["硕士-应用心理-直接就业占比"].append(len(direct) / app_psy_total * 100 if app_psy_total > 0 else 0)
        
        # 心理学（排除应用心理）
        psy_df = master_df[master_df[major_col].str.contains("心理学", na=False) & ~master_df[major_col].str.contains("应用心理", na=False)]
        psy_total = len(psy_df)
        grad_trends["硕士-心理学-国内升学占比"].append(len(psy_df[psy_df["_升学类型"] == "国内升学"]) / psy_total * 100 if psy_total > 0 else 0)
        grad_trends["硕士-心理学-出国占比"].append(len(psy_df[psy_df["_升学类型"] == "境外升学"]) / psy_total * 100 if psy_total > 0 else 0)
        direct = psy_df[(psy_df["_升学类型"] == "") & (~psy_df[outcome_col].str.contains("拟升学|待就业|暂不就业", na=False))]
        grad_trends["硕士-心理学-直接就业占比"].append(len(direct) / psy_total * 100 if psy_total > 0 else 0)
        
        # 博士 - 境外留学属于出国
        phd_df = df[df["_学历分类"] == "博士"]
        phd_total = len(phd_df)
        
        ra_count = len(phd_df[phd_df[outcome_col].str.contains("科研助理", na=False)])
        ra_trends["博士-科研助理占比"].append(ra_count / phd_total * 100 if phd_total > 0 else 0)
        
        grad_trends["博士-国内升学占比"].append(len(phd_df[phd_df["_升学类型"] == "国内升学"]) / phd_total * 100 if phd_total > 0 else 0)
        # 出国包含境外升学、境外留学以及单位所在地为出国/国外/境外
        phd_abroad = phd_df[(phd_df["_升学类型"] == "境外升学") | (phd_df[outcome_col].str.contains("境外留学", na=False)) | (phd_df[loc_col].str.contains("出国|国外|境外", na=False))]
        grad_trends["博士-出国占比"].append(len(phd_abroad) / phd_total * 100 if phd_total > 0 else 0)
        direct = phd_df[(phd_df["_升学类型"] == "") & (~phd_df[outcome_col].str.contains("拟升学|待就业|暂不就业|境外留学", na=False))]
        grad_trends["博士-直接就业占比"].append(len(direct) / phd_total * 100 if phd_total > 0 else 0)
        
        # 学硕升学比例
        master_df = df[df["_硕士类型"] == "学硕"]
        ms_total = len(master_df)
        if ms_total == 0:
            ms_trends["学硕升学比例"].append(0)
        else:
            ms_further = len(master_df[master_df["_升学类型"] != ""])
            ms_trends["学硕升学比例"].append(ms_further / ms_total * 100)

    return years, phd_trends, zone_trends, grad_trends, ra_trends, ms_trends

# --- 绘图函数 ---

def plot_trend_line(years, data_dict, title, ylabel, out_path, is_percent=False):
    """绘制趋势折线图"""
    fig, ax = plt.subplots(figsize=(12, 7))
    
    for label, values in data_dict.items():
        ax.plot(years, values, marker='o', linewidth=3, markersize=10, label=label)
        # 添加数值标注
        for x, y in zip(years, values):
            text = f"{y:.1f}%" if is_percent else f"{int(y)}"
            ax.text(x, y, text, ha='center', va='bottom', fontsize=13, fontweight='bold')

    ax.set_title(title, fontsize=22, pad=25, fontweight='bold')
    ax.set_ylabel(ylabel, fontsize=18)
    ax.set_xlabel("年份", fontsize=18)
    ax.set_xticks(years)
    ax.tick_params(axis='both', labelsize=16)
    ax.legend(fontsize=15)
    ax.grid(True, linestyle='--', alpha=0.6)
    
    fig.tight_layout()
    fig.savefig(out_path, dpi=200, bbox_inches='tight')
    plt.close(fig)

# --- 主流程 ---

def main():
    root = tk.Tk()
    root.withdraw()
    
    messagebox.showinfo("提示", "请选择多年的就业数据Excel文件（文件名需包含年份）。")
    file_paths = filedialog.askopenfilenames(title="选择多个Excel文件", filetypes=[("Excel files", "*.xlsx")])
    
    if not file_paths:
        return

    messagebox.showinfo("提示", "请选择保存目录")
    output_dir = filedialog.askdirectory(title="选择结果保存目录")
    if not output_dir:
        return

    image_dir = os.path.join(output_dir, "image_trends")
    if not os.path.exists(image_dir):
        os.makedirs(image_dir)

    try:
        years, phd, zones, grad, ra, ms = analyze_trends(file_paths)
        
        # 绘图
        paths = {}
        
        # (1) 博士趋势 - 科研 vs 企业 vs 其他（占比）
        p1 = os.path.join(image_dir, "博士去向趋势_科研与企业.png")
        plot_trend_line(years, {"科研占比": phd["科研占比"], "企业占比": phd["企业占比"], "其他占比": phd["其他占比"]}, "博士毕业生选择科研、企业与其他趋势", "占比 (%)", p1, is_percent=True)
        paths["phd_res_ent"] = p1
        
        # (2) 本科经济区趋势
        p2 = os.path.join(image_dir, "本科各经济区就业占比趋势.png")
        plot_trend_line(years, zones["本科"], "本科毕业生各经济区就业占比趋势", "占比 (%)", p2, is_percent=True)
        paths["zone_ug"] = p2
        
        # (3) 硕士经济区趋势
        p3 = os.path.join(image_dir, "硕士各经济区就业占比趋势.png")
        plot_trend_line(years, zones["硕士"], "硕士毕业生各经济区就业占比趋势", "占比 (%)", p3, is_percent=True)
        paths["zone_master"] = p3
        
        # (4) 博士经济区趋势
        p4 = os.path.join(image_dir, "博士各经济区就业占比趋势.png")
        plot_trend_line(years, zones["博士"], "博士毕业生各经济区就业占比趋势", "占比 (%)", p4, is_percent=True)
        paths["zone_phd"] = p4
        
        # (5) 本科毕业生去向趋势（占比）- 排除公费师范生
        p5 = os.path.join(image_dir, "本科毕业生去向趋势.png")
        ug_data = {k: grad[k] for k in grad if k.startswith("本科")}
        plot_trend_line(years, ug_data, "本科毕业生去向变化趋势", "占比 (%)", p5, is_percent=True)
        paths["ug"] = p5
        
        # (6) 硕士毕业生去向趋势 - 应用心理（占比）
        p6 = os.path.join(image_dir, "硕士-应用心理去向趋势.png")
        app_psy_data = {k: grad[k] for k in grad if "应用心理" in k}
        plot_trend_line(years, app_psy_data, "硕士-应用心理专业去向变化趋势", "占比 (%)", p6, is_percent=True)
        paths["master_app_psy"] = p6
        
        # (7) 硕士毕业生去向趋势 - 心理学（占比）
        p7 = os.path.join(image_dir, "硕士-心理学去向趋势.png")
        psy_data = {k: grad[k] for k in grad if "心理学" in k and "应用心理" not in k}
        plot_trend_line(years, psy_data, "硕士-心理学专业去向变化趋势", "占比 (%)", p7, is_percent=True)
        paths["master_psy"] = p7
        
        # (8) 博士毕业生去向趋势（占比）- 境外留学属于出国
        p8 = os.path.join(image_dir, "博士毕业生去向趋势.png")
        phd_data = {k: grad[k] for k in grad if k.startswith("博士")}
        plot_trend_line(years, phd_data, "博士毕业生去向变化趋势", "占比 (%)", p8, is_percent=True)
        paths["phd_grad"] = p8
        
        # (9) 科研助理占比趋势
        p9 = os.path.join(image_dir, "科研助理占比趋势.png")
        plot_trend_line(years, ra, "本硕博毕业生科研助理占比变化趋势", "占比 (%)", p9, is_percent=True)
        paths["ra_ratio"] = p9
        
        # (10) 学硕升学比例趋势
        p10 = os.path.join(image_dir, "学硕升学比例趋势.png")
        plot_trend_line(years, ms, "学术型硕士升学比例变化趋势", "占比 (%)", p10, is_percent=True)
        paths["ms_ratio"] = p10

        # 生成Word报告
        doc = Document()
        doc.add_heading("毕业生就业变化趋势分析报告", 0)
        
        doc.add_heading("一、博士毕业生就业趋势", level=1)
        doc.add_paragraph("下图展示了博士毕业生在科研单位、企业及其他去向之间的选择趋势。其中，科研单位包括高校、科研院所等学术机构；企业包括各类企业单位；其他包括出国深造、政府事业单位及未就业等情况。")
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.add_picture(paths["phd_res_ent"], width=Inches(6))
        
        doc.add_heading("二、核心经济区流向趋势", level=1)
        doc.add_paragraph("下图展示了本科毕业生在各核心经济区（京津冀、长三角、珠三角、成渝等）的就业占比变化。")
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.add_picture(paths["zone_ug"], width=Inches(6))
        
        doc.add_paragraph("下图展示了硕士毕业生在各核心经济区（京津冀、长三角、珠三角、成渝等）的就业占比变化。")
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.add_picture(paths["zone_master"], width=Inches(6))
        
        doc.add_paragraph("下图展示了博士毕业生在各核心经济区（京津冀、长三角、珠三角、成渝等）的就业占比变化。")
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.add_picture(paths["zone_phd"], width=Inches(6))
        
        doc.add_heading("三、毕业生去向趋势", level=1)
        doc.add_paragraph("下图展示了本科毕业生国内升学、出国留学及直接就业的变化趋势（占比，排除公费师范生）。")
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.add_picture(paths["ug"], width=Inches(6))
        
        doc.add_paragraph("下图展示了硕士-应用心理专业毕业生国内升学、出国留学及直接就业的变化趋势（占比）。")
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.add_picture(paths["master_app_psy"], width=Inches(6))
        
        doc.add_paragraph("下图展示了硕士-心理学专业毕业生国内升学、出国留学及直接就业的变化趋势（占比）。")
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.add_picture(paths["master_psy"], width=Inches(6))
        
        doc.add_paragraph("下图展示了博士毕业生国内升学、出国留学及直接就业的变化趋势（占比）。")
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.add_picture(paths["phd_grad"], width=Inches(6))
        
        doc.add_heading("四、科研助理占比趋势", level=1)
        doc.add_paragraph("下图展示了本科、硕士、博士毕业生选择科研助理岗位的占比变化趋势。")
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.add_picture(paths["ra_ratio"], width=Inches(6))
        
        doc.add_heading("五、学术型硕士升学比例趋势", level=1)
        doc.add_paragraph("下图展示了学术型硕士（不含应用心理专硕）毕业生继续深造（升学）的比例变化趋势。")
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.add_picture(paths["ms_ratio"], width=Inches(6))
        
        doc_path = os.path.join(output_dir, "毕业生就业趋势分析报告.docx")
        doc.save(doc_path)
        
        messagebox.showinfo("成功", f"趋势分析报告已生成！\n保存路径：{doc_path}")
        
    except Exception as e:
        messageback = traceback.format_exc()
        messagebox.showerror("错误", f"分析过程中出现错误：\n{messageback}")

if __name__ == "__main__":
    import traceback
    main()
