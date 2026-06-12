#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
党员微信群成员核对工具
======================

功能：对比 Excel 党员名单 与 微信群成员截图OCR文本，
      找出未进群或备注不规范的党员。

用法：
    python party_wechat_check.py --excel "名单.xlsx" --wechat-file "members.txt"
    python party_wechat_check.py                        # 交互式提示输入
    python party_wechat_check.py --excel "名单.xlsx" --wechat-file "members.txt" --output "结果.xlsx"
"""

__version__ = "1.1.0"

# ============================================================================
# Section 1: Imports and Constants
# ============================================================================
import argparse
import io
import re
import sys
from pathlib import Path
from typing import List, Optional, Tuple

# Windows 控制台 UTF-8 支持
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        try:
            sys.stdout = io.TextIOWrapper(
                sys.stdout.buffer, encoding='utf-8', errors='replace'
            )
        except Exception:
            pass
    try:
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        try:
            sys.stderr = io.TextIOWrapper(
                sys.stderr.buffer, encoding='utf-8', errors='replace'
            )
        except Exception:
            pass

import pandas as pd
from rapidfuzz import fuzz

# --- Rich (optional) ---
try:
    from rich.console import Console
    from rich.table import Table
    from rich.panel import Panel
    from rich import box
    RICH_AVAILABLE = True
    console = Console()
except ImportError:
    RICH_AVAILABLE = False
    console = None

# --- Constants ---
# 复姓集合（用于姓名提取时避免拆分复姓）
COMPOUND_SURNAMES = [
    '欧阳', '司马', '上官', '诸葛', '夏侯', '皇甫', '尉迟',
    '公羊', '慕容', '公孙', '令狐', '独孤', '长孙', '宇文',
    '东方', '赫连', '澹台', '申屠', '闻人', '闾丘', '太叔',
    '端木', '鲜于', '南门', '左丘', '百里', '东郭', '拓跋',
]

# Excel 姓名列关键词（按优先级排序）
NAME_COLUMN_KEYWORDS = [
    '姓名', '名字', '党员姓名', '学生姓名', '人员姓名',
    '名称', '党员', '成员', '联系人', 'name',
]

# 微信群文本中需要过滤的行关键词
WECHAT_SKIP_KEYWORDS = [
    '群成员', '群聊成员', '群聊', '群主', '群管理员', '管理员',
    '添加成员', '删除成员', '加入群聊', '修改群名',
    '置顶', '免打扰', '群公告',
    '搜索', '邀请', '添加', '投诉',
    '确定', '取消', '完成', '返回', '更多',
]

# 微信群昵称中常见的装饰性标签（正则模式）
WECHAT_DECORATION_PATTERNS = [
    re.compile(r'[\[【\(（].*?[\]】\)）]'),      # 括号内容：备注、部门等
    re.compile(r'[-—\-].*$'),                     # 横线后缀
    re.compile(r'[☀-➿]'),               # 杂项符号
    re.compile(r'\d{4,}'),                        # 长数字（可能是学号/工号）
    re.compile(r'[℀-⯿︀-﻿]'),  # 更多符号
]

# emoji 范围正则（BMP杂项符号 + SMP emoji）
# 注意：范围不能跨越 BMP 到 SMP，会误删中日韩文字！
EMOJI_PATTERN = re.compile(
    "[\U0001F600-\U0001F64F"   # Emoticons (SMP)
    "\U0001F300-\U0001F5FF"    # Misc Symbols & Pictographs (SMP)
    "\U0001F680-\U0001F6FF"    # Transport & Map (SMP)
    "\U0001F1E0-\U0001F1FF"    # Flags (SMP)
    "\U0001F900-\U0001F9FF"    # Supplemental Symbols (SMP)
    "\U0001FA00-\U0001FA6F"    # Chess Symbols (SMP)
    "\U0001FA70-\U0001FAFF"    # Symbols Extended-A (SMP)
    "\U0001F000-\U0001F02F"    # Mahjong Tiles (SMP)
    "\U0001F0A0-\U0001F0FF"    # Playing Cards (SMP)
    "\U0001F7E0-\U0001F7FF"    # Geometric Shapes Extended (SMP)
    "\U00002702-\U000027B0"    # Dingbats (BMP)
    "\U00002500-\U000027BF"    # Box Drawing & Geometric Shapes (BMP)
    "\U000024C2-\U000024C3"    # Circled M (BMP，仅单字符)
    "]+", flags=re.UNICODE
)

# 默认模糊匹配阈值
DEFAULT_THRESHOLD = 85


# ============================================================================
# Section 2: Text Normalization
# ============================================================================

def normalize_name(name: str) -> str:
    """规范化姓名：去空格、emoji、特殊字符、装饰标签。

    Args:
        name: 原始姓名（来自 Excel 或微信）

    Returns:
        规范化后的纯文本姓名
    """
    if not name:
        return ""
    if isinstance(name, float) and (name != name):  # NaN check (NaN != NaN)
        return ""

    name = str(name).strip()

    # 去除 emoji
    name = EMOJI_PATTERN.sub('', name)

    # 去除微信装饰标签
    for pattern in WECHAT_DECORATION_PATTERNS:
        name = pattern.sub('', name)

    # 去除所有空格（中文姓名不应有空格）
    name = re.sub(r'\s+', '', name)

    # 去除残留的特殊符号
    name = name.strip('-_=+*.,;:!?@#$%^&*()[]{}|\\/"\'`~❤️🔥⭐✨💯')

    return name


def extract_chinese_name_candidates(text: str) -> List[str]:
    """从一段文本中提取所有可能的中文姓名片段（2-4字连续中文）。

    能识别复姓（如欧阳、司马），避免拆分复姓产生错误候选。
    例如 "欧阳泽坪" 会提取 "欧阳泽坪"、"欧阳泽"、"欧阳"，但不会提取 "阳泽坪"。

    Args:
        text: 规范化后的微信昵称

    Returns:
        可能的姓名字符串列表，按长度降序
    """
    # 匹配连续中文字符
    chinese_chars = re.findall(r'[一-鿿]+', text)
    candidates = []
    for chunk in chinese_chars:
        n = len(chunk)
        # 定位所有复姓在 chunk 中的内部位置（避免拆分复姓）
        surname_inner_positions: set = set()
        for cs in COMPOUND_SURNAMES:
            pos = 0
            while True:
                pos = chunk.find(cs, pos)
                if pos == -1:
                    break
                # 复姓内部（第二个字及之后）不能作为候选起点
                for inner in range(pos + 1, pos + len(cs)):
                    surname_inner_positions.add(inner)
                pos += 1

        # 提取 2-4 字片段（滑动窗口），跳过会拆分复姓的起点
        for length in range(min(n, 4), 1, -1):
            for start in range(n - length + 1):
                if start in surname_inner_positions:
                    continue
                sub = chunk[start:start + length]
                if sub not in candidates:
                    candidates.append(sub)
    # 长度优先
    candidates.sort(key=lambda s: -len(s))
    return candidates


# Pre-compiled regexes for strip_common_prefixes
_STRIP_PREFIX_PATTERNS = [
    re.compile(r'^(硕士|博士)?\d{2,4}[级]?-?'),
    re.compile(r'^[A-Za-z]+\d*-?'),
    re.compile(r'^[本硕博]\d{4}-?'),
]


def strip_common_prefixes(name: str) -> str:
    """去除微信群常见前缀（年级、班级编号等）。"""
    for pat in _STRIP_PREFIX_PATTERNS:
        name = pat.sub('', name)
    return name


# ============================================================================
# Section 3: Excel Reading
# ============================================================================

def detect_name_column(df: pd.DataFrame) -> Optional[str]:
    """自动检测 DataFrame 中的姓名列。

    Args:
        df: 从 Excel 读取的 DataFrame

    Returns:
        检测到的列名，若找不到返回 None
    """
    columns = df.columns.tolist()

    # 精确匹配
    for col in columns:
        col_str = str(col).strip()
        for keyword in NAME_COLUMN_KEYWORDS:
            if col_str == keyword:
                return col

    # 包含匹配
    for col in columns:
        col_str = str(col).strip()
        for keyword in NAME_COLUMN_KEYWORDS:
            if keyword in col_str:
                return col

    return None


def read_party_names(
    excel_path: str,
    column_name: Optional[str] = None,
) -> Tuple[List[str], str]:
    """从 Excel 读取党员姓名列表。

    Args:
        excel_path: Excel 文件路径
        column_name: 手动指定姓名列名（None 则自动检测）

    Returns:
        (姓名列表, 使用的列名)

    Raises:
        FileNotFoundError: 文件不存在
        ValueError: 无法检测姓名列
    """
    path = Path(excel_path).expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(f"Excel 文件不存在: {path}")

    df = pd.read_excel(path)

    if column_name:
        if column_name not in df.columns:
            raise ValueError(
                f"找不到指定的列 '{column_name}'。"
                f"可用的列: {', '.join(df.columns.tolist())}"
            )
    else:
        column_name = detect_name_column(df)
        if column_name is None:
            raise ValueError(
                f"无法自动检测姓名列，请用 --name-column 指定。"
                f"可用的列: {', '.join(df.columns.tolist())}"
            )

    # 提取姓名并清洗
    raw_names = df[column_name].tolist()
    names = []
    for n in raw_names:
        cleaned = normalize_name(n)
        if cleaned and len(cleaned) >= 2:
            names.append(cleaned)

    return names, column_name


# ============================================================================
# Section 4: WeChat Text Parsing
# ============================================================================

def parse_wechat_text(text: str) -> List[str]:
    """解析微信群成员提取文本，返回昵称列表。

    处理 OCR 特有的问题：
    - 一行可能有多个名字连在一起（如 "朱永洁刘莹黄晓君"）
    - 名字间夹杂装饰文本（如 "黄辉虎23化....欧阳泽...戴雅琴"）
    - 含 UI 文本行（如 "搜索群成员"、"添加"）

    Args:
        text: 微信"提取文字"得到的原始文本

    Returns:
        清理后的微信昵称列表（包含拆分出的个别名字和原始行）
    """
    lines = text.strip().split('\n')
    members = set()  # 用集合去重

    for line in lines:
        line_orig = line.strip()
        if not line_orig:
            continue

        # 过滤无关行
        if any(kw in line_orig for kw in WECHAT_SKIP_KEYWORDS):
            continue

        # 过滤纯数字行、纯符号行
        if re.match(r'^[\d\s,.，。、;:：;!！?？]+$', line_orig):
            continue

        # 过滤纯英文+符号的单字符行
        if re.match(r'^[a-zA-Z0-9\W_]$', line_orig):
            continue

        clean_line = normalize_name(line_orig)

        # 提取该行中所有中文名片段（2-4 字连续中文）
        candidates = extract_chinese_name_candidates(clean_line)

        if candidates:
            # 添加所有候选名
            for cand in candidates:
                members.add(cand)
            # 同时保留完整清洗行（有些情况下整行就是一个名字+标签）
            if len(clean_line) >= 2:
                members.add(clean_line)
        elif clean_line and len(clean_line) >= 2:
            # 无可提取的中文名片段，直接用整行
            members.add(clean_line)

    # 去重后返回列表
    return sorted(members, key=lambda x: -len(x))


def read_wechat_members(file_path: str) -> List[str]:
    """从文本文件读取微信群成员列表。

    Args:
        file_path: 文本文件路径

    Returns:
        微信群成员昵称列表（原始格式）
    """
    path = Path(file_path).expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(f"微信群文本文件不存在: {path}")

    with open(path, 'r', encoding='utf-8') as f:
        text = f.read()

    return parse_wechat_text(text)


# ============================================================================
# Section 5: Matching Engine
# ============================================================================

def match_single_forward(
    excel_name: str,
    wechat_names: List[str],
    wechat_clean: List[str],
    threshold: int = DEFAULT_THRESHOLD,
) -> Tuple[Optional[str], Optional[str], float, str]:
    """正向匹配：Excel姓名 -> 微信群昵称。

    Args:
        excel_name: Excel 中的党员姓名（已规范化）
        wechat_names: 所有微信昵称（原始格式）
        wechat_clean: 所有微信昵称（已规范化，与 wechat_names 一一对应）
        threshold: 模糊匹配阈值 (0-100)

    Returns:
        (matched_wechat_name, matched_wechat_raw, confidence, method)
        未匹配则返回 (None, None, 0, '')
    """
    if not excel_name:
        return None, None, 0, ''

    best_score = 0
    best_match = None
    best_raw = None
    best_method = ''

    for w_raw, w_clean in zip(wechat_names, wechat_clean):
        # 方法 1: 直接包含（Excel姓名在微信昵称中）
        if excel_name in w_clean:
            best_score = 100
            best_match = w_clean
            best_raw = w_raw
            best_method = 'direct_contains'
            break  # perfect match, no need to continue
        # 方法 2: 微信昵称被Excel姓名包含（如备注格式相反）
        if w_clean and w_clean in excel_name:
            score = 100
            method = 'reverse_contains'
        # 方法 3: 去前缀后再尝试包含
        elif excel_name in strip_common_prefixes(w_clean):
            score = 95
            method = 'prefix_stripped_contains'
        else:
            # 方法 4: rapidfuzz 部分匹配
            score = fuzz.partial_ratio(excel_name, w_clean)
            method = 'fuzzy_partial'

        if score > best_score:
            best_score = score
            best_match = w_clean
            best_raw = w_raw
            best_method = method

    if best_score >= threshold and best_match:
        return best_match, best_raw, best_score, best_method
    elif best_score >= threshold - 10:
        return best_match, best_raw, best_score, 'uncertain_' + best_method
    else:
        return None, None, best_score, ''


def match_single_reverse(
    wechat_name: str,
    unmatched_excel: List[str],
    threshold: int = DEFAULT_THRESHOLD,
) -> Tuple[Optional[str], float]:
    """反向匹配：从微信昵称提取中文名片段，尝试匹配 Excel 名单。

    Args:
        wechat_name: 微信昵称（已规范化）
        unmatched_excel: 尚未匹配的 Excel 姓名列表

    Returns:
        (matched_excel_name, confidence) or (None, 0)
    """
    candidates = extract_chinese_name_candidates(wechat_name)
    if not candidates:
        return None, 0

    best_score = 0
    best_match = None

    for cand in candidates:
        for ename in unmatched_excel:
            # ename already normalized by read_party_names
            if cand in ename or ename in cand:
                score = fuzz.ratio(cand, ename)
                if score > best_score:
                    best_score = score
                    best_match = ename

    if best_score >= threshold and best_match:
        return best_match, best_score
    elif best_score >= threshold - 10:
        return best_match, best_score
    else:
        return None, best_score


def match_names(
    excel_names: List[str],
    wechat_names: List[str],
    threshold: int = DEFAULT_THRESHOLD,
) -> Tuple[List[dict], List[dict], List[dict]]:
    """执行完整的两轮匹配。

    Args:
        excel_names: Excel 党员姓名列表
        wechat_names: 微信群成员昵称列表（原始格式）
        threshold: 模糊匹配阈值 (0-100)

    Returns:
        (matched, not_found, uncertain) 三个列表
    """
    matched = []
    not_found = []
    uncertain = []
    consumed_wechat = set()  # 已被匹配的微信昵称索引

    # 预规范化所有微信昵称（避免在热循环中重复 normalize）
    wechat_clean = [normalize_name(w) for w in wechat_names]
    # 预建原始名→索引映射（O(1) 查找替代 O(n) 扫描）
    wechat_index = {w: j for j, w in enumerate(wechat_names)}

    # === 第一轮：正向匹配 ===
    for ename in excel_names:
        best_w, best_raw, score, method = match_single_forward(
            ename, wechat_names, wechat_clean, threshold
        )

        if best_w and method and not method.startswith('uncertain'):
            # 标记该微信昵称已被消费
            idx = wechat_index.get(best_raw)
            if idx is not None:
                consumed_wechat.add(idx)
            matched.append({
                'excel_name': ename,
                'wechat_name': best_w,
                'wechat_name_raw': best_raw,
                'method': method,
                'confidence': score,
            })
        elif best_w and method and method.startswith('uncertain'):
            uncertain.append({
                'excel_name': ename,
                'best_candidate_raw': best_raw,
                'best_candidate': best_w,
                'confidence': score,
                'method': method,
            })
        else:
            not_found.append({
                'excel_name': ename,
                'reason': 'no_match_pass1',
                'best_score': score,
            })

    # === 第二轮：反向匹配 ===
    # 用 dict 替代 list + enumerate 扫描（O(1) 查找）
    not_found_map = {item['excel_name']: item for item in not_found}
    for j, w_raw in enumerate(wechat_names):
        if j in consumed_wechat:
            continue
        w_clean = wechat_clean[j]
        best_e, score = match_single_reverse(
            w_clean, list(not_found_map.keys()), threshold
        )
        if best_e and score >= threshold and best_e in not_found_map:
            matched.append({
                'excel_name': best_e,
                'wechat_name': w_clean,
                'wechat_name_raw': w_raw,
                'method': 'reverse_pass2',
                'confidence': score,
            })
            del not_found_map[best_e]
            consumed_wechat.add(j)

    # 用更新后的 not_found_map 重建 not_found 列表
    not_found = list(not_found_map.values())

    # === 处理待确认项 ===
    # 将 confidence 足够高的 uncertain 提升为 matched
    confirmed_uncertain = []
    for item in uncertain:
        if item['confidence'] >= threshold:
            matched.append({
                'excel_name': item['excel_name'],
                'wechat_name': item['best_candidate'],
                'wechat_name_raw': item['best_candidate_raw'],
                'method': item['method'].replace('uncertain_', ''),
                'confidence': item['confidence'],
            })
            confirmed_uncertain.append(item)
    for item in confirmed_uncertain:
        uncertain.remove(item)

    return matched, not_found, uncertain


# ============================================================================
# Section 6: Output / Reporting
# ============================================================================

def print_results_plain(
    matched: List[dict],
    not_found: List[dict],
    uncertain: List[dict],
    excel_names: List[str],
    wechat_names: List[str],
    verbose: bool = False,
):
    """纯文本控制台输出（无 rich 时的降级方案）。"""
    total_excel = len(excel_names)
    total_wechat = len(wechat_names)
    total_matched = len(matched)
    total_not_found = len(not_found)
    total_uncertain = len(uncertain)

    sep = "=" * 60

    print(sep)
    print("  党员微信群成员核对结果")
    print(sep)

    # 汇总
    print(f"\n  Excel 党员总数:   {total_excel}")
    print(f"  微信群成员数:     {total_wechat}")
    print(f"  已匹配:           {total_matched} ({100*total_matched//max(total_excel,1)}%)")
    print(f"  未找到:           {total_not_found}")
    print(f"  待确认:           {total_uncertain}")

    # 未找到
    if not_found:
        print(f"\n{sep}")
        print("  ❌ 未找到（需跟进联系）")
        print(sep)
        for i, item in enumerate(not_found, 1):
            print(f"  {i:>3}. {item['excel_name']}")

    # 待确认
    if uncertain:
        print(f"\n{sep}")
        print("  ⚠️  待确认（建议手动核实）")
        print(sep)
        print(f"  {'序号':<6}{'Excel姓名':<12}{'微信候选':<18}{'相似度':<8}")
        print(f"  {'-'*44}")
        for i, item in enumerate(uncertain, 1):
            print(f"  {i:<6}{item['excel_name']:<12}"
                  f"{item['best_candidate']:<18}{item['confidence']:.0f}%")

    # 详细匹配
    if verbose and matched:
        print(f"\n{sep}")
        print("  ✅ 已匹配明细")
        print(sep)
        print(f"  {'序号':<6}{'Excel姓名':<12}{'微信昵称':<24}{'方式':<20}{'置信度':<8}")
        print(f"  {'-'*72}")
        for i, item in enumerate(matched, 1):
            print(f"  {i:<6}{item['excel_name']:<12}"
                  f"{item['wechat_name']:<24}{item['method']:<20}"
                  f"{item['confidence']:.0f}%")

    print(f"\n{sep}")
    print("  核对完成。")
    print(sep)


def print_results_rich(
    matched: List[dict],
    not_found: List[dict],
    uncertain: List[dict],
    excel_names: List[str],
    wechat_names: List[str],
    verbose: bool = False,
):
    """Rich 格式控制台输出。"""
    total_excel = len(excel_names)
    total_wechat = len(wechat_names)
    total_matched = len(matched)
    total_not_found = len(not_found)
    total_uncertain = len(uncertain)

    # 标题
    title = Panel.fit(
        "[bold cyan]党员微信群成员核对结果[/bold cyan]",
        border_style="cyan",
    )
    console.print()
    console.print(title)

    # 汇总面板
    summary_text = (
        f"  📋 Excel 党员总数:   [bold]{total_excel}[/bold]\n"
        f"  👥 微信群成员数:     [bold]{total_wechat}[/bold]\n"
        f"  ✅ 已匹配:           [bold green]{total_matched}[/bold green]"
        f" ({100*total_matched//max(total_excel,1)}%)\n"
        f"  ❌ 未找到:           [bold red]{total_not_found}[/bold red]\n"
        f"  ⚠️  待确认:           [bold yellow]{total_uncertain}[/bold yellow]"
    )
    console.print(Panel(summary_text, title="汇总", border_style="blue"))

    # 未找到表格
    if not_found:
        console.print()
        table = Table(
            title="❌ 未找到（需跟进联系）",
            box=box.SIMPLE_HEAVY,
            title_style="bold red",
            border_style="red",
        )
        table.add_column("序号", style="dim", width=6)
        table.add_column("Excel 姓名", style="bold red", width=16)

        for i, item in enumerate(not_found, 1):
            table.add_row(str(i), item['excel_name'])

        console.print(table)

    # 待确认表格
    if uncertain:
        console.print()
        table = Table(
            title="⚠️  待确认（建议手动核实）",
            box=box.SIMPLE_HEAVY,
            title_style="bold yellow",
            border_style="yellow",
        )
        table.add_column("序号", style="dim", width=6)
        table.add_column("Excel 姓名", style="bold yellow", width=16)
        table.add_column("微信中最接近", style="cyan", width=24)
        table.add_column("相似度", width=10)

        for i, item in enumerate(uncertain, 1):
            table.add_row(
                str(i),
                item['excel_name'],
                item['best_candidate'],
                f"{item['confidence']:.0f}%",
            )

        console.print(table)

    # 详细匹配表格
    if verbose and matched:
        console.print()
        table = Table(
            title="✅ 已匹配明细",
            box=box.SIMPLE_HEAVY,
            title_style="bold green",
            border_style="green",
        )
        table.add_column("序号", style="dim", width=6)
        table.add_column("Excel 姓名", style="bold green", width=16)
        table.add_column("微信昵称", style="cyan", width=28)
        table.add_column("匹配方式", width=22)
        table.add_column("置信度", width=8)

        for i, item in enumerate(matched, 1):
            table.add_row(
                str(i),
                item['excel_name'],
                item['wechat_name'],
                item['method'],
                f"{item['confidence']:.0f}%",
            )

        console.print(table)

    console.print()
    console.print("[bold green]✅ 核对完成。[/bold green]")


def print_results(
    matched: List[dict],
    not_found: List[dict],
    uncertain: List[dict],
    excel_names: List[str],
    wechat_names: List[str],
    verbose: bool = False,
):
    """输出核对结果（自动选择 Rich 或纯文本模式）。"""
    if RICH_AVAILABLE:
        print_results_rich(
            matched, not_found, uncertain,
            excel_names, wechat_names, verbose,
        )
    else:
        print_results_plain(
            matched, not_found, uncertain,
            excel_names, wechat_names, verbose,
        )


def export_to_excel(
    matched: List[dict],
    not_found: List[dict],
    uncertain: List[dict],
    wechat_names: List[str],
    output_path: str,
):
    """导出核对结果到 Excel 文件（多 Sheet）。"""
    path = Path(output_path).expanduser().resolve()

    # Sheet 1: 核对结果总表
    all_rows = []
    for item in matched:
        all_rows.append({
            'Excel姓名': item['excel_name'],
            '核对结果': '已匹配',
            '微信昵称': item['wechat_name'],
            '微信原始名': item['wechat_name_raw'],
            '匹配方式': item['method'],
            '置信度': item['confidence'],
        })
    for item in not_found:
        all_rows.append({
            'Excel姓名': item['excel_name'],
            '核对结果': '未找到',
            '微信昵称': '',
            '微信原始名': '',
            '匹配方式': '',
            '置信度': '',
        })
    for item in uncertain:
        all_rows.append({
            'Excel姓名': item['excel_name'],
            '核对结果': '待确认',
            '微信昵称': item['best_candidate'],
            '微信原始名': item['best_candidate_raw'],
            '匹配方式': item['method'],
            '置信度': item['confidence'],
        })
    df_all = pd.DataFrame(all_rows)

    # Sheet 2: 未找到
    df_not = pd.DataFrame([{
        'Excel姓名': item['excel_name'],
        '备注': item.get('reason', ''),
    } for item in not_found])

    # Sheet 3: 待确认
    df_uncertain = pd.DataFrame([{
        'Excel姓名': item['excel_name'],
        '微信候选': item['best_candidate'],
        '微信原始名': item['best_candidate_raw'],
        '相似度': item['confidence'],
    } for item in uncertain])

    # Sheet 4: 微信群全量成员
    df_wechat = pd.DataFrame([{
        '序号': i,
        '微信昵称': name,
    } for i, name in enumerate(wechat_names, 1)])

    with pd.ExcelWriter(path, engine='openpyxl') as writer:
        df_all.to_excel(writer, sheet_name='核对结果', index=False)
        df_not.to_excel(writer, sheet_name='未找到', index=False)
        df_uncertain.to_excel(writer, sheet_name='待确认', index=False)
        df_wechat.to_excel(writer, sheet_name='微信群全量', index=False)

    print(f"\n结果已导出到: {path}")


# ============================================================================
# Section 7: CLI Entry Point
# ============================================================================

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description='党员微信群成员核对工具 — 对比Excel党员名单与微信群成员，找出未进群或备注不规范的人',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
    python party_wechat_check.py --excel "名单.xlsx" --wechat-file "members.txt"
    python party_wechat_check.py --excel "名单.xlsx" --wechat-file "members.txt" --output "结果.xlsx" --verbose
    python party_wechat_check.py   # 交互式输入
        """,
    )

    parser.add_argument(
        '--excel', '-e',
        type=str,
        help='党员名单 Excel 文件路径 (.xlsx)',
    )
    parser.add_argument(
        '--wechat-file', '-w',
        type=str,
        help='微信群成员提取文本文件路径 (.txt)',
    )
    parser.add_argument(
        '--name-column', '-c',
        type=str,
        help='手动指定 Excel 中的姓名列名（默认自动检测）',
    )
    parser.add_argument(
        '--threshold', '-t',
        type=int,
        default=DEFAULT_THRESHOLD,
        help=f'模糊匹配阈值 0-100 (默认: {DEFAULT_THRESHOLD})',
    )
    parser.add_argument(
        '--output', '-o',
        type=str,
        help='导出结果到 Excel 文件 (.xlsx)',
    )
    parser.add_argument(
        '--verbose', '-v',
        action='store_true',
        help='显示全部已匹配明细',
    )
    parser.add_argument(
        '--version', '-V',
        action='version',
        version=f'%(prog)s {__version__}',
        help='显示版本号',
    )

    return parser


def main():
    parser = build_parser()
    args = parser.parse_args()

    # 交互式输入
    excel_path = args.excel
    wechat_path = args.wechat_file

    if not excel_path:
        excel_path = input("请输入党员名单 Excel 文件路径: ").strip().strip('"')
    if not wechat_path:
        wechat_path = input("请输入微信群成员文本文件路径: ").strip().strip('"')

    try:
        # Step 1: 读取 Excel
        print(f"\n📖 正在读取 Excel: {excel_path}")
        excel_names, used_column = read_party_names(excel_path, args.name_column)
        print(f"   → 从列 '{used_column}' 读取到 {len(excel_names)} 个姓名")

        # Step 2: 读取微信群
        print(f"\n📖 正在读取微信群文本: {wechat_path}")
        wechat_names = read_wechat_members(wechat_path)
        print(f"   → 识别到 {len(wechat_names)} 个微信群成员")

        # Step 3: 匹配
        print(f"\n🔍 正在匹配（阈值: {args.threshold}%）...")
        matched, not_found, uncertain = match_names(
            excel_names, wechat_names, args.threshold,
        )
        print(f"   → 匹配完成!")

        # Step 4: 输出结果
        print_results(
            matched, not_found, uncertain,
            excel_names, wechat_names,
            verbose=args.verbose,
        )

        # Step 5: 导出 Excel
        if args.output:
            export_to_excel(
                matched, not_found, uncertain,
                wechat_names, args.output,
            )

    except KeyboardInterrupt:
        print("\n\n⚠️  用户中断")
        sys.exit(130)
    except FileNotFoundError as e:
        print(f"\n❌ 文件错误: {e}")
        sys.exit(1)
    except ValueError as e:
        print(f"\n❌ 数据错误: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ 未知错误: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()
