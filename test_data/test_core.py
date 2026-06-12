#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
党员微信群核对工具 — 核心函数单元测试
======================================

运行方式:
    pip install pytest
    pytest test_data/test_core.py -v

覆盖:
    - normalize_name:    文本规范化
    - extract_chinese_name_candidates: 姓名片段提取（含复姓）
    - strip_common_prefixes: 前缀去除
    - detect_name_column: Excel列自动检测
    - parse_wechat_text: 微信群OCR文本解析
    - match_single_forward / match_single_reverse / match_names: 匹配引擎
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pytest
import pandas as pd

from party_wechat_check import (
    normalize_name,
    extract_chinese_name_candidates,
    strip_common_prefixes,
    detect_name_column,
    parse_wechat_text,
    match_single_forward,
    match_single_reverse,
    match_names,
    read_party_names,
    read_wechat_members,
    COMPOUND_SURNAMES,
    DEFAULT_THRESHOLD,
)

# ============================================================================
# normalize_name
# ============================================================================

class TestNormalizeName:
    def test_basic_chinese_name_passes_through(self):
        assert normalize_name("张三") == "张三"

    def test_strips_whitespace(self):
        assert normalize_name("  张三  ") == "张三"

    def test_strips_emoji(self):
        assert normalize_name("张三😊") == "张三"

    def test_strips_bracket_decorations(self):
        assert normalize_name("张三[01班]") == "张三"
        assert normalize_name("（宣传委员）张三") == "张三"

    def test_strips_dash_suffix(self):
        result = normalize_name("张三-党建办")
        assert "张三" in result
        assert "党建办" not in result

    def test_handles_empty(self):
        assert normalize_name("") == ""

    def test_handles_nan(self):
        import math
        assert normalize_name(float("nan")) == ""

    def test_handles_none(self):
        assert normalize_name(None) == ""


# ============================================================================
# extract_chinese_name_candidates
# ============================================================================

class TestExtractCandidates:
    def test_simple_name(self):
        cands = extract_chinese_name_candidates("张三")
        assert "张三" in cands

    def test_compound_surname_preserved(self):
        """复姓不应被拆分：'欧阳泽坪' 不应产生 '阳泽坪'"""
        cands = extract_chinese_name_candidates("欧阳泽坪")
        assert "欧阳泽坪" in cands
        assert "欧阳泽" in cands
        assert "欧阳" in cands
        assert "阳泽坪" not in cands  # 关键：不应拆分复姓

    def test_compound_surname_in_mixed_text(self):
        """复姓 + 单姓名混合"""
        cands = extract_chinese_name_candidates("欧阳泽坪张三")
        assert "欧阳泽坪" in cands
        assert "张三" in cands

    def test_returns_longest_first(self):
        """按长度降序排列"""
        cands = extract_chinese_name_candidates("欧阳泽坪")
        lengths = [len(c) for c in cands]
        assert lengths == sorted(lengths, reverse=True)

    def test_min_length_is_2(self):
        """最短候选为2字"""
        cands = extract_chinese_name_candidates("张三")
        assert all(len(c) >= 2 for c in cands)
        # 单字不应出现
        assert "张" not in cands


# ============================================================================
# strip_common_prefixes
# ============================================================================

class TestStripPrefixes:
    def test_grade_prefix(self):
        assert strip_common_prefixes("22级张三") == "张三"

    def test_master_prefix(self):
        assert strip_common_prefixes("硕2024张三") == "张三"

    def test_doctor_prefix(self):
        assert strip_common_prefixes("博士2023-张三") == "张三"

    def test_no_prefix_passes_through(self):
        assert strip_common_prefixes("张三") == "张三"


# ============================================================================
# detect_name_column
# ============================================================================

class TestDetectNameColumn:
    def test_detect_exact_match(self):
        df = pd.DataFrame({"姓名": ["张三", "李四"]})
        assert detect_name_column(df) == "姓名"

    def test_detect_contains_match(self):
        df = pd.DataFrame({"党员姓名": ["张三"]})
        assert detect_name_column(df) == "党员姓名"

    def test_no_match_returns_none(self):
        df = pd.DataFrame({"代号": ["001"]})
        assert detect_name_column(df) is None


# ============================================================================
# parse_wechat_text
# ============================================================================

class TestParseWechatText:
    def test_basic_names(self):
        text = "张三\n李四\n王五"
        members = parse_wechat_text(text)
        assert "张三" in members
        assert "李四" in members
        assert "王五" in members

    def test_filters_skip_keywords(self):
        text = "张三\n搜索群成员\n李四\n添加"
        members = parse_wechat_text(text)
        assert "张三" in members
        assert "李四" in members
        assert "搜索" not in str(members)

    def test_splits_multi_name_lines(self):
        """OCR 常见：一行多个名字连在一起"""
        text = "朱永洁刘莹黄晓君"
        members = parse_wechat_text(text)
        assert "朱永洁" in members
        assert "刘莹" in members
        assert "黄晓君" in members

    def test_dedup(self):
        text = "张三\n张三\n张三"
        members = parse_wechat_text(text)
        assert len([m for m in members if m == "张三"]) == 1

    def test_preserves_raw_lines_as_fallback(self):
        """纯中文行至少有整行作为候选"""
        text = "张三丰"
        members = parse_wechat_text(text)
        assert len(members) > 0


# ============================================================================
# match_single_forward
# ============================================================================

class TestMatchForward:
    def test_direct_contains(self):
        wechat_raw = ["张三-01班", "李四"]
        wechat_clean = [normalize_name(w) for w in wechat_raw]
        match, raw, score, method = match_single_forward(
            "张三", wechat_raw, wechat_clean, DEFAULT_THRESHOLD
        )
        assert match is not None
        assert method == "direct_contains"
        assert score == 100

    def test_no_match(self):
        wechat_raw = ["李四", "王五"]
        wechat_clean = [normalize_name(w) for w in wechat_raw]
        match, raw, score, method = match_single_forward(
            "张三", wechat_raw, wechat_clean, DEFAULT_THRESHOLD
        )
        assert match is None


# ============================================================================
# match_single_reverse
# ============================================================================

class TestMatchReverse:
    def test_extract_name_from_nickname(self):
        """从微信昵称中提取中文名匹配"""
        match, score = match_single_reverse(
            "欧阳泽坪和张三一起", ["欧阳泽坪"], DEFAULT_THRESHOLD
        )
        assert match == "欧阳泽坪"

    def test_no_candidates_returns_none(self):
        match, score = match_single_reverse(
            "abc123", ["张三"], DEFAULT_THRESHOLD
        )
        assert match is None


# ============================================================================
# match_names (integration)
# ============================================================================

class TestMatchNames:
    def test_full_match_flow(self):
        excel = ["张三", "李四", "王五", "赵六"]
        wechat = ["张三-01班", "王五", "赵六六", "陌生人"]
        matched, not_found, uncertain = match_names(excel, wechat, DEFAULT_THRESHOLD)

        matched_names = {m["excel_name"] for m in matched}
        assert "张三" in matched_names  # direct_contains
        assert "王五" in matched_names  # direct_contains
        assert "赵六" in matched_names  # fuzzy: 赵六 in 赵六六

        not_found_names = {n["excel_name"] for n in not_found}
        assert "李四" in not_found_names

    def test_compound_surname_matching(self):
        """复姓姓名应能正确匹配"""
        excel = ["欧阳泽坪", "司马光"]
        wechat = ["欧阳泽", "司马光同学"]
        matched, not_found, uncertain = match_names(excel, wechat, DEFAULT_THRESHOLD)

        matched_names = {m["excel_name"] for m in matched}
        assert "欧阳泽坪" in matched_names  # reverse_contains: 欧阳泽 in 欧阳泽坪
        assert "司马光" in matched_names    # direct_contains: 司马光 in 司马光同学


# ============================================================================
# End-to-end test with fixture files
# ============================================================================

class TestEndToEnd:
    def test_with_fixture_files(self):
        """用 test_data 目录的固定数据做端到端测试"""
        import os
        test_dir = os.path.dirname(os.path.abspath(__file__))
        excel_path = os.path.join(test_dir, "test_party.xlsx")
        wechat_path = os.path.join(test_dir, "test_wechat_members.txt")

        excel_names, col = read_party_names(excel_path)
        wechat_names = read_wechat_members(wechat_path)

        assert len(excel_names) == 10
        assert len(wechat_names) >= 10  # 13行原始数据，解析后应有足够条目

        matched, not_found, uncertain = match_names(excel_names, wechat_names)

        # 应该有至少 8 个匹配（10个里李四不在微信群，其他应匹配）
        assert len(matched) >= 8
        assert len(not_found) >= 1
        assert len(uncertain) == 0

        # 李四在名单但不在微信群
        not_found_names = {n["excel_name"] for n in not_found}
        assert "李四" in not_found_names


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
