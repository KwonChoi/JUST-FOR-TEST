#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
出境旅游信息查询 Web应用 - 测试用例生成器
支持输出 Excel(.xlsx) 和 XMind(.xmind) 两种格式
测试设计方法：等价类、边界值、场景法、错误推测
"""

import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import xmind


def create_excel_test_cases(output_path):
    """生成 Excel 格式测试用例"""
    wb = openpyxl.Workbook()

    # 统一样式定义
    header_font = Font(name='微软雅黑', bold=True, size=11, color='FFFFFF')
    header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    header_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    cell_font = Font(name='微软雅黑', size=10)
    cell_alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    center_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    thin_border = Border(
        left=Side(style='thin', color='B4B4B4'),
        right=Side(style='thin', color='B4B4B4'),
        top=Side(style='thin', color='B4B4B4'),
        bottom=Side(style='thin', color='B4B4B4')
    )

    priority_styles = {
        'P0': Font(name='微软雅黑', size=10, bold=True, color='C00000'),
        'P1': Font(name='微软雅黑', size=10, bold=True, color='E36C0A'),
        'P2': Font(name='微软雅黑', size=10, color='0070C0'),
        'P3': Font(name='微软雅黑', size=10, color='595959'),
    }

    method_fills = {
        '等价类': PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid'),
        '边界值': PatternFill(start_color='FCE4D6', end_color='FCE4D6', fill_type='solid'),
        '场景法': PatternFill(start_color='DDEBF7', end_color='DDEBF7', fill_type='solid'),
        '错误推测': PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid'),
    }

    # ==================== Sheet1: 测试用例 ====================
    ws = wb.active
    ws.title = '测试用例'

    headers = ['用例编号', '所属模块', '功能点', '用例标题', '前置条件', '测试步骤', '测试数据', '预期结果', '测试方法', '优先级']
    ws.append(headers)

    # 设置表头样式
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border

    # 测试用例数据
    test_cases = [
        # ========== 首页展示 ==========
        ('TC-HOME-001', '首页展示', '页面加载', '验证页面标题正确显示', '已打开应用首页', '1. 打开首页\n2. 查看浏览器标签页标题', '-', '页面标题显示为"出境旅游信息查询"', '场景法', 'P0'),
        ('TC-HOME-002', '首页展示', '页面加载', '验证Logo和搜索区域显示', '已打开应用首页', '1. 查看页面顶部', '-', '显示"🌍 出境旅游信息查询"标题、搜索输入框和搜索按钮', '场景法', 'P0'),
        ('TC-HOME-003', '首页展示', '区域筛选标签', '验证6个区域筛选标签完整显示', '已打开应用首页', '1. 查看筛选标签区域', '-', '依次显示：全部、亚洲、欧洲、美洲、非洲、大洋洲，且"全部"默认选中', '场景法', 'P0'),
        ('TC-HOME-004', '首页展示', '国家卡片', '验证默认展示4个国家卡片', '已打开应用首页', '1. 查看国家网格区域', '-', '显示日本、肯尼亚、法国、泰国4个国家卡片，每个卡片包含国旗、名称、区域、简介', '场景法', 'P0'),
        ('TC-HOME-005', '首页展示', '国家卡片', '验证国家卡片信息正确性', '已打开应用首页', '1. 查看日本卡片\n2. 核对国旗、名称、区域、简介', '-', '日本卡片显示：日本国旗、"日本"、"亚洲"、"樱花之国，传统文化与现代科技完美融合"', '场景法', 'P1'),
        ('TC-HOME-006', '首页展示', '国旗图片', '验证国旗图片正常加载', '已打开应用首页，网络正常', '1. 查看4个国家卡片的国旗图片', '-', '4个国家的国旗图片均正常显示，无破损或空白', '场景法', 'P1'),

        # ========== 搜索功能 ==========
        ('TC-SEARCH-001', '搜索功能', '中文搜索-有效全名', '输入有效中文国家全名"日本"进行搜索', '已打开应用首页', '1. 在搜索框输入"日本"\n2. 点击搜索按钮', '日本', '仅显示日本的国家卡片', '等价类', 'P0'),
        ('TC-SEARCH-002', '搜索功能', '中文搜索-有效全名', '输入有效中文国家全名"法国"进行搜索', '已打开应用首页', '1. 在搜索框输入"法国"\n2. 点击搜索按钮', '法国', '仅显示法国的国家卡片', '等价类', 'P0'),
        ('TC-SEARCH-003', '搜索功能', '英文搜索-有效全名', '输入有效英文国家全名"Japan"进行搜索', '已打开应用首页', '1. 在搜索框输入"Japan"\n2. 点击搜索按钮', 'Japan', '仅显示日本的国家卡片', '等价类', 'P0'),
        ('TC-SEARCH-004', '搜索功能', '英文搜索-有效全名', '输入有效英文国家全名"France"进行搜索', '已打开应用首页', '1. 在搜索框输入"France"\n2. 点击搜索按钮', 'France', '仅显示法国的国家卡片', '等价类', 'P0'),
        ('TC-SEARCH-005', '搜索功能', '中文搜索-部分匹配', '输入部分中文"日"进行搜索', '已打开应用首页', '1. 在搜索框输入"日"\n2. 点击搜索按钮', '日', '显示日本的国家卡片', '等价类', 'P1'),
        ('TC-SEARCH-006', '搜索功能', '中文搜索-部分匹配', '输入部分中文"泰"进行搜索', '已打开应用首页', '1. 在搜索框输入"泰"\n2. 点击搜索按钮', '泰', '显示泰国的国家卡片', '等价类', 'P1'),
        ('TC-SEARCH-007', '搜索功能', '英文搜索-部分匹配', '输入部分英文"Ja"进行搜索', '已打开应用首页', '1. 在搜索框输入"Ja"\n2. 点击搜索按钮', 'Ja', '显示日本(Japan)的国家卡片', '等价类', 'P1'),
        ('TC-SEARCH-008', '搜索功能', '英文搜索-部分匹配', '输入部分英文"Ke"进行搜索', '已打开应用首页', '1. 在搜索框输入"Ke"\n2. 点击搜索按钮', 'Ke', '显示肯尼亚(Kenya)的国家卡片', '等价类', 'P1'),
        ('TC-SEARCH-009', '搜索功能', '英文搜索-大小写混合', '输入大小写混合"jaPaN"进行搜索', '已打开应用首页', '1. 在搜索框输入"jaPaN"\n2. 点击搜索按钮', 'jaPaN', '显示日本的国家卡片（搜索不区分大小写）', '等价类', 'P1'),
        ('TC-SEARCH-010', '搜索功能', '无效搜索-不存在国家', '输入不存在的国家名"德国"进行搜索', '已打开应用首页', '1. 在搜索框输入"德国"\n2. 点击搜索按钮', '德国', '显示"没有找到符合条件的国家"提示信息', '等价类', 'P1'),
        ('TC-SEARCH-011', '搜索功能', '无效搜索-特殊字符', '输入特殊字符"!@#"进行搜索', '已打开应用首页', '1. 在搜索框输入"!@#"\n2. 点击搜索按钮', '!@#', '显示"没有找到符合条件的国家"提示信息', '等价类', 'P1'),
        ('TC-SEARCH-012', '搜索功能', '无效搜索-纯数字', '输入纯数字"123"进行搜索', '已打开应用首页', '1. 在搜索框输入"123"\n2. 点击搜索按钮', '123', '显示"没有找到符合条件的国家"提示信息', '等价类', 'P1'),
        ('TC-SEARCH-013', '搜索功能', '空值搜索', '搜索框为空时点击搜索', '已打开应用首页', '1. 确保搜索框为空\n2. 点击搜索按钮', '', '显示全部4个国家卡片', '等价类', 'P0'),
        ('TC-SEARCH-014', '搜索功能', '边界值-最小长度', '输入1个字符"法"进行搜索', '已打开应用首页', '1. 在搜索框输入"法"\n2. 点击搜索按钮', '法', '显示法国的国家卡片', '边界值', 'P1'),
        ('TC-SEARCH-015', '搜索功能', '边界值-正常长度', '输入正常长度字符"Thailand"进行搜索', '已打开应用首页', '1. 在搜索框输入"Thailand"\n2. 点击搜索按钮', 'Thailand', '显示泰国的国家卡片', '边界值', 'P1'),
        ('TC-SEARCH-016', '搜索功能', '边界值-超长字符串', '输入50个字符的超长字符串搜索', '已打开应用首页', '1. 在搜索框输入50个字符\n2. 点击搜索按钮', 'a' * 50, '显示"没有找到符合条件的国家"提示信息，页面不报错', '边界值', 'P2'),
        ('TC-SEARCH-017', '搜索功能', '搜索触发方式', '在搜索框按Enter键触发搜索', '已打开应用首页', '1. 在搜索框输入"日本"\n2. 按Enter键', '日本', '显示日本的国家卡片', '场景法', 'P1'),
        ('TC-SEARCH-018', '搜索功能', '搜索触发方式', '在搜索框输入后点击搜索按钮触发搜索', '已打开应用首页', '1. 在搜索框输入"日本"\n2. 点击搜索按钮', '日本', '显示日本的国家卡片', '场景法', 'P1'),
        ('TC-SEARCH-019', '搜索功能', '错误推测-快速输入', '快速连续输入并搜索', '已打开应用首页', '1. 快速输入"日日日本本本"\n2. 点击搜索', '日日日本本本', '页面正常响应，显示日本卡片或根据最终输入过滤', '错误推测', 'P2'),

        # ========== 区域筛选 ==========
        ('TC-FILTER-001', '区域筛选', '全部筛选', '点击"全部"标签显示所有国家', '已打开应用首页', '1. 点击"全部"筛选标签', '-', '显示日本、肯尼亚、法国、泰国4个国家卡片', '场景法', 'P0'),
        ('TC-FILTER-002', '区域筛选', '亚洲筛选', '点击"亚洲"标签显示亚洲国家', '已打开应用首页', '1. 点击"亚洲"筛选标签', '-', '显示日本、泰国2个国家卡片', '等价类', 'P0'),
        ('TC-FILTER-003', '区域筛选', '欧洲筛选', '点击"欧洲"标签显示欧洲国家', '已打开应用首页', '1. 点击"欧洲"筛选标签', '-', '显示法国1个国家卡片', '等价类', 'P0'),
        ('TC-FILTER-004', '区域筛选', '非洲筛选', '点击"非洲"标签显示非洲国家', '已打开应用首页', '1. 点击"非洲"筛选标签', '-', '显示肯尼亚1个国家卡片', '等价类', 'P0'),
        ('TC-FILTER-005', '区域筛选', '美洲筛选', '点击"美洲"标签显示美洲国家', '已打开应用首页', '1. 点击"美洲"筛选标签', '-', '显示"没有找到符合条件的国家"提示信息', '等价类', 'P1'),
        ('TC-FILTER-006', '区域筛选', '大洋洲筛选', '点击"大洋洲"标签显示大洋洲国家', '已打开应用首页', '1. 点击"大洋洲"筛选标签', '-', '显示"没有找到符合条件的国家"提示信息', '等价类', 'P1'),
        ('TC-FILTER-007', '区域筛选', '筛选状态保持', '切换筛选标签时高亮状态正确更新', '已打开应用首页', '1. 点击"亚洲"标签\n2. 观察标签样式\n3. 点击"欧洲"标签', '-', '"亚洲"标签取消高亮，"欧洲"标签变为高亮', '场景法', 'P1'),
        ('TC-FILTER-008', '区域筛选', '筛选+搜索组合', '先筛选亚洲再搜索日本', '已打开应用首页', '1. 点击"亚洲"筛选\n2. 在搜索框输入"日本"\n3. 点击搜索', '日本', '显示日本的国家卡片', '场景法', 'P1'),
        ('TC-FILTER-009', '区域筛选', '筛选+搜索组合', '先筛选欧洲再搜索日本（无交集）', '已打开应用首页', '1. 点击"欧洲"筛选\n2. 在搜索框输入"日本"\n3. 点击搜索', '日本', '显示"没有找到符合条件的国家"提示信息', '场景法', 'P1'),
        ('TC-FILTER-010', '区域筛选', '筛选+搜索组合', '先筛选非洲再搜索 Kenya', '已打开应用首页', '1. 点击"非洲"筛选\n2. 在搜索框输入"Kenya"\n3. 点击搜索', 'Kenya', '显示肯尼亚的国家卡片', '场景法', 'P1'),

        # ========== 国家详情模态窗口 ==========
        ('TC-DETAIL-001', '国家详情', '打开模态窗口', '点击日本卡片打开详情模态窗口', '已打开应用首页', '1. 点击日本国家卡片', '-', '模态窗口正常打开，显示日本详细信息', '场景法', 'P0'),
        ('TC-DETAIL-002', '国家详情', '打开模态窗口', '点击肯尼亚卡片打开详情模态窗口', '已打开应用首页', '1. 点击肯尼亚国家卡片', '-', '模态窗口正常打开，显示肯尼亚详细信息', '场景法', 'P0'),
        ('TC-DETAIL-003', '国家详情', '基本信息展示', '验证模态窗口头部信息正确', '已打开日本详情页', '1. 查看模态窗口头部', '-', '显示日本国旗、"日本"、"Japan • 东京 • 1.26亿人口"', '场景法', 'P0'),
        ('TC-DETAIL-004', '国家详情', '标签页展示', '验证6个信息标签页完整显示', '已打开任意国家详情页', '1. 查看标签栏', '-', '依次显示：签证信息、货币汇率、语言交流、习俗禁忌、玩法路线、高频词汇，且"签证信息"默认选中', '场景法', 'P0'),
        ('TC-DETAIL-005', '国家详情', '标签切换-货币汇率', '点击"货币汇率"标签显示货币内容', '已打开日本详情页', '1. 点击"货币汇率"标签', '-', '显示货币名称、代码、符号、汇率参考、支付建议', '场景法', 'P0'),
        ('TC-DETAIL-006', '国家详情', '标签切换-语言交流', '点击"语言交流"标签显示语言内容', '已打开日本详情页', '1. 点击"语言交流"标签', '-', '显示官方语言、常用问候语列表、实用短语列表', '场景法', 'P0'),
        ('TC-DETAIL-007', '国家详情', '标签切换-习俗禁忌', '点击"习俗禁忌"标签显示习俗内容', '已打开日本详情页', '1. 点击"习俗禁忌"标签', '-', '显示礼仪文化、用餐礼仪、公共行为等条目', '场景法', 'P0'),
        ('TC-DETAIL-008', '国家详情', '标签切换-玩法路线', '点击"玩法路线"标签显示路线内容', '已打开日本详情页', '1. 点击"玩法路线"标签', '-', '显示经典关西5日游、东京深度7日游等路线信息', '场景法', 'P0'),
        ('TC-DETAIL-009', '国家详情', '标签切换-高频词汇', '点击"高频词汇"标签显示词汇内容', '已打开日本详情页', '1. 点击"高频词汇"标签', '-', '显示寿司、拉面、温泉等中日英对照词汇', '场景法', 'P0'),
        ('TC-DETAIL-010', '国家详情', '标签切换-返回签证', '点击"签证信息"标签返回签证内容', '已切换到其他标签页', '1. 点击"签证信息"标签', '-', '显示签证类型、停留时间、办理难度、申请材料、申请建议', '场景法', 'P0'),
        ('TC-DETAIL-011', '国家详情', '标签高亮状态', '切换标签时高亮状态正确更新', '已打开详情页', '1. 点击"货币汇率"标签\n2. 观察标签样式', '-', '"货币汇率"标签变为高亮，其他标签取消高亮', '场景法', 'P1'),
        ('TC-DETAIL-012', '国家详情', '关闭模态窗口-X按钮', '点击X按钮关闭模态窗口', '已打开详情页', '1. 点击右上角的X按钮', '-', '模态窗口关闭，背景页面恢复可滚动', '场景法', 'P0'),
        ('TC-DETAIL-013', '国家详情', '关闭模态窗口-点击外部', '点击模态窗口外部区域关闭', '已打开详情页', '1. 点击模态窗口以外的深色背景区域', '-', '模态窗口关闭，背景页面恢复可滚动', '场景法', 'P0'),
        ('TC-DETAIL-014', '国家详情', '关闭模态窗口-ESC键', '按ESC键关闭模态窗口', '已打开详情页', '1. 按键盘ESC键', '-', '模态窗口关闭，背景页面恢复可滚动', '场景法', 'P0'),
        ('TC-DETAIL-015', '国家详情', '背景滚动锁定', '验证打开模态窗口后背景不可滚动', '已打开应用首页，页面可滚动', '1. 打开任意国家详情页\n2. 尝试滚动鼠标滚轮', '-', '背景页面固定不可滚动', '场景法', 'P1'),
        ('TC-DETAIL-016', '国家详情', '错误推测-快速切换标签', '快速连续切换不同标签页', '已打开详情页', '1. 快速点击不同标签多次', '-', '页面正常响应，最终显示最后点击的标签内容，无报错', '错误推测', 'P2'),

        # ========== 详情内容验证 ==========
        ('TC-CONTENT-001', '详情内容', '签证信息-类型', '验证日本签证类型显示正确', '已打开日本详情页-签证信息', '1. 查看签证类型', '-', '显示"签证类型：旅游签证"', '场景法', 'P1'),
        ('TC-CONTENT-002', '详情内容', '签证信息-停留时间', '验证日本停留时间显示正确', '已打开日本详情页-签证信息', '1. 查看停留时间', '-', '显示"停留时间：15天"', '场景法', 'P1'),
        ('TC-CONTENT-003', '详情内容', '签证信息-难度', '验证日本办理难度显示正确', '已打开日本详情页-签证信息', '1. 查看办理难度', '-', '显示"办理难度：中等"', '场景法', 'P1'),
        ('TC-CONTENT-004', '详情内容', '签证信息-申请材料', '验证日本申请材料列表完整', '已打开日本详情页-签证信息', '1. 查看申请材料列表', '-', '显示7项材料：护照原件、签证申请表、照片、在职证明、银行流水、机票、酒店', '场景法', 'P1'),
        ('TC-CONTENT-005', '详情内容', '签证信息-申请建议', '验证日本申请建议显示正确', '已打开日本详情页-签证信息', '1. 查看申请建议', '-', '显示"建议提前1-2个月申请，可通过旅行社代办简化材料"', '场景法', 'P1'),
        ('TC-CONTENT-006', '详情内容', '货币汇率-名称代码', '验证日本货币名称和代码', '已打开日本详情页-货币汇率', '1. 查看货币名称和代码', '-', '显示"货币名称：日元 (JPY)"', '场景法', 'P1'),
        ('TC-CONTENT-007', '详情内容', '货币汇率-符号', '验证日本货币符号', '已打开日本详情页-货币汇率', '1. 查看货币符号', '-', '显示"货币符号：¥"', '场景法', 'P1'),
        ('TC-CONTENT-008', '详情内容', '货币汇率-汇率', '验证日本汇率参考', '已打开日本详情页-货币汇率', '1. 查看汇率参考', '-', '显示"1日元 ≈ 0.048人民币"', '场景法', 'P1'),
        ('TC-CONTENT-009', '详情内容', '货币汇率-支付建议', '验证日本支付建议', '已打开日本详情页-货币汇率', '1. 查看支付建议', '-', '显示关于现金和信用卡的使用建议', '场景法', 'P1'),
        ('TC-CONTENT-010', '详情内容', '语言交流-官方语言', '验证日本官方语言', '已打开日本详情页-语言交流', '1. 查看官方语言', '-', '显示"官方语言：日语"', '场景法', 'P1'),
        ('TC-CONTENT-011', '详情内容', '语言交流-问候语', '验证日本常用问候语列表', '已打开日本详情页-语言交流', '1. 查看常用问候语', '-', '显示4条问候语：你好、谢谢、对不起、再见及对应日语', '场景法', 'P1'),
        ('TC-CONTENT-012', '详情内容', '语言交流-实用短语', '验证日本实用短语列表', '已打开日本详情页-语言交流', '1. 查看实用短语', '-', '显示3条短语：多少钱、不会说日语、厕所在哪里及对应日语', '场景法', 'P1'),
        ('TC-CONTENT-013', '详情内容', '习俗禁忌-条目数', '验证日本习俗条目数量', '已打开日本详情页-习俗禁忌', '1. 查看习俗禁忌内容', '-', '显示3个条目：礼仪文化、用餐礼仪、公共行为', '场景法', 'P1'),
        ('TC-CONTENT-014', '详情内容', '玩法路线-条目数', '验证日本路线数量', '已打开日本详情页-玩法路线', '1. 查看玩法路线', '-', '显示2条路线：经典关西5日游、东京深度7日游', '场景法', 'P1'),
        ('TC-CONTENT-015', '详情内容', '玩法路线-时长', '验证路线时长显示', '已打开日本详情页-玩法路线', '1. 查看第一条路线时长', '-', '显示"5天4夜"标签', '场景法', 'P1'),
        ('TC-CONTENT-016', '详情内容', '玩法路线-亮点', '验证路线亮点显示', '已打开日本详情页-玩法路线', '1. 查看第一条路线亮点', '-', '显示"大阪城 • 京都金阁寺 • 奈良公园 • 神户牛肉"', '场景法', 'P1'),
        ('TC-CONTENT-017', '详情内容', '高频词汇-数量', '验证日本高频词汇数量', '已打开日本详情页-高频词汇', '1. 查看高频词汇网格', '-', '显示6个词汇：寿司、拉面、温泉、新干线、便利店、地铁', '场景法', 'P1'),
        ('TC-CONTENT-018', '详情内容', '高频词汇-对照', '验证高频词汇中英对照', '已打开日本详情页-高频词汇', '1. 查看"寿司"词汇项', '-', '显示"寿司"和"Sushi"', '场景法', 'P1'),
        ('TC-CONTENT-019', '详情内容', '多国数据一致性', '验证不同国家详情数据不混淆', '已打开应用首页', '1. 点击日本卡片\n2. 查看签证类型\n3. 关闭后点击泰国卡片\n4. 查看签证类型', '-', '日本显示"旅游签证"，泰国显示"落地签/电子签"，数据不混淆', '场景法', 'P0'),
        ('TC-CONTENT-020', '详情内容', '肯尼亚特色词汇', '验证肯尼亚高频词汇包含野生动物术语', '已打开肯尼亚详情页-高频词汇', '1. 查看高频词汇', '-', '包含狮子、大象、犀牛、水牛、safari、国家公园等东非特色词汇', '场景法', 'P1'),

        # ========== 响应式设计 ==========
        ('TC-RESP-001', '响应式设计', '桌面端布局', '验证桌面端（1920px）布局正常', '使用桌面浏览器', '1. 设置浏览器宽度为1920px\n2. 刷新页面', '-', '国家卡片以网格形式排列，每行显示多个卡片，布局美观', '场景法', 'P1'),
        ('TC-RESP-002', '响应式设计', '平板端布局', '验证平板端（768px）布局正常', '使用平板或模拟器', '1. 设置浏览器宽度为768px\n2. 刷新页面', '-', '国家卡片每行显示2个，布局自适应，无元素重叠', '场景法', 'P1'),
        ('TC-RESP-003', '响应式设计', '手机端布局', '验证手机端（375px）布局正常', '使用手机或模拟器', '1. 设置浏览器宽度为375px\n2. 刷新页面', '-', '国家卡片每行显示1个，搜索框和筛选标签正常显示，无元素溢出', '场景法', 'P1'),
        ('TC-RESP-004', '响应式设计', '模态窗口响应式', '验证手机端模态窗口显示正常', '使用手机浏览器', '1. 点击任意国家卡片', '-', '模态窗口宽度适配屏幕，内容可滚动，标签可点击', '场景法', 'P1'),
        ('TC-RESP-005', '响应式设计', '横屏模式', '验证手机横屏布局正常', '使用手机浏览器', '1. 将手机旋转为横屏\n2. 刷新页面', '-', '页面布局自适应横屏宽度，无元素被截断', '错误推测', 'P2'),

        # ========== 键盘交互 ==========
        ('TC-KB-001', '键盘交互', 'Enter搜索', '在搜索框按Enter键触发搜索', '搜索框已聚焦', '1. 输入"法国"\n2. 按Enter键', '法国', '显示法国的国家卡片', '场景法', 'P1'),
        ('TC-KB-002', '键盘交互', 'ESC关闭弹窗', '按ESC键关闭模态窗口', '模态窗口已打开', '1. 按ESC键', '-', '模态窗口关闭', '场景法', 'P0'),
        ('TC-KB-003', '键盘交互', 'ESC无弹窗', '无弹窗时按ESC无异常', '首页无弹窗', '1. 按ESC键', '-', '页面无异常反应', '错误推测', 'P2'),

        # ========== 错误处理 ==========
        ('TC-ERR-001', '错误处理', '图片加载失败', '模拟国旗图片加载失败', '网络可拦截', '1. 拦截flagcdn.com域名请求\n2. 刷新页面', '-', '页面其他功能正常，图片位置可能显示空白或alt文本', '错误推测', 'P2'),
        ('TC-ERR-002', '错误处理', 'Google Fonts加载失败', '模拟Google Fonts加载失败', '网络可拦截', '1. 拦截fonts.googleapis.com请求\n2. 刷新页面', '-', '页面使用系统默认字体正常显示，功能不受影响', '错误推测', 'P2'),
        ('TC-ERR-003', '错误处理', '离线访问', '断网后访问已加载页面', '页面已加载', '1. 断开网络\n2. 点击国家卡片\n3. 切换标签', '-', '页面基础功能可用（数据在本地JS中），仅外部图片可能无法加载', '错误推测', 'P2'),
        ('TC-ERR-004', '错误处理', '快速点击卡片', '快速连续点击同一国家卡片', '已打开应用首页', '1. 快速双击日本卡片', '-', '仅打开一个模态窗口，不出现多个重叠窗口', '错误推测', 'P2'),
        ('TC-ERR-005', '错误处理', 'JS异常', '验证页面无JavaScript报错', '打开浏览器开发者工具', '1. 刷新首页\n2. 执行搜索\n3. 打开详情\n4. 切换标签\n5. 关闭弹窗', '-', 'Console中无红色报错信息', '错误推测', 'P1'),

        # ========== 场景法 - 端到端场景 ==========
        ('TC-SCENE-001', '场景测试', '主流程-浏览日本信息', '用户浏览日本的完整流程', '已打开应用首页', '1. 在搜索框输入"日本"\n2. 点击搜索\n3. 点击日本卡片\n4. 查看签证信息\n5. 切换到货币汇率\n6. 切换到高频词汇\n7. 点击X关闭', '日本', '顺利查看日本的所有信息并关闭弹窗', '场景法', 'P0'),
        ('TC-SCENE-002', '场景测试', '主流程-浏览肯尼亚信息', '用户浏览肯尼亚的完整流程', '已打开应用首页', '1. 点击"非洲"筛选\n2. 点击肯尼亚卡片\n3. 查看高频词汇（野生动物词汇）\n4. 切换到玩法路线\n5. 按ESC关闭', '-', '顺利查看肯尼亚的野生动物词汇和旅游路线', '场景法', 'P0'),
        ('TC-SCENE-003', '场景测试', '筛选+查看流程', '用户筛选欧洲后查看法国信息', '已打开应用首页', '1. 点击"欧洲"筛选\n2. 点击法国卡片\n3. 查看习俗禁忌\n4. 查看语言交流\n5. 关闭弹窗', '-', '顺利查看法国的文化和语言信息', '场景法', 'P0'),
        ('TC-SCENE-004', '场景测试', '搜索无结果流程', '用户搜索不存在的国家', '已打开应用首页', '1. 在搜索框输入"美国"\n2. 点击搜索\n3. 观察提示信息', '美国', '显示"没有找到符合条件的国家"，用户可重新搜索', '场景法', 'P1'),
        ('TC-SCENE-005', '场景测试', '筛选无结果流程', '用户筛选无数据的大洲', '已打开应用首页', '1. 点击"美洲"筛选\n2. 观察提示信息\n3. 点击"全部"恢复', '-', '显示无结果提示，切换回"全部"后正常显示所有国家', '场景法', 'P1'),
        ('TC-SCENE-006', '场景测试', '多国家对比流程', '用户对比两个国家的签证信息', '已打开应用首页', '1. 点击日本卡片\n2. 查看签证信息\n3. 关闭弹窗\n4. 点击泰国卡片\n5. 查看签证信息\n6. 关闭弹窗', '-', '两个国家的签证信息正确显示，无数据混淆', '场景法', 'P1'),
    ]

    for row_data in test_cases:
        ws.append(row_data)

    # 设置数据区域样式
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.border = thin_border
            cell.alignment = cell_alignment
            # 优先级字体颜色
            if cell.column == 10:  # 优先级列
                priority = cell.value
                if priority in priority_styles:
                    cell.font = priority_styles[priority]
                    cell.alignment = center_alignment
            # 测试方法背景色
            elif cell.column == 9:  # 测试方法列
                method = cell.value
                if method in method_fills:
                    cell.fill = method_fills[method]
                cell.alignment = center_alignment
            else:
                cell.font = cell_font

    # 设置列宽
    col_widths = [12, 12, 12, 35, 25, 30, 20, 35, 10, 8]
    for i, width in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width

    # 设置行高
    for row in range(2, ws.max_row + 1):
        ws.row_dimensions[row].height = 60

    ws.row_dimensions[1].height = 30
    ws.freeze_panes = 'A2'

    # ==================== Sheet2: 等价类分析 ====================
    ws2 = wb.create_sheet('等价类分析')
    eq_headers = ['模块', '输入项', '有效等价类', '代表值', '无效等价类', '代表值', '覆盖用例']
    ws2.append(eq_headers)
    for col_num, header in enumerate(eq_headers, 1):
        cell = ws2.cell(row=1, column=col_num)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border

    eq_data = [
        ('搜索功能', '搜索关键词', '存在的中文国家名', '日本', '不存在的国家名', '德国', 'TC-SEARCH-001/010'),
        ('搜索功能', '搜索关键词', '存在的英文国家名', 'Japan', '特殊字符', '!@#', 'TC-SEARCH-003/011'),
        ('搜索功能', '搜索关键词', '部分中文匹配', '日', '纯数字', '123', 'TC-SEARCH-005/012'),
        ('搜索功能', '搜索关键词', '部分英文匹配', 'Ja', '空字符串', '', 'TC-SEARCH-007/013'),
        ('搜索功能', '搜索关键词', '大小写混合', 'jaPaN', '超长字符串(50字符)', 'a'*50, 'TC-SEARCH-009/016'),
        ('区域筛选', '筛选条件', '有数据的大洲-亚洲', '亚洲', '无数据的大洲-美洲', '美洲', 'TC-FILTER-002/005'),
        ('区域筛选', '筛选条件', '有数据的大洲-欧洲', '欧洲', '无数据的大洲-大洋洲', '大洋洲', 'TC-FILTER-003/006'),
        ('区域筛选', '筛选条件', '有数据的大洲-非洲', '非洲', '-', '-', 'TC-FILTER-004'),
        ('区域筛选', '筛选条件', '全部大洲', '全部', '-', '-', 'TC-FILTER-001'),
    ]

    for row_data in eq_data:
        ws2.append(row_data)

    for row in ws2.iter_rows(min_row=2, max_row=ws2.max_row):
        for cell in row:
            cell.border = thin_border
            cell.alignment = cell_alignment
            cell.font = cell_font

    ws2_col_widths = [12, 12, 20, 15, 20, 15, 15]
    for i, width in enumerate(ws2_col_widths, 1):
        ws2.column_dimensions[get_column_letter(i)].width = width

    ws2.freeze_panes = 'A2'

    # ==================== Sheet3: 边界值分析 ====================
    ws3 = wb.create_sheet('边界值分析')
    bv_headers = ['模块', '输入项', '范围限制', 'min-1', 'min', 'min+1', 'nominal', 'max-1', 'max', 'max+1', '覆盖用例']
    ws3.append(bv_headers)
    for col_num, header in enumerate(bv_headers, 1):
        cell = ws3.cell(row=1, column=col_num)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border

    bv_data = [
        ('搜索功能', '搜索关键词长度', '1~20字符(合理)', '空字符串', '1字符(法)', '2字符(日本)', '5字符(Thailand)', '19字符', '20字符', '50字符(超长)', 'TC-SEARCH-013/014/015/016'),
        ('搜索功能', '搜索结果数量', '0~4个国家', '-', '0个(无结果)', '1个(法国)', '2个(亚洲筛选)', '3个', '4个(全部)', '-', 'TC-FILTER-001~006'),
        ('国家卡片', '国家卡片数量', '当前4个国家', '-', '1个卡片', '2个卡片', '4个卡片', '-', '-', '-', 'TC-HOME-004'),
        ('详情标签', '标签页数量', '6个标签', '-', '1个', '2个', '6个', '-', '-', '-', 'TC-DETAIL-004'),
        ('签证材料', '申请材料数量', '日本:7项', '-', '1项', '2项', '7项', '-', '-', '-', 'TC-CONTENT-004'),
        ('问候语', '问候语数量', '日本:4条', '-', '1条', '2条', '4条', '-', '-', '-', 'TC-CONTENT-011'),
        ('高频词汇', '词汇数量', '日本:6个', '-', '1个', '2个', '6个', '-', '-', '-', 'TC-CONTENT-017'),
    ]

    for row_data in bv_data:
        ws3.append(row_data)

    for row in ws3.iter_rows(min_row=2, max_row=ws3.max_row):
        for cell in row:
            cell.border = thin_border
            cell.alignment = cell_alignment
            cell.font = cell_font

    ws3_col_widths = [12, 14, 18, 12, 12, 12, 18, 10, 10, 14, 18]
    for i, width in enumerate(ws3_col_widths, 1):
        ws3.column_dimensions[get_column_letter(i)].width = width

    ws3.freeze_panes = 'A2'

    # ==================== Sheet4: 场景分析 ====================
    ws4 = wb.create_sheet('场景分析')
    scene_headers = ['场景编号', '场景类型', '场景名称', '触发条件', '操作步骤', '预期结果', '覆盖用例', '优先级']
    ws4.append(scene_headers)
    for col_num, header in enumerate(scene_headers, 1):
        cell = ws4.cell(row=1, column=col_num)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border

    scene_data = [
        ('SC-01', '基本流', '浏览日本完整信息', '用户想了解日本旅游信息', '1.搜索"日本"→2.点击日本卡片→3.查看签证→4.查看货币→5.查看词汇→6.关闭', '完整查看日本所有信息', 'TC-SCENE-001', 'P0'),
        ('SC-02', '基本流', '浏览肯尼亚野生动物信息', '用户对非洲safari感兴趣', '1.筛选非洲→2.点击肯尼亚→3.查看高频词汇→4.查看路线→5.ESC关闭', '查看肯尼亚特色词汇和路线', 'TC-SCENE-002', 'P0'),
        ('SC-03', '基本流', '筛选欧洲查看法国文化', '用户计划去法国旅行', '1.筛选欧洲→2.点击法国→3.查看习俗→4.查看语言→5.关闭', '了解法国文化和语言', 'TC-SCENE-003', 'P0'),
        ('SC-04', '备选流-搜索', '搜索后筛选组合查询', '用户想精确查找', '1.筛选亚洲→2.搜索"泰国"→3.查看详情→4.关闭', '显示泰国详情', 'TC-FILTER-008', 'P1'),
        ('SC-05', '备选流-搜索', '搜索无结果后重新搜索', '用户首次输错国家名', '1.搜索"美国"→2.看到无结果提示→3.清空搜索→4.搜索"日本"→5.查看详情', '最终成功查看日本信息', 'TC-SCENE-004', 'P1'),
        ('SC-06', '异常流-筛选', '筛选无数据大洲', '用户误点无数据大洲', '1.点击"美洲"→2.看到无结果提示→3.点击"全部"恢复', '恢复正常显示全部国家', 'TC-SCENE-005', 'P1'),
        ('SC-07', '异常流-筛选+搜索', '筛选和搜索无交集', '用户在错误筛选下搜索', '1.筛选欧洲→2.搜索"日本"→3.看到无结果提示', '显示无结果，提示信息正确', 'TC-FILTER-009', 'P1'),
        ('SC-08', '备选流-对比', '对比两个国家签证信息', '用户需要对比选择目的地', '1.点击日本→2.看签证→3.关闭→4.点击泰国→5.看签证→6.关闭', '两个国家签证信息正确显示且不混淆', 'TC-SCENE-006', 'P1'),
        ('SC-09', '异常流-网络', '网络异常下使用应用', '用户网络不稳定', '1.断网→2.刷新页面→3.点击卡片→4.切换标签', '基础功能可用，仅外部图片可能加载失败', 'TC-ERR-003', 'P2'),
        ('SC-10', '异常流-操作', '快速连续操作', '用户操作过快', '1.快速双击卡片→2.快速连续切换标签', '页面正常响应，无报错或异常', 'TC-ERR-004/TC-DETAIL-016', 'P2'),
    ]

    for row_data in scene_data:
        ws4.append(row_data)

    for row in ws4.iter_rows(min_row=2, max_row=ws4.max_row):
        for cell in row:
            cell.border = thin_border
            cell.alignment = cell_alignment
            cell.font = cell_font
            if cell.column == 8:
                priority = cell.value
                if priority in priority_styles:
                    cell.font = priority_styles[priority]
                cell.alignment = center_alignment

    ws4_col_widths = [10, 14, 20, 22, 35, 28, 16, 10]
    for i, width in enumerate(ws4_col_widths, 1):
        ws4.column_dimensions[get_column_letter(i)].width = width

    for row in range(2, ws4.max_row + 1):
        ws4.row_dimensions[row].height = 70

    ws4.freeze_panes = 'A2'

    # ==================== Sheet5: 错误推测清单 ====================
    ws5 = wb.create_sheet('错误推测清单')
    err_headers = ['序号', '推测类型', '推测描述', '风险场景', '对应测试点', '覆盖用例', '优先级']
    ws5.append(err_headers)
    for col_num, header in enumerate(err_headers, 1):
        cell = ws5.cell(row=1, column=col_num)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border

    err_data = [
        ('1', '并发操作', '快速连续点击同一卡片可能导致弹窗重复打开', '用户双击或快速点击', '验证快速点击只打开一个弹窗', 'TC-ERR-004', 'P2'),
        ('2', '并发操作', '快速连续切换标签可能导致内容错乱', '用户 impatient 快速点击', '验证快速切换标签最终显示正确', 'TC-DETAIL-016', 'P2'),
        ('3', '快速输入', '搜索框快速连续输入可能导致过滤逻辑异常', '用户打字速度快', '验证快速输入后搜索结果正确', 'TC-SEARCH-019', 'P2'),
        ('4', '网络异常', '外部图片资源加载失败时页面布局错乱', 'CDN不可用或网络差', '验证图片加载失败时页面仍可用', 'TC-ERR-001', 'P2'),
        ('5', '网络异常', 'Google Fonts加载失败导致文字不可见', 'fonts.googleapis.com被拦截', '验证字体加载失败时使用系统默认字体', 'TC-ERR-002', 'P2'),
        ('6', '网络异常', '断网后本地数据功能是否可用', '用户完全离线', '验证JS本地数据功能正常', 'TC-ERR-003', 'P2'),
        ('7', 'JS异常', '页面交互过程中出现JavaScript错误', '代码逻辑缺陷', '验证Console无报错', 'TC-ERR-005', 'P1'),
        ('8', '空状态', '筛选或搜索无结果时页面展示不完整', '数据为空时的边界', '验证无结果提示信息正确显示', 'TC-FILTER-005/006', 'P1'),
        ('9', '横屏适配', '手机横屏时模态窗口内容被截断', '用户旋转手机', '验证横屏模态窗口正常显示', 'TC-RESP-005', 'P2'),
        ('10', '键盘冲突', 'ESC键在有/无弹窗时的行为不一致', '用户误按ESC', '验证无弹窗时按ESC无异常', 'TC-KB-003', 'P2'),
        ('11', '数据一致性', '连续打开不同国家详情时数据混淆', '缓存或DOM未清理', '验证多国数据不混淆', 'TC-CONTENT-019', 'P0'),
        ('12', '图片溢出', '超长国家名或简介导致卡片布局破坏', '极端文本长度', '验证当前4个国家卡片布局正常', 'TC-HOME-004', 'P1'),
    ]

    for row_data in err_data:
        ws5.append(row_data)

    for row in ws5.iter_rows(min_row=2, max_row=ws5.max_row):
        for cell in row:
            cell.border = thin_border
            cell.alignment = cell_alignment
            cell.font = cell_font
            if cell.column == 7:
                priority = cell.value
                if priority in priority_styles:
                    cell.font = priority_styles[priority]
                cell.alignment = center_alignment

    ws5_col_widths = [8, 12, 28, 20, 25, 16, 10]
    for i, width in enumerate(ws5_col_widths, 1):
        ws5.column_dimensions[get_column_letter(i)].width = width

    for row in range(2, ws5.max_row + 1):
        ws5.row_dimensions[row].height = 50

    ws5.freeze_panes = 'A2'

    # 保存文件
    wb.save(output_path)
    print(f"Excel 测试用例已生成: {output_path}")
    print(f"  - 测试用例总数: {len(test_cases)}")
    print(f"  - Sheet数量: {len(wb.sheetnames)} ({', '.join(wb.sheetnames)})")

    # 统计摘要
    p0_count = sum(1 for tc in test_cases if tc[9] == 'P0')
    p1_count = sum(1 for tc in test_cases if tc[9] == 'P1')
    p2_count = sum(1 for tc in test_cases if tc[9] == 'P2')
    p3_count = sum(1 for tc in test_cases if tc[9] == 'P3')

    eq_count = sum(1 for tc in test_cases if tc[8] == '等价类')
    bv_count = sum(1 for tc in test_cases if tc[8] == '边界值')
    scene_count = sum(1 for tc in test_cases if tc[8] == '场景法')
    err_count = sum(1 for tc in test_cases if tc[8] == '错误推测')

    print(f"  - 优先级分布: P0={p0_count}, P1={p1_count}, P2={p2_count}, P3={p3_count}")
    print(f"  - 测试方法分布: 等价类={eq_count}, 边界值={bv_count}, 场景法={scene_count}, 错误推测={err_count}")


def create_xmind_test_cases(output_path):
    """生成 XMind 格式测试用例"""
    # XMind 标记器引用
    # 优先级标记: priority-1=红色, priority-2=橙色, priority-3=黄色, priority-4=绿色
    MARKER_PRIORITY = {
        'P0': 'priority-1',
        'P1': 'priority-2',
        'P2': 'priority-3',
        'P3': 'priority-4',
    }

    # 加载或创建工作簿
    if os.path.exists(output_path):
        workbook = xmind.load(output_path)
    else:
        workbook = xmind.load(output_path)

    sheet = workbook.getPrimarySheet()
    sheet.setTitle('测试用例')

    root = sheet.getRootTopic()
    root.setTitle('出境旅游信息查询 - 测试用例')

    # 辅助函数：添加子主题并设置标记
    def add_topic(parent, title, priority=None):
        topic = parent.addSubTopic()
        topic.setTitle(title)
        if priority and priority in MARKER_PRIORITY:
            topic.addMarker(MARKER_PRIORITY[priority])
        return topic

    # ========== 1. 首页展示 ==========
    mod1 = add_topic(root, '首页展示', 'P0')
    add_topic(mod1, 'TC-HOME-001 验证页面标题正确显示', 'P0')
    add_topic(mod1, 'TC-HOME-002 验证Logo和搜索区域显示', 'P0')
    add_topic(mod1, 'TC-HOME-003 验证6个区域筛选标签完整显示', 'P0')
    add_topic(mod1, 'TC-HOME-004 验证默认展示4个国家卡片', 'P0')
    add_topic(mod1, 'TC-HOME-005 验证国家卡片信息正确性', 'P1')
    add_topic(mod1, 'TC-HOME-006 验证国旗图片正常加载', 'P1')

    # ========== 2. 搜索功能 ==========
    mod2 = add_topic(root, '搜索功能', 'P0')
    eq2 = add_topic(mod2, '等价类测试', 'P0')
    add_topic(eq2, 'TC-SEARCH-001 中文全名"日本" → 显示日本', 'P0')
    add_topic(eq2, 'TC-SEARCH-002 中文全名"法国" → 显示法国', 'P0')
    add_topic(eq2, 'TC-SEARCH-003 英文全名"Japan" → 显示日本', 'P0')
    add_topic(eq2, 'TC-SEARCH-004 英文全名"France" → 显示法国', 'P0')
    add_topic(eq2, 'TC-SEARCH-005 部分中文"日" → 显示日本', 'P1')
    add_topic(eq2, 'TC-SEARCH-006 部分中文"泰" → 显示泰国', 'P1')
    add_topic(eq2, 'TC-SEARCH-007 部分英文"Ja" → 显示日本', 'P1')
    add_topic(eq2, 'TC-SEARCH-008 部分英文"Ke" → 显示肯尼亚', 'P1')
    add_topic(eq2, 'TC-SEARCH-009 大小写混合"jaPaN" → 显示日本', 'P1')
    add_topic(eq2, 'TC-SEARCH-010 不存在国家"德国" → 无结果', 'P1')
    add_topic(eq2, 'TC-SEARCH-011 特殊字符"!@#" → 无结果', 'P1')
    add_topic(eq2, 'TC-SEARCH-012 纯数字"123" → 无结果', 'P1')
    add_topic(eq2, 'TC-SEARCH-013 空字符串 → 显示全部', 'P0')

    bv2 = add_topic(mod2, '边界值测试', 'P1')
    add_topic(bv2, 'TC-SEARCH-014 1个字符"法" → 显示法国', 'P1')
    add_topic(bv2, 'TC-SEARCH-015 正常长度"Thailand" → 显示泰国', 'P1')
    add_topic(bv2, 'TC-SEARCH-016 50字符超长 → 无结果不报错', 'P2')

    scene2 = add_topic(mod2, '场景测试', 'P1')
    add_topic(scene2, 'TC-SEARCH-017 Enter键触发搜索', 'P1')
    add_topic(scene2, 'TC-SEARCH-018 点击搜索按钮触发搜索', 'P1')

    err2 = add_topic(mod2, '错误推测', 'P2')
    add_topic(err2, 'TC-SEARCH-019 快速连续输入 → 正常响应', 'P2')

    # ========== 3. 区域筛选 ==========
    mod3 = add_topic(root, '区域筛选', 'P0')
    eq3 = add_topic(mod3, '等价类测试', 'P0')
    add_topic(eq3, 'TC-FILTER-001 "全部" → 4个国家', 'P0')
    add_topic(eq3, 'TC-FILTER-002 "亚洲" → 日本+泰国', 'P0')
    add_topic(eq3, 'TC-FILTER-003 "欧洲" → 法国', 'P0')
    add_topic(eq3, 'TC-FILTER-004 "非洲" → 肯尼亚', 'P0')
    add_topic(eq3, 'TC-FILTER-005 "美洲" → 无结果', 'P1')
    add_topic(eq3, 'TC-FILTER-006 "大洋洲" → 无结果', 'P1')

    scene3 = add_topic(mod3, '场景测试', 'P1')
    add_topic(scene3, 'TC-FILTER-007 切换标签高亮状态更新', 'P1')
    add_topic(scene3, 'TC-FILTER-008 筛选亚洲+搜索日本 → 显示日本', 'P1')
    add_topic(scene3, 'TC-FILTER-009 筛选欧洲+搜索日本 → 无结果', 'P1')
    add_topic(scene3, 'TC-FILTER-010 筛选非洲+搜索Kenya → 显示肯尼亚', 'P1')

    # ========== 4. 国家详情模态窗口 ==========
    mod4 = add_topic(root, '国家详情弹窗', 'P0')
    open4 = add_topic(mod4, '打开弹窗', 'P0')
    add_topic(open4, 'TC-DETAIL-001 点击日本卡片打开弹窗', 'P0')
    add_topic(open4, 'TC-DETAIL-002 点击肯尼亚卡片打开弹窗', 'P0')
    add_topic(open4, 'TC-DETAIL-003 验证头部信息正确', 'P0')

    tab4 = add_topic(mod4, '标签页切换', 'P0')
    add_topic(tab4, 'TC-DETAIL-004 6个标签完整显示', 'P0')
    add_topic(tab4, 'TC-DETAIL-005 点击"货币汇率"显示货币内容', 'P0')
    add_topic(tab4, 'TC-DETAIL-006 点击"语言交流"显示语言内容', 'P0')
    add_topic(tab4, 'TC-DETAIL-007 点击"习俗禁忌"显示习俗内容', 'P0')
    add_topic(tab4, 'TC-DETAIL-008 点击"玩法路线"显示路线内容', 'P0')
    add_topic(tab4, 'TC-DETAIL-009 点击"高频词汇"显示词汇内容', 'P0')
    add_topic(tab4, 'TC-DETAIL-010 点击"签证信息"返回签证内容', 'P0')
    add_topic(tab4, 'TC-DETAIL-011 标签高亮状态正确更新', 'P1')

    close4 = add_topic(mod4, '关闭弹窗', 'P0')
    add_topic(close4, 'TC-DETAIL-012 点击X按钮关闭', 'P0')
    add_topic(close4, 'TC-DETAIL-013 点击外部区域关闭', 'P0')
    add_topic(close4, 'TC-DETAIL-014 按ESC键关闭', 'P0')

    scroll4 = add_topic(mod4, '滚动与状态', 'P1')
    add_topic(scroll4, 'TC-DETAIL-015 打开弹窗后背景不可滚动', 'P1')
    add_topic(scroll4, 'TC-DETAIL-016 快速切换标签 → 正常响应', 'P2')

    # ========== 5. 详情内容验证 ==========
    mod5 = add_topic(root, '详情内容验证', 'P1')
    visa5 = add_topic(mod5, '签证信息', 'P1')
    add_topic(visa5, 'TC-CONTENT-001 签证类型正确', 'P1')
    add_topic(visa5, 'TC-CONTENT-002 停留时间正确', 'P1')
    add_topic(visa5, 'TC-CONTENT-003 办理难度正确', 'P1')
    add_topic(visa5, 'TC-CONTENT-004 申请材料列表完整', 'P1')
    add_topic(visa5, 'TC-CONTENT-005 申请建议正确', 'P1')

    currency5 = add_topic(mod5, '货币汇率', 'P1')
    add_topic(currency5, 'TC-CONTENT-006 货币名称和代码', 'P1')
    add_topic(currency5, 'TC-CONTENT-007 货币符号', 'P1')
    add_topic(currency5, 'TC-CONTENT-008 汇率参考', 'P1')
    add_topic(currency5, 'TC-CONTENT-009 支付建议', 'P1')

    lang5 = add_topic(mod5, '语言交流', 'P1')
    add_topic(lang5, 'TC-CONTENT-010 官方语言', 'P1')
    add_topic(lang5, 'TC-CONTENT-011 常用问候语列表', 'P1')
    add_topic(lang5, 'TC-CONTENT-012 实用短语列表', 'P1')

    custom5 = add_topic(mod5, '习俗禁忌', 'P1')
    add_topic(custom5, 'TC-CONTENT-013 习俗条目数量正确', 'P1')

    route5 = add_topic(mod5, '玩法路线', 'P1')
    add_topic(route5, 'TC-CONTENT-014 路线数量正确', 'P1')
    add_topic(route5, 'TC-CONTENT-015 路线时长显示', 'P1')
    add_topic(route5, 'TC-CONTENT-016 路线亮点显示', 'P1')

    vocab5 = add_topic(mod5, '高频词汇', 'P1')
    add_topic(vocab5, 'TC-CONTENT-017 词汇数量正确', 'P1')
    add_topic(vocab5, 'TC-CONTENT-018 中英对照正确', 'P1')
    add_topic(vocab5, 'TC-CONTENT-020 肯尼亚特色野生动物词汇', 'P1')

    data5 = add_topic(mod5, '数据一致性', 'P0')
    add_topic(data5, 'TC-CONTENT-019 多国数据不混淆', 'P0')

    # ========== 6. 响应式设计 ==========
    mod6 = add_topic(root, '响应式设计', 'P1')
    add_topic(mod6, 'TC-RESP-001 桌面端1920px布局正常', 'P1')
    add_topic(mod6, 'TC-RESP-002 平板端768px布局正常', 'P1')
    add_topic(mod6, 'TC-RESP-003 手机端375px布局正常', 'P1')
    add_topic(mod6, 'TC-RESP-004 手机端模态窗口适配', 'P1')
    add_topic(mod6, 'TC-RESP-005 手机横屏布局正常', 'P2')

    # ========== 7. 键盘交互 ==========
    mod7 = add_topic(root, '键盘交互', 'P0')
    add_topic(mod7, 'TC-KB-001 Enter键触发搜索', 'P1')
    add_topic(mod7, 'TC-KB-002 ESC键关闭模态窗口', 'P0')
    add_topic(mod7, 'TC-KB-003 无弹窗时按ESC无异常', 'P2')

    # ========== 8. 错误处理 ==========
    mod8 = add_topic(root, '错误处理', 'P1')
    add_topic(mod8, 'TC-ERR-001 国旗图片加载失败', 'P2')
    add_topic(mod8, 'TC-ERR-002 Google Fonts加载失败', 'P2')
    add_topic(mod8, 'TC-ERR-003 断网后本地功能可用', 'P2')
    add_topic(mod8, 'TC-ERR-004 快速点击卡片只开一个弹窗', 'P2')
    add_topic(mod8, 'TC-ERR-005 Console无JS报错', 'P1')

    # ========== 9. 场景测试(端到端) ==========
    mod9 = add_topic(root, '端到端场景', 'P0')
    add_topic(mod9, 'TC-SCENE-001 浏览日本完整信息流程', 'P0')
    add_topic(mod9, 'TC-SCENE-002 浏览肯尼亚野生动物信息', 'P0')
    add_topic(mod9, 'TC-SCENE-003 筛选欧洲查看法国文化', 'P0')
    add_topic(mod9, 'TC-SCENE-004 搜索不存在国家流程', 'P1')
    add_topic(mod9, 'TC-SCENE-005 筛选无数据大洲流程', 'P1')
    add_topic(mod9, 'TC-SCENE-006 对比两个国家签证信息', 'P1')

    xmind.save(workbook, output_path)
    print(f"XMind 测试用例已生成: {output_path}")


def main():
    output_dir = os.path.dirname(os.path.abspath(__file__))
    excel_path = os.path.join(output_dir, '测试用例_出境旅游信息查询.xlsx')
    xmind_path = os.path.join(output_dir, '测试用例_出境旅游信息查询.xmind')

    print("=" * 60)
    print("出境旅游信息查询 Web应用 - 测试用例生成器")
    print("测试设计方法: 等价类 | 边界值 | 场景法 | 错误推测")
    print("=" * 60)
    print()

    create_excel_test_cases(excel_path)
    print()
    create_xmind_test_cases(xmind_path)

    print()
    print("=" * 60)
    print("生成完成!")
    print(f"Excel: {excel_path}")
    print(f"XMind: {xmind_path}")
    print("=" * 60)


if __name__ == '__main__':
    main()
