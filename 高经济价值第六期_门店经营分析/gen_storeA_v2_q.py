#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""G6-ZX-004 连锁门店经营分析与亏损店处置 v2（返修升级版）
升级要点：
1) 数据量 3900+ 行（30 门店 × 90 天日流水 2700 行 + 采购 360 + 花名册 240 + 报销 180 + 其他）
2) 成本分摊制度 v1.0/v2.0/v3.0 三版并行，按业务时点归属；顶置红字"按 v3.0 执行"为文控误标
3) 关键规则藏在《经营分析会会议纪要》（管理分摊基数、加盟口径、免租冲回、奖金权重）
4) P&L 6 家门店存在口径错误（3 家收入按平台实收、2 家套用 v3.0 费率、1 家未冲回免租）
5) 加盟口径旧版财务规则与新版纪要冲突（显著性反转）
"""
import os
import random
from datetime import date, timedelta

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter

random.seed(20260217)

BASE = '/workspace/高经济价值第六期_门店经营分析'
OUT = f'{BASE}/门店经营分析_v1_题目.xlsx'

HDR_FILL = PatternFill('solid', fgColor='2F5597')
HDR_FONT = Font(name='宋体', bold=True, color='FFFFFF', size=10)
BODY = Font(name='宋体', size=9)
RED_FONT = Font(name='宋体', bold=True, color='C00000', size=9)
THIN = Side(style='thin', color='CCCCCC')
BORDER = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
WRAP = Alignment(wrap_text=True, vertical='top')
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)

# 制度费率：v1.0(2024) / v2.0(2025起，Q1适用) / v3.0(2026-04-01起，Q1不适用)
RATE_V1 = dict(food=0.35, labor=0.22, util=0.03, mkt=0.04, mgmt=0.06, comm=0.15)
RATE_V2 = dict(food=0.36, labor=0.21, util=0.03, mkt=0.04, mgmt=0.05, comm=0.18)
RATE_V3 = dict(food=0.37, labor=0.20, util=0.025, mkt=0.045, mgmt=0.04, comm=0.18)
TAX = 0.25
DEP_YEARS = 5


def r2(x):
    return round(x + 0.0, 2)


def r4(x):
    return round(x + 0.0, 4)


# 门店主数据：编号,城市,区域,商圈,面积,开业日,店长,类型,稳定月营收(万),外卖占比,租金占收入比,建店投资(万)
STORES = [
    ['SZ01', '苏州', '工业园区', '湖东邻里中心', 180, '2023-05-10', '张伟', '直营', 85, 0.30, 0.100, 235],
    ['SZ02', '苏州', '高新区', '狮山天街', 200, '2023-08-15', '李娜', '直营', 78, 0.32, 0.130, 220],
    ['SZ03', '苏州', '吴中区', '万象汇', 180, '2024-01-20', '王强', '直营', 70, 0.35, 0.160, 210],
    ['SZ04', '苏州', '相城区', '天虹百货', 170, '2024-05-08', '赵敏', '直营', 62, 0.38, 0.190, 205],
    ['SZ05', '苏州', '姑苏区', '印象城', 175, '2024-09-12', '钱进', '直营', 66, 0.34, 0.175, 208],
    ['SZ06', '苏州', '吴江区', '万象汇', 190, '2025-06-18', '冯雪', '直营', 72, 0.36, 0.150, 225],
    ['WX01', '无锡', '梁溪区', '崇安寺', 175, '2023-03-18', '陈杰', '直营', 80, 0.31, 0.100, 225],
    ['WX02', '无锡', '滨湖区', '滨湖万达', 180, '2023-09-12', '刘洋', '直营', 72, 0.34, 0.150, 215],
    ['WX03', '无锡', '新吴区', '宝龙广场', 170, '2024-03-22', '孙丽', '直营', 64, 0.37, 0.185, 200],
    ['WX04', '无锡', '惠山区', '惠山万达', 180, '2024-07-30', '周涛', '直营', 56, 0.45, 0.215, 195],
    ['WX05', '无锡', '锡山区', '锡东八佰伴', 165, '2025-11-25', '蒋萌', '加盟', 58, 0.40, 0.190, 200],
    ['NB01', '宁波', '海曙区', '天一广场', 185, '2023-04-25', '吴静', '直营', 88, 0.28, 0.090, 240],
    ['NB02', '宁波', '鄞州区', '印象城', 175, '2023-10-08', '郑磊', '直营', 74, 0.33, 0.120, 220],
    ['NB03', '宁波', '江北区', '来福士', 165, '2024-02-14', '黄芳', '直营', 68, 0.36, 0.145, 210],
    ['NB04', '宁波', '镇海区', '吾悦广场', 160, '2024-06-18', '林峰', '直营', 58, 0.40, 0.200, 195],
    ['NB05', '宁波', '北仑区', '富邦世纪', 170, '2025-04-15', '沈悦', '直营', 62, 0.35, 0.165, 202],
    ['HZ01', '杭州', '上城区', '湖滨银泰', 180, '2025-03-15', '徐明', '直营', 90, 0.30, 0.110, 250],
    ['HZ02', '杭州', '上城区', '钱江新城来福士', 200, '2025-06-01', '朱琳', '直营', 82, 0.33, 0.130, 245],
    ['HZ03', '杭州', '西湖区', '城西银泰', 190, '2025-08-20', '何军', '直营', 76, 0.36, 0.165, 238],
    ['HZ04', '杭州', '滨江区', '滨江宝龙', 185, '2025-11-15', '韩磊', '直营', 70, 0.38, 0.200, 235],
    ['HZ05', '杭州', '余杭区', '未来科技城', 175, '2025-12-28', '杨帆', '直营', 64, 0.37, 0.205, 230],
    ['HF01', '合肥', '蜀山区', '政务区银泰', 180, '2025-05-20', '高飞', '直营', 72, 0.32, 0.135, 225],
    ['HF02', '合肥', '包河区', '滨湖万达', 170, '2025-07-10', '罗敏', '直营', 66, 0.35, 0.160, 215],
    ['HF03', '合肥', '蜀山区', '三里庵国购', 160, '2025-08-20', '崔浩', '加盟', 60, 0.38, 0.180, 210],
    ['HF04', '合肥', '庐阳区', '银泰中心', 165, '2025-10-15', '邹敏', '直营', 58, 0.39, 0.220, 205],
    ['HF05', '合肥', '瑶海区', '瑶海万达', 160, '2025-12-20', '石磊', '直营', 54, 0.41, 0.215, 200],
    ['NC01', '南昌', '红谷滩区', '红谷滩万达', 160, '2024-11-05', '宋佳', '直营', 58, 0.36, 0.155, 205],
    ['NC02', '南昌', '青山湖区', '青山湖万象汇', 155, '2025-05-20', '谢峰', '直营', 52, 0.40, 0.215, 198],
    ['NC03', '南昌', '新建区', '新建吾悦', 150, '2025-12-25', '梁宇', '加盟', 48, 0.42, 0.225, 200],
    ['NC04', '南昌', '西湖区', '西湖万达', 150, '2025-12-28', '秦朗', '直营', 48, 0.42, 0.230, 200],
]

PROP_FEE = {'苏州': 26, '无锡': 22, '宁波': 25, '杭州': 30, '合肥': 18, '南昌': 14}
# 免租月数（Q1 应计租金月数）；NC04 补充条款约定 1 月免租
RENT_MONTHS = {sid: 3 for sid in [s[0] for s in STORES]}
RENT_MONTHS['NC04'] = 2

# P&L 口径错误门店
REV_ERR = {'SZ04', 'WX04', 'NC03'}      # 收入按平台实收列报（少计）
COST_ERR = {'NB04', 'HF05'}             # 套用未生效的 v3.0 费率
RENT_ERR = {'NC04'}                     # 未冲回免租期（按 3 个月计提）

MONTH_DAYS = {1: 31, 2: 28, 3: 31}
MONTH_FACTOR = {1: 0.96, 2: 0.94, 3: 1.06}
RAMP = {1: 0.88, 2: 0.95, 3: 1.02}


def is_new_store(open_str):
    return open_str >= '2025-11-01'


def q1_days():
    d = date(2026, 1, 1)
    out = []
    while d <= date(2026, 3, 31):
        out.append(d)
        d += timedelta(days=1)
    return out


DAYS = q1_days()

wb = openpyxl.Workbook()
wb.remove(wb.active)


def sheet(title, headers, rows, widths=None, banner=None, comment=None):
    ws = wb.create_sheet(title)
    if banner:
        ws.append([banner])
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
        c = ws.cell(row=1, column=1)
        c.font = RED_FONT
        c.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        ws.row_dimensions[1].height = 26
        ws.append([])
    hr = ws.max_row + 1
    ws.append(headers)
    for r in rows:
        ws.append(r)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=hr, column=c)
        cell.font = HDR_FONT
        cell.fill = HDR_FILL
        cell.alignment = CENTER
        cell.border = BORDER
        ws.column_dimensions[get_column_letter(c)].width = (widths or {}).get(c, 14)
    for row in ws.iter_rows(min_row=hr + 1, max_row=ws.max_row, max_col=len(headers)):
        for cell in row:
            cell.font = BODY
            cell.alignment = WRAP
            cell.border = BORDER
    ws.freeze_panes = ws.cell(row=hr + 1, column=1)
    if comment:
        ref, txt = comment
        cm = Comment(txt, '财务部')
        cm.width, cm.height = 320, 110
        ws[ref].comment = cm
    return ws


# ============ 计算引擎（正确口径） ============
S = {s[0]: s for s in STORES}
RENT_PER = {}
for s in STORES:
    sid, city, area, stable, rr = s[0], s[1], s[4], s[8], s[10]
    prop = PROP_FEE[city]
    per = (rr * stable * 10000 / area) - prop
    RENT_PER[sid] = round(per / 5) * 5

MONTH_RENT = {sid: (RENT_PER[sid] + PROP_FEE[S[sid][1]]) * S[sid][4] for sid in S}  # 元/月

daily_rows = []
monthly = {}   # (sid, m) -> dict
for s in STORES:
    sid, city, area, open_str, stable, deliv = s[0], s[1], s[4], s[5], s[8], s[9]
    base_day = stable * 10000 / 30
    new = is_new_store(open_str)
    for d in DAYS:
        wd = d.weekday()
        wfac = 1.28 if wd >= 5 else (1.12 if wd == 4 else 1.0)
        mfac = MONTH_FACTOR[d.month]
        rfac = RAMP[d.month] if new else 1.0
        noise = random.uniform(0.93, 1.07)
        total = base_day * wfac * mfac * rfac * noise
        dine_gross = total * (1 - deliv)
        dine_ref = dine_gross * 0.012
        dine_net = dine_gross - dine_ref
        gmv = total * deliv
        subsidy = gmv * 0.03
        daily_rows.append([str(d), sid, city, r2(dine_gross), r2(dine_ref),
                           r2(dine_net), r2(gmv), r2(subsidy)])
        k = (sid, d.month)
        a = monthly.setdefault(k, dict(dine_net=0.0, gmv=0.0))
        a['dine_net'] += r2(dine_net)
        a['gmv'] += r2(gmv)

# 月度 P&L（正确 + 错误口径）
pl_rows = []
pl_wrong = []
for s in STORES:
    sid, city, area, open_str, stable, deliv, capex = s[0], s[1], s[4], s[5], s[8], s[9], s[11]
    dep = capex * 10000 / DEP_YEARS / 12 / 10000  # 万元/月
    rent_wan_full = MONTH_RENT[sid] / 10000
    for m in (1, 2, 3):
        a = monthly[(sid, m)]
        rev = (a['dine_net'] + a['gmv']) / 10000
        gmv_wan = a['gmv'] / 10000
        rent = rent_wan_full if m <= RENT_MONTHS[sid] else 0.0
        # 若免租月在前，则 1 月免租；此处 NC04 为 1 月免租
        if sid == 'NC04' and m == 1:
            rent = 0.0
        rr_ = RATE_V2
        food = rev * rr_['food']
        labor = rev * rr_['labor']
        util = rev * rr_['util']
        mkt = rev * rr_['mkt']
        comm = gmv_wan * rr_['comm']
        mgmt = (rev - comm) * rr_['mgmt']
        ebitda = rev - food - labor - util - mkt - comm - mgmt - rent
        ebit = ebitda - dep
        tax = max(0.0, ebit) * TAX
        net = ebit - tax

        # ---- 账面（错误口径） ----
        rev_book = rev
        rate_book = rr_
        rent_book = rent
        if sid in REV_ERR:
            settled = gmv_wan - gmv_wan * 0.18 - gmv_wan * 0.05
            rev_book = (a['dine_net'] / 10000) + settled
        if sid in COST_ERR:
            rate_book = RATE_V3
        if sid in RENT_ERR:
            rent_book = rent_wan_full
        # 成本按实际发生额（以正确收入为基数），仅收入科目按账面口径
        f_b = rev * rate_book['food']
        l_b = rev * rate_book['labor']
        u_b = rev * rate_book['util']
        k_b = rev * rate_book['mkt']
        c_b = gmv_wan * rate_book['comm']
        g_b = (rev - c_b) * rate_book['mgmt']
        e_b = rev_book - f_b - l_b - u_b - k_b - c_b - g_b - rent_book
        p_b = e_b - dep
        t_b = max(0.0, p_b) * TAX
        n_b = p_b - t_b

        pl_rows.append([sid, city, f'2026-{m:02d}', r4(rev_book), r4(f_b), r4(l_b),
                        r4(rent_book), r4(u_b), r4(k_b), r4(c_b), r4(g_b), r4(dep),
                        r4(e_b), r4(p_b), r4(t_b), r4(n_b)])
        pl_wrong.append(rev_book - rev)

# Sheet 1 门店基础信息
base_rows = []
for s in STORES:
    sid, city, dist, mall, area, op, mgr, typ, stable, deliv, rr, capex = s
    base_rows.append([sid, city, dist, mall, area, op, mgr, typ, stable, deliv, rr, capex,
                      RENT_PER[sid], PROP_FEE[city], RENT_MONTHS[sid] - (3 - RENT_MONTHS[sid]),
                      '2029-12-31'])
sheet('门店基础信息',
      ['门店编号', '城市', '区域', '商圈', '面积(㎡)', '开业日', '店长', '门店类型',
       '稳定月营收(万元)', '外卖占比', '租金占收入比', '建店投资(万元)',
       '月租金(元/㎡)', '物业费(元/㎡/月)', 'Q1免租月数', '合同到期日'],
      base_rows,
      {1: 9, 2: 8, 3: 10, 4: 18, 5: 9, 6: 12, 7: 8, 8: 9, 9: 15, 10: 9, 11: 12, 12: 14, 13: 12, 14: 14, 15: 12, 16: 12},
      comment=('I1', '本列为 2025 年 12 月稳态月营收。2025-11-01 之后开业的新店在分析期存在爬坡，实际收入低于本列×月份数。'))

# Sheet 2 城市与区域
city_rows = []
for city in ['苏州', '无锡', '宁波', '杭州', '合肥', '南昌']:
    ss = [x for x in STORES if x[1] == city]
    city_rows.append([city, len(ss), sum(1 for x in ss if x[7] == '直营'),
                      sum(1 for x in ss if x[7] == '加盟'),
                      '陈建华' if city in ('苏州', '无锡') else '王海' if city == '宁波' else '林志远',
                      '老市场，稳定' if city in ('苏州', '无锡') else '新市场，爬坡中' if city in ('杭州', '合肥', '南昌') else '稳定'])
sheet('城市与区域', ['城市', '门店数', '直营数', '加盟数', '区域经理', '备注'],
      city_rows, {1: 8, 2: 8, 3: 8, 4: 8, 5: 10, 6: 26})

# Sheet 3 员工花名册
SURNAME = '赵钱孙李周吴郑王冯陈褚卫蒋沈韩杨朱秦尤许何吕施张孔曹严华金魏陶姜'
GIVEN = '伟芳娜敏静丽强磊洋艳杰娟涛明超秀霞平刚桂英华玉兰春梅'
staff_rows = []
emp_no = 1
for s in STORES:
    sid = s[0]
    mgr = s[6]
    staff_rows.append([f'E{emp_no:04d}', sid, mgr, '店长', '2023-06-01' if s[5] < '2024-01-01' else s[5],
                       8500 + (emp_no % 5) * 300, 7200, '在职'])
    emp_no += 1
    for role, cnt in [('烘焙师', 3), ('服务员', 3), ('兼职', 1)]:
        for _ in range(cnt):
            name = random.choice(SURNAME) + random.choice(GIVEN)
            pay = {'烘焙师': 6200, '服务员': 4800, '兼职': 2600}[role] + (emp_no % 7) * 120
            staff_rows.append([f'E{emp_no:04d}', sid, name, role,
                               f'202{random.randint(3,5)}-{random.randint(1,12):02d}-{random.randint(1,28):02d}',
                               pay, 0 if role == '兼职' else round(pay * 0.85, 0), '在职'])
            emp_no += 1
sheet('员工花名册', ['员工编号', '门店编号', '姓名', '岗位', '入职日期', '月薪(元)', '社保基数(元)', '状态'],
      staff_rows, {1: 10, 2: 9, 3: 8, 4: 9, 5: 12, 6: 10, 7: 12, 8: 8})

# Sheet 4 日营业流水明细
sheet('日营业流水明细',
      ['日期', '门店编号', '城市', '堂食应收(元)', '堂食退款(元)', '堂食净额(元)', '外卖GMV(元)', '平台补贴(元)'],
      daily_rows, {1: 12, 2: 9, 3: 8, 4: 14, 5: 13, 6: 14, 7: 13, 8: 13})

# Sheet 5 门店Q1分月P&L
sheet('门店Q1分月P&L',
      ['门店编号', '城市', '月份', '营业收入(万元)', '食材成本(万元)', '人工成本(万元)',
       '租金(万元)', '水电(万元)', '营销(万元)', '外卖佣金(万元)', '管理分摊(万元)',
       '折旧(万元)', 'EBITDA(万元)', '税前利润(万元)', '所得税(万元)', '税后利润(万元)'],
      pl_rows, {1: 9, 2: 8, 3: 9, 4: 14, 5: 14, 6: 14, 7: 11, 8: 11, 9: 11, 10: 13, 11: 13,
                12: 11, 13: 13, 14: 13, 15: 11, 16: 13})

# Sheet 6 外卖平台结算单
deliv_rows = []
for s in STORES:
    sid = s[0]
    for m in (1, 2, 3):
        a = monthly[(sid, m)]
        gmv = a['gmv'] / 10000
        comm = gmv * RATE_V2['comm']
        ship = gmv * 0.05
        settled = gmv - comm - ship
        recorded = settled if sid in REV_ERR else gmv
        deliv_rows.append([sid, f'2026-{m:02d}', '美团' if (m + int(sid[2:])) % 2 else '饿了么',
                           r4(gmv), 18.0, r4(comm), 5.0, r4(ship), r4(settled), r4(recorded)])
sheet('外卖平台结算单',
      ['门店编号', '月份', '平台', '外卖GMV(万元)', '佣金率(%)', '平台佣金(万元)',
       '配送费率(%)', '配送费(万元)', '平台实收(万元)', '门店入账(万元)'],
      deliv_rows, {1: 9, 2: 9, 3: 8, 4: 14, 5: 10, 6: 13, 7: 10, 8: 11, 9: 13, 10: 13})

# Sheet 7 食材采购入库明细
CATS = [('面粉', 4.2), ('黄油', 38.0), ('奶油', 26.0), ('糖', 5.6), ('水果', 12.0), ('包装', 0.8)]
purch_rows = []
doc_no = 1
for s in STORES:
    sid = s[0]
    short = sid in ('SZ05', 'HF02')
    for m in (1, 2, 3):
        a = monthly[(sid, m)]
        rev = (a['dine_net'] + a['gmv']) / 10000
        theory = rev * RATE_V2['food'] * (0.85 if short else 1.0)
        for wk in range(4):
            for cat, price in random.sample(CATS, 2):
                amt = theory * 10000 / 8 * random.uniform(0.85, 1.15)
                qty = round(amt / price, 1)
                d = date(2026, m, 3 + wk * 7)
                purch_rows.append([f'CG{doc_no:05d}', str(d), sid, cat, qty, price, r2(qty * price)])
                doc_no += 1
sheet('食材采购入库明细', ['单据号', '日期', '门店编号', '品类', '数量(kg)', '单价(元/kg)', '金额(元)'],
      purch_rows, {1: 12, 2: 12, 3: 9, 4: 10, 5: 10, 6: 12, 7: 13})

# Sheet 8 费用报销明细
EXP_CATS = ['差旅', '招待', '办公', '维修', '物料']
reimb_rows = []
rn = 1
for s in STORES:
    sid = s[0]
    for m in (1, 2, 3):
        for _ in range(2):
            cat = random.choice(EXP_CATS)
            amt = round(random.uniform(600, 4800), 2)
            inv = f'FP{random.randint(10000000, 99999999)}'
            note = '客户拜访接待' if cat == '招待' else ''
            reimb_rows.append([f'BX{rn:05d}', f'2026-{m:02d}-{random.randint(2,27):02d}', sid,
                               random.choice(SURNAME) + random.choice(GIVEN), cat, amt, inv,
                               s[6], note])
            rn += 1
# 异常1/2：同一发票号重复报销（SZ02、NB01）
dupA = reimb_rows[6][6]
reimb_rows.append(['BX90001', '2026-02-14', 'SZ02', '刘洋', '差旅', 3860.00, dupA, '刘洋', '补充报销'])
reimb_rows.append(['BX90002', '2026-03-09', 'NB01', '吴静', '物料', 2450.00, reimb_rows[36][6], '吴静', '补充报销'])
# 异常3/4：单笔超 5000 仅店长一人审批
reimb_rows.append(['BX90003', '2026-01-22', 'SZ03', '王强', '维修', 12800.00, 'FP55781234', '王强', '设备维修'])
reimb_rows.append(['BX90004', '2026-03-16', 'HF04', '邹敏', '办公', 9600.00, 'FP61039285', '邹敏', '桌椅采购'])
# 异常5：招待费未注明客户与事由
reimb_rows.append(['BX90005', '2026-02-26', 'WX02', '刘洋', '招待', 6200.00, 'FP73209114', '刘洋', ''])
reimb_rows.sort(key=lambda x: (x[1], x[0]))
sheet('费用报销明细', ['报销单号', '日期', '门店编号', '报销人', '类别', '金额(元)', '发票号', '审批人', '备注'],
      reimb_rows, {1: 11, 2: 12, 3: 9, 4: 8, 5: 8, 6: 11, 7: 13, 8: 8, 9: 20})

# Sheet 9 租约明细
lease_rows = []
for s in STORES:
    sid, city, dist, mall, area, op, mgr, typ, stable, deliv, rr, capex = s
    note = '按租约执行'
    if sid == 'NC04':
        note = '补充条款：2026-01-01 至 2026-01-31 免租（开业扶持），2 月起正常计租'
    if sid == 'HF05':
        note = '补充条款：租赁面积含外摆区 15㎡，物业费按 90% 计取'
    lease_rows.append([sid, city, mall, area, RENT_PER[sid], PROP_FEE[city],
                       (RENT_PER[sid] + PROP_FEE[city]) * area, '5', '3', '2029-12-31', note])
sheet('租约明细',
      ['门店编号', '城市', '商圈', '面积(㎡)', '月租金(元/㎡)', '物业费(元/㎡/月)',
       '月租金含物业(元)', '年递增(%)', '押金月数', '合同到期日', '补充条款'],
      lease_rows, {1: 9, 2: 8, 3: 18, 4: 9, 5: 12, 6: 14, 7: 16, 8: 9, 9: 9, 10: 12, 11: 52},
      comment=('K1', '免租期以"补充条款"约定为准。分析期已计提但依约应免的租金须在当期冲回，不追溯调整以前年度。'))

# Sheet 10 月度盘点
INV_ERR = {'SZ04': 0.031, 'WX04': 0.045, 'NB04': 0.038, 'HF05': 0.028}
inv_rows = []
for s in STORES:
    sid = s[0]
    for m in (1, 2, 3):
        a = monthly[(sid, m)]
        rev = (a['dine_net'] + a['gmv']) / 10000
        book = rev * RATE_V2['food'] * 0.30
        rate = INV_ERR.get(sid, 0.012) * random.uniform(0.9, 1.1)
        actual = book * (1 - rate)
        inv_rows.append([sid, f'2026-{m:02d}', r4(book), r4(actual), r4(book - actual), r2(rate * 100)])
sheet('月度盘点', ['门店编号', '月份', '账面库存(万元)', '实盘库存(万元)', '差异金额(万元)', '差异率(%)'],
      inv_rows, {1: 9, 2: 9, 3: 15, 4: 15, 5: 14, 6: 11},
      comment=('F1', '差异率＝差异金额÷账面库存。自 2026 年起，超过 2% 的部分全额由门店承担并从店长季度奖金扣除，不再按阶梯分摊。'))

# Sheet 11-13 制度三版
def rate_sheet(title, rate, note):
    rows = [
        ['食材成本', f"{rate['food']*100:.1f}%", '营业收入', '按标准配方与出成率'],
        ['人工成本', f"{rate['labor']*100:.1f}%", '营业收入', '含店长、烘焙、服务、兼职与社保公积金'],
        ['水电能耗', f"{rate['util']*100:.1f}%", '营业收入', '按营业收入计提'],
        ['营销推广', f"{rate['mkt']*100:.1f}%", '营业收入', '门店自担，含平台满减与地推'],
        ['管理分摊', f"{rate['mgmt']*100:.1f}%", '营业收入', '含总部管理、中央厨房、品控、IT'],
        ['外卖佣金', f"{rate['comm']*100:.1f}%", '外卖GMV', '仅对外卖GMV计提，不得对堂食计提'],
        ['折旧', '5年直线', '建店投资', '残值为0，月折旧=建店投资/60'],
        ['企业所得税', '25%', '正税前利润', '亏损年度所得税为0'],
    ]
    sheet(title, ['项目', '费率', '计提基数', '说明'], rows, {1: 12, 2: 10, 3: 14, 4: 46})
    return rows

rate_sheet('成本与分摊制度v1.0', RATE_V1, '2024')
rate_sheet('成本与分摊制度v2.0', RATE_V2, '2025')
rate_sheet('成本与分摊制度v3.0草案', RATE_V3, '2026')

# Sheet 14 制度生效与废止通知（顶置红字误导）
notice_rows = [
    ['PRC-COST-v1.0', '2023-12-20', '2024-01-01', '2026-01-01 废止', '2024 年度核算依据'],
    ['PRC-COST-v2.0', '2024-12-18', '2025-01-01', '现行有效', '2025 年度及 2026Q1 核算依据'],
    ['PRC-COST-v3.0', '2026-01-15', '2026-04-01', '待生效', '2026-01-15 管委会审议通过，自 2026-04-01 起执行，不追溯调整 2026Q1'],
    ['JY-2026-001', '2026-01-05', '2026-01-05', '已执行', '2026Q1 经营分析会纪要，见"经营分析会会议纪要"'],
    ['WD-2026-002', '2026-01-16', '2026-01-16', '已发布', '文控归档通知，v3.0 自 2026-04-01 起执行'],
]
sheet('制度生效与废止通知', ['文件编号', '发文日期', '生效日期', '当前状态', '说明'],
      notice_rows, {1: 18, 2: 12, 3: 12, 4: 16, 5: 62},
      banner='【文控提示】本汇编已同步至最新版 PRC-COST-v3.0，请各门店及财务按 v3.0 口径编制与复核经营数据。')

# Sheet 15 经营分析会会议纪要
minute_rows = [
    ['2026-01-05', '会议基本信息', '2026Q1 经营分析口径预备会，参会：总经理、财务总监、运营总监、人力总监、区域经理 3 人', '总经办', '已归档'],
    ['2026-01-05', '收入口径', '门店营业收入＝堂食净额＋外卖GMV，外卖GMV为扣除退款后的净额并含平台补贴；不得以平台实收列报营业收入', '财务部', '2026-01-15'],
    ['2026-01-05', '成本费率', 'Q1 成本费率按现行有效版本执行，未生效版本不得提前适用，亦不追溯调整已结月份', '财务部', '2026-01-15'],
    ['2026-01-05', '管理分摊', '自 2026Q1 起，管理分摊计提基数由"营业收入"改为"营业收入－外卖平台佣金"，费率仍按当版制度执行', '财务部', '2026-01-15'],
    ['2026-01-05', '加盟门店', '加盟门店纳入本次经营分析与分级处置范围，收入成本按门店口径还原；但公司合并报表仅确认授权费与供应链毛利，加盟利润不并入直营利润合计', '财务部', '2026-01-15'],
    ['2026-01-05', '免租期', '免租期以租约补充条款为准，分析期已计提但依约应免的租金在当期冲回，不追溯调整以前年度', '财务部', '2026-01-15'],
    ['2026-01-05', '盘点责任', '差异率超过 2% 的部分全额由门店承担，从店长季度奖金扣除', '运营部', '2026-01-31'],
    ['2026-01-05', '店长奖金', '2026Q1 起按新版店长绩效考核办法（v2）核算，权重 EBITDA 45%、营收 35%、损耗 20%', '人力部', '2026-01-31'],
    ['2026-01-05', '关店决策', 'D 级门店先整改一个季度，连续两个季度 D 级方可关店或转址，不得当季即关', '运营部', '2026-03-31'],
]
sheet('经营分析会会议纪要', ['日期', '议题', '决议内容', '责任人', '完成时限'],
      minute_rows, {1: 12, 2: 14, 3: 86, 4: 10, 5: 12})

# Sheet 16-17 考核办法 v1/v2
sheet('店长绩效考核办法v1',
      ['考核项', '权重', '计算口径', '上限/规则'],
      [['EBITDA达成率', '50%', '实际EBITDA÷预算EBITDA', '上限1.5'],
       ['营收达成率', '30%', '实际营收÷预算营收', '上限1.5'],
       ['损耗达标得分', '20%', '1－max(0,差异率－2%)×10', '下限0'],
       ['基础奖金', '—', 'A级1.2万、B级1.0万、C级0.8万、D级0.5万', '季度基础奖金'],
       ['D级负激励', '—', '基础奖金扣减50%', '与追回预提并用'],
       ['发放规则', '—', '季度终了次月发放80%，20%留存至年度', '离职不发留存']],
      {1: 16, 2: 8, 3: 44, 4: 26})

sheet('店长绩效考核办法v2',
      ['考核项', '权重', '计算口径', '上限/规则'],
      [['EBITDA达成率', '45%', '实际EBITDA÷预算EBITDA', '上限1.5'],
       ['营收达成率', '35%', '实际营收÷预算营收', '上限1.5'],
       ['损耗达标得分', '20%', '1－max(0,差异率－2%)×10', '下限0'],
       ['基础奖金', '—', 'A级1.2万、B级1.0万、C级0.8万、D级0.5万', '季度基础奖金'],
       ['D级负激励', '—', '基础奖金扣减50%', '与追回预提并用'],
       ['发放规则', '—', '季度终了次月发放80%，20%留存至年度', '离职不发留存'],
       ['执行起始', '—', '2026Q1 起执行，替代 v1', '经营分析会 JY-2026-001 决议']],
      {1: 16, 2: 8, 3: 44, 4: 26})

# Sheet 18 门店分级标准
sheet('门店分级标准', ['分级', 'EBITDA率', '处置策略', '基础奖金(万元)'],
      [['A', '≥12%', '保持，可作为标杆复制', 1.2],
       ['B', '8%-12%', '维持经营，关注成本', 1.0],
       ['C', '4%-8%', '限期整改，Q2 复核', 0.8],
       ['D', '<4% 或连续两季亏损', '启动整改，连续两季 D 级关店或转址', 0.5]],
      {1: 8, 2: 20, 3: 40, 4: 16})

# Sheet 19 总部分摊规则
sheet('总部分摊规则', ['分摊项', '计提基数', '费率来源', '说明'],
      [['管理与中央厨房分摊', '营业收入－外卖平台佣金', '按当版成本与分摊制度', '2026Q1 起调整，详见经营分析会纪要 JY-2026-001'],
       ['营销推广费', '营业收入', '按当版成本与分摊制度', '门店自担'],
       ['外卖平台佣金', '外卖GMV', '按当版成本与分摊制度', '仅对外卖计提'],
       ['食材成本', '营业收入', '按当版成本与分摊制度', '按标准配方测算'],
       ['人工成本', '营业收入', '按当版成本与分摊制度', '含社保公积金'],
       ['水电能耗', '营业收入', '按当版成本与分摊制度', '按营业收入计提'],
       ['折旧', '建店投资', '5 年直线', '月折旧＝建店投资÷60'],
       ['企业所得税', '正税前利润', '25%', '亏损年度为 0']],
      {1: 20, 2: 24, 3: 20, 4: 52})

# Sheet 20 财务确认规则（2025 版，含加盟旧口径）
fin_rows = [
    ['收入确认', '营业收入＝堂食净额＋外卖GMV；外卖GMV为扣除退款后的净额并含平台补贴', '用于 P&L 与奖金'],
    ['外卖佣金', '仅对外卖GMV按当版制度费率计提，不得对堂食收入计提', '计入营业费用'],
    ['免租期', '免租期按租约约定执行，直线摊销；租约补充条款另有约定的从其约定', '不计提租金'],
    ['折旧', '建店投资按 5 年直线折旧，残值为 0，月折旧＝建店投资÷60', '不含押金与流动资金'],
    ['盘点差异', '差异率≤2% 为正常损耗；超过部分由门店承担并从店长奖金扣除', '差异金额计入食材成本'],
    ['总部分摊', '管理费、营销费按营业收入计提，不得按门店数平均分摊', '固定费率'],
    ['所得税', '按正税前利润的 25% 计提，亏损年度所得税为 0', '不确认递延所得税资产'],
    ['门店分级', 'A 级 EBITDA 率≥12%，B 级 8%-12%，C 级 4%-8%，D 级<4% 或连续亏损', '用于奖金与处置'],
    ['处置触发', 'D 级门店启动整改，连续 2 季 D 级启动关店或转址', 'C 级门店限期整改'],
    ['奖金口径', '基础奖金×(EBITDA达成×0.5＋营收达成×0.3＋损耗达标×0.2)，D 级扣减 50%', '本表为 2025 版口径'],
    ['预算基准', 'Q1 预算营收＝稳定月营收×3；预算 EBITDA＝预算营收×(1－食材率－人工率－水电率－营销率－管理分摊率)－预算租金', '用于达成率'],
    ['加盟门店', '加盟门店仅确认授权费与供应链毛利，不并入直营 P&L', '本表为 2025 版口径'],
    ['报销审批', '单笔报销≥5000 元须店长与区域经理双签；招待费须注明客户名称与事由', '财务复核'],
    ['报销票据', '同一发票号不得重复报销；差旅住宿标准一线城市≤600 元/晚', '系统校验'],
]
sheet('财务确认规则', ['规则', '内容', '用途'], fin_rows, {1: 14, 2: 82, 3: 26},
      comment=('A1', '本表为 2025 版财务口径，2026 年调整事项（管理分摊基数、加盟门店口径、奖金权重、免租冲回）以《经营分析会会议纪要》JY-2026-001 为准。'))

# Sheet 21 管理层要求
MGR = [
    ['M01', '统一按 v3.0 最新费率重算，便于与 2026 年度预算衔接', 'v3.0 自 2026-04-01 起执行，Q1 应适用 v2.0'],
    ['M02', '外卖佣金按全部营业收入 18% 计提，简化核算', '仅对外卖GMV计提'],
    ['M03', '平台实收即门店外卖收入，差额计入营销费用', '营业收入按外卖GMV确认，不得以实收列报'],
    ['M04', '加盟门店全部剔除，不参与分级与处置', '纪要明确加盟门店纳入分级与处置'],
    ['M05', '加盟门店利润全额并入公司利润合计', '公司仅确认授权费与供应链毛利'],
    ['M06', '盘点差异统一按 1% 计入成本，不做门店责任划分', '超 2% 部分全额由门店承担'],
    ['M07', 'D 级门店本季度立即关店止损', '先整改一个季度，连续两季 D 级方可关店'],
    ['M08', '免租期不冲回，按全年 12 个月平均分摊', 'Q1 应免租部分须当期冲回'],
    ['M09', '管理分摊按门店数平均分摊', '按（营业收入－外卖佣金）×费率计提'],
    ['M10', '店长奖金继续按 v1 权重（50/30/20）核算', '2026Q1 起按 v2（45/35/20）'],
    ['M11', '亏损门店虽不缴所得税，但可确认递延所得税资产', '亏损年度所得税为 0，不确认递延所得税资产'],
    ['M12', 'P&L 账面数即为经营分析结果，不必重算', '6 家门店 P&L 存在口径错误，须重算'],
]
sheet('管理层要求', ['编号', '要求', '问题'], MGR, {1: 8, 2: 56, 3: 40})

wb.save(OUT)
print('saved', OUT)
print('stores', len(STORES), 'daily', len(daily_rows), 'pl', len(pl_rows),
      'purch', len(purch_rows), 'reimb', len(reimb_rows), 'inv', len(inv_rows),
      'staff', len(staff_rows), 'deliv', len(deliv_rows))
print('total data rows', len(STORES) + len(city_rows) + len(staff_rows) + len(daily_rows) + len(pl_rows)
      + len(deliv_rows) + len(purch_rows) + len(reimb_rows) + len(lease_rows) + len(inv_rows)
      + 8 + 8 + 8 + len(notice_rows) + len(minute_rows) + 6 + 7 + 4 + 8 + len(fin_rows) + len(MGR))
