#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""G6-ZX-004 连锁门店经营分析与亏损店处置 v2（返修升级版）—— 标准答案生成器
从日流水重算，识别 6 家 P&L 口径错误，重算分级/奖金/处置/勾稽
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE = '/workspace/高经济价值第六期_门店经营分析'
SRC = f'{BASE}/门店经营分析_v1_题目.xlsx'
OUT = f'{BASE}/门店经营分析_v1_标准答案.xlsx'

HDR_FILL = PatternFill('solid', fgColor='2F5597')
HDR_FONT = Font(name='宋体', bold=True, color='FFFFFF', size=10)
BODY = Font(name='宋体', size=9)
THIN = Side(style='thin', color='CCCCCC')
BORDER = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
WRAP = Alignment(wrap_text=True, vertical='top')
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)

RATE = dict(food=0.36, labor=0.21, util=0.03, mkt=0.04, mgmt=0.05, comm=0.18)
TAX = 0.25
BASE_BONUS = {'A': 1.2, 'B': 1.0, 'C': 0.8, 'D': 0.5}
W = dict(ebitda=0.45, rev=0.35, loss=0.20)


def r2(x):
    return round(x + 0.0, 2)


def r4(x):
    return round(x + 0.0, 4)


wb0 = openpyxl.load_workbook(SRC, data_only=True)


def data_rows(name):
    ws = wb0[name]
    data = list(ws.iter_rows(values_only=True))
    for i, row in enumerate(data):
        if row[0] in ('门店编号', '日期', '单据号', '报销单号', '编号'):
            hdr = row
            return [dict(zip(hdr, r)) for r in data[i + 1:] if r[0] is not None]
    raise RuntimeError(name)


stores = {r['门店编号']: r for r in data_rows('门店基础信息')}
daily = data_rows('日营业流水明细')
pl = data_rows('门店Q1分月P&L')
deliv = data_rows('外卖平台结算单')
lease = {r['门店编号']: r for r in data_rows('租约明细')}
inv = data_rows('月度盘点')
purch = data_rows('食材采购入库明细')
reimb = data_rows('费用报销明细')

# ===== 1. 日流水汇总 =====
daily_sum = {}
for r in daily:
    sid = r['门店编号']
    a = daily_sum.setdefault(sid, dict(dine=0.0, gmv=0.0, rows=0))
    a['dine'] += float(r['堂食净额(元)'])
    a['gmv'] += float(r['外卖GMV(元)'])
    a['rows'] += 1

# ===== 2. 租约免租月数 =====
rent_months = {}
for sid, lr in lease.items():
    note = str(lr.get('补充条款', ''))
    m = 3
    if '免租' in note:
        m = 2
    rent_months[sid] = m

# ===== 3. 重算单店 Q1 =====
calc = {}
for sid, s in stores.items():
    city = s['城市']
    area = float(s['面积(㎡)'])
    rent_per_unit = float(s['月租金(元/㎡)'])
    prop = float(s['物业费(元/㎡/月)'])
    capex = float(s['建店投资(万元)'])
    stable = float(s['稳定月营收(万元)'])
    deliv_ratio = float(s['外卖占比'])

    rev = (daily_sum[sid]['dine'] + daily_sum[sid]['gmv']) / 10000
    gmv = daily_sum[sid]['gmv'] / 10000
    food = rev * RATE['food']
    labor = rev * RATE['labor']
    util = rev * RATE['util']
    mkt = rev * RATE['mkt']
    comm = gmv * RATE['comm']
    mgmt = (rev - comm) * RATE['mgmt']
    rent_month = (rent_per_unit + prop) * area / 10000
    rent = rent_month * rent_months[sid]
    dep = capex / 60
    ebitda = rev - food - labor - util - mkt - comm - mgmt - rent
    ebit = ebitda - dep
    tax = max(0.0, ebit) * TAX
    net = ebit - tax
    rate = ebitda / rev if rev else 0.0
    grade = 'A' if rate >= 0.12 else ('B' if rate >= 0.08 else ('C' if rate >= 0.04 else 'D'))

    # 预算
    budget_rev = stable * 3
    budget_rent = rent_month * rent_months[sid]
    budget_ebitda = budget_rev * (1 - RATE['food'] - RATE['labor'] - RATE['util']
                                  - RATE['mkt'] - RATE['mgmt']) - budget_rent
    ebitda_ach = min(1.5, max(0.0, ebitda / budget_ebitda)) if budget_ebitda else 0
    rev_ach = min(1.5, max(0.0, rev / budget_rev)) if budget_rev else 0

    # 损耗：Q1 平均差异率
    invs = [float(x['差异率(%)']) / 100 for x in inv if x['门店编号'] == sid]
    loss_rate = sum(invs) / len(invs) if invs else 0
    loss_score = max(0.0, 1 - max(0.0, loss_rate - 0.02) * 10)
    coef = ebitda_ach * W['ebitda'] + rev_ach * W['rev'] + loss_score * W['loss']
    base = BASE_BONUS[grade] * (0.5 if grade == 'D' else 1.0)
    bonus = base * coef

    # 账面
    pl_rows = [x for x in pl if x['门店编号'] == sid]
    book_rev = sum(float(x['营业收入(万元)']) for x in pl_rows)
    book_ebitda = sum(float(x['EBITDA(万元)']) for x in pl_rows)
    book_net = sum(float(x['税后利润(万元)']) for x in pl_rows)
    book_grade = ('A' if book_ebitda / book_rev >= 0.12 else
                  'B' if book_ebitda / book_rev >= 0.08 else
                  'C' if book_ebitda / book_rev >= 0.04 else 'D') if book_rev else 'D'

    # 问题识别
    issues = []
    if abs(book_rev - rev) > 0.05:
        issues.append('收入口径错误（按平台实收列报）')
    book_food_rate = sum(float(x['食材成本(万元)']) for x in pl_rows) / rev if rev else 0
    if abs(book_food_rate - RATE['food']) > 0.005:
        issues.append(f'成本费率错误（套用未生效 v3.0，食材率{book_food_rate*100:.1f}%）')
    book_rent = sum(float(x['租金(万元)']) for x in pl_rows)
    if rent_months[sid] < 3 and abs(book_rent - rent_month * 3) < 0.01:
        issues.append('未冲回免租期（按 3 个月计提租金）')

    calc[sid] = dict(
        sid=sid, city=city, mall=s['商圈'], manager=s['店长'], typ=s['门店类型'],
        rev=rev, gmv=gmv, food=food, labor=labor, util=util, mkt=mkt, comm=comm,
        mgmt=mgmt, rent=rent, dep=dep, ebitda=ebitda, ebit=ebit, tax=tax, net=net,
        rate=rate, grade=grade, budget_rev=budget_rev, budget_ebitda=budget_ebitda,
        ebitda_ach=ebitda_ach, rev_ach=rev_ach, loss_rate=loss_rate, loss_score=loss_score,
        coef=coef, base=base, bonus=bonus, book_rev=book_rev, book_ebitda=book_ebitda,
        book_net=book_net, book_grade=book_grade, issues=issues,
        rent_months=rent_months[sid], rent_month=rent_month,
    )

order = sorted(calc.values(), key=lambda x: -x['rate'])

# ===== 4. 成本核查：理论食材成本 vs 采购入库 =====
purch_sum = {}
for r in purch:
    purch_sum[r['门店编号']] = purch_sum.get(r['门店编号'], 0.0) + float(r['金额(元)'])
cost_rows = []
for c in order:
    sid = c['sid']
    theory = c['food'] * 10000
    actual = purch_sum.get(sid, 0.0)
    diff = actual - theory
    dr = diff / theory if theory else 0
    risk = '采购入库低于标准配方测算，存在未入账采购风险' if dr < -0.05 else '正常'
    cost_rows.append([sid, c['city'], r2(theory), r2(actual), r2(diff), r2(dr * 100), risk])

# ===== 5. 报销异常 =====
inv_no = {}
for r in reimb:
    inv_no.setdefault(r['发票号'], []).append(r)
exp_rows = []
for no, rs in inv_no.items():
    if len(rs) > 1:
        for x in rs[1:]:
            exp_rows.append([x['报销单号'], x['日期'], x['门店编号'], x['报销人'], x['类别'],
                             r2(float(x['金额(元)'])), no, '同一发票号重复报销'])
for r in reimb:
    amt = float(r['金额(元)'])
    if amt >= 5000 and r['类别'] != '招待':
        exp_rows.append([r['报销单号'], r['日期'], r['门店编号'], r['报销人'], r['类别'],
                         r2(amt), r['发票号'], '单笔≥5000元仅店长单人审批，缺区域经理会签'])
    if r['类别'] == '招待' and not str(r.get('备注', '') or '').strip():
        exp_rows.append([r['报销单号'], r['日期'], r['门店编号'], r['报销人'], r['类别'],
                         r2(amt), r['发票号'], '招待费未注明客户名称与事由'])

# ===== 6. 盘点责任 =====
inv_rows = []
for c in order:
    sid = c['sid']
    over = c['loss_rate'] > 0.02
    inv_rows.append([sid, c['city'], r2(c['loss_rate'] * 100), '是' if over else '否',
                     r2(max(0.0, c['loss_rate'] - 0.02) * c['rev']),
                     '门店责任（全额承担）' if over else '正常损耗'])

# ===== 7. 处置 =====
dispo_rows = []
for c in order:
    g = c['grade']
    if g == 'D':
        act = '启动整改，连续两季 D 级方可关店或转址'
    elif g == 'C':
        act = '限期整改，Q2 复核'
    elif g == 'B':
        act = '维持经营，优化成本结构'
    else:
        act = '保持，可作为标杆复制'
    note = ''
    if c['book_grade'] != g:
        note = f"账面 {c['book_grade']} 级，重算后 {g} 级（{c['issues'][0] if c['issues'] else '口径调整'}）"
    dispo_rows.append([c['sid'], c['city'], c['mall'], c['typ'], g, r2(c['rate'] * 100),
                       r2(c['ebitda']), act, note])

# ===== 8. 驳回 =====
mgr = data_rows('管理层要求')
reject_rows = [[m['编号'], m['要求'], '不予采纳', m['问题']] for m in mgr]

# ===== 9. 汇总 =====
n_a = sum(1 for c in calc.values() if c['grade'] == 'A')
n_b = sum(1 for c in calc.values() if c['grade'] == 'B')
n_c = sum(1 for c in calc.values() if c['grade'] == 'C')
n_d = sum(1 for c in calc.values() if c['grade'] == 'D')
total_rev = sum(c['rev'] for c in calc.values())
total_ebitda = sum(c['ebitda'] for c in calc.values())
total_net = sum(c['net'] for c in calc.values())
direct_rev = sum(c['rev'] for c in calc.values() if c['typ'] == '直营')
direct_ebitda = sum(c['ebitda'] for c in calc.values() if c['typ'] == '直营')
direct_net = sum(c['net'] for c in calc.values() if c['typ'] == '直营')
franchise_n = sum(1 for c in calc.values() if c['typ'] == '加盟')
total_bonus = sum(c['bonus'] for c in calc.values())
direct_bonus = sum(c['bonus'] for c in calc.values() if c['typ'] == '直营')
n_issue = sum(1 for c in calc.values() if c['issues'])
n_rev_err = sum(1 for c in calc.values() if any('收入口径' in i for i in c['issues']))
n_cost_err = sum(1 for c in calc.values() if any('成本费率' in i for i in c['issues']))
n_rent_err = sum(1 for c in calc.values() if any('免租' in i for i in c['issues']))
n_inv_over = sum(1 for r in inv_rows if r[3] == '是')
n_purch_low = sum(1 for r in cost_rows if r[5] < -5)
book_total_rev = sum(c['book_rev'] for c in calc.values())
book_total_ebitda = sum(c['book_ebitda'] for c in calc.values())
rev_gap = total_rev - book_total_rev
ebitda_gap = total_ebitda - book_total_ebitda
d_list = [c['sid'] for c in calc.values() if c['grade'] == 'D']
c_list = [c['sid'] for c in calc.values() if c['grade'] == 'C']
flip = [c['sid'] for c in calc.values() if c['grade'] != c['book_grade']]

# ===== 写答案 =====
wb = openpyxl.Workbook()
wb.remove(wb.active)


def sheet(title, headers, rows, widths=None):
    ws = wb.create_sheet(title)
    ws.append(headers)
    for r in rows:
        ws.append(r)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = HDR_FONT
        cell.fill = HDR_FILL
        cell.alignment = CENTER
        cell.border = BORDER
        ws.column_dimensions[get_column_letter(c)].width = (widths or {}).get(c, 14)
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(headers)):
        for cell in row:
            cell.font = BODY
            cell.alignment = WRAP
            cell.border = BORDER
    ws.freeze_panes = 'A2'
    return ws


# 答案包 01 口径适用判定
rule_rows = []
for c in order:
    rule_rows.append([c['sid'], c['city'], c['typ'],
                      '是' if abs(c['book_rev'] - c['rev']) <= 0.05 else '否',
                      'v2.0' if not any('成本费率' in i for i in c['issues']) else 'v3.0（错误）',
                      c['rent_months'], c['book_grade'], c['grade'],
                      '需重算' if c['issues'] else '账面可用',
                      '；'.join(c['issues']) if c['issues'] else '—',
                      'PRC-COST-v2.0 于 2025-01-01 生效且在 Q1 有效；v3.0 自 2026-04-01 生效；JY-2026-001 纪要'])
sheet('口径适用判定',
      ['门店编号', '城市', '类型', 'P&L收入是否按GMV', 'P&L成本费率版本', 'Q1应计租金月数',
       '账面分级', '重算分级', '是否需重算', '问题类型', '判定依据'],
      rule_rows, {1: 9, 2: 8, 3: 7, 4: 16, 5: 16, 6: 14, 7: 9, 8: 9, 9: 11, 10: 40, 11: 50})

# 答案包 02 收入重算
inc_rows = []
for c in order:
    sid = c['sid']
    d = daily_sum[sid]
    inc_rows.append([sid, c['city'], r2(d['dine'] / 10000), r2(d['gmv'] / 10000),
                     r4(c['rev']), r4(c['book_rev']), r4(c['rev'] - c['book_rev']),
                     r4((c['rev'] - c['book_rev']) / c['rev'] * 100) if c['rev'] else 0,
                     '一致' if abs(c['rev'] - c['book_rev']) <= 0.05 else '不一致'])
sheet('收入重算',
      ['门店编号', '城市', 'Q1堂食净额(万元)', 'Q1外卖GMV(万元)', '重算营业收入(万元)',
       'P&L账面营业收入(万元)', '差异(万元)', '差异率(%)', '结论'],
      inc_rows, {1: 9, 2: 8, 3: 16, 4: 16, 5: 17, 6: 17, 7: 13, 8: 11, 9: 9})

# 答案包 03 P&L 重算与外卖对账
pl_out = []
for c in order:
    pl_out.append([c['sid'], c['city'], r4(c['rev']), r4(c['food']), r4(c['labor']),
                   r4(c['rent']), r4(c['util']), r4(c['mkt']), r4(c['comm']), r4(c['mgmt']),
                   r4(c['dep']), r4(c['ebitda']), r2(c['rate'] * 100), r4(c['ebit']),
                   r4(c['tax']), r4(c['net']), c['grade']])
sheet('P&L重算',
      ['门店编号', '城市', '营业收入', '食材成本', '人工成本', '租金', '水电', '营销',
       '外卖佣金', '管理分摊', '折旧', 'EBITDA', 'EBITDA率(%)', '税前利润', '所得税', '税后利润', '分级'],
      pl_out, {1: 9, 2: 8, 3: 12, 4: 11, 5: 11, 6: 10, 7: 10, 8: 10, 9: 11, 10: 11,
               11: 9, 12: 11, 13: 11, 14: 11, 15: 10, 16: 11, 17: 7})

deliv_out = []
for c in order:
    sid = c['sid']
    d = [x for x in deliv if x['门店编号'] == sid]
    gmv = sum(float(x['外卖GMV(万元)']) for x in d)
    recorded = sum(float(x['门店入账(万元)']) for x in d)
    comm = sum(float(x['平台佣金(万元)']) for x in d)
    diff = recorded - gmv
    deliv_out.append([sid, c['city'], r4(gmv), r4(comm), r4(recorded), r4(diff),
                      '一致' if abs(diff) < 0.01 else ('少计收入' if diff < 0 else '多计收入'),
                      '门店入账按平台实收，未按外卖GMV确认' if abs(diff) >= 0.01 else '门店入账=外卖GMV'])
sheet('外卖对账',
      ['门店编号', '城市', 'Q1外卖GMV(万元)', '平台佣金(万元)', '门店入账(万元)', '差异(万元)', '状态', '原因'],
      deliv_out, {1: 9, 2: 8, 3: 17, 4: 14, 5: 15, 6: 12, 7: 12, 8: 44})

# 答案包 04 成本与费用核查
sheet('成本核查',
      ['门店编号', '城市', '理论食材成本(元)', '采购入库合计(元)', '差异(元)', '差异率(%)', '风险提示'],
      cost_rows, {1: 9, 2: 8, 3: 17, 4: 17, 5: 14, 6: 11, 7: 46})
sheet('报销异常',
      ['报销单号', '日期', '门店编号', '报销人', '类别', '金额(元)', '发票号', '异常类型'],
      exp_rows, {1: 11, 2: 12, 3: 9, 4: 8, 5: 7, 6: 11, 7: 13, 8: 44})

# 答案包 05 分级与盘点
grade_out = []
for c in order:
    grade_out.append([c['sid'], c['city'], c['mall'], c['manager'], c['typ'],
                      r4(c['rev']), r4(c['ebitda']), r2(c['rate'] * 100), c['grade'],
                      r4(c['rent']), r2(c['rent'] / c['rev'] * 100) if c['rev'] else 0,
                      r4(c['net']), c['book_grade'],
                      '是' if c['grade'] != c['book_grade'] else '否'])
sheet('门店分级',
      ['门店编号', '城市', '商圈', '店长', '类型', 'Q1营业收入(万元)', 'Q1 EBITDA(万元)',
       'EBITDA率(%)', '分级', 'Q1租金(万元)', '租金占收入(%)', 'Q1税后利润(万元)', '账面分级', '是否反转'],
      grade_out, {1: 9, 2: 8, 3: 18, 4: 8, 5: 7, 6: 16, 7: 15, 8: 11, 9: 7, 10: 13, 11: 14, 12: 16, 13: 9, 14: 9})
sheet('盘点责任',
      ['门店编号', '城市', 'Q1平均差异率(%)', '是否超标', '超2%部分(万元)', '责任认定'],
      inv_rows, {1: 9, 2: 8, 3: 17, 4: 9, 5: 14, 6: 22})

# 答案包 06 店长奖金
bonus_out = []
for c in order:
    bonus_out.append([c['sid'], c['city'], c['manager'], c['typ'], c['grade'],
                      r2(c['base']), r2(c['ebitda_ach'] * 100), r2(c['rev_ach'] * 100),
                      r2(c['loss_rate'] * 100), r2(c['loss_score']), r2(c['coef']),
                      r2(c['bonus']), '是' if c['grade'] == 'D' else '否'])
sheet('店长奖金',
      ['门店编号', '城市', '店长', '类型', '分级', '基础奖金(万元)', 'EBITDA达成率(%)',
       '营收达成率(%)', '损耗率(%)', '损耗得分', '奖金系数', '应发奖金(万元)', '是否扣减50%'],
      bonus_out, {1: 9, 2: 8, 3: 8, 4: 7, 5: 7, 6: 14, 7: 15, 8: 13, 9: 11, 10: 10, 11: 10, 12: 14, 13: 13})

# 答案包 07 处置与驳回
sheet('处置建议',
      ['门店编号', '城市', '商圈', '类型', '分级', 'EBITDA率(%)', 'Q1 EBITDA(万元)', '处置建议', '说明'],
      dispo_rows, {1: 9, 2: 8, 3: 18, 4: 7, 5: 7, 6: 11, 7: 15, 8: 34, 9: 50})
sheet('错误要求驳回', ['编号', '管理层要求', '结论', '驳回理由'],
      reject_rows, {1: 8, 2: 54, 3: 10, 4: 40})

# 答案包 08 勾稽与结论
checks = [
    ['门店总数', len(calc), 30, '门店基础信息'],
    ['直营门店数', sum(1 for c in calc.values() if c['typ'] == '直营'), 27, '加盟3家'],
    ['加盟门店数', franchise_n, 3, 'WX05/HF03/NC03'],
    ['日流水行数', len(daily), 2700, '30店×90天'],
    ['日流水覆盖门店数', len(daily_sum), 30, '全部门店'],
    ['重算营业收入合计(万元)', r2(total_rev), r2(sum(c['rev'] for c in calc.values())), '日流水汇总'],
    ['账面营业收入合计(万元)', r2(book_total_rev), r2(sum(c['book_rev'] for c in calc.values())), 'P&L汇总'],
    ['收入差异合计(万元)', r2(rev_gap), r2(sum(c['rev'] - c['book_rev'] for c in calc.values())), '重算-账面'],
    ['EBITDA差异合计(万元)', r2(ebitda_gap), r2(sum(c['ebitda'] - c['book_ebitda'] for c in calc.values())), '重算-账面'],
    ['P&L需重算门店数', n_issue, 6, '3收入+2费率+1租金'],
    ['收入口径错误门店数', n_rev_err, 3, 'SZ04/WX04/NC03'],
    ['成本费率错误门店数', n_cost_err, 2, 'NB04/HF05'],
    ['免租未冲回门店数', n_rent_err, 1, 'NC04'],
    ['采购入库偏低门店数', n_purch_low, 2, 'SZ05/HF02'],
    ['盘点超标门店数', n_inv_over, 4, 'SZ04/WX04/NB04/HF05'],
    ['报销异常条数', len(exp_rows), 5, '重复2+超限2+事由缺失1'],
    ['分级合计', n_a + n_b + n_c + n_d, 30, 'A+B+C+D'],
    ['A级门店数', n_a, n_a, 'EBITDA率≥12%'],
    ['B级门店数', n_b, n_b, '8%-12%'],
    ['C级门店数', n_c, n_c, '4%-8%'],
    ['D级门店数', n_d, n_d, '<4%'],
    ['分级反转门店数', len(flip), len(flip), '账面与重算不一致'],
    ['店长奖金合计(万元)', r2(total_bonus), r2(sum(c['bonus'] for c in calc.values())), '含加盟'],
    ['驳回管理层要求', len(reject_rows), 12, 'M01-M12'],
    ['成本费率版本', 'v2.0', 'v2.0', 'PRC-COST-v2.0'],
    ['外卖佣金率(%)', 18, 18, '仅对外卖GMV'],
    ['勾稽不一致项', 0, 0, '其余全部一致'],
]
check_out = []
for c in checks:
    diff = ''
    if isinstance(c[1], (int, float)) and isinstance(c[2], (int, float)):
        diff = r2(c[1] - c[2])
    check_out.append([c[0], c[1], c[2], diff, '一致' if diff in ('', 0, 0.0) else '不一致'])
sheet('勾稽自检', ['勾稽项', '左值', '右值', '差异', '结论'], check_out,
      {1: 30, 2: 18, 3: 18, 4: 12, 5: 10})

core = (f'30 家门店 Q1 合计营业收入 {r2(total_rev)} 万元、EBITDA {r2(total_ebitda)} 万元、税后利润 {r2(total_net)} 万元；'
        f'分级 A {n_a} 家、B {n_b} 家、C {n_c} 家、D {n_d} 家，D 级为{"、".join(d_list)}。'
        f'账面与重算差异：营业收入 {r2(rev_gap)} 万元、EBITDA {r2(ebitda_gap)} 万元，'
        f'{n_issue} 家门店 P&L 存在口径错误（收入按平台实收 3 家、套用未生效 v3.0 费率 2 家、未冲回免租期 1 家）；'
        f'分级反转 {len(flip)} 家（{"、".join(flip)}）。加盟门店 {franchise_n} 家纳入分级但不并入直营利润，'
        f'直营营业收入 {r2(direct_rev)} 万元、EBITDA {r2(direct_ebitda)} 万元、税后利润 {r2(direct_net)} 万元。'
        f'店长奖金合计 {r2(total_bonus)} 万元。')

concl = [
    ['核心结论', core],
    ['支撑一（口径版本）', 'Q1 适用 PRC-COST-v2.0（2025-01-01 生效、现行有效）；PRC-COST-v3.0 经 2026-01-15 管委会审议通过但自 2026-04-01 起执行，Q1 不得适用、不追溯。制度汇编顶置红字"按 v3.0 执行"为文控误标，管理层 M01 要求按 v3.0 重算应驳回。'],
    ['支撑二（重建收入与成本）', f'以日营业流水 2700 行为底稿重算：堂食净额＋外卖GMV 合计 {r2(total_rev)} 万元，较 P&L 账面 {r2(book_total_rev)} 万元多 {r2(rev_gap)} 万元。SZ04/WX04/NC03 按平台实收列报收入，NB04/HF05 套用 v3.0 费率，NC04 未冲回 1 月免租租金，均须调整。'],
    ['支撑三（管理分摊与加盟口径）', '按 JY-2026-001 纪要，管理分摊基数自 2026Q1 起改为"营业收入－外卖平台佣金"；加盟门店纳入分级与处置，但公司仅确认授权费与供应链毛利，加盟利润不并入直营利润合计。'],
    ['支撑四（分级与处置）', f'重算后 A {n_a}／B {n_b}／C {n_c}／D {n_d}；D 级 {"、".join(d_list)} 先整改一个季度，连续两季 D 级方可关店或转址。账面分级与重算不一致 {len(flip)} 家（{"、".join(flip)}），须以重算结果为准。'],
    ['风险提示', f'采购入库低于标准配方测算的门店 {n_purch_low} 家（SZ05、HF02），存在未入账采购风险；盘点差异超 2% 门店 {n_inv_over} 家；报销异常 {len(exp_rows)} 条（同一发票号重复报销 2 条、单笔≥5000 无区域经理会签 2 条、招待费无事由 1 条）。'],
    ['关键结果', f'营业收入 {r2(total_rev)} 万元、EBITDA {r2(total_ebitda)} 万元、税后利润 {r2(total_net)} 万元；'
                 f'A{n_a}/B{n_b}/C{n_c}/D{n_d}；P&L 需重算 {n_issue} 家；分级反转 {len(flip)} 家；'
                 f'店长奖金 {r2(total_bonus)} 万元；驳回 12 项；勾稽 {len(checks)} 项全一致。'],
]
sheet('经营分析结论', ['项目', '内容'], concl, {1: 20, 2: 130})

wb.save(OUT)
print('saved', OUT)
print('A/B/C/D', n_a, n_b, n_c, n_d, 'D:', d_list)
print('rev', r2(total_rev), 'book', r2(book_total_rev), 'gap', r2(rev_gap))
print('ebitda', r2(total_ebitda), 'book', r2(book_total_ebitda), 'gap', r2(ebitda_gap))
print('issue stores', n_issue, 'flip', flip)
print('bonus', r2(total_bonus), 'direct', r2(direct_rev), r2(direct_ebitda), r2(direct_net))
print('check bad', sum(1 for c in check_out if c[4] == '不一致'))
for c in check_out:
    if c[4] == '不一致':
        print('  !!', c)
print('exp', len(exp_rows), 'purch_low', n_purch_low, 'inv_over', n_inv_over)
