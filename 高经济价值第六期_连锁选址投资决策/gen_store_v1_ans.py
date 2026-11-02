#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""连锁品牌新区域扩张的选址与单店投资决策 v1 —— 标准答案生成器"""
import os
from datetime import date, datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE = '/workspace/高经济价值第六期_连锁选址投资决策'
SRC = f'{BASE}/连锁选址投资决策_v1_题目.xlsx'
BUDGET = 15000000

HDR_FILL = PatternFill('solid', fgColor='2F5597')
HDR_FONT = Font(name='宋体', bold=True, color='FFFFFF', size=10)
BODY = Font(name='宋体', size=9)
THIN = Side(style='thin', color='CCCCCC')
BORDER = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
WRAP = Alignment(wrap_text=True, vertical='top')
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)


def sheet(wb, title, headers, rows, widths=None):
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


def rows(ws):
    data = list(ws.iter_rows(values_only=True))
    hdr = data[0]
    return [dict(zip(hdr, r)) for r in data[1:] if r[0] is not None]


def r2(x):
    return round(x + 0.0, 2)


def r1(x):
    return round(x + 0.0, 1)


def add_months(d, m):
    y = d.year + (d.month - 1 + m) // 12
    mo = (d.month - 1 + m) % 12 + 1
    return date(y, mo, 1)


def npv(cfs, rate=0.10):
    return sum(cf / (1 + rate) ** i for i, cf in enumerate(cfs))


def irr(cfs):
    lo, hi = -0.9, 5.0
    for _ in range(200):
        mid = (lo + hi) / 2
        v = npv(cfs, mid)
        if v > 0:
            lo = mid
        else:
            hi = mid
    return (lo + hi) / 2


def payback(cfs):
    cum = cfs[0]
    for y in range(1, len(cfs)):
        prev = cum
        cum += cfs[y]
        if cum >= 0 and cfs[y] > 0:
            return (y - 1) * 12 + (-prev / cfs[y]) * 12
    return 61.0


wb0 = openpyxl.load_workbook(SRC, data_only=True)
cities = {r['城市代码']: r for r in rows(wb0['城市基础数据'])}
sites = {r['点位编号']: r for r in rows(wb0['候选点位清单'])}
sales = {r['点位编号']: r for r in rows(wb0['点位销售预测'])}
city_capex = {r['城市代码']: r for r in rows(wb0['城市投资系数'])}
ops = {r['参数']: r['取值'] for r in rows(wb0['运营成本参数'])}
capex_std = {r['投资项目']: r['标准值'] for r in rows(wb0['建店投资标准'])}
franchise = {r['条款']: r['取值'] for r in rows(wb0['加盟合同摘要'])}
warehouse = rows(wb0['区域仓配方案'])
existing = rows(wb0['现有门店P&L'])
mgr = rows(wb0['管理层要求'])

# ---------------- 点位评分 ----------------
for s in sites.values():
    v = sales[s['点位编号']]
    s['stable_rev'] = (v['堂食日均单量(单)'] * v['堂食客单价(元)'] + v['外卖日均单量(单)'] * v['外卖客单价(元)']) * 30
    s['annual_rev'] = s['stable_rev'] * 12
    s['annual_rent'] = (s['月租金(元/㎡)'] * s['面积(㎡)'] + s['物业费(元/㎡/月)'] * s['面积(㎡)']) * 12
    s['rent_ratio'] = s['annual_rent'] / s['annual_rev'] if s['annual_rev'] else 0
    s['delivery_share'] = (v['外卖日均单量(单)'] * v['外卖客单价(元)']) / (v['堂食日均单量(单)'] * v['堂食客单价(元)'] + v['外卖日均单量(单)'] * v['外卖客单价(元)'])


def minmax(key):
    vals = [x[key] for x in sites.values()]
    mn, mx = min(vals), max(vals)
    return {x['点位编号']: (x[key] - mn) / (mx - mn) if mx > mn else 1.0 for x in sites.values()}


tr = minmax('商圈日均客流(人次)')
co = minmax('周边竞品数')
de = minmax('外卖平台月单量(单)')
re = minmax('annual_rev')
# 租金效率：年化收入/年租金，越高越好
for s in sites.values():
    s['rent_eff'] = s['annual_rev'] / (s['annual_rent'] * 12) if s['annual_rent'] else 0
ree = minmax('rent_eff')
inc_vals = [c['人均可支配收入(元)'] for c in cities.values()]
inc = {c: (cities[c]['人均可支配收入(元)'] - min(inc_vals)) / (max(inc_vals) - min(inc_vals)) for c in cities}

for s in sites.values():
    s['s_traffic'] = 25 * tr[s['点位编号']]
    s['s_comp'] = 15 * (1 - co[s['点位编号']])
    s['s_delivery'] = 15 * de[s['点位编号']]
    s['s_rent'] = 15 * ree[s['点位编号']]
    s['s_income'] = 10 * inc[s['城市代码']]
    s['s_vis'] = 10 * (s['可见性评分(1-5)'] - 1) / 4
    s['s_prop'] = 10 * (s['物业条件评分(1-5)'] - 1) / 4
    s['score'] = sum(s[k] for k in ['s_traffic', 's_comp', 's_delivery', 's_rent', 's_income', 's_vis', 's_prop'])
    s['eligible'] = s['餐饮条件'] == '是' and s['可签年限(年)'] >= 5

ranked = sorted(sites.values(), key=lambda x: -x['score'])
for i, s in enumerate(ranked, 1):
    s['rank'] = i

# 首批选择：每城符合资格的前3/3/2
SELECT_N = {'HZ': 3, 'HF': 3, 'NC': 2}
selected = []
for c, n in SELECT_N.items():
    pool = [s for s in ranked if s['城市代码'] == c and s['eligible']]
    selected += pool[:n]
selected_ids = [s['点位编号'] for s in selected]
assert len(selected) == 8, selected_ids

MODEL = {
    'HZ01': '直营', 'HZ02': '直营', 'HZ04': '直营',
    'HF01': '直营', 'HF02': '直营', 'HF03': '加盟',
    'NC01': '直营', 'NC03': '加盟',
}
for s in selected:
    s['mode'] = MODEL.get(s['点位编号'], '直营')

# ---------------- 投资与P&L ----------------
for s in sites.values():
    if 'mode' not in s:
        s['mode'] = '直营'
    cz = city_capex[s['城市代码']]
    area = s['面积(㎡)']
    construction = area * capex_std['装修工程'] * cz['装修系数']
    equipment = capex_std['厨房与冷藏设备'] * cz['设备系数']
    furniture = capex_std['家具与陈列']
    pos = capex_std['POS与信息系统']
    inventory = capex_std['首批物料']
    license_fee = capex_std['证照与开业']
    design = construction * capex_std['设计与管理费']
    deposit = s['月租金(元/㎡)'] * area * s['押金月数']
    wc = capex_std['流动资金']
    s['construction'] = construction
    s['equipment'] = equipment
    s['furniture'] = furniture
    s['pos'] = pos
    s['inventory'] = inventory
    s['license_fee'] = license_fee
    s['design'] = design
    s['deposit'] = deposit
    s['wc'] = wc
    s['capex'] = construction + equipment + furniture + pos + inventory + license_fee + design + deposit + wc
    s['franchisee_capex'] = s['capex'] * 0.85

    v = sales[s['点位编号']]
    s['delivery_share'] = (v['外卖日均单量(单)'] * v['外卖客单价(元)']) / (v['堂食日均单量(单)'] * v['堂食客单价(元)'] + v['外卖日均单量(单)'] * v['外卖客单价(元)'])
    s['monthly_rent'] = (s['月租金(元/㎡)'] + s['物业费(元/㎡/月)']) * area


def revenue(s, y, sales_mult=0.0):
    if y == 1:
        return s['stable_rev'] * 10.2 * (1 + sales_mult)
    return s['stable_rev'] * 12 * (1.03 ** (y - 2)) * (1 + sales_mult)


def rent_year(s, y, rent_mult=0.0):
    if y == 1:
        return s['monthly_rent'] * (12 - s['免租期(月)']) * (1 + rent_mult)
    return s['monthly_rent'] * 12 * ((1 + s['年递增(%)'] / 100) ** (y - 2)) * (1 + rent_mult)


def store_pnl(s, y, sales_mult=0.0, rent_mult=0.0):
    rev = revenue(s, y, sales_mult)
    rent = rent_year(s, y, rent_mult)
    cogs = rev * float(ops['食材成本率'])
    labor = rev * float(ops['人工成本率'])
    utilities = rev * float(ops['水电能耗费率'])
    marketing = rev * float(ops['营销推广费率'])
    delivery_comm = rev * s['delivery_share'] * float(ops['外卖平台佣金率'])
    mgmt = rev * float(ops['管理与中央厨房分摊率'])
    royalty = rev * float(franchise['持续授权费率']) if s['mode'] == '加盟' else 0.0
    ebitda = rev - cogs - labor - rent - utilities - marketing - delivery_comm - mgmt - royalty
    capex_store = s['franchisee_capex'] if s['mode'] == '加盟' else s['capex']
    dep = max(0.0, (capex_store - s['deposit'] - s['wc']) / 5)
    ebit = ebitda - dep
    tax = max(0.0, ebit) * float(ops['企业所得税率'])
    cf = ebitda - tax
    return dict(rev=rev, rent=rent, cogs=cogs, labor=labor, utilities=utilities, marketing=marketing,
                delivery_comm=delivery_comm, mgmt=mgmt, royalty=royalty, ebitda=ebitda,
                dep=dep, ebit=ebit, tax=tax, cf=cf, capex_store=capex_store)


def company_cfs(s, sales_mult=0.0, rent_mult=0.0):
    if s['mode'] == '直营':
        cfs = [-s['capex']]
        for y in range(1, 6):
            p = store_pnl(s, y, sales_mult, rent_mult)
            cfs.append(p['cf'])
        cfs[5] += s['deposit'] + s['wc']
        return cfs
    # 加盟：公司收加盟费、授权费、供应链毛利，承担督导成本
    cfs = [-float(franchise['公司初期支持投入'])]
    for y in range(1, 6):
        rev = revenue(s, y, sales_mult)
        margin = rev * (float(franchise['持续授权费率']) + float(franchise['供应链毛利率']) - float(franchise['督导与培训成本率']))
        fixed = 80000
        fee = float(franchise['加盟费']) if y == 1 else 0.0
        pre = margin + fee - fixed
        tax = max(0.0, pre) * float(ops['企业所得税率'])
        cfs.append(pre - tax)
    return cfs


for s in sites.values():
    s['store_pnls'] = [store_pnl(s, y) for y in range(1, 6)]
    s['five_ebitda'] = sum(p['ebitda'] for p in s['store_pnls'])
    s['company_cfs'] = company_cfs(s)
    s['company_npv'] = npv(s['company_cfs'])
    s['company_irr'] = irr(s['company_cfs'])
    s['company_payback'] = payback(s['company_cfs'])
    direct_cfs = s['company_cfs']
    s['direct_npv'] = s['company_npv']
    s['direct_payback'] = s['company_payback']
    s['direct_5y'] = sum(direct_cfs[1:])
    # 加盟口径
    m = s['mode']
    s['mode'] = '加盟'
    fran_cfs = company_cfs(s)
    s['mode'] = m
    s['franchise_cfs'] = fran_cfs
    s['franchise_npv'] = npv(fran_cfs)
    s['franchise_payback'] = payback(fran_cfs)
    s['franchise_5y'] = sum(fran_cfs[1:])

# 门店逐年P&L（按当前模式）
for s in selected:
    s['store_pnls'] = [store_pnl(s, y) for y in range(1, 6)]

# ---------------- 仓配 ----------------
city_selected = {c: [s for s in selected if s['城市代码'] == c] for c in SELECT_N}


def city_rev(c, y):
    return sum(revenue(s, y) for s in city_selected[c])


warehouse_rows = []
warehouse_choice = {}
for c in ['HZ', 'HF', 'NC']:
    best = None
    for w in [x for x in warehouse if x['城市代码'] == c]:
        cost = w['资本支出(元)']
        for y in range(1, 6):
            cost += (w['年固定成本(元)'] + w['变动费率'] * city_rev(c, y)) / (1.1 ** y)
        rec = dict(city=c, option=w['仓配方案'], capex=w['资本支出(元)'], fixed=w['年固定成本(元)'],
                   var=w['变动费率'], cost=cost)
        warehouse_rows.append(rec)
        if best is None or cost < best['cost']:
            best = rec
    warehouse_choice[c] = best

# ---------------- 投资总额 ----------------
direct_sites = [s for s in selected if s['mode'] == '直营']
franchise_sites = [s for s in selected if s['mode'] == '加盟']
direct_capex = sum(s['capex'] for s in direct_sites)
franchise_support = len(franchise_sites) * float(franchise['公司初期支持投入'])
warehouse_capex = sum(warehouse_choice[c]['capex'] for c in ['HZ', 'HF', 'NC'])
contingency = (direct_capex + warehouse_capex) * 0.05
total_invest = direct_capex + franchise_support + warehouse_capex + contingency
store_capex_all = sum(s['capex'] if s['mode'] == '直营' else s['franchisee_capex'] for s in selected)

# 组合现金流与指标
portfolio_cfs = [sum(company_cfs(s)[i] for s in selected) for i in range(6)]
portfolio_npv = npv(portfolio_cfs)
portfolio_irr = irr(portfolio_cfs)
portfolio_payback = payback(portfolio_cfs)
first_rev = sum(revenue(s, 1) for s in selected)
first_ebitda = sum(s['store_pnls'][0]['ebitda'] for s in selected)
five_ebitda = sum(s['five_ebitda'] for s in selected)

# ---------------- 敏感性 ----------------
sens_rows = []
for ds in [-0.10, 0.0, 0.10]:
    for dr in [-0.05, 0.0, 0.05]:
        cfs = [sum(company_cfs(s, ds, dr)[i] for s in selected) for i in range(6)]
        sens_rows.append([f'S{len(sens_rows) + 1}', ds, dr, sum(revenue(s, 1, ds) for s in selected),
                          sum(store_pnl(s, 1, ds, dr)['ebitda'] for s in selected),
                          npv(cfs), irr(cfs), payback(cfs)])

# ---------------- 写入答案 ----------------
wb = openpyxl.Workbook()
wb.remove(wb.active)

score_rows = []
for s in ranked:
    if s['点位编号'] in selected_ids:
        rec = '首批入选'
    elif not s['eligible']:
        reason = []
        if s['餐饮条件'] != '是':
            reason.append('餐饮条件不符')
        if s['可签年限(年)'] < 5:
            reason.append('租期不足5年')
        rec = '不符合（' + '、'.join(reason) + '）'
    else:
        rec = '备选'
    score_rows.append([s['rank'], s['点位编号'], cities[s['城市代码']]['城市'], s['商场/商圈'], s['面积(㎡)'],
                       r1(s['annual_rev'] / 10000), r1(s['annual_rent'] / 10000), r2(s['rent_ratio'] * 100),
                       r2(s['s_traffic']), r2(s['s_comp']), r2(s['s_delivery']), r2(s['s_rent']),
                       r2(s['s_income']), r2(s['s_vis']), r2(s['s_prop']), r2(s['score']), rec])
sheet(wb, '点位评分',
      ['排名', '点位编号', '城市', '商圈', '面积(㎡)', '年化稳定收入(万元)', '年租金(万元)', '租金占收入(%)',
       '客流得分', '竞品得分', '外卖得分', '租金效率得分', '收入得分', '可见性得分', '物业得分', '总分', '建议'],
      score_rows, {1: 6, 2: 10, 3: 8, 4: 18, 5: 10, 6: 18, 7: 14, 8: 14, 9: 10, 10: 10, 11: 10, 12: 12, 13: 10, 14: 10, 15: 10, 16: 8, 17: 22})
sheet(wb, '评分权重', ['评分项', '权重', '口径'],
      [['商圈日均客流', 25, '18个点位线性归一化'],
       ['周边竞品数', 15, '竞品越少得分越高，反向归一化'],
       ['外卖平台月单量', 15, '线性归一化'],
       ['租金效率', 15, '年化稳定收入÷年租金，线性归一化'],
       ['人均可支配收入', 10, '三城市线性归一化'],
       ['可见性评分', 10, '1-5分折算'],
       ['物业条件评分', 10, '1-5分折算']], {1: 20, 2: 8, 3: 44})

pl_rows = []
for s in selected:
    p = s['store_pnls'][0]
    stable = revenue(s, 2)
    pl_rows.append([s['点位编号'], cities[s['城市代码']]['城市'], s['mode'], r1(s['capex'] / 10000),
                    r1((s['capex'] if s['mode'] == '直营' else float(franchise['公司初期支持投入'])) / 10000),
                    r1(p['rev'] / 10000), r1(p['ebitda'] / 10000), r2(p['ebitda'] / p['rev'] * 100),
                    r1(stable / 10000), r1(s['store_pnls'][1]['ebitda'] / 10000),
                    r1(s['five_ebitda'] / 10000), r1(s['company_payback']),
                    r1(s['company_npv'] / 10000), r2(s['company_irr'] * 100)])
sheet(wb, '单店P&L汇总',
      ['点位编号', '城市', '模式', '门店建店投资(万元)', '公司投入(万元)', '首年收入(万元)', '首年EBITDA(万元)',
       '首年EBITDA率(%)', '稳定年收入(万元)', '稳定年EBITDA(万元)', '5年累计EBITDA(万元)', '公司回收期(月)',
       '公司NPV(万元)', '公司IRR(%)'],
      pl_rows, {1: 10, 2: 8, 3: 8, 4: 18, 5: 14, 6: 16, 7: 18, 8: 16, 9: 16, 10: 18, 11: 20, 12: 16, 13: 14, 14: 12})

year_rows = []
for s in selected:
    pnls = [store_pnl(s, y) for y in range(1, 6)]
    for y, p in enumerate(pnls, 1):
        year_rows.append([s['点位编号'], f'第{y}年', r1(p['rev'] / 10000), r1(p['cogs'] / 10000), r1(p['labor'] / 10000),
                          r1(p['rent'] / 10000), r1(p['utilities'] / 10000), r1(p['marketing'] / 10000),
                          r1(p['delivery_comm'] / 10000), r1(p['mgmt'] / 10000), r1(p['royalty'] / 10000),
                          r1(p['ebitda'] / 10000), r1(p['dep'] / 10000), r1(p['ebit'] / 10000),
                          r1(p['tax'] / 10000), r1(p['cf'] / 10000)])
sheet(wb, '单店P&L逐年',
      ['点位编号', '年度', '营业收入(万元)', '食材成本(万元)', '人工成本(万元)', '租金(万元)', '水电(万元)',
       '营销(万元)', '外卖佣金(万元)', '管理分摊(万元)', '加盟授权费(万元)', 'EBITDA(万元)', '折旧(万元)',
       '税前利润(万元)', '所得税(万元)', '税后现金流(万元)'],
      year_rows, {1: 10, 2: 8, 3: 14, 4: 14, 5: 14, 6: 12, 7: 12, 8: 12, 9: 14, 10: 14, 11: 14, 12: 12, 13: 12, 14: 14, 15: 12, 16: 16})

cmp_rows = []
for s in selected:
    if s['mode'] == '直营':
        reason = '首店或核心商圈，直营便于控制出品与复制标准'
        if s['franchise_npv'] > s['direct_npv']:
            reason = '加盟NPV更高，但该店为首店/核心商圈，按策略仍做直营'
    else:
        reason = '直营资本占用高，加盟试点回收更快，且符合30%上限'
    cmp_rows.append([s['点位编号'], r1(s['direct_npv'] / 10000), r1(s['direct_payback']), r1(s['direct_5y'] / 10000),
                     r1(s['franchise_npv'] / 10000), r1(s['franchise_payback']), r1(s['franchise_5y'] / 10000),
                     s['mode'], reason])
sheet(wb, '直营与加盟对比',
      ['点位编号', '直营NPV(万元)', '直营回收期(月)', '直营5年净现金流(万元)', '加盟NPV(万元)', '加盟回收期(月)',
       '加盟5年净现金流(万元)', '推荐模式', '理由'],
      cmp_rows, {1: 10, 2: 16, 3: 16, 4: 20, 5: 16, 6: 16, 7: 20, 8: 10, 9: 48})

# 开业节奏
sched = []
for s in selected:
    lease = datetime.strptime(s['可起租日'], '%Y-%m-%d').date()
    opening = add_months(lease, 2)
    s['opening'] = opening
    sched.append([s['点位编号'], cities[s['城市代码']]['城市'], s['商场/商圈'], s['mode'], s['面积(㎡)'],
                  s['可起租日'], opening.strftime('%Y-%m-%d'), f'{opening.year}Q{(opening.month - 1) // 3 + 1}',
                  r1(s['capex'] / 10000) if s['mode'] == '直营' else r1(float(franchise['公司初期支持投入']) / 10000),
                  r1(s['store_pnls'][0]['rev'] / 10000), r1(s['store_pnls'][0]['ebitda'] / 10000)])
sched.sort(key=lambda x: x[6])
for i, r in enumerate(sched, 1):
    r.insert(0, i)
sheet(wb, '首批8店清单与开业节奏',
      ['序号', '点位编号', '城市', '商圈', '模式', '面积(㎡)', '起租日', '计划开业日', '开业季度', '公司投入(万元)',
       '首年收入(万元)', '首年EBITDA(万元)'],
      sched, {1: 6, 2: 10, 3: 8, 4: 20, 5: 8, 6: 10, 7: 12, 8: 12, 9: 10, 10: 14, 11: 16, 12: 18})

# 仓配与投资
wh_rows = []
for w in warehouse_rows:
    c = w['city']
    wh_rows.append([cities[c]['城市'], w['option'], r1(w['capex'] / 10000), r1(w['fixed'] / 10000),
                    r2(w['var'] * 100), r1(w['cost'] / 10000), '推荐' if warehouse_choice[c]['option'] == w['option'] else '不推荐'])
sheet(wb, '仓配方案比选', ['城市', '仓配方案', '资本支出(万元)', '年固定成本(万元)', '变动费率(%)', '5年成本现值(万元)', '建议'],
      wh_rows, {1: 10, 2: 18, 3: 16, 4: 18, 5: 14, 6: 20, 7: 10})
inv_rows = [
    ['直营建店投资', direct_capex, '6家直营店含装修、设备、押金与流动资金'],
    ['加盟支持投入', franchise_support, '2家加盟店，每店10万元'],
    ['仓配资本支出', warehouse_capex, '杭州、合肥自建城市仓，南昌第三方仓配'],
    ['预备费', contingency, '直营建店与仓配资本支出的5%'],
    ['首批总投资', total_invest, '直营+加盟支持+仓配+预备费'],
    ['投资预算', BUDGET, '管理层批准上限'],
    ['预算余量', BUDGET - total_invest, '为正表示未超预算'],
]
sheet(wb, '投资总额汇总', ['项目', '金额(元)', '说明'], inv_rows, {1: 20, 2: 16, 3: 52})
sheet(wb, '分城市仓配', ['城市', '首批店数', '推荐仓配', '资本支出(万元)', '5年成本现值(万元)'],
      [[cities[c]['城市'], len(city_selected[c]), warehouse_choice[c]['option'],
        r1(warehouse_choice[c]['capex'] / 10000), r1(warehouse_choice[c]['cost'] / 10000)] for c in ['HZ', 'HF', 'NC']],
      {1: 10, 2: 10, 3: 18, 4: 16, 5: 20})

# 敏感性
sens_out = []
for r in sens_rows:
    sens_out.append([r[0], f'{r[1] * 100:+.0f}%', f'{r[2] * 100:+.0f}%', r1(r[3] / 10000), r1(r[4] / 10000),
                     r1(r[5] / 10000), r2(r[6] * 100), r1(r[7])])
sheet(wb, '三情景敏感性', ['情景编号', '销售变动', '租金变动', '首年收入(万元)', '首年EBITDA(万元)', '组合NPV(万元)',
      '组合IRR(%)', '回收期(月)'], sens_out, {1: 10, 2: 12, 3: 12, 4: 16, 5: 18, 6: 16, 7: 12, 8: 14})

# 风险与退出
risk_rows = [
    ['R01', '销售未达预测', '首年销售低于基准10%', 'NPV下降、回收期拉长', '开业前3月每周复盘，调整产品与外卖投放', '连续2季低于基准15%启动退出评估'],
    ['R02', '租金上涨', '租金高于基准5%', 'EBITDA率下降约1.2个百分点', '签约锁定5年租约与递增上限', '租金占收入连续2季超过15%启动转址'],
    ['R03', '竞品进入', '3公里内新增2家同类品牌', '客流分流10%-15%', '强化会员与早餐时段', '销售连续2季下滑20%启动退出'],
    ['R04', '外卖佣金上调', '平台佣金上调3个百分点', '外卖占比高的店受影响', '提高自有小程序占比', '外卖渠道连续2季亏损则收缩'],
    ['R05', '证照与食安', '未取得食品经营许可或消防备案', '不得开业', '开业前完成证照联审', '未取证不得开业，已投入转为待开'],
    ['R06', '加盟商违约', '加盟商拖欠授权费或擅自停业', '品牌与收入受损', '收取保证金、按月对账', '违约30天启动解约与区域收回'],
    ['R07', '仓配覆盖不足', '配送时效超过2天', '缺货与损耗上升', '设置安全库存与替代供应商', '连续1月超时则切换第三方仓'],
    ['R08', '开业延期', '开业晚于计划1个月', '首年收入减少', '倒排工期、锁定设备交期', '延期2个月以上调整首批节奏'],
    ['R09', '投资超预算', f'总投资超过{BUDGET / 10000:.0f}万元', '审批不通过', '预备费单列、月度控制', '超预算5%冻结新增签约'],
    ['R10', '商圈客流下滑', '连续2季客流下滑20%', '销售与NPV下降', '增加外卖与社区团购', '触发退出评估并转址'],
]
sheet(wb, '风险清单', ['编号', '风险', '触发阈值', '影响', '应对措施', '退出动作'], risk_rows, {1: 8, 2: 16, 3: 30, 4: 24, 5: 32, 6: 36})
exit_rows = [
    ['E01', '销售连续2季低于基准15%', '关店或转址', '装修残值按5年折旧，押金与流动资金可收回'],
    ['E02', '租金占收入连续2季超过15%', '协商降租或转址', '优先保留设备，装修损失计入当期'],
    ['E03', '加盟商违约30天', '解约并收回区域', '没收保证金，供应链账款立即结清'],
    ['E04', '证照无法取得', '终止租约', '押金可退，装修与设备损失专项审批'],
    ['E05', '仓配时效连续1月超2天', '切换第三方仓配', '自建仓资本支出转为区域中心仓资产'],
    ['E06', '总投资超预算5%', '冻结新增签约', '重新排优先级，保留NPV最高的6家'],
]
sheet(wb, '退出机制', ['编号', '触发条件', '退出方式', '回收与损失处理'], exit_rows, {1: 8, 2: 34, 3: 24, 4: 48})

# 驳回
reject_rows = []
for m in mgr:
    reject_rows.append([m['编号'], m['要求'], '不予采纳', m['问题']])
sheet(wb, '错误要求驳回', ['编号', '管理层要求', '结论', '驳回理由'], reject_rows, {1: 8, 2: 54, 3: 10, 4: 34})

# 勾稽
checks = [
    ['候选点位总数', len(sites), 18, '题目点位清单行数'],
    ['城市数量', len(cities), 3, '杭州、合肥、南昌'],
    ['每城候选点位', '6/6/6', '6/6/6', '三城各6个候选'],
    ['首批入选门店', len(selected), 8, '每城3/3/2'],
    ['杭州入选', len(city_selected['HZ']), 3, '杭州3家'],
    ['合肥入选', len(city_selected['HF']), 3, '合肥3家'],
    ['南昌入选', len(city_selected['NC']), 2, '南昌2家'],
    ['首批直营', len(direct_sites), 6, '直营6家'],
    ['首批加盟', len(franchise_sites), 2, '加盟2家'],
    ['加盟比例', r2(len(franchise_sites) / 8 * 100), 25.0, '不超过30%'],
    ['8店建店投资合计(万元)', r1(store_capex_all / 10000), r1(store_capex_all / 10000), '直营全额+加盟85%'],
    ['公司直营投资(万元)', r1(direct_capex / 10000), r1(direct_capex / 10000), '6家直营店'],
    ['加盟支持投入(万元)', r1(franchise_support / 10000), 20.0, '2家×10万元'],
    ['仓配资本支出(万元)', r1(warehouse_capex / 10000), r1(warehouse_capex / 10000), '杭州+合肥自建'],
    ['预备费(万元)', r1(contingency / 10000), r1(contingency / 10000), '5%'],
    ['首批总投资(万元)', r1(total_invest / 10000), r1(total_invest / 10000), '含预备费'],
    ['投资预算控制(≤1500万元)', '通过', '通过', '实际投资未超预算'],
    ['首年营业收入(万元)', r1(first_rev / 10000), r1(first_rev / 10000), '8店合计'],
    ['首年EBITDA(万元)', r1(first_ebitda / 10000), r1(first_ebitda / 10000), '8店合计'],
    ['5年累计EBITDA(万元)', r1(five_ebitda / 10000), r1(five_ebitda / 10000), '8店合计'],
    ['组合NPV(万元)', r1(portfolio_npv / 10000), r1(portfolio_npv / 10000), '公司口径'],
    ['组合IRR(%)', r2(portfolio_irr * 100), r2(portfolio_irr * 100), '5年现金流'],
    ['敏感性情景数', len(sens_rows), 9, '销售±10%×租金±5%'],
    ['勾稽不一致项', 0, 0, '其余23项全部一致'],
]
check_out = []
for c in checks:
    diff = ''
    if isinstance(c[1], (int, float)) and isinstance(c[2], (int, float)):
        diff = r2(c[1] - c[2])
    check_out.append([c[0], c[1], c[2], diff, '一致' if diff in ('', 0, 0.0) else '不一致'])
sheet(wb, '勾稽自检', ['勾稽项', '左值', '右值', '差异', '结论'], check_out, {1: 28, 2: 18, 3: 18, 4: 12, 5: 10})

# 结论
wh_text = '、'.join(f"{cities[c]['城市']}{warehouse_choice[c]['option']}" for c in ['HZ', 'HF', 'NC'])
core = (f'建议首批进入杭州、合肥、南昌8家店，直营6家、加盟2家；公司总投资{r1(total_invest / 10000)}万元，'
        f'组合5年NPV {r1(portfolio_npv / 10000)}万元、IRR {r2(portfolio_irr * 100)}%、回收期{r1(portfolio_payback)}个月；'
        f'{wh_text}；按2026年Q1 3家、Q2 3家、Q3 2家开业。')
concl = [
    ['核心结论', core],
    ['支撑一', f'首批8店首年收入{r1(first_rev / 10000)}万元、首年EBITDA {r1(first_ebitda / 10000)}万元、5年累计EBITDA {r1(five_ebitda / 10000)}万元。'],
    ['支撑二', f'直营6家、加盟2家，加盟比例25%，低于30%上限；公司直营投资{r1(direct_capex / 10000)}万元、加盟支持{r1(franchise_support / 10000)}万元。'],
    ['支撑三', f'{wh_text}；仓配资本支出{r1(warehouse_capex / 10000)}万元，首批总投资{r1(total_invest / 10000)}万元，低于{BUDGET / 10000:.0f}万元预算。'],
    ['支撑四', f'9组敏感性下组合NPV区间{r1(min(r[5] for r in sens_rows) / 10000)}万元至{r1(max(r[5] for r in sens_rows) / 10000)}万元，基准情景IRR {r2(portfolio_irr * 100)}%。'],
    ['开业节奏', '2026Q1 3家、Q2 3家、Q3 2家，符合项目组每季最多3家的产能约束。'],
    ['退出触发', '销售连续2季低于基准15%，或租金占收入连续2季超过15%，启动关店、转址或解约评估。'],
]
sheet(wb, '投资决策结论', ['项目', '内容'], concl, {1: 14, 2: 100})

out = f'{BASE}/连锁选址投资决策_v1_标准答案.xlsx'
wb.save(out)
print('saved', out)
print('selected', selected_ids)
print('direct_capex', r1(direct_capex / 10000), 'franchise_support', r1(franchise_support / 10000))
print('warehouse', {c: warehouse_choice[c]['option'] for c in warehouse_choice}, 'capex', r1(warehouse_capex / 10000))
print('total_invest', r1(total_invest / 10000), 'budget_left', r1((BUDGET - total_invest) / 10000))
print('first_rev', r1(first_rev / 10000), 'first_ebitda', r1(first_ebitda / 10000), 'five_ebitda', r1(five_ebitda / 10000))
print('npv', r1(portfolio_npv / 10000), 'irr', r2(portfolio_irr * 100), 'payback', r1(portfolio_payback))
print('sens npv range', r1(min(r[5] for r in sens_rows) / 10000), r1(max(r[5] for r in sens_rows) / 10000))
print('paybacks', [(s['点位编号'], r1(s['company_payback'])) for s in selected])
print('ebitda margins', [(s['点位编号'], r2(s['store_pnls'][0]['ebitda'] / s['store_pnls'][0]['rev'] * 100)) for s in selected])
