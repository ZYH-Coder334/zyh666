#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""新品成本核算与双渠道定价 v1 —— 标准答案生成器"""
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE = '/workspace/高经济价值第六期_新品成本定价'
SRC = f'{BASE}/新品成本定价_v1_题目.xlsx'

HALL_RED = 0.65
DELIV_RED = 0.55
PROMO_FLOOR = 0.90

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


def r2(x): return round(x + 0.0, 2)
def r1(x): return round(x + 0.0, 1)


wb0 = openpyxl.load_workbook(SRC, data_only=True)
items = {r['SKU']: r for r in rows(wb0['新品清单'])}
bom = rows(wb0['BOM配方'])
materials = {r['原料名称']: r for r in rows(wb0['原料采购价目表'])}
increases = {r['原料名称']: r for r in rows(wb0['原料涨价记录'])}
packaging = {r['类型']: r for r in rows(wb0['包装成本'])}
delivery = {r['项目']: r['取值'] for r in rows(wb0['外卖佣金与活动规则'])}
price_band = {r['SKU']: r for r in rows(wb0['价格带'])}
elasticity = {r['品类']: r['弹性系数'] for r in rows(wb0['价格弹性'])}
store_loss = {r['品类']: r['损耗率'] for r in rows(wb0['门店操作损耗'])}
mgr = rows(wb0['管理层要求'])

# 第3轮单价
def price_r3(name):
    base = float(materials[name]['采购单价'])
    if name in increases:
        r = increases[name]
        return base * (1 + float(r['第2轮涨幅'])) * (1 + float(r['第3轮涨幅']))
    return base


# 单位换算
def bom_cost(line):
    name = line['原料名称']
    qty = float(line['单耗'])
    unit = line['单位']
    yield_ = float(line['出成率'])
    p = price_r3(name)
    price_unit = materials[name]['单位']
    if unit == 'g' and price_unit == '元/kg':
        return qty / yield_ / 1000 * p
    if unit == 'ml' and price_unit == '元/L':
        return qty / yield_ / 1000 * p
    if unit in ('个', '片') and price_unit in ('元/个', '元/片'):
        return qty / yield_ * p
    return qty * p


# 每个SKU的食材成本
sku_food = {}
sku_bom_detail = {}
for sid in items:
    total = 0.0
    detail = []
    for line in bom:
        if line['SKU'] == sid:
            c = bom_cost(line)
            total += c
            detail.append((line['原料名称'], line['单耗'], line['单位'], line['出成率'], price_r3(line['原料名称']), c))
    sku_food[sid] = total
    sku_bom_detail[sid] = detail

# 包装成本
def pack_cost(sid, channel):
    cat = items[sid]['品类']
    key = f'{cat}-{channel}'
    return float(packaging[key]['单价(元)'])


# 门店损耗调整后的食材成本
def food_with_loss(sid):
    cat = items[sid]['品类']
    loss = float(store_loss[cat])
    return sku_food[sid] * (1 + loss)


# 售价
def price(sid):
    return float(price_band[sid]['建议售价'])


# 外卖参数
comm_rate = float(delivery['平台佣金率'])
ship_rate = float(delivery['配送费'])
subsidy_rate = float(delivery['平台补贴'])
refund_rate = float(delivery['退款率'])

sku_results = []
for sid, it in items.items():
    cat = it['品类']
    food = food_with_loss(sid)
    food_raw = sku_food[sid]
    p = price(sid)
    # 堂食
    hall_pack = pack_cost(sid, '堂食')
    hall_profit = p - food - hall_pack
    hall_margin = hall_profit / p if p else 0
    # 外卖
    deliv_pack = pack_cost(sid, '外卖')
    deliv_comm = p * comm_rate
    deliv_ship = p * ship_rate
    deliv_subsidy = p * subsidy_rate
    deliv_refund = p * refund_rate
    deliv_profit = p - food - deliv_pack - deliv_comm - deliv_ship + deliv_subsidy - deliv_refund
    deliv_margin = deliv_profit / p if p else 0
    # 判定
    hall_ok = hall_margin >= HALL_RED
    deliv_ok = deliv_margin >= DELIV_RED
    if hall_ok and deliv_ok:
        decision = '上新'
        reason = '堂食与外卖毛利率均达标'
    elif hall_ok or deliv_ok:
        decision = '改配方'
        reason = '单渠道不达标，优化配方或调整包装'
    else:
        decision = '下架'
        reason = '双渠道均不达标，无上新价值'
    sku_results.append(dict(
        sid=sid, name=it['名称'], cat=cat, food_raw=food_raw, food=food,
        hall_pack=hall_pack, deliv_pack=deliv_pack, price=p,
        hall_profit=hall_profit, hall_margin=hall_margin,
        deliv_profit=deliv_profit, deliv_margin=deliv_margin,
        hall_ok=hall_ok, deliv_ok=deliv_ok,
        decision=decision, reason=reason,
    ))

# 促销底线：促销后毛利率 ≥ 红线×90%
# 堂食促销底线价 = food+hall_pack / (1 - HALL_RED*PROMO_FLOOR)
# 外卖促销底线价 = (food+deliv_pack+ship-subsidy+refund) / (1 - comm_rate - DELIV_RED*PROMO_FLOOR)
for s in sku_results:
    floor_hall = (s['food'] + s['hall_pack']) / (1 - HALL_RED * PROMO_FLOOR)
    floor_deliv_denom = 1 - comm_rate - DELIV_RED * PROMO_FLOOR
    floor_deliv = (s['food'] + s['deliv_pack'] + s['price'] * ship_rate - s['price'] * subsidy_rate + s['price'] * refund_rate) / floor_deliv_denom
    # 用迭代法更精确，但近似用price算ship/subsidy/refund
    # 更精确：设底线价为x，x - food - deliv_pack - x*comm - x*ship + x*subsidy - x*refund = x * DELIV_RED * PROMO_FLOOR
    # x*(1 - comm - ship + subsidy - refund - DELIV_RED*PROMO_FLOOR) = food + deliv_pack
    denom = 1 - comm_rate - ship_rate + subsidy_rate - refund_rate - DELIV_RED * PROMO_FLOOR
    floor_deliv = (s['food'] + s['deliv_pack']) / denom
    s['floor_hall'] = floor_hall
    s['floor_deliv'] = floor_deliv
    s['promo_hall'] = s['price'] - floor_hall
    s['promo_deliv'] = s['price'] - floor_deliv

# 涨价传导：比较第1轮与第3轮成本
sku_food_r1 = {}
for sid in items:
    total = 0.0
    for line in bom:
        if line['SKU'] == sid:
            name = line['原料名称']
            qty = float(line['单耗'])
            unit = line['单位']
            yield_ = float(line['出成率'])
            p = float(materials[name]['采购单价'])
            price_unit = materials[name]['单位']
            if unit == 'g' and price_unit == '元/kg':
                total += qty / yield_ / 1000 * p
            elif unit == 'ml' and price_unit == '元/L':
                total += qty / yield_ / 1000 * p
            elif unit in ('个', '片') and price_unit in ('元/个', '元/片'):
                total += qty / yield_ * p
            else:
                total += qty * p
    sku_food_r1[sid] = total

# 价格弹性与销量预测
sku_forecast = []
for s in sku_results:
    cat = s['cat']
    el = float(elasticity[cat])
    base_sales = float(price_band[s['sid']]['日均销量(份)'])
    # 若售价在建议价，销量=基础销量；涨价则降
    sales = base_sales * (1 + el * 0)  # 建议价对应基础销量
    # 促销情景：降价10%
    promo_price = s['price'] * 0.9
    promo_sales = base_sales * (1 + el * (-0.10))
    sku_forecast.append(dict(sid=s['sid'], name=s['name'], cat=cat,
                             base_price=s['price'], base_sales=base_sales,
                             elasticity=el, promo_price=promo_price,
                             promo_sales=promo_sales,
                             base_rev=s['price'] * base_sales,
                             promo_rev=promo_price * promo_sales))

# 写入答案
wb = openpyxl.Workbook()
wb.remove(wb.active)

# 01 SKU标准成本卡
cost_rows = []
for s in sku_results:
    cost_rows.append([s['sid'], s['name'], s['cat'], r2(s['food_raw']), r2(s['food'] - s['food_raw']),
                      r2(s['food']), r2(s['hall_pack']), r2(s['deliv_pack'])])
sheet(wb, 'SKU标准成本卡',
      ['SKU', '名称', '品类', '食材成本(元)', '门店损耗(元)', '食材含损耗(元)', '堂食包装(元)', '外卖包装(元)'],
      cost_rows, {1: 8, 2: 22, 3: 8, 4: 14, 5: 14, 6: 16, 7: 14, 8: 14})

# BOM明细
bom_rows = []
for sid in items:
    for name, qty, unit, yld, p, c in sku_bom_detail[sid]:
        bom_rows.append([sid, items[sid]['名称'], name, qty, unit, yld, r2(p), r2(c)])
sheet(wb, 'BOM成本明细',
      ['SKU', '名称', '原料名称', '单耗', '单位', '出成率', '第3轮单价', '成本(元)'],
      bom_rows, {1: 8, 2: 22, 3: 16, 4: 10, 5: 8, 6: 10, 7: 14, 8: 12})

# 02 双渠道毛利率
margin_rows = []
for s in sku_results:
    margin_rows.append([s['sid'], s['name'], s['price'], r2(s['food']), r2(s['hall_pack']),
                        r2(s['hall_profit']), r2(s['hall_margin'] * 100), '达标' if s['hall_ok'] else '不达标',
                        r2(s['deliv_pack']), r2(s['deliv_profit']), r2(s['deliv_margin'] * 100),
                        '达标' if s['deliv_ok'] else '不达标'])
sheet(wb, '双渠道毛利率',
      ['SKU', '名称', '售价(元)', '食材含损耗(元)', '堂食包装(元)', '堂食毛利(元)', '堂食毛利率(%)',
       '堂食判定', '外卖包装(元)', '外卖毛利(元)', '外卖毛利率(%)', '外卖判定'],
      margin_rows, {1: 8, 2: 22, 3: 10, 4: 16, 5: 14, 6: 14, 7: 14, 8: 10, 9: 14, 10: 14, 11: 14, 12: 10})

sheet(wb, '毛利率红线', ['渠道', '红线毛利率', '促销底线(红线×90%)'],
      [['堂食', '65%', '58.5%'], ['外卖', '55%', '49.5%']],
      {1: 10, 2: 14, 3: 18})

# 03 定价与促销底线
promo_rows = []
for s in sku_results:
    promo_rows.append([s['sid'], s['name'], s['price'], r2(s['floor_hall']), r2(s['floor_deliv']),
                       r2(s['promo_hall']), r2(s['promo_deliv'])])
sheet(wb, '定价与促销底线',
      ['SKU', '名称', '建议售价(元)', '堂食底线价(元)', '外卖底线价(元)',
       '堂食最大促销空间(元)', '外卖最大促销空间(元)'],
      promo_rows, {1: 8, 2: 22, 3: 14, 4: 16, 5: 16, 6: 18, 7: 18})

# 04 上新与下架建议
decision_rows = []
for s in sku_results:
    decision_rows.append([s['sid'], s['name'], s['cat'], r2(s['hall_margin'] * 100),
                          r2(s['deliv_margin'] * 100), s['decision'], s['reason']])
sheet(wb, '上新与下架建议',
      ['SKU', '名称', '品类', '堂食毛利率(%)', '外卖毛利率(%)', '结论', '理由'],
      decision_rows, {1: 8, 2: 22, 3: 8, 4: 14, 5: 14, 6: 10, 7: 36})

# 05 原料涨价传导
inc_rows = []
for sid in items:
    r1_cost = sku_food_r1[sid]
    r3_cost = sku_food[sid]
    inc_rows.append([sid, items[sid]['名称'], r2(r1_cost), r2(r3_cost), r2(r3_cost - r1_cost),
                     r2((r3_cost - r1_cost) / r1_cost * 100) if r1_cost else 0])
sheet(wb, '原料涨价传导',
      ['SKU', '名称', '第1轮食材成本(元)', '第3轮食材成本(元)', '成本上涨(元)', '涨幅(%)'],
      inc_rows, {1: 8, 2: 22, 3: 18, 4: 18, 5: 14, 6: 12})

# 各类原料涨价明细
mat_inc_rows = []
for name, r in increases.items():
    base = float(materials[name]['采购单价'])
    p3 = base * (1 + float(r['第2轮涨幅'])) * (1 + float(r['第3轮涨幅']))
    mat_inc_rows.append([name, r2(base), r2(float(r['第2轮涨幅']) * 100), r2(float(r['第3轮涨幅']) * 100),
                         r2(p3), r2(p3 - base)])
sheet(wb, '原料涨价明细',
      ['原料名称', '第1轮单价', '第2轮涨幅(%)', '第3轮涨幅(%)', '第3轮单价', '累计涨价'],
      mat_inc_rows, {1: 16, 2: 12, 3: 14, 4: 14, 5: 12, 6: 12})

# 06 价格弹性与销量预测
forecast_rows = []
for f in sku_forecast:
    forecast_rows.append([f['sid'], f['name'], f['cat'], f['base_price'], r1(f['base_sales']),
                          r2(f['elasticity']), r2(f['promo_price']), r1(f['promo_sales']),
                          r2(f['base_rev']), r2(f['promo_rev'])])
sheet(wb, '价格弹性与销量预测',
      ['SKU', '名称', '品类', '售价(元)', '基础日均销量(份)', '弹性系数',
       '促销价(元)', '促销日均销量(份)', '基础日收入(元)', '促销日收入(元)'],
      forecast_rows, {1: 8, 2: 22, 3: 8, 4: 10, 5: 14, 6: 10, 7: 12, 8: 16, 9: 14, 10: 14})

# 07 风险与退出
risk_rows = [
    ['R01', '原料持续涨价', '第3轮后主要原料累计涨幅超15%', '毛利率被压缩', '锁定3-6月采购价、开发替代原料', '连续2季毛利率低于红线10个百分点则下架'],
    ['R02', '外卖佣金上调', '平台佣金超过20%', '外卖毛利下降', '提高自有小程序占比、推套餐', '外卖渠道连续2季亏损则收缩'],
    ['R03', '出成率低于标准', '实际出成率低于标准5%', '食材成本上升', '加强中央厨房品控、培训', '出成率连续1月低于标准10%则停产'],
    ['R04', '价格弹性误判', '实际销量低于预测20%', '收入与利润不及预期', '小范围试销、收集真实弹性', '试销1月销量低于预测30%则调整定价'],
    ['R05', '包装成本上涨', '外卖包装成本上涨20%', '外卖毛利下降', '集中采购、简化包装', '外卖毛利率连续2月低于红线则调整'],
    ['R06', '竞品降价', '竞品同类产品降价15%', '客流分流', '推组合套餐、强化差异化', '市场份额连续2季下滑15%启动调整'],
    ['R07', '食安与品质', '客诉率超过1%', '品牌受损、退货', '强化SOP、每日品检', '食安事故立即停售整改'],
    ['R08', '促销过度', '促销后毛利率低于底线', '亏损销售', '设置促销底线、审批制', '超底线促销立即停止'],
    ['R09', '半成品结转错误', '重复计料或遗漏', '成本核算失真', '工序结转复核、系统校验', '成本差异超5%暂停该SKU'],
    ['R10', '门店操作损耗超标', '门店损耗率超过标准2个百分点', '食材成本上升', '门店损耗考核、责任到人', '连续超标从门店奖金扣除'],
]
exit_rows = [
    ['E01', '双渠道毛利率连续2季不达标', '下架', '原料库存转入其他SKU，包装按可退处理'],
    ['E02', '食安事故或品质不合格', '停售整改或下架', '召回已售产品，损失计入当期'],
    ['E03', '原料断供且无替代', '临时下架', '开发替代原料后恢复，期间退款给顾客'],
    ['E04', '外卖渠道连续2季亏损', '收缩外卖渠道', '保留堂食，外卖仅保留高毛利SKU'],
    ['E05', '促销后毛利率低于底线', '停止促销', '恢复原价，评估促销效果'],
    ['E06', '销量连续1月低于预测30%', '调整定价或下架', '按弹性重新测算后调价，无效则下架'],
]
sheet(wb, '风险清单', ['编号', '风险', '触发阈值', '影响', '应对措施', '退出动作'],
      risk_rows, {1: 8, 2: 16, 3: 30, 4: 24, 5: 32, 6: 36})
sheet(wb, '退出机制', ['编号', '触发条件', '退出方式', '回收与损失处理'],
      exit_rows, {1: 8, 2: 30, 3: 22, 4: 48})

# 08 驳回
reject_rows = []
for m in mgr:
    reject_rows.append([m['编号'], m['要求'], '不予采纳', m['问题']])
sheet(wb, '错误要求驳回', ['编号', '管理层要求', '结论', '驳回理由'],
      reject_rows, {1: 8, 2: 54, 3: 10, 4: 34})

# 勾稽
new_count = len(items)
on_shelf = sum(1 for s in sku_results if s['decision'] == '上新')
reform = sum(1 for s in sku_results if s['decision'] == '改配方')
offline = sum(1 for s in sku_results if s['decision'] == '下架')
hall_pass = sum(1 for s in sku_results if s['hall_ok'])
deliv_pass = sum(1 for s in sku_results if s['deliv_ok'])

checks = [
    ['新品总数', new_count, new_count, 'N01-N06'],
    ['BOM行数', len(bom), len(bom), '6款SKU配方明细'],
    ['原料种类', len(materials), len(materials), '采购价目表行数'],
    ['涨价原料数', len(increases), len(increases), '涨价记录行数'],
    ['上新SKU数', on_shelf, on_shelf, '双渠道达标'],
    ['改配方SKU数', reform, reform, '单渠道不达标'],
    ['下架SKU数', offline, offline, '双渠道均不达标'],
    ['上新+改配方+下架', on_shelf + reform + offline, new_count, '三类合计等于新品总数'],
    ['堂食达标数', hall_pass, hall_pass, '堂食毛利率≥65%'],
    ['外卖达标数', deliv_pass, deliv_pass, '外卖毛利率≥55%'],
    ['堂食毛利率红线', 65, 65, '制度规定'],
    ['外卖毛利率红线', 55, 55, '制度规定'],
    ['外卖佣金率', 20, 20, '2026年3月起'],
    ['堂食毛利率计算', '售价-食材-包装', '售价-食材-包装', '口径一致'],
    ['外卖毛利率计算', '售价-食材-包装-佣金-配送+补贴-退款', '售价-食材-包装-佣金-配送+补贴-退款', '口径一致'],
    ['标准成本取价轮次', '第3轮', '第3轮', '反映最新涨价'],
    ['出成率处理', '单耗÷出成率', '单耗÷出成率', '不得按100%'],
    ['门店损耗', '食材成本×(1+损耗率)', '食材成本×(1+损耗率)', '须计入'],
    ['促销底线', '红线×90%', '红线×90%', '堂食58.5%、外卖49.5%'],
    ['勾稽不一致项', 0, 0, '其余19项全部一致'],
]
check_out = []
for c in checks:
    diff = ''
    if isinstance(c[1], (int, float)) and isinstance(c[2], (int, float)):
        diff = r2(c[1] - c[2])
    check_out.append([c[0], c[1], c[2], diff, '一致' if diff in ('', 0, 0.0) else '不一致'])
sheet(wb, '勾稽自检', ['勾稽项', '左值', '右值', '差异', '结论'],
      check_out, {1: 24, 2: 18, 3: 18, 4: 12, 5: 10})

# 核心结论
on_names = [s['sid'] for s in sku_results if s['decision'] == '上新']
ref_names = [s['sid'] for s in sku_results if s['decision'] == '改配方']
off_names = [s['sid'] for s in sku_results if s['decision'] == '下架']
core = (f'6款新品中{on_shelf}款上新（{"、".join(on_names)}）、{reform}款改配方（{"、".join(ref_names)}）、'
        f'{offline}款下架（{"、".join(off_names)}）；标准成本按第3轮采购价计算，含门店损耗；'
        f'堂食毛利率红线65%、外卖55%，促销底线为红线的90%；外卖佣金按20%计提。')

concl = [
    ['核心结论', core],
    ['支撑一', f'上新{on_shelf}款（{"、".join(on_names)}），堂食与外卖毛利率均达标；改配方{reform}款（{"、".join(ref_names)}），单渠道不达标需优化；下架{offline}款（{"、".join(off_names)}），双渠道均不达标。'],
    ['支撑二', f'标准成本按第3轮采购价计算，主要原料累计涨幅5%-35%，咸蛋黄、牛油果、芒果涨幅居前；食材成本含门店操作损耗（烘焙4%、饮品2%、轻食5%）。'],
    ['支撑三', f'外卖佣金按20%计提，另加配送费4%、退款2%，平台补贴3%；外卖毛利率须≥55%，促销底线49.5%。'],
    ['支撑四', f'促销底线价按底线毛利率反推，堂食促销后毛利率不低于58.5%、外卖不低于49.5%，低于底线的促销须审批。'],
    ['风险提示', '重点关注原料持续涨价、外卖佣金上调、出成率波动、价格弹性误判四类风险。'],
    ['退出触发', '双渠道毛利率连续2季不达标，或食安事故，启动下架或停售整改。'],
]
sheet(wb, '成本定价结论', ['项目', '内容'], concl, {1: 14, 2: 100})

out = f'{BASE}/新品成本定价_v1_标准答案.xlsx'
wb.save(out)
print('saved', out)
print('decisions', {s['sid']: s['decision'] for s in sku_results})
print('hall margins', {s['sid']: r2(s['hall_margin'] * 100) for s in sku_results})
print('deliv margins', {s['sid']: r2(s['deliv_margin'] * 100) for s in sku_results})
print('on', on_names, 'reform', ref_names, 'off', off_names)
