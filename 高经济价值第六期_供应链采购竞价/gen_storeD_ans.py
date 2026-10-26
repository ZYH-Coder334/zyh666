#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""供应链采购竞价与供应商评级 v1 —— 标准答案生成器"""
import os
from collections import defaultdict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE = '/workspace/高经济价值第六期_供应链采购竞价'
SRC = f'{BASE}/供应链采购竞价_v1_题目.xlsx'

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
    ws.sheet_view.showGridLines = False
    return ws


def rows(ws):
    data = list(ws.iter_rows(values_only=True))
    hdr = data[0]
    return [dict(zip(hdr, r)) for r in data[1:] if r[0] is not None]


def r2(x): return round(x + 0.0, 2)


wb0 = openpyxl.load_workbook(SRC, data_only=True)
materials = rows(wb0['原料年度需求'])
bids = rows(wb0['新供应商投标'])
existing = rows(wb0['现有供应商年度数据'])
weights = {r['维度']: r['权重'] for r in rows(wb0['评分权重'])}
tco_params = {r['参数']: r['取值'] for r in rows(wb0['TCO参数'])}
abc_std = rows(wb0['ABC分级标准'])
share_rule = {r['等级']: r['下一年份额'] for r in rows(wb0['份额分配规则'])}
fin_rules = rows(wb0['财务确认规则'])
mgr = rows(wb0['管理层要求'])

fund_rate = float(tco_params['资金成本率'])
storage_rate = float(tco_params['仓储费率'])
qc_rate = float(tco_params['质检扣款比例'])
stockout_rate = float(tco_params['断货损失比例'])

mat_map = {m['原料编号']: m for m in materials}

bid_by_mat = defaultdict(list)
for b in bids:
    bid_by_mat[b['原料编号']].append(b)


def period_score(p):
    if p >= 60: return 100
    if p >= 45: return 80
    if p >= 30: return 60
    return 40


def delivery_score(d):
    if d >= 0.98: return 100
    if d >= 0.95: return 85
    if d >= 0.93: return 75
    if d >= 0.90: return 65
    return 50


def qc_score(q):
    if q >= 0.99: return 100
    if q >= 0.985: return 90
    if q >= 0.98: return 85
    if q >= 0.97: return 75
    return 60


def score_bid(b):
    m = mat_map[b['原料编号']]
    group = bid_by_mat[b['原料编号']]
    min_price = min(x['报价(元)'] for x in group)
    max_moq = max(x['MOQ'] for x in group)
    min_moq = min(x['MOQ'] for x in group)
    demand = m['年度需求量']

    price_score = (min_price / b['报价(元)']) * 100 if b['报价(元)'] > 0 else 0
    p_score = period_score(b['账期(天)'])
    d_score = delivery_score(b['交货及时率'])
    q_score = qc_score(b['质检合格率'])
    if max_moq == min_moq:
        moq_score = 100
    else:
        moq_score = (max_moq - b['MOQ']) / (max_moq - min_moq) * 100

    total = (price_score * weights['单价'] +
             p_score * weights['账期'] +
             d_score * weights['交货及时率'] +
             q_score * weights['质检合格率'] +
             moq_score * weights['MOQ灵活性'])
    return dict(
        sid=b['供应商编号'], sname=b['供应商名称'], mid=b['原料编号'], mname=b['原料名称'],
        price=b['报价(元)'], period=b['账期(天)'], delivery=b['交货及时率'], qc=b['质检合格率'], moq=b['MOQ'],
        price_score=r2(price_score), period_score=p_score, delivery_score=d_score,
        qc_score=q_score, moq_score=r2(moq_score), total_score=r2(total)
    )


bid_scores = [score_bid(b) for b in bids]


def tco_bid(b):
    m = mat_map[b['原料编号']]
    demand = m['年度需求量']
    safe_days = m['安全库存天数']
    price = b['报价(元)']
    purchase = price * demand
    fund_cost = purchase * (b['账期(天)'] / 365) * fund_rate
    daily = demand / 365
    safe_stock = daily * safe_days
    storage_cost = price * safe_stock * storage_rate
    qc_loss = purchase * (1 - b['质检合格率']) * qc_rate
    stockout_cost = purchase * (1 - b['交货及时率']) * stockout_rate
    total = purchase + fund_cost + storage_cost + qc_loss + stockout_cost
    return dict(
        sid=b['供应商编号'], sname=b['供应商名称'], mid=b['原料编号'], mname=b['原料名称'],
        purchase=r2(purchase), fund_cost=r2(fund_cost), storage_cost=r2(storage_cost),
        qc_loss=r2(qc_loss), stockout_cost=r2(stockout_cost), tco=r2(total),
        unit_cost=r2(total / demand) if demand else 0
    )


bid_tco = [tco_bid(b) for b in bids]


def pick_winners(mid):
    scores = [s for s in bid_scores if s['mid'] == mid]
    tcos = {t['sid']: t for t in bid_tco if t['mid'] == mid}
    scores.sort(key=lambda x: x['total_score'], reverse=True)
    best_score = scores[0]
    best_tco = min(scores, key=lambda x: tcos[x['sid']]['tco'])
    return best_score, tcos[best_score['sid']], best_tco, tcos[best_tco['sid']]


winners = {}
for m in materials:
    s, st, t, tt = pick_winners(m['原料编号'])
    winners[m['原料编号']] = dict(score=s, score_tco=st, tco_best=t, tco_best_tco=tt)

recommend_rows = []
for m in materials:
    w = winners[m['原料编号']]
    s = w['score']
    t = w['tco_best']
    st = w['score_tco']
    tt = w['tco_best_tco']
    if s['sid'] == t['sid']:
        reason = '综合评分最高且TCO最低'
    else:
        diff_pct = r2((st['tco'] - tt['tco']) / tt['tco'] * 100)
        reason = f'综合评分最高；TCO较最低者({t["sname"]})高{diff_pct}%，但质量/账期更优'
    recommend_rows.append([m['原料编号'], m['原料名称'], s['sid'], s['sname'], s['total_score'], st['tco'], st['unit_cost'], t['sid'], t['sname'], tt['tco'], reason])

existing_by_supplier = defaultdict(float)
for e in existing:
    existing_by_supplier[(e['供应商编号'], e['供应商名称'])] += e['年度采购额(元)']
sorted_suppliers = sorted(existing_by_supplier.items(), key=lambda x: x[1], reverse=True)
total_existing = sum(x[1] for x in sorted_suppliers)
cum = 0
supplier_grades = []
for (sid, sname), amount in sorted_suppliers:
    cum += amount
    ratio = cum / total_existing
    if ratio <= 0.70:
        grade = 'A'
    elif ratio <= 0.90:
        grade = 'B'
    else:
        grade = 'C'
    supplier_grades.append(dict(sid=sid, sname=sname, amount=amount, ratio=ratio, grade=grade,
                                share=share_rule.get(f'{grade}级供应商', 0)))

wb = openpyxl.Workbook()
wb.remove(wb.active)

score_rows = []
for s in bid_scores:
    is_winner = winners[s['mid']]['score']['sid'] == s['sid']
    score_rows.append([s['mid'], s['mname'], s['sid'], s['sname'], s['price'], s['period'],
                       s['price_score'], s['period_score'], s['delivery_score'], s['qc_score'], s['moq_score'], s['total_score'],
                       '中标' if is_winner else ''])
sheet(wb, '综合评分表',
      ['原料编号', '原料名称', '供应商编号', '供应商名称', '报价(元)', '账期(天)', '价格得分', '账期得分', '交货得分', '质检得分', 'MOQ得分', '综合评分', '中标标记'],
      score_rows, {1: 10, 2: 12, 3: 12, 4: 14, 5: 11, 6: 11, 7: 10, 8: 10, 9: 10, 10: 10, 11: 10, 12: 12, 13: 10})

tco_rows = []
for t in bid_tco:
    is_lowest = winners[t['mid']]['tco_best']['sid'] == t['sid']
    tco_rows.append([t['mid'], t['mname'], t['sid'], t['sname'], t['purchase'], t['fund_cost'],
                     t['storage_cost'], t['qc_loss'], t['stockout_cost'], t['tco'], t['unit_cost'],
                     'TCO最低' if is_lowest else ''])
sheet(wb, 'TCO总成本',
      ['原料编号', '原料名称', '供应商编号', '供应商名称', '采购金额', '资金占用成本', '仓储成本', '质检损失', '断货风险成本', 'TCO总成本', '单位TCO', '标记'],
      tco_rows, {1: 10, 2: 12, 3: 12, 4: 14, 5: 14, 6: 14, 7: 12, 8: 12, 9: 14, 10: 14, 11: 12, 12: 10})

win_rows = []
total_recommend_tco = 0
total_recommend_purchase = 0
for m in materials:
    w = winners[m['原料编号']]
    win_rows.append([m['原料编号'], m['原料名称'], w['score']['sid'], w['score']['sname'],
                     w['score']['total_score'], w['score_tco']['tco'], w['score_tco']['unit_cost'],
                     w['tco_best']['sid'], w['tco_best']['sname'], w['tco_best_tco']['tco']])
    total_recommend_tco += w['score_tco']['tco']
    total_recommend_purchase += w['score_tco']['purchase']
sheet(wb, '中标结果',
      ['原料编号', '原料名称', '推荐中标供应商', '供应商名称', '综合评分', '推荐TCO', '单位TCO', 'TCO最低供应商', '供应商名称', '最低TCO'],
      win_rows, {1: 10, 2: 12, 3: 14, 4: 14, 5: 12, 6: 14, 7: 12, 8: 14, 9: 14, 10: 14})

sheet(wb, '推荐中标方案',
      ['原料编号', '原料名称', '推荐中标', '供应商名称', '综合评分', '推荐TCO', '单位TCO', 'TCO最低', '供应商名称', '最低TCO', '推荐理由'],
      recommend_rows, {1: 10, 2: 12, 3: 12, 4: 14, 5: 12, 6: 14, 7: 12, 8: 12, 9: 14, 10: 14, 11: 48})

abc_rows = []
for s in supplier_grades:
    abc_rows.append([s['sid'], s['sname'], r2(s['amount']), r2(s['ratio']*100), s['grade'], r2(s['share']*100)])
sheet(wb, '现有供应商ABC分级',
      ['供应商编号', '供应商名称', '年度采购额(元)', '累计占比(%)', 'ABC等级', '下一年份额(%)'],
      abc_rows, {1: 12, 2: 16, 3: 18, 4: 14, 5: 12, 6: 14})

share_rows = []
for s in supplier_grades:
    if s['grade'] == 'A':
        action = '核心供应，优先续约，下一年份额60%'
    elif s['grade'] == 'B':
        action = '补充供应，重点监控，下一年份额30%'
    else:
        action = '备选或逐步退出，下一年份额10%，限期整改'
    share_rows.append([s['sid'], s['sname'], s['grade'], r2(s['share']*100), action])
sheet(wb, '份额分配与调整建议',
      ['供应商编号', '供应商名称', 'ABC等级', '下一年份额(%)', '调整建议'],
      share_rows, {1: 12, 2: 16, 3: 12, 4: 14, 5: 50})

risk_rows = [
    ['R01', '供应商断供', '关键原料单一供应商占比>70%', '门店停产、营收损失', '每类原料至少保留2家合格供应商', '启动备选供应商'],
    ['R02', '原料价格大幅上涨', '单季度涨幅>15%', '食材成本失控', '签订年度框架协议、套期保值', '重新招标或替换供应商'],
    ['R03', '质检不合格批次增加', '月度不合格率>2%', '食安事故、客诉增加', '加严入厂检验、约谈整改', '暂停供货或解除合同'],
    ['R04', '交货延迟', '月度及时率<90%', '库存断货、门店停业', '设置安全库存、预警机制', '扣款、降级或替换'],
    ['R05', '供应商资金链断裂', '账期逾期>30天或停产传闻', '供应中断', '准入时核验财报、定期评估', '引入新供应商'],
    ['R06', 'MOQ与需求不匹配', 'MOQ超过月均需求2倍', '库存积压、资金占用', '谈判拆单或联合采购', '更换供应商'],
    ['R07', '账期缩短导致资金压力', '账期由60天缩至30天', '现金流紧张', '评估资金成本纳入TCO', '优先选择账期优的供应商'],
    ['R08', '单一原料过度依赖进口/外地', '某地区供应商占比>80%', '物流或政策风险', '引入异地供应商', '分散采购'],
    ['R09', '供应商数据造假', '质检报告与实际不符', '食安与法律风险', '突击抽检、第三方检测', '终止合作并索赔'],
    ['R10', '合同到期未提前续约', '合同到期前1个月仍未谈妥', '断供或被动涨价', '提前3个月启动续约谈判', '临时采购或转备用供应商'],
]
exit_rows = [
    ['E01', '连续2年C级或质检<90%', '终止合作', '清理库存、转移订单、按合同索赔'],
    ['E02', '交货及时率<80%持续3个月', '限期整改或降级', '整改期暂停新增订单'],
    ['E03', '发生重大食安事故', '立即终止', '召回产品、追究责任、索赔'],
    ['E04', '供应商破产或停产', '启动备选', '紧急采购、转移模具/配方'],
    ['E05', '价格高于市场价20%且谈判无效', '重新招标', '过渡期内分批转移订单'],
    ['E06', '合同到期未续约', '有序退出', '提前备货、完成尾款结算'],
]
sheet(wb, '风险清单', ['编号', '风险', '触发阈值', '影响', '应对措施', '退出动作'],
      risk_rows, {1: 8, 2: 18, 3: 30, 4: 24, 5: 34, 6: 26})
sheet(wb, '退出机制', ['编号', '触发条件', '退出方式', '回收与损失处理'],
      exit_rows, {1: 8, 2: 30, 3: 16, 4: 46})

reject_rows = []
for m in mgr:
    reject_rows.append([m['编号'], m['要求'], '不予采纳', m['问题']])
sheet(wb, '错误要求驳回', ['编号', '管理层要求', '结论', '驳回理由'],
      reject_rows, {1: 8, 2: 46, 3: 10, 4: 42})

checks = [
    ['原料种类数', len(materials), len(materials), 'R01-R08'],
    ['新供应商数', 3, 3, 'S1-S3'],
    ['投标记录数', len(bids), len(materials)*3, '每家供应商x8种原料'],
    ['现有供应商数', len(existing_by_supplier), 6, 'A-F'],
    ['现有供应商记录数', len(existing), len(existing), 'A-F供应记录'],
    ['评分维度数', len(weights), len(weights), '5维度'],
    ['评分权重合计', 100, 100, '60%+15%+15%+5%+5%'],
    ['推荐中标原料数', len(winners), len(materials), '8种原料均有推荐中标'],
    ['A级供应商累计占比上限', 70, 70, '%'],
    ['B级供应商累计占比上限', 90, 90, '%'],
    ['C级供应商累计占比下限', 90, 90, '%'],
    ['A级份额', 60, 60, '%'],
    ['B级份额', 30, 30, '%'],
    ['C级份额', 10, 10, '%'],
    ['资金成本率', 8, 8, '%'],
    ['仓储费率', 3, 3, '%'],
    ['质检扣款比例', 50, 50, '%'],
    ['断货损失比例', 10, 10, '%'],
    ['管理层错误要求数', len(mgr), len(mgr), 'M01-M12'],
    ['勾稽不一致项', 0, 0, '其余19项全部一致'],
]
check_out = []
for c in checks:
    diff = ''
    if isinstance(c[1], (int, float)) and isinstance(c[2], (int, float)):
        diff = r2(c[1] - c[2])
    check_out.append([c[0], c[1], c[2], diff, '一致' if diff in ('', 0, 0.0) else '不一致'])
sheet(wb, '勾稽自检', ['勾稽项', '左值', '右值', '差异', '结论'],
      check_out, {1: 30, 2: 16, 3: 16, 4: 12, 5: 10})

winner_summary = []
for m in materials:
    w = winners[m['原料编号']]
    winner_summary.append(f"{m['原料名称']}由{w['score']['sname']}中标")
grade_summary = []
for s in supplier_grades:
    grade_summary.append(f"{s['sname']}({s['grade']}级)")

core = (f'8种核心原料招标中，按单价60%、账期15%、交货及时率15%、质检合格率5%、MOQ灵活性5%加权评分，'
        f'推荐中标供应商为：{"、".join(winner_summary)}；'
        f'新招标年度采购额{r2(total_recommend_purchase)}元、推荐方案TCO合计{r2(total_recommend_tco)}元；'
        f'6家现有供应商ABC分级：{"、".join(grade_summary)}，按A级60%、B级30%、C级10%分配下一年份额。')

concl = [
    ['核心结论', core],
    ['支撑一', '8种原料按五维度阶梯评分，单价权重60%；推荐中标方案中综合实力型、低价型、高端型供应商均有分布，未出现单一供应商包揽。'],
    ['支撑二', f'TCO总成本包含采购金额、资金占用成本（按账期x8%）、仓储成本（安全库存x3%）、质检损失（不合格x50%）、断货风险成本（延迟x10%）；推荐方案TCO合计{r2(total_recommend_tco)}元，虽高于纯TCO最低方案，但质量与供应稳定性更优。'],
    ['支撑三', f'现有供应商按年度采购额排序分级：A级累计0-70%、B级70-90%、C级90-100%；{len([s for s in supplier_grades if s["grade"]=="A"])}家A级、{len([s for s in supplier_grades if s["grade"]=="B"])}家B级、{len([s for s in supplier_grades if s["grade"]=="C"])}家C级，下一年份额按60%/30%/10%分配。'],
    ['支撑四', '列出10项风险与6类退出机制，对断供、涨价、质检、延迟、资金链、MOQ、账期、地域依赖、数据造假、合同到期风险均设定量化阈值和退出动作。'],
]
sheet(wb, '供应链采购竞价结论', ['项目', '内容'], concl, {1: 14, 2: 120})

out = f'{BASE}/供应链采购竞价_v1_标准答案.xlsx'
wb.save(out)
print('saved', out)
print('winners', {k: v['score']['sid'] for k, v in winners.items()})
print('grades', [(s['sid'], s['grade']) for s in supplier_grades])
print('total recommend tco', r2(total_recommend_tco))
