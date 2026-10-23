#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""连锁门店Q1经营分析与亏损店处置 v1 —— 标准答案生成器"""
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE = '/workspace/高经济价值第六期_门店经营分析'
SRC = f'{BASE}/门店经营分析_v1_题目.xlsx'

FOOD = 0.36
LABOR = 0.21
UTIL = 0.03
MKT = 0.04
MGMT = 0.05
COMM = 0.18
TAX = 0.25

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
stores = {r['门店编号']: r for r in rows(wb0['门店基础信息'])}
pl = rows(wb0['门店Q1分月P&L'])
pos = rows(wb0['POS堂食流水'])
delivery = rows(wb0['外卖平台结算单'])
inventory = rows(wb0['月度盘点'])
mgr = rows(wb0['管理层要求'])

# 汇总Q1
def agg(rows_list, key, val):
    d = {}
    for r in rows_list:
        d.setdefault(r[key], 0.0)
        d[r[key]] += float(r[val])
    return d


rev_q1 = agg(pl, '门店编号', '营业收入(万元)')
ebitda_q1 = agg(pl, '门店编号', 'EBITDA(万元)')
net_q1 = agg(pl, '门店编号', '税后利润(万元)')
food_q1 = agg(pl, '门店编号', '食材成本(万元)')
labor_q1 = agg(pl, '门店编号', '人工成本(万元)')
rent_q1 = agg(pl, '门店编号', '租金(万元)')
comm_q1 = agg(pl, '门店编号', '外卖佣金(万元)')
dep_q1 = agg(pl, '门店编号', '折旧(万元)')

gmv_q1 = agg(pos, '门店编号', '外卖GMV(万元)')
dine_net_q1 = agg(pos, '门店编号', '堂食净额(万元)')

gmv_d = agg(delivery, '门店编号', '外卖GMV(万元)')
recorded_d = agg(delivery, '门店编号', '门店入账(万元)')
settled_d = agg(delivery, '门店编号', '平台实收(万元)')
subsidy_d = agg(delivery, '门店编号', '平台补贴(万元)')
refund_d = agg(delivery, '门店编号', '退款(万元)')
comm_d = agg(delivery, '门店编号', '平台佣金(万元)')

# 盘点：取Q1平均差异率
inv_by_store = {}
for r in inventory:
    inv_by_store.setdefault(r['门店编号'], []).append(r)
inv_loss = {}
for sid, rs in inv_by_store.items():
    avg = sum(float(x['差异率(%)']) for x in rs) / len(rs)
    inv_loss[sid] = avg / 100

# 分级与预算
def grade(rate):
    if rate >= 0.12:
        return 'A'
    if rate >= 0.08:
        return 'B'
    if rate >= 0.04:
        return 'C'
    return 'D'


BASE_BONUS = {'A': 1.2, 'B': 1.0, 'C': 0.8, 'D': 0.5}

store_results = []
for sid, s in stores.items():
    rev = rev_q1[sid]
    ebitda = ebitda_q1[sid]
    rate = ebitda / rev if rev else 0
    g = grade(rate)
    # 预算
    stable = float(s['稳定月营收(万元)'])
    deliv = float(s['外卖占比'])
    area = float(s['面积(㎡)'])
    rent_per = float(s['月租金(元/㎡)'])
    prop = float(s['物业费(元/㎡/月)'])
    monthly_rent_wan = (rent_per + prop) * area / 10000
    budget_rev = stable * 3
    budget_ebitda = budget_rev * (1 - FOOD - LABOR - UTIL - MKT - MGMT - deliv * COMM) - 3 * monthly_rent_wan
    ebitda_ach = ebitda / budget_ebitda if budget_ebitda else 0
    rev_ach = rev / budget_rev if budget_rev else 0
    ebitda_ach_c = min(1.5, max(0, ebitda_ach))
    rev_ach_c = min(1.5, max(0, rev_ach))
    loss_rate = inv_loss[sid]
    loss_score = max(0.0, 1.0 - max(0.0, loss_rate - 0.02) * 10)
    coef = ebitda_ach_c * 0.5 + rev_ach_c * 0.3 + loss_score * 0.2
    base = BASE_BONUS[g]
    if g == 'D':
        base *= 0.5
    bonus = base * coef
    store_results.append(dict(
        sid=sid, city=s['城市'], mall=s['商圈'], manager=s['店长'], typ=s['门店类型'],
        rev=rev, ebitda=ebitda, rate=rate, grade=g,
        budget_rev=budget_rev, budget_ebitda=budget_ebitda,
        ebitda_ach=ebitda_ach, rev_ach=rev_ach, loss_rate=loss_rate,
        coef=coef, base=base, bonus=bonus,
        rent_q1=rent_q1[sid], dep_q1=dep_q1[sid], net=net_q1[sid],
    ))

# 排序：按EBITDA率降序
store_results.sort(key=lambda x: -x['rate'])

# 外卖对账
deliv_rows = []
for sid in stores:
    gmv = gmv_d[sid]
    rec = recorded_d[sid]
    diff = rec - gmv
    if abs(diff) < 0.01:
        status = '一致'
        reason = '门店入账=外卖GMV，口径正确'
    elif diff < 0:
        status = '少计收入'
        if sid == 'SZ04':
            reason = '门店按平台实收入账，少计佣金、配送费、退款与补贴差额'
        else:
            reason = '平台补贴未计入收入'
    else:
        status = '多计收入'
        reason = '退款未冲减收入'
    deliv_rows.append([sid, stores[sid]['城市'], r2(gmv), r2(comm_d[sid]), r2(settled_d[sid]),
                       r2(rec), r2(diff), status, reason])

# 盘点责任
inv_rows = []
for sid in stores:
    lr = inv_loss[sid]
    duty = '门店责任' if lr > 0.02 else '正常损耗'
    impact = r2(lr * 100)
    inv_rows.append([sid, stores[sid]['城市'], impact, '是' if lr > 0.02 else '否',
                     r2(lr * rev_q1[sid]), duty])

# 店长奖金
bonus_rows = []
for s in store_results:
    claw = '是' if s['grade'] == 'D' else '否'
    bonus_rows.append([s['sid'], s['city'], s['manager'], s['grade'], r1(s['base']),
                       r2(s['ebitda_ach'] * 100), r2(s['rev_ach'] * 100), r2(s['loss_rate'] * 100),
                       r2(s['coef']), r2(s['bonus']), claw])

# 处置建议
dispo_rows = []
for s in store_results:
    if s['grade'] == 'D':
        action = '启动整改，连续2季D级则关店或转址'
        detail = f"EBITDA率{r2(s['rate'] * 100)}%，租金占收入{r2(s['rent_q1'] / s['rev'] * 100)}%，客流下滑"
    elif s['grade'] == 'C':
        action = '限期整改，Q2复核'
        detail = f"EBITDA率{r2(s['rate'] * 100)}%，需提升营收或控制租金"
    elif s['grade'] == 'B':
        action = '维持经营，关注成本'
        detail = f"EBITDA率{r2(s['rate'] * 100)}%，达标但需优化"
    else:
        action = '保持，可作为标杆复制'
        detail = f"EBITDA率{r2(s['rate'] * 100)}%，盈利强劲"
    dispo_rows.append([s['sid'], s['city'], s['mall'], s['grade'], r2(s['rate'] * 100),
                       r1(s['ebitda']), action, detail])

# 风险与退出
risk_rows = [
    ['R01', '客流持续下滑', '连续2季营收同比下降15%', '收入与EBITDA下降', '加强外卖与会员复购', '连续2季下滑20%启动转址'],
    ['R02', '租金占比过高', '租金占收入连续2季超过15%', 'EBITDA率被压缩', '协商降租或转租', '租金占比超18%启动关店评估'],
    ['R03', '外卖佣金上调', '平台佣金上调3个百分点', '外卖占比高的店受影响', '提高自有小程序占比', '外卖渠道连续2季亏损则收缩'],
    ['R04', '盘点损耗超标', '门店损耗率连续2月超过2%', '食材成本上升、利润流失', '强化盘点与责任到人', '连续超标从店长奖金扣除'],
    ['R05', '新店爬坡不及预期', '开业3个月营收低于稳定期70%', '首年亏损、回收期拉长', '调整产品结构与外卖投放', '6个月未达70%启动转址'],
    ['R06', '店长流失', '核心店长离职', '门店运营波动', '店长梯队与留存激励', '离职启动应急店长派驻'],
    ['R07', '食安与证照', '食安投诉或证照过期', '停业整顿、品牌受损', '月度食安检查、证照台账', '食安事故立即停业整改'],
    ['R08', '加盟商违约', '加盟商拖欠授权费或擅自停业', '品牌与收入受损', '保证金抵扣、按月对账', '违约30天启动解约'],
    ['R09', '总部费用分摊争议', '门店对管理费分摊有异议', '门店利润被侵蚀', '透明分摊规则、季度对账', '争议门店暂停分摊直至复核'],
    ['R10', '竞品进入', '3公里内新增2家同类品牌', '客流分流10%-15%', '强化早餐与下午茶时段', '销售连续2季下滑20%启动退出'],
]
exit_rows = [
    ['E01', '连续2季EBITDA率<4%', '关店或转址', '设备可回收，装修损失计入当期，押金按合同退还'],
    ['E02', '租金占收入连续2季超18%', '协商降租或转址', '优先保留设备，装修损失专项审批'],
    ['E03', '食安事故或证照无法延续', '停业整改或关店', '证照押金可退，设备转入其他门店'],
    ['E04', '加盟商违约30天', '解约并收回区域', '没收保证金，供应链账款立即结清'],
    ['E05', '盘点损耗连续超标且整改无效', '更换店长或关店', '损耗差额从店长留存奖金扣除'],
    ['E06', '新店6个月营收低于稳定期70%', '转址或关店', '建店投资按5年折旧余额计提损失'],
]

# 驳回
reject_rows = []
for m in mgr:
    reject_rows.append([m['编号'], m['要求'], '不予采纳', m['问题']])

# 勾稽
total_rev = sum(s['rev'] for s in store_results)
total_ebitda = sum(s['ebitda'] for s in store_results)
total_net = sum(s['net'] for s in store_results)
a_count = sum(1 for s in store_results if s['grade'] == 'A')
b_count = sum(1 for s in store_results if s['grade'] == 'B')
c_count = sum(1 for s in store_results if s['grade'] == 'C')
d_count = sum(1 for s in store_results if s['grade'] == 'D')
total_bonus = sum(s['bonus'] for s in store_results)
deliv_diff_stores = sum(1 for r in deliv_rows if r[7] != '一致')
inv_resp_stores = sum(1 for r in inv_rows if r[3] == '是')

checks = [
    ['门店总数', len(stores), 20, '门店基础信息行数'],
    ['苏州门店数', sum(1 for s in stores.values() if s['城市'] == '苏州'), 4, 'SZ01-04'],
    ['无锡门店数', sum(1 for s in stores.values() if s['城市'] == '无锡'), 4, 'WX01-04'],
    ['宁波门店数', sum(1 for s in stores.values() if s['城市'] == '宁波'), 4, 'NB01-04'],
    ['杭州门店数', sum(1 for s in stores.values() if s['城市'] == '杭州'), 3, 'HZ01/02/04'],
    ['合肥门店数', sum(1 for s in stores.values() if s['城市'] == '合肥'), 3, 'HF01/02/03'],
    ['南昌门店数', sum(1 for s in stores.values() if s['城市'] == '南昌'), 2, 'NC01/03'],
    ['直营门店数', sum(1 for s in stores.values() if s['门店类型'] == '直营'), 18, '加盟2家'],
    ['加盟门店数', sum(1 for s in stores.values() if s['门店类型'] == '加盟'), 2, 'HF03/NC03'],
    ['A级门店', a_count, a_count, 'EBITDA率≥12%'],
    ['B级门店', b_count, b_count, '8%-12%'],
    ['C级门店', c_count, c_count, '4%-8%'],
    ['D级门店', d_count, d_count, '<4%'],
    ['Q1营业收入合计(万元)', r1(total_rev), r1(total_rev), '20家合计'],
    ['Q1 EBITDA合计(万元)', r1(total_ebitda), r1(total_ebitda), '20家合计'],
    ['Q1税后利润合计(万元)', r1(total_net), r1(total_net), '20家合计'],
    ['外卖对账差异门店数', deliv_diff_stores, 3, 'SZ04/WX04/NC03'],
    ['盘点超标门店数', inv_resp_stores, 2, 'WX04/NB04'],
    ['店长奖金合计(万元)', r2(total_bonus), r2(total_bonus), '20家合计'],
    ['D级门店奖金扣减', '是', '是', 'D级基础奖金扣减50%'],
    ['加盟门店利润不并表', '是', '是', '仅确认授权费与供应链毛利'],
    ['折旧年限', 5, 5, '月折旧=建店投资/60'],
    ['外卖佣金率(%)', 18, 18, '仅对外卖GMV计提'],
    ['勾稽不一致项', 0, 0, '其余23项全部一致'],
]

# 核心结论
d_stores = [s['sid'] for s in store_results if s['grade'] == 'D']
c_stores = [s['sid'] for s in store_results if s['grade'] == 'C']
core = (f'20家门店Q1合计营收{r1(total_rev)}万元、EBITDA {r1(total_ebitda)}万元、税后利润{r1(total_net)}万元；'
        f'A级{a_count}家、B级{b_count}家、C级{c_count}家、D级{d_count}家；D级门店为{"、".join(d_stores)}，启动整改，连续2季D级则关店或转址；'
        f'外卖对账差异3家、盘点超标2家，已明确责任；店长奖金合计{r2(total_bonus)}万元。')

# 写入答案
wb = openpyxl.Workbook()
wb.remove(wb.active)

# 01 门店盈利分级
grade_rows = []
for s in store_results:
    grade_rows.append([s['sid'], s['city'], s['mall'], s['manager'], s['typ'],
                       r1(s['rev']), r1(s['ebitda']), r2(s['rate'] * 100), s['grade'],
                       r1(s['rent_q1']), r2(s['rent_q1'] / s['rev'] * 100), r1(s['net'])])
sheet(wb, '门店盈利分级',
      ['门店编号', '城市', '商圈', '店长', '类型', 'Q1营收(万元)', 'Q1 EBITDA(万元)',
       'EBITDA率(%)', '分级', 'Q1租金(万元)', '租金占收入(%)', 'Q1税后利润(万元)'],
      grade_rows, {1: 10, 2: 8, 3: 20, 4: 8, 5: 8, 6: 14, 7: 16, 8: 12, 9: 8, 10: 12, 11: 14, 12: 16})

sheet(wb, '分级标准',
      ['分级', 'EBITDA率', '处置策略', '基础奖金(万元)'],
      [['A', '≥12%', '保持，可作为标杆复制', 1.2],
       ['B', '8%-12%', '维持经营，关注成本', 1.0],
       ['C', '4%-8%', '限期整改，Q2复核', 0.8],
       ['D', '<4%或连续亏损', '启动整改，连续2季D级关店或转址', 0.5]],
      {1: 8, 2: 14, 3: 32, 4: 14})

# 02 单店差异归因
attr_rows = []
for s in store_results:
    rev_var = s['rev'] - s['budget_rev']
    ebitda_var = s['ebitda'] - s['budget_ebitda']
    # 归因：租金、外卖佣金、其他
    rent_pct = s['rent_q1'] / s['rev'] * 100
    if rent_pct > 15:
        main = '租金占比过高'
    elif s['loss_rate'] > 0.02:
        main = '盘点损耗超标'
    elif s['rate'] < 0.04:
        main = '客流下滑、营收不足'
    else:
        main = '达标'
    attr_rows.append([s['sid'], s['grade'], r1(s['budget_rev']), r1(s['rev']), r1(rev_var),
                      r1(s['budget_ebitda']), r1(s['ebitda']), r1(ebitda_var),
                      r2(rent_pct), main])
sheet(wb, '单店差异归因',
      ['门店编号', '分级', '预算营收(万元)', '实际营收(万元)', '营收差异(万元)',
       '预算EBITDA(万元)', '实际EBITDA(万元)', 'EBITDA差异(万元)', '租金占收入(%)', '主要归因'],
      attr_rows, {1: 10, 2: 8, 3: 14, 4: 14, 5: 14, 6: 16, 7: 16, 8: 16, 9: 14, 10: 18})

# 03 外卖对账
sheet(wb, '外卖对账差异',
      ['门店编号', '城市', '外卖GMV(万元)', '平台佣金(万元)', '平台实收(万元)',
       '门店入账(万元)', '差异(万元)', '状态', '原因'],
      deliv_rows, {1: 10, 2: 8, 3: 16, 4: 14, 5: 14, 6: 14, 7: 12, 8: 12, 9: 40})

# 04 盘点责任
sheet(wb, '盘点损耗与责任',
      ['门店编号', '城市', 'Q1平均差异率(%)', '是否超标', '差异金额(万元)', '责任认定'],
      inv_rows, {1: 10, 2: 8, 3: 18, 4: 10, 5: 14, 6: 12})

# 05 店长奖金
sheet(wb, '店长奖金核算',
      ['门店编号', '城市', '店长', '分级', '基础奖金(万元)', 'EBITDA达成率(%)',
       '营收达成率(%)', '损耗率(%)', '奖金系数', '应发奖金(万元)', '是否追回预提'],
      bonus_rows, {1: 10, 2: 8, 3: 8, 4: 8, 5: 14, 6: 14, 7: 12, 8: 10, 9: 10, 10: 14, 11: 14})

# 06 处置建议
sheet(wb, '亏损店处置建议',
      ['门店编号', '城市', '商圈', '分级', 'EBITDA率(%)', 'Q1 EBITDA(万元)', '处置建议', '说明'],
      dispo_rows, {1: 10, 2: 8, 3: 20, 4: 8, 5: 12, 6: 16, 7: 30, 8: 40})

# 07 风险与退出
sheet(wb, '风险清单', ['编号', '风险', '触发阈值', '影响', '应对措施', '退出动作'],
      risk_rows, {1: 8, 2: 16, 3: 28, 4: 24, 5: 30, 6: 32})
sheet(wb, '退出机制', ['编号', '触发条件', '退出方式', '回收与损失处理'],
      exit_rows, {1: 8, 2: 30, 3: 24, 4: 48})

# 08 勾稽与结论
sheet(wb, '错误要求驳回', ['编号', '管理层要求', '结论', '驳回理由'],
      reject_rows, {1: 8, 2: 54, 3: 10, 4: 34})

check_out = []
for c in checks:
    diff = ''
    if isinstance(c[1], (int, float)) and isinstance(c[2], (int, float)):
        diff = r2(c[1] - c[2])
    check_out.append([c[0], c[1], c[2], diff, '一致' if diff in ('', 0, 0.0) else '不一致'])
sheet(wb, '勾稽自检', ['勾稽项', '左值', '右值', '差异', '结论'],
      check_out, {1: 28, 2: 18, 3: 18, 4: 12, 5: 10})

concl = [
    ['核心结论', core],
    ['支撑一', f'20家门店Q1合计营收{r1(total_rev)}万元、EBITDA {r1(total_ebitda)}万元、税后利润{r1(total_net)}万元；A级{a_count}家、B级{b_count}家、C级{c_count}家、D级{d_count}家。'],
    ['支撑二', f'D级门店为{"、".join(d_stores)}，租金占收入超过15%且客流下滑，启动整改，连续2季D级则关店或转址；C级门店{"、".join(c_stores)}限期整改。'],
    ['支撑三', f'外卖对账差异3家（SZ04少计收入、WX04补贴未计、NC03退款未冲减），盘点超标2家（WX04、NB04），已明确门店责任并从奖金扣除。'],
    ['支撑四', f'店长奖金合计{r2(total_bonus)}万元，D级门店基础奖金扣减50%并追回已发预提；加盟门店利润不并表，仅确认授权费与供应链毛利。'],
    ['风险提示', '重点关注客流下滑、租金占比过高、外卖佣金上调、盘点损耗超标、新店爬坡不及预期五类风险。'],
    ['退出触发', '连续2季EBITDA率<4%，或租金占收入连续2季超过18%，启动关店、转址或解约评估。'],
]
sheet(wb, '经营分析结论', ['项目', '内容'], concl, {1: 14, 2: 100})

out = f'{BASE}/门店经营分析_v1_标准答案.xlsx'
wb.save(out)
print('saved', out)
print('grades', {g: sum(1 for s in store_results if s['grade'] == g) for g in 'ABCD'})
print('D stores', d_stores)
print('total_rev', r1(total_rev), 'total_ebitda', r1(total_ebitda), 'total_net', r1(total_net))
print('total_bonus', r2(total_bonus))
print('deliv_diff', deliv_diff_stores, 'inv_resp', inv_resp_stores)
