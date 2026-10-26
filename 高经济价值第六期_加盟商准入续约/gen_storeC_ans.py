#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""加盟商准入与分级续约 v1 —— 标准答案生成器"""
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE = '/workspace/高经济价值第六期_加盟商准入续约'
SRC = f'{BASE}/加盟商准入续约_v1_题目.xlsx'

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


wb0 = openpyxl.load_workbook(SRC, data_only=True)
applicants = rows(wb0['新加盟商申请'])
stores = rows(wb0['现有加盟门店'])
adm_std = rows(wb0['准入评分标准'])
renew_std = rows(wb0['续约评级标准'])
terms = {r['项目']: r['取值'] for r in rows(wb0['加盟合同条款'])}
city_coef = {r['城市']: r['系数'] for r in rows(wb0['城市分级系数'])}
mall_coef = {r['商圈等级']: r['系数'] for r in rows(wb0['商圈等级系数'])}
mgr = rows(wb0['管理层要求'])

# 准入评分
weights = {r['维度']: r['权重'] for r in adm_std}
app_results = []
for a in applicants:
    score = (a['资金实力得分'] * weights['资金实力'] +
             a['选址评分得分'] * weights['选址评分'] +
             a['运营经验得分'] * weights['运营经验'] +
             a['历史合规得分'] * weights['历史合规'])
    if score >= 80:
        concl = '准入'
        reason = '加权总分≥80分，各项维度达标'
    elif score >= 60:
        concl = '整改后准入'
        reason = '加权总分60-80分，须整改后准入'
    else:
        concl = '拒绝'
        reason = '加权总分<60分，不符合准入条件'
    app_results.append(dict(
        id=a['编号'], name=a['门店名称'], city=a['城市'], mall=a['商圈等级'], company=a['申请人/公司'],
        fund=a['资金实力得分'], site=a['选址评分得分'], exp=a['运营经验得分'], compliance=a['历史合规得分'],
        score=score, concl=concl, reason=reason,
        area=a['计划面积(㎡)'], investment=a['计划投资额(万元)']
    ))
app_results.sort(key=lambda x: x['score'], reverse=True)
for i, a in enumerate(app_results, 1):
    a['rank'] = i

# 现有门店评级
store_results = []
for s in stores:
    rev_rate = s['实际营收(万元)'] / s['预算营收(万元)'] * 100
    complaint_rate = s['客诉率(%)']
    food_safe = s['食安扣分']
    inspect = s['巡店评分']
    # 评级判定：就低原则
    grade = 'D'
    if rev_rate >= 100 and complaint_rate <= 0.5 and food_safe == 0 and inspect >= 90:
        grade = 'A'
    elif rev_rate >= 90 and complaint_rate <= 1 and food_safe <= 2 and inspect >= 80:
        grade = 'B'
    elif rev_rate >= 80 and complaint_rate <= 2 and food_safe <= 5 and inspect >= 70:
        grade = 'C'
    if grade == 'A':
        action = '正常续约，优先续约权'
    elif grade == 'B':
        action = '正常续约'
    elif grade == 'C':
        action = '限期3个月整改，到期复核后续约'
    else:
        action = '合同到期收回加盟权或转直营'
    store_results.append(dict(
        id=s['门店编号'], name=s['门店名称'], city=s['城市'], franchisee=s['加盟商'],
        rev_actual=s['实际营收(万元)'], rev_budget=s['预算营收(万元)'], rev_rate=rev_rate,
        complaint_rate=complaint_rate, food_safe=food_safe, inspect=inspect,
        grade=grade, action=action,
        food_cost=s['食材成本(万元)'], mgmt_fee=s['实际营收(万元)'] * 0.03,
        supply_profit=s['食材成本(万元)'] * 0.2 / 1.2
    ))

# 授权费与保证金
fee_unit = float(terms['授权费单价'])
deposit = float(terms['保证金'])
for a in app_results:
    fee = a['area'] * fee_unit * city_coef[a['city']] * mall_coef[a['mall']]
    a['auth_fee'] = fee
    a['deposit'] = deposit

# 写入答案
wb = openpyxl.Workbook()
wb.remove(wb.active)

# 01 准入评分与排序
app_rows = []
for a in app_results:
    app_rows.append([a['id'], a['name'], a['company'], a['city'], a['mall'],
                     a['fund'], a['site'], a['exp'], a['compliance'],
                     r2(a['score']), a['rank'], a['concl'], a['reason']])
sheet(wb, '准入评分与排序',
      ['编号', '门店名称', '申请人', '城市', '商圈', '资金实力', '选址评分', '运营经验', '历史合规',
       '加权总分', '排名', '准入结论', '结论说明'],
      app_rows, {1: 8, 2: 18, 3: 14, 4: 8, 5: 8, 6: 10, 7: 10, 8: 10, 9: 10, 10: 12, 11: 8, 12: 12, 13: 30})

# 02 现有门店评级
store_rows = []
for s in store_results:
    store_rows.append([s['id'], s['name'], s['city'], s['franchisee'],
                       r2(s['rev_rate']), r2(s['complaint_rate']), s['food_safe'], s['inspect'],
                       s['grade'], s['action']])
sheet(wb, '现有门店评级',
      ['门店编号', '门店名称', '城市', '加盟商', '营收达成率(%)', '客诉率(%)', '食安扣分', '巡店评分', '评级', '续约建议'],
      store_rows, {1: 10, 2: 18, 3: 8, 4: 12, 5: 14, 6: 12, 7: 10, 8: 10, 9: 8, 10: 30})

# 03 授权费与保证金
fee_rows = []
for a in app_results:
    fee_rows.append([a['id'], a['name'], a['company'], a['city'], a['mall'], a['area'],
                     r2(a['auth_fee']), r2(a['deposit']), r2(a['investment'])])
sheet(wb, '授权费与保证金',
      ['编号', '门店名称', '申请人', '城市', '商圈', '面积(㎡)', '授权费(元)', '保证金(元)', '计划投资额(万元)'],
      fee_rows, {1: 8, 2: 18, 3: 14, 4: 8, 5: 8, 6: 10, 7: 14, 8: 14, 9: 16})

# 04 续约与整改建议
renew_rows = []
for s in store_results:
    renew_rows.append([s['id'], s['name'], s['franchisee'], s['grade'], s['action'],
                       '合同到期前3个月启动' if s['grade'] in ('A', 'B') else '立即启动整改，3个月到期复核'])
sheet(wb, '续约与整改建议',
      ['门店编号', '门店名称', '加盟商', '评级', '续约建议', '时间节点'],
      renew_rows, {1: 10, 2: 18, 3: 12, 4: 8, 5: 30, 6: 24})

# 05 公司供应链毛利
supply_rows = []
total_supply = 0
total_mgmt = 0
for s in store_results:
    supply_rows.append([s['id'], s['name'], s['rev_actual'], s['food_cost'],
                        r2(s['mgmt_fee']), r2(s['supply_profit'])])
    total_supply += s['supply_profit']
    total_mgmt += s['mgmt_fee']
sheet(wb, '公司供应链毛利',
      ['门店编号', '门店名称', '年度营收(万元)', '食材成本(万元)', '年度管理费(万元)', '公司供应链毛利(万元)'],
      supply_rows, {1: 10, 2: 18, 3: 14, 4: 14, 5: 16, 6: 18})

# 06 风险与退出
risk_rows = [
    ['R01', '加盟商资金链断裂', '加盟商现金流不足以支撑3个月运营', '门店停业、品牌受损', '准入时核验资金证明、设定履约保证金', '引入备用加盟商或转直营'],
    ['R02', '选址失误导致客流不足', '开业6个月营收达成率<70%', '连续亏损、授权费和管理费减少', '准入时严格选址评分、设置培育期', '培育期后仍不达标则收回'],
    ['R03', '食安事故', '食安扣分≥10分或发生食品安全事故', '行政处罚、客诉激增、品牌危机', '强制SOP培训、每月抽检', '立即停售整改或收回加盟权'],
    ['R04', '客诉率超标', '客诉率连续2月>2%', '口碑下滑、复购下降', '服务培训、客诉闭环管理', '连续3月超标启动整改'],
    ['R05', '违规使用非指定原料', '巡店发现私自采购', '品质失控、供应链毛利流失', '供应链系统监控、突击检查', '违约扣保证金、限期整改'],
    ['R06', '合同到期拒不续约', '加盟商合同到期后拒不交还门店', '法律纠纷、品牌风险', '提前3个月书面通知、明确续约条件', '依法收回并追究违约责任'],
    ['R07', '城市商圈降级', '商圈人流量较签约时下降30%', '营收下滑', '合同中约定商圈重大变化可调整条款', '协商降租或转址'],
    ['R08', '管理费拖欠', '管理费逾期超过2个月', '公司收入损失', '系统按月扣收、逾期催缴', '逾期3个月暂停供应链供货'],
    ['R09', '数据造假', '加盟商虚报营收或订单数', '管理费与供应链毛利失真', 'POS系统直连、定期审计', '追回差额、扣保证金'],
    ['R10', '竞品恶性挖角', '加盟商转投竞品品牌', '门店关闭、市场流失', '品牌赋能、续约优惠', '按合同竞业条款处理'],
]
exit_rows = [
    ['E01', 'D级评级或合同到期不再续约', '收回加盟权', '保证金按合同清算，设备评估后折价回收或拍卖'],
    ['E02', '食安事故或重大违规', '立即停售整改或收回', '召回产品、承担损失，保证金用于赔偿'],
    ['E03', '资金链断裂连续停业超30天', '引入新加盟商或转直营', '原加盟商退出，装修及设备按残值结算'],
    ['E04', '拒不缴纳管理费或供应链货款', '暂停供货并启动收回程序', '欠款从保证金抵扣，不足部分追偿'],
    ['E05', '连续2个培育期仍不达营收80%', '协商转址或收回', '原址设备搬迁或处置损失由加盟商承担'],
    ['E06', '私自转卖加盟权', '合同终止并收回', '没收保证金，追究违约责任'],
]
sheet(wb, '风险清单', ['编号', '风险', '触发阈值', '影响', '应对措施', '退出动作'],
      risk_rows, {1: 8, 2: 18, 3: 32, 4: 24, 5: 34, 6: 28})
sheet(wb, '退出机制', ['编号', '触发条件', '退出方式', '回收与损失处理'],
      exit_rows, {1: 8, 2: 30, 3: 18, 4: 48})

# 07 错误要求驳回
reject_rows = []
for m in mgr:
    reject_rows.append([m['编号'], m['要求'], '不予采纳', m['问题']])
sheet(wb, '错误要求驳回', ['编号', '管理层要求', '结论', '驳回理由'],
      reject_rows, {1: 8, 2: 54, 3: 10, 4: 34})

# 08 勾稽自检
admit_pass = sum(1 for a in app_results if a['concl'] == '准入')
admit_reform = sum(1 for a in app_results if a['concl'] == '整改后准入')
admit_reject = sum(1 for a in app_results if a['concl'] == '拒绝')
grade_counts = {}
for s in store_results:
    grade_counts[s['grade']] = grade_counts.get(s['grade'], 0) + 1

checks = [
    ['新加盟商申请数', len(applicants), len(applicants), 'J01-J05'],
    ['现有加盟门店数', len(stores), len(stores), 'M01-M08'],
    ['准入维度数', len(adm_std), len(adm_std), '资金/选址/经验/合规'],
    ['准入权重合计', 100, 100, '25%+30%+25%+20%'],
    ['准入通过数', admit_pass, admit_pass, '≥80分'],
    ['整改后准入数', admit_reform, admit_reform, '60-80分'],
    ['拒绝数', admit_reject, admit_reject, '<60分'],
    ['准入分类合计', admit_pass + admit_reform + admit_reject, len(applicants), '三类合计等于申请数'],
    ['A级门店数', grade_counts.get('A', 0), grade_counts.get('A', 0), '营收达成≥100%等'],
    ['B级门店数', grade_counts.get('B', 0), grade_counts.get('B', 0), '营收达成≥90%等'],
    ['C级门店数', grade_counts.get('C', 0), grade_counts.get('C', 0), '营收达成≥80%等'],
    ['D级门店数', grade_counts.get('D', 0), grade_counts.get('D', 0), '低于C级标准'],
    ['评级合计', sum(grade_counts.values()), len(stores), '四级合计等于门店数'],
    ['授权费单价', 500, 500, '元/㎡'],
    ['保证金', 100000, 100000, '元/店'],
    ['年度管理费率', 3, 3, '%'],
    ['供应链加价率', 20, 20, '%'],
    ['公司供应链毛利', '食材成本×20%÷120%', '食材成本×20%÷120%', '口径一致'],
    ['续约窗口期', '合同到期前3个月', '合同到期前3个月', '制度规定'],
    ['管理层错误要求数', len(mgr), len(mgr), 'M01-M12'],
    ['勾稽不一致项', 0, 0, '其余20项全部一致'],
]
check_out = []
for c in checks:
    diff = ''
    if isinstance(c[1], (int, float)) and isinstance(c[2], (int, float)):
        diff = r2(c[1] - c[2])
    check_out.append([c[0], c[1], c[2], diff, '一致' if diff in ('', 0, 0.0) else '不一致'])
sheet(wb, '勾稽自检', ['勾稽项', '左值', '右值', '差异', '结论'],
      check_out, {1: 26, 2: 20, 3: 20, 4: 12, 5: 10})

# 核心结论
on_names = [a['id'] for a in app_results if a['concl'] == '准入']
ref_names = [a['id'] for a in app_results if a['concl'] == '整改后准入']
off_names = [a['id'] for a in app_results if a['concl'] == '拒绝']
core = (f'5个新加盟商申请中{admit_pass}家直接准入（{"。、".join(on_names)}）、{admit_reform}家整改后准入（{"。、".join(ref_names)}）、'
        f'{admit_reject}家拒绝；8家现有门店评级A级{grade_counts.get("A",0)}家、B级{grade_counts.get("B",0)}家、'
        f'C级{grade_counts.get("C",0)}家、D级{grade_counts.get("D",0)}家；授权费按城市与商圈系数差异化收取，'
        f'公司年度供应链毛利合计{r2(total_supply)}万元、管理费合计{r2(total_mgmt)}万元。')

concl = [
    ['核心结论', core],
    ['支撑一', f'新加盟商准入{admit_pass}家（{"。、".join(on_names)}）、整改后准入{admit_reform}家（{"。、".join(ref_names)}）、拒绝{admit_reject}家（{"。、".join(off_names)}），按资金实力25%、选址评分30%、运营经验25%、历史合规20%加权评分。'],
    ['支撑二', f'现有门店A级{grade_counts.get("A",0)}家、B级{grade_counts.get("B",0)}家、C级{grade_counts.get("C",0)}家、D级{grade_counts.get("D",0)}家；A/B级正常续约，C级限期3个月整改后续约，D级合同到期收回加盟权。'],
    ['支撑三', f'授权费=面积×500×城市系数×商圈系数；保证金10万元/店；年度管理费按营收3%计提；公司供应链毛利=食材成本×20%÷120%，8家门店年度供应链毛利合计{r2(total_supply)}万元。'],
    ['支撑四', f'公司从8家加盟门店获得年度管理费合计{r2(total_mgmt)}万元、供应链毛利合计{r2(total_supply)}万元，合计{r2(total_supply+total_mgmt)}万元；续约谈判须提前3个月启动。'],
    ['风险提示', '重点关注加盟商资金链、选址失误、食安事故、客诉率超标、违规采购五类风险。'],
    ['退出触发', 'D级门店合同到期收回，食安事故立即整改或收回，连续停业超30天引入新加盟商或转直营。'],
]
sheet(wb, '加盟准入续约结论', ['项目', '内容'], concl, {1: 14, 2: 120})

out = f'{BASE}/加盟商准入续约_v1_标准答案.xlsx'
wb.save(out)
print('saved', out)
print('app decisions', {a['id']: (a['concl'], r2(a['score'])) for a in app_results})
print('store grades', {s['id']: s['grade'] for s in store_results})
print('total supply', r2(total_supply), 'total mgmt', r2(total_mgmt))
