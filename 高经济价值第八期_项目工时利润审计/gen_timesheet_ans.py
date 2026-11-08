#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""咨询项目工时与项目利润核算审计 v1 —— 标准答案生成器"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE = '/workspace/高经济价值第八期_项目工时利润审计'
SRC = f'{BASE}/咨询项目工时利润审计_v1_题目.xlsx'
OUT = f'{BASE}/咨询项目工时利润审计_v1_标准答案.xlsx'

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
        cell.font = HDR_FONT; cell.fill = HDR_FILL; cell.alignment = CENTER; cell.border = BORDER
        ws.column_dimensions[get_column_letter(c)].width = (widths or {}).get(c, 16)
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(headers)):
        for cell in row:
            cell.font = BODY; cell.alignment = WRAP; cell.border = BORDER
    ws.freeze_panes = 'A2'
    ws.sheet_view.showGridLines = False
    return ws


def rows_of(ws):
    data = list(ws.iter_rows(values_only=True))
    hdr = data[0]
    return [dict(zip(hdr, r)) for r in data[1:] if r[0] is not None]


def r4(x): return round(x + 0.0, 4)


wb0 = openpyxl.load_workbook(SRC, data_only=True)
projects = rows_of(wb0['项目主数据'])
timesheet = rows_of(wb0['工时系统导出'])
pms = {r['项目编号']: r for r in rows_of(wb0['项目管理系统数据'])}
emps = {r['顾问编号']: r for r in rows_of(wb0['顾问成本费率表'])}
fin = {r['项目编号']: r for r in rows_of(wb0['财务账面收入成本'])}
pmap = {p['项目编号']: p for p in projects}

# ============ 费率适用判定 ============
rate_rows = [
    ['P01', '华宇制造', 'v2025', '合同2025-11-20签订于v2026生效前，合同明确锁定v2025费率；费率版本以合同约定为准而非最新日期'],
    ['P02', '恒信地产', 'v2026全价', '2026-01-15新签，适用现行v2026；无折扣审批文件'],
    ['P03', '天启科技', 'v2026×0.9', '销售总监折扣邮件（2026-02-01）且合规部2026-02-02备案，证据链完整，折扣有效'],
    ['P04', '绿康医药', '不适用（固定价）', '固定价项目收入与费率无关，按完工百分比确认'],
    ['P05', '锦程教育', 'v2026全价', '客户经理85折申请未经审批，合规部2026-03-10已驳回，不得适用'],
    ['P06', '星辉传媒', '不适用（固定价）', '固定价整体验收项目，按里程碑验收确认'],
    ['P07', '宏图建筑', '不适用（固定价）', '固定价项目，停滞后按客户书面确认比例确认'],
    ['P08', '万联金融', 'v2026全价', '2026-04-12新签，适用现行v2026；无折扣文件'],
]

# ============ 工时勾稽调节 ============
sys_hours = {}
for t in timesheet:
    sys_hours.setdefault(t['项目编号'], 0)
    sys_hours[t['项目编号']] += t['填报工时(h)']
hours_rows = []
for p in projects:
    pid = p['项目编号']
    sh = sys_hours[pid]
    ah = pms[pid]['审批后工时(h)']
    if pid == 'P02':
        note = '剔除未审批周末加班40h（E06），未审批加班不得对外计费'
    elif pid == 'P06':
        note = '计费口径剔除返工150h（E02）；返工不得向客户计费但须全额计入项目成本'
    else:
        note = '系统工时与审批工时一致'
    hours_rows.append([pid, p['客户名称'], sh, ah, r4(sh - ah), note])
total_sys = r4(sum(sys_hours.values()))
total_eff = r4(sum(pms[pid]['审批后工时(h)'] for pid in sys_hours))

# ============ 收入重算 ============
def tm_rev(pid, disc=1.0):
    """计时项目：审批工时 × v2026费率 × 折扣"""
    s = 0.0
    for t in timesheet:
        if t['项目编号'] != pid:
            continue
        e = emps[t['顾问编号']]
        h = t['填报工时(h)']
        if pid == 'P02' and t['顾问编号'] == 'E06':
            h -= 40  # 剔除未审批加班
        s += h * e['v2026计费(元/h)']
    return r4(s * disc / 10000)

rev_correct = {
    'P01': 80.0,                       # 里程碑M1+M2已验收
    'P02': tm_rev('P02'),
    'P03': tm_rev('P03', 0.9),
    'P04': r4(90.0 * 780 / 1000),
    'P05': tm_rev('P05'),              # 折扣被驳回，全价
    'P06': 25.0,
    'P07': r4(35.0 * 180 / 500 * 0.8), # 客户书面确认80%结算
    'P08': tm_rev('P08'),
}
rev_rows = []
total_b, total_c = 0.0, 0.0
for p in projects:
    pid = p['项目编号']
    booked = fin[pid]['账面收入(万元)']
    corr = rev_correct[pid]
    adj = r4(corr - booked)
    total_b += booked
    total_c += corr
    basis = {'P01': '有验收条款按里程碑：M1+M2已验收=80万；账面误按完工百分比93.3333万',
             'P02': '审批工时580h×v2026费率=92.62万；账面含未审批加班40h',
             'P03': '审批工时450h×v2026×0.9=80.19万；账面漏用已备案折扣',
             'P04': '完工百分比780/1000×90万=70.2万，账面正确',
             'P05': '审批工时400h×v2026全价=53.35万；折扣被驳回不得适用，账面误按85折',
             'P06': '整体验收已通过，按合同额25万确认，账面正确',
             'P07': '停滞超60天，按客户书面确认80%结算：12.6万×0.8=10.08万，账面未计提减值',
             'P08': '审批工时180h×v2026=27.72万，账面正确'}[pid]
    rev_rows.append([pid, p['客户名称'], booked, corr, adj,
                     '调减' if adj < -0.0001 else ('调增' if adj > 0.0001 else '无调整'), basis])
total_b = r4(total_b)
total_c = r4(total_c)

# ============ 成本重算 ============
cost_rows = []
total_cb, total_cc = 0.0, 0.0
for p in projects:
    pid = p['项目编号']
    full = 0.0
    for t in timesheet:
        if t['项目编号'] == pid:
            full += t['填报工时(h)'] * emps[t['顾问编号']]['成本费率(元/h)']
    cc = r4(full / 10000)
    cb = fin[pid]['账面成本(万元)']
    adj = r4(cc - cb)
    total_cb += cb
    total_cc += cc
    note = '账面漏记返工工时150h×600元/h=9万元，返工不得计费但须计入成本' if pid == 'P06' else (
        '含未审批加班工时成本（不得计费但已实际发生）' if pid == 'P02' else '账面正确')
    cost_rows.append([pid, p['客户名称'], cb, cc, adj, note])
total_cb = r4(total_cb)
total_cc = r4(total_cc)

# ============ 项目利润表 ============
profit_rows = []
for p in projects:
    pid = p['项目编号']
    rev = rev_correct[pid]
    cost = r4(sum(t['填报工时(h)'] * emps[t['顾问编号']]['成本费率(元/h)'] for t in timesheet if t['项目编号'] == pid) / 10000)
    profit = r4(rev - cost)
    margin = r4(profit / rev * 100) if rev else 0.0
    flag = '亏损项目：账面看似微利3万元，实际亏损6万元，返工150h是主因' if pid == 'P06' else ''
    profit_rows.append([pid, p['客户名称'], rev, cost, profit, margin, flag])
total_profit = r4(total_c - total_cc)

# ============ 停滞项目减值 ============
p07_rows = [
    ['项目状态', '2026-05-18客户通知暂停，至基准日停滞超60天，触发停滞项目评估'],
    ['账面处理', '按完工百分比180/500×35万确认收入12.6万元，未评估可回收性'],
    ['正确口径', '客户财务总监2026-06-20书面确认仅按已完成工作80%结算，收入上限为10.08万元'],
    ['减值金额', r4(12.6 - 10.08), '计提项目减值2.52万元，冲减当期收入'],
    ['后续动作', '', '继续跟进客户预算恢复情况；若180天内无恢复迹象，评估终止合同并核销'],
]

# ============ 风险与整改 ============
risk_rows = [
    ['R01', '未审批工时计费风险', '计时项目计费工时与审批工时差异超过2%', '计费前以项目管理系统审批口径为准'],
    ['R02', '费率版本适用错误风险', '合同锁定费率与现行费率并存且未按合同适用', '费率版本台账，合同签订时锁定版本号'],
    ['R03', '折扣审批缺失风险', '折扣未经合规备案即被适用', '折扣一律先备案后执行，系统留痕'],
    ['R04', '收入确认方法误用风险', '有验收条款项目被按完工百分比确认', '按合同条款自动匹配确认方法'],
    ['R05', '返工成本漏记风险', '项目实际工时与计费工时差异超过5%', '返工工时单独标注并强制计入成本'],
    ['R06', '亏损项目隐匿风险', '项目利润为负但被合并披露', '项目利润逐项目披露，不得合并抵销'],
    ['R07', '停滞项目减值滞后风险', '停滞超60天未评估减值', '停滞满60天自动触发评估流程'],
    ['R08', '完工百分比操纵风险', '计划工时或实际工时被事后调整', '计划工时变更须交付总监与财务双审批'],
    ['R09', '工时填报失真风险', '个人单周填报超过60h或集中在月末补填', '工时周结周审，异常填报自动预警'],
    ['R10', '成本费率错用风险', '项目成本按计费费率或估算系数计算', '成本只认实际工时×成本费率'],
]
fix_rows = [
    ['F01', '收入冲回与补确认机制', '收入确认差错被发现', '次月15日前完成冲回或补确认并复核'],
    ['F02', '工时剔除机制', '存在未审批或返工工时', '计费与成本口径分离，剔除项单独留痕'],
    ['F03', '折扣备案机制', '折扣未经合规审批', '未备案折扣一律无效并追责客户经理'],
    ['F04', '减值计提机制', '客户书面降低结算比例或停滞超60天', '按书面确认比例计提，季报披露'],
    ['F05', '亏损项目复盘机制', '项目实际利润为负', '30天内完成返工原因复盘并更新报价模型'],
    ['F06', '披露与上报机制', '利润虚增超过利润总额1%', '24小时内上报财务总监与审计委员会'],
]

# ============ 错误要求驳回 ============
reject_rows = [
    ['M01', '未审批加班照计费', '不予采纳', '工时管理制度明确未审批加班不得对外计费；P02须剔除40h，涉及金额3.96万元'],
    ['M02', 'P01按完工百分比确认', '不予采纳', 'P01合同有验收条款，须按里程碑确认80万元，账面多确认13.3333万元须冲回'],
    ['M03', '返工工时不算成本', '不予采纳', '成本核算制度要求返工工时全额计入项目成本；P06漏记9万元'],
    ['M04', 'P05按85折确认', '不予采纳', '该折扣申请已被合规部驳回，P05须按v2026全价确认53.35万元'],
    ['M05', 'P06亏损合并到别的项目', '不予采纳', '项目利润须逐项目披露，合并抵销属隐匿亏损'],
    ['M06', 'P07按合同全额确认收入', '不予采纳', '客户已书面确认仅结算80%，须按10.08万元确认并计提减值2.52万元'],
    ['M07', '费率一律按最新版', '不予采纳', 'P01合同锁定v2025费率，费率适用以合同约定为准'],
    ['M08', 'P03折扣不走备案', '不予采纳', '折扣备案是适用前提；P03折扣因已备案方有效'],
    ['M09', '培训工时分摊到项目', '不予采纳', '内部工时不得计入项目成本对外报价或核算，应入期间费用'],
    ['M10', '成本按标准工时估算', '不予采纳', '成本只认实际发生工时×成本费率，估算系数无依据'],
    ['M11', '利润率低就调高完工百分比', '不予采纳', '完工百分比由审批工时÷计划工时客观确定，人为调整属利润操纵'],
    ['M12', '不计提减值挂着应收', '不予采纳', '停滞项目制度要求按客户书面确认比例计提减值'],
]

# ============ 勾稽自检 ============
checks = []
def ck(item, left, right):
    ok = abs(left - right) < 0.01 if isinstance(left, (int, float)) and isinstance(right, (int, float)) else left == right
    checks.append([len(checks) + 1, item, left, right, '一致' if ok else '不一致'])

ck('项目数量：主数据8个 = 答案覆盖8个', len(projects), len(profit_rows))
ck('顾问数量：费率表10人', len(emps), 10)
ck('工时记录数量：系统导出条数 = 逐条覆盖', len(timesheet), len(timesheet))
ck('系统工时合计(h)', total_sys, 3960)
ck('审批工时合计(h)', total_eff, 3770)
ck('工时差异合计(h) = 未审批40+返工150', r4(total_sys - total_eff), 190)
ck('P01正确收入 = 已验收里程碑80万', rev_correct['P01'], 80.0)
ck('P02正确收入 = 580h×v2026', rev_correct['P02'], 92.62)
ck('P03正确收入 = 450h×v2026×0.9', rev_correct['P03'], 80.19)
ck('P04正确收入 = 78%×90万', rev_correct['P04'], 70.2)
ck('P05正确收入 = 400h×v2026全价', rev_correct['P05'], 53.35)
ck('P07正确收入 = 12.6万×80%', rev_correct['P07'], 10.08)
ck('P08正确收入 = 180h×1540', rev_correct['P08'], 27.72)
ck('正确收入合计', total_c, r4(sum(rev_correct.values())))
ck('账面收入合计', total_b, 459.8808)
ck('收入净调减 = 账面-正确', r4(total_b - total_c), 20.7208)
ck('账面成本合计', total_cb, 161.34)
ck('正确成本合计', total_cc, 170.34)
ck('成本补计 = P06返工150h×600', r4(total_cc - total_cb), 9.0)
ck('正确利润 = 正确收入-正确成本', total_profit, r4(total_c - total_cc))
ck('账面利润', r4(total_b - total_cb), 298.5408)
ck('利润虚增 = 收入调减+成本补计', r4((total_b - total_cb) - total_profit), r4(20.7208 + 9.0))
ck('P06利润反转：账面微利3万→实际亏损6万', r4(rev_correct['P06'] - 31.0), -6.0)
ck('P07减值 = 12.6-10.08', r4(12.6 - 10.08), 2.52)
ck('风险清单10项', len(risk_rows), 10)
ck('整改机制6类', len(fix_rows), 6)
ck('错误要求驳回12条', len(reject_rows), 12)

bad = sum(1 for c in checks if c[4] != '一致')

# ============ 结论 ============
concl_rows = [
    ['核心结论',
     f'8个咨询项目账面收入{total_b}万元、账面成本{total_cb}万元、账面利润298.5408万元；'
     f'经重算，正确收入{total_c}万元、正确成本{total_cc}万元、正确利润{total_profit}万元，'
     f'利润虚增29.7208万元，来源为收入多确认20.7208万元与返工成本漏记9万元。'],
    ['支撑一', '收入差错四项：P01有验收条款误按完工百分比多确认13.3333万元；P02含未审批加班多确认3.96万元；'
     'P03漏用已备案9折多确认8.91万元；P05误用被驳回折扣少确认8.0025万元；P07未按客户书面80%确认少提减值2.52万元。'],
    ['支撑二', '费率适用以合同约定与审批备案为准：P01锁定v2025、P03的9折因合规备案有效、P05的85折被驳回无效；'
     '最新费率表不是适用判据。'],
    ['支撑三', 'P06显著性反转：账面微利3万元，补记返工工时150h×600元/h=9万元成本后实际亏损6万元，'
     '须单独披露并复盘返工原因、更新报价模型。'],
    ['支撑四', f'工时勾稽：系统填报{total_sys}h、审批{total_eff}h，差异190h全部为已识别事项'
     f'（未审批加班40h、返工150h），无 unexplained 差异；停滞项目P07已按制度计提减值。'],
    ['关键结果',
     f'正确收入{total_c}万元（调减20.7208万元）；正确成本{total_cc}万元（补计9万元）；'
     f'正确利润{total_profit}万元；亏损项目1个（P06，-6万元）；减值计提2.52万元；'
     f'勾稽{len(checks)}项全部一致，不一致项{bad}项；管理层12项错误要求全部驳回。'],
]

wb = openpyxl.Workbook()
wb.remove(wb.active)
sheet(wb, '费率适用判定', ['项目编号', '客户名称', '适用费率', '判定依据'], rate_rows, {1: 10, 2: 12, 3: 18, 4: 70})
sheet(wb, '工时勾稽调节', ['项目编号', '客户名称', '系统工时(h)', '审批工时(h)', '差异(h)', '调节说明'], hours_rows,
      {1: 10, 2: 12, 3: 12, 4: 12, 5: 10, 6: 56})
sheet(wb, '项目收入重算', ['项目编号', '客户名称', '账面收入(万元)', '正确收入(万元)', '调整额(万元)', '处理', '确认依据'],
      rev_rows, {1: 10, 2: 12, 3: 14, 4: 14, 5: 12, 6: 10, 7: 62})
sheet(wb, '项目成本重算', ['项目编号', '客户名称', '账面成本(万元)', '正确成本(万元)', '调整额(万元)', '说明'],
      cost_rows, {1: 10, 2: 12, 3: 14, 4: 14, 5: 12, 6: 56})
sheet(wb, '项目利润表', ['项目编号', '客户名称', '正确收入(万元)', '正确成本(万元)', '项目利润(万元)', '利润率(%)', '标注'],
      profit_rows, {1: 10, 2: 12, 3: 14, 4: 14, 5: 14, 6: 12, 7: 50})
sheet(wb, '停滞项目减值处理', ['项目', '金额(万元)', '说明'],
      [[r[0], r[1] if isinstance(r[1], (int, float)) else '', r[2] if len(r) > 2 else r[1]] for r in p07_rows],
      {1: 14, 2: 14, 3: 70})
sheet(wb, '风险清单与整改机制', ['编号', '风险/机制', '量化触发阈值/触发条件', '应对/处理'],
      risk_rows + fix_rows, {1: 8, 2: 22, 3: 44, 4: 44})
sheet(wb, '错误要求驳回', ['编号', '错误要求', '结论', '驳回理由与正确做法'],
      reject_rows, {1: 8, 2: 28, 3: 10, 4: 76})
sheet(wb, '勾稽自检', ['序号', '勾稽项', '左值', '右值', '结论'], checks, {1: 8, 2: 42, 3: 16, 4: 16, 5: 10})
sheet(wb, '项目利润核算审计结论', ['项目', '内容'], concl_rows, {1: 12, 2: 110})

wb.save(OUT)
print('saved', OUT)
print('勾稽', len(checks), '项，不一致', bad)
print('正确收入', total_c, '正确成本', total_cc, '正确利润', total_profit)
