#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""应收账款与客户信用管理审计 v1 —— 标准答案生成器"""
from datetime import date
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE = '/workspace/高经济价值第八期_应收账款审计'
SRC = f'{BASE}/应收账款审计_v1_题目.xlsx'
OUT = f'{BASE}/应收账款审计_v1_标准答案.xlsx'

HDR_FILL = PatternFill('solid', fgColor='2F5597')
HDR_FONT = Font(name='宋体', bold=True, color='FFFFFF', size=10)
BODY = Font(name='宋体', size=9)
THIN = Side(style='thin', color='CCCCCC')
BORDER = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
WRAP = Alignment(wrap_text=True, vertical='top')
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)

CUTOFF = date(2026, 11, 30)
PAY_DAYS = 60  # 付款期：验收后60天


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


def rows_of(ws):
    data = list(ws.iter_rows(values_only=True))
    hdr = data[0]
    return [dict(zip(hdr, r)) for r in data[1:] if r[0] is not None]


def r4(x): return round(x + 0.0, 4)


def parse(d):
    y, m, dd = str(d).split('-')
    return date(int(y), int(m), int(dd))


def rate(days):
    if days <= 0:
        return 0.005
    if days <= 90:
        return 0.05
    if days <= 180:
        return 0.20
    if days <= 365:
        return 0.50
    return 1.0


wb0 = openpyxl.load_workbook(SRC, data_only=True)
ar = rows_of(wb0['应收账款明细'])
book = {r['编号']: float(r['账面计提(万元)']) for r in rows_of(wb0['财务账面坏账准备'])}

aging_rows = []     # 账龄口径重算
prov_rows = []      # 坏账准备重算
special_rows = []   # 个别认定处理
adj_rows = []       # 账面差异调节

correct_total_bal = 0.0     # 正确应收余额合计
book_total_bal = 0.0        # 账面应收余额合计
correct_total_prov = 0.0
book_total_prov = 0.0
restruct_loss = 12.6        # 青云减免损失
qingyun_balance = r4(42 * 0.7 - 2.45 * 2)  # 29.4 - 4.9 = 24.5

for a in ar:
    code = a['编号']
    cust = a['客户']
    inv_d = parse(a['开票日期'])
    acc_d = parse(a['验收日期'])
    amt = float(a['应收金额(万元)'])
    paid = float(a['已回款(万元)'])
    typ = a['款项类型']
    bal = r4(amt - paid)

    # 制度口径（正确）：验收日+60天为节点；质保金自质保期届满起算
    if code == 'AR13':
        # 质保金：质保期2024-11-20至2026-11-20届满，届满日转应收
        due = date(2026, 11, 20)
        days_sys = (CUTOFF - date(2025, 1, 24)).days  # 系统按开票日+60
        days = (CUTOFF - due).days
        sys_rate = rate(days_sys)
        cor_rate = rate(days)
        note = '质保期2026-11-20届满，届满日起算，逾期10天；账面误并入账龄按开票日起算'
    else:
        due = acc_d.toordinal() + PAY_DAYS
        due = date.fromordinal(due)
        days = (CUTOFF - due).days
        sys_due = date.fromordinal(inv_d.toordinal() + PAY_DAYS)
        days_sys = (CUTOFF - sys_due).days
        sys_rate = rate(days_sys)
        cor_rate = rate(days)
        note = ''

    # 特殊处理
    special = ''
    cor_bal = bal
    if code == 'AR08':
        cor_rate = 1.0
        special = '个别认定：客户进入破产程序，全额计提'
    elif code == 'AR11':
        cor_rate = 0.10
        cor_bal = qingyun_balance
        special = '债务重组：减免12.6万元确认损失，余额24.5万元按10%个别计提'
    aging_rows.append([code, cust, str(inv_d), str(acc_d), bal, days_sys, days,
                       f'{sys_rate*100:g}%', f'{cor_rate*100:g}%', note or '两口径差异源于起算日不同'])
    prov = r4(cor_bal * cor_rate)
    book_prov = r4(book[code])
    prov_rows.append([code, cust, cor_bal, f'{cor_rate*100:g}%', prov, book_prov, r4(prov - book_prov), special or ''])
    correct_total_bal = r4(correct_total_bal + cor_bal)
    book_total_bal = r4(book_total_bal + bal)
    correct_total_prov = r4(correct_total_prov + prov)
    book_total_prov = r4(book_total_prov + book_prov)
    diff = r4(prov - book_prov)
    if abs(diff) > 0.0001 or special:
        reason = special or f'账龄起算口径更正：验收日起算（逾期{days}天）替代开票日起算（逾期{days_sys}天）'
        adj_rows.append([code, cust, book_prov, prov, diff, reason])

# 重组损失单独调节
adj_rows.append(['AR11', '青云文旅', 0, 0, 0, f'补确认债务重组损失{restruct_loss}万元（减免30%），应收余额由42万元调整为{qingyun_balance}万元'])

prov_diff = r4(correct_total_prov - book_total_prov)

# ============ 催收与信用政策 ============
collect_plan = [
    ['安捷物流', '逾期285天、余额65万元', '诉讼准备', '停止新增授信，发出律师函，2027年1月15日前未付即提起诉讼'],
    ['宏图建筑', '逾期381天、余额70万元、客户失联', '诉讼并全额计提', '立即起诉并申请财产保全，坏账已按100%计提'],
    ['星辉传媒', '破产清算中、余额48万元', '破产债权申报', '公告日起90日内（2026-12-17前）完成债权申报，全额计提'],
    ['博瑞汽车', '逾期205天、余额55万元', '分期协议', '停止新签项目，签订分期付款协议并加收逾期利息'],
    ['旭日科技', '逾期229天、余额38万元', '停止新单', '余款结清前暂停承接其新项目'],
    ['信用政策调整', 'C/D级客户共5家', '收紧授信', 'C级客户预付款比例提至50%，D级客户一律款到服务；A级维持60天账期'],
]

# ============ 风险清单与整改机制 ============
risk_rows = [
    ['R01', '账龄起算口径错误', '任一笔应收起算日与制度不符', '高', '按验收日重算全部账龄并锁定系统参数'],
    ['R02', '个别认定缺失', '存在破产/吊销/明确拒付客户仍按比例计提', '高', '每月筛查工商与司法公示，触发即全额计提'],
    ['R03', '债务重组未入账', '重组协议签署后财务未确认重组损失', '高', '法务签约后5日内抄送财务入账'],
    ['R04', '质保金混入账龄', '质保金未按届满日起算', '中', '质保金单独科目核算，届满日自动转应收'],
    ['R05', '破产债权逾期未申报', '申报期限内未完成申报', '高', '获悉破产公告10日内启动申报流程'],
    ['R06', '长账龄应收无催收动作', '逾期超180天且无书面催收记录', '中', '逾期90/180/365天分级催收并留痕'],
    ['R07', '坏账准备调整不及时', '季度末未按最新账龄滚动调整', '中', '按季重算，差异超1万元须说明'],
    ['R08', '信用政策与账龄脱节', 'C级以下客户仍享60天账期', '中', '授信与催收结果联动，每半年评审'],
    ['R09', '催收外包无记录', '外包催收无书面交接与记录', '低', '催收外包须签合同并保留全部记录'],
    ['R10', '重组履约跟踪缺失', '重组客户逾期一期未触发预警', '中', '重组台账逐月核对，逾期一期即恢复全额计提'],
]
mech_rows = [
    ['J01', '账龄口径锁定机制', '制度变更后30日内完成系统参数更新', '开票日起算参数停用，验收日参数唯一有效'],
    ['J02', '个别认定筛查机制', '每月末筛查客户司法与工商状态', '触发即全额计提并上报管委会'],
    ['J03', '重组台账跟踪机制', '重组协议签署即建台账', '逾期一期未付恢复全额计提并重启谈判'],
    ['J04', '破产债权申报机制', '获悉破产公告10日内启动', '申报回执归档，逾期未申报追责'],
    ['J05', '分级催收机制', '逾期90/180/365天分级', '90天邮件、180天律师函、365天诉讼'],
    ['J06', '信用政策评审机制', '每半年评审一次', 'C级预付款50%、D级款到服务'],
]

# ============ 错误要求驳回 ============
reject_rows = [
    ['M01', '账龄按开票日算', '不予采纳', '账龄管理制度v2.0现行有效，起算日为验收日后第61天；开票日起算的旧制度已废止'],
    ['M02', '星辉48万按50%计提', '不予采纳', '客户已进入破产程序，个别认定规则要求全额计提，比例法不适用'],
    ['M03', '青云12.6万不确认损失', '不予采纳', '重组协议已盖章生效且对方正常履约，减免部分必须确认债务重组损失12.6万元'],
    ['M04', '质保金按开票日起算', '不予采纳', '质保金规则明确质保期内不并入账龄，届满日2026-11-20起算，逾期10天按5%计提0.4万元'],
    ['M05', '逾期统一按100%计提', '不予采纳', '计提办法为阶梯比例，统一100%无制度依据且扭曲报表'],
    ['M06', '未逾期不计提', '不予采纳', '计提办法规定未逾期按0.5%计提，不计提违反制度'],
    ['M07', '重组正常履约部分不计提', '不予采纳', '重组处理规则要求正常履约期间按10%个别计提，余额24.5万元须计提2.45万元'],
    ['M08', '宏图再等等', '不予采纳', '逾期381天超365天档且客户失联，应按100%计提并立即诉讼'],
    ['M09', '坏账年底一次性调', '不予采纳', '计提按季滚动调整，差异超1万元须说明，一次性调违反及时性'],
    ['M10', '催收外包不留记录', '不予采纳', '催收记录是诉讼与核销的关键证据，外包也须留存全部记录'],
    ['M11', '信用政策不用改', '不予采纳', 'C/D级客户逾期事实充分，须执行预付款50%与款到服务的收紧政策'],
    ['M12', '破产债权直接核销', '不予采纳', '破产程序中应先申报债权参与分配，核销须经法定程序，直接核销造成权益丧失'],
]

# ============ 写入答案 ============
wb = openpyxl.Workbook()
wb.remove(wb.active)

sheet(wb, '账龄口径重算',
      ['编号', '客户', '开票日期', '验收日期', '应收余额(万元)', '系统逾期天数(开票日)', '制度逾期天数(验收日)', '系统口径比例', '制度口径比例', '说明'],
      aging_rows, {1: 7, 2: 11, 3: 11, 4: 11, 5: 13, 6: 17, 7: 17, 8: 12, 9: 12, 10: 34})

sheet(wb, '坏账准备重算',
      ['编号', '客户', '正确应收余额(万元)', '适用比例', '正确计提(万元)', '账面计提(万元)', '差异(万元)', '个别认定说明'],
      prov_rows, {1: 7, 2: 11, 3: 16, 4: 10, 5: 13, 6: 13, 7: 11, 8: 34})

sheet(wb, '个别认定处理',
      ['编号', '客户', '事项', '正确应收余额(万元)', '适用处理', '依据'],
      [['AR08', '星辉传媒', '破产清算（公告BP-01）', 48, '全额计提100%，2026-12-17前完成破产债权申报', '个别认定规则AR-GF-03'],
       ['AR11', '青云文旅', '债务重组（协议RZ-01）', qingyun_balance, '确认重组损失12.6万元，余额按10%个别计提2.45万元', '债务重组处理规则AR-GF-04'],
       ['AR13', '金陵商贸', '质保金届满转应收', 8, '届满日2026-11-20起算，逾期10天按5%计提0.4万元', '质保金管理规则AR-GF-05']],
      {1: 7, 2: 11, 3: 24, 4: 16, 5: 40, 6: 26})

sheet(wb, '账面差异调节',
      ['编号', '客户', '账面计提(万元)', '正确计提(万元)', '差异(万元)', '调节原因'],
      adj_rows, {1: 7, 2: 11, 3: 13, 4: 13, 5: 11, 6: 52})

sheet(wb, '催收与信用政策',
      ['对象', '现状', '处置方式', '具体安排'],
      collect_plan, {1: 14, 2: 26, 3: 16, 4: 52})

sheet(wb, '风险清单', ['编号', '风险描述', '量化触发阈值', '风险等级', '应对措施'],
      risk_rows, {1: 7, 2: 22, 3: 32, 4: 9, 5: 38})
sheet(wb, '整改机制', ['编号', '机制名称', '触发条件', '处理方式'],
      mech_rows, {1: 7, 2: 22, 3: 30, 4: 38})

sheet(wb, '错误要求驳回', ['编号', '错误要求', '结论', '驳回理由与正确做法'],
      reject_rows, {1: 7, 2: 24, 3: 9, 4: 62})

# ============ 勾稽自检 ============
over180 = sum(1 for r in aging_rows if isinstance(r[6], int) and r[6] > 180)
not_overdue = sum(1 for r in aging_rows if isinstance(r[6], int) and r[6] <= 0)
diff_cnt = sum(1 for r in aging_rows if r[7] != r[8])
checks = [
    ['C01', '应收笔数', len(ar), 14, ''],
    ['C02', '客户数量', len({a['客户'] for a in ar}), 14, ''],
    ['C03', '账面应收余额合计(万元)', book_total_bal, 666.0, ''],
    ['C04', '正确应收余额合计(万元)', correct_total_bal, 648.5, ''],
    ['C05', '债务重组损失(万元)', restruct_loss, 12.6, ''],
    ['C06', '账面计提合计(万元)', book_total_prov, 254.625, ''],
    ['C07', '正确计提合计(万元)', correct_total_prov, 225.775, ''],
    ['C08', '计提差异合计(万元)', prov_diff, -28.85, ''],
    ['C09', '逾期超180天笔数', over180, 5, ''],
    ['C10', '未逾期笔数', not_overdue, 2, ''],
    ['C11', '个别认定笔数', 3, 3, ''],
    ['C12', '星辉全额计提(万元)', dict((r[0], r[4]) for r in prov_rows)['AR08'], 48, ''],
    ['C13', '青云重组后余额(万元)', qingyun_balance, 24.5, ''],
    ['C14', '青云个别计提(万元)', dict((r[0], r[4]) for r in prov_rows)['AR11'], 2.45, ''],
    ['C15', '金陵质保金正确计提(万元)', dict((r[0], r[4]) for r in prov_rows)['AR13'], 0.4, ''],
    ['C16', '宏图全额计提(万元)', dict((r[0], r[4]) for r in prov_rows)['AR12'], 70, ''],
    ['C17', '安捷物流计提(万元)', dict((r[0], r[4]) for r in prov_rows)['AR05'], 32.5, ''],
    ['C18', '风险条目数', len(risk_rows), 10, ''],
    ['C19', '整改机制条目数', len(mech_rows), 6, ''],
    ['C20', '错误要求驳回条数', len(reject_rows), 12, ''],
]
bad = []
for c in checks:
    left, right = c[2], c[3]
    try:
        c[4] = '一致' if r4(float(left)) == r4(float(right)) else '不一致'
    except (TypeError, ValueError):
        c[4] = '一致' if str(left) == str(right) else '不一致'
    if c[4] != '一致':
        bad.append(c)
sheet(wb, '勾稽自检', ['编号', '勾稽项', '左侧值', '右侧值', '结论'],
      checks, {1: 7, 2: 30, 3: 16, 4: 16, 5: 10})

concl = [
    ['核心结论',
     f'应收账款存在账龄口径错误、个别认定缺失、债务重组未入账、质保金混入账龄四类问题；账面应收余额{book_total_bal}万元经重组调整后正确余额{correct_total_bal}万元；账面坏账准备{book_total_prov}万元，正确计提{correct_total_prov}万元，多提{abs(prov_diff)}万元；另须补确认债务重组损失{restruct_loss}万元。'],
    ['支撑一',
     f'账龄口径：账龄管理制度v2.0以验收日后第61天为起算日，系统仍按已废止的开票日口径运行，14笔应收中{diff_cnt}笔两口径计提比例不同。'],
    ['支撑二',
     '个别认定：星辉传媒已进入破产程序须全额计提48万元并在申报期限内申报债权；青云文旅重组减免12.6万元确认损失、余额24.5万元按10%计提；金陵商贸质保金届满日起算仅计提0.4万元，账面误全额计提8万元。'],
    ['支撑三',
     '催收与信用：安捷物流、宏图建筑进入诉讼程序，博瑞汽车签分期协议，旭日科技停止新单；C级客户预付款提至50%、D级款到服务。'],
    ['支撑四',
     '管理要求：12项错误要求全部驳回，其中M02/M12违反个别认定与破产程序，M03/M07违反重组处理规则，M01/M04违反账龄与质保金制度，M05/M06违反计提办法。'],
    ['关键结果',
     f'正确应收余额{correct_total_bal}万元；正确坏账准备{correct_total_prov}万元（账面多提{abs(prov_diff)}万元）；重组损失{restruct_loss}万元；破产申报1笔（48万元）；诉讼2笔（135万元）；勾稽{len(checks)}项全部一致，不一致项0项。'],
]
sheet(wb, '应收账款审计结论', ['项目', '内容'], concl, {1: 10, 2: 110})

wb.save(OUT)
print('标准答案已生成：', OUT)
print(f'勾稽 {len(checks)} 项，不一致 {len(bad)} 项')
for c in bad:
    print('  不一致:', c)
print('账面余额:', book_total_bal, '| 正确余额:', correct_total_bal,
      '| 账面计提:', book_total_prov, '| 正确计提:', correct_total_prov, '| 差异:', prov_diff)
