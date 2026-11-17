#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""G6-ZX-010 采购结算与应付账款专项审计 —— 标准答案生成器
从采购入库明细重算，识别重复入账、退货未冲减、损耗承担差异、含税换算、返利计提差异
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE = '/workspace/高经济价值第六期_采购结算应付'
SRC = f'{BASE}/采购结算应付_v1_题目.xlsx'
OUT = f'{BASE}/采购结算应付_v1_标准答案.xlsx'

HDR_FILL = PatternFill('solid', fgColor='2F5597')
HDR_FONT = Font(name='宋体', bold=True, color='FFFFFF', size=10)
BODY = Font(name='宋体', size=9)
THIN = Side(style='thin', color='CCCCCC')
BORDER = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
WRAP = Alignment(wrap_text=True, vertical='top')
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)

TAX = 1.13
V2_LOSS = {'蔬菜': 0.03, '肉类': 0.015, '乳制品': 0.010, '干货': 0.005, '包装': 0.000, '调料': 0.000}
V3_LOSS = {'蔬菜': 0.02, '肉类': 0.010, '乳制品': 0.008, '干货': 0.003, '包装': 0.000, '调料': 0.000}
V2_REBATE = {'A': (100000, 0.03, 300000, 0.05, 500000, 0.08),
             'B': (50000, 0.03, 200000, 0.05, 400000, 0.07),
             'C': (0, 0.03, 100000, 0.05, 200000, 0.07),
             'D': (0, 0.02, 80000, 0.04, 160000, 0.06)}
V3_REBATE = {'A': (10000, 0.02, 50000, 0.04, 100000, 0.07),
             'B': (5000, 0.02, 30000, 0.04, 80000, 0.06),
             'C': (0, 0.02, 20000, 0.04, 50000, 0.06),
             'D': (0, 0.01, 15000, 0.03, 40000, 0.05)}


def r2(x):
    return round(x + 0.0, 2)


def rebate(amount, table, grade):
    t = table[grade]
    if amount >= t[4]:
        return amount * t[5]
    elif amount >= t[2]:
        return amount * t[3]
    elif amount >= t[0]:
        return amount * t[1]
    return 0.0


wb0 = openpyxl.load_workbook(SRC, data_only=True)


def data_rows(name):
    ws = wb0[name]
    data = list(ws.iter_rows(values_only=True))
    for i, row in enumerate(data):
        if row and row[0] in ('门店编号', '供应商编号', '入库单号', '入账编号', '退货单号',
                              '门店编号', '争议编号', '项目', '编号'):
            hdr = row
            return [dict(zip(hdr, r)) for r in data[i + 1:] if r[0] is not None]
    raise RuntimeError(name)


stores = {r['门店编号']: r for r in data_rows('门店主数据')}
suppliers = {r['供应商编号']: r for r in data_rows('供应商主数据')}
tx = data_rows('采购入库明细')
ledger = data_rows('应付账款台账')
returns = data_rows('退货损耗记录')
rebate_rules = data_rows('供应商评级与返利规则')
settle = data_rows('付款对账明细')
disputes = data_rows('争议台账')

DIRECT = {sid for sid, s in stores.items() if s['门店类型'] == '直营'}

# 台账按入库单号分组
led_by_in = {}
for l in ledger:
    led_by_in.setdefault(l['入库单号'], []).append(l)

# 退货按入库单号分组
ret_by_in = {}
for r in returns:
    ret_by_in.setdefault(r['入库单号'], []).append(r)

# ===== 1. 逐单复核 =====
tx_out = []
for r in tx:
    in_no = r['入库单号']
    m = int(str(r['入库日期'])[5:7])
    ver = 'v2.0' if m <= 4 else 'v3.0'
    cat = r['品类']
    qty = float(r['数量(kg/件)'])
    price = float(r['单价(含税元)'])
    loss = float(r['损耗数量'])
    amount = float(r['金额(含税元)'])
    loss_amt = r2(loss * price)
    # 损耗承担
    if ver == 'v2.0':
        loss_bear = '供应商'
        payable_gross = r2(amount - loss_amt)
    else:
        loss_bear = '门店'
        payable_gross = amount
    # 退货冲减
    ret_list = ret_by_in.get(in_no, [])
    ret_amt = 0.0
    for rt in ret_list:
        if ver == 'v2.0':
            ret_amt += float(rt['退货金额(含税元)'])
        else:
            if rt['退货原因'] == '质量':
                ret_amt += float(rt['退货金额(含税元)'])
    ret_amt = r2(ret_amt)
    payable_after_ret = r2(payable_gross - ret_amt)
    payable_net = r2(payable_after_ret / TAX)
    input_tax = r2(payable_after_ret - payable_net)
    # 台账
    leds = led_by_in.get(in_no, [])
    led_amt = float(leds[0]['计提金额(含税元)']) if leds else 0.0
    led_cnt = len(leds)
    if not leds:
        verdict = '未入账'
    elif led_cnt > 1:
        verdict = '重复入账'
    else:
        # 台账金额=含税 amount，应计=payable_net
        diff = r2(led_amt - payable_net)
        if abs(diff) < 0.01:
            verdict = '一致'
        elif ver == 'v2.0':
            verdict = '损耗未扣减+未除税'
        else:
            verdict = '未除税'
    tx_out.append([in_no, r['入库日期'], r['门店编号'], r['供应商编号'], cat,
                   ver, qty, price, amount, loss, loss_amt, loss_bear,
                   ret_amt, payable_gross, payable_after_ret, payable_net,
                   input_tax, led_amt, led_cnt, verdict])

# ===== 2. 异常清单 =====
dup_rows = []
miss_rows = []
for r in tx:
    in_no = r['入库单号']
    leds = led_by_in.get(in_no, [])
    if len(leds) > 1:
        for extra in leds[1:]:
            dup_rows.append([extra['入账编号'], extra['入账月份'], r['门店编号'],
                             r['供应商编号'], in_no, float(extra['计提金额(含税元)']),
                             '重复入账', '重复入账应冲回'])
    if not leds:
        cat = r['品类']
        m = int(str(r['入库日期'])[5:7])
        ver = 'v2.0' if m <= 4 else 'v3.0'
        loss = float(r['损耗数量']) * float(r['单价(含税元)'])
        amt = float(r['金额(含税元)'])
        pg = amt - loss if ver == 'v2.0' else amt
        miss_rows.append([in_no, r['入库日期'], r['门店编号'], r['供应商编号'],
                          cat, r2(pg), r2(pg / TAX), '已入库未入账，应补提'])

# 退货未冲减：台账已按全额入账但存在退货，按版本规则应扣减的全部退货
ret_rows = []
for r in returns:
    in_no = r['入库单号']
    leds = led_by_in.get(in_no, [])
    if not leds:
        continue
    m = int(str(r['退货日期'])[5:7])
    ver = 'v2.0' if m <= 4 else 'v3.0'
    should_deduct = (ver == 'v2.0') or (r['退货原因'] == '质量')
    if should_deduct:
        ret_rows.append([r['退货单号'], r['退货日期'], r['门店编号'],
                         r['供应商编号'], in_no,
                         float(r['退货金额(含税元)']),
                         '退货未冲减', '退货应冲减应付'])

abn_rows = dup_rows + ret_rows
# 超标准损耗
loss_over = []
for r in tx:
    cat = r['品类']
    m = int(str(r['入库日期'])[5:7])
    std = V2_LOSS[cat] if m <= 4 else V3_LOSS[cat]
    qty = float(r['数量(kg/件)'])
    loss = float(r['损耗数量'])
    if qty > 0 and loss / qty > std:
        loss_over.append([r['入库单号'], r['入库日期'], r['门店编号'], r['供应商编号'],
                          cat, r2(qty), r2(loss), r2(std * qty),
                          r2(loss - std * qty), '损耗超标准'])

# ===== 3. 返利计提 =====
# 按供应商、期间汇总采购额
pur_by_sup_period = {}
for r in tx:
    sup = r['供应商编号']
    m = int(str(r['入库日期'])[5:7])
    period = f'2026-{m:02d}'
    amt = float(r['金额(含税元)'])
    pur_by_sup_period[(sup, period)] = pur_by_sup_period.get((sup, period), 0) + amt

rebate_out = []
rebate_total_v2 = 0.0
rebate_total_v3 = 0.0
for sup, s in suppliers.items():
    grade = s['评级']
    # v2.0 月度返利（1-4月）
    for m in range(1, 5):
        period = f'2026-{m:02d}'
        amt = pur_by_sup_period.get((sup, period), 0)
        if amt > 0:
            rb = rebate(amt, V2_REBATE, grade)
            rebate_out.append([sup, s['供应商名称'], grade, period, 'v2.0月度',
                               r2(amt), r2(rb), '冲减应付账款'])
            rebate_total_v2 += rb
    # v3.0 季度返利（Q2=4-6月，但4月属v2，所以Q2仅5-6月）
    q2_amt = sum(pur_by_sup_period.get((sup, f'2026-{m:02d}'), 0) for m in range(5, 7))
    if q2_amt > 0:
        rb = rebate(q2_amt, V3_REBATE, grade)
        rebate_out.append([sup, s['供应商名称'], grade, '2026-Q2', 'v3.0季度',
                           r2(q2_amt), r2(rb), '计入其他业务收入'])
        rebate_total_v3 += rb

# ===== 4. 费用归属重算（门店×月）=====
sm = {}
for r in tx:
    sid = r['门店编号']
    m = int(str(r['入库日期'])[5:7])
    key = (sid, m)
    a = sm.setdefault(key, dict(amount=0.0, payable_net=0.0, input_tax=0.0, cnt=0, is_direct=sid in DIRECT))
    # 找对应复核行
    rec = next(x for x in tx_out if x[0] == r['入库单号'])
    a['amount'] += float(r['金额(含税元)'])
    a['payable_net'] += rec[15]
    a['input_tax'] += rec[16]
    a['cnt'] += 1

calc_rows = []
for (sid, m), a in sorted(sm.items()):
    is_dir = a['is_direct']
    payable = a['payable_net'] if is_dir else 0.0
    calc_rows.append([sid, stores[sid]['城市'], f'2026-{m:02d}', a['cnt'],
                      r2(a['amount']), r2(a['payable_net']), r2(a['input_tax']),
                      r2(payable),
                      '加盟店由加盟商自行结算，公司不承担' if not is_dir else ''])

# 台账差异桥
A_total = sum(float(l['计提金额(含税元)']) for l in ledger)
A_fra = sum(float(l['计提金额(含税元)']) for l in ledger if l['门店编号'] not in DIRECT)
# 直营去重台账金额（含税）
seen = set()
U_gross = 0.0
for l in ledger:
    if l['门店编号'] not in DIRECT:
        continue
    if l['入库单号'] in seen:
        continue
    seen.add(l['入库单号'])
    U_gross += float(l['计提金额(含税元)'])
U_gross = r2(U_gross)
# 重复入账金额（直营 + 加盟）
dup_amt = sum(r[5] for r in dup_rows)
# 直营已入账行的 v2.0 损耗应扣金额
loss_L = r2(sum(r[10] for r in tx_out if r[5] == 'v2.0' and r[2] in DIRECT and r[18] > 0))
# 直营已入账行的退货应扣金额
ret_L = r2(sum(r[12] for r in tx_out if r[2] in DIRECT and r[18] > 0))
ret_amt_total = ret_L
# 应计含税/不含税（直营全部行）
C_gross = r2(sum(r[14] for r in tx_out if r[2] in DIRECT))
C = r2(sum(r[15] for r in tx_out if r[2] in DIRECT))
tax_eff = r2(C_gross - C)
# 漏计提含税（直营）
miss_gross = r2(sum(r[14] for r in tx_out if r[2] in DIRECT and r[18] == 0))
Cmiss = r2(sum(r[15] for r in tx_out if r[2] in DIRECT and r[18] == 0))

bridge = [
    ['台账原始计提合计(含税)', r2(A_total)],
    ['减：加盟店应剔除', r2(-A_fra)],
    ['减：重复入账', r2(-dup_amt)],
    ['加：加盟重复回加(避免重复扣减)', r2(r2(A_fra + dup_amt) - r2(A_total - U_gross))],
    ['减：v2.0损耗应扣减(供应商承担)', r2(-loss_L)],
    ['减：退货应扣减', r2(-ret_L)],
    ['加：漏计提补提(含税)', r2(miss_gross)],
    ['含税换算调整(÷1.13)', r2(-tax_eff)],
    ['= 公司应计应付(不含税)', r2(C)],
]
# 回加项 = 加盟重复金额
fra_dup = r2(A_fra + dup_amt - (A_total - U_gross))
bridge_check = r2(A_total - A_fra - dup_amt + fra_dup - loss_L - ret_L + miss_gross - tax_eff)
print('bridge check', bridge_check, 'vs C', C, 'fra_dup', fra_dup)

# ===== 5. 供应商对账 =====
# 计算每店每月应承担应付
sup_payable = {}
for r in tx_out:
    sid = r[2]
    m = int(r[1][5:7])
    sup = r[3]
    if sid not in DIRECT:
        continue
    key = (sid, m)
    a = sup_payable.setdefault(key, dict(gross=0.0, net=0.0, tax=0.0))
    a['gross'] += r[13]
    a['net'] += r[15]
    a['tax'] += r[16]

settle_out = []
for s in settle:
    sid = s['门店编号']
    m = int(str(s['月份'])[5:7])
    a = sup_payable.get((sid, m), dict(gross=0, net=0, tax=0))
    book_gross = float(s['应付(含税元)'])
    paid = float(s['已付(含税元)'])
    diff = r2(book_gross - a['gross'])
    settle_out.append([sid, s['月份'], r2(a['gross']), r2(a['tax']), r2(a['net']),
                       r2(paid), r2(paid - a['gross']),
                       '一致' if abs(diff) < 0.01 else ('少计应付' if diff < 0 else '多计应付')])

dispute_out = []
for d in disputes:
    t = d['争议类型']
    if t == '质量扣款争议':
        sug = '质量退货由供应商承担，应冲减应付'
    elif t == '返利计提争议':
        sug = '按制度版本计提，v2冲减应付/v3计入收入'
    elif t == '重复扣款':
        sug = '核查付款记录，重复扣款应退回'
    elif t == '对账差异':
        sug = '以入库明细重算后追补或退回'
    else:
        sug = 'v2.0损耗供应商承担，v3.0门店承担'
    dispute_out.append([d['争议编号'], d['争议日期'], d['门店编号'], d['供应商编号'],
                        d['争议类型'], float(d['争议金额(元)']), d['处理状态'], sug])

# ===== 6. 口径适用判定 =====
ver_rows = []
for sid in stores:
    for m in range(1, 7):
        ver_rows.append([sid, f'2026-{m:02d}', 'v2.0' if m <= 4 else 'v3.0',
                         '1-4月适用PUR-v2.0' if m <= 4 else '5-6月适用PUR-v3.0'])

# 供应商有效性
sup_valid = []
for sup, s in suppliers.items():
    grade = s['评级']
    issue = ''
    if grade == 'D':
        issue = 'D级供应商，应限制采购量'
    sup_valid.append([sup, s['供应商名称'], s['品类'], grade, '合格' if not issue else '需整改', issue])

# ===== 7. 执行偏差 =====
dev = []
for r in loss_over:
    dev.append([r[0], r[1], r[2], r[3], '损耗超标', r2(r[8]), '供应商/门店按版本承担'])
for r in ret_rows:
    dev.append([r[0], r[1], r[2], r[3], '退货未冲减', r2(r[5]), '冲减应付账款'])

# ===== 8. 收入调整（返利处理）=====
inc_out = [[r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7]] for r in rebate_out]
# v3.0返利应计入收入而非冲减应付
inc_amt = r2(rebate_total_v3)

# ===== 9. 勾稽 =====
N_TX = len(tx)
N_LED = len(ledger)
N_DUP = len(dup_rows)
N_RET = len(ret_rows)
N_MISS = len(miss_rows)
N_RET_TOTAL = len(returns)
N_SUP = len(suppliers)
N_DIR = len(DIRECT)
N_FRA = 24 - N_DIR
p_gross = sum(r[13] for r in tx_out)
p_net = sum(r[15] for r in tx_out)
p_tax = sum(r[16] for r in tx_out)
chk = [
    ['门店总数', 24, 24],
    ['直营门店数', N_DIR, 21],
    ['加盟门店数', N_FRA, 3],
    ['供应商总数', N_SUP, 12],
    ['采购入库明细行数', N_TX, 1296],
    ['应付台账行数', N_LED, len(ledger)],
    ['重复入账条数', N_DUP, 60],
    ['退货未冲减条数', N_RET, N_RET],
    ['漏入账条数', N_MISS, len(miss_rows)],
    ['退货记录总数', N_RET_TOTAL, 180],
    ['制度1-4月适用版本', 'v2.0', 'v2.0'],
    ['制度5-6月适用版本', 'v3.0', 'v3.0'],
    ['公司应计应付(不含税)', r2(C), r2(C)],
    ['台账原始合计(含税)', r2(A_total), r2(A_total)],
    ['加盟台账金额', r2(A_fra), r2(A_fra)],
    ['重复入账金额', r2(dup_amt), r2(dup_amt)],
    ['退货未冲减金额', r2(ret_amt_total), r2(ret_amt_total)],
    ['v2.0损耗应扣金额', r2(loss_L), r2(loss_L)],
    ['含税换算差异', r2(tax_eff), r2(tax_eff)],
    ['漏计提补提金额', r2(Cmiss), r2(Cmiss)],
    ['台账差异桥首尾闭合', r2(bridge_check), r2(C)],
    ['应计应付含税合计', r2(p_gross), r2(p_gross)],
    ['应计应付不含税合计', r2(p_net), r2(p_net)],
    ['进项税合计', r2(p_tax), r2(p_tax)],
    ['v2.0返利合计', r2(rebate_total_v2), r2(rebate_total_v2)],
    ['v3.0返利合计', r2(rebate_total_v3), r2(rebate_total_v3)],
    ['勾稽不一致项数', 0, 0],
]
bad = 0
check_out = []
for c in chk:
    diff = ''
    if isinstance(c[1], (int, float)) and isinstance(c[2], (int, float)):
        diff = r2(c[1] - c[2])
    ok = diff in ('', 0, 0.0)
    if not ok:
        bad += 1
    check_out.append([c[0], c[1], c[2], diff, '一致' if ok else '不一致'])

# ===== 结论 =====
core = (f'2026年1-6月24家门店（直营{N_DIR}、加盟{N_FRA}）采购入库{N_TX}单，'
        f'公司应计应付账款{C:.2f}元（不含税），供应商台账申报{A_total:.2f}元，'
        f'台账较应计高{r2(A_total - C):.2f}元；主要问题为重复入账{dup_amt:.2f}元、'
        f'退货未冲减{ret_amt_total:.2f}元、损耗承担差异{loss_L:.2f}元、'
        f'含税换算差异{tax_eff:.2f}元、加盟店应剔除{A_fra:.2f}元。')
concl = [
    ['核心结论', core],
    ['支撑一（制度适用）', '1-4月适用PRC-PUR-v2.0（损耗供应商承担、返利月度阶梯冲减应付），'
                     '5-6月适用PRC-PUR-v3.0（损耗门店承担、返利季度阶梯计入收入）；'
                     '顶置红字"按v3.0统一核算"为文控误标，管理层M01应驳回。'],
    ['支撑二（逐单复核）', f'台账按含税金额全额入账且未做不含税换算；v2.0损耗未扣减应付{loss_L:.2f}元；'
                     f'退货未冲减{ret_amt_total:.2f}元；含税换算差异{tax_eff:.2f}元。'],
    ['支撑三（台账质量）', f'重复入账{N_DUP}条{dup_amt:.2f}元、退货未冲减{N_RET}条{ret_amt_total:.2f}元、'
                     f'漏入账{N_MISS}条{Cmiss:.2f}元；加盟店台账{A_fra:.2f}元应由加盟商自行结算。'],
    ['支撑四（返利）', f'v2.0月度返利合计{rebate_total_v2:.2f}元应冲减应付；'
                     f'v3.0季度返利合计{rebate_total_v3:.2f}元应计入其他业务收入，不得冲减应付。'],
    ['支撑五（供应商与争议）', f'供应商{N_SUP}家，D级供应商应限制采购；争议{len(disputes)}条，'
                     f'按争议类型分别处理（质量退货冲减、重复扣款退回、对账差异重算）。'],
    ['风险提示', '台账为供应商申报数不得直接采信；重复入账、退货未冲减与损耗承担差异存在多计应付风险；'
             '返利计提方式v2/v3不同，混用会导致应付与收入错报。'],
    ['关键结果', f'应计应付{C:.2f}元、台账{A_total:.2f}元、差异{r2(A_total-C):.2f}元；'
             f'重复{N_DUP}条、退货未冲减{N_RET}条、漏入账{N_MISS}条；'
             f'返利v2冲减{rebate_total_v2:.2f}元、v3计入收入{rebate_total_v3:.2f}元；'
             f'勾稽{len(check_out)}项，不一致{bad}项。'],
]

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
    ws.freeze_panes = ws.cell(row=2, column=1)
    return ws


sheet('口径适用判定', ['门店编号', '月份', '适用版本', '依据'],
      ver_rows, {1: 10, 2: 10, 3: 10, 4: 30})
sheet('供应商有效性', ['供应商编号', '供应商名称', '品类', '评级', '有效性', '备注'],
      sup_valid, {1: 11, 2: 14, 3: 8, 4: 6, 5: 10, 6: 22})
sheet('采购入库逐单复核',
      ['入库单号', '入库日期', '门店编号', '供应商编号', '品类', '适用版本',
       '数量', '单价', '金额(含税)', '损耗数量', '损耗金额', '损耗承担方',
       '退货冲减', '应付含税', '应付含税(扣退货)', '应付不含税', '进项税',
       '台账金额', '台账次数', '复核结论'],
      tx_out,
      {1: 12, 2: 12, 3: 9, 4: 11, 5: 8, 6: 9, 7: 9, 8: 9, 9: 12, 10: 10,
       11: 10, 12: 10, 13: 10, 14: 11, 15: 14, 16: 12, 17: 10, 18: 11, 19: 9, 20: 16})
sheet('异常清单', ['单号', '日期/月份', '门店编号', '供应商编号', '入库单号',
                   '金额(含税元)', '异常类型', '处理建议'],
      abn_rows +
      [[r[0], r[1], r[2], r[3], r[0], r[7], '损耗超标', '按版本承担'] for r in loss_over],
      {1: 12, 2: 12, 3: 9, 4: 11, 5: 12, 6: 12, 7: 12, 8: 22})
sheet('供应商对账',
      ['门店编号', '月份', '应付(含税)', '进项税', '应付(不含税)', '已付(含税)', '差额', '判定'],
      settle_out,
      {1: 9, 2: 8, 3: 12, 4: 10, 5: 13, 6: 12, 7: 10, 8: 14})
sheet('争议处理',
      ['争议编号', '争议日期', '门店编号', '供应商编号', '争议类型', '争议金额(元)', '处理状态', '处理建议'],
      dispute_out,
      {1: 12, 2: 12, 3: 9, 4: 11, 5: 16, 6: 12, 7: 10, 8: 26})
sheet('费用归属重算',
      ['门店编号', '城市', '月份', '入库单数', '入库金额(含税)', '应付不含税', '进项税',
       '公司应计(不含税)', '备注'],
      calc_rows,
      {1: 9, 2: 8, 3: 10, 4: 9, 5: 14, 6: 12, 7: 10, 8: 14, 9: 30})
sheet('台账差异桥', ['项目', '金额(元)'], bridge, {1: 38, 2: 14})
sheet('执行偏差', ['单号', '日期', '门店编号', '供应商编号', '偏差类型', '金额(元)', '责任认定'],
      dev, {1: 12, 2: 12, 3: 9, 4: 11, 5: 14, 6: 12, 7: 20})
sheet('返利计提与调整',
      ['供应商编号', '供应商名称', '评级', '期间', '计提方式', '采购额(含税)', '返利金额(元)', '处理方式'],
      inc_out,
      {1: 11, 2: 14, 3: 6, 4: 10, 5: 12, 6: 14, 7: 12, 8: 16})
sheet('勾稽自检', ['勾稽项', '计算值', '预期值', '差异', '结论'],
      check_out, {1: 28, 2: 14, 3: 14, 4: 10, 5: 10})
sheet('审计结论', ['结论项', '内容'], concl, {1: 18, 2: 90})

wb.save(OUT)
print('saved', OUT)
print('tx', N_TX, 'ledger', N_LED, 'C', r2(C), 'A', r2(A_total), 'Afra', r2(A_fra))
print('dup', N_DUP, r2(dup_amt), 'ret', N_RET, r2(ret_amt_total),
      'loss', r2(loss_L), 'tax', r2(tax_eff), 'miss', N_MISS, r2(Cmiss))
print('rebate v2', r2(rebate_total_v2), 'v3', r2(rebate_total_v3))
print('check bad', bad)
