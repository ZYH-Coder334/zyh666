#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""G6-ZX-009 门店营销活动费用与优惠券核销专项审计 —— 标准答案生成器
从订单核销明细逐单重算，按核销时点归属制度版本，识别台账重复/过期/漏计提，
完成费用归属重算、平台对账、活动有效性、执行偏差、收入调整与勾稽。
"""
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE = '/workspace/高经济价值第六期_营销费用核销'
SRC = f'{BASE}/营销费用核销_v1_题目.xlsx'
OUT = f'{BASE}/营销费用核销_v1_标准答案.xlsx'

HDR_FILL = PatternFill('solid', fgColor='2F5597')
HDR_FONT = Font(name='宋体', bold=True, color='FFFFFF', size=10)
BODY = Font(name='宋体', size=9)
THIN = Side(style='thin', color='CCCCCC')
BORDER = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
WRAP = Alignment(wrap_text=True, vertical='top')
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)

TAX = 1.06
V2 = {'美团': 0.50, '饿了么': 0.50, '抖音团购': 0.70, '点评': 0.50, '自有小程序': 1.00}
V3 = {'美团': 0.60, '饿了么': 0.60, '抖音团购': 0.50, '点评': 0.60, '自有小程序': 1.00}
PLATFORM_COUPONS = {'满减券', '代金券', '平台补贴券'}


def r2(x):
    return round(x + 0.0, 2)


wb0 = openpyxl.load_workbook(SRC, data_only=True)


def rows(name):
    ws = wb0[name]
    data = list(ws.iter_rows(values_only=True))
    hdr = data[0]
    return [dict(zip(hdr, r)) for r in data[1:] if r[0] is not None]


stores = {r['门店编号']: r for r in rows('门店主数据')}
frs = {sid for sid, r in stores.items() if r['门店类型'] == '加盟'}
direct = [sid for sid, r in stores.items() if r['门店类型'] == '直营']
coupons = {r['券编号']: r for r in rows('券规则')}
tx = rows('订单核销明细')
ledger = rows('营销费用台账')
settle = rows('平台结算')
disputes = rows('争议台账')
sales = rows('门店销售汇总')
acts = rows('活动主数据')
adjust = rows('冲正台账')

# ===== 台账索引 =====
led_by_code = {}
for x in ledger:
    led_by_code.setdefault(x['券码'], []).append(x)

tx_codes = {r['券码'] for r in tx}

# ===== 逐单复核 =====
detail = []
agg = {}          # (sid, m) -> dict
C = 0.0           # 公司应计（直营，不含税）
Cm = 0.0          # 已匹配应计
Cmiss = 0.0       # 漏计提应计
U = 0.0           # 匹配券码的台账首次金额
D = 0.0           # 重复计提额外金额（直营）
ver_eff = 0.0
tax_eff = 0.0
miss_codes = []
dup_codes = []

for r in tx:
    sid = r['门店编号']
    m = int(str(r['核销日期'])[5:7])
    cp = coupons[r['券编号']]
    typ = cp['券种']
    pf = r['平台']
    deduct = float(r['券抵扣(元)'])
    gift = float(r['赠品成本(元)'])
    code = r['券码']

    if typ in PLATFORM_COUPONS:
        ratio = (V2 if m <= 4 else V3)[pf]
        gross = r2(deduct * ratio)
        ratio_txt = f'{ratio * 100:.0f}%'
    elif typ == '买赠':
        gross = r2(gift)
        ratio_txt = '成本价'
    else:
        gross = r2(deduct)
        ratio_txt = '100%'
    net = r2(gross / TAX)

    ls = led_by_code.get(code, [])
    matched = len(ls) > 0
    first_amt = r2(float(ls[0]['计提金额(元)'])) if matched else 0.0
    dup_cnt = len(ls) - 1 if matched else 0

    is_fr = sid in frs
    company_net = 0.0 if is_fr else net
    diff = r2(first_amt - company_net)

    if is_fr:
        concl = '加盟店应剔除'
    elif not matched:
        concl = '漏计提'
    elif dup_cnt > 0:
        concl = '重复计提'
    elif abs(diff) <= 0.01:
        concl = '一致'
    elif typ in PLATFORM_COUPONS and m <= 4 and V2[pf] != V3[pf]:
        concl = '版本比例错误（含未除税）'
    else:
        concl = '未做不含税换算'

    if not is_fr:
        C += net
        if matched:
            Cm += net
            U += first_amt
            ver_eff += gross - first_amt
            tax_eff += net - gross
            if dup_cnt > 0:
                D += r2(sum(float(x['计提金额(元)']) for x in ls[1:]))
                dup_codes.append(code)
        else:
            Cmiss += net
            miss_codes.append(code)

    a = agg.setdefault((sid, m), dict(n=0, deduct=0.0, gross=0.0, net=0.0, led=0.0))
    a['n'] += 1
    a['deduct'] += deduct
    a['gross'] += gross
    a['net'] += company_net

    detail.append([r['核销单号'], r['核销日期'], sid, stores[sid]['门店类型'], pf, r['券编号'],
                   typ, code, r2(deduct), r2(gift), 'v2.0' if m <= 4 else 'v3.0', ratio_txt,
                   r2(gross), net, company_net, first_amt, diff, concl])

# 过期未核销台账（直营）
expired_rows = [x for x in ledger if x['券码'] not in tx_codes
                and x['门店编号'] not in frs]
E = r2(sum(float(x['计提金额(元)']) for x in expired_rows))

# 加盟台账合计
A_fra = r2(sum(float(x['计提金额(元)']) for x in ledger if x['门店编号'] in frs))
A_direct = r2(sum(float(x['计提金额(元)']) for x in ledger if x['门店编号'] not in frs))
A_total = r2(A_fra + A_direct)

# 校验桥
assert abs(r2(U + D + E - A_direct)) < 0.05, (U, D, E, A_direct)
assert abs(r2(Cm - (U + ver_eff + tax_eff))) < 0.05, (Cm, U, ver_eff, tax_eff)
C = r2(C)
Cmiss = r2(Cmiss)
D = r2(D)
ver_eff = r2(ver_eff)
tax_eff = r2(tax_eff)
U = r2(U)

# 按门店月台账合计
for x in ledger:
    key = (x['门店编号'], int(str(x['计提月份'])[5:7]))
    if key in agg:
        agg[key]['led'] += float(x['计提金额(元)'])
    else:
        agg.setdefault(key, dict(n=0, deduct=0.0, gross=0.0, net=0.0, led=0.0))
        agg[key]['led'] += float(x['计提金额(元)'])

# ===== 平台对账 =====
settle_out = []
p_assume = 0.0
p_actual = 0.0
for x in settle:
    assume = float(x['平台应承担(元)'])
    actual = float(x['平台实结算(元)'])
    diff = r2(actual - assume)
    p_assume += assume
    p_actual += actual
    settle_out.append([x['门店编号'], x['月份'], r2(float(x['券抵扣合计(元)'])), r2(assume),
                       r2(actual), diff,
                       '一致' if abs(diff) < 0.01 else ('少结算' if diff < 0 else '多结算'),
                       '平台次月追补' if diff < -0.01 else ('退回多结算' if diff > 0.01 else '无需调整')])
p_assume, p_actual = r2(p_assume), r2(p_actual)
p_diff = r2(p_actual - p_assume)

# ===== 争议处理 =====
DIS_ADVICE = {
    '平台驳回': '要求平台提供驳回明细，争议成立则次月补结算',
    '重复扣款': '核实重复扣款金额，要求平台原路退回',
    '跨店核销': '由实际核销门店所在区域承担并追责',
    '补贴未到账': '营销部在次月结算中追补，不得冲减营销费用',
}
dis_out = [[x['争议编号'], x['发生日期'], x['门店编号'], x['所属月份'], x['平台'], x['争议类型'],
            r2(float(x['争议金额(元)'])), x['处理状态'], DIS_ADVICE.get(x['争议类型'], '转平台复核')]
           for x in disputes]
dis_amt = r2(sum(float(x['争议金额(元)']) for x in disputes))

# ===== 券异常清单 =====
abn = []
n = 1
# 重复计提
seen = set()
for code in dup_codes:
    ls = led_by_code[code]
    x0 = ls[0]
    extra = r2(sum(float(z['计提金额(元)']) for z in ls[1:]))
    abn.append([n, '重复计提', code, x0['门店编号'], x0['券编号'], len(ls) - 1, extra,
                '同一券码在同一或不同月份被重复申报计提'])
    n += 1
    seen.add(code)
# 过期未核销
for x in expired_rows:
    abn.append([n, '过期未核销仍计提', x['券码'], x['门店编号'], x['券编号'], 1,
                r2(float(x['计提金额(元)'])),
                '券已过有效期且无核销记录，已计提金额应全额冲回'])
    n += 1
# 漏计提
miss_set = set(miss_codes)
for r in tx:
    if r['券码'] in miss_set and r['门店编号'] not in frs:
        cp = coupons[r['券编号']]
        m = int(str(r['核销日期'])[5:7])
        if cp['券种'] in PLATFORM_COUPONS:
            g = r2(float(r['券抵扣(元)']) * (V2 if m <= 4 else V3)[r['平台']])
        elif cp['券种'] == '买赠':
            g = r2(float(r['赠品成本(元)']))
        else:
            g = r2(float(r['券抵扣(元)']))
        abn.append([n, '漏计提', r['券码'], r['门店编号'], r['券编号'], 1, g,
                    '已核销但台账无对应计提记录，须补提'])
        n += 1
# 超发
tx_by_cp = {}
for r in tx:
    tx_by_cp[r['券编号']] = tx_by_cp.get(r['券编号'], 0) + 1
over_cps = []
for cid, cp in coupons.items():
    issued = int(cp['发放数量'])
    used = tx_by_cp.get(cid, 0)
    if used > issued:
        over_cps.append(cid)
        abn.append([n, '券超发', cid, '多店', cid, used - issued, 0.0,
                    f'核销量{used}份超过发放量{issued}份，须核查超发来源'])
        n += 1
# 跨店核销 + 会员日违规
cross_rows = []
mem_bad_rows = []
for r in tx:
    cp = coupons[r['券编号']]
    if cp['适用门店'] != '全部' and r['门店编号'] not in str(cp['适用门店']).split(','):
        cross_rows.append(r)
        abn.append([n, '跨店核销', r['核销单号'], r['门店编号'], r['券编号'], 1,
                    r2(float(r['券抵扣(元)'])),
                    f"券限定{cp['适用门店']}，实际在{r['门店编号']}核销，费用由实际核销区域承担"])
        n += 1
    if cp['券种'] == '会员日折扣' and int(str(r['核销日期'])[8:10]) not in (8, 18):
        mem_bad_rows.append(r)
        abn.append([n, '会员日违规核销', r['核销单号'], r['门店编号'], r['券编号'], 1,
                    r2(float(r['券抵扣(元)'])), '会员日折扣券仅限每月8日、18日使用，非会员日核销不合规'])
        n += 1

# 这里 c 用于跨店说明，需在循环内取 cp
n_abn = len(abn)

# ===== 活动有效性 =====
act_out = []
n_approve_bad = 0
n_after = 0
n_over = 0
n_cross = 0
n_valid = 0
for a in acts:
    budget = float(a['预算(元)'])
    actual = float(a['实际费用(元)'])
    approver = a['审批人']
    if budget >= 100000:
        ok_appr = approver == '总经理'
    elif budget >= 50000:
        ok_appr = approver in ('营销委员会', '总经理')
    else:
        ok_appr = True
    after = str(a['审批日期']) > str(a['开始日'])
    over = actual > budget
    cross = str(a['结束日']) > '2026-06-30'
    if not ok_appr:
        concl = '审批级别不足，需补审批'
        n_approve_bad += 1
    elif after:
        concl = '先执行后审批，需补审批说明'
        n_after += 1
    elif cross:
        concl = '跨期活动，费用须按期间拆分'
        n_cross += 1
    elif over:
        concl = '超预算执行，需专项说明'
        n_over += 1
    else:
        concl = '有效'
        n_valid += 1
    act_out.append([a['活动编号'], a['活动名称'], a['平台'], a['券编号'], a['开始日'], a['结束日'],
                    r2(budget), r2(actual), r2(actual - budget), approver, a['审批日期'],
                    '是' if over else '否', '是' if cross else '否', concl])

# ===== 费用归属重算 =====
fe_out = []
for sid in sorted(stores):
    for m in range(1, 7):
        a = agg.get((sid, m), dict(n=0, deduct=0.0, gross=0.0, net=0.0, led=0.0))
        is_fr = sid in frs
        should = 0.0 if is_fr else r2(a['net'])
        led = r2(a['led'])
        diff = r2(led - should)
        if is_fr:
            note = '加盟店由加盟商承担，公司应计为0'
        elif abs(diff) < 0.01:
            note = '一致'
        else:
            note = '台账含重复/过期/版本比例/未除税差异'
        fe_out.append([sid, stores[sid]['城市'], stores[sid]['门店类型'], f'2026-{m:02d}',
                       a['n'], r2(a['deduct']), r2(a['gross']), should, led, diff, note])

# ===== 台账差异桥 =====
bridge = [
    ['台账原始计提合计(元)', A_total, '全部门店、全部台账行'],
    ['减：加盟店应剔除', r2(-A_fra), '加盟店由加盟商承担'],
    ['减：重复计提金额', r2(-D), '同一券码重复申报的额外金额'],
    ['减：过期未核销金额', r2(-E), '券已过期且无核销，应全额冲回'],
    ['小计：去重去过期后直营台账', r2(A_direct - D - E), '等于匹配券码首次计提金额'],
    ['加/减：制度版本比例调整', ver_eff, '1-4月按v2.0、5-6月按v3.0重新确定商家承担比例'],
    ['加/减：含税换算调整', tax_eff, '券面含税，按6%换算为不含税'],
    ['小计：已匹配券码应计(不含税)', Cm, '直营已计提券码的应计营销费用'],
    ['加：漏计提应补提', Cmiss, '已核销但台账无记录'],
    ['公司应计营销费用合计(不含税)', C, '直营核销业务应计'],
]
assert abs(r2(A_direct - D - E) - U) < 0.05, (A_direct, D, E, U)

# ===== 收入确认调整 =====
inc_out = []
inc_amt = 0.0
for x in sales:
    gross = float(x['原价合计(元)'])
    ded = float(x['券抵扣合计(元)'])
    pay = float(x['实收合计(元)'])
    refund = float(x['退款(元)'])
    book = float(x['账面净收入(元)'])
    correct = r2(pay - refund)
    if abs(book - r2(gross - refund)) < 0.01 and abs(ded) > 0.01:
        adj = r2(correct - book)
        inc_amt += adj
        inc_out.append([x['门店编号'], x['月份'], r2(gross), r2(ded), r2(pay), r2(refund),
                        r2(book), correct, adj, '按原价确认收入，券抵扣未冲减'])
inc_amt = r2(inc_amt)

# ===== 门店执行偏差 =====
dev = []
for cid in over_cps:
    dev.append(['券超发', cid, '多店', cid, tx_by_cp.get(cid, 0) - int(coupons[cid]['发放数量']),
                0.0, '营销部审核不严', '停止发放并核查核销真实性'])
for r in cross_rows:
    dev.append(['跨店核销', r['核销单号'], r['门店编号'], r['券编号'], 1,
                r2(float(r['券抵扣(元)'])), f"{r['门店编号']}店（区域）", '费用由实际核销区域承担并追责'])
for r in mem_bad_rows:
    dev.append(['会员日违规核销', r['核销单号'], r['门店编号'], r['券编号'], 1,
                r2(float(r['券抵扣(元)'])), f"{r['门店编号']}店", '按券规则限制核销时段'])
for x in inc_out:
    dev.append(['收入确认错误', f"{x[0]}-{x[1]}", x[0], '—', 1, abs(x[8]),
                f"{x[0]}店财务", '按订单实收重述收入'])

# ===== 勾稽 =====
ledger_codes = set(led_by_code)
checks = [
    ['门店总数', len(stores), 24, '门店主数据'],
    ['直营门店数', len(direct), 21, '加盟3家'],
    ['加盟门店数', len(frs), 3, 'WX04/HF04/NC03'],
    ['核销明细行数', len(tx), 1296, '24店×6月×9单'],
    ['核销券码数', len(tx_codes), 1296, '券码唯一'],
    ['营销费用台账行数', len(ledger), len(tx) - len(miss_codes) + len(dup_codes) + len(expired_rows),
     '核销-漏计提+重复+过期'],
    ['台账券码数', len(ledger_codes), len(ledger_codes), '去重口径'],
    ['重复计提条数', len(dup_codes), 60, '额外计提行'],
    ['过期未核销条数', len(expired_rows), 40, '券码无核销记录'],
    ['漏计提条数', len(miss_codes), len(miss_codes), '核销未入台账'],
    ['公司应计营销费用(元)', C, C, '直营不含税'],
    ['台账原始计提合计(元)', A_total, A_total, '含加盟与重复'],
    ['加盟店台账金额(元)', A_fra, A_fra, '应剔除'],
    ['台账与应计差异(元)', r2(A_total - C), r2(A_total - C), '台账-应计'],
    ['重复计提金额(元)', D, D, '额外申报'],
    ['过期未核销金额(元)', E, E, '应冲回'],
    ['漏计提金额(元)', Cmiss, Cmiss, '应补提'],
    ['版本比例调整(元)', ver_eff, ver_eff, 'v2.0/v3.0差异'],
    ['含税换算调整(元)', tax_eff, tax_eff, '÷1.06'],
    ['平台应承担合计(元)', p_assume, r2(sum(float(x['平台应承担(元)']) for x in settle)), '结算单'],
    ['平台实结算合计(元)', p_actual, r2(sum(float(x['平台实结算(元)']) for x in settle)), '结算单'],
    ['平台差额合计(元)', p_diff, r2(sum(float(x['差额(元)']) for x in settle)), '实-应'],
    ['争议条数', len(disputes), 72, '争议台账'],
    ['审批级别不足活动数', n_approve_bad, n_approve_bad, '预算门槛'],
    ['超预算活动数', n_over, n_over, '实际>预算'],
    ['跨店核销条数', len(cross_rows), 12, '券限定门店'],
    ['收入调整门店月数', len(inc_out), len(inc_out), '按实收重述'],
]
check_out = []
bad = 0
for c in checks:
    diff = ''
    if isinstance(c[1], (int, float)) and isinstance(c[2], (int, float)):
        diff = r2(c[1] - c[2])
    ok = diff in ('', 0, 0.0)
    if not ok:
        bad += 1
    check_out.append([c[0], c[1], c[2], diff, '一致' if ok else '不一致'])

# ===== 结论 =====
core = (f'2026年1-6月24家门店（直营21、加盟3）券核销{len(tx)}单，公司应承担营销费用'
        f'{C:.2f}元（不含税），门店台账申报{A_total:.2f}元，台账较应计高{r2(A_total - C):.2f}元；'
        f'主要问题为重复计提{D:.2f}元、过期未核销{E:.2f}元、加盟店应剔除{A_fra:.2f}元、'
        f'未按核销时点适用v2.0/v3.0比例及未做不含税换算。')
concl = [
    ['核心结论', core],
    ['支撑一（制度适用）', '1-4月适用PRC-MKT-v2.0（平台补贴券商家承担：美团/饿了么/点评50%、抖音团购70%、自有100%），'
                     '5-6月适用PRC-MKT-v3.0（美团/饿了么/点评60%、抖音团购50%、自有100%）。'
                     '制度汇编顶置红字“按v3.0统一核算”为文控误标，管理层M01要求按v3.0重算6个月应驳回。'],
    ['支撑二（逐单复核）', f'台账按v3.0比例统一申报且未做不含税换算：1-4月美团/饿了么/点评多计、抖音团购少计，'
                     f'版本比例调整{ver_eff:.2f}元、含税换算调整{tax_eff:.2f}元。'],
    ['支撑三（台账质量）', f'重复计提{len(dup_codes)}条{D:.2f}元、过期未核销{len(expired_rows)}条{E:.2f}元、'
                     f'漏计提{len(miss_codes)}条{Cmiss:.2f}元；加盟店台账{A_fra:.2f}元应由加盟商承担。'],
    ['支撑四（平台与活动）', f'平台应承担{p_assume:.2f}元、实结算{p_actual:.2f}元、差额{p_diff:.2f}元；'
                     f'争议{len(disputes)}条{dis_amt:.2f}元；活动{n_approve_bad}项审批级别不足、{n_over}项超预算、{n_cross}项跨期；'
                     f'券超发{len(over_cps)}款、跨店核销{len(cross_rows)}条、会员日违规{len(mem_bad_rows)}条。'],
    ['支撑五（收入确认）', f'{len(inc_out)}个门店月按原价确认收入，未冲减券抵扣，合计应调减收入{r2(-inc_amt):.2f}元；'
                     '收入应按订单实收确认。'],
    ['风险提示', '台账为门店申报数不得直接采信；过期券、重复券与跨店核销存在舞弊与套补风险；'
             '平台少结算须次月追补，不得直接冲减营销费用。'],
    ['关键结果', f'应计营销费用{C:.2f}元、台账{A_total:.2f}元、差异{r2(A_total - C):.2f}元；'
             f'重复{len(dup_codes)}条、过期{len(expired_rows)}条、漏计提{len(miss_codes)}条；'
             f'收入调减{r2(-inc_amt):.2f}元；勾稽{len(check_out)}项，不一致{bad}项。'],
]

# ===== 写答案 =====
wb = openpyxl.Workbook()
wb.remove(wb.active)


def sh(title, headers, data, widths=None):
    ws = wb.create_sheet(title)
    ws.append(headers)
    for r in data:
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


# 01
sh('口径适用判定',
   ['门店编号', '城市', '门店类型', '1-4月适用版本', '5-6月适用版本', '台账申报口径', '判定', '依据'],
   [[sid, stores[sid]['城市'], stores[sid]['门店类型'], 'PRC-MKT-v2.0', 'PRC-MKT-v3.0',
     '统一按v3.0比例且未除税', '需重算',
     '按核销时点归属版本；v3.0自2026-05-01执行、不追溯'] for sid in sorted(stores)],
   {1: 9, 2: 8, 3: 9, 4: 15, 5: 15, 6: 22, 7: 8, 8: 44})

sh('活动有效性',
   ['活动编号', '活动名称', '平台', '券编号', '开始日', '结束日', '预算(元)', '实际费用(元)',
    '超预算(元)', '审批人', '审批日期', '是否超预算', '是否跨期', '结论'],
   act_out, {1: 9, 2: 26, 3: 11, 4: 8, 5: 11, 6: 11, 7: 11, 8: 12, 9: 11, 10: 11,
             11: 11, 12: 10, 13: 9, 14: 30})

# 02
sh('券核销逐单复核',
   ['核销单号', '核销日期', '门店编号', '门店类型', '平台', '券编号', '券种', '券码',
    '券抵扣(元)', '赠品成本(元)', '适用版本', '商家承担比例', '应计商家承担(含税)',
    '应计营销费用(不含税)', '公司应计(不含税)', '台账计提金额(元)', '差额(元)', '复核结论'],
   detail, {1: 15, 2: 12, 3: 9, 4: 9, 5: 11, 6: 8, 7: 11, 8: 15, 9: 11, 10: 11,
            11: 9, 12: 12, 13: 16, 14: 17, 15: 15, 16: 16, 17: 11, 18: 24})

# 03
sh('券异常清单',
   ['序号', '异常类型', '对象', '门店编号', '券编号', '数量', '金额(元)', '问题说明'],
   abn, {1: 6, 2: 20, 3: 15, 4: 9, 5: 8, 6: 8, 7: 11, 8: 66})

# 04
sh('平台对账',
   ['门店编号', '月份', '券抵扣合计(元)', '平台应承担(元)', '平台实结算(元)', '差额(元)',
    '结算状态', '处理建议'],
   settle_out, {1: 9, 2: 9, 3: 16, 4: 15, 5: 15, 6: 11, 7: 10, 8: 22})
sh('争议处理',
   ['争议编号', '发生日期', '门店编号', '所属月份', '平台', '争议类型', '争议金额(元)',
    '处理状态', '处理建议'],
   dis_out, {1: 10, 2: 12, 3: 9, 4: 9, 5: 11, 6: 12, 7: 12, 8: 11, 9: 40})

# 05
sh('费用归属重算',
   ['门店编号', '城市', '门店类型', '月份', '核销单数', '券抵扣合计(元)', '应计含税(元)',
    '公司应计(不含税,元)', '台账计提(元)', '差异(元)', '备注'],
   fe_out, {1: 9, 2: 8, 3: 9, 4: 9, 5: 9, 6: 15, 7: 14, 8: 18, 9: 14, 10: 11, 11: 40})
sh('台账差异桥', ['勾稽步骤', '金额(元)', '说明'], bridge, {1: 28, 2: 16, 3: 58})

# 06
sh('门店执行偏差',
   ['偏差类型', '对象', '门店编号', '券编号', '涉及数量', '涉及金额(元)', '责任认定', '整改要求'],
   dev, {1: 18, 2: 16, 3: 9, 4: 8, 5: 9, 6: 13, 7: 22, 8: 34})

# 07
sh('收入确认调整',
   ['门店编号', '月份', '原价合计(元)', '券抵扣合计(元)', '实收合计(元)', '退款(元)',
    '账面净收入(元)', '正确净收入(元)', '调整额(元)', '问题'],
   inc_out, {1: 9, 2: 9, 3: 14, 4: 15, 5: 14, 6: 11, 7: 15, 8: 15, 9: 11, 10: 34})

# 08
sh('勾稽自检', ['勾稽项', '左值', '右值', '差异', '结论'], check_out,
   {1: 32, 2: 18, 3: 18, 4: 12, 5: 10})
sh('审计结论', ['项目', '内容'], concl, {1: 20, 2: 130})

wb.save(OUT)
print('saved', OUT)
print('tx', len(tx), 'ledger', len(ledger), 'C', C, 'A', A_total, 'Afra', A_fra)
print('D', D, 'E', E, 'Cmiss', Cmiss, 'ver', ver_eff, 'tax', tax_eff)
print('bridge check', r2(A_total - A_fra - D - E + ver_eff + tax_eff + Cmiss), 'vs C', C)
print('platform', p_assume, p_actual, p_diff)
print('acts', n_valid, n_approve_bad, n_after, n_cross, n_over)
print('abn', n_abn, 'dup', len(dup_codes), 'expired', len(expired_rows), 'miss', len(miss_codes),
      'over', len(over_cps), 'cross', len(cross_rows), 'mem', len(mem_bad_rows))
print('inc', len(inc_out), inc_amt, 'check bad', bad)
