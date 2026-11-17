#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""G6-ZX-009 门店营销活动费用与优惠券核销专项审计 —— 题目数据生成器

沿用已通过题目（0876 / 0880）的经验：
1) 约 3240 行数据，审计期间横跨制度变更日（1-4 月 v2.0 / 5-6 月 v3.0），按业务核销时点归属；
2) 已发台账是门店申报的计提数，含重复计提、过期未核销、漏计提，不能直接当应计依据；
3) 券面值含税、营销费用按不含税（6%）列示，须逐单换算；
4) 加盟店营销费用公司不承担，但仍纳入门店执行偏差检查；
5) 关键规则藏在《营销委员会会议纪要》与工作表批注；
6) 制度汇编顶置红字"按 v3.0 统一核算"为文控误标，构成显著性反转。
"""
import os
import random
from datetime import date, timedelta

import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

random.seed(20260609)

BASE = '/workspace/高经济价值第六期_营销费用核销'
os.makedirs(BASE, exist_ok=True)
OUT = f'{BASE}/营销费用核销_v1_题目.xlsx'

HDR_FILL = PatternFill('solid', fgColor='2F5597')
HDR_FONT = Font(name='宋体', bold=True, color='FFFFFF', size=10)
BODY = Font(name='宋体', size=9)
RED_FONT = Font(name='宋体', bold=True, color='C00000', size=9)
THIN = Side(style='thin', color='CCCCCC')
BORDER = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
WRAP = Alignment(wrap_text=True, vertical='top')
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)


def r2(x):
    return round(x + 0.0, 2)


TAX = 1.06
V2_RATIO = {'美团': 0.50, '饿了么': 0.50, '抖音团购': 0.70, '点评': 0.50, '自有小程序': 1.00}
V3_RATIO = {'美团': 0.60, '饿了么': 0.60, '抖音团购': 0.50, '点评': 0.60, '自有小程序': 1.00}
PLATFORMS = ['美团', '饿了么', '抖音团购', '点评', '自有小程序']

# 门店编号, 城市, 区域, 商圈, 店长, 类型, 开业日期
STORES = [
    ['SZ01', '苏州', '工业园区', '湖东邻里中心', '张伟', '直营', '2023-05-10'],
    ['SZ02', '苏州', '高新区', '狮山天街', '李娜', '直营', '2023-08-15'],
    ['SZ03', '苏州', '吴中区', '万象汇', '王强', '直营', '2024-01-20'],
    ['SZ04', '苏州', '相城区', '天虹百货', '赵敏', '直营', '2024-05-08'],
    ['SZ05', '苏州', '姑苏区', '印象城', '钱进', '直营', '2024-09-12'],
    ['WX01', '无锡', '梁溪区', '崇安寺', '陈杰', '直营', '2023-03-18'],
    ['WX02', '无锡', '滨湖区', '滨湖万达', '刘洋', '直营', '2023-09-12'],
    ['WX03', '无锡', '新吴区', '宝龙广场', '孙丽', '直营', '2024-03-22'],
    ['WX04', '无锡', '惠山区', '惠山万达', '周涛', '加盟', '2024-07-30'],
    ['NB01', '宁波', '海曙区', '天一广场', '吴静', '直营', '2023-04-25'],
    ['NB02', '宁波', '鄞州区', '印象城', '郑磊', '直营', '2023-10-08'],
    ['NB03', '宁波', '江北区', '来福士', '黄芳', '直营', '2024-02-14'],
    ['NB04', '宁波', '镇海区', '吾悦广场', '林峰', '直营', '2024-06-18'],
    ['HZ01', '杭州', '上城区', '湖滨银泰', '徐明', '直营', '2025-03-15'],
    ['HZ02', '杭州', '上城区', '钱江新城来福士', '朱琳', '直营', '2025-06-01'],
    ['HZ03', '杭州', '西湖区', '城西银泰', '何军', '直营', '2025-08-20'],
    ['HZ04', '杭州', '滨江区', '滨江宝龙', '韩磊', '直营', '2025-11-15'],
    ['HF01', '合肥', '蜀山区', '政务区银泰', '高飞', '直营', '2025-05-20'],
    ['HF02', '合肥', '包河区', '滨湖万达', '罗敏', '直营', '2025-07-10'],
    ['HF03', '合肥', '蜀山区', '三里庵国购', '崔浩', '直营', '2025-08-20'],
    ['HF04', '合肥', '庐阳区', '银泰中心', '邹敏', '加盟', '2025-10-15'],
    ['NC01', '南昌', '红谷滩区', '红谷滩万达', '宋佳', '直营', '2024-11-05'],
    ['NC02', '南昌', '青山湖区', '青山湖万象汇', '谢峰', '直营', '2025-05-20'],
    ['NC03', '南昌', '新建区', '新建吾悦', '梁宇', '加盟', '2025-12-25'],
]
S = {s[0]: s for s in STORES}
FRS = {s[0] for s in STORES if s[5] == '加盟'}
DIRECT = [s[0] for s in STORES if s[5] == '直营']

# ============ 券规则 ============
COUPON_TYPES = ['满减券', '满减券', '代金券', '折扣券', '买赠', '会员日折扣', '平台补贴券', '满减券', '代金券', '折扣券']
FACES = [10, 15, 20, 25, 30, 40, 50]
RATES = [0.85, 0.88, 0.90, 0.95]
GIFTS = [8, 12, 15, 18, 22]
THRESH = [30, 50, 60, 80, 100]
PLATFORM_COUPONS = {'满减券', '代金券', '平台补贴券'}
RESTRICT = {
    5: ['SZ01', 'SZ02', 'SZ03'],
    11: ['WX01', 'WX02'],
    17: ['NB01', 'NB02', 'NB03'],
    23: ['HZ01', 'HZ02'],
    29: ['HF01', 'HF02', 'HF03'],
    35: ['NC01', 'NC02'],
}

coupons = []
k = 0
for pi, pf in enumerate(PLATFORMS):
    cnt = 10 if pi < 3 else 9
    for j in range(cnt):
        typ = COUPON_TYPES[j % len(COUPON_TYPES)]
        face = FACES[(pi * 3 + j) % len(FACES)]
        rate = RATES[(pi + j) % len(RATES)]
        gift = GIFTS[(pi * 2 + j) % len(GIFTS)]
        thresh = THRESH[(pi + j * 2) % len(THRESH)]
        cid = f'Q{pi + 1:02d}{j + 1:02d}'
        if typ in PLATFORM_COUPONS:
            label = f'{face}元'
        elif typ == '折扣券':
            label = f'{rate * 10:.1f}折'
        elif typ == '会员日折扣':
            label = '85折'
        else:
            label = '赠品1份'
        end = '2026-06-30'
        if k in (8, 16, 24, 32, 40, 46):
            end = '2026-04-30'
        start = '2026-03-01' if k in (12, 20, 28, 36) else '2026-01-01'
        scope = '全部'
        if k in RESTRICT:
            scope = ','.join(RESTRICT[k])
        coupons.append(dict(idx=k, id=cid, name=f'{pf}{typ}{label}活动券', typ=typ, platform=pf,
                            face=face, rate=rate, gift=gift, thresh=thresh, label=label,
                            start=start, end=end, scope=scope, issued=0))
        k += 1
CP = {c['id']: c for c in coupons}

# ============ 核销计划 ============
plan = []
for s in STORES:
    cand = [c for c in coupons if c['scope'] == '全部' or s[0] in str(c['scope']).split(',')]
    for m in range(1, 7):
        for _ in range(9):
            plan.append(dict(s=s, m=m, c=random.choice(cand), mem_viol=False))

# 跨店核销 12 条
viol = [('SZ05', 5), ('WX03', 5), ('NB04', 5), ('HZ03', 5), ('HF04', 5), ('NC03', 5),
        ('SZ04', 17), ('WX04', 17), ('WX01', 23), ('HZ04', 17), ('HF04', 17), ('NC03', 17)]
for k, (sid, ci) in enumerate(viol):
    pos = [i for i, p in enumerate(plan) if p['s'][0] == sid]
    plan[pos[(k * 2) % len(pos)]]['c'] = coupons[ci]

# 会员日券非会员日核销 8 条
mem_cp = [c for c in coupons if c['typ'] == '会员日折扣' and c['scope'] == '全部']
mem_bad = ['SZ02', 'WX02', 'NB02', 'HZ02', 'HF02', 'NC02', 'SZ04', 'WX03']
for k, sid in enumerate(mem_bad):
    pos = [i for i, p in enumerate(plan) if p['s'][0] == sid]
    p = plan[pos[(k * 3 + 1) % len(pos)]]
    p['c'] = mem_cp[k % len(mem_cp)]
    p['mem_viol'] = True

for i, p in enumerate(plan):
    c = p['c']
    if c['typ'] == '会员日折扣':
        day = 22 if p['mem_viol'] else random.choice([8, 18])
    else:
        day = random.randint(1, 28)
    p['date'] = f"2026-{p['m']:02d}-{day:02d}"
    orig = round(c['thresh'] + random.uniform(8, 220), 2)
    if c['typ'] in PLATFORM_COUPONS:
        deduct = min(float(c['face']), round(orig * 0.5, 2))
    elif c['typ'] == '折扣券':
        deduct = round(orig * (1 - c['rate']), 2)
    elif c['typ'] == '会员日折扣':
        deduct = round(orig * 0.15, 2)
    else:
        deduct = 0.0
    p['orig'] = orig
    p['deduct'] = deduct
    p['gift'] = float(c['gift']) if c['typ'] == '买赠' else 0.0
    p['pay'] = round(orig - deduct, 2)
    p['code'] = f'CP{202600000 + i:09d}'

# 券发放数量（制造超发）
cnt_by_cp = {}
for p in plan:
    cnt_by_cp[p['c']['id']] = cnt_by_cp.get(p['c']['id'], 0) + 1
for c in coupons:
    n = cnt_by_cp.get(c['id'], 0)
    if c['idx'] in (3, 9, 14, 21, 28, 33):
        c['issued'] = max(10, int(n * 0.65))
    else:
        c['issued'] = max(60, int(n * random.uniform(1.6, 3.0)))

# ============ 营销费用台账（门店申报口径：统一按 v3.0 比例、未除税） ============
ledger = []
ln = 1
missing_idx = set()
for i, p in enumerate(plan):
    if i % 43 == 0 and p['s'][5] == '直营':
        missing_idx.add(i)
        continue
    c = p['c']
    if c['typ'] in PLATFORM_COUPONS:
        amt = r2(p['deduct'] * V3_RATIO[c['platform']])
    elif c['typ'] == '买赠':
        amt = r2(p['gift'])
    else:
        amt = r2(p['deduct'])
    ledger.append(dict(no=f'TJ{ln:06d}', month=f"2026-{p['m']:02d}", sid=p['s'][0],
                       act=f"A{c['idx'] + 1:02d}", cid=c['id'], code=p['code'],
                       amt=amt, basis='门店月度申报', status='已计提'))
    ln += 1

# 重复计提 60 条（仅直营，便于公司口径统计）
dup_src = [i for i, p in enumerate(plan) if i not in missing_idx and p['s'][5] == '直营']
dup_src = dup_src[3::17][:60]
for i in dup_src:
    p = plan[i]
    c = p['c']
    if c['typ'] in PLATFORM_COUPONS:
        amt = r2(p['deduct'] * V3_RATIO[c['platform']])
    elif c['typ'] == '买赠':
        amt = r2(p['gift'])
    else:
        amt = r2(p['deduct'])
    ledger.append(dict(no=f'TJ{ln:06d}', month=f"2026-{p['m']:02d}", sid=p['s'][0],
                       act=f"A{c['idx'] + 1:02d}", cid=c['id'], code=p['code'],
                       amt=amt, basis='门店补报', status='已发'))
    ln += 1

# 过期未核销仍计提 40 条
expired_cps = [c for c in coupons if c['end'] == '2026-04-30']
for k in range(40):
    c = expired_cps[k % len(expired_cps)]
    sid = DIRECT[k % len(DIRECT)]
    amt = r2(float(c['face'] if c['typ'] in PLATFORM_COUPONS else 20) * 1.0)
    ledger.append(dict(no=f'TJ{ln:06d}', month='2026-05' if k % 2 == 0 else '2026-06',
                       sid=sid, act=f"A{c['idx'] + 1:02d}", cid=c['id'],
                       code=f'CP9{202600000 + k:09d}', amt=amt,
                       basis='门店月度申报', status='已发'))
    ln += 1

ledger.sort(key=lambda x: x['no'])

# ============ 活动主数据 ============
activities = []
d0 = date(2026, 1, 1)
OVER_ACT = {16, 18, 20, 22, 24, 26, 28, 30}
for c in coupons:
    aid = f"A{c['idx'] + 1:02d}"
    if c['idx'] < 12:
        budget = 60000 + (c['idx'] % 5) * 15000
    elif c['idx'] in (40, 41, 42, 43):
        budget = 52000 + (c['idx'] % 4) * 9000
    else:
        budget = 30000 + (c['idx'] * 7 % 9) * 8000
    if c['idx'] in OVER_ACT:
        actual = r2(budget * 1.08)
    else:
        actual = r2(budget * random.uniform(0.78, 0.97))
    st = d0 + timedelta(days=c['idx'] * 3)
    en = st + timedelta(days=20 + c['idx'] % 40)
    if c['idx'] in (40, 41, 42, 43):
        en = date(2026, 7, 15)
    if c['idx'] < 12:
        approver = '区域经理'
    elif budget >= 100000:
        approver = '总经理'
    else:
        approver = '营销委员会'
    adate = st - timedelta(days=3 + c['idx'] % 13)
    if c['idx'] in (12, 13, 14, 15):
        adate = st + timedelta(days=5)
    activities.append([aid, c['name'], f"{c['platform']}·{c['scope']}", c['platform'], c['id'],
                       str(st), str(en), budget, approver, str(adate), actual,
                       '已结束' if en <= date(2026, 6, 30) else '执行中'])

# ============ 平台结算与争议 ============
settle = []
for si, sid in enumerate(DIRECT):
    for m in range(1, 7):
        recs = [p for p in plan if p['s'][0] == sid and p['m'] == m]
        ded = r2(sum(p['deduct'] for p in recs))
        ratio = V2_RATIO if m <= 4 else V3_RATIO
        assume = r2(sum(r2(p['deduct'] * (1 - ratio[p['c']['platform']]))
                        for p in recs if p['c']['typ'] in PLATFORM_COUPONS))
        diff = 0.0
        if (si + m) % 5 == 0:
            diff = -r2(assume * random.uniform(0.03, 0.12))
        elif (si + m) % 17 == 0:
            diff = r2(assume * 0.02)
        actual = r2(assume + diff)
        status = '一致' if abs(diff) < 0.01 else ('少结算' if diff < 0 else '多结算')
        settle.append([sid, f'2026-{m:02d}', ded, assume, actual, r2(diff), status])

disputes = []
DTYPES = ['平台驳回', '重复扣款', '跨店核销', '补贴未到账']
for k in range(72):
    sid = DIRECT[(k * 5) % len(DIRECT)]
    m = k % 6 + 1
    disputes.append([f'ZY{k + 1:04d}', f'2026-{m:02d}-{random.randint(3, 26):02d}', sid,
                     f'2026-{m:02d}', random.choice(PLATFORMS), DTYPES[k % 4],
                     r2(random.uniform(50, 800)),
                     '待平台复核' if k % 3 == 0 else ('已处理' if k % 3 == 1 else '已驳回')])

# ============ 门店销售汇总 ============
sales = []
for si, s in enumerate(STORES):
    sid = s[0]
    for m in range(1, 7):
        recs = [p for p in plan if p['s'][0] == sid and p['m'] == m]
        ded = r2(sum(p['deduct'] for p in recs))
        orders = random.randint(380, 720)
        gross = r2(orders * random.uniform(78, 108))
        pay = r2(gross - ded)
        refund = r2(gross * random.uniform(0.008, 0.02))
        if (si * 6 + m) % 7 == 0:
            book = r2(gross - refund)
        else:
            book = r2(pay - refund)
        sales.append([sid, f'2026-{m:02d}', orders, gross, ded, pay, refund, book, '门店月报'])

# ============ 冲正台账 ============
ADJ_TYPES = [
    ('重复计提冲回', '同一券码重复计提，门店自行冲回'),
    ('过期券冲回', '券已过有效期未核销，冲回已计提'),
    ('跨店核销调整', '券限定门店与实际核销门店不一致'),
    ('收入冲减-券抵扣', '将券抵扣作为销售折让冲减营业收入'),
    ('平台争议调整', '平台争议款项暂估调整'),
]
adjust = []
for k in range(60):
    typ, reason = ADJ_TYPES[k % 5]
    sid = STORES[(k * 7) % len(STORES)][0]
    m = k % 6 + 1
    adjust.append([f'CZ{k + 1:04d}', f'2026-{m:02d}-{random.randint(2, 27):02d}', sid, typ,
                   r2(random.uniform(300, 12000)), reason,
                   '财务经理' if k % 2 else '门店店长'])

# ============ 写工作簿 ============
wb = openpyxl.Workbook()
wb.remove(wb.active)


def sheet(title, headers, rows, widths=None, banner=None, comment=None):
    ws = wb.create_sheet(title)
    if banner:
        ws.append([banner])
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
        c = ws.cell(row=1, column=1)
        c.font = RED_FONT
        c.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        ws.row_dimensions[1].height = 28
        ws.append([])
    hr = ws.max_row + 1
    ws.append(headers)
    for row in rows:
        ws.append(row)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=hr, column=c)
        cell.font = HDR_FONT
        cell.fill = HDR_FILL
        cell.alignment = CENTER
        cell.border = BORDER
        ws.column_dimensions[get_column_letter(c)].width = (widths or {}).get(c, 14)
    for row in ws.iter_rows(min_row=hr + 1, max_row=ws.max_row, max_col=len(headers)):
        for cell in row:
            cell.font = BODY
            cell.alignment = WRAP
            cell.border = BORDER
    ws.freeze_panes = ws.cell(row=hr + 1, column=1)
    if comment:
        ref, txt = comment
        cm = Comment(txt, '营销委员会')
        cm.width, cm.height = 330, 120
        ws[ref].comment = cm
    return ws


# 01 门店与活动主数据
sheet('门店主数据', ['门店编号', '城市', '区域', '商圈', '店长', '门店类型', '开业日期'],
      STORES, {1: 9, 2: 8, 3: 10, 4: 20, 5: 8, 6: 9, 7: 12})
sheet('活动主数据',
      ['活动编号', '活动名称', '门店范围', '平台', '券编号', '开始日', '结束日',
       '预算(元)', '审批人', '审批日期', '实际费用(元)', '状态'],
      activities, {1: 9, 2: 26, 3: 22, 4: 11, 5: 8, 6: 11, 7: 11, 8: 12, 9: 12, 10: 11, 11: 13, 12: 9})

# 02 订单核销明细
tx_rows = []
for p in plan:
    tx_rows.append([f"HX{p['code'][2:]}", p['date'], p['s'][0], p['c']['platform'], p['c']['id'], p['code'],
                    f"DD{p['code'][2:]}", p['orig'], p['pay'], p['deduct'], p['gift'],
                    f"M{random.randint(100000, 999999)}"])
tx_rows.sort(key=lambda x: (x[1], x[0]))
sheet('订单核销明细',
      ['核销单号', '核销日期', '门店编号', '平台', '券编号', '券码', '订单号',
       '订单原价(元)', '订单实收(元)', '券抵扣(元)', '赠品成本(元)', '会员ID'],
      tx_rows, {1: 15, 2: 12, 3: 9, 4: 11, 5: 8, 6: 15, 7: 17, 8: 12, 9: 12, 10: 11, 11: 11, 12: 10})

# 03 平台结算与争议
sheet('平台结算', ['门店编号', '月份', '券抵扣合计(元)', '平台应承担(元)', '平台实结算(元)',
                '差额(元)', '结算状态'],
      settle, {1: 9, 2: 9, 3: 16, 4: 15, 5: 15, 6: 11, 7: 10})
sheet('争议台账', ['争议编号', '发生日期', '门店编号', '所属月份', '平台', '争议类型',
                '争议金额(元)', '处理状态'],
      disputes, {1: 10, 2: 12, 3: 9, 4: 9, 5: 11, 6: 12, 7: 12, 8: 11})

# 04 营销费用台账
ledger_rows = [[x['no'], x['month'], x['sid'], x['act'], x['cid'], x['code'], x['amt'],
                x['basis'], x['status']] for x in ledger]
sheet('营销费用台账',
      ['计提编号', '计提月份', '门店编号', '活动编号', '券编号', '券码', '计提金额(元)',
       '计提依据', '状态'],
      ledger_rows, {1: 11, 2: 10, 3: 9, 4: 9, 5: 8, 6: 15, 7: 13, 8: 14, 9: 9},
      comment=('A1', '本台账为门店自行申报的计提数，未经营销部复核，不得直接作为应计营销费用依据；'
                     '同一券码可能被多次申报，已过期未核销券亦可能仍在台账中。'))

# 05 制度文件与会议纪要
sheet('营销费用管理制度v2.0', ['条款', '内容'],
      [['文件编号', 'PRC-MKT-v2.0'],
       ['适用范围', '2025-07-01 至 2026-04-30（含）'],
       ['营销费用定义', '因优惠券、折扣、买赠等促销活动由商家承担的让利金额'],
       ['平台补贴券承担比例', '美团50%、饿了么50%、点评50%、抖音团购70%、自有小程序100%'],
       ['折扣券承担比例', '商家全额承担，按订单原价减实收计算'],
       ['会员日折扣', '每月8日、18日，商家全额承担折扣额'],
       ['买赠', '按赠品成本价计入营销费用'],
       ['含税口径', '券面值、折扣额、赠品成本均为含税金额'],
       ['计提时点', '按券核销时点归属会计期间'],
       ['加盟门店', '加盟门店营销费用由加盟商自行承担，公司不承担'],
       ['收入确认', '按订单实收确认营业收入'],
       ['审批权限', '单场活动预算≥5万元须营销委员会审批，≥10万元须总经理审批']],
      {1: 20, 2: 86})

sheet('营销费用管理制度v3.0', ['条款', '内容'],
      [['文件编号', 'PRC-MKT-v3.0'],
       ['审议与生效', '2026-03-20 经营销委员会审议通过，自2026-05-01起执行'],
       ['适用范围', '2026-05-01 起发生的核销业务'],
       ['平台补贴券承担比例', '美团60%、饿了么60%、点评60%、抖音团购50%、自有小程序100%'],
       ['折扣券承担比例', '商家全额承担，按订单原价减实收计算'],
       ['会员日折扣', '每月8日、18日，商家全额承担折扣额'],
       ['买赠', '按赠品成本价计入营销费用'],
       ['含税口径', '券面值、折扣额、赠品成本均为含税金额'],
       ['计提时点', '按券核销时点归属会计期间'],
       ['券码去重', '同一券码只能计提一次，重复计提须冲回'],
       ['过期券', '未核销且已过有效期的券不得计提，已计提的须冲回'],
       ['加盟门店', '加盟门店营销费用由加盟商自行承担，公司不承担'],
       ['收入确认', '按订单实收确认营业收入，券抵扣不得重复冲减收入']],
      {1: 22, 2: 88})

sheet('制度生效与废止通知', ['文件编号', '发文日期', '生效日期', '当前状态', '说明'],
      [['PRC-MKT-v2.0', '2025-06-20', '2025-07-01', '2026-04-30 废止', '2026年1-4月营销费用核算依据'],
       ['PRC-MKT-v3.0', '2026-03-22', '2026-05-01', '现行有效', '2026-03-20审议通过，自2026-05-01起执行，不追溯调整1-4月'],
       ['MKT-2026-001', '2026-01-08', '2026-01-08', '已执行', '2026年营销费用专项审计口径会纪要'],
       ['WD-2026-003', '2026-03-25', '2026-03-25', '已发布', '文控归档通知，v3.0自2026-05-01起执行']],
      {1: 18, 2: 12, 3: 12, 4: 18, 5: 70},
      banner='【文控提示】本汇编已同步最新版 PRC-MKT-v3.0，请各门店及财务按 v3.0 口径统一核算。')

sheet('营销委员会会议纪要', ['日期', '议题', '决议内容', '责任人', '完成时限'],
      [['2026-01-08', '会议基本信息', '2026年营销费用专项审计口径会，参会：总经理、财务总监、营销总监、审计部、区域经理6人', '总经办', '已归档'],
       ['2026-01-08', '制度适用', '营销费用按券核销时点归属制度版本，1-4月适用v2.0、5-6月适用v3.0；不得按计提时点或发放时点归属，也不得统一套用最新版', '财务部', '2026-01-15'],
       ['2026-01-08', '台账性质', '营销费用台账为门店申报数，含重复计提与过期券，须按券码去重、剔除过期未核销后方可作为应计基础', '审计部', '2026-01-31'],
       ['2026-01-08', '含税口径', '券面值含税，营销费用按不含税列示，增值税率6%，应计金额＝含税金额÷1.06', '财务部', '2026-02-15'],
       ['2026-01-08', '加盟门店', '加盟门店营销费用由加盟商承担，公司不承担，不得计入公司营销费用', '营销部', '2026-02-15'],
       ['2026-01-08', '收入确认', '营业收入按订单实收确认，券抵扣不得作为销售折让重复冲减收入', '财务部', '2026-02-15'],
       ['2026-01-08', '活动审批', '单场活动预算≥5万元须营销委员会审批，≥10万元须总经理审批；先执行后审批的须补审批并说明', '营销部', '2026-03-31'],
       ['2026-01-08', '平台争议', '平台少结算或重复扣款的，由营销部在次月结算中追补，不得直接冲减营销费用', '营销部', '2026-03-31'],
       ['2026-01-08', '跨店核销', '券限定门店与实际核销门店不一致的，费用由实际核销门店所在区域承担并追责', '运营部', '2026-03-31']],
      {1: 12, 2: 14, 3: 88, 4: 10, 5: 12})

# 06 券规则与价格
coupon_rows = [[c['id'], c['name'], c['typ'], c['platform'], c['label'], c['thresh'],
                c['gift'], c['start'], c['end'], c['scope'], c['issued']] for c in coupons]
sheet('券规则', ['券编号', '券名称', '券种', '平台', '面值/折扣', '使用门槛(元)', '赠品成本(元)',
              '有效期起', '有效期止', '适用门店', '发放数量'],
      coupon_rows, {1: 8, 2: 28, 3: 11, 4: 11, 5: 10, 6: 12, 7: 12, 8: 11, 9: 11, 10: 18, 11: 10},
      comment=('L1', '券面值、折扣与赠品成本均为含税口径；平台补贴券的商家承担比例按核销时点适用制度版本确定，'
                     '不得按发放时点或计提时点确定。'))

# 07 门店销售汇总与冲正
sheet('门店销售汇总', ['门店编号', '月份', '订单数', '原价合计(元)', '券抵扣合计(元)',
                  '实收合计(元)', '退款(元)', '账面净收入(元)', '数据来源'],
      sales, {1: 9, 2: 9, 3: 9, 4: 14, 5: 15, 6: 14, 7: 11, 8: 15, 9: 10},
      comment=('H1', '账面净收入为门店月报列报数，收入确认口径以制度与《营销委员会会议纪要》为准。'))
sheet('冲正台账', ['冲正编号', '日期', '门店编号', '冲正类型', '金额(元)', '原因', '审批人'],
      adjust, {1: 10, 2: 12, 3: 9, 4: 18, 5: 11, 6: 42, 7: 10})

# 08 财务口径与管理要求
sheet('财务口径说明', ['规则', '内容', '适用年度'],
      [['营销费用确认', '营销费用按券面值全额计入，不区分商家承担比例', '2025版'],
       ['制度版本', '统一按最新版制度核算，便于年度预算衔接', '2025版'],
       ['加盟门店', '加盟门店营销费用并入公司营销费用合计', '2025版'],
       ['收入确认', '营业收入按订单原价确认，券抵扣计入营销费用', '2025版'],
       ['含税口径', '营销费用按含税金额列示', '2025版'],
       ['活动审批', '单场活动预算≥3万元须区域经理审批', '2025版'],
       ['券码管理', '未要求按券码去重', '2025版'],
       ['过期券', '已计提券不因过期冲回', '2025版'],
       ['争议处理', '平台争议金额直接冲减营销费用', '2025版'],
       ['金额单位', '元，保留两位小数', '通用']],
      {1: 16, 2: 70, 3: 10},
      comment=('A1', '本表为2025版财务口径，2026年调整事项（制度版本归属、券码去重、过期券冲回、加盟剔除、'
                     '不含税列示、收入按实收确认）以《营销委员会会议纪要》MKT-2026-001 为准。'))

sheet('管理层要求', ['编号', '要求', '问题'],
      [['M01', '统一按最新版 v3.0 比例重算全部6个月，便于与年度预算衔接', 'v3.0自2026-05-01起执行，1-4月应适用v2.0'],
       ['M02', '营销费用按台账合计数确认，不再逐单复核', '台账含重复计提、过期券与漏计提，不能直接采信'],
       ['M03', '券面值即为营销费用，不必做不含税换算', '券面含税，应按6%换算为不含税'],
       ['M04', '加盟门店营销费用并入公司合计，口径统一', '加盟店由加盟商承担，公司不承担'],
       ['M05', '过期未核销券已发即计入费用，不冲回', '未核销券不得计提，已计提须冲回'],
       ['M06', '同一券码多次计提按累计数确认，不做事后调整', '同一券码只能计提一次，重复须冲回'],
       ['M07', '收入按原价确认，券抵扣作为销售折让冲减收入', '收入按实收确认，不得重复冲减'],
       ['M08', '平台少结算部分直接冲减营销费用', '应由营销部次月追补，不得冲减费用'],
       ['M09', '活动预算5万元以下无需审批，直接执行', '≥5万元须营销委员会审批，与金额门槛无关'],
       ['M10', '跨店核销属系统问题，费用仍由公司统一承担', '由实际核销门店所在区域承担并追责'],
       ['M11', '会计提月份归属制度版本，1-4月按计提月判断', '应按券核销时点归属版本'],
       ['M12', '台账与账面差异小于1%可不调整', '差异须逐条列示并调整，不得设重要性豁免']],
      {1: 8, 2: 58, 3: 46})

wb.save(OUT)
print('saved', OUT)
print('stores', len(STORES), 'coupons', len(coupons), 'tx', len(tx_rows),
      'ledger', len(ledger), 'settle', len(settle), 'disputes', len(disputes),
      'sales', len(sales), 'adjust', len(adjust), 'activities', len(activities))
print('total data rows', len(STORES) + len(activities) + len(tx_rows) + len(settle) + len(disputes)
      + len(ledger) + 12 + 13 + 5 + 9 + len(coupons) + len(sales) + len(adjust) + 10 + 12)
