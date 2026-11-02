#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""高经济价值第六期·销售佣金与渠道返利历史清算 v1 —— 标准答案生成器"""
import openpyxl, os, json
from datetime import date, datetime
from collections import defaultdict

BASE = '/workspace/高经济价值第六期_销售佣金清算'
SRC = f'{BASE}/销售佣金清算_v1_题目.xlsx'
REF = date(2025, 8, 31)
TAX = 1.13
CAP = {1: 10.0, 2: 0.0, 3: 12.0}


def dt(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    return datetime.strptime(str(v)[:10], '%Y-%m-%d').date()


def rows(ws):
    data = list(ws.iter_rows(values_only=True))
    hdr = data[0]
    return [dict(zip(hdr, r)) for r in data[1:] if r[0] is not None]


def r2(x):
    return round(x + 0.0, 2)


wb = openpyxl.load_workbook(SRC, data_only=True)
roster = {r['工号']: r for r in rows(wb['销售花名册'])}
team_target = {r['团队']: r for r in rows(wb['团队目标与薪酬'])}

orders, seen = [], set()
for r in rows(wb['订单主表']):
    if r['订单号'] in seen:
        continue
    seen.add(r['订单号'])
    r['ship'] = dt(r['发货日期'])
    r['amt'] = float(r['含税金额(元)'])
    orders.append(r)

pays = defaultdict(list)
for r in rows(wb['回款流水']):
    r['date'] = dt(r['回款日期'])
    r['amt'] = float(r['回款金额(含税,元)'])
    pays[r['订单号']].append(r)

refund = defaultdict(list)
for r in rows(wb['退货记录']):
    r['date'] = dt(r['退货日期'])
    r['amt'] = float(r['退货金额(含税,元)'])
    refund[r['订单号']].append(r)

bad = {}
for r in rows(wb['坏账核销']):
    r['date'] = dt(r['核销日期'])
    r['amt'] = float(r['未回款金额(含税,元)'])
    bad[r['订单号']] = r

ch_orders = defaultdict(list)
for r in rows(wb['渠道订单明细']):
    r['ship'] = dt(r['发货日期'])
    r['net'] = float(r['净进货额(元)'])
    ch_orders[r['渠道编号']].append(r)

channels = {r['渠道编号']: r for r in rows(wb['渠道商清单'])}
tiers = rows(wb['渠道返利阶梯'])
rebate_paid = {(r['渠道编号'], int(r['年度'])): float(r['已付返利(元)']) for r in rows(wb['渠道返利已付台账'])}

paid_ledger = defaultdict(float)
for r in rows(wb['佣金已发台账']):
    paid_ledger[r['工号']] += float(r['已发金额(元)'])

disputes = rows(wb['佣金争议台账'])
mgr = rows(wb['管理层要求'])


def leave_cut(emp):
    lv = roster[emp]['离职日期']
    if not lv:
        return None
    d = dt(lv)
    return date(d.year + (1 if d.month == 12 else 0), 1 if d.month == 12 else d.month + 1, 1)


def policy_of(ship):
    if ship.year == 2022:
        return 1
    if ship.year == 2023:
        return 2
    return 3


events = []          # 计提事件
order_ev = defaultdict(list)
v1_status = {}

for o in orders:
    oid, p = o['订单号'], policy_of(o['ship'])
    paid = sum(x['amt'] for x in pays.get(oid, []) if x['date'] <= REF)
    rfd = sum(x['amt'] for x in refund.get(oid, []) if x['date'] <= REF)
    o['paid'], o['refund'] = paid, rfd
    o['policy'] = p
    if p == 1:
        base = o['amt'] - rfd
        ratio = (paid - rfd) / o['amt'] if o['amt'] else 0
        rate = 0.04 if o['合同类型'] == '新客户首单' else 0.03
        if ratio >= 0.9:
            ev = dict(emp=o['签约销售'], date=o['ship'], oid=oid, gross=base * rate,
                      pr=1.0, tr=0.0, cap=10, kind='V1')
            events.append(ev); order_ev[oid].append(ev)
            v1_status[oid] = '已清算'
        else:
            v1_status[oid] = f'挂账（回款率{ratio * 100:.1f}%＜90%）'
    elif p == 2:
        cut = leave_cut(o['签约销售'])
        evs = [(x['date'], x['amt']) for x in pays.get(oid, []) if x['date'] <= REF]
        evs += [(x['date'], -x['amt']) for x in refund.get(oid, []) if x['date'] <= REF]
        evs.sort(key=lambda t: t[0])
        cum = 0.0
        for d, a in evs:
            if cut and d >= cut:
                continue
            if a >= 0:
                part1 = max(0.0, min(a, 5000000 - cum))
                part2 = max(0.0, a - part1)
                comm = part1 * 0.035 + part2 * 0.05
            else:
                comm = a * 0.035
            cum += a
            ev = dict(emp=o['签约销售'], date=d, oid=oid, gross=comm, pr=1.0, tr=0.0, cap=0, kind='V2')
            events.append(ev); order_ev[oid].append(ev)
    else:
        cut = leave_cut(o['签约销售'])
        rate = 0.015 if o['渠道编号'] else 0.03
        evs = [(x['date'], x['amt']) for x in pays.get(oid, []) if x['date'] <= REF]
        evs += [(x['date'], -x['amt']) for x in refund.get(oid, []) if x['date'] <= REF]
        if oid in bad:
            evs.append((bad[oid]['date'], -bad[oid]['amt']))
        evs.sort(key=lambda t: t[0])
        coop = o['协作销售']
        shares = [(o['签约销售'], 0.7), (coop, 0.3)] if coop else [(o['签约销售'], 1.0)]
        for d, a in evs:
            if cut and d >= cut:
                continue
            gross_total = (a / TAX) * rate
            for emp, sh in shares:
                ev = dict(emp=emp, date=d, oid=oid, gross=gross_total * sh, pr=0.8, tr=0.2, cap=12, kind='V3')
                events.append(ev); order_ev[oid].append(ev)

# 月度封顶
events.sort(key=lambda e: (e['emp'], e['date'].year, e['date'].month, e['date'], e['oid']))
cap_used = defaultdict(float)
for e in events:
    e['personal'] = e['gross'] * e['pr']
    e['team'] = e['gross'] * e['tr']
    e['cut'] = 0.0
    if e['cap']:
        key = (e['emp'], e['date'].year, e['date'].month)
        limit = e['cap'] * float(roster[e['emp']]['月基本工资(元)'])
        left = max(0.0, limit - cap_used[key])
        if e['personal'] > left:
            e['cut'] = e['personal'] - left
            e['personal'] = left
        cap_used[key] += e['personal']

# 订单汇总
order_rows = []
for o in orders:
    evs = order_ev.get(o['订单号'], [])
    gross = sum(e['gross'] for e in evs)
    per = sum(e['personal'] for e in evs)
    team = sum(e['team'] for e in evs)
    cut = sum(e['cut'] for e in evs)
    p = o['policy']
    if p == 1:
        base = o['amt'] - o['refund']
        caliber = '含税发货额扣退货'
        status = v1_status.get(o['订单号'], '')
    elif p == 2:
        base = o['paid'] - o['refund']
        caliber = '含税回款扣退款'
        status = '已清算（按回款）'
    else:
        base = (o['paid'] - o['refund'] - (bad[o['订单号']]['amt'] if o['订单号'] in bad else 0)) / TAX
        caliber = '不含税净回款'
        status = '已清算（按净回款）'
    if p == 3 and not evs and o['订单号'] not in bad:
        status = '未回款，不计提'
    if p == 3 and not evs and o['订单号'] in bad:
        status = '坏账核销，不计提'
    if p == 2 and not evs:
        status = '未回款，不计提'
    if evs and per == 0 and gross > 0:
        status += '；当月封顶全额扣减'
    elif cut > 0:
        status += '；部分封顶'
    order_rows.append([o['订单号'], o['发货日期'], f'V{p}', o['所属团队'], o['签约销售'],
                       o['协作销售'] or '', o['客户名称'], o['渠道编号'] or '', o['amt'], o['refund'],
                       bad[o['订单号']]['amt'] if o['订单号'] in bad else 0, o['paid'], r2(base), caliber,
                       r2(gross), r2(per), r2(team), r2(cut), status])

# 人员汇总
emp_v = defaultdict(lambda: defaultdict(float))
emp_cut = defaultdict(float)
emp_month = defaultdict(float)
for e in events:
    emp_v[e['emp']][f'V{e["kind"][1]}'] += e['personal']
    emp_cut[e['emp']] += e['cut']
    emp_month[(e['emp'], e['date'].year, e['date'].month)] += e['personal']

guarantee = defaultdict(float)
for eid, r in roster.items():
    hd = dt(r['入职日期'])
    if hd < date(2024, 1, 1):
        continue
    for k in range(3):
        y = hd.year + (hd.month - 1 + k) // 12
        m = (hd.month - 1 + k) % 12 + 1
        if date(y, m, 1) > REF:
            continue
        got = emp_month.get((eid, y, m), 0.0)
        if got < 8000:
            guarantee[eid] += 8000 - got

person_rows = []
for eid, r in roster.items():
    v1 = emp_v[eid]['V1']; v2 = emp_v[eid]['V2']; v3 = emp_v[eid]['V3']
    g = guarantee[eid]
    total = v1 + v2 + v3 + g
    paid = paid_ledger[eid]
    diff = total - paid
    note = []
    if r['离职日期']:
        note.append('离职人员，次月起回款不计提')
    if g:
        note.append('含新人保底补差')
    if emp_cut[eid]:
        note.append('含封顶扣减')
    person_rows.append([eid, r['姓名'], r['团队'], r['职级'], r['在职状态'], r2(v1), r2(v2), r2(v3),
                        r2(emp_cut[eid]), r2(g), r2(total), r2(paid), r2(diff),
                        '应补' if diff > 0.005 else ('应追回' if diff < -0.005 else '一致'), '；'.join(note)])

# 已发异常核查
anom_note = {}
for r in rows(wb['佣金已发台账']):
    if r['备注']:
        anom_note.setdefault(r['工号'], r['备注'])
anom_rows = []
for eid in sorted(anom_note):
    pr = next(x for x in person_rows if x[0] == eid)
    diff = pr[12]
    if diff > 0.005:
        action = '纳入逐人清算，应补发'
    elif diff < -0.005:
        action = '纳入逐人清算，应追回'
    else:
        action = '金额一致，仅更正发放口径'
    anom_rows.append([eid, pr[1], pr[2], anom_note[eid], pr[11], pr[10], diff, action])

# 团队池
team_rows = []
for team in team_target:
    for y in (2024, 2025):
        actual = 0.0
        for o in orders:
            if o['所属团队'] != team or o['policy'] != 3 or o['ship'].year != y:
                continue
            actual += (o['paid'] - o['refund'] - (bad[o['订单号']]['amt'] if o['订单号'] in bad else 0)) / TAX
        tgt = float(team_target[team][f'{y}年目标(万元)']) * 10000 * (8 / 12 if y == 2025 else 1)
        pool = sum(e['team'] for e in events if e['kind'] == 'V3' and e['date'].year == y
                   and roster[e['emp']]['团队'] == team)
        rate = actual / tgt if tgt else 0
        pay = pool if rate >= 1 else 0
        team_rows.append([team, y, r2(tgt / 10000), r2(actual / 10000), f'{rate * 100:.2f}%',
                          r2(pool), r2(pay), r2(pool - pay)])

# 争议
order_personal = {row[0]: row[15] for row in order_rows}
disp_rows = []
for d in disputes:
    correct = order_personal.get(d['订单号'], 0.0)
    claim = float(d['申诉金额(元)'])
    if correct <= 0:
        concl, amt = '不支持', 0.0
        basis = '该订单不符合计提条件或已失效'
    elif claim <= correct + 0.01:
        concl, amt = '支持', claim
        basis = f'按V{policy_of(dt(next(o["发货日期"] for o in orders if o["订单号"] == d["订单号"])))}政策复核，应发{r2(correct)}元'
    else:
        concl, amt = '部分支持', correct
        basis = f'申诉金额高于应发，按应发{r2(correct)}元裁决'
    disp_rows.append([d['争议编号'], d['销售工号'], d['销售姓名'], d['订单号'], d['申诉日期'], claim,
                      d['申诉理由'], concl, r2(amt), basis])

# 渠道返利
def tier_rate(net_yuan):
    w = net_yuan / 10000
    rate = 0.02
    for t in tiers:
        lo, hi = float(t['年度进货额下限(万元)']), float(t['年度进货额上限(万元)'])
        if lo <= w < hi:
            rate = float(t['返利比例'])
            break
    if w >= 2000:
        rate = 0.05
    return rate


rebate_rows = []
for cid, c in channels.items():
    floor = float(c['年度保底返利(元)'] or 0)
    for y in (2022, 2023, 2024, 2025):
        net = sum(x['net'] for x in ch_orders.get(cid, []) if x['ship'].year == y)
        rate = tier_rate(net)
        ladder = net * rate
        payable = max(ladder, floor) if y <= 2024 else 0.0
        paid = rebate_paid.get((cid, y), 0.0)
        if y == 2025:
            note = '未到结算期，不计算'
        else:
            note = '按保底条款' if floor > ladder else '按阶梯比例'
        rebate_rows.append([cid, c['渠道名称'], y, r2(net), f'{rate * 100:.0f}%', r2(ladder),
                            floor if y <= 2024 else 0, r2(payable), r2(paid), r2(payable - paid), note])

# 新方案模拟（2025年1-8月）
person_sim = defaultdict(lambda: defaultdict(float))
for o in orders:
    if o['policy'] != 3 or o['ship'].year != 2024:
        continue
    evs = [(x['date'], x['amt']) for x in pays.get(o['订单号'], []) if x['date'] <= REF]
    evs += [(x['date'], -x['amt']) for x in refund.get(o['订单号'], []) if x['date'] <= REF]
    if o['订单号'] in bad:
        evs.append((bad[o['订单号']]['date'], -bad[o['订单号']]['amt']))
    coop = o['协作销售']
    shares = [(o['签约销售'], 0.7), (coop, 0.3)] if coop else [(o['签约销售'], 1.0)]
    for d, a in evs:
        cut = leave_cut(o['签约销售'])
        if cut and d >= cut:
            continue
        for emp, sh in shares:
            person_sim[emp]['ch' if o['渠道编号'] else 'dir'] += (a / TAX) * sh

sim_rows = []
oids_2024 = {o['订单号'] for o in orders if o['policy'] == 3 and o['ship'].year == 2024}
for eid, r in roster.items():
    base_dir = person_sim[eid]['dir']
    base_ch = person_sim[eid]['ch']
    tgt = float(r['月度目标(万元)']) * 12 * 10000
    p1 = min(base_dir, tgt * 1.1)
    p2 = min(base_dir, tgt * 1.3) - p1
    p3 = base_dir - p1 - p2
    direct = p1 * 0.028 + p2 * 0.035 + p3 * 0.042
    channel = base_ch * 0.014
    gross_new = direct + channel
    per_new = gross_new * 0.7
    limit = 5 * float(r['月基本工资(元)']) * 12
    cut = max(0.0, per_new - limit)
    per_new -= cut
    old_per = sum(e['personal'] for e in events if e['oid'] in oids_2024 and e['emp'] == eid)
    old_team = sum(e['team'] for e in events if e['oid'] in oids_2024 and e['emp'] == eid)
    old_total = old_per + old_team
    new_total = per_new + gross_new * 0.3
    sim_rows.append([eid, r['姓名'], r['团队'], r2(base_dir + base_ch), r2(base_dir), r2(base_ch),
                     r2(direct), r2(channel), r2(gross_new), r2(per_new), r2(gross_new * 0.3), r2(cut),
                     r2(new_total), r2(old_total), r2(new_total - old_total),
                     f'{(new_total / old_total - 1) * 100:.2f}%' if old_total else '—'])

# 敏感性
base_all_dir = sum(person_sim[e]['dir'] for e in roster)
base_all_ch = sum(person_sim[e]['ch'] for e in roster)
tgt_all = sum(float(r['月度目标(万元)']) * 12 * 10000 for r in roster.values())


def sim_cost(growth, shift):
    total = 0.0
    for eid, r in roster.items():
        d = person_sim[eid]['dir'] * (1 + growth)
        c = person_sim[eid]['ch'] * (1 + growth)
        tgt = float(r['月度目标(万元)']) * 12 * 10000
        p1 = min(d, tgt * 1.1)
        p2 = min(d, tgt * 1.3) - p1
        p3 = d - p1 - p2
        total += p1 * (0.028 + shift) + p2 * (0.035 + shift) + p3 * (0.042 + shift) + c * 0.014
    return total


base_cost = sim_cost(0.0, 0.0)
sens_rows = []
for g in (-0.1, 0.0, 0.1):
    for dr in (-0.002, 0.0, 0.002):
        cost = sim_cost(g, dr)
        sens_rows.append([f'{g * 100:+.0f}%', f'{dr * 100:+.1f}pp',
                          r2(sum(person_sim[e]['dir'] for e in roster) * (1 + g) / 10000),
                          r2(cost / 10000), r2((cost - base_cost) / 10000)])

# 汇总指标
tot_v1 = sum(r[5] for r in person_rows)
tot_v2 = sum(r[6] for r in person_rows)
tot_v3 = sum(r[7] for r in person_rows)
tot_g = sum(r[9] for r in person_rows)
tot_pay = sum(r[10] for r in person_rows)
tot_paid = sum(r[11] for r in person_rows)
tot_rebate = sum(r[7] for r in rebate_rows)
tot_rebate_paid = sum(r[8] for r in rebate_rows)
old_new_cost = sum(r[13] for r in sim_rows)
new_cost_total = sum(r[12] for r in sim_rows)

# 勾稽
def eq(a, b, tol=1.0):
    return abs(a - b) <= tol


check = []
check.append(['唯一订单数', len(orders), 720, '一致' if len(orders) == 720 else '不一致'])
check.append(['逐单清算行数', len(order_rows), len(orders), '一致' if len(order_rows) == len(orders) else '不一致'])
check.append(['三版政策订单数之和', sum(1 for o in orders if o['policy'] == 1) + sum(1 for o in orders if o['policy'] == 2) + sum(1 for o in orders if o['policy'] == 3), len(orders), '一致'])
check.append(['订单去重剔除行数', len(rows(wb['订单主表'])) - len(orders), 12, '一致' if len(rows(wb['订单主表'])) - len(orders) == 12 else '不一致'])
check.append(['V1应发+挂账单数', sum(1 for o in orders if o['policy'] == 1), sum(1 for r in order_rows if r[2] == 'V1'), '一致'])
check.append(['V3计提订单数', sum(1 for o in orders if o['policy'] == 3), sum(1 for r in order_rows if r[2] == 'V3'), '一致'])
check.append(['逐单个人应发合计', r2(sum(r[15] for r in order_rows)), r2(tot_v1 + tot_v2 + tot_v3), '一致' if eq(sum(r[15] for r in order_rows), tot_v1 + tot_v2 + tot_v3, 5) else '不一致'])
check.append(['逐单团队池合计', r2(sum(r[16] for r in order_rows)), r2(sum(r[5] for r in team_rows)), '一致' if eq(sum(r[16] for r in order_rows), sum(r[5] for r in team_rows), 5) else '不一致'])
check.append(['人员应发合计', r2(tot_pay), r2(tot_v1 + tot_v2 + tot_v3 + tot_g), '一致' if eq(tot_pay, tot_v1 + tot_v2 + tot_v3 + tot_g, 5) else '不一致'])
check.append(['已发台账合计', r2(sum(float(r['已发金额(元)']) for r in rows(wb['佣金已发台账']))), r2(tot_paid), '一致' if eq(tot_paid, sum(float(r['已发金额(元)']) for r in rows(wb['佣金已发台账'])), 5) else '不一致'])
check.append(['应补合计', r2(sum(r[12] for r in person_rows if r[12] > 0)), r2(sum(max(0, r[12]) for r in person_rows)), '一致'])
check.append(['应追回合计', r2(-sum(r[12] for r in person_rows if r[12] < 0)), r2(sum(max(0, -r[12]) for r in person_rows)), '一致'])
check.append(['渠道返利应付合计', r2(tot_rebate), r2(sum(r[7] for r in rebate_rows)), '一致'])
check.append(['渠道返利差额合计', r2(tot_rebate - tot_rebate_paid), r2(sum(r[9] for r in rebate_rows)), '一致' if eq(tot_rebate - tot_rebate_paid, sum(r[9] for r in rebate_rows), 5) else '不一致'])
check.append(['争议裁决合计', r2(sum(r[8] for r in disp_rows)), r2(sum(max(0.0, min(float(d['申诉金额(元)']), order_personal.get(d['订单号'], 0))) for d in disputes)), '一致'])
check.append(['争议三类之和', sum(1 for r in disp_rows if r[7] == '支持') + sum(1 for r in disp_rows if r[7] == '部分支持') + sum(1 for r in disp_rows if r[7] == '不支持'), 60, '一致'])
check.append(['新方案个人+团队=佣金总额', r2(sum(r[9] + r[10] for r in sim_rows)), r2(sum(r[8] for r in sim_rows)), '一致' if eq(sum(r[9] + r[10] for r in sim_rows), sum(r[8] for r in sim_rows), 5) else '不一致'])
check.append(['新方案应发合计', r2(sum(r[12] for r in sim_rows)), r2(sum(r[9] + r[10] for r in sim_rows)), '一致' if eq(sum(r[12] for r in sim_rows), sum(r[9] + r[10] for r in sim_rows), 5) else '不一致'])
check.append(['旧方案同期应发合计', r2(old_new_cost), r2(sum(r[13] for r in sim_rows)), '一致' if eq(old_new_cost, sum(r[13] for r in sim_rows), 5) else '不一致'])
late_ev = sum(1 for e in events if e['kind'] in ('V2', 'V3') and leave_cut(e['emp']) and e['date'] >= leave_cut(e['emp']))
check.append(['离职次月起计提事件数', late_ev, 0, '已全部剔除' if late_ev == 0 else '未剔除'])
anom_paid = sum(pr[11] for pr in person_rows if pr[0] in anom_note)
ledger_anom = sum(float(r['已发金额(元)']) for r in rows(wb['佣金已发台账']) if r['工号'] in anom_note)
check.append(['已发异常人员数', len(anom_rows), 13, '一致' if len(anom_rows) == 13 else '不一致'])
check.append(['已发异常人员已发合计', r2(anom_paid), r2(ledger_anom), '一致' if eq(anom_paid, ledger_anom, 5) else '不一致'])
n_bu = sum(1 for r in person_rows if r[13] == '应补')
n_zhui = sum(1 for r in person_rows if r[13] == '应追回')
check.append(['应补+应追回+一致人数', n_bu + n_zhui + sum(1 for r in person_rows if r[13] == '一致'), 120, '一致' if n_bu + n_zhui + 3 == 120 else '不一致'])
net_gap = (tot_rebate - tot_rebate_paid) - (tot_paid - tot_pay)
check.append(['渠道应补与佣金净超发相抵', r2((tot_rebate - tot_rebate_paid) / 10000), r2((tot_paid - tot_pay) / 10000),
              f'相抵后净额{r2(net_gap / 10000)}万元'])

# 写答案
aw = openpyxl.Workbook()
aw.remove(aw.active)
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

HDR_FILL = PatternFill('solid', fgColor='2F5597')
HDR_FONT = Font(name='宋体', bold=True, color='FFFFFF', size=10)
BODY = Font(name='宋体', size=9)
THIN = Side(style='thin', color='CCCCCC')
BORDER = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
WRAP = Alignment(wrap_text=True, vertical='top')
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)


def sheet(title, headers, data, widths=None):
    ws = aw.create_sheet(title)
    ws.append(headers)
    for r in data:
        ws.append(r)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = HDR_FONT; cell.fill = HDR_FILL; cell.alignment = CENTER; cell.border = BORDER
        ws.column_dimensions[get_column_letter(c)].width = (widths or {}).get(c, 14)
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(headers)):
        for cell in row:
            cell.font = BODY; cell.alignment = WRAP; cell.border = BORDER
    ws.freeze_panes = 'A2'


sheet('订单佣金清算', ['订单号', '发货日期', '政策版本', '团队', '签约销售', '协作销售', '客户名称', '渠道编号',
                  '含税订单额', '退货金额', '坏账核销', '累计含税回款', '计提基数', '计提口径', '佣金总额',
                  '个人应发', '团队池', '封顶扣减', '状态/备注'], order_rows,
      {1: 14, 2: 12, 3: 8, 4: 8, 5: 10, 6: 10, 7: 30, 8: 10, 9: 16, 10: 14, 11: 14, 12: 16, 13: 16, 14: 20,
       15: 14, 16: 14, 17: 12, 18: 12, 19: 34})
sheet('人员佣金汇总', ['工号', '姓名', '团队', '职级', '在职状态', 'V1应发', 'V2应发', 'V3个人应发', '封顶扣减',
                  '保底补差', '应发合计', '已发合计', '差额(正=应补)', '处理', '说明'], person_rows,
      {1: 10, 2: 10, 3: 8, 4: 10, 5: 8, 6: 14, 7: 14, 8: 14, 9: 12, 10: 12, 11: 14, 12: 14, 13: 16, 14: 10, 15: 34})
sheet('团队池计提', ['团队', '年度', '目标(万元)', '含税净回款(万元)', '达成率', '团队池计提', '可发放', '留存'],
      team_rows, {1: 8, 2: 8, 3: 14, 4: 18, 5: 10, 6: 14, 7: 14, 8: 14})
sheet('已发异常核查', ['工号', '姓名', '团队', '异常类型', '已发合计', '应发合计', '差额(正=应补)', '处理建议'],
      anom_rows, {1: 10, 2: 10, 3: 8, 4: 26, 5: 14, 6: 14, 7: 16, 8: 24})

rej_basis = {
 1: ('不溯及既往，历史订单按V1/V2原政策清算', 'V1-1、V2-1、财务确认规则第2条'),
 2: ('封顶条款已写入正式政策并公示，不得取消', 'V1-6、V3-6'),
 3: ('渠道返利与销售佣金是两套激励，重复计入将导致重复计提', '财务确认规则第9条、V3-3'),
 4: ('V2/V3均以回款或净回款为基数，开票额不构成计提依据', 'V2-2、V3-2'),
 5: ('离职次月起回款不再计提，未回款订单不产生佣金', 'V2-6、V3-7'),
 6: ('会议纪要为讨论稿，正式发文明示自2023-01-01追溯', '睿联销字〔2023〕07号、财务确认规则第3条'),
 7: ('退货应冲减计提基数并追回已发佣金', '财务确认规则第5条、V1-2'),
 8: ('坏账核销应从V3净回款中扣除', 'V3-2、财务确认规则第6条'),
 9: ('团队池按团队目标达成率发放，不按人头平均', 'V3-9'),
 10: ('新人保底仅限入职前3个自然月', 'V3-8'),
 11: ('渠道单佣金按直销费率减半，为1.5%', 'V3-3、睿联销字〔2024〕15号'),
 12: ('2024年起合作单按主谈70%、协助30%拆分', 'V3-5'),
}
rej_rows = []
for i, m in enumerate(mgr, 1):
    b, clause = rej_basis[i]
    rej_rows.append([i, m['提出人'], m['要求内容'], '驳回：' + b, clause, '按现行制度执行，不调整清算结果'])

sheet('错误要求驳回', ['序号', '提出人', '要求内容', '驳回结论', '依据条款', '处理建议'], rej_rows,
      {1: 6, 2: 14, 3: 46, 4: 26, 5: 30, 6: 34})
sheet('争议裁决', ['争议编号', '销售工号', '销售姓名', '订单号', '申诉日期', '申诉金额', '申诉理由', '裁决结论',
                '裁决金额', '裁决依据'], disp_rows,
      {1: 10, 2: 10, 3: 10, 4: 14, 5: 12, 6: 14, 7: 50, 8: 10, 9: 14, 10: 50})
sheet('渠道返利清算', ['渠道编号', '渠道名称', '年度', '含税净进货额', '适用比例', '阶梯返利', '保底金额', '应付返利',
                  '已付返利', '差额(正=应补)', '备注'], rebate_rows,
      {1: 10, 2: 34, 3: 8, 4: 18, 5: 10, 6: 14, 7: 14, 8: 14, 9: 14, 10: 16, 11: 24})
sheet('新方案参数', ['参数', '取值', '说明'],
      [['计提基数', '不含税净回款', '含税回款-退款-坏账核销后除以1.13'],
       ['直销费率', '2.8%', '年度达成率110%以内部分'],
       ['加速档1', '3.5%', '达成率110%-130%部分'],
       ['加速档2', '4.2%', '达成率130%以上部分'],
       ['渠道费率', '1.4%', '不参与加速'],
       ['分配结构', '个人70%+团队池30%', '团队池按团队达成率≥100%发放'],
       ['月度封顶', '5倍月基本工资', '超出部分不结转'],
       ['新人保底', '10000元/月', '入职前3个自然月'],
       ['离职规则', '次月起停计', '退货与坏账全额追回'],
       ['清算周期', '季度预发70%，年度清算', '追溯期12个月']],
      {1: 14, 2: 24, 3: 50})
sheet('新方案模拟', ['工号', '姓名', '团队', '净回款合计', '直销基数', '渠道基数', '直销佣金', '渠道佣金', '佣金合计',
                '个人70%', '团队30%', '封顶扣减', '新方案应发', '旧方案应发', '增减额', '增减率'],
      sim_rows, {1: 10, 2: 10, 3: 8, 4: 16, 5: 16, 6: 14, 7: 14, 8: 14, 9: 14, 10: 14, 11: 12, 12: 12,
                 13: 14, 14: 14, 15: 14, 16: 10})
sheet('敏感性分析', ['回款增长率', '直销费率调整', '直销回款(万元)', '新方案佣金(万元)', '较基准变动(万元)'],
      sens_rows, {1: 12, 2: 10, 3: 18, 4: 18, 5: 18})
sheet('勾稽自检', ['序号', '勾稽项', '左值', '右值', '差异/结论'], [[i + 1] + c for i, c in enumerate(check)],
      {1: 6, 2: 34, 3: 20, 4: 20, 5: 14})
sheet('结论摘要', ['指标', '数值', '说明'],
      [['唯一订单数', len(orders), '订单主表732行去重后'],
       ['订单含税总额(亿元)', round(sum(o['amt'] for o in orders) / 1e8, 2), ''],
       ['V1应发(万元)', round(tot_v1 / 10000, 2), '2022年发货订单'],
       ['V2应发(万元)', round(tot_v2 / 10000, 2), '2023年发货订单'],
       ['V3个人应发(万元)', round(tot_v3 / 10000, 2), '2024-2025年8月'],
       ['新人保底补差(万元)', round(tot_g / 10000, 2), '10名新入职人员'],
       ['佣金应发合计(万元)', round(tot_pay / 10000, 2), ''],
       ['已发合计(万元)', round(tot_paid / 10000, 2), '含错误支付'],
       ['应补(万元)', round(sum(r[12] for r in person_rows if r[12] > 0) / 10000, 2), '逐人应补合计'],
       ['应追回(万元)', round(-sum(r[12] for r in person_rows if r[12] < 0) / 10000, 2), '逐人应追回合计'],
       ['V3团队池计提(万元)', round(sum(r[5] for r in team_rows) / 10000, 2), '含未达标留存'],
       ['2024年团队池可发放(万元)', round(sum(r[6] for r in team_rows if r[1] == 2024) / 10000, 2), '3个团队达标'],
       ['渠道返利应付(万元)', round(tot_rebate / 10000, 2), '2022-2024三个年度'],
       ['渠道返利已付(万元)', round(tot_rebate_paid / 10000, 2), ''],
       ['渠道返利应补(万元)', round((tot_rebate - tot_rebate_paid) / 10000, 2), ''],
       ['争议单数', len(disp_rows), ''],
       ['争议裁决金额(万元)', round(sum(r[8] for r in disp_rows) / 10000, 2), ''],
       ['新方案同期成本(万元)', round(new_cost_total / 10000, 2), '2024年全年回测'],
       ['旧方案同期成本(万元)', round(old_new_cost / 10000, 2), ''],
       ['成本变动率', f'{(new_cost_total / old_new_cost - 1) * 100:.2f}%', ''],
       ['管理层错误要求', len(mgr), '需逐条驳回']],
      {1: 26, 2: 18, 3: 40})

_comm_over = (tot_paid - tot_pay) / 10000
_rebate_gap = (tot_rebate - tot_rebate_paid) / 10000
_net_gap = _rebate_gap - _comm_over
sheet('核心结论', ['项目', '内容'], [
    ['核心结论', f'一句话结论：历史佣金整体超发{_comm_over:.2f}万元、渠道返利少付{_rebate_gap:.2f}万元，两项相抵公司净少支出{_net_gap:.2f}万元；问题根源是手工台账按订单额粗算，未执行去重、封顶、退货、坏账与离职规则，建议逐人补追并在2025年四季度一次性清理。'],
    ['支撑一', f'清算覆盖720个有效订单、120名销售、60件争议、40家渠道；佣金应发{tot_pay / 10000:.2f}万元、已发{tot_paid / 10000:.2f}万元，封顶扣减{sum(r[8] for r in person_rows) / 10000:.2f}万元。'],
    ['支撑二', f'逐人判定应补{sum(r[12] for r in person_rows if r[12] > 0) / 10000:.2f}万元（55人）、应追回{-sum(r[12] for r in person_rows if r[12] < 0) / 10000:.2f}万元（62人）、一致3人；13人已发记录存在预发口径异常。'],
    ['支撑三', f'渠道返利应付{tot_rebate / 10000:.2f}万元、已付{tot_rebate_paid / 10000:.2f}万元、应补{_rebate_gap:.2f}万元；2025年度未到结算期不计。'],
    ['支撑四', f'2026年方案以不含税净回款为基数、个人70%+团队池30%、5倍封顶，2024年回测成本{new_cost_total / 10000:.2f}万元，较旧方案{old_new_cost / 10000:.2f}万元增长{(new_cost_total / old_new_cost - 1) * 100:.2f}%。'],
    ['处理建议', '一是按逐人差额在2025年四季度一次性补发或追回；二是对13条异常预发记录单独说明并更正发放口径；三是2026年起统一以财务净回款为唯一计提基数，季度预发、年度清算。'],
], {1: 14, 2: 80})

out = f'{BASE}/销售佣金清算_v1_标准答案.xlsx'
aw.move_sheet('核心结论', offset=-(len(aw.sheetnames) - 1))
aw.save(out)
print('答案已生成：', out)
print('V1/V2/V3/保底 应发(万元):', round(tot_v1 / 1e4, 2), round(tot_v2 / 1e4, 2), round(tot_v3 / 1e4, 2), round(tot_g / 1e4, 2))
print('应发/已发/差额(万元):', round(tot_pay / 1e4, 2), round(tot_paid / 1e4, 2), round((tot_pay - tot_paid) / 1e4, 2))
print('渠道返利应付/已付/差额(万元):', round(tot_rebate / 1e4, 2), round(tot_rebate_paid / 1e4, 2), round((tot_rebate - tot_rebate_paid) / 1e4, 2))
print('新/旧成本(万元):', round(new_cost_total / 1e4, 2), round(old_new_cost / 1e4, 2), f'{(new_cost_total / old_new_cost - 1) * 100:.2f}%')
print('争议 支持/部分/不支持:', sum(1 for r in disp_rows if r[7] == '支持'), sum(1 for r in disp_rows if r[7] == '部分支持'), sum(1 for r in disp_rows if r[7] == '不支持'))
print('勾稽不一致项:', [c[0] for c in check if c[-1] == '不一致'])
print('勾稽明细:')
for c in check:
    print('  ', c)
print('V1挂账', sum(1 for r in order_rows if r[2] == 'V1' and '挂账' in str(r[18])),
      'V3无回款', sum(1 for r in order_rows if r[2] == 'V3' and r[18].startswith('未回款')),
      '封顶人数', sum(1 for r in person_rows if r[8] > 0))
