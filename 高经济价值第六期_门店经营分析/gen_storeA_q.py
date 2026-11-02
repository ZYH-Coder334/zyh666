#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""连锁门店Q1经营分析与亏损店处置 v1 —— 题目数据生成器"""
import os
from datetime import date
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE = '/workspace/高经济价值第六期_门店经营分析'
os.makedirs(BASE, exist_ok=True)

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


# 成本费率
FOOD = 0.36
LABOR = 0.21
UTIL = 0.03
MKT = 0.04
MGMT = 0.05
COMM = 0.18
TAX = 0.25

# 门店主数据：编号, 城市, 区域, 商圈, 面积㎡, 开业日, 店长, 类型,
# 稳定月营收(万元), 外卖占比, 租金占收入比, 建店投资(万元)
STORES = [
    ['SZ01', '苏州', '工业园区', '湖东邻里中心', 180, '2023-05-10', '张伟', '直营', 85, 0.30, 0.10, 235],
    ['SZ02', '苏州', '高新区', '狮山天街', 200, '2023-08-15', '李娜', '直营', 78, 0.32, 0.15, 220],
    ['SZ03', '苏州', '吴中区', '万象汇', 180, '2024-01-20', '王强', '直营', 70, 0.35, 0.15, 210],
    ['SZ04', '苏州', '相城区', '天虹百货', 170, '2024-05-08', '赵敏', '直营', 62, 0.38, 0.19, 205],
    ['WX01', '无锡', '梁溪区', '崇安寺', 175, '2023-03-18', '陈杰', '直营', 80, 0.31, 0.10, 225],
    ['WX02', '无锡', '滨湖区', '滨湖万达', 180, '2023-09-12', '刘洋', '直营', 72, 0.34, 0.15, 215],
    ['WX03', '无锡', '新吴区', '宝龙广场', 170, '2024-03-22', '孙丽', '直营', 64, 0.37, 0.18, 200],
    ['WX04', '无锡', '惠山区', '惠山万达', 180, '2024-07-30', '周涛', '直营', 56, 0.45, 0.20, 195],
    ['NB01', '宁波', '海曙区', '天一广场', 185, '2023-04-25', '吴静', '直营', 88, 0.28, 0.09, 240],
    ['NB02', '宁波', '鄞州区', '印象城', 175, '2023-10-08', '郑磊', '直营', 74, 0.33, 0.12, 220],
    ['NB03', '宁波', '江北区', '来福士', 165, '2024-02-14', '黄芳', '直营', 68, 0.36, 0.14, 210],
    ['NB04', '宁波', '镇海区', '吾悦广场', 160, '2024-06-18', '林峰', '直营', 58, 0.40, 0.17, 195],
    ['HZ01', '杭州', '上城区', '湖滨银泰', 180, '2025-11-15', '徐明', '直营', 95, 0.30, 0.11, 250],
    ['HZ02', '杭州', '上城区', '钱江新城来福士', 200, '2025-12-01', '朱琳', '直营', 82, 0.33, 0.12, 245],
    ['HZ04', '杭州', '西湖区', '城西银泰', 190, '2025-12-15', '何军', '直营', 70, 0.38, 0.19, 235],
    ['HF01', '合肥', '蜀山区', '政务区银泰', 180, '2025-11-20', '高飞', '直营', 72, 0.32, 0.15, 225],
    ['HF02', '合肥', '包河区', '滨湖万达', 170, '2025-12-10', '罗敏', '直营', 66, 0.35, 0.15, 215],
    ['HF03', '合肥', '蜀山区', '三里庵国购', 160, '2025-12-20', '梁宇', '加盟', 60, 0.45, 0.12, 210],
    ['NC01', '南昌', '红谷滩区', '红谷滩万达', 160, '2025-12-05', '宋佳', '直营', 58, 0.36, 0.15, 205],
    ['NC03', '南昌', '青山湖区', '青山湖万象汇', 150, '2025-12-25', '谢峰', '加盟', 48, 0.42, 0.22, 200],
]

# 租约补充：月租金(元/㎡), 物业费(元/㎡/月), 免租期剩余(月), 年递增%, 押金月数, 合同到期
LEASE = {
    'SZ01': [472, 25, 0, 5, 3, '2028-05-09'],
    'SZ02': [585, 28, 0, 5, 3, '2028-08-14'],
    'SZ03': [583, 24, 0, 5, 3, '2029-01-19'],
    'SZ04': [693, 22, 0, 6, 3, '2029-05-07'],
    'WX01': [457, 26, 0, 5, 3, '2028-03-17'],
    'WX02': [600, 23, 0, 5, 3, '2028-09-11'],
    'WX03': [678, 20, 0, 5, 3, '2029-03-21'],
    'WX04': [622, 18, 0, 5, 3, '2029-07-29'],
    'NB01': [428, 30, 0, 5, 3, '2028-04-24'],
    'NB02': [507, 27, 0, 5, 3, '2028-10-07'],
    'NB03': [576, 22, 0, 5, 3, '2029-02-13'],
    'NB04': [616, 18, 0, 5, 3, '2029-06-17'],
    'HZ01': [581, 32, 0, 5, 3, '2030-11-14'],
    'HZ02': [492, 30, 0, 5, 3, '2030-11-30'],
    'HZ04': [700, 25, 0, 5, 3, '2030-12-14'],
    'HF01': [600, 18, 0, 5, 3, '2030-11-19'],
    'HF02': [582, 16, 0, 5, 3, '2030-12-09'],
    'HF03': [450, 15, 0, 5, 3, '2030-12-19'],
    'NC01': [544, 14, 0, 5, 3, '2030-12-04'],
    'NC03': [704, 12, 0, 5, 3, '2030-12-24'],
}

# Q1月度系数（春节扰动）
MONTH_FACTOR = {1: 0.93, 2: 0.95, 3: 1.04}

# 12月开业店爬坡系数
def ramp(sid):
    return 0.95 if sid in ('HZ02', 'HZ04', 'HF03', 'NC03') else 1.0


def monthly_rev(rev, deliv, m, sid):
    return rev * MONTH_FACTOR[m] * ramp(sid)


wb = openpyxl.Workbook()
wb.remove(wb.active)

# ---- 01 门店基础与组织架构 ----
store_rows = []
for s in STORES:
    sid, city, dist, mall, area, op, mgr, typ, rev, deliv, rr, capex = s
    l = LEASE[sid]
    store_rows.append([sid, city, dist, mall, area, op, mgr, typ, rev, deliv, rr, capex,
                       l[0], l[1], l[2], l[3], l[4], l[5]])
sheet(wb, '门店基础信息',
      ['门店编号', '城市', '区域', '商圈', '面积(㎡)', '开业日', '店长', '门店类型',
       '稳定月营收(万元)', '外卖占比', '租金占收入比', '建店投资(万元)',
       '月租金(元/㎡)', '物业费(元/㎡/月)', '免租期剩余(月)', '年递增(%)', '押金月数', '合同到期日'],
      store_rows, {1: 10, 2: 8, 3: 10, 4: 20, 5: 10, 6: 12, 7: 8, 8: 10, 9: 16, 10: 10, 11: 14, 12: 16, 13: 14, 14: 16, 15: 14, 16: 10, 17: 8, 18: 12})

sheet(wb, '城市与区域',
      ['城市', '门店数', '直营数', '加盟数', '区域经理', '备注'],
      [['苏州', 4, 4, 0, '陈建华', '老市场，稳定增长'],
       ['无锡', 4, 4, 0, '陈建华', '惠山万达客流下滑'],
       ['宁波', 4, 4, 0, '王海', '镇海吾悦损耗偏高'],
       ['杭州', 3, 3, 0, '林志远', '新店爬坡中'],
       ['合肥', 3, 2, 1, '林志远', 'HF03为加盟试点'],
       ['南昌', 2, 1, 1, '林志远', 'NC03爬坡不及预期']],
      {1: 8, 2: 8, 3: 8, 4: 8, 5: 10, 6: 30})

# ---- 02 门店Q1 P&L（分月）----
pl_rows = []
for s in STORES:
    sid, city, dist, mall, area, op, mgr, typ, rev, deliv, rr, capex = s
    l = LEASE[sid]
    monthly_rent = (l[0] + l[1]) * area
    dep = capex * 10000 / 60
    for m in (1, 2, 3):
        r = monthly_rev(rev, deliv, m, sid)  # 万元
        food = r * FOOD
        labor = r * LABOR
        util = r * UTIL
        mkt = r * MKT
        comm = r * deliv * COMM
        mgmt = r * MGMT
        rent = monthly_rent / 10000
        ebitda = r - food - labor - util - mkt - comm - mgmt - rent
        ebit = ebitda - dep / 10000
        tax = max(0.0, ebit) * TAX
        net = ebit - tax
        pl_rows.append([sid, city, f'2026-{m:02d}', round(r, 2), round(food, 2), round(labor, 2),
                        round(rent, 2), round(util, 2), round(mkt, 2), round(comm, 2),
                        round(mgmt, 2), round(dep / 10000, 2), round(ebitda, 2),
                        round(ebit, 2), round(tax, 2), round(net, 2)])
sheet(wb, '门店Q1分月P&L',
      ['门店编号', '城市', '月份', '营业收入(万元)', '食材成本(万元)', '人工成本(万元)',
       '租金(万元)', '水电(万元)', '营销(万元)', '外卖佣金(万元)', '管理分摊(万元)',
       '折旧(万元)', 'EBITDA(万元)', '税前利润(万元)', '所得税(万元)', '税后利润(万元)'],
      pl_rows, {1: 10, 2: 8, 3: 10, 4: 16, 5: 14, 6: 14, 7: 12, 8: 12, 9: 12, 10: 14, 11: 14, 12: 12, 13: 14, 14: 14, 15: 12, 16: 14})

# ---- 03 POS流水与外卖结算 ----
pos_rows = []
for s in STORES:
    sid, city, dist, mall, area, op, mgr, typ, rev, deliv, rr, capex = s
    for m in (1, 2, 3):
        r = monthly_rev(rev, deliv, m, sid)
        dine = r * (1 - deliv)
        delivery_gmv = r * deliv
        dine_refund = dine * 0.01
        dine_net = dine - dine_refund
        pos_rows.append([sid, city, f'2026-{m:02d}', round(dine, 2), round(dine_refund, 2),
                         round(dine_net, 2), round(delivery_gmv, 2)])
sheet(wb, 'POS堂食流水',
      ['门店编号', '城市', '月份', '堂食应收(万元)', '堂食退款(万元)', '堂食净额(万元)', '外卖GMV(万元)'],
      pos_rows, {1: 10, 2: 8, 3: 10, 4: 16, 5: 14, 6: 16, 7: 16})

# 外卖结算单：门店、月份、平台、外卖GMV、佣金率、平台佣金、配送费率、配送费、退款率、退款、平台补贴、平台实收、门店入账
# 故意在3家门店设置入账差异
ERR_DELIVERY = {'SZ04', 'WX04', 'NC03'}
delivery_rows = []
for s in STORES:
    sid, city, dist, mall, area, op, mgr, typ, rev, deliv, rr, capex = s
    for m in (1, 2, 3):
        r = monthly_rev(rev, deliv, m, sid)
        gmv = r * deliv
        comm_rate = COMM
        comm = gmv * comm_rate
        ship_rate = 0.05
        ship = gmv * ship_rate
        ref_rate = 0.02
        refund = gmv * ref_rate
        subsidy = gmv * 0.03
        settled = gmv - comm - ship - refund + subsidy
        if sid in ERR_DELIVERY:
            if sid == 'SZ04':
                recorded = settled  # 错误：按实收入账，少计收入
            elif sid == 'WX04':
                recorded = gmv - subsidy  # 错误：平台补贴未计入收入
            else:  # NC03
                recorded = gmv + refund  # 错误：退款未冲减
        else:
            recorded = gmv
        delivery_rows.append([sid, f'2026-{m:02d}', '美团' if m % 2 else '饿了么',
                              round(gmv, 2), round(comm_rate * 100, 1), round(comm, 2),
                              round(ship_rate * 100, 1), round(ship, 2), round(ref_rate * 100, 1),
                              round(refund, 2), round(subsidy, 2), round(settled, 2), round(recorded, 2)])
sheet(wb, '外卖平台结算单',
      ['门店编号', '月份', '平台', '外卖GMV(万元)', '佣金率(%)', '平台佣金(万元)',
       '配送费率(%)', '配送费(万元)', '退款率(%)', '退款(万元)', '平台补贴(万元)',
       '平台实收(万元)', '门店入账(万元)'],
      delivery_rows, {1: 10, 2: 10, 3: 8, 4: 16, 5: 10, 6: 14, 7: 10, 8: 12, 9: 10, 10: 12, 11: 14, 12: 14, 13: 14})

# ---- 04 排班工时与工资 ----
schedule_rows = []
for s in STORES:
    sid, city, dist, mall, area, op, mgr, typ, rev, deliv, rr, capex = s
    for m in (1, 2, 3):
        r = monthly_rev(rev, deliv, m, sid)
        labor_wan = r * LABOR  # 万元
        # 拆分岗位：店长、烘焙师、服务员、兼职
        total_labor_yuan = labor_wan * 10000
        mgr_pay = total_labor_yuan * 0.18
        baker_pay = total_labor_yuan * 0.40
        svc_pay = total_labor_yuan * 0.30
        part_pay = total_labor_yuan * 0.12
        social = total_labor_yuan * 0.28
        for role, pay in [('店长', mgr_pay), ('烘焙师', baker_pay), ('服务员', svc_pay), ('兼职', part_pay)]:
            hours = pay / 35 if role == '兼职' else pay / 40
            schedule_rows.append([sid, f'2026-{m:02d}', role, round(hours, 1),
                                  40 if role != '兼职' else 35, round(pay, 2),
                                  round(pay * 0.28, 2), round(pay * 1.28, 2)])
        # 社保汇总行
        schedule_rows.append([sid, f'2026-{m:02d}', '社保公积金合计', '', '', '', round(social, 2), round(social, 2)])
sheet(wb, '排班工时与工资',
      ['门店编号', '月份', '岗位', '排班工时', '时薪(元)', '应发工资(元)', '社保公积金(元)', '用工成本(元)'],
      schedule_rows, {1: 10, 2: 10, 3: 12, 4: 10, 5: 10, 6: 14, 7: 14, 8: 14})

# ---- 05 租约与免租期 ----
lease_rows = []
for s in STORES:
    sid, city, dist, mall, area, op, mgr, typ, rev, deliv, rr, capex = s
    l = LEASE[sid]
    lease_rows.append([sid, city, mall, area, l[0], l[1], l[2], l[3], l[4], l[5],
                       round((l[0] + l[1]) * area, 2), '按租约执行'])
sheet(wb, '租约明细',
      ['门店编号', '城市', '商圈', '面积(㎡)', '月租金(元/㎡)', '物业费(元/㎡/月)',
       '免租期剩余(月)', '年递增(%)', '押金月数', '合同到期日', '月租金含物业(元)', '备注'],
      lease_rows, {1: 10, 2: 8, 3: 20, 4: 10, 5: 12, 6: 14, 7: 12, 8: 10, 9: 8, 10: 12, 11: 18, 12: 16})

# ---- 06 盘点损耗与总部分摊 ----
# 盘点：账面库存、实盘、差异金额、差异率、责任
# 2家超标：WX04, NB04
INVENTORY_ERR = {'WX04': 0.045, 'NB04': 0.038}  # 损耗率
inv_rows = []
for s in STORES:
    sid, city, dist, mall, area, op, mgr, typ, rev, deliv, rr, capex = s
    for m in (1, 2, 3):
        r = monthly_rev(rev, deliv, m, sid)
        food_wan = r * FOOD
        book = food_wan * 0.35  # 期末账面库存约为食材成本的35%
        if sid in INVENTORY_ERR:
            loss_rate = INVENTORY_ERR[sid]
        else:
            loss_rate = 0.012
        actual = book * (1 - loss_rate)
        diff = book - actual
        diff_rate = diff / book if book else 0
        inv_rows.append([sid, f'2026-{m:02d}', round(book, 2), round(actual, 2),
                         round(diff, 2), round(diff_rate * 100, 2)])
sheet(wb, '月度盘点',
      ['门店编号', '月份', '账面库存(万元)', '实盘库存(万元)', '差异金额(万元)', '差异率(%)'],
      inv_rows, {1: 10, 2: 10, 3: 16, 4: 16, 5: 14, 6: 12})

sheet(wb, '总部分摊规则',
      ['分摊项', '计提基数', '费率', '说明'],
      [['管理与中央厨房分摊', '门店营业收入', '6%', '含总部管理、中央厨房、品控、IT'],
       ['营销推广费', '门店营业收入', '4%', '门店自担，含平台满减与地推'],
       ['外卖平台佣金', '外卖GMV', '18%', '仅对外卖收入计提'],
       ['食材成本', '门店营业收入', '36%', '按标准配方与出成率'],
       ['人工成本', '门店营业收入', '21%', '含店长、烘焙、服务、兼职与社保'],
       ['水电能耗', '门店营业收入', '3%', '按营业收入计提'],
       ['折旧', '建店投资', '5年直线', '残值为0，月折旧=建店投资/60'],
       ['企业所得税', '正税前利润', '25%', '亏损年度所得税为0']],
      {1: 22, 2: 16, 3: 12, 4: 40})

# ---- 07 店长绩效考核办法 ----
sheet(wb, '店长绩效考核办法',
      ['考核项', '权重', '计算口径', '达标线', '奖金系数规则'],
      [['EBITDA达成率', 0.50, '实际EBITDA÷预算EBITDA', '≥100%', '达成率×0.5，上限1.5'],
       ['营收达成率', 0.30, '实际营收÷预算营收', '≥100%', '达成率×0.3，上限1.5'],
       ['损耗率达标', 0.20, '实际损耗率≤2%', '≤2%', '达标得满分，每超1个百分点扣0.1'],
       ['基础奖金', '', '门店类型与城市系数', 'A级1.2万、B级1万、C级0.8万、D级0.5万', '季度基础奖金'],
       ['负激励', '', 'D级门店或连续2季亏损', 'EBITDA率<4%', '扣减50%基础奖金，追回已发预提'],
       ['发放规则', '', '季度终了次月发放80%，20%留存至年度', '留存部分年度考核后发放', '离职不予发放留存部分']],
      {1: 16, 2: 8, 3: 28, 4: 18, 5: 30})

# ---- 08 财务口径与管理层要求 ----
FIN_RULES = [
    ['收入确认', '营业收入=堂食净额+外卖GMV，外卖GMV含平台补贴，退款冲减收入', '用于P&L与奖金'],
    ['外卖佣金', '仅对外卖GMV按18%计提，不得对堂食收入计提', '佣金计入营业费用'],
    ['免租期摊销', '免租期内租金按直线法在整个租期内分摊，Q1无免租期剩余', '老店免租期已用完'],
    ['折旧口径', '建店投资按5年直线折旧，残值为0，月折旧=建店投资/60', '不含押金与流动资金'],
    ['盘点差异', '差异率≤2%为正常损耗，>2%由门店承担并从奖金扣除', '差异金额计入食材成本'],
    ['总部分摊', '管理费6%、营销费4%按营业收入计提，不得按门店数分摊', '固定费率'],
    ['所得税', '按正税前利润的25%计提，亏损年度所得税为0', '不得税前弥补以前年度亏损'],
    ['门店分级', 'A级EBITDA率≥12%，B级8%-12%，C级4%-8%，D级<4%或连续亏损', '用于奖金与处置'],
    ['处置触发', 'D级门店启动整改，连续2季D级启动关店或转址', 'C级门店限期整改'],
    ['奖金口径', '基础奖金×(EBITDA达成×0.5+营收达成×0.3+损耗达标×0.2)，D级扣减50%', '季度核算'],
    ['预算基准', 'Q1预算营收=稳定月营收×3，预算EBITDA按标准费率测算', '用于达成率'],
    ['加盟门店', '加盟门店P&L按门店口径核算，公司仅确认授权费与供应链毛利', '不并入直营利润'],
]
sheet(wb, '财务确认规则', ['规则', '内容', '用途'], FIN_RULES, {1: 18, 2: 60, 3: 28})

MGR = [
    ['M01', '为快速出报表，外卖佣金按全部营业收入18%计提', '佣金仅对外卖收入计提'],
    ['M02', '盘点差异统一按1%计入成本，不做门店责任划分', '差异率>2%须认定门店责任'],
    ['M03', '亏损门店店长奖金按0.8万基础奖金全额发放', 'D级门店扣减50%基础奖金'],
    ['M04', '租金占收入超过15%的门店直接关店', '应先整改，连续2季D级才关店'],
    ['M05', '新店Q1营收按稳定月营收×3直接确认', '12月开业店有爬坡系数'],
    ['M06', '外卖平台补贴不计入营业收入', '外卖GMV含平台补贴，应计入收入'],
    ['M07', '折旧按3年计提以加快回收', '制度规定5年直线折旧'],
    ['M08', '总部管理费按门店数平均分摊', '按营业收入6%计提'],
    ['M09', '门店退款不冲减收入，计入营销费用', '退款应冲减营业收入'],
    ['M10', '加盟门店利润全额并入公司利润', '加盟仅确认授权费与供应链毛利'],
    ['M11', '所得税按EBITDA的25%计提', '按正税前利润（EBITDA-折旧）的25%'],
    ['M12', 'D级门店全部立即关店，不做整改', '应先启动整改，连续2季D级才关店'],
]
sheet(wb, '管理层要求', ['编号', '要求', '问题'], MGR, {1: 8, 2: 56, 3: 34})

out = f'{BASE}/门店经营分析_v1_题目.xlsx'
wb.save(out)
print('saved', out)
print('stores', len(STORES), 'pl_rows', len(pl_rows))
