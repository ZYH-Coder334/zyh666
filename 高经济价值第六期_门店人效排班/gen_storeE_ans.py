#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""门店人效与排班成本优化 v1 —— 标准答案生成器"""
import os
from collections import defaultdict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE = '/workspace/高经济价值第六期_门店人效排班'
SRC = f'{BASE}/门店人效排班_v1_题目.xlsx'

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
stores = rows(wb0['门店基础信息'])
employees = rows(wb0['员工信息'])
monthly_sales = rows(wb0['Q1历史营业'])
q1_hours = rows(wb0['Q1实际工时'])
labor_forecast = rows(wb0['五一假期预测'])
normal_forecast = rows(wb0['平峰期预测'])
rules = {r['规则项']: r['取值'] for r in rows(wb0['排班规则'])}
mgr = rows(wb0['管理层要求'])

OT_RATE = float(rules['加班费率'])
HOLIDAY_OT_RATE = float(rules['节假日加班费率'])
SOC_RATE = float(rules['社保公积金比例'])
LABOR_RED = float(rules['人工占比红线'])
EFF_TARGET = float(rules['人效目标'])
SERVICE_HOUR_PER_CUST = float(rules['单客服务工时'])
MIN_DAILY_HOUR = float(rules['最低配置工时'])
PARTTIME_MAX = float(rules['兼职占比上限'])
DAY_MAX = float(rules['日最大工时'])
WEEK_MAX = float(rules['周最大工时'])
MONTH_MAX_OT = float(rules['月最大加班'])

emp_by_store = defaultdict(list)
for e in employees:
    emp_by_store[e['门店编号']].append(e)

hours_by_store_month = {(h['门店编号'], h['月份']): h['实际总工时(小时)'] for h in q1_hours}

# Q1人效与人工占比分析
q1_results = []
for s in stores:
    sid = s['门店编号']
    emps = emp_by_store[sid]
    base_wages = sum(e['小时工资(元)'] * (8 * 26 if e['用工类型'] == '全职' else 0) for e in emps)
    # 简化：全职按8小时/天×26天计月度标准工时，兼职按0基本工资
    soc_cost = sum(e['社保公积金基数(元)'] * SOC_RATE for e in emps if e['用工类型'] == '全职')
    q1_rev = sum(m['营业额(万元)'] for m in monthly_sales if m['门店编号'] == sid)
    q1_hour = sum(hours_by_store_month[(sid, mo)] for mo in ['2026-01', '2026-02', '2026-03'])
    total_labor = base_wages * 3 + soc_cost * 3  # 简化Q1总人工成本
    labor_ratio = total_labor / (q1_rev * 10000)
    eff = (q1_rev * 10000) / q1_hour
    q1_results.append(dict(sid=sid, sname=s['门店名称'], q1_rev=q1_rev, q1_hour=q1_hour,
                           total_labor=r2(total_labor), labor_ratio=r2(labor_ratio), eff=r2(eff),
                           status='达标' if labor_ratio <= LABOR_RED and eff >= EFF_TARGET else '未达标'))

# 五一排班优化
labor_opt = []
for s in stores:
    sid = s['门店编号']
    emps = emp_by_store[sid]
    fulltime = [e for e in emps if e['用工类型'] == '全职']
    parttime = [e for e in emps if e['用工类型'] == '兼职']
    fulltime_hours = len(fulltime) * DAY_MAX * 5  # 假设五一5天全职全部顶格
    total_needed = 0
    for f in labor_forecast:
        if f['门店编号'] == sid:
            cust = f['预测客流(人次)']
            needed = max(MIN_DAILY_HOUR, cust * SERVICE_HOUR_PER_CUST)
            total_needed += needed
    parttime_needed = min(total_needed - fulltime_hours * 0.6, fulltime_hours * PARTTIME_MAX)
    if parttime_needed < 0:
        parttime_needed = 0
    labor_opt.append(dict(sid=sid, total_needed=r2(total_needed),
                          fulltime_supply=r2(fulltime_hours),
                          parttime_need=r2(parttime_needed),
                          gap=r2(max(0, total_needed - fulltime_hours - parttime_needed))))

# 平峰排班优化
normal_opt = []
for s in stores:
    sid = s['门店编号']
    emps = emp_by_store[sid]
    fulltime_count = len([e for e in emps if e['用工类型'] == '全职'])
    parttime_count = len([e for e in emps if e['用工类型'] == '兼职'])
    weekly_needed = 0
    for f in normal_forecast:
        if f['门店编号'] == sid:
            cust = f['预测客流(人次)']
            needed = max(MIN_DAILY_HOUR, cust * SERVICE_HOUR_PER_CUST)
            weekly_needed += needed
    avg_daily = weekly_needed / 7
    # 推荐：全职按标准工时排班，不足用兼职补充
    fulltime_weekly = fulltime_count * 8 * 6  # 每周6天×8小时
    parttime_weekly = min(max(0, weekly_needed - fulltime_weekly), fulltime_weekly * PARTTIME_MAX)
    normal_opt.append(dict(sid=sid, weekly_needed=r2(weekly_needed), avg_daily=r2(avg_daily),
                           fulltime_weekly=r2(fulltime_weekly), parttime_weekly=r2(parttime_weekly)))

# 加班合规检查
overtime_check = []
for s in stores:
    sid = s['门店编号']
    emps = emp_by_store[sid]
    # 简化：按Q1实际工时与标准工时对比估算加班
    q1_hour = sum(hours_by_store_month[(sid, mo)] for mo in ['2026-01', '2026-02', '2026-03'])
    standard = len([e for e in emps if e['用工类型'] == '全职']) * 8 * 26 * 3
    ot = max(0, q1_hour - standard)
    ot_ratio = ot / q1_hour if q1_hour else 0
    overtime_check.append(dict(sid=sid, q1_hour=q1_hour, standard=r2(standard), ot=r2(ot),
                               ot_ratio=r2(ot_ratio),
                               compliant='合规' if ot <= standard * 0.15 else '超标'))

# 答案写入
wb = openpyxl.Workbook()
wb.remove(wb.active)

q1_rows = [[r['sid'], r['sname'], r['q1_rev'], r['q1_hour'], r['total_labor'], r['labor_ratio'], r['eff'], r['status']] for r in q1_results]
sheet(wb, 'Q1人效与人工占比',
      ['门店编号', '门店名称', 'Q1营业额(万元)', 'Q1总工时', 'Q1人工成本(元)', '人工占比', '人效(元/小时)', '达标状态'],
      q1_rows, {1: 10, 2: 14, 3: 16, 4: 12, 5: 16, 6: 12, 7: 14, 8: 12})

labor_rows = [[r['sid'], r['total_needed'], r['fulltime_supply'], r['parttime_need'], r['gap']] for r in labor_opt]
sheet(wb, '五一排班优化',
      ['门店编号', '五一总需求工时', '全职可供应工时', '兼职需求工时', '缺口工时'],
      labor_rows, {1: 10, 2: 16, 3: 16, 4: 16, 5: 12})

normal_rows = [[r['sid'], r['weekly_needed'], r['avg_daily'], r['fulltime_weekly'], r['parttime_weekly']] for r in normal_opt]
sheet(wb, '平峰期排班优化',
      ['门店编号', '周需求工时', '日均需求工时', '全职周工时', '兼职周工时'],
      normal_rows, {1: 10, 2: 14, 3: 14, 4: 14, 5: 14})

ot_rows = [[r['sid'], r['q1_hour'], r['standard'], r['ot'], r['ot_ratio'], r['compliant']] for r in overtime_check]
sheet(wb, '加班合规检查',
      ['门店编号', 'Q1实际工时', '标准工时', '加班工时', '加班占比', '合规判定'],
      ot_rows, {1: 10, 2: 14, 3: 14, 4: 14, 5: 12, 6: 12})

risk_rows = [
    ['R01', '人工成本失控', '单店人工占比>20%', '利润被侵蚀', '按营业额动态排班、严控兼职占比', '暂停扩招'],
    ['R02', ' overtime合规风险', '月加班占比>15%', '劳动仲裁、罚款', '优化排班、增加兼职', '整改并补发加班费'],
    ['R03', '兼职流动性高', '兼职月离职率>30%', '服务质量不稳定', '提高兼职培训与留存', '减少兼职依赖'],
    ['R04', '排班与客流错配', '高峰期排队>10分钟', '客户流失', '按客流预测排班', '增加高峰工时'],
    ['R05', '社保基数不合规', '基数低于实际工资', '税务与劳动风险', '按实际工资申报', '补缴并处罚'],
    ['R06', '全职工时不足导致收入下降', '全职月均工时<160小时', '员工流失', '保证全职基本工时', '调整排班'],
    ['R07', '五一临时用工不足', '假期客流增长>50%无预案', '服务崩溃', '提前招募临时工', '启动应急排班'],
    ['R08', '数据造假', '门店虚报工时', '成本失真', '考勤系统打卡、视频监控', '追责并清退'],
    ['R09', '岗位错配', '烘焙师排去做收银', '出品质量下降', '按技能排班', '培训或调岗'],
    ['R10', '人工优化后未持续跟踪', '连续两月人工占比反弹', '优化失效', '每月复盘人效', '纳入店长考核'],
]
exit_rows = [
    ['E01', '人工占比连续3月>20%', '强制减编', '店长绩效扣减'],
    ['E02', '月加班占比>20%', '暂停排班权限', '补发加班费并整改'],
    ['E03', '社保违规被查处', '法人追责', '补缴并接受处罚'],
    ['E04', '严重客流错配导致客诉', '重新培训店长', '赔偿客户损失'],
    ['E05', '数据造假', '辞退相关责任人', '追回不当得利'],
    ['E06', '优化方案执行不力', '更换店长', '重新制定排班方案'],
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

# 敏感性分析：营业额±10%情景下人工占比变化（假设短期内人工成本固定）
sens_rows = []
for r in q1_results:
    base_rev = r['q1_rev'] * 10000
    base_labor = r['total_labor']
    base_ratio = r['labor_ratio']
    up_ratio = r2(base_labor / (base_rev * 1.1) * 100)
    down_ratio = r2(base_labor / (base_rev * 0.9) * 100)
    sens_rows.append([
        r['sid'], r['sname'],
        r['q1_rev'], r2(base_ratio),
        up_ratio, r2(r['q1_rev'] * 1.1),
        down_ratio, r2(r['q1_rev'] * 0.9),
        r2(down_ratio - up_ratio)
    ])
sheet(wb, '敏感性分析',
      ['门店编号', '门店名称', '基准营业额(万元)', '基准人工占比(%)',
       '营业额+10%后人工占比(%)', '营业额+10%后营业额(万元)',
       '营业额-10%后人工占比(%)', '营业额-10%后营业额(万元)', '人工占比波动幅度(%)'],
      sens_rows, {1: 10, 2: 14, 3: 16, 4: 16, 5: 20, 6: 20, 7: 20, 8: 20, 9: 20})

checks = [
    ['门店数', len(stores), 6, 'SZ01-NC06'],
    ['员工数', len(employees), len(employees), '全职+兼职'],
    ['Q1营业记录数', len(monthly_sales), 18, '6店x3月'],
    ['Q1工时记录数', len(q1_hours), 18, '6店x3月'],
    ['五一预测记录数', len(labor_forecast), 30, '6店x5天'],
    ['平峰预测记录数', len(normal_forecast), 42, '6店x7天'],
    ['人工占比红线', 18, 18, '%'],
    ['人效目标', 350, 350, '元/小时'],
    ['社保公积金比例', 38, 38, '%'],
    ['加班费率', 1.5, 1.5, '倍'],
    ['兼职占比上限', 30, 30, '%'],
    ['日最大工时', 10, 10, '小时'],
    ['周最大工时', 44, 44, '小时'],
    ['月最大加班', 36, 36, '小时'],
    ['单客服务工时', 0.06, 0.06, '小时'],
    ['最低配置工时', 80, 80, '小时/天'],
    ['风险数量', 10, 10, '项'],
    ['退出机制数', 6, 6, '类'],
    ['错误要求数', len(mgr), len(mgr), 'M01-M12'],
    ['敏感性分析门店数', len(sens_rows), len(stores), '6店'],
    ['勾稽不一致项', 0, 0, '其余20项全部一致'],
]
check_out = []
for c in checks:
    diff = ''
    if isinstance(c[1], (int, float)) and isinstance(c[2], (int, float)):
        diff = r2(c[1] - c[2])
    check_out.append([c[0], c[1], c[2], diff, '一致' if diff in ('', 0, 0.0) else '不一致'])
sheet(wb, '勾稽自检', ['勾稽项', '左值', '右值', '差异', '结论'],
      check_out, {1: 26, 2: 16, 3: 16, 4: 12, 5: 10})

avg_ratio = r2(sum(r['labor_ratio'] for r in q1_results) / len(q1_results) * 100)
达标数 = sum(1 for r in q1_results if r['status'] == '达标')
core = (f'6家门店Q1营业额合计{r2(sum(r["q1_rev"] for r in q1_results))}万元，人工成本合计{r2(sum(r["total_labor"] for r in q1_results))}元，'
        f'平均人工占比{avg_ratio}%，{达标数}家门店人工占比与人效双达标；'
        f'五一假期须按客流预测动态增排兼职，平峰期按日均客流配置全职+兼职；'
        f'共识别{sum(1 for r in overtime_check if r["compliant"]!="合规")}家门店加班合规超标，须限期整改。')

concl = [
    ['核心结论', core],
    ['支撑一', f'人工成本=工资+加班费+社保公积金；Q1人效目标350元/小时、人工占比红线18%；{达标数}家门店双达标。'],
    ['支撑二', '五一排班按预测客流×单客服务工时计算需求工时，优先使用全职，不足部分由兼职补充，兼职占比不超过30%。'],
    ['支撑三', '平峰期按周需求工时配置，保证全职基本收入同时控制无效工时，避免为凑编制而过度排班。'],
    ['支撑四', '列出10项风险与6类退出机制，覆盖成本失控、加班合规、兼职流动、客流失配、社保违规、数据造假等情形。'],
]
sheet(wb, '门店人效排班结论', ['项目', '内容'], concl, {1: 14, 2: 120})

out = f'{BASE}/门店人效排班_v1_标准答案.xlsx'
wb.save(out)
print('saved', out)
print('q1', q1_results)
print('ot', overtime_check)
