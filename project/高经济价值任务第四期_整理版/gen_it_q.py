#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""IT项目群后评价 - 输入附件生成（确定性分类）"""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from docx import Document
import zipfile, os, random, shutil
import numpy as np

random.seed(88)
plt.rcParams['font.sans-serif'] = ['WenQuanYi Micro Hei', 'SimHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

BASE = '/workspace/IT项目后评价_附件'
INPUT_DIR = os.path.join(BASE, '输入附件')
os.makedirs(INPUT_DIR, exist_ok=True)

hf = Font(name='微软雅黑', size=11, bold=True, color='FFFFFF')
hfill = PatternFill(start_color='2E75B6', end_color='2E75B6', fill_type='solid')
cf = Font(name='微软雅黑', size=10)
thin = Side(border_style='thin', color='000000')
bd = Border(left=thin, right=thin, top=thin, bottom=thin)
ctr = Alignment(horizontal='center', vertical='center', wrap_text=True)

def style(ws, nc):
    for c in ws[1]:
        c.font = hf; c.fill = hfill; c.alignment = ctr; c.border = bd
    for r in ws.iter_rows(min_row=2):
        for c in r:
            c.font = cf; c.border = bd; c.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

def mkzip(folder, name):
    zp = os.path.join(INPUT_DIR, name)
    with zipfile.ZipFile(zp, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(folder):
            for f in files:
                fp = os.path.join(root, f)
                zf.write(fp, os.path.relpath(fp, folder))
    shutil.rmtree(folder)
    print(f'  {name} ({os.path.getsize(zp)/1024:.1f}KB)')

# ========== 压缩包1: IT项目合同与预算 ==========
print('=== 01 IT项目合同与预算数据 ===')
d = os.path.join(INPUT_DIR, '_t1'); os.makedirs(d)

# 1.1 三年IT投资趋势PNG
fig, ax = plt.subplots(figsize=(10, 6))
years = ['2023年', '2024年', '2025年']
erp = [0.8, 0.6, 0.4]; crm = [0.3, 0.5, 0.2]; mes = [0.2, 0.4, 0.6]; data_mid = [0.1, 0.3, 0.5]; other = [0.2, 0.2, 0.3]
x = np.arange(3); w = 0.15
ax.bar(x-2*w, erp, w, label='ERP', color='#4472C4')
ax.bar(x-w, crm, w, label='CRM', color='#ED7D31')
ax.bar(x, mes, w, label='MES', color='#A5A5A5')
ax.bar(x+w, data_mid, w, label='数据中台', color='#FFC000')
ax.bar(x+2*w, other, w, label='其他', color='#5B9BD5')
ax.set_xlabel('年度'); ax.set_ylabel('投资额（亿元）'); ax.set_title('2023-2025年IT投资趋势')
ax.set_xticks(x); ax.set_xticklabels(years); ax.legend(fontsize=8); ax.set_ylim(0, 1.2)
for i in range(3):
    total = erp[i]+crm[i]+mes[i]+data_mid[i]+other[i]
    ax.text(x[i]+w*0.5, total+0.03, f'{total:.1f}亿', ha='center', fontsize=9)
plt.tight_layout(); fig.savefig(os.path.join(d, '三年IT投资趋势.png'), dpi=150); plt.close()

# 1.2 项目预算执行情况PNG
fig, ax = plt.subplots(figsize=(12, 6))
pnames = ['ERP升级', 'CRM系统', 'MES系统', '数据中台', 'OA升级', '信息安全', 'BI报表', '供应链平台', 'HR系统', '费控系统']
budgets = [3500, 2000, 4000, 3000, 800, 1500, 1200, 2800, 1000, 900]
actuals = [4200, 1600, 4800, 3200, 760, 1350, 1400, 2500, 950, 850]
xp = range(10)
ax.bar([i-0.2 for i in xp], budgets, 0.4, label='预算', color='#4472C4', alpha=0.7)
ax.bar([i+0.2 for i in xp], actuals, 0.4, label='实际', color='#ED7D31', alpha=0.7)
ax.set_ylabel('金额（万元）'); ax.set_title('2025年IT项目预算执行情况')
ax.set_xticks(xp); ax.set_xticklabels(pnames, rotation=30, ha='right'); ax.legend()
plt.tight_layout(); fig.savefig(os.path.join(d, '项目预算执行情况.png'), dpi=150); plt.close()

# 1.3 IT项目合同台账Excel（30个项目）
wb = openpyxl.Workbook(); ws = wb.active; ws.title = '项目合同台账'
ws.append(['序号', '项目编号', '项目名称', '项目类型', '合同金额(万元)', '预算金额(万元)', '实际支出(万元)', '供应商编号', '验收状态', '付款状态', '项目状态'])
# 30个项目: 8超预算 + 5未验收已付款 + 4功能未达标 + 3预算超支但ROI高 + 10正常
# 1-8: 超预算（实际>预算20%以上）
pnames_list = ['ERP核心升级', 'CRM客户管理', 'MES制造执行', '数据中台建设', 'OA协同升级',
               '信息安全体系', 'BI智能报表', '供应链平台', 'HR人力系统', '费控报销系统',
               '智能仓储WMS', '物联网IoT平台', '移动门户APP', '主数据MDM', '档案管理系统',
               '视频会议系统', '统一认证SSO', '日志分析平台', '容器云平台', 'API网关',
               'DevOps流水线', '数据治理平台', '智能客服系统', '电子签章系统', '培训学习平台',
               '资产管理系统', '采购协同平台', '质量追溯系统', '能耗监控平台', '安全运营中心']
for i in range(1, 31):
    pname = pnames_list[i-1]
    ptype = ['ERP', 'CRM', 'MES', '数据', 'OA', '安全', 'BI', '供应链', 'HR', '费控',
             'WMS', 'IoT', '移动', 'MDM', '档案', '视频', '安全', '日志', '云', 'API',
             'DevOps', '治理', '客服', '签章', '培训', '资产', '采购', '质量', '能耗', '安全'][i-1]
    budget = random.randint(500, 5000)
    contract = budget
    if i <= 8:
        actual = int(budget * random.uniform(1.25, 1.6))
        vstatus = '已验收'; pstatus = '已完成'; pstate = '超预算'
    elif i <= 13:
        actual = int(budget * random.uniform(0.8, 0.95))
        vstatus = '未验收'; pstatus = '已付款'; pstate = '未验收已付款'
    elif i <= 17:
        actual = int(budget * random.uniform(0.9, 1.05))
        vstatus = '已验收'; pstatus = '功能待整改'; pstate = '功能未达标'
    elif i <= 20:
        actual = int(budget * random.uniform(1.15, 1.3))
        vstatus = '已验收'; pstatus = '已完成'; pstate = '预算超支但ROI高'
    else:
        actual = int(budget * random.uniform(0.85, 1.1))
        vstatus = '已验收'; pstatus = '已完成'; pstate = '正常'
    sid = f'IT_SUP_{((i-1) % 25) + 1:03d}'
    ws.append([i, f'ITP{i:03d}', pname, ptype, contract, budget, actual, sid, vstatus, pstatus, pstate])
style(ws, 11); wb.save(os.path.join(d, '项目合同台账.xlsx'))

# 1.4 项目ROI数据Excel
wb = openpyxl.Workbook(); ws = wb.active; ws.title = '项目ROI'
ws.append(['项目编号', '项目名称', '投资额(万元)', '年化收益(万元)', 'ROI(%)', '回收期(年)', '收益类型', '备注'])
for i in range(1, 31):
    inv = random.randint(500, 5000)
    if i <= 20:
        roi_val = round(random.uniform(15, 45), 1)
        payback = round(random.uniform(2.5, 6.0), 1)
        rtype = '可量化'; note = ''
    else:
        roi_val = round(random.uniform(5, 15), 1)
        payback = round(random.uniform(6.0, 10.0), 1)
        rtype = '间接收益'; note = '品牌提升、效率改善等间接收益'
    if i <= 8:
        note = '预算超支但ROI仍达标'
    if i <= 20 and i > 17:
        note = '预算超支但ROI高于行业平均，属合理超支'
    annual = int(inv * roi_val / 100)
    ws.append([f'ITP{i:03d}', pnames_list[i-1], inv, annual, roi_val, payback, rtype, note])
style(ws, 8); wb.save(os.path.join(d, '项目ROI数据.xlsx'))

mkzip(d, '01IT项目合同与预算数据.rar')

# ========== 压缩包2: 系统交付与验收 ==========
print('=== 02 系统交付与验收数据 ===')
d = os.path.join(INPUT_DIR, '_t2'); os.makedirs(d)

# 2.1 系统功能完成度PNG
fig, ax = plt.subplots(figsize=(10, 6))
sys_names = ['ERP核心', 'CRM', 'MES', '数据中台', 'OA', '信息安全', 'BI', '供应链', 'HR', '费控']
planned = [100, 100, 100, 100, 100, 100, 100, 100, 100, 100]
actual_func = [95, 88, 72, 85, 98, 92, 68, 90, 96, 94]
x = range(10)
ax.bar([i-0.2 for i in x], planned, 0.4, label='计划完成度%', color='#4472C4', alpha=0.7)
ax.bar([i+0.2 for i in x], actual_func, 0.4, label='实际完成度%', color='#ED7D31', alpha=0.7)
ax.set_ylabel('功能完成度(%)'); ax.set_title('核心系统功能完成度对比')
ax.set_xticks(x); ax.set_xticklabels(sys_names, rotation=30, ha='right'); ax.legend(); ax.set_ylim(0, 120)
ax.axhline(y=90, color='r', linestyle='--', alpha=0.5, label='达标线90%')
for i in range(10):
    ax.text(i+0.2, actual_func[i]+2, f'{actual_func[i]}%', ha='center', fontsize=8)
plt.tight_layout(); fig.savefig(os.path.join(d, '系统功能完成度.png'), dpi=150); plt.close()

# 2.2 用户满意度PNG
fig, ax = plt.subplots(figsize=(10, 6))
sys_list = ['ERP', 'CRM', 'MES', '数据中台', 'OA', '安全', 'BI', '供应链', 'HR', '费控']
scores = [72, 65, 58, 78, 85, 80, 55, 70, 82, 76]
colors_s = ['#FF6B6B' if s < 60 else '#FFC000' if s < 70 else '#70AD47' for s in scores]
ax.bar(sys_list, scores, color=colors_s, width=0.6)
ax.set_ylabel('满意度评分'); ax.set_title('系统用户满意度评分')
ax.axhline(y=60, color='r', linestyle='--', alpha=0.5, label='及格线60')
ax.set_ylim(0, 100); ax.legend()
for i, s in enumerate(scores):
    ax.text(i, s+1, str(s), ha='center', fontsize=9)
plt.xticks(rotation=30, ha='right'); plt.tight_layout()
fig.savefig(os.path.join(d, '用户满意度评分.png'), dpi=150); plt.close()

# 2.3 系统验收明细Excel（40个系统）
wb = openpyxl.Workbook(); ws = wb.active; ws.title = '系统验收明细'
ws.append(['序号', '系统编号', '系统名称', '所属项目', '计划验收日期', '实际验收日期', '功能完成度(%)', '验收结论', '上线状态', '备注'])
sys_full = []
for i in range(1, 41):
    proj_idx = ((i-1) % 30) + 1
    sname = f'子系统{i}'
    pcode = f'ITP{proj_idx:03d}'
    func = random.randint(60, 100)
    if i <= 8:
        vconcl = '功能未达标'; ustatus = '已上线'; note = f'功能完成度{func}%低于90%达标线'
    elif i <= 13:
        vconcl = '未验收'; ustatus = '已上线'; note = '未完成验收即上线，违规'
    elif i <= 19:
        vconcl = '分期验收通过'; ustatus = '部分上线'; note = '分期交付，当前功能完成度合理'
    else:
        vconcl = '验收通过'; ustatus = '已上线'; note = ''
    ws.append([i, f'SYS{i:03d}', sname, pcode, f'2025-{random.randint(1,9):02d}-{random.randint(1,28):02d}',
               f'2025-{random.randint(1,9):02d}-{random.randint(1,28):02d}', func, vconcl, ustatus, note])
style(ws, 10); wb.save(os.path.join(d, '系统验收明细.xlsx'))

mkzip(d, '02系统交付与验收数据.rar')

# ========== 压缩包3: 供应商与招标 ==========
print('=== 03 供应商与招标数据 ===')
d = os.path.join(INPUT_DIR, '_t3'); os.makedirs(d)

# 3.1 招标方式分布PNG
fig, ax = plt.subplots(figsize=(8, 6))
labels = ['公开招标', '竞争性谈判', '单一来源', '邀请招标', '紧急采购']
sizes = [12, 6, 3, 6, 3]
colors = ['#4472C4', '#ED7D31', '#FFC000', '#5B9BD5', '#FF6B6B']
ax.pie(sizes, labels=labels, autopct='%1.0f%%', colors=colors, startangle=90)
ax.set_title('IT项目招标方式分布（30个项目）')
plt.tight_layout(); fig.savefig(os.path.join(d, '招标方式分布.png'), dpi=150); plt.close()

# 3.2 供应商名录Excel（25家）
wb = openpyxl.Workbook(); ws = wb.active; ws.title = '供应商名录'
ws.append(['序号', '供应商编号', '供应商名称', '服务领域', '年合同额(万元)', '合作年限', '招标方式', '供应商等级', '备注'])
# 25家: 5交付质量差 + 4价格虚高 + 3独家 + 3价格偏高但技术好 + 10正常
for i in range(1, 26):
    area = ['ERP', 'CRM', 'MES', '数据', '安全', '基础设施', '运维', '开发外包'][i % 8]
    amt = round(random.uniform(100, 2000), 2)
    years = random.randint(1, 10)
    if i <= 5:
        grade = 'D'; excl = '竞争性谈判'; note = '连续两期交付质量评分低于60分'
    elif i <= 9:
        grade = 'C'; excl = '邀请招标'; note = '合同单价高于市场均价20%以上'
    elif i <= 12:
        grade = 'B'; excl = '独家单一来源'; note = '独家供应，存在技术锁定风险'
    elif i <= 15:
        grade = 'A'; excl = '公开招标'; note = '单价偏高10%但技术评分95分以上，综合性价比高'
    else:
        grade = random.choice(['A', 'B', 'B']); excl = '公开招标'; note = ''
    ws.append([i, f'IT_SUP_{i:03d}', f'华{area}供应商{i}号', area, amt, years, excl, grade, note])
style(ws, 9); wb.save(os.path.join(d, '供应商名录.xlsx'))

# 3.3 供应商绩效评分Excel
wb = openpyxl.Workbook(); ws = wb.active; ws.title = '供应商绩效'
ws.append(['供应商编号', '交付质量(30分)', '进度(25分)', '价格(20分)', '服务(15分)', '技术(10分)', '总分(100分)', '评级'])
for i in range(1, 26):
    if i <= 5:
        q = random.randint(12, 18); sch = random.randint(10, 15); pr = random.randint(12, 16)
        sv = random.randint(6, 10); te = random.randint(5, 8)
    elif i <= 9:
        q = random.randint(22, 26); sch = random.randint(18, 22); pr = random.randint(10, 14)
        sv = random.randint(10, 13); te = random.randint(7, 9)
    elif i <= 12:
        q = random.randint(24, 28); sch = random.randint(20, 24); pr = random.randint(14, 17)
        sv = random.randint(11, 14); te = random.randint(6, 9)
    elif i <= 15:
        q = random.randint(27, 30); sch = random.randint(23, 25); pr = random.randint(13, 16)
        sv = random.randint(13, 15); te = random.randint(9, 10)
    else:
        q = random.randint(22, 28); sch = random.randint(18, 24); pr = random.randint(15, 19)
        sv = random.randint(10, 14); te = random.randint(7, 10)
    total = q + sch + pr + sv + te
    grade = 'A' if total >= 85 else 'B' if total >= 70 else 'C' if total >= 60 else 'D'
    ws.append([f'IT_SUP_{i:03d}', q, sch, pr, sv, te, total, grade])
style(ws, 8); wb.save(os.path.join(d, '供应商绩效评分表.xlsx'))

mkzip(d, '03供应商与招标数据.rar')

# ========== 压缩包4: 运维与安全 ==========
print('=== 04 运维与安全数据 ===')
d = os.path.join(INPUT_DIR, '_t4'); os.makedirs(d)

# 4.1 系统可用性PNG
fig, ax = plt.subplots(figsize=(10, 6))
sys_names = ['ERP', 'CRM', 'MES', '数据中台', 'OA', '安全', 'BI', '供应链']
avail = [99.5, 98.2, 97.1, 99.8, 99.9, 99.6, 96.5, 98.8]
colors_a = ['#FF6B6B' if a < 99 else '#70AD47' for a in avail]
ax.bar(sys_names, avail, color=colors_a, width=0.6)
ax.set_ylabel('可用性(%)'); ax.set_title('核心系统可用性')
ax.axhline(y=99, color='r', linestyle='--', alpha=0.5, label='SLA标准99%')
ax.set_ylim(95, 100.5); ax.legend()
for i, a in enumerate(avail):
    ax.text(i, a+0.1, f'{a}%', ha='center', fontsize=9)
plt.xticks(rotation=30, ha='right'); plt.tight_layout()
fig.savefig(os.path.join(d, '系统可用性.png'), dpi=150); plt.close()

# 4.2 安全事件统计PNG
fig, ax = plt.subplots(figsize=(10, 6))
months = [f'{m}月' for m in range(1, 13)]
leak = [0, 0, 1, 0, 0, 1, 0, 0, 1, 0, 0, 0]
auth = [0, 1, 0, 0, 1, 0, 0, 0, 0, 1, 0, 0]
vuln = [2, 3, 1, 4, 2, 3, 1, 2, 3, 1, 2, 3]
x = range(12)
ax.bar([i-0.2 for i in x], leak, 0.2, label='数据泄露', color='#FF6B6B')
ax.bar([i for i in x], auth, 0.2, label='权限违规', color='#FFC000')
ax.bar([i+0.2 for i in x], vuln, 0.2, label='漏洞事件', color='#5B9BD5')
ax.set_xlabel('月份'); ax.set_ylabel('事件数'); ax.set_title('2025年安全事件统计')
ax.set_xticks(x); ax.set_xticklabels(months); ax.legend()
plt.tight_layout(); fig.savefig(os.path.join(d, '安全事件统计.png'), dpi=150); plt.close()

# 4.3 运维工单明细Excel（50条）
wb = openpyxl.Workbook(); ws = wb.active; ws.title = '运维工单'
ws.append(['序号', '工单编号', '系统编号', '故障类型', '报障时间', '响应时间(分钟)', '解决时间(小时)', '处理人', '工单状态', '备注'])
# 50条: 10重复故障 + 8响应超时 + 5复杂升级合理 + 27正常
for i in range(1, 51):
    sid = f'SYS{random.randint(1, 40):03d}'
    ftype = random.choice(['系统故障', '性能问题', '功能异常', '数据错误', '权限问题'])
    rtime = random.randint(5, 30)
    stime = round(random.uniform(0.5, 8.0), 1)
    if i <= 10:
        status = '重复故障'; note = '同类型故障30天内重复出现3次以上'
        rtime = random.randint(5, 15)
    elif i <= 18:
        status = '响应超时'; note = f'响应时间{rtime}分钟超过SLA 15分钟标准'
        rtime = random.randint(20, 60)
    elif i <= 23:
        status = '复杂升级'; note = '响应超时但属复杂问题升级处理，三线支持合理超时'
        rtime = random.randint(30, 90)
    else:
        status = '正常'; note = ''
        rtime = random.randint(5, 15)
    ws.append([i, f'TK{i:04d}', sid, ftype,
               f'2025-{random.randint(1,9):02d}-{random.randint(1,28):02d} {random.randint(8,18):02d}:{random.randint(0,59):02d}',
               rtime, stime, f'运维工程师{random.randint(1,8)}', status, note])
style(ws, 10); wb.save(os.path.join(d, '运维工单明细.xlsx'))

# 4.4 安全事件台账Excel
wb = openpyxl.Workbook(); ws = wb.active; ws.title = '安全事件'
ws.append(['序号', '事件编号', '事件类型', '发生时间', '影响系统', '严重等级', '处置状态', '是否有审批', '备注'])
# 7条安全事件: 3数据泄露 + 2权限违规 + 2有审批合理
for i in range(1, 8):
    if i <= 3:
        etype = '数据泄露'; level = '严重'; appr = '否'
        note = f'未经审批的数据导出操作，需追责'
    elif i <= 5:
        etype = '权限违规'; level = '中等'; appr = '否'
        note = '越权访问敏感数据，需整改'
    else:
        etype = '权限调整'; level = '低'; appr = '是'
        note = '虽有触发告警但经安全委员会审批，属合理操作'
    ws.append([i, f'SEC{i:03d}', etype,
               f'2025-{random.randint(1,9):02d}-{random.randint(1,28):02d}',
               f'SYS{random.randint(1,40):03d}', level, '已处置', appr, note])
style(ws, 9); wb.save(os.path.join(d, '安全事件台账.xlsx'))

mkzip(d, '04运维与安全数据.rar')

# ========== 压缩包5: 管理制度 ==========
print('=== 05 IT管理制度 ===')
d = os.path.join(INPUT_DIR, '_t5'); os.makedirs(d)
policies = [
    {'name': '01IT项目管理办法.docx', 'title': 'IT项目管理办法',
     'secs': ['总则', '项目立项', '项目采购', '项目实施', '项目验收', '项目变更', '附则'],
     'issues': '问题：1.验收标准未量化；2.变更审批权限未按金额分级；3.紧急采购流程缺失'},
    {'name': '02IT供应商管理办法.docx', 'title': 'IT供应商管理办法',
     'secs': ['总则', '准入条件', '招标管理', '合同管理', '绩效考核', '退出机制', '附则'],
     'issues': '问题：1.单一来源采购无审批；2.绩效未与续约挂钩；3.退出机制不完善'},
    {'name': '03IT运维管理办法.docx', 'title': 'IT运维管理办法',
     'secs': ['总则', '运维组织', '工单管理', 'SLA管理', '变更管理', '应急管理', '附则'],
     'issues': '问题：1.SLA标准未按系统分级；2.重复故障无升级机制；3.运维外包无审批'},
    {'name': '04信息安全管理办法.docx', 'title': '信息安全管理办法',
     'secs': ['总则', '数据安全', '权限管理', '事件响应', '审计管理', '附则'],
     'issues': '问题：1.数据导出无审批；2.权限调整无审计；3.安全事件未分级'},
    {'name': '05IT投资管理办法.docx', 'title': 'IT投资管理办法',
     'secs': ['总则', '投资审批', '预算管理', '效益评价', '后评价', '附则'],
     'issues': '问题：1.投资审批未按金额分级；2.超预算无预警；3.后评价未制度化'},
]
for p in policies:
    doc = Document(); doc.add_heading(p['title'], level=0)
    for sec in p['secs']:
        doc.add_heading(sec, level=1)
        doc.add_paragraph(f'{sec}具体条款内容详见正式制度文件。')
    doc.add_heading('现有制度问题诊断', level=1)
    doc.add_paragraph(p['issues'])
    doc.save(os.path.join(d, p['name']))
mkzip(d, '05IT管理制度.rar')

# ========== 压缩包6: 管理层要求 ==========
print('=== 06 管理层要求 ===')
d = os.path.join(INPUT_DIR, '_t6'); os.makedirs(d)
doc = Document(); doc.add_heading('管理层要求汇总', level=0)
reqs = [
    (1, 'CEO', '超预算项目不追责继续执行', '追责可能影响团队积极性'),
    (2, 'CFO', '未验收项目继续付款不暂停', '暂停付款影响供应商关系'),
    (3, 'CIO', '独家供应商不引入竞争保持稳定', '切换成本高风险大'),
    (4, 'CEO', '安全事件不向董事会报告', '报告可能引起不必要的关注'),
    (5, 'CIO', '功能未达标项目不整改直接使用', '整改耗时影响业务'),
    (6, 'CFO', '运维外包不审批直接执行', '审批流程太慢'),
    (7, 'CIO', '数据迁移不走测试直接执行', '测试环境不具备'),
    (8, 'CEO', '系统上线不走验收流程', '验收流程太长影响上线'),
    (9, 'CIO', '供应商绩效不与续约挂钩', '绩效差的也维持原合同'),
    (10, 'CFO', 'IT预算偏差不分析直接通过', '分析需要时间和人力'),
]
for num, src, content, reason in reqs:
    doc.add_heading(f'要求{num}', level=1)
    p = doc.add_paragraph(); p.add_run('来源：').bold = True; p.add_run(src)
    p = doc.add_paragraph(); p.add_run('要求：').bold = True; p.add_run(content)
    p = doc.add_paragraph(); p.add_run('理由：').bold = True; p.add_run(reason)
    doc.add_paragraph()
doc.save(os.path.join(d, '管理层要求.docx'))

# CIO邮件PNG
fig, ax = plt.subplots(figsize=(10, 4)); ax.axis('off')
ax.text(0.05, 0.9, '发件人：CIO办公室', fontsize=11, transform=ax.transAxes)
ax.text(0.05, 0.75, '收件人：IT管理部全体', fontsize=11, transform=ax.transAxes)
ax.text(0.05, 0.6, '主题：关于IT项目后评价工作的指导意见', fontsize=11, transform=ax.transAxes)
ax.text(0.05, 0.3, '各位：\n经管理层讨论，IT项目后评价工作涉及面广，\n为稳妥起见，请暂缓执行超预算追责、未验收暂停付款、\n供应商引入竞争等措施，先维持现状，后续再议。\n\nCIO\n2025年9月20日', fontsize=10, transform=ax.transAxes)
rect = plt.Rectangle((0,0),1,1, transform=ax.transAxes, fill=False, edgecolor='gray')
ax.add_patch(rect); plt.tight_layout()
fig.savefig(os.path.join(d, 'CIO邮件.png'), dpi=150); plt.close()
mkzip(d, '06管理层要求.rar')

print('\n=== 输入附件生成完成 ===')
