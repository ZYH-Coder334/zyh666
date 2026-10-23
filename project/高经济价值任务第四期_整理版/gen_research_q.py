#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""高校科研经费审计 - 输入附件生成"""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from docx import Document
import zipfile, os, random, shutil
import numpy as np

random.seed(66)
plt.rcParams['font.sans-serif'] = ['WenQuanYi Micro Hei', 'SimHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

BASE = '/workspace/科研经费审计_附件'
INPUT_DIR = os.path.join(BASE, '输入附件')
os.makedirs(INPUT_DIR, exist_ok=True)

hf = Font(name='微软雅黑', size=11, bold=True, color='FFFFFF')
hfill = PatternFill(start_color='7030A0', end_color='7030A0', fill_type='solid')
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

# ========== 压缩包1: 科研项目与经费预算 ==========
print('=== 01 科研项目与经费预算数据 ===')
d = os.path.join(INPUT_DIR, '_t1'); os.makedirs(d)

# 1.1 三年科研经费投入趋势PNG
fig, ax = plt.subplots(figsize=(10, 6))
years = ['2023年', '2024年', '2025年']
national = [0.8, 1.0, 1.2]; provincial = [0.3, 0.4, 0.5]; enterprise = [0.2, 0.3, 0.4]; univ = [0.15, 0.2, 0.25]
x = np.arange(3); w = 0.18
ax.bar(x-1.5*w, national, w, label='国家级项目', color='#4472C4')
ax.bar(x-0.5*w, provincial, w, label='省部级项目', color='#ED7D31')
ax.bar(x+0.5*w, enterprise, w, label='横向课题', color='#A5A5A5')
ax.bar(x+1.5*w, univ, w, label='校级项目', color='#FFC000')
ax.set_xlabel('年度'); ax.set_ylabel('经费（亿元）'); ax.set_title('2023-2025年科研经费投入趋势')
ax.set_xticks(x); ax.set_xticklabels(years); ax.legend(fontsize=8); ax.set_ylim(0, 1.5)
for i in range(3):
    total = national[i]+provincial[i]+enterprise[i]+univ[i]
    ax.text(x[i], total+0.03, f'{total:.2f}亿', ha='center', fontsize=9)
plt.tight_layout(); fig.savefig(os.path.join(d, '三年科研经费投入趋势.png'), dpi=150); plt.close()

# 1.2 项目经费预算执行PNG
fig, ax = plt.subplots(figsize=(12, 6))
pnames = ['国家重点研发', '国自然基金', '省重点研发', '横向课题A', '横向课题B', '校自主科研', '人才启动', '平台建设']
budgets = [5000, 800, 2000, 300, 150, 500, 200, 1000]
actuals = [5800, 760, 2400, 350, 120, 550, 250, 900]
xp = range(8)
ax.bar([i-0.2 for i in xp], budgets, 0.4, label='预算', color='#4472C4', alpha=0.7)
ax.bar([i+0.2 for i in xp], actuals, 0.4, label='实际', color='#ED7D31', alpha=0.7)
ax.set_ylabel('金额（万元）'); ax.set_title('2025年科研项目经费预算执行情况')
ax.set_xticks(xp); ax.set_xticklabels(pnames, rotation=20, ha='right'); ax.legend()
plt.tight_layout(); fig.savefig(os.path.join(d, '经费预算执行情况.png'), dpi=150); plt.close()

# 1.3 科研项目台账Excel（40个项目）
wb = openpyxl.Workbook(); ws = wb.active; ws.title = '科研项目台账'
ws.append(['序号', '项目编号', '项目名称', '项目负责人', '经费来源', '立项金额(万元)', '到位金额(万元)', '已支出(万元)', '结题状态', '项目状态标记'])
# 40个: 8超预算 + 6经费挪用 + 5结题未验收 + 4预算结余大 + 3超支但成果突出 + 14正常
proj_names = ['纳米材料制备工艺', '高性能合金研究', '智能传感技术', '新能源电池', '量子计算基础',
              '人工智能算法', '生物制药工艺', '环境治理技术', '海洋工程材料', '航空航天结构',
              '芯片设计EDA', '基因编辑技术', '碳捕集利用', '柔性电子器件', '机器人控制',
              '光通信技术', '储能材料开发', '智能交通系统', '土壤修复技术', '纳米药物递送',
              '增材制造工艺', '新型催化剂', '功能涂料研发', '声学超材料', '光电显示技术',
              '生物传感器', '燃料电池堆', '摩擦学基础', '复合材料力学', '数据安全加密',
              '脑机接口', '机器视觉', '自然语言处理', '区块链溯源', '智慧农业物联网',
              '核废料处理', '深海装备', '卫星通信', '高超音速', '量子保密通信']
sources = ['国家重点研发', '国自然基金', '省重点研发', '横向课题', '校级自主', '人才启动', '平台建设']
for i in range(1, 41):
    pname = proj_names[i-1]
    src = sources[i % 7]
    budget = random.randint(100, 5000)
    arrived = int(budget * random.uniform(0.85, 1.0))
    if i <= 8:
        spent = int(budget * random.uniform(1.25, 1.6)); vstat = '在研'; pstate = '超预算执行'
    elif i <= 14:
        spent = int(budget * random.uniform(0.8, 0.95)); vstat = '在研'; pstate = '经费挪用嫌疑'
    elif i <= 19:
        spent = int(budget * random.uniform(0.9, 1.0)); vstat = '结题未验收'; pstate = '结题未验收'
    elif i <= 23:
        spent = int(budget * random.uniform(0.3, 0.45)); vstat = '在研'; pstate = '预算结余大'
    elif i <= 26:
        spent = int(budget * random.uniform(1.15, 1.35)); vstat = '已结题'; pstate = '超支但成果突出'
    else:
        spent = int(budget * random.uniform(0.85, 1.05)); vstat = random.choice(['在研', '已结题']); pstate = '正常'
    ws.append([i, f'KY{i:03d}', pname, f'课题组{i:02d}号', src, budget, arrived, spent, vstat, pstate])
style(ws, 10); wb.save(os.path.join(d, '科研项目台账.xlsx'))

# 1.4 项目成果数据Excel
wb = openpyxl.Workbook(); ws = wb.active; ws.title = '项目成果'
ws.append(['项目编号', '项目名称', '论文(篇)', '专利(项)', '获奖(项)', '成果转化(万元)', '综合评级', '备注'])
for i in range(1, 41):
    if i <= 19:  # 结题未验收前的
        papers = random.randint(2, 8); patents = random.randint(1, 5); awards = random.randint(0, 2)
        transfer = 0; grade = 'B'; note = ''
    elif i <= 26:  # 超支但成果突出
        papers = random.randint(5, 12); patents = random.randint(3, 8); awards = random.randint(1, 3)
        transfer = round(random.uniform(50, 300), 1); grade = 'A'; note = '成果突出，超支合理'
    else:
        papers = random.randint(1, 5); patents = random.randint(0, 3); awards = random.randint(0, 1)
        transfer = round(random.uniform(0, 50), 1); grade = random.choice(['A', 'B', 'B', 'C']); note = ''
    if i <= 14 and i > 8:
        note = '经费使用异常，成果待重新评估'
    ws.append([f'KY{i:03d}', proj_names[i-1], papers, patents, awards, transfer, grade, note])
style(ws, 8); wb.save(os.path.join(d, '项目成果数据.xlsx'))

mkzip(d, '01科研项目与经费预算数据.rar')

# ========== 压缩包2: 经费支出明细 ==========
print('=== 02 经费支出明细数据 ===')
d = os.path.join(INPUT_DIR, '_t2'); os.makedirs(d)

# 2.1 支出结构PNG
fig, ax = plt.subplots(figsize=(8, 6))
labels = ['设备费', '材料费', '测试化验加工费', '燃料动力费', '差旅会议费', '出版文献费', '劳务费', '专家咨询费', '其他']
sizes = [25, 20, 15, 8, 10, 5, 10, 4, 3]
colors = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47', '#9B59B6', '#E74C3C', '#F0E68C']
ax.pie(sizes, labels=labels, autopct='%1.0f%%', colors=colors, startangle=90)
ax.set_title('2025年科研经费支出结构')
plt.tight_layout(); fig.savefig(os.path.join(d, '支出结构分布.png'), dpi=150); plt.close()

# 2.2 各类经费支出对比PNG
fig, ax = plt.subplots(figsize=(10, 6))
cats = ['设备费', '材料费', '测试费', '差旅费', '劳务费', '其他']
planned = [2500, 2000, 1500, 1000, 1000, 800]
actual = [2800, 2400, 1200, 1600, 1300, 700]
x = range(6)
ax.bar([i-0.2 for i in x], planned, 0.4, label='预算', color='#4472C4', alpha=0.7)
ax.bar([i+0.2 for i in x], actual, 0.4, label='实际', color='#ED7D31', alpha=0.7)
ax.set_ylabel('金额（万元）'); ax.set_title('各类经费预算与实际支出对比')
ax.set_xticks(x); ax.set_xticklabels(cats, rotation=20); ax.legend()
plt.tight_layout(); fig.savefig(os.path.join(d, '各类经费支出对比.png'), dpi=150); plt.close()

# 2.3 经费支出明细Excel（60条）
wb = openpyxl.Workbook(); ws = wb.active; ws.title = '经费支出明细'
ws.append(['序号', '凭证编号', '项目编号', '支出类别', '支出金额(万元)', '支出日期', '收款单位', '发票编号', '审批人', '备注'])
# 60条: 10超范围 + 8虚列 + 6拆分采购 + 4跨项目调剂 + 3正常但金额大 + 29正常
for i in range(1, 61):
    pcode = f'KY{random.randint(1,40):03d}'
    cat = random.choice(['设备费', '材料费', '测试化验加工费', '燃料动力费', '差旅会议费', '出版文献费', '劳务费', '专家咨询费'])
    amt = round(random.uniform(0.5, 50), 2)
    if i <= 10:
        cat = '劳务费'; note = '支出内容与项目任务书无关，属超范围支出'
    elif i <= 18:
        cat = '材料费'; note = '发票金额与实际采购金额不符，存在虚列嫌疑'
    elif i <= 24:
        cat = '设备费'; note = '将超过招标限额的采购拆分为多笔小额采购'
    elif i <= 28:
        note = '从A项目调剂至B项目，未经审批'
    elif i <= 31:
        note = '金额较大但属大型设备采购，已按规定招标'  # 反直觉正常项
    else:
        note = ''
    ws.append([i, f'PZ{i:04d}', pcode, cat, amt,
               f'2025-{random.randint(1,9):02d}-{random.randint(1,28):02d}',
               f'供应商{random.randint(1,20)}', f'FP{random.randint(10000,99999)}',
               f'审批人{random.randint(1,8)}', note])
style(ws, 10); wb.save(os.path.join(d, '经费支出明细.xlsx'))

mkzip(d, '02经费支出明细数据.rar')

# ========== 压缩包3: 采购与资产 ==========
print('=== 03 采购与资产数据 ===')
d = os.path.join(INPUT_DIR, '_t3'); os.makedirs(d)

# 3.1 采购方式分布PNG
fig, ax = plt.subplots(figsize=(8, 6))
labels = ['公开招标', '竞争性谈判', '询价采购', '单一来源', '紧急采购']
sizes = [15, 5, 8, 4, 3]
colors = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#FF6B6B']
ax.pie(sizes, labels=labels, autopct='%1.0f%%', colors=colors, startangle=90)
ax.set_title('2025年科研设备采购方式分布（35笔）')
plt.tight_layout(); fig.savefig(os.path.join(d, '采购方式分布.png'), dpi=150); plt.close()

# 3.2 采购记录Excel（35笔）
wb = openpyxl.Workbook(); ws = wb.active; ws.title = '采购记录'
ws.append(['序号', '采购编号', '项目编号', '采购品名', '采购金额(万元)', '采购方式', '供应商', '是否审批', '招标限额', '备注'])
# 35笔: 6拆分规避 + 4超限额未招标 + 3单一来源合理 + 2紧急采购合理 + 17正常
equip_names = ['光谱仪', '显微镜', '离心机', '培养箱', '干燥箱', '光谱分析仪', '电子天平', '反应釜',
               '测试机', '分析仪', '传感器', '控制器', '服务器', '工作站', '软件License']
for i in range(1, 36):
    ename = equip_names[i % len(equip_names)]
    amt = round(random.uniform(5, 200), 2)
    limit = 50  # 招标限额50万
    if i <= 6:
        method = '询价采购'; appr = '是'; note = '单笔低于限额但同类累计超限额，拆分规避招标'
    elif i <= 10:
        method = '询价采购'; appr = '否'; note = f'单笔{amt}万超招标限额{limit}万，未走招标程序'
    elif i <= 13:
        method = '单一来源'; appr = '是'; note = '专有技术设备，单一来源采购经审批合理'  # 反直觉正常项
    elif i <= 15:
        method = '紧急采购'; appr = '是'; note = '实验急需设备，紧急采购经审批合理'  # 反直觉正常项
    else:
        method = random.choice(['公开招标', '询价采购']); appr = '是'; note = ''
    ws.append([i, f'CG{i:03d}', f'KY{random.randint(1,40):03d}', ename, amt, method,
               f'供应商{random.randint(1,20)}', appr, f'{limit}万', note])
style(ws, 10); wb.save(os.path.join(d, '采购记录.xlsx'))

# 3.3 设备资产台账Excel（30台）
wb = openpyxl.Workbook(); ws = wb.active; ws.title = '设备资产台账'
ws.append(['序号', '资产编号', '设备名称', '所属项目', '购置金额(万元)', '购入日期', '使用状态', '共享状态', '利用率(%)', '备注'])
# 30台: 8闲置 + 5共享 + 4利用率低但合理 + 13正常
for i in range(1, 31):
    dname = equip_names[i % len(equip_names)]
    amt = round(random.uniform(10, 300), 2)
    if i <= 8:
        ustate = '闲置'; share = '否'; util = random.randint(0, 10); note = '设备闲置超6个月，需调配'
    elif i <= 13:
        ustate = '在用'; share = '是'; util = random.randint(60, 90); note = '已纳入大型仪器共享平台'
    elif i <= 17:
        ustate = '在用'; share = '否'; util = random.randint(15, 30); note = '利用率低但属专属实验定制设备，合理低利用'  # 反直觉正常项
    else:
        ustate = '在用'; share = random.choice(['是', '否']); util = random.randint(40, 85); note = ''
    ws.append([i, f'ZC{i:03d}', dname, f'KY{random.randint(1,40):03d}', amt,
               f'2025-{random.randint(1,9):02d}-{random.randint(1,28):02d}', ustate, share, util, note])
style(ws, 10); wb.save(os.path.join(d, '设备资产台账.xlsx'))

mkzip(d, '03采购与资产数据.rar')

# ========== 压缩包4: 成果与绩效 ==========
print('=== 04 成果与绩效数据 ===')
d = os.path.join(INPUT_DIR, '_t4'); os.makedirs(d)

# 4.1 成果产出趋势PNG
fig, ax = plt.subplots(figsize=(10, 6))
years = ['2023', '2024', '2025']
papers = [120, 145, 168]; patents = [35, 42, 55]; awards = [8, 10, 12]
x = range(3)
ax.plot(x, papers, 'o-', label='论文(篇)', color='#4472C4', linewidth=2)
ax.plot(x, patents, 's-', label='专利(项)', color='#ED7D31', linewidth=2)
ax.plot(x, awards, '^-', label='获奖(项)', color='#70AD47', linewidth=2)
ax.set_xlabel('年度'); ax.set_ylabel('数量'); ax.set_title('2023-2025年科研成果产出趋势')
ax.set_xticks(x); ax.set_xticklabels(years); ax.legend()
for j, vals in enumerate([papers, patents, awards]):
    for i, v in enumerate(vals):
        ax.text(i, v+3, str(v), ha='center', fontsize=8)
plt.tight_layout(); fig.savefig(os.path.join(d, '成果产出趋势.png'), dpi=150); plt.close()

# 4.2 成果转化金额PNG
fig, ax = plt.subplots(figsize=(10, 6))
depts = ['材料学院', '化工学院', '机械学院', '信息学院', '生医学院', '环建学院', '理学院', '其他']
transfer = [350, 280, 220, 180, 150, 100, 80, 50]
colors_t = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47', '#9B59B6', '#E74C3C']
ax.bar(depts, transfer, color=colors_t, width=0.6)
ax.set_ylabel('转化金额（万元）'); ax.set_title('2025年各学院成果转化金额')
ax.set_ylim(0, 400)
for i, v in enumerate(transfer):
    ax.text(i, v+5, f'{v}万', ha='center', fontsize=9)
plt.xticks(rotation=20, ha='right'); plt.tight_layout()
fig.savefig(os.path.join(d, '成果转化金额.png'), dpi=150); plt.close()

# 4.3 项目绩效评分Excel（40个项目）
wb = openpyxl.Workbook(); ws = wb.active; ws.title = '项目绩效'
ws.append(['项目编号', '项目名称', '学术成果(30分)', '技术指标(25分)', '经济效益(20分)', '人才培养(15分)', '经费执行(10分)', '总分(100分)', '评级', '备注'])
for i in range(1, 41):
    if i <= 8:
        ac = random.randint(20, 28); ti = random.randint(18, 23); ec = random.randint(12, 16)
        tr = random.randint(10, 13); fe = random.randint(3, 5)
    elif i <= 14:
        ac = random.randint(18, 24); ti = random.randint(16, 20); ec = random.randint(10, 14)
        tr = random.randint(8, 12); fe = random.randint(4, 6)
    elif i <= 19:
        ac = random.randint(15, 22); ti = random.randint(14, 20); ec = random.randint(10, 15)
        tr = random.randint(9, 12); fe = random.randint(6, 8)
    elif i <= 23:
        ac = random.randint(20, 26); ti = random.randint(18, 23); ec = random.randint(14, 18)
        tr = random.randint(11, 14); fe = random.randint(2, 4)
    elif i <= 26:
        ac = random.randint(26, 30); ti = random.randint(22, 25); ec = random.randint(16, 20)
        tr = random.randint(13, 15); fe = random.randint(3, 5)
    else:
        ac = random.randint(20, 28); ti = random.randint(18, 24); ec = random.randint(13, 18)
        tr = random.randint(10, 14); fe = random.randint(7, 9)
    total = ac+ti+ec+tr+fe
    grade = 'A' if total >= 85 else 'B' if total >= 70 else 'C'
    note = ''
    if i <= 14 and i > 8: note = '经费使用异常，绩效待重新评估'
    if i <= 26 and i > 23: note = '超支但成果突出，绩效优秀'
    ws.append([f'KY{i:03d}', proj_names[i-1], ac, ti, ec, tr, fe, total, grade, note])
style(ws, 10); wb.save(os.path.join(d, '项目绩效评分表.xlsx'))

mkzip(d, '04成果与绩效数据.rar')

# ========== 压缩包5: 管理制度 ==========
print('=== 05 科研管理制度 ===')
d = os.path.join(INPUT_DIR, '_t5'); os.makedirs(d)
policies = [
    {'name': '01科研经费管理办法.docx', 'title': '科研经费管理办法',
     'secs': ['总则', '预算管理', '支出管理', '决算管理', '监督检查', '附则'],
     'issues': '问题：1.预算调整审批权限未按金额分级；2.跨项目经费调剂无审批流程；3.超范围支出无预警机制'},
    {'name': '02科研采购管理办法.docx', 'title': '科研采购管理办法',
     'secs': ['总则', '采购计划', '采购方式', '招标管理', '合同管理', '验收付款', '附则'],
     'issues': '问题：1.招标限额未区分设备类型；2.拆分采购无识别机制；3.单一来源采购审批不严'},
    {'name': '03科研资产管理办法.docx', 'title': '科研资产管理办法',
     'secs': ['总则', '资产购置', '资产使用', '共享管理', '处置管理', '附则'],
     'issues': '问题：1.闲置设备无调配机制；2.共享平台覆盖率不足；3.资产处置审批不严'},
    {'name': '04科研成果管理办法.docx', 'title': '科研成果管理办法',
     'secs': ['总则', '成果认定', '转化管理', '知识产权', '奖励机制', '附则'],
     'issues': '问题：1.成果认定标准不量化；2.转化收益分配不合理；3.知识产权保护不足'},
    {'name': '05科研绩效评价办法.docx', 'title': '科研绩效评价办法',
     'secs': ['总则', '评价体系', '评价指标', '评价流程', '结果应用', '附则'],
     'issues': '问题：1.指标权重不合理；2.经费执行率权重过低；3.评价结果未与后续立项挂钩'},
]
for p in policies:
    doc = Document(); doc.add_heading(p['title'], level=0)
    for sec in p['secs']:
        doc.add_heading(sec, level=1)
        doc.add_paragraph(f'{sec}具体条款内容详见正式制度文件。本制度适用于学校科研管理各环节。')
    doc.add_heading('现有制度问题诊断', level=1)
    doc.add_paragraph(p['issues'])
    doc.save(os.path.join(d, p['name']))
mkzip(d, '05科研管理制度.rar')

# ========== 压缩包6: 管理层要求 ==========
print('=== 06 管理层要求 ===')
d = os.path.join(INPUT_DIR, '_t6'); os.makedirs(d)
doc = Document(); doc.add_heading('管理层要求汇总', level=0)
reqs = [
    (1, '校长', '超预算项目不追责继续执行', '追责影响科研积极性'),
    (2, '财务处长', '经费挪用项目不处理', '处理可能引发课题组不满'),
    (3, '科研院长', '结题未验收项目直接通过', '验收流程太长影响新项目立项'),
    (4, '设备处长', '拆分采购不追究', '拆分采购是课题组常见做法'),
    (5, '校长', '闲置设备不调配', '调配可能引发课题组矛盾'),
    (6, '财务处长', '超范围支出不追回', '追回影响课题组运转'),
    (7, '科研院长', '虚列支出不审计', '审计需要大量人力'),
    (8, '设备处长', '单一来源采购不审批', '专有设备需要快速采购'),
    (9, '校长', '预算结余大项目不调整', '调整可能影响科研进度'),
    (10, '科研院长', '绩效评价只看论文不看经费执行', '论文是核心指标'),
]
for num, src, content, reason in reqs:
    doc.add_heading(f'要求{num}', level=1)
    p = doc.add_paragraph(); p.add_run('来源：').bold = True; p.add_run(src)
    p = doc.add_paragraph(); p.add_run('要求：').bold = True; p.add_run(content)
    p = doc.add_paragraph(); p.add_run('理由：').bold = True; p.add_run(reason)
    doc.add_paragraph()
doc.save(os.path.join(d, '管理层要求.docx'))

# 校长邮件PNG
fig, ax = plt.subplots(figsize=(10, 4)); ax.axis('off')
ax.text(0.05, 0.9, '发件人：校长办公室', fontsize=11, transform=ax.transAxes)
ax.text(0.05, 0.75, '收件人：科研处、财务处、设备处', fontsize=11, transform=ax.transAxes)
ax.text(0.05, 0.6, '主题：关于科研经费审计工作的指导意见', fontsize=11, transform=ax.transAxes)
ax.text(0.05, 0.3, '各位：\n经校务会讨论，科研经费审计工作涉及面广，\n为保护科研人员积极性，请暂缓执行超预算追责、\n经费挪用处理、拆分采购追究等措施，先维持现状。\n\n校长\n2025年9月25日', fontsize=10, transform=ax.transAxes)
rect = plt.Rectangle((0,0),1,1, transform=ax.transAxes, fill=False, edgecolor='gray')
ax.add_patch(rect); plt.tight_layout()
fig.savefig(os.path.join(d, '校长邮件.png'), dpi=150); plt.close()
mkzip(d, '06管理层要求.rar')

print('\n=== 输入附件生成完成 ===')
