#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""义务教育经费管理使用审计 - 输入附件生成"""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from docx import Document
import zipfile, os, random, shutil
import numpy as np

random.seed(55)
plt.rcParams['font.sans-serif'] = ['WenQuanYi Micro Hei', 'SimHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

BASE = '/workspace/义务教育经费审计_附件'
INPUT_DIR = os.path.join(BASE, '输入附件')
os.makedirs(INPUT_DIR, exist_ok=True)

hf = Font(name='微软雅黑', size=11, bold=True, color='FFFFFF')
hfill = PatternFill(start_color='2F5496', end_color='2F5496', fill_type='solid')
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

# ========== 压缩包1: 学校基本信息与经费预算 ==========
print('=== 01 学校基本信息与经费预算 ===')
d = os.path.join(INPUT_DIR, '_t1'); os.makedirs(d)

# 1.1 三年义务教育经费投入趋势PNG
fig, ax = plt.subplots(figsize=(10, 6))
years = ['2023年', '2024年', '2025年']
public_fund = [0.35, 0.38, 0.42]
nutrition = [0.12, 0.14, 0.16]
maintain = [0.08, 0.09, 0.10]
training = [0.05, 0.06, 0.07]
x = np.arange(3); w = 0.18
ax.bar(x-1.5*w, public_fund, w, label='生均公用经费', color='#4472C4')
ax.bar(x-0.5*w, nutrition, w, label='营养餐补助', color='#ED7D31')
ax.bar(x+0.5*w, maintain, w, label='校舍维修', color='#A5A5A5')
ax.bar(x+1.5*w, training, w, label='教师培训', color='#FFC000')
ax.set_xlabel('年度'); ax.set_ylabel('经费（亿元）'); ax.set_title('2023-2025年义务教育经费投入趋势')
ax.set_xticks(x); ax.set_xticklabels(years); ax.legend(fontsize=8); ax.set_ylim(0, 0.5)
for i in range(3):
    total = public_fund[i]+nutrition[i]+maintain[i]+training[i]
    ax.text(x[i], total+0.01, f'{total:.2f}亿', ha='center', fontsize=9)
plt.tight_layout(); fig.savefig(os.path.join(d, '三年义务教育经费投入趋势.png'), dpi=150); plt.close()

# 1.2 学校经费预算执行PNG
fig, ax = plt.subplots(figsize=(12, 6))
snames = ['实验小学', '中心小学', '第一中学', '第二中学', '育才小学', '阳光中学', '新华小学', '和平中学']
budgets = [800, 650, 1200, 1000, 450, 900, 380, 850]
actuals = [960, 620, 1450, 1050, 400, 1100, 350, 800]
xp = range(8)
ax.bar([i-0.2 for i in xp], budgets, 0.4, label='预算', color='#4472C4', alpha=0.7)
ax.bar([i+0.2 for i in xp], actuals, 0.4, label='实际', color='#ED7D31', alpha=0.7)
ax.set_ylabel('金额（万元）'); ax.set_title('2025年部分学校经费预算执行情况')
ax.set_xticks(xp); ax.set_xticklabels(snames, rotation=20, ha='right'); ax.legend()
plt.tight_layout(); fig.savefig(os.path.join(d, '学校经费预算执行.png'), dpi=150); plt.close()

# 1.3 学校台账Excel（30所学校）
wb = openpyxl.Workbook(); ws = wb.active; ws.title = '学校台账'
ws.append(['序号', '学校编号', '学校名称', '学校类型', '学生人数', '在编教师', '临聘教师', '生均公用经费(元/人)', '到位金额(万元)', '已支出(万元)', '结余金额(万元)', '学校状态标记'])
# 30所: 8超预算 + 6经费挪用 + 5结余过大 + 4生均经费不足 + 3超支但绩效好 + 4正常
school_names = ['实验小学', '中心小学', '第一中学', '第二中学', '育才小学', '阳光中学', '新华小学', '和平中学',
                '建国小学', '新民中学', '工农小学', '红旗中学', '团结小学', '胜利中学', '光明小学',
                '朝阳中学', '迎春小学', '解放中学', '幸福小学', '民主中学', '文明小学', '和谐中学',
                '进步小学', '复兴中学', '繁荣小学', '昌盛中学', '兴旺小学', '发达中学', '如意小学', '吉祥中学']
for i in range(1, 31):
    sname = school_names[i-1]
    stype = '小学' if i % 2 == 1 else '中学'
    students = random.randint(300, 2500)
    teachers = int(students * random.uniform(0.05, 0.08))
    temp_teachers = random.randint(0, 8)
    per_student = random.randint(600, 900)
    arrived = int(students * per_student / 10000)
    if i <= 8:
        spent = int(arrived * random.uniform(1.25, 1.6)); state = '超预算执行'
    elif i <= 14:
        spent = int(arrived * random.uniform(0.85, 0.95)); state = '经费挪用嫌疑'
    elif i <= 19:
        spent = int(arrived * random.uniform(0.35, 0.45)); state = '结余过大'
    elif i <= 23:
        per_student = random.randint(400, 550); arrived = int(students * per_student / 10000)
        spent = int(arrived * random.uniform(0.85, 0.95)); state = '生均经费不足'
    elif i <= 26:
        spent = int(arrived * random.uniform(1.15, 1.3)); state = '超支但绩效突出'
    else:
        spent = int(arrived * random.uniform(0.85, 1.0)); state = '正常'
    surplus = arrived - spent
    ws.append([i, f'XX{i:03d}', sname, stype, students, teachers, temp_teachers, per_student, arrived, spent, surplus, state])
style(ws, 12); wb.save(os.path.join(d, '学校台账.xlsx'))

# 1.4 学校绩效数据Excel
wb = openpyxl.Workbook(); ws = wb.active; ws.title = '学校绩效'
ws.append(['学校编号', '学校名称', '教学质量(30分)', '安全管理(25分)', '经费执行(20分)', '师资建设(15分)', '社会满意度(10分)', '总分(100分)', '评级', '备注'])
for i in range(1, 31):
    if i <= 8:
        tq = random.randint(22, 28); sm = random.randint(20, 24); fe = random.randint(3, 6)
        tb = random.randint(10, 13); ss = random.randint(7, 9); note = '超预算但绩效正常'
    elif i <= 14:
        tq = random.randint(18, 24); sm = random.randint(16, 20); fe = random.randint(4, 7)
        tb = random.randint(8, 12); ss = random.randint(6, 8); note = '经费使用异常'
    elif i <= 19:
        tq = random.randint(20, 26); sm = random.randint(18, 22); fe = random.randint(8, 10)
        tb = random.randint(9, 12); ss = random.randint(6, 9); note = '结余过大执行率低'
    elif i <= 23:
        tq = random.randint(15, 20); sm = random.randint(16, 20); fe = random.randint(15, 18)
        tb = random.randint(8, 11); ss = random.randint(5, 7); note = '经费不足影响教学'
    elif i <= 26:
        tq = random.randint(27, 30); sm = random.randint(23, 25); fe = random.randint(4, 6)
        tb = random.randint(13, 15); ss = random.randint(9, 10); note = '超支但绩效突出（合理）'
    else:
        tq = random.randint(22, 28); sm = random.randint(20, 24); fe = random.randint(15, 19)
        tb = random.randint(10, 14); ss = random.randint(7, 9); note = ''
    total = tq+sm+fe+tb+ss
    grade = 'A' if total >= 85 else 'B' if total >= 70 else 'C'
    ws.append([f'XX{i:03d}', school_names[i-1], tq, sm, fe, tb, ss, total, grade, note])
style(ws, 10); wb.save(os.path.join(d, '学校绩效数据.xlsx'))

mkzip(d, '01学校基本信息与经费预算数据.rar')

# ========== 压缩包2: 经费支出明细 ==========
print('=== 02 经费支出明细 ===')
d = os.path.join(INPUT_DIR, '_t2'); os.makedirs(d)

fig, ax = plt.subplots(figsize=(8, 6))
labels = ['办公费', '印刷费', '水电费', '差旅费', '维修维护费', '会议费', '培训费', '专用材料费', '劳务费', '其他']
sizes = [15, 8, 10, 5, 20, 5, 8, 15, 10, 4]
colors = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47', '#9B59B6', '#E74C3C', '#F0E68C', '#BDC3C7']
ax.pie(sizes, labels=labels, autopct='%1.0f%%', colors=colors, startangle=90)
ax.set_title('2025年义务教育经费支出结构')
plt.tight_layout(); fig.savefig(os.path.join(d, '支出结构分布.png'), dpi=150); plt.close()

fig, ax = plt.subplots(figsize=(10, 6))
cats = ['办公费', '维修费', '培训费', '差旅费', '会议费']
planned = [150, 200, 80, 50, 50]
actual = [180, 280, 60, 120, 70]
x = range(5)
ax.bar([i-0.2 for i in x], planned, 0.4, label='预算', color='#4472C4', alpha=0.7)
ax.bar([i+0.2 for i in x], actual, 0.4, label='实际', color='#ED7D31', alpha=0.7)
ax.set_ylabel('金额（万元）'); ax.set_title('各类经费预算与实际支出对比')
ax.set_xticks(x); ax.set_xticklabels(cats); ax.legend()
plt.tight_layout(); fig.savefig(os.path.join(d, '各类经费支出对比.png'), dpi=150); plt.close()

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '经费支出明细'
ws.append(['序号', '凭证编号', '学校编号', '支出类别', '支出金额(万元)', '支出日期', '收款单位', '发票编号', '审批人', '备注'])
# 60条: 10超范围 + 8虚列 + 6拆分 + 4挪用营养餐 + 3正常大额 + 29正常
for i in range(1, 61):
    scode = f'XX{random.randint(1,30):03d}'
    cat = random.choice(['办公费', '印刷费', '水电费', '差旅费', '维修维护费', '会议费', '培训费', '专用材料费', '劳务费'])
    amt = round(random.uniform(0.1, 30), 2)
    if i <= 10:
        cat = '劳务费'; note = '发放人员补助与公用经费用途不符，超范围支出'
    elif i <= 18:
        cat = '办公费'; note = '发票金额与实际采购金额不符，存在虚列'
    elif i <= 24:
        cat = '维修维护费'; note = '将超招标限额的维修拆分为多笔小额'
    elif i <= 28:
        cat = '营养餐补助'; note = '营养餐补助经费挪用于办公支出'
    elif i <= 31:
        note = '大型维修项目，已按规定招标'  # 反直觉正常项
    else:
        note = ''
    ws.append([i, f'PZ{i:04d}', scode, cat, amt,
               f'2025-{random.randint(1,9):02d}-{random.randint(1,28):02d}',
               f'供应商{random.randint(1,20)}', f'FP{random.randint(10000,99999)}',
               f'审批人{random.randint(1,8)}', note])
style(ws, 10); wb.save(os.path.join(d, '经费支出明细.xlsx'))

mkzip(d, '02经费支出明细数据.rar')

# ========== 压缩包3: 采购与资产 ==========
print('=== 03 采购与资产 ===')
d = os.path.join(INPUT_DIR, '_t3'); os.makedirs(d)

fig, ax = plt.subplots(figsize=(8, 6))
labels = ['公开招标', '竞争性谈判', '询价采购', '单一来源', '紧急采购']
sizes = [12, 5, 8, 4, 3]
colors = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#FF6B6B']
ax.pie(sizes, labels=labels, autopct='%1.0f%%', colors=colors, startangle=90)
ax.set_title('2025年义务教育学校采购方式分布（32笔）')
plt.tight_layout(); fig.savefig(os.path.join(d, '采购方式分布.png'), dpi=150); plt.close()

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '采购记录'
ws.append(['序号', '采购编号', '学校编号', '采购品名', '采购金额(万元)', '采购方式', '供应商', '是否审批', '招标限额', '备注'])
items = ['教学多媒体', '课桌椅', '实验器材', '图书', '体育器材', '办公电脑', '打印机', '空调', '安保设备', '食堂设备']
for i in range(1, 33):
    iname = items[i % len(items)]
    amt = round(random.uniform(1, 80), 2)
    if i <= 6:
        method = '询价采购'; note = '同类采购累计超限额但单笔低于限额，拆分规避'
    elif i <= 10:
        method = '询价采购'; note = f'单笔{amt}万超招标限额5万，未走招标'
    elif i <= 13:
        method = '单一来源'; note = '专有教学设备，单一来源经审批合理'  # 反直觉正常项
    elif i <= 15:
        method = '紧急采购'; note = '教学急需设备，紧急采购经审批合理'  # 反直觉正常项
    else:
        method = random.choice(['公开招标', '询价采购']); note = ''
    ws.append([i, f'CG{i:03d}', f'XX{random.randint(1,30):03d}', iname, amt, method,
               f'供应商{random.randint(1,20)}', '是', '5万', note])
style(ws, 10); wb.save(os.path.join(d, '采购记录.xlsx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '设备资产台账'
ws.append(['序号', '资产编号', '设备名称', '所属学校', '购置金额(万元)', '购入日期', '使用状态', '共享状态', '利用率(%)', '备注'])
for i in range(1, 31):
    dname = items[i % len(items)]
    amt = round(random.uniform(0.5, 50), 2)
    if i <= 8:
        ustate = '闲置'; share = '否'; util = random.randint(0, 10); note = '设备闲置超6个月'
    elif i <= 13:
        ustate = '在用'; share = '是'; util = random.randint(60, 90); note = '已纳入区域共享'
    elif i <= 17:
        ustate = '在用'; share = '否'; util = random.randint(15, 30); note = '专用实验设备利用率低但合理'  # 反直觉正常项
    else:
        ustate = '在用'; share = random.choice(['是', '否']); util = random.randint(40, 85); note = ''
    ws.append([i, f'ZC{i:03d}', dname, f'XX{random.randint(1,30):03d}', amt,
               f'2025-{random.randint(1,9):02d}-{random.randint(1,28):02d}', ustate, share, util, note])
style(ws, 10); wb.save(os.path.join(d, '设备资产台账.xlsx'))

mkzip(d, '03采购与资产数据.rar')

# ========== 压缩包4: 营养餐与资助 ==========
print('=== 04 营养餐与资助 ===')
d = os.path.join(INPUT_DIR, '_t4'); os.makedirs(d)

fig, ax = plt.subplots(figsize=(10, 6))
months = [f'{m}月' for m in range(1, 13)]
planned = [120, 120, 120, 120, 120, 120, 80, 80, 120, 120, 120, 120]
actual = [120, 115, 118, 122, 110, 125, 75, 78, 120, 115, 120, 118]
x = range(12)
ax.plot(months, planned, 's-', label='预算', color='#4472C4', linewidth=2)
ax.plot(months, actual, 'o-', label='实际', color='#ED7D31', linewidth=2)
ax.fill_between(range(12), planned, actual, alpha=0.15, color='red')
ax.set_xlabel('月份'); ax.set_ylabel('支出（万元）'); ax.set_title('2025年营养餐补助月度支出')
ax.set_xticks(x); ax.set_xticklabels(months, rotation=45); ax.legend()
plt.tight_layout(); fig.savefig(os.path.join(d, '营养餐补助支出趋势.png'), dpi=150); plt.close()

fig, ax = plt.subplots(figsize=(8, 6))
labels = ['小学', '初中', '特殊教育']
sizes = [60, 35, 5]
colors = ['#4472C4', '#ED7D31', '#FFC000']
ax.pie(sizes, labels=labels, autopct='%1.0f%%', colors=colors, startangle=90)
ax.set_title('受资助学生学段分布')
plt.tight_layout(); fig.savefig(os.path.join(d, '受资助学生分布.png'), dpi=150); plt.close()

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '营养餐支出明细'
ws.append(['序号', '支出编号', '学校编号', '支出月份', '支出金额(万元)', '采购食材', '供应商', '验收人', '备注'])
# 50条: 8挪用 + 6虚列 + 5标准超标 + 3正常大额 + 28正常
for i in range(1, 51):
    scode = f'XX{random.randint(1,30):03d}'
    month = random.randint(1, 12)
    amt = round(random.uniform(0.5, 15), 2)
    food = random.choice(['大米', '蔬菜', '肉类', '食用油', '调味品', '蛋奶'])
    if i <= 8:
        note = '营养餐补助经费挪用于非食品采购'
    elif i <= 14:
        note = '食材采购量与学生人数不符，存在虚列'
    elif i <= 19:
        note = '单餐标准超标（5元/人，规定4元/人）'
    elif i <= 22:
        note = '大批量食材集中采购，已按规定招标'  # 反直觉正常项
    else:
        note = ''
    ws.append([i, f'YY{i:04d}', scode, f'{month}月', amt, food,
               f'供应商{random.randint(1,15)}', f'验收人{random.randint(1,8)}', note])
style(ws, 9); wb.save(os.path.join(d, '营养餐支出明细.xlsx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '学生资助明细'
ws.append(['序号', '资助编号', '学生编号', '学校编号', '学段', '家庭情况', '资助类型', '资助金额(元/学期)', '发放状态', '备注'])
# 40名: 6不符条件 + 4重复发放 + 3金额超标 + 27正常
for i in range(1, 41):
    scode = f'XX{random.randint(1,30):03d}'
    stage = random.choice(['小学', '初中', '特殊教育'])
    fam = random.choice(['低保家庭', '残疾儿童', '孤儿', '建档立卡', '一般困难'])
    atype = random.choice(['寄宿生补助', '营养改善计划', '免学杂费', '助学金'])
    amt = random.randint(500, 2000)
    if i <= 6:
        note = '家庭情况证明材料不全，不符合资助条件'
    elif i <= 10:
        note = '同一学生重复领取多类资助'
    elif i <= 13:
        note = '资助金额超出标准（2000元，标准1500元）'
    else:
        note = ''
    ws.append([i, f'ZZ{i:04d}', f'S{i:04d}', scode, stage, fam, atype, amt, '已发放', note])
style(ws, 10); wb.save(os.path.join(d, '学生资助明细.xlsx'))

mkzip(d, '04营养餐与资助数据.rar')

# ========== 压缩包5: 教育经费管理制度 ==========
print('=== 05 教育经费管理制度 ===')
d = os.path.join(INPUT_DIR, '_t5'); os.makedirs(d)
policies = [
    {'name': '01义务教育经费管理办法.docx', 'title': '义务教育经费管理办法',
     'secs': ['总则', '预算管理', '支出管理', '决算管理', '监督检查', '附则'],
     'issues': '问题：1.预算调整审批未按金额分级；2.超范围支出无预警；3.结余资金处理不明确'},
    {'name': '02学校采购管理办法.docx', 'title': '学校采购管理办法',
     'secs': ['总则', '采购计划', '采购方式', '招标管理', '合同管理', '验收付款', '附则'],
     'issues': '问题：1.招标限额统一5万未区分；2.拆分采购无识别；3.单一来源审批不严'},
    {'name': '03营养餐管理办法.docx', 'title': '营养餐管理办法',
     'secs': ['总则', '资金管理', '采购验收', '食品质量', '监督公示', '附则'],
     'issues': '问题：1.单餐标准执行不严；2.验收记录不规范；3.公示制度未落实'},
    {'name': '04学生资助管理办法.docx', 'title': '学生资助管理办法',
     'secs': ['总则', '认定标准', '资助类型', '资金发放', '监督管理', '附则'],
     'issues': '问题：1.认定材料审核不严；2.重复发放无核查；3.资助标准执行不严'},
    {'name': '05学校资产管理办法.docx', 'title': '学校资产管理办法',
     'secs': ['总则', '资产购置', '资产使用', '共享管理', '处置管理', '附则'],
     'issues': '问题：1.闲置设备无调配；2.共享机制不健全；3.处置审批不严'},
]
for p in policies:
    doc = Document(); doc.add_heading(p['title'], level=0)
    for sec in p['secs']:
        doc.add_heading(sec, level=1)
        doc.add_paragraph(f'{sec}具体条款内容详见正式制度文件。')
    doc.add_heading('现有制度问题诊断', level=1)
    doc.add_paragraph(p['issues'])
    doc.save(os.path.join(d, p['name']))
mkzip(d, '05教育经费管理制度.rar')

# ========== 压缩包6: 管理层要求 ==========
print('=== 06 管理层要求 ===')
d = os.path.join(INPUT_DIR, '_t6'); os.makedirs(d)
doc = Document(); doc.add_heading('管理层要求汇总', level=0)
reqs = [
    (1, '教育局长', '超预算学校不追责', '追责影响学校积极性'),
    (2, '财务科长', '经费挪用不处理', '处理引发学校不满'),
    (3, '基教科长', '结余过大不调整', '调整影响学校运转'),
    (4, '设备科长', '拆分采购不追究', '拆分采购是学校常见做法'),
    (5, '教育局长', '闲置设备不调配', '调配引发校际矛盾'),
    (6, '财务科长', '营养餐挪用不追回', '追回影响学生供餐'),
    (7, '资助科长', '不符条件资助不追回', '追回影响学生家庭'),
    (8, '设备科长', '单一来源不审批', '教学设备需快速采购'),
    (9, '教育局长', '生均经费不足不追加', '追加需要财政统筹'),
    (10, '基教科长', '绩效只看升学率不看经费执行', '升学率是核心指标'),
]
for num, src, content, reason in reqs:
    doc.add_heading(f'要求{num}', level=1)
    p = doc.add_paragraph(); p.add_run('来源：').bold = True; p.add_run(src)
    p = doc.add_paragraph(); p.add_run('要求：').bold = True; p.add_run(content)
    p = doc.add_paragraph(); p.add_run('理由：').bold = True; p.add_run(reason)
    doc.add_paragraph()
doc.save(os.path.join(d, '管理层要求.docx'))

fig, ax = plt.subplots(figsize=(10, 4)); ax.axis('off')
ax.text(0.05, 0.9, '发件人：教育局办公室', fontsize=11, transform=ax.transAxes)
ax.text(0.05, 0.75, '收件人：财务科、基教科、设备科', fontsize=11, transform=ax.transAxes)
ax.text(0.05, 0.6, '主题：关于义务教育经费审计工作的指导意见', fontsize=11, transform=ax.transAxes)
ax.text(0.05, 0.3, '各位：\n经局务会讨论，义务教育经费审计工作涉及面广，\n为保护学校办学积极性，请暂缓执行超预算追责、\n经费挪用处理、拆分采购追究等措施，先维持现状。\n\n教育局长\n2025年9月28日', fontsize=10, transform=ax.transAxes)
rect = plt.Rectangle((0,0),1,1, transform=ax.transAxes, fill=False, edgecolor='gray')
ax.add_patch(rect); plt.tight_layout()
fig.savefig(os.path.join(d, '教育局长邮件.png'), dpi=150); plt.close()
mkzip(d, '06管理层要求.rar')

print('\n=== 输入附件生成完成 ===')
