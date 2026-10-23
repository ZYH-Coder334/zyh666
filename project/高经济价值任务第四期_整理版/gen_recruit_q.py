#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""中小学招生入学合规审查 - 输入附件生成"""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from docx import Document
import zipfile, os, random, shutil
import numpy as np

random.seed(77)
plt.rcParams['font.sans-serif'] = ['WenQuanYi Micro Hei', 'SimHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

BASE = '/workspace/招生合规审查_附件'
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

# ========== 压缩包1: 招生计划与学校信息 ==========
print('=== 01 招生计划与学校信息 ===')
d = os.path.join(INPUT_DIR, '_t1'); os.makedirs(d)

fig, ax = plt.subplots(figsize=(10, 6))
years = ['2023年', '2024年', '2025年']
plan = [8000, 8500, 9000]
actual = [8200, 8900, 9800]
x = np.arange(3); w = 0.3
ax.bar(x-w/2, plan, w, label='招生计划', color='#4472C4')
ax.bar(x+w/2, actual, w, label='实际招生', color='#ED7D31')
ax.set_ylabel('招生人数'); ax.set_title('2023-2025年招生计划与实际执行')
ax.set_xticks(x); ax.set_xticklabels(years); ax.legend()
for i in range(3):
    ax.text(x[i]-w/2, plan[i]+50, str(plan[i]), ha='center', fontsize=9)
    ax.text(x[i]+w/2, actual[i]+50, str(actual[i]), ha='center', fontsize=9)
plt.tight_layout(); fig.savefig(os.path.join(d, '三年招生计划执行.png'), dpi=150); plt.close()

fig, ax = plt.subplots(figsize=(12, 6))
snames = ['实验小学', '中心小学', '第一中学', '第二中学', '育才小学', '阳光中学', '新华小学', '和平中学']
plans = [400, 350, 600, 500, 300, 450, 250, 400]
actuals = [480, 340, 720, 510, 280, 540, 230, 410]
xp = range(8)
ax.bar([i-0.2 for i in xp], plans, 0.4, label='计划', color='#4472C4', alpha=0.7)
ax.bar([i+0.2 for i in xp], actuals, 0.4, label='实际', color='#ED7D31', alpha=0.7)
ax.set_ylabel('人数'); ax.set_title('2025年部分学校招生计划执行')
ax.set_xticks(xp); ax.set_xticklabels(snames, rotation=20, ha='right'); ax.legend()
plt.tight_layout(); fig.savefig(os.path.join(d, '学校招生计划执行.png'), dpi=150); plt.close()

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '学校招生台账'
ws.append(['序号', '学校编号', '学校名称', '学校类型', '招生计划数', '实际招生数', '特长生计划', '特长生实际', '划片范围', '学校状态标记'])
# 25所: 8超额 + 6跨片 + 5特长生违规 + 3反直觉正常(政策允许扩招) + 3正常
snames = ['实验小学', '中心小学', '第一中学', '第二中学', '育才小学', '阳光中学', '新华小学', '和平中学',
          '建国小学', '新民中学', '工农小学', '红旗中学', '团结小学', '胜利中学', '光明小学',
          '朝阳中学', '迎春小学', '解放中学', '幸福小学', '民主中学', '文明小学', '和谐中学',
          '进步小学', '复兴中学', '繁荣小学']
for i in range(1, 26):
    sname = snames[i-1]
    stype = '小学' if i % 2 == 1 else '中学'
    plan = random.randint(200, 600)
    tal_plan = random.randint(0, 30)
    if i <= 8:
        actual = int(plan * random.uniform(1.15, 1.4)); state = '超额招生'
    elif i <= 14:
        actual = plan; state = '跨片区招生'
    elif i <= 19:
        actual = plan; tal_actual = tal_plan + random.randint(5, 15); state = '特长生违规'
    elif i <= 22:
        actual = int(plan * random.uniform(1.08, 1.12)); state = '政策允许扩招（合理）'
    else:
        actual = plan; state = '正常'
    ws.append([i, f'XX{i:03d}', sname, stype, plan, actual, tal_plan, tal_plan if i>19 else tal_plan,
               f'片区{i//2+1}', state])
style(ws, 10); wb.save(os.path.join(d, '学校招生台账.xlsx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '招生政策'
ws.append(['政策编号', '政策名称', '核心条款', '适用范围', '备注'])
policies = [
    ('ZC001', '义务教育招生工作意见', '免试就近入学，严禁考试选拔', '全部学校', '核心政策'),
    ('ZC002', '招生计划管理办法', '严格按计划招生，超额需审批', '全部学校', '超10%需报教育局'),
    ('ZC003', '特长生招生管理办法', '特长生比例不超5%，需公示', '有特长生资格学校', '含认定标准'),
    ('ZC004', '均衡编班指导意见', '严禁分重点班快慢班，随机编班', '全部学校', '需公示编班方案'),
    ('ZC005', '民办学校招生管理办法', '报名超计划实行摇号，不得提前录取', '民办学校', '2025年执行'),
]
for p in policies: ws.append(p)
style(ws, 5); wb.save(os.path.join(d, '招生政策文件.xlsx'))

mkzip(d, '01招生计划与学校信息.rar')

# ========== 压缩包2: 报名录取记录 ==========
print('=== 02 报名录取记录 ===')
d = os.path.join(INPUT_DIR, '_t2'); os.makedirs(d)

fig, ax = plt.subplots(figsize=(10, 6))
months = ['3月', '4月', '5月', '6月', '7月', '8月']
signups = [500, 1200, 3500, 2800, 800, 200]
ax.bar(months, signups, color='#4472C4', alpha=0.7)
ax.set_ylabel('报名人数'); ax.set_title('2025年招生报名月度分布')
for i, v in enumerate(signups):
    ax.text(i, v+50, str(v), ha='center', fontsize=9)
plt.tight_layout(); fig.savefig(os.path.join(d, '招生报名趋势.png'), dpi=150); plt.close()

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '报名录取记录'
ws.append(['序号', '报名编号', '学生姓名', '学校编号', '户籍片区', '房产片区', '报名类型', '录取状态', '录取时间', '备注'])
# 200条: 30跨片 + 25提前录取 + 20违规择校 + 15特长生违规 + 10正常跨片(政策允许) + 100正常
for i in range(1, 201):
    scode = f'XX{random.randint(1,25):03d}'
    area = random.randint(1, 12)
    btype = random.choice(['划片入学', '特长生', '民办摇号', '转学插班'])
    if i <= 30:
        btype = '划片入学'; note = '户籍房产均不在划片范围，跨片录取违规'
    elif i <= 55:
        btype = '民办摇号'; note = '在摇号前已提前录取，违规'
    elif i <= 75:
        btype = '划片入学'; note = '缴纳择校费后录取，违规'
    elif i <= 90:
        btype = '特长生'; note = '特长生认定材料不全，违规'
    elif i <= 100:
        btype = '划片入学'; note = '符合随迁子女政策，跨片合理（反直觉正常项）'
    else:
        note = ''
    ws.append([i, f'BM{i:04d}', f'学生{i:03d}', scode, f'片区{area}', f'片区{area if i>30 else (area+1)%12+1}',
               btype, '已录取', f'2025-{random.randint(5,8):02d}-{random.randint(1,28):02d}', note])
style(ws, 10); wb.save(os.path.join(d, '报名录取记录.xlsx'))

mkzip(d, '02报名录取记录.rar')

# ========== 压缩包3: 特长生与编班 ==========
print('=== 03 特长生与编班 ===')
d = os.path.join(INPUT_DIR, '_t3'); os.makedirs(d)

fig, ax = plt.subplots(figsize=(8, 6))
labels = ['艺术', '体育', '科技', '其他']
sizes = [40, 35, 15, 10]
colors = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000']
ax.pie(sizes, labels=labels, autopct='%1.0f%%', colors=colors, startangle=90)
ax.set_title('特长生类别分布')
plt.tight_layout(); fig.savefig(os.path.join(d, '特长生类别分布.png'), dpi=150); plt.close()

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '特长生认定'
ws.append(['序号', '认定编号', '学生姓名', '学校编号', '特长类别', '认定证书', '认定等级', '评审委员会', '公示情况', '备注'])
# 50人: 10材料不全 + 8证书造假 + 5比例超标 + 4正常合理 + 23正常
for i in range(1, 51):
    scode = f'XX{random.randint(1,25):03d}'
    cat = random.choice(['艺术', '体育', '科技', '其他'])
    cert = f'{cat}等级{random.randint(1,10)}级'
    level = random.choice(['国家级', '省级', '市级', '区级'])
    if i <= 10:
        note = '认定材料不全，缺少获奖证书原件'
    elif i <= 18:
        note = '认定证书经核实为伪造'
    elif i <= 23:
        note = '特长生比例超过5%上限'
    elif i <= 27:
        note = '证书等级高且材料齐全，认定合理（反直觉正常项）'
    else:
        note = ''
    ws.append([i, f'TC{i:03d}', f'学生{i:03d}', scode, cat, cert, level,
               f'评审组{random.randint(1,5)}', '已公示' if i > 10 else '未公示', note])
style(ws, 10); wb.save(os.path.join(d, '特长生认定记录.xlsx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '编班方案'
ws.append(['序号', '方案编号', '学校编号', '班级数', '总人数', '平均班额', '编班方式', '是否公示', '是否有重点班', '备注'])
# 25所: 8重点班 + 6不均衡 + 3正常合理(分层教学允许) + 8正常
for i in range(1, 26):
    scode = f'XX{i:03d}'
    class_num = random.randint(4, 12)
    total = random.randint(180, 600)
    avg = total // class_num
    if i <= 8:
        method = '按成绩分班'; focus = '是'; note = '分重点班和普通班，违规'
    elif i <= 14:
        method = '随机编班'; focus = '否'; note = '各班人数差异超过10%，不均衡'
    elif i <= 17:
        method = '分层走班'; focus = '否'; note = '分层教学有教育局批文，合理（反直觉正常项）'
    else:
        method = '随机编班'; focus = '否'; note = ''
    ws.append([i, f'BB{i:03d}', scode, class_num, total, avg, method, '是' if i > 8 else '否', focus, note])
style(ws, 10); wb.save(os.path.join(d, '编班方案.xlsx'))

mkzip(d, '03特长生与编班.rar')

# ========== 压缩包4: 收费与转学 ==========
print('=== 04 收费与转学 ===')
d = os.path.join(INPUT_DIR, '_t4'); os.makedirs(d)

fig, ax = plt.subplots(figsize=(10, 6))
types = ['择校费', '捐资助学', '住宿费', '校服费', '教辅费', '其他']
amounts = [120, 80, 45, 30, 25, 15]
colors = ['#FF6B6B', '#FFC000', '#4472C4', '#70AD47', '#A5A5A5', '#BDC3C7']
ax.bar(types, amounts, color=colors, width=0.6)
ax.set_ylabel('金额（万元）'); ax.set_title('2025年招生相关收费分布')
for i, v in enumerate(amounts):
    ax.text(i, v+3, f'{v}万', ha='center', fontsize=9)
plt.xticks(rotation=15, ha='right'); plt.tight_layout()
fig.savefig(os.path.join(d, '招生收费分布.png'), dpi=150); plt.close()

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '收费记录'
ws.append(['序号', '收费编号', '学校编号', '学生姓名', '收费项目', '收费金额(元)', '收费日期', '是否公示', '是否开发票', '备注'])
# 150条: 25择校费 + 15捐资助学 + 20乱收费 + 10正常住宿 + 80正常
for i in range(1, 151):
    scode = f'XX{random.randint(1,25):03d}'
    if i <= 25:
        item = '择校费'; amt = random.randint(10000, 50000); note = '违规收取择校费'
    elif i <= 40:
        item = '捐资助学'; amt = random.randint(5000, 30000); note = '强制捐资助学，违规'
    elif i <= 60:
        item = '赞助费'; amt = random.randint(3000, 20000); note = '无收费依据的乱收费'
    elif i <= 70:
        item = '住宿费'; amt = random.randint(500, 1500); note = '住宿生正常收费（反直觉正常项）'
    else:
        item = random.choice(['校服费', '教辅费', '餐费']); amt = random.randint(100, 800); note = ''
    ws.append([i, f'SF{i:04d}', scode, f'学生{i:03d}', item, amt,
               f'2025-{random.randint(6,9):02d}-{random.randint(1,28):02d}',
               '否' if i <= 60 else '是', '否' if i <= 40 else '是', note])
style(ws, 10); wb.save(os.path.join(d, '收费记录.xlsx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '转学记录'
ws.append(['序号', '转学编号', '学生姓名', '原学校', '转入学校', '转学原因', '申请日期', '审批状态', '是否有空挂', '备注'])
# 80条: 15违规转学 + 10空挂学籍 + 8未审批 + 5正常合理 + 42正常
reasons = ['家庭搬迁', '工作调动', '户籍迁移', '照顾子女', '其他']
for i in range(1, 81):
    if i <= 15:
        reason = '其他'; note = '转学原因不实，违规转学'
    elif i <= 25:
        reason = '家庭搬迁'; note = '学生未实际到校，空挂学籍'
    elif i <= 33:
        reason = '工作调动'; note = '转学未经教育局审批'
    elif i <= 38:
        reason = '户籍迁移'; note = '材料齐全，转学合理（反直觉正常项）'
    else:
        reason = random.choice(reasons); note = ''
    ws.append([i, f'ZX{i:04d}', f'学生{i:03d}', f'XX{random.randint(1,25):03d}',
               f'XX{random.randint(1,25):03d}', reason, f'2025-{random.randint(1,9):02d}-{random.randint(1,28):02d}',
               '已审批' if i > 25 else '未审批', '否' if i > 15 else '是', note])
style(ws, 10); wb.save(os.path.join(d, '转学记录.xlsx'))

mkzip(d, '04收费与转学.rar')

# ========== 压缩包5: 招生管理制度 ==========
print('=== 05 招生管理制度 ===')
d = os.path.join(INPUT_DIR, '_t5'); os.makedirs(d)
policies = [
    {'name': '01招生工作管理办法.docx', 'title': '招生工作管理办法',
     'secs': ['总则', '招生计划', '报名录取', '编班管理', '收费管理', '监督检查', '附则'],
     'issues': '问题：1.超额招生无审批流程；2.跨片招生无审核；3.违规招生无问责'},
    {'name': '02特长生招生管理办法.docx', 'title': '特长生招生管理办法',
     'secs': ['总则', '认定标准', '认定程序', '评审管理', '公示监督', '附则'],
     'issues': '问题：1.认定标准不明确；2.评审委员会组成不规范；3.公示制度不落实'},
    {'name': '03均衡编班管理办法.docx', 'title': '均衡编班管理办法',
     'secs': ['总则', '编班原则', '编班程序', '监督管理', '附则'],
     'issues': '问题：1.分层教学无明确规定；2.编班方案不公示；3.重点班查处不力'},
    {'name': '04招生收费管理办法.docx', 'title': '招生收费管理办法',
     'secs': ['总则', '收费项目', '收费标准', '票据管理', '监督检查', '附则'],
     'issues': '问题：1.择校费禁而不止；2.捐资助学不规范；3.乱收费无问责'},
    {'name': '05转学插班管理办法.docx', 'title': '转学插班管理办法',
     'secs': ['总则', '转学条件', '转学程序', '学籍管理', '附则'],
     'issues': '问题：1.转学条件审核不严；2.空挂学籍无核查；3.审批程序不规范'},
]
for p in policies:
    doc = Document(); doc.add_heading(p['title'], level=0)
    for sec in p['secs']:
        doc.add_heading(sec, level=1)
        doc.add_paragraph(f'{sec}具体条款内容详见正式制度文件。')
    doc.add_heading('现有制度问题诊断', level=1)
    doc.add_paragraph(p['issues'])
    doc.save(os.path.join(d, p['name']))
mkzip(d, '05招生管理制度.rar')

# ========== 压缩包6: 管理层要求 ==========
print('=== 06 管理层要求 ===')
d = os.path.join(INPUT_DIR, '_t6'); os.makedirs(d)
doc = Document(); doc.add_heading('管理层要求汇总', level=0)
reqs = [
    (1, '教育局长', '超额招生不追责', '超额招生缓解学位紧张'),
    (2, '基教科长', '跨片招生不处理', '跨片招生照顾优质生源'),
    (3, '教育局长', '特长生认定不严格', '特长生是学校特色'),
    (4, '基教科长', '重点班不取消', '重点班有利于升学'),
    (5, '财务科长', '择校费不退还', '择校费是学校收入'),
    (6, '基教科长', '捐资助学不叫停', '捐资助学补充办学经费'),
    (7, '教育局长', '空挂学籍不清理', '清理影响学校考核'),
    (8, '基教科长', '乱收费不查处', '查处影响学校积极性'),
    (9, '教育局长', '转学审批不严格', '转学都是合理需求'),
    (10, '基教科长', '编班方案不公示', '公示引发家长争议'),
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
ax.text(0.05, 0.75, '收件人：基教科、财务科、各学校', fontsize=11, transform=ax.transAxes)
ax.text(0.05, 0.6, '主题：关于招生入学审计工作的指导意见', fontsize=11, transform=ax.transAxes)
ax.text(0.05, 0.3, '各位：\n招生工作涉及面广，为维护学校办学积极性，\n请暂缓执行超额招生追责、跨片招生处理、\n重点班取消、择校费退还等措施，先维持现状。\n\n教育局长\n2025年9月28日', fontsize=10, transform=ax.transAxes)
rect = plt.Rectangle((0,0),1,1, transform=ax.transAxes, fill=False, edgecolor='gray')
ax.add_patch(rect); plt.tight_layout()
fig.savefig(os.path.join(d, '教育局长邮件.png'), dpi=150); plt.close()
mkzip(d, '06管理层要求.rar')

print('\n=== 输入附件生成完成 ===')
