#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""中小学教育信息化建设与专项资金使用审计 - 输入附件生成"""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from docx import Document
import zipfile, os, random, shutil

random.seed(33)
plt.rcParams['font.sans-serif'] = ['WenQuanYi Micro Hei', 'SimHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

BASE = '/workspace/教育信息化审计_附件'
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

# ========== 压缩包1: 学校与信息化制度 ==========
print('=== 01 学校与信息化制度 ===')
d = os.path.join(INPUT_DIR, '_t1'); os.makedirs(d)

fig, ax = plt.subplots(figsize=(10, 6))
types = ['小学', '初中', '九年一贯制', '完全中学']
counts = [14, 9, 5, 2]
colors = ['#4472C4', '#70AD47', '#ED7D31', '#A5A5A5']
ax.bar(types, counts, color=colors, width=0.5)
ax.set_ylabel('学校数'); ax.set_title('30所学校类型分布')
for i, v in enumerate(counts):
    ax.text(i, v+0.2, str(v), ha='center', fontsize=9)
plt.tight_layout(); fig.savefig(os.path.join(d, '学校类型分布.png'), dpi=150); plt.close()

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '学校清单'
ws.append(['序号', '学校编号', '学校名称', '学校类型', '学生人数', '信息化投入', '信息化部门', '状态标记'])
# 30所: 8挪用 + 6闲置 + 5培训形式 + 3合规保留 + 8正常
for i in range(1, 31):
    if i <= 8:
        state = '信息化专项资金被挪用'
    elif i <= 14:
        state = '信息化设备大量闲置'
    elif i <= 19:
        state = '教师信息技术培训流于形式'
    elif i <= 22:
        state = '设备使用率高，信息化成效好（反直觉正常项）'
    else:
        state = '正常'
    ws.append([i, f'XX{i:03d}', f'第{i}学校', random.choice(['小学', '初中', '九年一贯制', '完全中学']),
               random.randint(200, 3000), random.randint(10, 200), '信息中心', state])
style(ws, 8); wb.save(os.path.join(d, '学校清单.xlsx'))

mkzip(d, '01学校与信息化制度.rar')

# ========== 压缩包2: 设备采购记录 ==========
print('=== 02 设备采购记录 ===')
d = os.path.join(INPUT_DIR, '_t2'); os.makedirs(d)

fig, ax = plt.subplots(figsize=(10, 6))
cats = ['未招标', '超预算', '验收不合格', '合规保留', '正常']
counts = [25, 15, 10, 10, 40]
colors = ['#FF6B6B', '#FF9F43', '#FECA57', '#1DD1A1', '#48DBFB']
ax.bar(cats, counts, color=colors, width=0.5)
ax.set_ylabel('记录数'); ax.set_title('100条设备采购记录分类')
for i, v in enumerate(counts):
    ax.text(i, v+1, str(v), ha='center', fontsize=9)
plt.tight_layout(); fig.savefig(os.path.join(d, '设备采购分类.png'), dpi=150); plt.close()

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '设备采购记录'
ws.append(['序号', '采购编号', '所属学校', '设备名称', '采购数量', '预算金额', '实际金额', '采购方式', '验收状态', '状态标记'])
# 100条: 25未招标 + 15超预算 + 10验收不合格 + 10合规保留 + 40正常
devices = ['多媒体一体机', '计算机', '服务器', '交换机', '路由器', '投影仪', '电子白板', '监控设备']
for i in range(1, 101):
    school = f'XX{random.randint(1,30):03d}'
    dev = random.choice(devices)
    budget = random.randint(1, 50)
    if i <= 25:
        state = '未按规定进行政府采购'
        way = '自行采购'
    elif i <= 40:
        state = '实际采购金额超出预算'
        way = '公开招标'
    elif i <= 50:
        state = '设备验收不合格，未整改'
        way = '公开招标'
    elif i <= 60:
        state = '单一来源采购理由充分（反直觉正常项）'
        way = '单一来源'
    else:
        state = '正常'
        way = '公开招标'
    ws.append([i, f'CG{i:04d}', school, dev, random.randint(1, 100), budget,
               budget + random.randint(0, 20) if i > 40 else budget + random.randint(5, 30),
               way, '合格' if i > 50 else '不合格', state])
style(ws, 10); wb.save(os.path.join(d, '设备采购记录.xlsx'))

mkzip(d, '02设备采购记录.rar')

# ========== 压缩包3: 软件采购记录 ==========
print('=== 03 软件采购记录 ===')
d = os.path.join(INPUT_DIR, '_t3'); os.makedirs(d)

fig, ax = plt.subplots(figsize=(10, 6))
cats = ['重复采购', '使用率低', '未招标', '合规保留', '正常']
counts = [12, 10, 8, 5, 25]
colors = ['#FF6B6B', '#FF9F43', '#FECA57', '#1DD1A1', '#48DBFB']
ax.bar(cats, counts, color=colors, width=0.5)
ax.set_ylabel('记录数'); ax.set_title('60条软件采购记录分类')
for i, v in enumerate(counts):
    ax.text(i, v+0.5, str(v), ha='center', fontsize=9)
plt.tight_layout(); fig.savefig(os.path.join(d, '软件采购分类.png'), dpi=150); plt.close()

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '软件采购记录'
ws.append(['序号', '采购编号', '所属学校', '软件名称', '授权方式', '采购金额', '采购方式', '使用率', '状态标记'])
# 60条: 12重复 + 10使用率低 + 8未招标 + 5合规保留 + 25正常
softs = ['教学管理系统', '办公软件', '杀毒软件', '数据库', '课件制作', '阅卷系统']
for i in range(1, 61):
    school = f'XX{random.randint(1,30):03d}'
    soft = random.choice(softs)
    if i <= 12:
        state = '与已采购软件功能重复'
    elif i <= 22:
        state = '软件使用率极低，基本闲置'
    elif i <= 30:
        state = '未按规定进行政府采购'
    elif i <= 35:
        state = '正版软件授权，合规（反直觉正常项）'
    else:
        state = '正常'
    ws.append([i, f'RJ{i:04d}', school, soft, random.choice(['永久授权', '年授权', '订阅']),
               random.randint(1, 30), '公开招标' if i > 30 else '自行采购',
               f'{random.randint(0, 100)}%', state])
style(ws, 9); wb.save(os.path.join(d, '软件采购记录.xlsx'))

mkzip(d, '03软件采购记录.rar')

# ========== 压缩包4: 设备使用台账 ==========
print('=== 04 设备使用台账 ===')
d = os.path.join(INPUT_DIR, '_t4'); os.makedirs(d)

fig, ax = plt.subplots(figsize=(10, 6))
cats = ['闲置', '使用率低', '未入账', '合规保留', '正常']
counts = [30, 25, 15, 10, 70]
colors = ['#FF6B6B', '#FF9F43', '#FECA57', '#1DD1A1', '#48DBFB']
ax.bar(cats, counts, color=colors, width=0.5)
ax.set_ylabel('记录数'); ax.set_title('150条设备使用台账分类')
for i, v in enumerate(counts):
    ax.text(i, v+1, str(v), ha='center', fontsize=9)
plt.tight_layout(); fig.savefig(os.path.join(d, '设备使用分类.png'), dpi=150); plt.close()

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '设备使用台账'
ws.append(['序号', '台账编号', '所属学校', '设备名称', '设备编号', '启用日期', '月使用时长', '保管人', '状态标记'])
# 150条: 30闲置 + 25使用率低 + 15未入账 + 10合规保留 + 70正常
for i in range(1, 151):
    school = f'XX{random.randint(1,30):03d}'
    dev = random.choice(devices)
    if i <= 30:
        state = '设备长期闲置，未投入使用'
        hours = 0
    elif i <= 55:
        state = '设备使用率远低于预期'
        hours = random.randint(1, 10)
    elif i <= 70:
        state = '设备未登记入账'
        hours = random.randint(0, 20)
    elif i <= 80:
        state = '校际共享设备，使用合理（反直觉正常项）'
        hours = random.randint(20, 80)
    else:
        state = '正常'
        hours = random.randint(30, 120)
    ws.append([i, f'SY{i:04d}', school, dev, f'DEV{i:05d}', f'2024-{random.randint(1,12):02d}-01',
               hours, f'保管人{i}', state])
style(ws, 9); wb.save(os.path.join(d, '设备使用台账.xlsx'))

mkzip(d, '04设备使用台账.rar')

# ========== 压缩包5: 培训记录与网络安全 ==========
print('=== 05 培训记录与网络安全 ===')
d = os.path.join(INPUT_DIR, '_t5'); os.makedirs(d)

fig, ax = plt.subplots(figsize=(10, 6))
months = [f'{m}月' for m in range(1, 13)]
counts = [2, 1, 3, 2, 4, 3, 2, 5, 3, 4, 3, 2]
ax.bar(months, counts, color='#4472C4', alpha=0.7)
ax.set_ylabel('培训次数'); ax.set_title('2024年信息技术培训月度分布')
for i, v in enumerate(counts):
    if v > 0:
        ax.text(i, v+0.1, str(v), ha='center', fontsize=9)
plt.xticks(rotation=45); plt.tight_layout()
fig.savefig(os.path.join(d, '培训分布.png'), dpi=150); plt.close()

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '培训记录'
ws.append(['序号', '培训编号', '所属学校', '培训主题', '参训人数', '培训时长', '考核方式', '考核结果', '状态标记'])
# 40条: 12流于形式 + 8未考核 + 5超范围 + 3合规保留 + 12正常
topics = ['多媒体教学', '办公软件', '网络安全', '数据分析', '在线教育', '信息素养']
for i in range(1, 41):
    school = f'XX{random.randint(1,30):03d}'
    topic = random.choice(topics)
    if i <= 12:
        state = '培训流于形式，签到代签'
        result = '未考核'
    elif i <= 20:
        state = '培训后未组织考核'
        result = '未考核'
    elif i <= 25:
        state = '培训内容超出信息化范围'
        result = '合格'
    elif i <= 28:
        state = '培训实效显著，教师应用能力提升（反直觉正常项）'
        result = '优秀'
    else:
        state = '正常'
        result = '合格'
    ws.append([i, f'PX{i:03d}', school, topic, random.randint(10, 100),
               f'{random.randint(1, 8)}小时', '考试' if i > 20 else '签到',
               result, state])
style(ws, 9); wb.save(os.path.join(d, '培训记录.xlsx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '网络安全检查报告'
ws.append(['序号', '报告编号', '所属学校', '检查日期', '等保级别', '等保备案', '隐患数量', '整改情况', '数据备份', '状态标记'])
# 15份: 6等保未备案 + 5隐患未整改 + 2备份缺失 + 2合规保留
for i in range(1, 16):
    school = f'XX{random.randint(1,30):03d}'
    if i <= 6:
        state = '未按规定进行网络安全等级保护备案'
        record = '未备案'
    elif i <= 11:
        state = '安全隐患未按要求整改'
        record = '已备案'
    elif i <= 13:
        state = '重要数据未定期备份'
        record = '已备案'
    else:
        state = '等保合规，隐患已整改（反直觉正常项）'
        record = '已备案'
    ws.append([i, f'AQ{i:03d}', school, f'2024-{random.randint(1,12):02d}-15',
               random.choice(['二级', '三级']), record, random.randint(0, 10),
               '已整改' if i > 11 else '未整改',
               '有备份' if i > 13 else '无备份', state])
style(ws, 10); wb.save(os.path.join(d, '网络安全检查报告.xlsx'))

mkzip(d, '05培训记录与网络安全.rar')

# ========== 压缩包6: 信息化制度与管理层要求 ==========
print('=== 06 信息化制度与管理层要求 ===')
d = os.path.join(INPUT_DIR, '_t6'); os.makedirs(d)
policies = [
    {'name': '01教育信息化专项资金管理办法.docx', 'title': '教育信息化专项资金管理办法',
     'secs': ['总则', '资金来源', '使用范围', '采购管理', '监督检查', '附则'],
     'issues': '问题：1.使用范围不严；2.采购不规范；3.监督缺失'},
    {'name': '02设备采购管理办法.docx', 'title': '设备采购管理办法',
     'secs': ['总则', '采购方式', '采购程序', '验收管理', '资产管理', '附则'],
     'issues': '问题：1.采购方式不规范；2.验收不严；3.资产不入账'},
    {'name': '03软件采购管理办法.docx', 'title': '软件采购管理办法',
     'secs': ['总则', '需求论证', '采购方式', '授权管理', '使用评估', '附则'],
     'issues': '问题：1.需求论证缺失；2.重复采购；3.使用率低'},
    {'name': '04教师信息技术培训办法.docx', 'title': '教师信息技术培训办法',
     'secs': ['总则', '培训内容', '培训方式', '考核管理', '效果评估', '附则'],
     'issues': '问题：1.培训流于形式；2.考核缺失；3.效果不评估'},
    {'name': '05网络安全管理办法.docx', 'title': '网络安全管理办法',
     'secs': ['总则', '等级保护', '安全防护', '数据备份', '应急处置', '附则'],
     'issues': '问题：1.等保未备案；2.防护不到位；3.备份缺失'},
]
for p in policies:
    doc = Document(); doc.add_heading(p['title'], level=0)
    for sec in p['secs']:
        doc.add_heading(sec, level=1)
        doc.add_paragraph(f'{sec}具体条款内容详见正式制度文件。')
    doc.add_heading('现有制度问题诊断', level=1)
    doc.add_paragraph(p['issues'])
    doc.save(os.path.join(d, p['name']))

doc = Document(); doc.add_heading('管理层要求汇总', level=0)
reqs = [
    (1, '教育局副局长', '信息化专项资金挪用', '统筹使用更合理'),
    (2, '信息中心主任', '设备采购不招标', '招标耽误时间'),
    (3, '学校校长', '闲置设备不处理', '处理影响资产价值'),
    (4, '信息中心主任', '软件重复采购不清理', '清理工作量大'),
    (5, '学校校长', '培训不考核', '考核影响教师积极性'),
    (6, '信息中心主任', '网络安全等保不备案', '备案增加费用'),
    (7, '教育局副局长', '安全隐患不整改', '整改影响教学'),
    (8, '学校校长', '设备未入账不补登', '补登增加工作量'),
    (9, '信息中心主任', '数据备份不做', '备份占用存储'),
    (10, '教育局副局长', '培训流于形式不整改', '整改影响培训指标'),
]
for num, src, content, reason in reqs:
    doc.add_heading(f'要求{num}', level=1)
    p = doc.add_paragraph(); p.add_run('来源：').bold = True; p.add_run(src)
    p = doc.add_paragraph(); p.add_run('要求：').bold = True; p.add_run(content)
    p = doc.add_paragraph(); p.add_run('理由：').bold = True; p.add_run(reason)
    doc.add_paragraph()
doc.save(os.path.join(d, '管理层要求.docx'))

fig, ax = plt.subplots(figsize=(10, 4)); ax.axis('off')
ax.text(0.05, 0.9, '发件人：信息中心', fontsize=11, transform=ax.transAxes)
ax.text(0.05, 0.75, '收件人：各学校、信息中心', fontsize=11, transform=ax.transAxes)
ax.text(0.05, 0.6, '主题：关于教育信息化建设的指导意见', fontsize=11, transform=ax.transAxes)
ax.text(0.05, 0.3, '各位：\n教育信息化建设任务重，为保障工作顺利开展，\n请简化采购流程、闲置设备暂不处理、\n培训和网络安全可适当放宽要求。\n\n教育局副局长\n2025年9月28日', fontsize=10, transform=ax.transAxes)
rect = plt.Rectangle((0,0),1,1, transform=ax.transAxes, fill=False, edgecolor='gray')
ax.add_patch(rect); plt.tight_layout()
fig.savefig(os.path.join(d, '教育局副局长邮件.png'), dpi=150); plt.close()
mkzip(d, '06信息化制度与管理层要求.rar')

print('\n=== 输入附件生成完成 ===')
