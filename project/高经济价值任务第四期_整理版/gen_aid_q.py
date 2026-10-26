#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""中小学学生资助与教育扶贫资金使用审计 - 输入附件生成"""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from docx import Document
import zipfile, os, random, shutil

random.seed(77)
plt.rcParams['font.sans-serif'] = ['WenQuanYi Micro Hei', 'SimHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

BASE = '/workspace/学生资助扶贫审计_附件'
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

# ========== 压缩包1: 学校与资助制度 ==========
print('=== 01 学校与资助制度 ===')
d = os.path.join(INPUT_DIR, '_t1'); os.makedirs(d)

fig, ax = plt.subplots(figsize=(10, 6))
types = ['小学', '初中', '九年一贯制', '完全中学']
counts = [10, 6, 3, 1]
colors = ['#4472C4', '#70AD47', '#ED7D31', '#A5A5A5']
ax.bar(types, counts, color=colors, width=0.5)
ax.set_ylabel('学校数'); ax.set_title('20所学校类型分布')
for i, v in enumerate(counts):
    ax.text(i, v+0.1, str(v), ha='center', fontsize=9)
plt.tight_layout(); fig.savefig(os.path.join(d, '学校类型分布.png'), dpi=150); plt.close()

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '学校清单'
ws.append(['序号', '学校编号', '学校名称', '学校类型', '学生人数', '资助人数', '资助管理部门', '状态标记'])
# 20所: 6制度缺失 + 5公示不到位 + 4档案不全 + 3合规保留 + 2正常
for i in range(1, 21):
    if i <= 6:
        state = '资助管理制度缺失'
    elif i <= 11:
        state = '资助名单公示不到位'
    elif i <= 15:
        state = '资助档案资料不全'
    elif i <= 18:
        state = '制度健全公示规范（反直觉正常项）'
    else:
        state = '正常'
    ws.append([i, f'XX{i:03d}', f'第{i}学校', random.choice(['小学', '初中', '九年一贯制', '完全中学']),
               random.randint(200, 2000), random.randint(20, 200), '资助管理办公室', state])
style(ws, 8); wb.save(os.path.join(d, '学校清单.xlsx'))

mkzip(d, '01学校与资助制度.rar')

# ========== 压缩包2: 资助学生名册 ==========
print('=== 02 资助学生名册 ===')
d = os.path.join(INPUT_DIR, '_t2'); os.makedirs(d)

fig, ax = plt.subplots(figsize=(10, 6))
cats = ['认定不精准', '重复享受', '已脱贫仍享受', '合规保留', '正常']
counts = [50, 30, 20, 30, 170]
colors = ['#FF6B6B', '#FF9F43', '#FECA57', '#1DD1A1', '#48DBFB']
ax.bar(cats, counts, color=colors, width=0.5)
ax.set_ylabel('人数'); ax.set_title('300名资助学生分类')
for i, v in enumerate(counts):
    ax.text(i, v+2, str(v), ha='center', fontsize=9)
plt.tight_layout(); fig.savefig(os.path.join(d, '资助学生分类.png'), dpi=150); plt.close()

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '资助学生名册'
ws.append(['序号', '学生编号', '学生姓名', '身份证号', '所属学校', '家庭情况', '资助类型', '资助金额', '认定日期', '状态标记'])
# 300人: 50认定不精准 + 30重复享受 + 20已脱贫仍享受 + 30合规保留 + 170正常
fam_types = ['建档立卡', '低保', '残疾', '孤儿', '特困供养', '一般困难']
fund_types = ['助学金', '免学杂费', '营养餐', '寄宿生补助']
for i in range(1, 301):
    fam = random.choice(fam_types)
    ftype = random.choice(fund_types)
    school = f'XX{random.randint(1,20):03d}'
    if i <= 50:
        state = '资助认定条件不符，家庭收入高于标准'
    elif i <= 80:
        state = '同一学生在多所学校重复享受资助'
    elif i <= 100:
        state = '家庭已脱贫仍继续享受资助'
    elif i <= 130:
        state = '边缘户家庭，认定合理（反直觉正常项）'
    else:
        state = '正常'
    ws.append([i, f'XS{i:04d}', f'学生{i}', f'1101012008{i:08d}', school, fam, ftype,
               random.choice([500, 800, 1000, 1250, 2000]), f'2024-09-01', state])
style(ws, 10); wb.save(os.path.join(d, '资助学生名册.xlsx'))

mkzip(d, '02资助学生名册.rar')

# ========== 压缩包3: 资助发放记录 ==========
print('=== 03 资助发放记录 ===')
d = os.path.join(INPUT_DIR, '_t3'); os.makedirs(d)

fig, ax = plt.subplots(figsize=(10, 6))
months = [f'{m}月' for m in range(1, 13)]
amounts = [80, 95, 85, 110, 90, 100, 105, 98, 88, 92, 102, 95]
ax.plot(months, amounts, marker='o', color='#4472C4', linewidth=2)
ax.fill_between(range(12), amounts, alpha=0.2, color='#4472C4')
ax.set_ylabel('发放金额(万元)'); ax.set_title('2024年资助金月度发放趋势')
plt.xticks(rotation=45); plt.tight_layout()
fig.savefig(os.path.join(d, '资助发放趋势.png'), dpi=150); plt.close()

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '资助发放记录'
ws.append(['序号', '发放编号', '学生编号', '资助类型', '应发金额', '实发金额', '发放日期', '发放方式', '签收情况', '状态标记'])
# 200条: 25未及时发放 + 15金额不足 + 10截留挪用 + 10合规保留 + 140正常
for i in range(1, 201):
    xs = f'XS{random.randint(1,300):04d}'
    ftype = random.choice(fund_types)
    payable = random.choice([500, 800, 1000, 1250, 2000])
    if i <= 25:
        state = '资助金未及时发放，超过规定时限'
        actual = 0
    elif i <= 40:
        state = '实发金额低于应发金额'
        actual = payable - random.randint(50, 200)
    elif i <= 50:
        state = '资助金被截留挪用'
        actual = 0
    elif i <= 60:
        state = '提前发放符合规定（反直觉正常项）'
        actual = payable
    else:
        state = '正常'
        actual = payable
    ws.append([i, f'FF{i:04d}', xs, ftype, payable, actual,
               f'2024-{random.randint(1,12):02d}-{random.randint(1,28):02d}',
               '银行转账', '已签收' if i > 25 else '未签收', state])
style(ws, 10); wb.save(os.path.join(d, '资助发放记录.xlsx'))

mkzip(d, '03资助发放记录.rar')

# ========== 压缩包4: 免学杂费记录 ==========
print('=== 04 免学杂费记录 ===')
d = os.path.join(INPUT_DIR, '_t4'); os.makedirs(d)

fig, ax = plt.subplots(figsize=(10, 6))
cats = ['超范围免', '重复免', '应免未免', '合规保留', '正常']
counts = [15, 10, 8, 7, 60]
colors = ['#FF6B6B', '#FF9F43', '#FECA57', '#1DD1A1', '#48DBFB']
ax.bar(cats, counts, color=colors, width=0.5)
ax.set_ylabel('记录数'); ax.set_title('100条免学杂费记录分类')
for i, v in enumerate(counts):
    ax.text(i, v+1, str(v), ha='center', fontsize=9)
plt.tight_layout(); fig.savefig(os.path.join(d, '免学杂费分类.png'), dpi=150); plt.close()

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '免学杂费记录'
ws.append(['序号', '记录编号', '学生编号', '所属学校', '学段', '应免金额', '实免金额', '批准日期', '状态标记'])
# 100条: 15超范围免 + 10重复免 + 8应免未免 + 7合规保留 + 60正常
for i in range(1, 101):
    xs = f'XS{random.randint(1,300):04d}'
    school = f'XX{random.randint(1,20):03d}'
    stage = random.choice(['小学', '初中'])
    exempt = random.choice([200, 300, 500])
    if i <= 15:
        state = '非义务教育阶段违规免学杂费'
        actual = exempt
    elif i <= 25:
        state = '同一学生重复免学杂费'
        actual = exempt
    elif i <= 33:
        state = '符合条件学生未免学杂费'
        actual = 0
    elif i <= 40:
        state = '符合减免政策，合规（反直觉正常项）'
        actual = exempt
    else:
        state = '正常'
        actual = exempt
    ws.append([i, f'MZ{i:04d}', xs, school, stage, exempt, actual,
               f'2024-09-01', state])
style(ws, 9); wb.save(os.path.join(d, '免学杂费记录.xlsx'))

mkzip(d, '04免学杂费记录.rar')

# ========== 压缩包5: 扶贫资金拨付台账 ==========
print('=== 05 扶贫资金拨付台账 ===')
d = os.path.join(INPUT_DIR, '_t5'); os.makedirs(d)

fig, ax = plt.subplots(figsize=(10, 6))
sources = ['中央财政', '省级财政', '市级财政', '县级财政']
amounts = [320, 180, 150, 100]
colors = ['#4472C4', '#70AD47', '#ED7D31', '#FFC000']
ax.bar(sources, amounts, color=colors, width=0.5)
ax.set_ylabel('金额(万元)'); ax.set_title('教育扶贫资金来源分布')
for i, v in enumerate(amounts):
    ax.text(i, v+5, str(v), ha='center', fontsize=9)
plt.tight_layout(); fig.savefig(os.path.join(d, '扶贫资金来源.png'), dpi=150); plt.close()

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '扶贫资金拨付台账'
ws.append(['序号', '拨付编号', '资金来源', '拨付对象', '拨付金额', '拨付日期', '到账日期', '用途', '状态标记'])
# 80条: 15挪用 + 10截留 + 10超范围 + 5合规保留 + 40正常
sources = ['中央财政', '省级财政', '市级财政', '县级财政']
uses = ['学生资助', '校舍维修', '教学设备', '教师培训', '营养餐']
for i in range(1, 81):
    src = random.choice(sources)
    obj = f'XX{random.randint(1,20):03d}'
    amount = random.randint(5, 50)
    use = random.choice(uses)
    if i <= 15:
        state = '扶贫资金被挪用至非教育用途'
    elif i <= 25:
        state = '扶贫资金被截留未及时拨付'
    elif i <= 35:
        state = '扶贫资金超出规定使用范围'
    elif i <= 40:
        state = '专项用于建档立卡学生，合规（反直觉正常项）'
    else:
        state = '正常'
    ws.append([i, f'FP{i:04d}', src, obj, amount,
               f'2024-{random.randint(1,12):02d}-01',
               f'2024-{random.randint(1,12):02d}-15', use, state])
style(ws, 9); wb.save(os.path.join(d, '扶贫资金拨付台账.xlsx'))

mkzip(d, '05扶贫资金拨付台账.rar')

# ========== 压缩包6: 资助制度与管理层要求 ==========
print('=== 06 资助制度与管理层要求 ===')
d = os.path.join(INPUT_DIR, '_t6'); os.makedirs(d)
policies = [
    {'name': '01学生资助管理办法.docx', 'title': '学生资助管理办法',
     'secs': ['总则', '资助对象认定', '资助标准', '申请审批', '发放管理', '附则'],
     'issues': '问题：1.认定标准不明确；2.审批程序简化；3.发放不及时'},
    {'name': '02免学杂费管理办法.docx', 'title': '免学杂费管理办法',
     'secs': ['总则', '减免范围', '减免标准', '申请程序', '资金管理', '附则'],
     'issues': '问题：1.范围界定不清；2.重复减免；3.应免未免'},
    {'name': '03教育扶贫资金管理办法.docx', 'title': '教育扶贫资金管理办法',
     'secs': ['总则', '资金来源', '使用范围', '拨付管理', '监督检查', '附则'],
     'issues': '问题：1.使用范围不严；2.拨付不及时；3.监督缺失'},
    {'name': '04资助公示制度.docx', 'title': '资助公示制度',
     'secs': ['总则', '公示内容', '公示方式', '公示期限', '异议处理', '附则'],
     'issues': '问题：1.公示内容不全；2.公示期限不足；3.异议处理不规范'},
    {'name': '05资助档案管理制度.docx', 'title': '资助档案管理制度',
     'secs': ['总则', '档案内容', '建档要求', '保管期限', '查阅使用', '附则'],
     'issues': '问题：1.档案内容不全；2.建档不及时；3.保管不规范'},
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
    (1, '教育局副局长', '放宽资助认定标准', '让更多学生享受资助'),
    (2, '资助办主任', '资助金不及时发放', '集中发放节省时间'),
    (3, '学校校长', '挪用扶贫资金维修校舍', '校舍维修更急需'),
    (4, '资助办主任', '资助名单不公示', '公示影响学生隐私'),
    (5, '教育局副局长', '非义务教育阶段免学杂费', '减轻家长负担'),
    (6, '学校校长', '已脱贫学生继续享受资助', '避免学生心理落差'),
    (7, '资助办主任', '资助档案不齐全', '档案占用空间'),
    (8, '教育局副局长', '扶贫资金截留用于其他', '统筹使用更合理'),
    (9, '学校校长', '重复享受资助不清理', '清理工作量大'),
    (10, '资助办主任', '应免未免不整改', '整改影响学校收入'),
]
for num, src, content, reason in reqs:
    doc.add_heading(f'要求{num}', level=1)
    p = doc.add_paragraph(); p.add_run('来源：').bold = True; p.add_run(src)
    p = doc.add_paragraph(); p.add_run('要求：').bold = True; p.add_run(content)
    p = doc.add_paragraph(); p.add_run('理由：').bold = True; p.add_run(reason)
    doc.add_paragraph()
doc.save(os.path.join(d, '管理层要求.docx'))

fig, ax = plt.subplots(figsize=(10, 4)); ax.axis('off')
ax.text(0.05, 0.9, '发件人：教育局资助办', fontsize=11, transform=ax.transAxes)
ax.text(0.05, 0.75, '收件人：各学校、资助管理办公室', fontsize=11, transform=ax.transAxes)
ax.text(0.05, 0.6, '主题：关于学生资助工作的指导意见', fontsize=11, transform=ax.transAxes)
ax.text(0.05, 0.3, '各位：\n学生资助工作涉及面广，为保障工作顺利开展，\n请放宽资助认定标准、资助金集中发放、\n扶贫资金统筹使用、资助名单不公示。\n\n教育局副局长\n2025年9月28日', fontsize=10, transform=ax.transAxes)
rect = plt.Rectangle((0,0),1,1, transform=ax.transAxes, fill=False, edgecolor='gray')
ax.add_patch(rect); plt.tight_layout()
fig.savefig(os.path.join(d, '教育局副局长邮件.png'), dpi=150); plt.close()
mkzip(d, '06资助制度与管理层要求.rar')

print('\n=== 输入附件生成完成 ===')
