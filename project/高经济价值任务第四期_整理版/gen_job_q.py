#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""高校毕业生就业质量与就业数据真实性审计 - 输入附件生成"""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from docx import Document
import zipfile, os, random, shutil

random.seed(55)
plt.rcParams['font.sans-serif'] = ['WenQuanYi Micro Hei', 'SimHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

BASE = '/workspace/毕业生就业审计_附件'
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

# ========== 压缩包1: 高校与毕业生样本 ==========
print('=== 01 高校与毕业生样本 ===')
d = os.path.join(INPUT_DIR, '_t1'); os.makedirs(d)

fig, ax = plt.subplots(figsize=(10, 6))
types = ['综合', '理工', '医药', '农林', '师范', '财经']
counts = [5, 6, 3, 2, 2, 2]
colors = ['#4472C4', '#70AD47', '#ED7D31', '#A5A5A5', '#FF6B6B', '#9B59B6']
ax.bar(types, counts, color=colors, width=0.5)
ax.set_ylabel('高校数'); ax.set_title('20所高校类型分布')
for i, v in enumerate(counts):
    ax.text(i, v+0.1, str(v), ha='center', fontsize=9)
plt.tight_layout(); fig.savefig(os.path.join(d, '高校类型分布.png'), dpi=150); plt.close()

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '高校清单'
ws.append(['序号', '高校编号', '高校名称', '高校类型', '毕业生人数', '就业率', '就业部门', '状态标记'])
univs = ['理工大学', '科技大学', '医科大学', '农业大学', '师范大学', '财经大学',
         '工业大学', '化工大学', '海洋大学', '交通大学', '建筑大学', '轻工大学',
         '纺织大学', '邮电大学', '地质大学', '林业大学', '中医药大学', '外国语大学',
         '政法大学', '美术学院']
for i in range(1, 21):
    if i <= 6:
        state = '就业率数据造假'
    elif i <= 11:
        state = '就业数据不完整'
    elif i <= 15:
        state = '就业指导服务缺失'
    elif i <= 18:
        state = '就业率真实偏高（反直觉正常项）'
    else:
        state = '正常'
    ws.append([i, f'GX{i:03d}', univs[i-1], random.choice(['综合', '理工', '医药', '农林', '师范', '财经']),
               random.randint(500, 8000), f'{random.uniform(75, 98):.1f}%', '就业指导中心', state])
style(ws, 8); wb.save(os.path.join(d, '高校清单.xlsx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '毕业生样本'
ws.append(['序号', '毕业生编号', '姓名', '身份证号', '所属高校', '专业', '学历', '就业状态', '就业单位', '状态标记'])
# 500人: 80造假 + 60虚高 + 50灵活不规范 + 40质量差 + 30合规保留 + 240正常
for i in range(1, 501):
    univ = f'GX{random.randint(1,20):03d}'
    major = random.choice(['计算机', '机械', '医学', '农学', '教育', '金融', '法学', '英语', '艺术', '土木'])
    edu = random.choice(['专科', '本科', '硕士', '博士'])
    if i <= 80:
        state = '虚假就业证明，实际未就业'
    elif i <= 140:
        state = '未就业登记为已就业，就业率虚高'
    elif i <= 190:
        state = '灵活就业认定不规范'
    elif i <= 230:
        state = '就业质量差，薪酬低于当地最低工资'
    elif i <= 260:
        state = '灵活就业真实有效（反直觉正常项）'
    else:
        state = '正常'
    ws.append([i, f'BY{i:04d}', f'毕业生{i}', f'1101012002{i:08d}', univ, major, edu,
               '已就业' if i > 140 else '待就业', f'某公司{i}', state])
style(ws, 10); wb.save(os.path.join(d, '毕业生样本.xlsx'))

mkzip(d, '01高校与毕业生样本.rar')

# ========== 压缩包2: 就业协议 ==========
print('=== 02 就业协议 ===')
d = os.path.join(INPUT_DIR, '_t2'); os.makedirs(d)

fig, ax = plt.subplots(figsize=(10, 6))
cats = ['虚假协议', '不规范', '单位不存在', '合规保留', '正常']
counts = [60, 40, 30, 20, 150]
colors = ['#FF6B6B', '#FF9F43', '#FECA57', '#1DD1A1', '#48DBFB']
ax.bar(cats, counts, color=colors, width=0.5)
ax.set_ylabel('份数'); ax.set_title('300份就业协议分类')
for i, v in enumerate(counts):
    ax.text(i, v+2, str(v), ha='center', fontsize=9)
plt.tight_layout(); fig.savefig(os.path.join(d, '就业协议分类.png'), dpi=150); plt.close()

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '就业协议'
ws.append(['序号', '协议编号', '毕业生编号', '用人单位', '岗位', '薪资', '签订日期', '协议状态', '状态标记'])
# 300份: 60虚假 + 40不规范 + 30单位不存在 + 20合规保留 + 150正常
for i in range(1, 301):
    by = f'BY{random.randint(1,500):04d}'
    if i <= 60:
        state = '就业协议系伪造，单位公章不实'
    elif i <= 100:
        state = '协议条款不规范，缺必备条款'
    elif i <= 130:
        state = '用人单位经核查不存在'
    elif i <= 150:
        state = '短期实习协议真实有效（反直觉正常项）'
    else:
        state = '正常'
    ws.append([i, f'XY{i:04d}', by, f'某公司{i}', random.choice(['工程师', '销售', '行政', '教师', '医生']),
               random.randint(3000, 15000), f'2025-{random.randint(1,8):02d}-15',
               '有效' if i > 100 else '无效', state])
style(ws, 9); wb.save(os.path.join(d, '就业协议.xlsx'))

mkzip(d, '02就业协议.rar')

# ========== 压缩包3: 灵活就业认定 ==========
print('=== 03 灵活就业认定 ===')
d = os.path.join(INPUT_DIR, '_t3'); os.makedirs(d)

fig, ax = plt.subplots(figsize=(8, 6))
labels = ['认定不规范', '虚假灵活就业', '超时未认定', '合规保留', '正常']
sizes = [30, 25, 10, 8, 7]
colors = ['#FF6B6B', '#FF9F43', '#FECA57', '#1DD1A1', '#48DBFB']
ax.pie(sizes, labels=labels, autopct='%1.0f%%', colors=colors, startangle=90)
ax.set_title('80份灵活就业认定分布')
plt.tight_layout(); fig.savefig(os.path.join(d, '灵活就业分布.png'), dpi=150); plt.close()

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '灵活就业认定'
ws.append(['序号', '认定编号', '毕业生编号', '就业形式', '收入证明', '认定日期', '认定人', '状态标记'])
# 80份: 30不规范 + 25虚假 + 10超时 + 8合规保留 + 7正常
forms = ['自主创业', '自由职业', '网络平台就业', '家庭作坊', '个体经营']
for i in range(1, 81):
    by = f'BY{random.randint(1,500):04d}'
    form = random.choice(forms)
    if i <= 30:
        state = '灵活就业认定材料不规范'
    elif i <= 55:
        state = '虚假灵活就业，实际未就业'
    elif i <= 65:
        state = '超过规定时限未完成认定'
    elif i <= 73:
        state = '自主创业有营业执照，合规（反直觉正常项）'
    else:
        state = '正常'
    ws.append([i, f'LH{i:03d}', by, form, '有' if i > 30 else '无',
               f'2025-{random.randint(1,8):02d}-15', f'认定人{i}', state])
style(ws, 8); wb.save(os.path.join(d, '灵活就业认定.xlsx'))

mkzip(d, '03灵活就业认定.rar')

# ========== 压缩包4: 用人单位回访 ==========
print('=== 04 用人单位回访 ===')
d = os.path.join(INPUT_DIR, '_t4'); os.makedirs(d)

fig, ax = plt.subplots(figsize=(10, 6))
results = ['查无此人', '单位否认', '联系失效', '确认就业', '正常']
counts = [30, 20, 15, 10, 25]
colors = ['#FF6B6B', '#FF9F43', '#FECA57', '#1DD1A1', '#48DBFB']
ax.bar(results, counts, color=colors, width=0.5)
ax.set_ylabel('记录数'); ax.set_title('100条用人单位回访结果')
for i, v in enumerate(counts):
    ax.text(i, v+1, str(v), ha='center', fontsize=9)
plt.tight_layout(); fig.savefig(os.path.join(d, '回访结果分布.png'), dpi=150); plt.close()

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '用人单位回访'
ws.append(['序号', '回访编号', '毕业生编号', '用人单位', '联系电话', '回访结果', '回访日期', '状态标记'])
# 100条: 30查无此人 + 20单位否认 + 15联系失效 + 10合规保留 + 25正常
for i in range(1, 101):
    by = f'BY{random.randint(1,500):04d}'
    if i <= 30:
        state = '回访查无此人，就业信息虚假'
        result = '查无此人'
    elif i <= 50:
        state = '用人单位否认录用该毕业生'
        result = '单位否认'
    elif i <= 65:
        state = '联系方式失效，无法核实'
        result = '联系失效'
    elif i <= 75:
        state = '单位确认毕业生真实就业（反直觉正常项）'
        result = '确认就业'
    else:
        state = '正常'
        result = '确认就业'
    ws.append([i, f'HF{i:03d}', by, f'某公司{i}', f'138{random.randint(10000000,99999999)}',
               result, f'2025-{random.randint(1,9):02d}-15', state])
style(ws, 8); wb.save(os.path.join(d, '用人单位回访.xlsx'))

mkzip(d, '04用人单位回访.rar')

# ========== 压缩包5: 就业质量问卷 ==========
print('=== 05 就业质量问卷 ===')
d = os.path.join(INPUT_DIR, '_t5'); os.makedirs(d)

fig, ax = plt.subplots(figsize=(10, 6))
satisfaction = ['非常满意', '满意', '一般', '不满意', '非常不满意']
counts = [30, 60, 50, 40, 20]
colors = ['#1DD1A1', '#48DBFB', '#FECA57', '#FF9F43', '#FF6B6B']
ax.bar(satisfaction, counts, color=colors, width=0.5)
ax.set_ylabel('人数'); ax.set_title('200份就业质量问卷满意度分布')
for i, v in enumerate(counts):
    ax.text(i, v+1, str(v), ha='center', fontsize=9)
plt.tight_layout(); fig.savefig(os.path.join(d, '满意度分布.png'), dpi=150); plt.close()

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '就业质量问卷'
ws.append(['序号', '问卷编号', '毕业生编号', '月薪', '专业匹配度', '工作满意度', '五险一金', '状态标记'])
# 200份: 40薪酬低 + 30专业不匹配 + 20满意度低 + 10合规保留 + 100正常
for i in range(1, 201):
    by = f'BY{random.randint(1,500):04d}'
    if i <= 40:
        state = '薪酬低于当地最低工资标准'
        salary = random.randint(1500, 2400)
    elif i <= 70:
        state = '专业与岗位严重不匹配'
        salary = random.randint(3000, 8000)
    elif i <= 90:
        state = '工作满意度极低，频繁加班'
        salary = random.randint(3000, 8000)
    elif i <= 100:
        state = '自愿选择低薪兴趣岗位（反直觉正常项）'
        salary = random.randint(2000, 3500)
    else:
        state = '正常'
        salary = random.randint(4000, 15000)
    ws.append([i, f'WJ{i:04d}', by, salary,
               random.choice(['完全匹配', '基本匹配', '不匹配']),
               random.choice(['非常满意', '满意', '一般', '不满意', '非常不满意']),
               '有' if i > 40 else '无', state])
style(ws, 8); wb.save(os.path.join(d, '就业质量问卷.xlsx'))

mkzip(d, '05就业质量问卷.rar')

# ========== 压缩包6: 就业制度与管理层要求 ==========
print('=== 06 就业制度与管理层要求 ===')
d = os.path.join(INPUT_DIR, '_t6'); os.makedirs(d)
policies = [
    {'name': '01毕业生就业统计管理办法.docx', 'title': '毕业生就业统计管理办法',
     'secs': ['总则', '统计范围', '统计标准', '数据审核', '上报流程', '附则'],
     'issues': '问题：1.统计标准不明确；2.数据审核不严；3.上报不及时'},
    {'name': '02就业协议管理办法.docx', 'title': '就业协议管理办法',
     'secs': ['总则', '协议签订', '协议审核', '协议归档', '违约责任', '附则'],
     'issues': '问题：1.协议审核不严；2.虚假协议未识别；3.归档不规范'},
    {'name': '03灵活就业认定办法.docx', 'title': '灵活就业认定办法',
     'secs': ['总则', '认定标准', '认定材料', '认定程序', '动态管理', '附则'],
     'issues': '问题：1.认定标准模糊；2.材料审核不严；3.动态管理缺失'},
    {'name': '04用人单位回访制度.docx', 'title': '用人单位回访制度',
     'secs': ['总则', '回访方式', '回访比例', '回访记录', '结果处理', '附则'],
     'issues': '问题：1.回访比例不足；2.回访记录不全；3.结果处理不规范'},
    {'name': '05就业指导服务办法.docx', 'title': '就业指导服务办法',
     'secs': ['总则', '指导内容', '服务方式', '师资配备', '效果评估', '附则'],
     'issues': '问题：1.指导内容单一；2.服务方式落后；3.师资不足'},
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
    (1, '就业处处长', '就业率造假虚报', '提高学校排名'),
    (2, '就业指导中心主任', '虚假就业证明不核实', '核实影响就业率'),
    (3, '二级学院院长', '灵活就业认定放宽', '提高灵活就业率'),
    (4, '就业处处长', '用人单位回访不做', '回访发现问题影响数据'),
    (5, '二级学院院长', '未就业登记为就业', '完成就业率指标'),
    (6, '就业指导中心主任', '就业协议不审核', '审核耽误时间'),
    (7, '就业处处长', '就业质量差不上报', '影响学校声誉'),
    (8, '二级学院院长', '虚假灵活就业不清理', '清理影响就业率'),
    (9, '就业指导中心主任', '回访查无此人不处理', '处理影响数据'),
    (10, '就业处处长', '就业指导服务不开展', '占用教学时间'),
]
for num, src, content, reason in reqs:
    doc.add_heading(f'要求{num}', level=1)
    p = doc.add_paragraph(); p.add_run('来源：').bold = True; p.add_run(src)
    p = doc.add_paragraph(); p.add_run('要求：').bold = True; p.add_run(content)
    p = doc.add_paragraph(); p.add_run('理由：').bold = True; p.add_run(reason)
    doc.add_paragraph()
doc.save(os.path.join(d, '管理层要求.docx'))

fig, ax = plt.subplots(figsize=(10, 4)); ax.axis('off')
ax.text(0.05, 0.9, '发件人：就业处', fontsize=11, transform=ax.transAxes)
ax.text(0.05, 0.75, '收件人：各学院、就业指导中心', fontsize=11, transform=ax.transAxes)
ax.text(0.05, 0.6, '主题：关于毕业生就业统计工作的指导意见', fontsize=11, transform=ax.transAxes)
ax.text(0.05, 0.3, '各位：\n毕业生就业率关系学校声誉和排名，\n请务必确保就业率达标，\n灵活就业认定可适当放宽，\n用人单位回访可简化。\n\n就业处处长\n2025年9月28日', fontsize=10, transform=ax.transAxes)
rect = plt.Rectangle((0,0),1,1, transform=ax.transAxes, fill=False, edgecolor='gray')
ax.add_patch(rect); plt.tight_layout()
fig.savefig(os.path.join(d, '就业处处长邮件.png'), dpi=150); plt.close()
mkzip(d, '06就业制度与管理层要求.rar')

print('\n=== 输入附件生成完成 ===')
