#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""高校实验室安全管理与科研伦理合规审查 - 输入附件生成"""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from docx import Document
import zipfile, os, random, shutil
import numpy as np

random.seed(99)
plt.rcParams['font.sans-serif'] = ['WenQuanYi Micro Hei', 'SimHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

BASE = '/workspace/实验室安全伦理审查_附件'
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

# ========== 压缩包1: 高校实验室基本信息 ==========
print('=== 01 高校实验室基本信息 ===')
d = os.path.join(INPUT_DIR, '_t1'); os.makedirs(d)

fig, ax = plt.subplots(figsize=(10, 6))
types = ['化学', '生物', '物理', '材料', '医学', '其他']
counts = [25, 20, 15, 10, 8, 2]
colors = ['#4472C4', '#70AD47', '#ED7D31', '#A5A5A5', '#FF6B6B', '#9B59B6']
ax.bar(types, counts, color=colors, width=0.5)
ax.set_ylabel('实验室数'); ax.set_title('80个实验室类型分布')
for i, v in enumerate(counts):
    ax.text(i, v+0.3, str(v), ha='center', fontsize=9)
plt.tight_layout(); fig.savefig(os.path.join(d, '实验室类型分布.png'), dpi=150); plt.close()

fig, ax = plt.subplots(figsize=(10, 6))
levels = ['一级', '二级', '三级', '四级']
counts = [40, 25, 10, 5]
colors = ['#70AD47', '#FFC000', '#ED7D31', '#FF6B6B']
ax.bar(levels, counts, color=colors, width=0.5)
ax.set_ylabel('实验室数'); ax.set_title('实验室生物安全等级分布')
for i, v in enumerate(counts):
    ax.text(i, v+0.3, str(v), ha='center', fontsize=9)
plt.tight_layout(); fig.savefig(os.path.join(d, '安全等级分布.png'), dpi=150); plt.close()

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '实验室台账'
ws.append(['序号', '实验室编号', '实验室名称', '所属高校', '实验室类型', '生物安全等级', '面积(㎡)', '负责人', '安全员', '上次安全检查', '状态标记'])
# 80个: 12制度缺失 + 10安全员无资质 + 8超期未检查 + 5反直觉正常(新实验室正常) + 45正常
univs = ['理工大学', '科技大学', '医科大学', '农业大学', '师范大学', '工业大学', '化工大学', '海洋大学',
         '交通大学', '建筑大学', '轻工大学', '纺织大学', '邮电大学', '地质大学', '林业大学']
types = ['化学', '生物', '物理', '材料', '医学', '其他']
levels = ['一级', '二级', '三级', '四级']
for i in range(1, 81):
    univ = univs[(i-1) % 15]
    ltype = random.choice(types)
    level = random.choices(levels, weights=[40, 25, 10, 5])[0]
    if i <= 12:
        state = '安全管理制度缺失'
    elif i <= 22:
        state = '安全员无资质证书'
    elif i <= 30:
        state = '超期未进行安全检查'
    elif i <= 35:
        state = '新建成实验室，各项制度齐全（反直觉正常项）'
    else:
        state = '正常'
    ws.append([i, f'LAB{i:03d}', f'{univ}{ltype}实验室{i}', univ, ltype, level,
               random.randint(30, 300), f'负责人{i}', f'安全员{i}',
               f'2024-{random.randint(1,12):02d}-15', state])
style(ws, 11); wb.save(os.path.join(d, '实验室台账.xlsx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '高校清单'
ws.append(['序号', '高校编号', '高校名称', '高校类型', '实验室数量', '科研人员', '安全管理部门', '备注'])
for i in range(1, 16):
    ws.append([i, f'GX{i:03d}', univs[i-1], random.choice(['综合', '理工', '医药', '农林']),
               random.randint(3, 10), random.randint(50, 500),
               '安全管理处', '正常'])
style(ws, 8); wb.save(os.path.join(d, '高校清单.xlsx'))

mkzip(d, '01高校实验室基本信息.rar')

# ========== 压缩包2: 危险化学品管理 ==========
print('=== 02 危险化学品管理 ===')
d = os.path.join(INPUT_DIR, '_t2'); os.makedirs(d)

fig, ax = plt.subplots(figsize=(10, 6))
cats = ['采购', '存储', '使用', '废弃']
amounts = [120, 100, 95, 30]
colors = ['#4472C4', '#70AD47', '#ED7D31', '#FF6B6B']
ax.bar(cats, amounts, color=colors, width=0.5)
ax.set_ylabel('记录数'); ax.set_title('危险化学品管理各环节记录数')
for i, v in enumerate(amounts):
    ax.text(i, v+2, str(v), ha='center', fontsize=9)
plt.tight_layout(); fig.savefig(os.path.join(d, '化学品管理分布.png'), dpi=150); plt.close()

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '危险化学品台账'
ws.append(['序号', '化学品编号', '化学品名称', 'CAS号', '危险类别', '所属实验室', '库存量(g/ml)', '存储位置', '双人双锁', '采购日期', '状态标记'])
# 200条: 25台账不符 + 20存储违规 + 15废弃违规 + 10反直觉正常(易制毒合法备案) + 130正常
chems = ['硫酸', '盐酸', '硝酸', '氢氧化钠', '乙醇', '甲醇', '丙酮', '乙醚', '苯', '甲苯',
         '二甲苯', '氯仿', '四氯化碳', '乙酸', '氨水', '过氧化氢', '高锰酸钾', '重铬酸钾',
         '甲醛', '二甲亚砜']
hazards = ['腐蚀品', '易燃液体', '氧化剂', '毒害品', '爆炸品', '压缩气体']
for i in range(1, 201):
    chem = chems[(i-1) % len(chems)]
    haz = random.choice(hazards)
    lab = f'LAB{random.randint(1,80):03d}'
    if i <= 25:
        state = '台账与实际库存不符'
    elif i <= 45:
        state = '存储位置不符合危险等级要求'
    elif i <= 60:
        state = '废弃化学品未按规定处置'
    elif i <= 70:
        state = '易制毒化学品已在公安备案（反直觉正常项）'
    else:
        state = '正常'
    ws.append([i, f'HX{i:04d}', chem, f'CAS-{random.randint(10000,99999)}-{random.randint(1,9)}',
               haz, lab, random.randint(10, 5000), f'存储柜{random.randint(1,20)}',
               '是' if i > 45 else '否', f'2024-{random.randint(1,12):02d}-15', state])
style(ws, 11); wb.save(os.path.join(d, '危险化学品台账.xlsx'))

mkzip(d, '02危险化学品管理.rar')

# ========== 压缩包3: 生物安全与辐射安全 ==========
print('=== 03 生物安全与辐射安全 ===')
d = os.path.join(INPUT_DIR, '_t3'); os.makedirs(d)

fig, ax = plt.subplots(figsize=(8, 6))
labels = ['生物安全违规', '辐射安全违规', '合规']
sizes = [18, 12, 70]
colors = ['#FF6B6B', '#FFC000', '#70AD47']
ax.pie(sizes, labels=labels, autopct='%1.0f%%', colors=colors, startangle=90)
ax.set_title('生物与辐射安全合规分布（50条）')
plt.tight_layout(); fig.savefig(os.path.join(d, '安全合规分布.png'), dpi=150); plt.close()

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '生物安全记录'
ws.append(['序号', '记录编号', '所属实验室', '生物安全等级', '实验内容', '操作人员', '防护措施', '废弃物处理', '审批状态', '状态标记'])
# 50条: 15等级不符 + 10防护缺失 + 8废弃物违规 + 5反直觉正常(低风险实验) + 12正常
for i in range(1, 51):
    lab = f'LAB{random.randint(1,80):03d}'
    level = random.choice(['一级', '二级', '三级', '四级'])
    if i <= 15:
        state = '实验内容与生物安全等级不符'
    elif i <= 25:
        state = '操作人员防护措施缺失'
    elif i <= 33:
        state = '生物废弃物未高压灭菌处理'
    elif i <= 38:
        state = '低风险分子生物学实验，合规（反直觉正常项）'
    else:
        state = '正常'
    ws.append([i, f'BS{i:03d}', lab, level, f'实验内容{i}', f'操作人员{i}',
               '完整' if i > 15 else '缺失', '合规' if i > 25 else '违规',
               '已审批' if i > 33 else '未审批', state])
style(ws, 10); wb.save(os.path.join(d, '生物安全记录.xlsx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '辐射安全记录'
ws.append(['序号', '记录编号', '所属实验室', '辐射源类型', '活度(Bq)', '操作人员', '个人剂量计', '防护设施', '审批状态', '状态标记'])
# 30条: 8无剂量计 + 6防护缺失 + 4超期未检测 + 3反直觉正常(低剂量豁免) + 9正常
sources = ['X射线', 'γ射线', '中子源', '放射性同位素']
for i in range(1, 31):
    lab = f'LAB{random.randint(1,80):03d}'
    src = random.choice(sources)
    if i <= 8:
        state = '操作人员未佩戴个人剂量计'
    elif i <= 14:
        state = '辐射防护设施缺失或失效'
    elif i <= 18:
        state = '辐射设备超期未检测'
    elif i <= 21:
        state = '低剂量豁免源，合规（反直觉正常项）'
    else:
        state = '正常'
    ws.append([i, f'RS{i:03d}', lab, src, random.randint(1000, 1000000),
               f'操作人员{i}', '已佩戴' if i > 8 else '未佩戴',
               '完备' if i > 14 else '缺失', '已审批' if i > 18 else '未审批', state])
style(ws, 10); wb.save(os.path.join(d, '辐射安全记录.xlsx'))

mkzip(d, '03生物安全与辐射安全.rar')

# ========== 压缩包4: 实验动物与人类受试者伦理 ==========
print('=== 04 实验动物与人类受试者伦理 ===')
d = os.path.join(INPUT_DIR, '_t4'); os.makedirs(d)

fig, ax = plt.subplots(figsize=(10, 6))
types = ['小鼠', '大鼠', '兔', '豚鼠', '其他']
counts = [18, 6, 3, 2, 1]
colors = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#9B59B6']
ax.bar(types, counts, color=colors, width=0.5)
ax.set_ylabel('审批份数'); ax.set_title('实验动物伦理审批分布（30份）')
for i, v in enumerate(counts):
    ax.text(i, v+0.2, str(v), ha='center', fontsize=9)
plt.tight_layout(); fig.savefig(os.path.join(d, '动物伦理分布.png'), dpi=150); plt.close()

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '实验动物伦理审批'
ws.append(['序号', '审批编号', '所属高校', '实验项目', '动物种类', '动物数量', '伦理委员会意见', '批准日期', '有效期', '状态标记'])
# 30份: 8未审批 + 6超范围 + 4到期未续 + 3反直觉正常(替代方法) + 9正常
animals = ['小鼠', '大鼠', '兔', '豚鼠', '斑马鱼']
for i in range(1, 31):
    univ = f'GX{random.randint(1,15):03d}'
    animal = random.choice(animals)
    if i <= 8:
        state = '未经伦理委员会审批'
    elif i <= 14:
        state = '实验内容超出批准范围'
    elif i <= 18:
        state = '伦理审批到期未续'
    elif i <= 21:
        state = '采用替代方法减少动物使用，合规（反直觉正常项）'
    else:
        state = '正常'
    ws.append([i, f'AL{i:03d}', univ, f'实验项目{i}', animal, random.randint(10, 500),
               '批准' if i > 8 else '未审批', f'2024-{random.randint(1,12):02d}-01',
               '1年' if i > 18 else '已过期', state])
style(ws, 10); wb.save(os.path.join(d, '实验动物伦理审批.xlsx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '人类受试者伦理审批'
ws.append(['序号', '审批编号', '所属高校', '研究项目', '受试者人数', '知情同意', '隐私保护', '伦理委员会意见', '批准日期', '状态标记'])
# 25份: 7无知情同意 + 5隐私泄露 + 4未审批 + 3反直觉正常(匿名数据) + 6正常
for i in range(1, 26):
    univ = f'GX{random.randint(1,15):03d}'
    if i <= 7:
        state = '受试者未签署知情同意书'
    elif i <= 12:
        state = '受试者隐私信息未脱敏处理'
    elif i <= 16:
        state = '未经伦理委员会审批'
    elif i <= 19:
        state = '使用公开匿名数据，合规（反直觉正常项）'
    else:
        state = '正常'
    ws.append([i, f'HS{i:03d}', univ, f'研究项目{i}', random.randint(20, 500),
               '已签署' if i > 7 else '未签署', '已脱敏' if i > 12 else '未脱敏',
               '批准' if i > 16 else '未审批', f'2024-{random.randint(1,12):02d}-01', state])
style(ws, 10); wb.save(os.path.join(d, '人类受试者伦理审批.xlsx'))

mkzip(d, '04实验动物与人类受试者伦理.rar')

# ========== 压缩包5: 安全事故与应急预案 ==========
print('=== 05 安全事故与应急预案 ===')
d = os.path.join(INPUT_DIR, '_t5'); os.makedirs(d)

fig, ax = plt.subplots(figsize=(10, 6))
months = [f'{m}月' for m in range(1, 13)]
accidents = [1, 0, 2, 1, 0, 3, 1, 2, 0, 1, 2, 1]
ax.bar(months, accidents, color='#FF6B6B', alpha=0.7)
ax.set_ylabel('事故数'); ax.set_title('2024年安全事故月度分布')
for i, v in enumerate(accidents):
    if v > 0:
        ax.text(i, v+0.1, str(v), ha='center', fontsize=9)
plt.xticks(rotation=45); plt.tight_layout()
fig.savefig(os.path.join(d, '安全事故分布.png'), dpi=150); plt.close()

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '安全事故记录'
ws.append(['序号', '事故编号', '所属实验室', '事故类型', '发生时间', '严重等级', '人员伤亡', '经济损失', '处置情况', '状态标记'])
# 15起: 6隐瞒不报 + 5处置不当 + 2超期上报 + 2反直觉正常(及时处置) + 0? 不对应该15
# 15起: 5隐瞒 + 4处置不当 + 3超期上报 + 3及时处置(反直觉正常) = 15
for i in range(1, 16):
    lab = f'LAB{random.randint(1,80):03d}'
    atype = random.choice(['化学品泄漏', '火灾', '烫伤', '割伤', '感染', '辐射'])
    severity = random.choice(['一般', '较大', '重大'])
    if i <= 5:
        state = '事故隐瞒未上报'
    elif i <= 9:
        state = '事故处置不当'
    elif i <= 12:
        state = '超期上报（超过24小时）'
    else:
        state = '及时上报并妥善处置（反直觉正常项）'
    ws.append([i, f'AC{i:03d}', lab, atype, f'2024-{random.randint(1,12):02d}-{random.randint(1,28):02d}',
               severity, random.randint(0, 3), random.randint(0, 50),
               '已处置' if i > 9 else '处置不当', state])
style(ws, 10); wb.save(os.path.join(d, '安全事故记录.xlsx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '应急预案'
ws.append(['序号', '预案编号', '所属高校', '预案类型', '编制日期', '演练记录', '评审状态', '更新状态', '状态标记'])
# 15份: 6未演练 + 5未评审 + 2超期未更新 + 2反直觉正常(完善) = 15
for i in range(1, 16):
    univ = f'GX{random.randint(1,15):03d}'
    ptype = random.choice(['火灾', '化学品泄漏', '生物安全', '辐射事故', '人员伤害'])
    if i <= 6:
        state = '应急预案从未演练'
    elif i <= 11:
        state = '预案未经专家评审'
    elif i <= 13:
        state = '预案超期未更新'
    else:
        state = '预案完善且定期演练（反直觉正常项）'
    ws.append([i, f'EP{i:03d}', univ, ptype, f'2023-{random.randint(1,12):02d}-01',
               '有' if i > 6 else '无', '已评审' if i > 11 else '未评审',
               '已更新' if i > 13 else '超期', state])
style(ws, 9); wb.save(os.path.join(d, '应急预案.xlsx'))

mkzip(d, '05安全事故与应急预案.rar')

# ========== 压缩包6: 安全管理制度与管理层要求 ==========
print('=== 06 安全管理制度与管理层要求 ===')
d = os.path.join(INPUT_DIR, '_t6'); os.makedirs(d)
policies = [
    {'name': '01实验室安全管理办法.docx', 'title': '实验室安全管理办法',
     'secs': ['总则', '安全责任', '安全检查', '隐患整改', '事故处理', '附则'],
     'issues': '问题：1.安全检查频次不足；2.隐患整改无跟踪；3.事故上报不及时'},
    {'name': '02危险化学品管理办法.docx', 'title': '危险化学品管理办法',
     'secs': ['总则', '采购管理', '存储管理', '使用管理', '废弃处置', '附则'],
     'issues': '问题：1.台账不规范；2.存储不分类；3.废弃处置无记录'},
    {'name': '03生物安全管理办法.docx', 'title': '生物安全管理办法',
     'secs': ['总则', '等级管理', '操作规范', '废弃物处理', '监督检查', '附则'],
     'issues': '问题：1.等级认定不严；2.操作无规范；3.废弃物处理不规范'},
    {'name': '04科研伦理审查办法.docx', 'title': '科研伦理审查办法',
     'secs': ['总则', '伦理委员会', '审查程序', '动物伦理', '人类受试者', '附则'],
     'issues': '问题：1.委员会组成不规范；2.审查程序简化；3.跟踪审查缺失'},
    {'name': '05辐射安全管理办法.docx', 'title': '辐射安全管理办法',
     'secs': ['总则', '辐射源管理', '人员防护', '监测检测', '应急处置', '附则'],
     'issues': '问题：1.个人剂量监测缺失；2.防护不到位；3.检测不及时'},
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
    (1, '科研副校长', '安全事故隐瞒不上报', '上报影响学校声誉'),
    (2, '实验室主任', '危险化学品台账不整改', '整改影响实验进度'),
    (3, '科研副院长', '生物安全实验不审批', '审批耽误研究时间'),
    (4, '动物中心主任', '动物伦理不补审批', '补审影响动物实验'),
    (5, '科研副院长', '人类受试者不签知情同意', '签同意书影响招募'),
    (6, '实验室主任', '辐射安全检测不做', '检测增加费用'),
    (7, '科研副校长', '应急预案不演练', '演练影响正常教学'),
    (8, '实验室主任', '化学品废弃不按规定', '规定处置成本高'),
    (9, '科研副院长', '安全检查不整改', '整改影响科研产出'),
    (10, '实验室主任', '安全员无资质不培训', '培训占用时间'),
]
for num, src, content, reason in reqs:
    doc.add_heading(f'要求{num}', level=1)
    p = doc.add_paragraph(); p.add_run('来源：').bold = True; p.add_run(src)
    p = doc.add_paragraph(); p.add_run('要求：').bold = True; p.add_run(content)
    p = doc.add_paragraph(); p.add_run('理由：').bold = True; p.add_run(reason)
    doc.add_paragraph()
doc.save(os.path.join(d, '管理层要求.docx'))

fig, ax = plt.subplots(figsize=(10, 4)); ax.axis('off')
ax.text(0.05, 0.9, '发件人：科研处', fontsize=11, transform=ax.transAxes)
ax.text(0.05, 0.75, '收件人：各实验室、安全管理处', fontsize=11, transform=ax.transAxes)
ax.text(0.05, 0.6, '主题：关于实验室安全审查的指导意见', fontsize=11, transform=ax.transAxes)
ax.text(0.05, 0.3, '各位：\n实验室安全审查涉及面广，为保障科研工作正常开展，\n请暂缓执行安全事故上报、危险化学品整改、\n生物安全审批等措施，先维持现状。\n\n科研副校长\n2025年9月28日', fontsize=10, transform=ax.transAxes)
rect = plt.Rectangle((0,0),1,1, transform=ax.transAxes, fill=False, edgecolor='gray')
ax.add_patch(rect); plt.tight_layout()
fig.savefig(os.path.join(d, '科研副校长邮件.png'), dpi=150); plt.close()
mkzip(d, '06安全管理制度与管理层要求.rar')

print('\n=== 输入附件生成完成 ===')
