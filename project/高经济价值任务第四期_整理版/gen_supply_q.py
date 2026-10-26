#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""供应链成本优化咨询项目 - 输入附件生成（修正版）
强制指定精确分类数量，确保数据与题目描述完全一致"""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from docx import Document
import zipfile
import os
import random
import shutil
import numpy as np

random.seed(42)

plt.rcParams['font.sans-serif'] = ['WenQuanYi Micro Hei', 'SimHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

BASE = '/workspace/供应链成本优化_附件'
INPUT_DIR = os.path.join(BASE, '输入附件')
os.makedirs(INPUT_DIR, exist_ok=True)

header_font = Font(name='微软雅黑', size=11, bold=True, color='FFFFFF')
header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
cell_font = Font(name='微软雅黑', size=10)
thin = Side(border_style='thin', color='000000')
border = Border(left=thin, right=thin, top=thin, bottom=thin)
center = Alignment(horizontal='center', vertical='center', wrap_text=True)

def style_ws(ws, ncols):
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.font = cell_font
            cell.border = border
            cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

def make_zip(folder_path, zip_name):
    zip_path = os.path.join(INPUT_DIR, zip_name)
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(folder_path):
            for f in files:
                fp = os.path.join(root, f)
                arcname = os.path.relpath(fp, folder_path)
                zf.write(fp, arcname)
    shutil.rmtree(folder_path)
    size = os.path.getsize(zip_path)
    print(f'  生成: {zip_name} ({size/1024:.1f}KB)')

# ==================== 压缩包1: 近三年采购成本与供应商数据 ====================
print('=== 生成压缩包1: 近三年采购成本与供应商数据 ===')
d1 = os.path.join(INPUT_DIR, '_temp_01')
os.makedirs(d1, exist_ok=True)

# 1.1 三年采购成本趋势图PNG
for year in [2023, 2024, 2025]:
    fig, ax = plt.subplots(figsize=(10, 6))
    quarters = ['Q1', 'Q2', 'Q3', 'Q4']
    if year == 2023:
        vals = [1.8, 2.0, 2.1, 2.3]
    elif year == 2024:
        vals = [2.1, 2.2, 2.3, 2.5]
    else:
        vals = [2.3, 2.4, 2.5, 2.6]
    categories = ['原材料', '包材', '辅料', '物流', '其他']
    data = []
    for q in range(4):
        data.append([vals[q]*0.55, vals[q]*0.15, vals[q]*0.1, vals[q]*0.12, vals[q]*0.08])
    data = np.array(data)
    x = np.arange(4)
    w = 0.15
    colors = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5']
    for i, cat in enumerate(categories):
        ax.bar(x + i*w, data[:, i], w, label=cat, color=colors[i])
    ax.set_xlabel('季度')
    ax.set_ylabel('采购成本（亿元）')
    ax.set_title(f'{year}年采购成本明细')
    ax.set_xticks(x + w*2)
    ax.set_xticklabels(quarters)
    ax.legend(loc='upper left', fontsize=8)
    ax.set_ylim(0, 3.5)
    for i in range(4):
        ax.text(x[i] + w*2, data[i].sum() + 0.05, f'{data[i].sum():.2f}亿', ha='center', fontsize=8)
    plt.tight_layout()
    fig.savefig(os.path.join(d1, f'{year}年采购成本明细.png'), dpi=150)
    plt.close()

# 1.2 供应商名录Excel（60家）- 确定性分类
wb = openpyxl.Workbook()
ws = wb.active
ws.title = '供应商名录'
ws.append(['序号', '供应商编号', '供应商名称', '物料类别', '年采购额(万元)', '合作年限', '独家/多家', '供应商等级', '备注'])
categories = ['原材料-钢材', '原材料-塑料', '原材料-电子料', '包材', '辅料', '物流', '其他']

suppliers = []
for i in range(1, 61):
    cat = categories[(i-1) % 7]
    amt = round(random.uniform(50, 3000), 2)
    years = random.randint(1, 15)
    # 默认全部为多家、B级、无备注
    suppliers.append([i, f'GYS{i:03d}', f'华{cat}供应商{i}号', cat, amt, years, '多家', 'B', ''])

# 强制指定8家价格虚高 GYS026-GYS033
for idx in range(26, 34):
    s = suppliers[idx-1]
    s[6] = '多家'
    s[7] = 'B'
    s[8] = '采购单价高于市场均价15%以上'

# 强制指定12家绩效D级 GYS020-GYS025, GYS036-GYS041
for idx in list(range(20, 26)) + list(range(36, 42)):
    s = suppliers[idx-1]
    s[6] = '多家'
    s[7] = 'D'
    s[8] = '连续两年绩效评分低于60分'

# 强制指定5家独家供应商 GYS042-GYS046
for idx in range(42, 47):
    s = suppliers[idx-1]
    s[6] = '独家'
    s[7] = 'B'
    s[8] = '独家供应，需评估替代风险'

# 强制指定10家价格偏高但质量好 GYS047-GYS056
for idx in range(47, 57):
    s = suppliers[idx-1]
    s[6] = '多家'
    s[7] = 'A'
    s[8] = '单价偏高5-8%但良品率99.5%以上，综合成本更低'

# 其余25家 GYS001-GYS019, GYS057-GYS060 保持默认（多家、B级、正常）
for idx in list(range(1, 20)) + list(range(57, 61)):
    s = suppliers[idx-1]
    s[6] = '多家'
    s[7] = random.choice(['A', 'B', 'C'])
    s[8] = ''

for s in suppliers:
    ws.append(s)
style_ws(ws, 9)
wb.save(os.path.join(d1, '供应商名录.xlsx'))

# 1.3 供应商绩效评分表 - 确定性分类
wb = openpyxl.Workbook()
ws = wb.active
ws.title = '供应商绩效评分'
ws.append(['供应商编号', '质量评分(30分)', '交期评分(25分)', '价格评分(20分)', '服务评分(15分)', '创新评分(10分)', '总分(100分)', '评级'])

for i in range(1, 61):
    sid = f'GYS{i:03d}'
    if 20 <= i <= 25 or 36 <= i <= 41:
        # D级：总分<60
        q = random.randint(15, 20); d = random.randint(10, 15); p = random.randint(10, 14)
        sv = random.randint(5, 8); inn = random.randint(5, 7)
    elif 26 <= i <= 33:
        # 价格虚高：价格评分低但其他正常
        q = random.randint(25, 30); d = random.randint(20, 25); p = random.randint(10, 13)
        sv = random.randint(10, 15); inn = random.randint(7, 10)
    elif 42 <= i <= 46:
        # 独家：质量交期正常但创新低
        q = random.randint(22, 28); d = random.randint(18, 23); p = random.randint(14, 18)
        sv = random.randint(10, 14); inn = random.randint(5, 8)
    elif 47 <= i <= 56:
        # 价格偏高但质量好：质量评分高
        q = random.randint(28, 30); d = random.randint(22, 25); p = random.randint(13, 16)
        sv = random.randint(12, 15); inn = random.randint(8, 10)
    else:
        # 正常
        q = random.randint(22, 28); d = random.randint(18, 23); p = random.randint(15, 19)
        sv = random.randint(10, 14); inn = random.randint(7, 10)
    total = q + d + p + sv + inn
    if total >= 85: grade = 'A'
    elif total >= 70: grade = 'B'
    elif total >= 60: grade = 'C'
    else: grade = 'D'
    ws.append((sid, q, d, p, sv, inn, total, grade))
style_ws(ws, 8)
wb.save(os.path.join(d1, '供应商绩效评分表.xlsx'))

make_zip(d1, '01近三年采购成本与供应商数据.rar')

# ==================== 压缩包2: 库存与仓储数据 ====================
print('=== 生成压缩包2: 库存与仓储数据 ===')
d2 = os.path.join(INPUT_DIR, '_temp_02')
os.makedirs(d2, exist_ok=True)

# 2.1 库存明细表PNG
fig, ax = plt.subplots(figsize=(12, 8))
sku_types = ['原材料', '半成品', '成品', '包材', '辅料', '备件']
sku_counts = [80, 40, 30, 25, 15, 10]
sku_values = [4500, 2200, 1800, 600, 400, 200]
colors = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47']
x = range(len(sku_types))
ax.bar(x, sku_values, color=colors, width=0.5)
ax.set_xlabel('物料类型')
ax.set_ylabel('库存金额（万元）')
ax.set_title('2025年库存明细（按物料类型）')
ax.set_xticks(x)
ax.set_xticklabels(sku_types)
for i, v in enumerate(sku_values):
    ax.text(i, v + 50, f'{v}万\n({sku_counts[i]}个SKU)', ha='center', fontsize=9)
ax.set_ylim(0, 5500)
plt.tight_layout()
fig.savefig(os.path.join(d2, '2025年库存明细表.png'), dpi=150)
plt.close()

# 2.2 库龄分析表PNG
fig, ax = plt.subplots(figsize=(10, 6))
aging_labels = ['0-30天', '31-60天', '61-90天', '91-180天', '181-365天', '>365天']
aging_counts = [85, 45, 25, 20, 15, 10]
aging_values = [3200, 1800, 1200, 1500, 800, 600]
colors_age = ['#70AD47', '#5B9BD5', '#FFC000', '#ED7D31', '#FF6B6B', '#C00000']
ax.bar(aging_labels, aging_values, color=colors_age, width=0.6)
ax.set_xlabel('库龄')
ax.set_ylabel('库存金额（万元）')
ax.set_title('2025年库龄分析表')
ax.set_ylim(0, 3500)
for i, (c, v) in enumerate(zip(aging_counts, aging_values)):
    ax.text(i, v + 30, f'{v}万\n({c}个SKU)', ha='center', fontsize=9)
ax.annotate('滞销库存\n需处理', xy=(4, 800), xytext=(4.5, 2000),
            arrowprops=dict(arrowstyle='->', color='red'), fontsize=10, color='red')
ax.annotate('过期库存\n需报废', xy=(5, 600), xytext=(5.2, 1800),
            arrowprops=dict(arrowstyle='->', color='red'), fontsize=10, color='red')
plt.tight_layout()
fig.savefig(os.path.join(d2, '2025年库龄分析表.png'), dpi=150)
plt.close()

# 2.3 仓库布局图PNG
fig, ax = plt.subplots(figsize=(10, 8))
ax.set_xlim(0, 100)
ax.set_ylim(0, 80)
ax.set_title('仓库布局图', fontsize=14)
ax.plot([0, 100, 100, 0, 0], [0, 0, 80, 80, 0], 'k-', linewidth=2)
zones = [('原材料区A', 5, 50, 25, 25, '#4472C4'),
         ('原材料区B', 5, 20, 25, 25, '#5B9BD5'),
         ('半成品区', 35, 50, 25, 25, '#ED7D31'),
         ('成品区', 35, 20, 25, 25, '#A5A5A5'),
         ('包材区', 65, 50, 20, 25, '#FFC000'),
         ('辅料区', 65, 20, 20, 25, '#70AD47'),
         ('备件区', 90, 20, 8, 55, '#9B59B6'),
         ('收货区', 5, 60, 25, 15, '#E74C3C'),
         ('发货区', 35, 60, 25, 15, '#E74C3C'),
         ('通道', 60, 5, 35, 70, '#F0F0F0')]
for name, x, y, w, h, color in zones:
    rect = plt.Rectangle((x, y), w, h, facecolor=color, alpha=0.3, edgecolor='black')
    ax.add_patch(rect)
    ax.text(x+w/2, y+h/2, name, ha='center', va='center', fontsize=8)
ax.set_xlabel('长度（米）')
ax.set_ylabel('宽度（米）')
ax.set_aspect('equal')
plt.tight_layout()
fig.savefig(os.path.join(d2, '仓库布局图.png'), dpi=150)
plt.close()

# 2.4 安全库存参数表 - 确定性分类（精确控制数量）
wb = openpyxl.Workbook()
ws = wb.active
ws.title = '安全库存参数'
ws.append(['SKU编号', '物料名称', '日均消耗量', '采购提前期(天)', '安全库存系数', '设定安全库存', '实际库存', '库存状态'])

# 200个SKU: 20虚高 + 10缺料 + 15偏高合理 + 155正常
# 前20个(1-20): 库存虚高
for i in range(1, 21):
    daily = round(random.uniform(100, 500), 1)
    lead = random.randint(5, 20)
    coef = round(random.uniform(1.2, 1.5), 1)
    safety = int(daily * lead * coef)
    actual = safety * random.randint(3, 5)  # 实际>安全2倍以上
    ws.append((f'SKU{i:04d}', f'物料{i}', daily, lead, coef, safety, actual, '库存虚高'))

# 21-30: 缺料风险
for i in range(21, 31):
    daily = round(random.uniform(50, 300), 1)
    lead = random.randint(7, 25)
    coef = round(random.uniform(1.3, 1.8), 1)
    safety = int(daily * lead * coef)
    actual = int(safety * random.uniform(0.3, 0.7))  # 实际<安全80%
    ws.append((f'SKU{i:04d}', f'物料{i}', daily, lead, coef, safety, actual, '缺料风险'))

# 31-45: 安全库存偏高但合理（季节性备货）
for i in range(31, 46):
    daily = round(random.uniform(10, 80), 1)
    lead = random.randint(10, 30)
    coef = round(random.uniform(1.5, 2.0), 1)
    safety = int(daily * lead * coef)
    actual = int(safety * random.uniform(1.3, 1.8))  # 偏高但合理
    ws.append((f'SKU{i:04d}', f'物料{i}', daily, lead, coef, safety, actual, '安全库存合理但偏高（季节性备货）'))

# 46-200: 正常
for i in range(46, 201):
    daily = round(random.uniform(20, 400), 1)
    lead = random.randint(3, 20)
    coef = round(random.uniform(1.2, 1.6), 1)
    safety = int(daily * lead * coef)
    actual = int(safety * random.uniform(0.9, 1.3))  # 正常范围
    ws.append((f'SKU{i:04d}', f'物料{i}', daily, lead, coef, safety, actual, '正常'))

style_ws(ws, 8)
wb.save(os.path.join(d2, '安全库存参数表.xlsx'))

make_zip(d2, '02现有库存与仓储数据.rar')

# ==================== 压缩包3: 物流运输数据 ====================
print('=== 生成压缩包3: 物流运输数据 ===')
d3 = os.path.join(INPUT_DIR, '_temp_03')
os.makedirs(d3, exist_ok=True)

# 3.1 两年运输费用趋势PNG
for year in [2024, 2025]:
    fig, ax = plt.subplots(figsize=(10, 6))
    months = [f'{m}月' for m in range(1, 13)]
    if year == 2024:
        vals = [420, 380, 410, 450, 480, 520, 550, 530, 490, 460, 440, 470]
    else:
        vals = [450, 410, 440, 480, 510, 550, 580, 560, 520, 490, 470, 500]
    ax.plot(months, vals, 'o-', color='#4472C4', linewidth=2, markersize=6)
    ax.fill_between(range(12), vals, alpha=0.15, color='#4472C4')
    ax.set_xlabel('月份')
    ax.set_ylabel('运输费用（万元）')
    ax.set_title(f'{year}年运输费用明细')
    ax.set_ylim(0, 650)
    for i, v in enumerate(vals):
        ax.text(i, v + 15, str(v), ha='center', fontsize=8)
    ax.axhline(y=sum(vals)/12, color='r', linestyle='--', alpha=0.5, label=f'月均{sum(vals)/12:.0f}万')
    ax.legend()
    plt.xticks(rotation=45)
    plt.tight_layout()
    fig.savefig(os.path.join(d3, f'{year}年运输费用明细.png'), dpi=150)
    plt.close()

# 3.2 运输路线时效表 - 确定性分类（精确控制数量）
wb = openpyxl.Workbook()
ws = wb.active
ws.title = '运输路线时效'
ws.append(['路线编号', '起点', '终点', '运输方式', '标准时效(天)', '实际时效(天)', '年运费(万元)', '承运商', '路线状态'])
cities = ['上海', '广州', '深圳', '北京', '成都', '武汉', '西安', '杭州', '南京', '重庆']

# 50条路线: 10运费虚高 + 5可合并 + 5紧急运输合理 + 30正常
# 1-10: 运费虚高需重新招标
for i in range(1, 11):
    origin = cities[(i-1) % 10]
    dest = cities[i % 10]
    ws.append((f'RT{i:03d}', origin, dest, '公路', random.randint(2, 5), random.randint(2, 5),
               round(random.uniform(150, 250), 2), f'物流公司{(i%6)+1}', '运费虚高需重新招标'))

# 11-15: 路线可合并优化
for i in range(11, 16):
    origin = cities[(i-1) % 10]
    dest = cities[(i+2) % 10]
    ws.append((f'RT{i:03d}', origin, dest, '公路', random.randint(2, 5), random.randint(2, 5),
               round(random.uniform(80, 150), 2), f'物流公司{(i%6)+1}', '路线可合并优化'))

# 16-20: 紧急运输高价但合理
for i in range(16, 21):
    origin = cities[(i-1) % 10]
    dest = cities[(i+3) % 10]
    ws.append((f'RT{i:03d}', origin, dest, '航空', 1, 1,
               round(random.uniform(100, 200), 2), f'物流公司5', '紧急运输（合理高价）'))

# 21-50: 正常
for i in range(21, 51):
    origin = cities[(i-1) % 10]
    dest = cities[(i+1) % 10]
    ws.append((f'RT{i:03d}', origin, dest, random.choice(['公路', '铁路', '多式联运']),
               random.randint(1, 7), random.randint(1, 7),
               round(random.uniform(20, 120), 2), f'物流公司{(i%6)+1}', '正常'))

style_ws(ws, 9)
wb.save(os.path.join(d3, '运输路线时效表.xlsx'))

# 3.3 第三方物流合同摘要Word
from docx import Document
doc = Document()
doc.add_heading('第三方物流合同摘要', level=1)
contracts = [
    {'name': '物流公司1-干线运输合同', 'term': '2023.1.1-2025.12.31', 'amount': '年费450万', 'notes': '上海至广州干线，每日2班次，含保险'},
    {'name': '物流公司2-城配合同', 'term': '2024.1.1-2026.12.31', 'amount': '年费280万', 'notes': '上海市内配送，含装卸上架'},
    {'name': '物流公司3-冷链合同', 'term': '2023.6.1-2025.5.31', 'amount': '年费180万', 'notes': '已到期未续签但仍在服务，合同存在法律风险'},
    {'name': '物流公司4-铁路运输合同', 'term': '2024.3.1-2026.2.28', 'amount': '年费220万', 'notes': '上海至成都铁路运输，价格高于市场15%'},
    {'name': '物流公司5-航空急件合同', 'term': '2025.1.1-2025.12.31', 'amount': '年费120万', 'notes': '航空急件，单价偏高但时效保障合理'},
    {'name': '物流公司6-仓储外包合同', 'term': '2022.1.1-2025.12.31', 'amount': '年费350万', 'notes': '外地仓储外包，利用率仅40%'},
]
for c in contracts:
    doc.add_heading(c['name'], level=2)
    p = doc.add_paragraph()
    p.add_run('合同期限：').bold = True
    p.add_run(c['term'])
    p = doc.add_paragraph()
    p.add_run('合同金额：').bold = True
    p.add_run(c['amount'])
    p = doc.add_paragraph()
    p.add_run('备注：').bold = True
    p.add_run(c['notes'])
    doc.add_paragraph()
doc.save(os.path.join(d3, '第三方物流合同摘要.docx'))

make_zip(d3, '03物流运输数据.rar')

# ==================== 压缩包4: 供应链管理制度 ====================
print('=== 生成压缩包4: 供应链管理制度 ===')
d4 = os.path.join(INPUT_DIR, '_temp_04')
os.makedirs(d4, exist_ok=True)

policies = [
    {'name': '01采购管理制度.docx', 'title': '采购管理制度',
     'sections': ['第一章 总则', '第二章 采购计划管理', '第三章 供应商管理', '第四章 采购审批权限', '第五章 采购合同管理', '第六章 采购验收与付款', '第七章 附则'],
     'issues': '问题：1.采购审批权限未按金额分级；2.紧急采购流程缺失；3.供应商考核周期为年度，频次不足'},
    {'name': '02供应商准入管理办法.docx', 'title': '供应商准入管理办法',
     'sections': ['第一章 总则', '第二章 供应商准入条件', '第三章 供应商准入流程', '第四章 供应商分级管理', '第五章 供应商考核', '第六章 供应商退出', '第七章 附则'],
     'issues': '问题：1.准入条件未明确资质要求；2.分级标准仅按采购额，未含绩效；3.退出机制不完善'},
    {'name': '03库存管理办法.docx', 'title': '库存管理办法',
     'sections': ['第一章 总则', '第二章 库存分类', '第三章 安全库存设定', '第四章 库存盘点', '第五章 滞销库存处理', '第六章 库存周转考核', '第七章 附则'],
     'issues': '问题：1.安全库存计算方法未按物料ABC分类；2.盘点周期统一为年度，不合理；3.滞销处理流程不明确'},
    {'name': '04物流服务商管理办法.docx', 'title': '物流服务商管理办法',
     'sections': ['第一章 总则', '第二章 物流服务商选择', '第三章 物流合同管理', '第四章 物流绩效考核', '第五章 物流费用审核', '第六章 附则'],
     'issues': '问题：1.物流商选择未引入竞争机制；2.费用审核无标准；3.绩效考核指标缺失'},
]
for p in policies:
    doc = Document()
    doc.add_heading(p['title'], level=0)
    for sec in p['sections']:
        doc.add_heading(sec, level=1)
        doc.add_paragraph(f'{sec}具体条款内容详见正式制度文件。本制度适用于公司供应链管理各环节，自发布之日起执行。')
    doc.add_heading('现有制度问题诊断', level=1)
    doc.add_paragraph(p['issues'])
    doc.save(os.path.join(d4, p['name']))

make_zip(d4, '04供应链管理制度.rar')

# ==================== 压缩包5: 生产计划与需求预测 ====================
print('=== 生成压缩包5: 生产计划与需求预测 ===')
d5 = os.path.join(INPUT_DIR, '_temp_05')
os.makedirs(d5, exist_ok=True)

# 5.1 生产计划表PNG
fig, ax = plt.subplots(figsize=(12, 6))
months = [f'{m}月' for m in range(1, 13)]
plan = [120, 115, 130, 140, 150, 160, 170, 165, 155, 145, 135, 150]
actual = [118, 110, 125, 138, 145, 155, 168, 160, 150, 140, 130, 145]
x = range(12)
ax.bar([i-0.2 for i in x], plan, 0.4, label='计划产量', color='#4472C4', alpha=0.7)
ax.bar([i+0.2 for i in x], actual, 0.4, label='实际产量', color='#ED7D31', alpha=0.7)
ax.set_xlabel('月份')
ax.set_ylabel('产量（万件）')
ax.set_title('2025年生产计划与实际产量对比')
ax.set_xticks(x)
ax.set_xticklabels(months, rotation=45)
ax.legend()
ax.set_ylim(0, 200)
plt.tight_layout()
fig.savefig(os.path.join(d5, '2025年生产计划表.png'), dpi=150)
plt.close()

# 5.2 需求预测表PNG
fig, ax = plt.subplots(figsize=(12, 6))
forecast = [125, 120, 135, 145, 155, 165, 175, 170, 160, 150, 140, 155]
actual_sales = [122, 115, 130, 142, 148, 158, 172, 165, 155, 142, 128, 148]
ax.plot(months, forecast, 's-', color='#4472C4', linewidth=2, markersize=6, label='需求预测')
ax.plot(months, actual_sales, 'o-', color='#ED7D31', linewidth=2, markersize=6, label='实际销量')
ax.fill_between(range(12), forecast, actual_sales, alpha=0.15, color='red')
ax.set_xlabel('月份')
ax.set_ylabel('销量（万件）')
ax.set_title('2025年需求预测与实际销量对比')
ax.set_xticks(x)
ax.set_xticklabels(months, rotation=45)
ax.legend()
ax.set_ylim(0, 200)
plt.tight_layout()
fig.savefig(os.path.join(d5, '2025年需求预测表.png'), dpi=150)
plt.close()

# 5.3 销售订单明细Excel
wb = openpyxl.Workbook()
ws = wb.active
ws.title = '销售订单明细'
ws.append(['订单编号', '客户编号', '产品编号', '订单量', '下单日期', '交期', '状态', '备注'])
for i in range(1, 81):
    qty = random.randint(500, 5000)
    status = random.choice(['已完成', '已完成', '已完成', '生产中', '已逾期', '已取消'])
    note = ''
    if status == '已逾期':
        note = '交期延误，客户投诉'
    elif status == '已取消':
        note = '客户取消订单，库存积压'
    ws.append((f'SO{i:04d}', f'C{random.randint(1,30):03d}', f'P{random.randint(1,20):03d}',
               qty, f'2025-{random.randint(1,9):02d}-{random.randint(1,28):02d}',
               f'2025-{random.randint(1,9):02d}-{random.randint(1,28):02d}', status, note))
style_ws(ws, 8)
wb.save(os.path.join(d5, '销售订单明细.xlsx'))

# 5.4 物料清单BOM Excel
wb = openpyxl.Workbook()
ws = wb.active
ws.title = 'BOM清单'
ws.append(['产品编号', '产品名称', '物料编号', '物料名称', '用量', '单位', '物料类型', '采购方式'])
products = [('P001', '产品A'), ('P002', '产品B'), ('P003', '产品C'), ('P004', '产品D'),
            ('P005', '产品E'), ('P006', '产品F'), ('P007', '产品G'), ('P008', '产品H')]
materials = [('M001', '钢材A', '原材料', '批量采购'), ('M002', '钢材B', '原材料', '批量采购'),
             ('M003', '塑料颗粒A', '原材料', '批量采购'), ('M004', '电子元件A', '原材料', '单件采购'),
             ('M005', '包装箱', '包材', '批量采购'), ('M006', '标签', '包材', '批量采购'),
             ('M007', '润滑油', '辅料', '定期采购'), ('M008', '螺丝', '辅料', '批量采购')]
for p in products:
    for m in materials:
        ws.append((p[0], p[1], m[0], m[1], round(random.uniform(0.1, 5.0), 2), random.choice(['个', 'kg', '米', '件']), m[2], m[3]))
style_ws(ws, 8)
wb.save(os.path.join(d5, '物料清单BOM.xlsx'))

make_zip(d5, '05生产计划与需求预测.rar')

# ==================== 压缩包6: 管理层要求 ====================
print('=== 生成压缩包6: 管理层要求 ===')
d6 = os.path.join(INPUT_DIR, '_temp_06')
os.makedirs(d6, exist_ok=True)

doc = Document()
doc.add_heading('管理层要求汇总', level=0)
reqs = [
    (1, 'CEO', '供应商淘汰方案暂缓执行，维持现有供应商体系不变', '供应商淘汰可能短期影响生产'),
    (2, '采购总监', '价格虚高的供应商不调换，继续合作', '调换供应商需要重新走准入流程'),
    (3, 'CFO', '库存优化方案不报总经理，直接由供应链部执行', '库存削减涉及账面资产减少'),
    (4, '运营VP', '独家供应商不引入竞争，保持稳定', '独家供应风险但切换成本高'),
    (5, 'CEO', '物流合同到期的不重新招标，直接续签', '重新招标耗时且可能影响交付'),
    (6, '采购总监', '紧急采购不走审批，直接执行', '紧急采购流程太慢'),
    (7, 'CFO', '滞销库存不处理，等市场回暖再卖', '处理滞销库存会产生损失'),
    (8, '运营VP', '安全库存不调整，按现有标准执行', '降低安全库存可能导致缺料'),
    (9, 'CEO', '需求预测偏差不分析，直接用销售部数据', '分析偏差需要时间和人力'),
    (10, '采购总监', '供应商绩效不与采购份额挂钩', '绩效差的供应商也维持原份额'),
]
for num, source, content, reason in reqs:
    doc.add_heading(f'要求{num}', level=1)
    p = doc.add_paragraph()
    p.add_run('来源：').bold = True
    p.add_run(source)
    p = doc.add_paragraph()
    p.add_run('要求内容：').bold = True
    p.add_run(content)
    p = doc.add_paragraph()
    p.add_run('理由：').bold = True
    p.add_run(reason)
    doc.add_paragraph()
doc.save(os.path.join(d6, '管理层要求.docx'))

# CEO邮件PNG
fig, ax = plt.subplots(figsize=(10, 4))
ax.axis('off')
ax.text(0.05, 0.9, '发件人：CEO办公室', fontsize=11, transform=ax.transAxes)
ax.text(0.05, 0.75, '收件人：供应链管理部全体', fontsize=11, transform=ax.transAxes)
ax.text(0.05, 0.6, '主题：关于供应链优化方案的指导意见', fontsize=11, transform=ax.transAxes)
ax.text(0.05, 0.35, '各位：\n经管理层讨论，供应链优化方案涉及面广，为稳妥起见，\n请暂缓执行供应商淘汰、库存削减、物流招标等措施，\n先维持现状，后续再议。\n\nCEO\n2025年9月15日', fontsize=10, transform=ax.transAxes)
rect = plt.Rectangle((0, 0), 1, 1, transform=ax.transAxes, fill=False, edgecolor='gray')
ax.add_patch(rect)
plt.tight_layout()
fig.savefig(os.path.join(d6, 'CEO邮件.png'), dpi=150)
plt.close()

make_zip(d6, '06管理层要求.rar')

print('\n=== 输入附件全部生成完成（修正版） ===')
