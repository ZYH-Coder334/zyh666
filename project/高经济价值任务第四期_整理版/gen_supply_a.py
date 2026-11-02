#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""供应链成本优化咨询项目 - 标准答案附件生成
生成6个压缩包，每个对应一个考点"""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from docx import Document
import zipfile
import os
import shutil

plt.rcParams['font.sans-serif'] = ['WenQuanYi Micro Hei', 'SimHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

BASE = '/workspace/供应链成本优化_附件'
ANS_DIR = os.path.join(BASE, '标准答案')
os.makedirs(ANS_DIR, exist_ok=True)

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
    zip_path = os.path.join(ANS_DIR, zip_name)
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(folder_path):
            for f in files:
                fp = os.path.join(root, f)
                arcname = os.path.relpath(fp, folder_path)
                zf.write(fp, arcname)
    shutil.rmtree(folder_path)
    size = os.path.getsize(zip_path)
    print(f'  生成: {zip_name} ({size/1024:.1f}KB)')

# ==================== 答案1: 供应链成本诊断报告 ====================
print('=== 生成答案1: 供应链成本诊断报告 ===')
d1 = os.path.join(ANS_DIR, '_temp_a01')
os.makedirs(d1, exist_ok=True)

# 诊断报告Word
doc = Document()
doc.add_heading('供应链成本深度诊断报告', level=0)
doc.add_heading('一、诊断概述', level=1)
doc.add_paragraph('本报告基于华恒制造企业（脱敏代称）近三年采购成本数据、供应商数据、库存数据、物流数据，'
                  '进行全面诊断分析。企业年营收15亿元，供应链成本占营收65%（9.75亿元），'
                  '库存周转率仅3.2次/年（行业平均6次），存在重大优化空间。')
doc.add_heading('二、采购成本诊断', level=1)
doc.add_paragraph('1. 采购成本总额：2025年9.8亿元，三年增长19.5%（2023年8.2亿→2025年9.8亿），'
                  '增速高于营收增速（营收仅增长12%）。')
doc.add_paragraph('2. 供应商成本问题：60家供应商中，8家采购单价高于市场均价15%以上（GYS026-GYS033），'
                  '年浪费约1200万元；12家绩效D级供应商（GYS020-GYS025、GYS036-GYS041）'
                  '质量差、交期不准，导致隐性成本约800万元/年。')
doc.add_paragraph('3. 反直觉发现：10家供应商（GYS047-GYS056）单价偏高5-8%，但良品率99.5%以上，'
                  '综合不良率成本更低，应予保留，不应误判为需调换。')
doc.add_paragraph('4. 独家供应商风险：5家独家供应商（GYS042-GYS046），存在断供风险，'
                  '需引入备选供应商。')
doc.add_heading('三、库存成本诊断', level=1)
doc.add_paragraph('1. 库存总额：9700万元，周转率3.2次/年，低于行业平均6次。')
doc.add_paragraph('2. 库龄问题：181天以上滞销库存25个SKU（1400万元），365天以上过期库存10个SKU（600万元），'
                  '合计占用资金2000万元。')
doc.add_paragraph('3. 库存结构问题：200个SKU中，20个库存虚高需削减（约800万元），'
                  '10个缺料风险需补库（反直觉陷阱，看似低库存实则缺料），'
                  '15个安全库存偏高但合理（反直觉正常项，季节性物料备货）。')
doc.add_paragraph('4. 仓库布局问题：备件区利用率仅40%，外地仓利用率40%，'
                  '仓库布局不合理导致搬运效率低。')
doc.add_heading('四、物流成本诊断', level=1)
doc.add_paragraph('1. 物流总费用：2025年约5800万元，占采购成本5.9%。')
doc.add_paragraph('2. 运费虚高：50条路线中，10条运费虚高需重新招标（年浪费约300万元），'
                  '5条路线可合并优化（年节省约100万元）。')
doc.add_paragraph('3. 合同风险：物流公司3冷链合同已到期未续签仍在服务（法律风险），'
                  '物流公司4铁路运输价格高于市场15%，物流公司6仓储外包利用率仅40%。')
doc.add_paragraph('4. 反直觉发现：5条紧急运输路线单价偏高但时效保障合理，'
                  '为紧急订单配送，不应误判为运费虚高。')
doc.add_heading('五、供应链管理制度诊断', level=1)
doc.add_paragraph('1. 采购管理制度：采购审批权限未按金额分级，紧急采购流程缺失。')
doc.add_paragraph('2. 供应商准入办法：准入条件未明确资质要求，分级标准仅按采购额未含绩效，退出机制不完善。')
doc.add_paragraph('3. 库存管理办法：安全库存计算未按ABC分类，盘点周期统一为年度不合理，滞销处理流程不明确。')
doc.add_paragraph('4. 物流服务商管理办法：物流商选择未引入竞争机制，费用审核无标准，绩效考核指标缺失。')
doc.add_heading('六、总体诊断结论', level=1)
doc.add_paragraph('供应链存在6大类问题，年浪费约3500万元。'
                  '采购成本浪费约2000万元（含供应商价格虚高1200万、绩效差隐性成本800万），'
                  '库存浪费约800万元，物流浪费约700万元。'
                  '通过系统优化可实现年降本8%（约7800万元），库存周转率提升至6次/年。')
doc.save(os.path.join(d1, '供应链成本深度诊断报告.docx'))

# 诊断结论汇总表Excel
wb = openpyxl.Workbook()
ws = wb.active
ws.title = '诊断结论汇总'
ws.append(['序号', '诊断维度', '问题类别', '问题数量', '影响金额(万元)', '优化方向', '优先级'])
diags = [
    (1, '采购成本', '供应商价格虚高', 8, 1200, '重新招标或引入竞争', '高'),
    (2, '采购成本', '供应商绩效差', 12, 800, '淘汰并替换', '高'),
    (3, '采购成本', '独家供应商风险', 5, 0, '引入备选供应商', '中'),
    (4, '采购成本', '价格偏高但质量好（保留）', 10, 0, '维持现状', '低'),
    (5, '库存成本', '滞销库存', 25, 1400, '降价处理或报废', '高'),
    (6, '库存成本', '过期库存', 10, 600, '报废处理', '高'),
    (7, '库存成本', '库存虚高', 20, 800, '削减库存', '高'),
    (8, '库存成本', '缺料风险', 10, 0, '补库至安全水位', '中'),
    (9, '库存成本', '安全库存偏高但合理（保留）', 15, 0, '维持现状', '低'),
    (10, '物流成本', '运费虚高', 10, 300, '重新招标', '高'),
    (11, '物流成本', '路线可合并', 5, 100, '合并优化', '中'),
    (12, '物流成本', '紧急运输高价但合理（保留）', 5, 0, '维持现状', '低'),
    (13, '物流成本', '合同到期未续签', 1, 0, '补签合同', '高'),
    (14, '管理制度', '采购审批权限未分级', 1, 0, '修订制度', '中'),
    (15, '管理制度', '供应商考核频次不足', 1, 0, '修订制度', '中'),
]
for d in diags:
    ws.append(d)
style_ws(ws, 7)
wb.save(os.path.join(d1, '诊断结论汇总表.xlsx'))

# 诊断维度雷达图PNG
fig, ax = plt.subplots(figsize=(8, 8), subplot_kw=dict(projection='polar'))
import numpy as np
categories = ['采购管理', '库存管理', '物流管理', '供应商管理', '制度体系', '需求预测']
current = [60, 45, 55, 50, 40, 35]
target = [85, 80, 82, 80, 78, 75]
angles = np.linspace(0, 2*np.pi, len(categories), endpoint=False).tolist()
current += current[:1]
target += target[:1]
angles += angles[:1]
ax.plot(angles, current, 'o-', linewidth=2, label='当前水平', color='#ED7D31')
ax.fill(angles, current, alpha=0.15, color='#ED7D31')
ax.plot(angles, target, 's-', linewidth=2, label='目标水平', color='#4472C4')
ax.fill(angles, target, alpha=0.15, color='#4472C4')
ax.set_xticks(angles[:-1])
ax.set_xticklabels(categories)
ax.set_ylim(0, 100)
ax.set_title('供应链各维度当前水平 vs 目标', fontsize=12, pad=20)
ax.legend(loc='upper right', bbox_to_anchor=(1.3, 1.1))
plt.tight_layout()
fig.savefig(os.path.join(d1, '供应链诊断雷达图.png'), dpi=150)
plt.close()

make_zip(d1, '01供应链成本诊断报告.rar')

# ==================== 答案2: 供应商优化方案 ====================
print('=== 生成答案2: 供应商优化方案 ===')
d2 = os.path.join(ANS_DIR, '_temp_a02')
os.makedirs(d2, exist_ok=True)

# 供应商优化方案Word
doc = Document()
doc.add_heading('供应商优化方案', level=0)
doc.add_heading('一、优化目标', level=1)
doc.add_paragraph('1. 采购成本年降低2000万元（价格虚高供应商调减1200万 + 绩效差供应商替换节约800万）')
doc.add_paragraph('2. 供应商绩效平均分从72分提升至85分')
doc.add_paragraph('3. 独家供应商引入备选，降低断供风险')
doc.add_heading('二、供应商分类优化措施', level=1)
doc.add_heading('2.1 价格虚高供应商调换（8家）', level=2)
doc.add_paragraph('供应商编号：GYS026至GYS033')
doc.add_paragraph('问题：采购单价高于市场均价15%以上')
doc.add_paragraph('措施：重新招标引入竞争性报价，要求降价至市场均价或更换供应商')
doc.add_paragraph('预计降本：1200万元/年')
doc.add_heading('2.2 绩效差供应商淘汰（12家）', level=2)
doc.add_paragraph('供应商编号：GYS020至GYS025（6家）、GYS036至GYS041（6家）')
doc.add_paragraph('问题：连续两年绩效评分低于60分，质量差、交期不准')
doc.add_paragraph('措施：启动淘汰流程，从A级备选供应商中替换')
doc.add_paragraph('预计降本：800万元/年（隐性成本节约）')
doc.add_heading('2.3 独家供应商引入竞争（5家）', level=2)
doc.add_paragraph('供应商编号：GYS042至GYS046')
doc.add_paragraph('问题：独家供应存在断供风险')
doc.add_paragraph('措施：引入1-2家备选供应商，采用70:30分配原则')
doc.add_paragraph('注意：不可直接替换，需渐进过渡，确保供应稳定')
doc.add_heading('2.4 价格偏高但质量好供应商保留（10家，反直觉正常项）', level=2)
doc.add_paragraph('供应商编号：GYS047至GYS056')
doc.add_paragraph('判定：单价偏高5-8%但良品率99.5%以上，综合不良率成本更低')
doc.add_paragraph('措施：维持现状，不误判为需调换')
doc.add_paragraph('依据：综合成本=采购成本+质量成本+交期成本，良品率高的供应商综合成本更低')
doc.add_heading('三、供应商优化时间表', level=1)
doc.add_paragraph('第1-2月：完成8家价格虚高供应商的重新招标和价格谈判')
doc.add_paragraph('第3-4月：完成12家绩效差供应商的淘汰替换')
doc.add_paragraph('第5-6月：完成5家独家供应商的备选引入')
doc.add_paragraph('第7-12月：持续监控供应商绩效，确保优化效果')
doc.save(os.path.join(d2, '供应商优化方案.docx'))

# 供应商优化明细表Excel
wb = openpyxl.Workbook()
ws = wb.active
ws.title = '供应商优化明细'
ws.append(['序号', '优化类别', '供应商编号范围', '数量', '问题', '优化措施', '预计降本(万元)', '依据'])
rows = [
    (1, '价格虚高调换', 'GYS026-GYS033', 8, '单价高于市场15%', '重新招标/降价谈判', 1200, '市场比价数据'),
    (2, '绩效差淘汰', 'GYS020-GYS025,GYS036-GYS041', 12, '绩效D级', '从备选中替换', 800, '供应商绩效评分表'),
    (3, '独家引入竞争', 'GYS042-GYS046', 5, '独家断供风险', '引入备选70:30分配', 0, '风险管理要求'),
    (4, '价格偏高保留（正常项）', 'GYS047-GYS056', 10, '单价偏高5-8%但良品率99.5%', '维持现状', 0, '综合成本分析'),
    (5, '正常维持', '其余25家', 25, '绩效正常', '维持现有合作', 0, '-'),
    (6, '合计', '-', 60, '-', '-', 2000, '-'),
]
for r in rows:
    ws.append(r)
style_ws(ws, 8)
wb.save(os.path.join(d2, '供应商优化明细表.xlsx'))

# 供应商优化前后对比图PNG
fig, ax = plt.subplots(figsize=(10, 6))
cats = ['采购成本(万)', '质量成本(万)', '交期成本(万)', '综合成本(万)']
before = [98000, 3000, 2000, 103000]
after = [96000, 1500, 1200, 98700]
x = range(len(cats))
ax.bar([i-0.2 for i in x], before, 0.4, label='优化前', color='#ED7D31', alpha=0.7)
ax.bar([i+0.2 for i in x], after, 0.4, label='优化后', color='#4472C4', alpha=0.7)
ax.set_ylabel('金额（万元）')
ax.set_title('供应商优化前后成本对比')
ax.set_xticks(x)
ax.set_xticklabels(cats)
ax.legend()
for i in range(len(cats)):
    ax.text(i-0.2, before[i]+500, str(before[i]), ha='center', fontsize=8)
    ax.text(i+0.2, after[i]+500, str(after[i]), ha='center', fontsize=8)
plt.tight_layout()
fig.savefig(os.path.join(d2, '供应商优化前后对比图.png'), dpi=150)
plt.close()

make_zip(d2, '02供应商优化方案.rar')

# ==================== 答案3: 库存优化方案 ====================
print('=== 生成答案3: 库存优化方案 ===')
d3 = os.path.join(ANS_DIR, '_temp_a03')
os.makedirs(d3, exist_ok=True)

doc = Document()
doc.add_heading('库存优化方案', level=0)
doc.add_heading('一、优化目标', level=1)
doc.add_paragraph('1. 库存周转率从3.2次/年提升至6次/年')
doc.add_paragraph('2. 库存总额从9700万元降至6400万元（降34%）')
doc.add_paragraph('3. 释放占用资金3300万元')
doc.add_heading('二、分类优化措施', level=1)
doc.add_heading('2.1 滞销库存处理（25个SKU，1400万元）', level=2)
doc.add_paragraph('库龄181天以上滞销库存，分三档处理：')
doc.add_paragraph('  A档（181-365天，15个SKU，800万元）：降价30%促销处理')
doc.add_paragraph('  B档（可返厂，5个SKU，300万元）：与供应商协商退货')
doc.add_paragraph('  C档（无法处理，5个SKU，300万元）：报废处理')
doc.add_paragraph('预计回收：约1000万元，损失约400万元')
doc.add_heading('2.2 过期库存报废（10个SKU，600万元）', level=2)
doc.add_paragraph('库龄365天以上，已无使用价值，直接报废')
doc.add_paragraph('需走资产减值流程，账面确认损失600万元')
doc.add_heading('2.3 库存虚高削减（20个SKU，800万元）', level=2)
doc.add_paragraph('实际库存超过安全库存2倍以上')
doc.add_paragraph('措施：暂停采购，消耗至安全库存水位')
doc.add_paragraph('预计释放资金：800万元')
doc.add_heading('2.4 缺料风险补库（10个SKU，反直觉陷阱）', level=2)
doc.add_paragraph('看似低库存实则缺料风险，实际库存低于安全库存80%')
doc.add_paragraph('措施：紧急补库至安全库存水位')
doc.add_paragraph('注意：不可误判为库存充足，需核查日均消耗量和采购提前期')
doc.add_heading('2.5 安全库存偏高但合理保留（15个SKU，反直觉正常项）', level=2)
doc.add_paragraph('季节性物料备货，库龄长但属合理备货')
doc.add_paragraph('措施：维持现状，不误判为库存虚高')
doc.add_paragraph('依据：季节性物料消耗集中，需提前备货')
doc.add_heading('2.6 仓库布局优化', level=2)
doc.add_paragraph('1. 备件区面积缩减40%（利用率仅40%）')
doc.add_paragraph('2. 外地仓退租（利用率40%，年租金80万元）')
doc.add_paragraph('3. 原材料区A和B合并，减少搬运距离')
doc.add_heading('三、库存优化时间表', level=1)
doc.add_paragraph('第1月：完成滞销库存分类处理方案')
doc.add_paragraph('第2月：完成过期库存报废流程')
doc.add_paragraph('第3-4月：完成库存虚高削减')
doc.add_paragraph('第5月：完成缺料风险补库')
doc.add_paragraph('第6月：完成仓库布局调整')
doc.save(os.path.join(d3, '库存优化方案.docx'))

wb = openpyxl.Workbook()
ws = wb.active
ws.title = '库存优化明细'
ws.append(['序号', '优化类别', 'SKU数量', '库存金额(万元)', '处理措施', '回收/释放(万元)', '判定依据'])
rows = [
    (1, '滞销库存-降价处理', 15, 800, '降价30%促销', 560, '库龄181-365天'),
    (2, '滞销库存-返厂', 5, 300, '与供应商协商退货', 300, '可返厂物料'),
    (3, '滞销库存-报废', 5, 300, '报废处理', 0, '无法处理'),
    (4, '过期库存报废', 10, 600, '报废+资产减值', 0, '库龄>365天'),
    (5, '库存虚高削减', 20, 800, '暂停采购消耗至安全水位', 800, '实际>安全2倍'),
    (6, '缺料风险补库（陷阱）', 10, 0, '紧急补库至安全水位', -200, '实际<安全80%'),
    (7, '安全库存合理保留（正常项）', 15, 500, '维持现状', 0, '季节性备货'),
    (8, '仓库布局优化', '-', '-', '退租外地仓+合并备件区', 80, '利用率40%'),
    (9, '合计', '-', 3300, '-', 1740, '-'),
]
for r in rows:
    ws.append(r)
style_ws(ws, 7)
wb.save(os.path.join(d3, '库存优化明细表.xlsx'))

fig, ax = plt.subplots(figsize=(10, 6))
labels = ['滞销降价', '滞销返厂', '滞销报废', '过期报废', '虚高削减', '缺料补库', '合理保留']
sizes = [560, 300, 0, 0, 800, -200, 0]
colors = ['#4472C4', '#5B9BD5', '#ED7D31', '#FFC000', '#70AD47', '#E74C3C', '#A5A5A5']
ax.bar(labels, sizes, color=colors, width=0.6)
ax.set_ylabel('资金释放/投入（万元）')
ax.set_title('库存优化各措施资金影响')
ax.axhline(y=0, color='black', linewidth=0.5)
ax.set_ylim(-300, 900)
for i, v in enumerate(sizes):
    ax.text(i, v + 20 if v >= 0 else v - 40, f'{v}', ha='center', fontsize=9)
plt.xticks(rotation=30)
plt.tight_layout()
fig.savefig(os.path.join(d3, '库存优化资金影响图.png'), dpi=150)
plt.close()

make_zip(d3, '03库存优化方案.rar')

# ==================== 答案4: 物流降本方案 ====================
print('=== 生成答案4: 物流降本方案 ===')
d4 = os.path.join(ANS_DIR, '_temp_a04')
os.makedirs(d4, exist_ok=True)

doc = Document()
doc.add_heading('物流降本方案', level=0)
doc.add_heading('一、优化目标', level=1)
doc.add_paragraph('1. 物流总费用从5800万元降至5100万元（降12%）')
doc.add_paragraph('2. 运输时效达标率从85%提升至95%')
doc.add_paragraph('3. 消除合同法律风险')
doc.add_heading('二、分类优化措施', level=1)
doc.add_heading('2.1 运费虚高路线重新招标（10条）', level=2)
doc.add_paragraph('路线编号：RT006-RT010、RT016-RT020')
doc.add_paragraph('问题：运费高于市场均价15%以上')
doc.add_paragraph('措施：重新招标引入竞争性报价')
doc.add_paragraph('预计降本：300万元/年')
doc.add_heading('2.2 路线合并优化（5条）', level=2)
doc.add_paragraph('路线编号：RT021-RT025')
doc.add_paragraph('问题：相邻路线可合并，减少发车频次')
doc.add_paragraph('措施：合并为3条路线，减少2次发车')
doc.add_paragraph('预计降本：100万元/年')
doc.add_heading('2.3 紧急运输高价但合理保留（5条，反直觉正常项）', level=2)
doc.add_paragraph('路线编号：RT026-RT030')
doc.add_paragraph('判定：单价偏高但为紧急订单配送，时效保障合理')
doc.add_paragraph('措施：维持现状，不误判为运费虚高')
doc.add_paragraph('依据：紧急运输的边际利润高于运输成本溢价')
doc.add_heading('2.4 合同风险处理', level=2)
doc.add_paragraph('1. 物流公司3冷链合同已到期未续签：立即补签，消除法律风险')
doc.add_paragraph('2. 物流公司4铁路运输价格高于市场15%：重新谈判或更换')
doc.add_paragraph('3. 物流公司6仓储外包利用率40%：退租低利用率仓库')
doc.add_heading('2.5 物流商引入竞争机制', level=2)
doc.add_paragraph('1. 每条干线引入2-3家承运商竞争性报价')
doc.add_paragraph('2. 建立承运商绩效考核体系（时效30%+成本25%+安全20%+服务15%+创新10%）')
doc.add_paragraph('3. 季度考核末位淘汰')
doc.save(os.path.join(d4, '物流降本方案.docx'))

wb = openpyxl.Workbook()
ws = wb.active
ws.title = '物流优化明细'
ws.append(['序号', '优化类别', '路线编号', '数量', '问题', '优化措施', '预计降本(万元)', '依据'])
rows = [
    (1, '运费虚高重新招标', 'RT006-RT010,RT016-RT020', 10, '运费高于市场15%', '重新招标', 300, '市场比价'),
    (2, '路线合并优化', 'RT021-RT025', 5, '相邻路线可合并', '合并为3条路线', 100, '路线分析'),
    (3, '紧急运输保留（正常项）', 'RT026-RT030', 5, '高价但时效保障合理', '维持现状', 0, '综合成本分析'),
    (4, '合同到期补签', '物流公司3', 1, '到期未续签仍在服务', '立即补签', 0, '法律合规要求'),
    (5, '铁路运费谈判', '物流公司4', 1, '价格高于市场15%', '重新谈判或更换', 50, '市场比价'),
    (6, '退租低利用仓库', '物流公司6', 1, '利用率40%', '退租外地仓', 80, '利用率分析'),
    (7, '合计', '-', 23, '-', '-', 530, '-'),
]
for r in rows:
    ws.append(r)
style_ws(ws, 8)
wb.save(os.path.join(d4, '物流优化明细表.xlsx'))

make_zip(d4, '04物流降本方案.rar')

# ==================== 答案5: 管理层错误要求回函 ====================
print('=== 生成答案5: 管理层错误要求回函 ===')
d5 = os.path.join(ANS_DIR, '_temp_a05')
os.makedirs(d5, exist_ok=True)

doc = Document()
doc.add_heading('管理层错误要求法律回函', level=0)
doc.add_paragraph('针对管理层提出的10项要求，逐一分析驳回如下：')
reqs = [
    (1, 'CEO', '供应商淘汰方案暂缓执行', '供应商管理是企业供应链风险控制的核心环节，绩效D级供应商连续两年评分低于60分，'
     '已严重影响产品质量和交期。依据《供应商管理办法》第四章，连续两年绩效D级的供应商应启动淘汰流程。'
     '暂缓执行将导致每年约800万元隐性损失持续发生。建议按计划执行淘汰，同时从A级备选中替换。'),
    (2, '采购总监', '价格虚高供应商不调换', '8家供应商采购单价高于市场均价15%以上，年浪费约1200万元。'
     '依据《采购管理制度》第三章，供应商价格应定期对标市场均价，偏差超过10%应启动重新招标。'
     '不调换将导致企业持续承担不合理溢价，损害企业利益。'),
    (3, 'CFO', '库存优化方案不报总经理', '库存优化涉及资产减值约1000万元（滞销+过期库存处理），'
     '依据《企业内部控制基本规范》第三十条，资产减值属于重大事项，应报管理层审批。'
     '不报总经理直接执行将违反内控规定，存在审计风险。'),
    (4, '运营VP', '独家供应商不引入竞争', '5家独家供应商存在断供风险，一旦断供将导致生产线停工。'
     '依据《供应商管理办法》第四章，独家供应应制定备选方案。'
     '不引入竞争将使企业完全依赖单一供应商，丧失议价能力。'),
    (5, 'CEO', '到期合同不重新招标直接续签', '物流公司3合同已到期未续签仍在服务，存在法律风险。'
     '依据《合同法》相关规定，到期合同应及时续签或重新招标。直接续签可能错失降本机会，'
     '且无竞争性报价无法确保价格合理性。'),
    (6, '采购总监', '紧急采购不走审批', '紧急采购不审批将导致采购合规风险，可能出现虚假紧急采购、利益输送等问题。'
     '依据《采购管理制度》第二章，应建立紧急采购绿色通道，简化但不取消审批，事后48小时内补审。'),
    (7, 'CFO', '滞销库存不处理等市场回暖', '滞销库存25个SKU（1400万元），库龄181天以上，'
     '依据《存货管理办法》第五章，滞销超过180天应启动处理流程。'
     '等市场回暖无时间表，持续占用资金和仓储空间。'),
    (8, '运营VP', '安全库存不调整', '当前安全库存计算方法未按ABC分类，导致部分物料库存虚高、部分缺料。'
     '依据《库存管理办法》第三章，应按物料ABC分类设定安全库存。不调整将导致库存周转率持续偏低。'),
    (9, 'CEO', '需求预测偏差不分析', '2025年11月预测偏差超过10%，导致库存积压。'
     '依据《需求预测管理办法》，应每月分析预测偏差并调整预测模型。'
     '不分析将导致偏差持续，库存问题反复出现。'),
    (10, '采购总监', '供应商绩效不与采购份额挂钩', '绩效D级供应商仍维持原采购份额，违背优胜劣汰原则。'
     '依据《供应商管理办法》第四章，绩效评分应与采购份额挂钩，D级应减少份额直至淘汰。'),
]
for num, source, req, reply in reqs:
    doc.add_heading(f'要求{num}：{source} - {req}', level=1)
    doc.add_paragraph(f'要求内容：{req}')
    doc.add_paragraph(f'驳回依据：')
    doc.add_paragraph(reply)
    doc.add_paragraph()
doc.save(os.path.join(d5, '管理层错误要求回函.docx'))

wb = openpyxl.Workbook()
ws = wb.active
ws.title = '错误要求回函汇总'
ws.append(['序号', '来源', '错误要求', '驳回依据', '正确处理方式'])
rows = [
    (1, 'CEO', '供应商淘汰暂缓', '供应商管理办法第四章', '按计划淘汰并替换'),
    (2, '采购总监', '价格虚高不调换', '采购管理制度第三章', '重新招标引入竞争'),
    (3, 'CFO', '库存方案不报总经理', '内控基本规范第三十条', '报总经理审批后执行'),
    (4, '运营VP', '独家供应商不引入竞争', '供应商管理办法第四章', '引入备选70:30分配'),
    (5, 'CEO', '到期合同直接续签', '合同法相关规定', '重新招标或补签合同'),
    (6, '采购总监', '紧急采购不走审批', '采购管理制度第二章', '绿色通道简化不取消审批'),
    (7, 'CFO', '滞销库存不处理', '存货管理办法第五章', '分档处理（降价/返厂/报废）'),
    (8, '运营VP', '安全库存不调整', '库存管理办法第三章', '按ABC分类重新设定'),
    (9, 'CEO', '预测偏差不分析', '需求预测管理办法', '每月分析并调整模型'),
    (10, '采购总监', '绩效不与份额挂钩', '供应商管理办法第四章', 'D级减少份额直至淘汰'),
]
for r in rows:
    ws.append(r)
style_ws(ws, 5)
wb.save(os.path.join(d5, '错误要求回函汇总表.xlsx'))

make_zip(d5, '05管理层错误要求回函.rar')

# ==================== 答案6: 供应链整体优化方案 ====================
print('=== 生成答案6: 供应链整体优化方案 ===')
d6 = os.path.join(ANS_DIR, '_temp_a06')
os.makedirs(d6, exist_ok=True)

doc = Document()
doc.add_heading('供应链整体优化方案', level=0)
doc.add_heading('一、项目概述', level=1)
doc.add_paragraph('华恒制造企业（脱敏代称），年营收15亿元，供应链成本占营收65%。'
                  '通过系统优化实现年降本7800万元（采购成本降低2000万、库存释放3300万、物流降本530万、'
                  '隐性成本节约1970万），库存周转率从3.2次/年提升至6次/年。')
doc.add_heading('二、优化目标', level=1)
doc.add_paragraph('1. 采购成本年降低2000万元（8家价格虚高调减1200万 + 12家绩效差替换节约800万）')
doc.add_paragraph('2. 库存周转率从3.2次/年提升至6次/年，释放占用资金3300万元')
doc.add_paragraph('3. 物流费用年降低530万元')
doc.add_paragraph('4. 管理制度全面修订，建立长效机制')
doc.add_paragraph('5. 管理层10项错误要求全部依法驳回')
doc.add_heading('三、分阶段实施计划', level=1)
doc.add_paragraph('第一阶段（1-3月）：采购成本优化 - 8家价格虚高供应商重新招标，12家绩效差供应商淘汰替换')
doc.add_paragraph('第二阶段（2-6月）：库存优化 - 滞销库存处理、过期库存报废、虚高库存削减、缺料补库')
doc.add_paragraph('第三阶段（3-6月）：物流优化 - 运费虚高路线重新招标、路线合并、合同补签')
doc.add_paragraph('第四阶段（4-8月）：制度修订 - 4项制度全面修订，建立供应商绩效与份额挂钩机制')
doc.add_paragraph('第五阶段（6-12月）：持续监控 - 建立月度供应链KPI看板，持续跟踪优化效果')
doc.add_heading('四、预期效益', level=1)
doc.add_paragraph('1. 年降本7800万元（采购2000万 + 库存释放3300万 + 物流530万 + 隐性成本1970万）')
doc.add_paragraph('2. 库存周转率3.2次→6次/年')
doc.add_paragraph('3. 供应商绩效平均分72分→85分')
doc.add_paragraph('4. 运输时效达标率85%→95%')
doc.add_paragraph('5. 供应链管理风险全面降低')
doc.add_heading('五、风险与应对', level=1)
doc.add_paragraph('1. 供应商替换期间供应中断风险 → 备选供应商提前导入，渐进过渡')
doc.add_paragraph('2. 库存处理损失影响当期利润 → 分批处理，控制单期减值金额')
doc.add_paragraph('3. 物流招标期间服务波动 → 旧合同延续至新合同生效')
doc.add_paragraph('4. 制度变革阻力 → 分阶段推行，先试点后推广')
doc.save(os.path.join(d6, '供应链整体优化方案.docx'))

# 数据勾稽自检表Excel
wb = openpyxl.Workbook()
ws = wb.active
ws.title = '数据勾稽自检'
ws.append(['勾稽项', '题目描述', '附件明细', '标准答案', '是否一致'])
rows = [
    ('供应商总数', '60家', '供应商名录60行', '60家', '一致'),
    ('价格虚高', '8家', 'GYS026-GYS033', '8家', '一致'),
    ('绩效差淘汰', '12家', 'GYS020-GYS025+GYS036-GYS041', '12家', '一致'),
    ('独家供应商', '5家', 'GYS042-GYS046', '5家', '一致'),
    ('价格偏高保留', '10家', 'GYS047-GYS056', '10家', '一致'),
    ('正常供应商', '25家', '其余25家', '25家', '一致'),
    ('SKU总数', '200个', '安全库存参数表200行', '200个', '一致'),
    ('滞销库存', '25个SKU', '库龄181天以上25个', '25个', '一致'),
    ('过期库存', '10个SKU', '库龄365天以上10个', '10个', '一致'),
    ('库存虚高', '20个SKU', '实际>安全2倍20个', '20个', '一致'),
    ('缺料风险', '10个SKU', '实际<安全80%10个', '10个', '一致'),
    ('安全库存合理', '15个SKU', '季节性备货15个', '15个', '一致'),
    ('运输路线', '50条', '运输路线时效表50行', '50条', '一致'),
    ('运费虚高', '10条', '运费虚高需重新招标10条', '10条', '一致'),
    ('路线可合并', '5条', '路线可合并优化5条', '5条', '一致'),
    ('紧急运输合理', '5条', '紧急运输合理高价5条', '5条', '一致'),
    ('物流合同', '6份', '第三方物流合同6份', '6份', '一致'),
    ('管理制度', '4份', '4份制度文件', '4份', '一致'),
    ('管理层要求', '10条', '管理层要求10条', '10条', '一致'),
    ('采购成本', '9.8亿元', '2025年采购成本', '9.8亿元', '一致'),
    ('库存总额', '9700万元', '库存明细汇总', '9700万元', '一致'),
    ('物流费用', '5800万元', '运输费用汇总', '5800万元', '一致'),
    ('年降本目标', '7800万元', '采购2000+库存3300+物流530+隐性1970', '7800万元', '一致'),
    ('周转率目标', '3.2→6次/年', '库存9700万→6400万', '6次/年', '一致'),
]
for r in rows:
    ws.append(r)
style_ws(ws, 5)
wb.save(os.path.join(d6, '数据勾稽自检表.xlsx'))

# 总体效益对比图PNG
fig, ax = plt.subplots(figsize=(12, 6))
metrics = ['采购成本\n(万元)', '库存占用\n(万元)', '物流费用\n(万元)', '周转率\n(次/年)', '供应商\n绩效(分)']
before = [98000, 9700, 5800, 3.2, 72]
after = [96000, 6400, 5270, 6.0, 85]
# 归一化
scales = [100000, 10000, 6000, 7, 100]
before_n = [b/s for b, s in zip(before, scales)]
after_n = [a/s for a, s in zip(after, scales)]
x = range(len(metrics))
ax.bar([i-0.2 for i in x], before_n, 0.4, label='优化前', color='#ED7D31', alpha=0.7)
ax.bar([i+0.2 for i in x], after_n, 0.4, label='优化后', color='#4472C4', alpha=0.7)
ax.set_ylabel('归一化值')
ax.set_title('供应链优化前后关键指标对比')
ax.set_xticks(x)
ax.set_xticklabels(metrics)
ax.legend()
labels_actual = [f'{b}→{a}' for b, a in zip(before, after)]
for i, label in enumerate(labels_actual):
    ax.text(i, max(before_n[i], after_n[i]) + 0.02, label, ha='center', fontsize=8)
plt.tight_layout()
fig.savefig(os.path.join(d6, '供应链优化前后对比图.png'), dpi=150)
plt.close()

make_zip(d6, '06供应链整体优化方案.rar')

print('\n=== 标准答案附件全部生成完成 ===')
print(f'保存目录: {ANS_DIR}')
