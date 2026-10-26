#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""义务教育经费审计 - 标准答案附件生成"""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from docx import Document
import zipfile, os, shutil, random
import numpy as np

plt.rcParams['font.sans-serif'] = ['WenQuanYi Micro Hei', 'SimHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

BASE = '/workspace/义务教育经费审计_附件'
ANS_DIR = os.path.join(BASE, '标准答案')
os.makedirs(ANS_DIR, exist_ok=True)

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
    zp = os.path.join(ANS_DIR, name)
    with zipfile.ZipFile(zp, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(folder):
            for f in files:
                fp = os.path.join(root, f)
                zf.write(fp, os.path.relpath(fp, folder))
    shutil.rmtree(folder)
    print(f'  {name} ({os.path.getsize(zp)/1024:.1f}KB)')

# ========== 答案1: 经费使用诊断报告 ==========
print('=== 01 经费使用诊断报告 ===')
d = os.path.join(ANS_DIR, '_a1'); os.makedirs(d)
doc = Document(); doc.add_heading('义务教育经费使用合规诊断报告', level=0)
doc.add_heading('一、诊断概述', level=1)
doc.add_paragraph('某区30所中小学三年义务教育经费投入1.86亿元。'
    '诊断发现8所学校超预算执行（超支约960万元），6所学校经费挪用嫌疑，'
    '5所学校结余过大（执行率低于50%），4所学校生均经费不足。'
    '60条支出明细中10条超范围、8条虚列、6条拆分、4条营养餐挪用。'
    '32笔采购中6笔拆分规避、4笔超限额未招标。'
    '50条营养餐支出中8条挪用、6条虚列、5条标准超标。'
    '40名受资助学生中6名不符条件、4名重复领取。'
    '总体经费执行合规率约70%，低于目标85%。')
doc.add_heading('二、学校经费诊断', level=1)
doc.add_paragraph('1. 超预算8所（XX001至XX008），超支25%-60%，约960万元。')
doc.add_paragraph('2. 经费挪用6所（XX009至XX014），支出与预算不符。')
doc.add_paragraph('3. 结余过大5所（XX015至XX019），执行率低于50%。')
doc.add_paragraph('4. 生均经费不足4所（XX020至XX023），低于600元/人标准。')
doc.add_paragraph('5. 超支但绩效突出3所（XX024至XX026），超支但教学质量和满意度均超目标（反直觉正常项）。')
doc.add_heading('三、总体结论', level=1)
doc.add_paragraph('义务教育经费存在5大类15项问题，年浪费约1800万元。'
    '通过整改可实现年节约1500万元，合规率从70%提升至85%。')
doc.save(os.path.join(d, '经费使用诊断报告.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '诊断结论汇总'
ws.append(['序号', '诊断维度', '问题类别', '数量', '影响金额(万元)', '优化方向', '优先级'])
diags = [
    (1, '学校经费', '超预算执行', 8, 960, '追责并报审批', '高'),
    (2, '学校经费', '经费挪用', 6, 400, '追回并追责', '高'),
    (3, '学校经费', '结余过大', 5, 500, '调整或收回', '中'),
    (4, '学校经费', '生均经费不足', 4, 0, '追加经费', '高'),
    (5, '学校经费', '超支绩效突出（保留）', 3, 0, '维持', '低'),
    (6, '支出明细', '超范围支出', 10, 150, '追回', '高'),
    (7, '支出明细', '虚列支出', 8, 100, '追回追责', '高'),
    (8, '支出明细', '拆分采购', 6, 200, '合并采购', '中'),
    (9, '支出明细', '营养餐挪用', 4, 80, '追回', '高'),
    (10, '采购', '拆分规避招标', 6, 300, '合并招标', '高'),
    (11, '采购', '超限额未招标', 4, 200, '补招标', '高'),
    (12, '资产', '设备闲置', 8, 300, '调配', '中'),
    (13, '营养餐', '标准超标', 5, 50, '按标准执行', '中'),
    (14, '资助', '不符条件资助', 6, 80, '追回', '高'),
    (15, '资助', '重复发放', 4, 30, '追回', '高'),
]
for row in diags: ws.append(row)
style(ws, 7); wb.save(os.path.join(d, '诊断结论汇总表.xlsx'))

fig, ax = plt.subplots(figsize=(8, 8), subplot_kw=dict(projection='polar'))
cats = ['学校经费', '支出明细', '采购合规', '资产管理', '营养餐', '学生资助']
cur = [55, 60, 50, 65, 60, 58]; tgt = [85, 82, 88, 80, 82, 85]
angles = np.linspace(0, 2*np.pi, len(cats), endpoint=False).tolist()
cur += cur[:1]; tgt += tgt[:1]; angles += angles[:1]
ax.plot(angles, cur, 'o-', linewidth=2, label='当前', color='#ED7D31')
ax.fill(angles, cur, alpha=0.15, color='#ED7D31')
ax.plot(angles, tgt, 's-', linewidth=2, label='目标', color='#2F5496')
ax.fill(angles, tgt, alpha=0.15, color='#2F5496')
ax.set_xticks(angles[:-1]); ax.set_xticklabels(cats)
ax.set_ylim(0, 100); ax.set_title('义务教育经费管理各维度当前 vs 目标', fontsize=12, pad=20)
ax.legend(loc='upper right', bbox_to_anchor=(1.3, 1.1))
plt.tight_layout(); fig.savefig(os.path.join(d, '诊断雷达图.png'), dpi=150); plt.close()
mkzip(d, '01经费使用诊断报告.rar')

# ========== 答案2: 学校经费审查答案 ==========
print('=== 02 学校经费审查答案 ===')
d = os.path.join(ANS_DIR, '_a2'); os.makedirs(d)
doc = Document(); doc.add_heading('学校经费审查方案', level=0)
doc.add_heading('一、审查范围', level=1)
doc.add_paragraph('30所中小学，重点审查预算执行、经费挪用、结余、生均经费。')
doc.add_heading('二、分类审查结论', level=1)
doc.add_paragraph('1. 超预算追责（8所）：XX001至XX008，超支25%-60%，约960万元，应追责。')
doc.add_paragraph('2. 经费挪用整改（6所）：XX009至XX014，应追回并追责。')
doc.add_paragraph('3. 结余过大调整（5所）：XX015至XX019，执行率低于50%，应调整预算。')
doc.add_paragraph('4. 生均经费追加（4所）：XX020至XX023，低于600元/人标准，应追加。')
doc.add_paragraph('5. 超支绩效突出保留（3所，反直觉正常项）：XX024至XX026，不误判。')
doc.save(os.path.join(d, '学校经费审查方案.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '学校审查明细'
ws.append(['序号', '审查类别', '学校编号', '数量', '问题', '措施', '影响金额(万元)', '依据'])
rows = [
    (1, '超预算追责', 'XX001-008', 8, '超预算25%-60%', '追责并报审批', 960, '义务教育经费管理办法'),
    (2, '经费挪用', 'XX009-014', 6, '支出与预算不符', '追回并追责', 400, '义务教育经费管理办法'),
    (3, '结余过大', 'XX015-019', 5, '执行率<50%', '调整或收回', 500, '义务教育经费管理办法'),
    (4, '生均经费不足', 'XX020-023', 4, '低于600元/人', '追加经费', 0, '义务教育经费保障机制'),
    (5, '超支绩效突出（保留）', 'XX024-026', 3, '超支但绩效好', '维持', 0, '学校绩效评价'),
    (6, '正常', '其余4所', 4, '无问题', '维持', 0, '-'),
    (7, '合计', '-', 30, '-', '-', 1860, '-'),
]
for r in rows: ws.append(r)
style(ws, 8); wb.save(os.path.join(d, '学校审查明细表.xlsx'))

fig, ax = plt.subplots(figsize=(10, 6))
cats = ['超预算(8)', '挪用(6)', '结余大(5)', '经费不足(4)', '超支绩效好(3)', '正常(4)']
amounts = [960, 400, 500, 0, 0, 0]
colors = ['#FF6B6B', '#FFC000', '#ED7D31', '#5B9BD5', '#70AD47', '#A5A5A5']
ax.bar(cats, amounts, color=colors, width=0.6)
ax.set_ylabel('影响金额（万元）'); ax.set_title('各类问题影响金额')
ax.set_ylim(0, 1100)
for i, v in enumerate(amounts):
    ax.text(i, v+20, f'{v}万', ha='center', fontsize=9)
plt.xticks(rotation=15, ha='right'); plt.tight_layout()
fig.savefig(os.path.join(d, '问题影响金额图.png'), dpi=150); plt.close()
mkzip(d, '02学校经费审查答案.rar')

# ========== 答案3: 支出采购审查答案 ==========
print('=== 03 支出采购审查答案 ===')
d = os.path.join(ANS_DIR, '_a3'); os.makedirs(d)
doc = Document(); doc.add_heading('经费支出与采购审查方案', level=0)
doc.add_heading('一、支出明细审查', level=1)
doc.add_paragraph('1. 超范围10条（PZ0001至PZ0010），追回约150万元。')
doc.add_paragraph('2. 虚列8条（PZ0011至PZ0018），追回约100万元。')
doc.add_paragraph('3. 拆分6条（PZ0019至PZ0024），合并采购。')
doc.add_paragraph('4. 营养餐挪用4条（PZ0025至PZ0028），追回约80万元。')
doc.add_paragraph('5. 正常大额3条（PZ0029至PZ0031），已招标（反直觉正常项）。')
doc.add_heading('二、采购审查', level=1)
doc.add_paragraph('1. 拆分规避6笔（CG001至CG006），合并招标。')
doc.add_paragraph('2. 超限额4笔（CG007至CG010），补招标。')
doc.add_paragraph('3. 单一来源3笔（CG011至CG013），经审批合理（反直觉正常项）。')
doc.add_paragraph('4. 紧急采购2笔（CG014至CG015），经审批合理（反直觉正常项）。')
doc.add_heading('三、资产审查', level=1)
doc.add_paragraph('1. 闲置8台（ZC001至ZC008），调配。')
doc.add_paragraph('2. 利用率低合理4台（ZC014至ZC017），专用实验设备（反直觉正常项）。')
doc.save(os.path.join(d, '支出采购审查方案.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '支出采购审查明细'
ws.append(['序号', '类别', '编号范围', '数量', '问题', '措施', '影响金额(万元)', '依据'])
rows = [
    (1, '超范围支出', 'PZ0001-0010', 10, '与预算用途不符', '追回', 150, '义务教育经费管理办法'),
    (2, '虚列支出', 'PZ0011-0018', 8, '发票不符', '追回追责', 100, '义务教育经费管理办法'),
    (3, '拆分采购', 'PZ0019-0024', 6, '拆分规避', '合并采购', 200, '采购管理办法'),
    (4, '营养餐挪用', 'PZ0025-0028', 4, '挪作他用', '追回', 80, '营养餐管理办法'),
    (5, '正常大额（保留）', 'PZ0029-0031', 3, '已招标', '维持', 0, '采购管理办法'),
    (6, '拆分规避招标', 'CG001-006', 6, '拆分规避', '合并招标', 300, '采购管理办法'),
    (7, '超限额未招标', 'CG007-010', 4, '超限额', '补招标', 200, '采购管理办法'),
    (8, '单一来源合理（保留）', 'CG011-013', 3, '专有设备', '维持', 0, '采购管理办法'),
    (9, '紧急采购合理（保留）', 'CG014-015', 2, '教学急需', '维持', 0, '采购管理办法'),
    (10, '设备闲置', 'ZC001-008', 8, '闲置>6月', '调配', 300, '资产管理办法'),
    (11, '利用率低合理（保留）', 'ZC014-017', 4, '专用设备', '维持', 0, '资产管理办法'),
]
for r in rows: ws.append(r)
style(ws, 8); wb.save(os.path.join(d, '支出采购审查明细表.xlsx'))
mkzip(d, '03支出采购审查答案.rar')

# ========== 答案4: 营养餐与资助审查答案 ==========
print('=== 04 营养餐与资助审查答案 ===')
d = os.path.join(ANS_DIR, '_a4'); os.makedirs(d)
doc = Document(); doc.add_heading('营养餐与学生资助审查方案', level=0)
doc.add_heading('一、营养餐审查', level=1)
doc.add_paragraph('1. 挪用8条（YY0001至YY0008），追回约80万元。')
doc.add_paragraph('2. 虚列6条（YY0009至YY0014），追回约60万元。')
doc.add_paragraph('3. 标准超标5条（YY0015至YY0019），按4元/人执行。')
doc.add_paragraph('4. 正常大额3条（YY0020至YY0022），已招标（反直觉正常项）。')
doc.add_heading('二、学生资助审查', level=1)
doc.add_paragraph('1. 不符条件6名（ZZ001至ZZ006），追回资助金约80万元。')
doc.add_paragraph('2. 重复发放4名（ZZ007至ZZ010），追回重复部分约30万元。')
doc.add_paragraph('3. 金额超标3名（ZZ011至ZZ013），按标准1500元执行。')
doc.add_heading('三、整改要求', level=1)
doc.add_paragraph('营养餐应严格按4元/人标准执行，建立食材采购验收和公示制度。'
    '学生资助应严格审核认定材料，建立重复发放核查机制。')
doc.save(os.path.join(d, '营养餐资助审查方案.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '营养餐资助审查明细'
ws.append(['序号', '类别', '编号范围', '数量', '问题', '措施', '影响金额(万元)', '依据'])
rows = [
    (1, '营养餐挪用', 'YY0001-0008', 8, '挪作他用', '追回', 80, '营养餐管理办法'),
    (2, '营养餐虚列', 'YY0009-0014', 6, '虚列采购', '追回追责', 60, '营养餐管理办法'),
    (3, '标准超标', 'YY0015-0019', 5, '5元超标准4元', '按标准执行', 50, '营养餐管理办法'),
    (4, '正常大额（保留）', 'YY0020-0022', 3, '已招标', '维持', 0, '采购管理办法'),
    (5, '不符条件资助', 'ZZ001-006', 6, '材料不全', '追回', 80, '学生资助管理办法'),
    (6, '重复发放', 'ZZ007-010', 4, '重复领取', '追回', 30, '学生资助管理办法'),
    (7, '金额超标', 'ZZ011-013', 3, '超1500元标准', '按标准执行', 0, '学生资助管理办法'),
]
for r in rows: ws.append(r)
style(ws, 8); wb.save(os.path.join(d, '营养餐资助审查明细表.xlsx'))
mkzip(d, '04营养餐资助审查答案.rar')

# ========== 答案5: 管理层错误要求回函 ==========
print('=== 05 管理层错误要求回函 ===')
d = os.path.join(ANS_DIR, '_a5'); os.makedirs(d)
doc = Document(); doc.add_heading('管理层错误要求回函', level=0)
reqs = [
    (1, '教育局长', '超预算不追责', '义务教育经费管理办法规定超预算20%以上应报审批并追责。'),
    (2, '财务科长', '经费挪用不处理', '义务教育经费管理办法规定经费挪用应追回并追责。'),
    (3, '基教科长', '结余过大不调整', '义务教育经费管理办法规定执行率低于50%应调整预算。'),
    (4, '设备科长', '拆分采购不追究', '采购管理办法规定不得拆分规避招标。'),
    (5, '教育局长', '闲置设备不调配', '资产管理办法规定闲置超6个月应调配。'),
    (6, '财务科长', '营养餐挪用不追回', '营养餐管理办法规定营养餐经费应专款专用。'),
    (7, '资助科长', '不符条件资助不追回', '学生资助管理办法规定不符条件应追回资助金。'),
    (8, '设备科长', '单一来源不审批', '采购管理办法规定单一来源应审批。'),
    (9, '教育局长', '生均经费不足不追加', '义务教育经费保障机制规定生均经费不得低于标准。'),
    (10, '基教科长', '绩效只看升学率', '学校绩效评价办法规定应多维评价含经费执行。'),
]
for num, src, req, reply in reqs:
    doc.add_heading(f'要求{num}：{src} - {req}', level=1)
    doc.add_paragraph(f'驳回依据：{reply}')
    doc.add_paragraph()
doc.save(os.path.join(d, '管理层错误要求回函.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '回函汇总'
ws.append(['序号', '来源', '错误要求', '驳回依据', '正确处理'])
for r in reqs:
    ws.append([r[0], r[1], r[2], r[3].split('。')[0], r[3].split('。')[1] if '。' in r[3] else ''])
style(ws, 5); wb.save(os.path.join(d, '错误要求回函汇总表.xlsx'))
mkzip(d, '05管理层错误要求回函.rar')

# ========== 答案6: 整体方案与勾稽自检 ==========
print('=== 06 整体方案与勾稽自检 ===')
d = os.path.join(ANS_DIR, '_a6'); os.makedirs(d)
doc = Document(); doc.add_heading('义务教育经费审计整体方案', level=0)
doc.add_heading('一、项目概述', level=1)
doc.add_paragraph('某区30所中小学三年经费1.86亿元。审计发现5大类15项问题，年浪费约1800万元。'
    '通过整改可实现年节约1500万元，合规率从70%提升至85%。')
doc.add_heading('二、优化目标', level=1)
doc.add_paragraph('1. 学校经费：超预算追责960万、挪用追回400万、结余调整500万\n'
    '2. 支出采购：超范围追回150万、虚列追回100万、采购整改500万\n'
    '3. 营养餐资助：挪用追回80万、资助追回110万\n4. 管理层10项错误要求全部驳回\n'
    '5. 经费执行合规率从70%提升至85%')
doc.add_heading('三、分阶段实施', level=1)
doc.add_paragraph('第一阶段（1-2月）：超预算追责、挪用追回、虚列追回\n'
    '第二阶段（2-4月）：拆分合并、超限额补招标、闲置调配\n'
    '第三阶段（3-5月）：营养餐整改、资助追回\n'
    '第四阶段（4-8月）：制度修订、长效机制')
doc.add_heading('四、预期效益', level=1)
doc.add_paragraph('1. 年节约1500万元\n2. 合规率从70%提升至85%\n'
    '3. 采购合规率从60%提升至90%\n4. 营养餐资金使用率达100%\n'
    '5. 资助准确率达100%')
doc.save(os.path.join(d, '义务教育审计整体方案.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '数据勾稽自检'
ws.append(['勾稽项', '题目描述', '附件明细', '标准答案', '是否一致'])
rows = [
    ('学校总数', '30所', '台账30行', '30所', '一致'),
    ('超预算', '8所', 'XX001-008', '8所', '一致'),
    ('经费挪用', '6所', 'XX009-014', '6所', '一致'),
    ('结余过大', '5所', 'XX015-019', '5所', '一致'),
    ('生均经费不足', '4所', 'XX020-023', '4所', '一致'),
    ('超支绩效突出', '3所', 'XX024-026', '3所', '一致'),
    ('正常学校', '4所', '其余4所', '4所', '一致'),
    ('支出明细总数', '60条', '明细60行', '60条', '一致'),
    ('超范围支出', '10条', 'PZ0001-0010', '10条', '一致'),
    ('虚列支出', '8条', 'PZ0011-0018', '8条', '一致'),
    ('拆分采购', '6条', 'PZ0019-0024', '6条', '一致'),
    ('营养餐挪用', '4条', 'PZ0025-0028', '4条', '一致'),
    ('正常大额', '3条', 'PZ0029-0031', '3条', '一致'),
    ('正常支出', '29条', '其余29条', '29条', '一致'),
    ('采购总数', '32笔', '采购32行', '32笔', '一致'),
    ('拆分规避', '6笔', 'CG001-006', '6笔', '一致'),
    ('超限额未招标', '4笔', 'CG007-010', '4笔', '一致'),
    ('单一来源合理', '3笔', 'CG011-013', '3笔', '一致'),
    ('紧急采购合理', '2笔', 'CG014-015', '2笔', '一致'),
    ('设备总数', '30台', '台账30行', '30台', '一致'),
    ('闲置设备', '8台', 'ZC001-008', '8台', '一致'),
    ('营养餐总数', '50条', '明细50行', '50条', '一致'),
    ('营养餐挪用', '8条', 'YY0001-0008', '8条', '一致'),
    ('营养餐虚列', '6条', 'YY0009-0014', '6条', '一致'),
    ('资助总数', '40名', '明细40行', '40名', '一致'),
    ('不符条件', '6名', 'ZZ001-006', '6名', '一致'),
    ('重复发放', '4名', 'ZZ007-010', '4名', '一致'),
    ('管理制度', '5份', '5个Word', '5份', '一致'),
    ('管理层要求', '10条', '要求10条', '10条', '一致'),
    ('经费总额', '1.86亿元', '趋势图', '1.86亿元', '一致'),
    ('年节约目标', '1500万元', '各项加总', '1500万元', '一致'),
    ('合规率目标', '70%→85%', '-', '85%', '一致'),
]
for r in rows: ws.append(r)
style(ws, 5); wb.save(os.path.join(d, '数据勾稽自检表.xlsx'))

fig, ax = plt.subplots(figsize=(12, 6))
metrics = ['超预算(万)', '挪用(万)', '采购违规(万)', '营养餐(万)', '资助违规(万)', '合规率(%)']
before = [960, 400, 500, 140, 110, 70]
after = [0, 0, 0, 0, 0, 85]
scales = [1000, 500, 600, 200, 150, 100]
bn = [b/s for b, s in zip(before, scales)]
an = [a/s for a, s in zip(after, scales)]
x = range(len(metrics))
ax.bar([i-0.2 for i in x], bn, 0.4, label='优化前', color='#ED7D31', alpha=0.7)
ax.bar([i+0.2 for i in x], an, 0.4, label='优化后', color='#2F5496', alpha=0.7)
ax.set_ylabel('归一化值'); ax.set_title('义务教育经费审计优化前后对比')
ax.set_xticks(x); ax.set_xticklabels(metrics); ax.legend()
for i in range(len(metrics)):
    ax.text(i, max(bn[i], an[i])+0.02, f'{before[i]}→{after[i]}', ha='center', fontsize=8)
plt.tight_layout(); fig.savefig(os.path.join(d, '优化前后对比图.png'), dpi=150); plt.close()
mkzip(d, '06整体方案与勾稽自检.rar')

print('\n=== 标准答案生成完成 ===')
