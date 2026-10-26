#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""教师职称审计 - 标准答案附件生成"""
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

BASE = '/workspace/教师职称审计_附件'
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

# ========== 答案1: 人事合规诊断报告 ==========
print('=== 01 人事合规诊断报告 ===')
d = os.path.join(ANS_DIR, '_a1'); os.makedirs(d)
doc = Document(); doc.add_heading('教师人事管理合规诊断报告', level=0)
doc.add_heading('一、诊断概述', level=1)
doc.add_paragraph('某市25所中小学500名教师人事管理审查。发现教师名册中15人职称不符、12人学历造假、8人教龄造假。'
    '80名职称评审中12人条件不符、10人论文造假、8人程序违规。'
    '60名岗位竞聘中10人条件不符、8人暗箱操作、5人未公示。'
    '120份聘用合同中20份代课教师未签合同、15份条款违规、10份到期未续签。'
    '30套绩效分配中8套未公示、6套分配不公、5套依据不足。'
    '50名代课教师中15人未缴社保、12人长期代课、10人工资偏低。')
doc.add_heading('二、核心问题', level=1)
doc.add_paragraph('1. 职称评审不规范30人。2. 岗位竞聘违规23人。3. 聘用合同违规45份。4. 绩效分配不公19套。5. 代课教师权益受损37人。')
doc.add_heading('三、总体结论', level=1)
doc.add_paragraph('教师人事管理存在5大类15项问题，涉及违规金额约300万元（社保补缴、工资差额等）。'
    '通过整改可实现合规率从60%提升至90%。')
doc.save(os.path.join(d, '人事合规诊断报告.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '诊断结论汇总'
ws.append(['序号', '诊断维度', '问题类别', '数量', '影响金额(万元)', '优化方向', '优先级'])
diags = [
    (1, '教师名册', '职称不符', 15, 0, '撤销职称重新评定', '高'),
    (2, '教师名册', '学历造假', 12, 0, '追责并处理', '高'),
    (3, '教师名册', '教龄造假', 8, 0, '核实并纠正', '高'),
    (4, '教师名册', '职称达标（保留）', 5, 0, '维持', '低'),
    (5, '职称评审', '条件不符', 12, 0, '取消评审结果', '高'),
    (6, '职称评审', '论文造假', 10, 0, '取消资格追责', '高'),
    (7, '职称评审', '程序违规', 8, 0, '重新评审', '高'),
    (8, '职称评审', '条件达标（保留）', 5, 0, '维持', '低'),
    (9, '岗位竞聘', '条件不符', 10, 0, '取消竞聘资格', '高'),
    (10, '岗位竞聘', '暗箱操作', 8, 0, '调查追责', '高'),
    (11, '岗位竞聘', '未公示', 5, 0, '补公示', '中'),
    (12, '聘用合同', '未签合同', 20, 80, '补签合同+双倍工资', '高'),
    (13, '聘用合同', '条款违规', 15, 0, '修正条款', '高'),
    (14, '聘用合同', '到期未续签', 10, 0, '补续签', '中'),
    (15, '绩效分配', '未公示', 8, 0, '补公示', '高'),
    (16, '绩效分配', '分配不公', 6, 0, '重新分配', '高'),
    (17, '代课教师', '未缴社保', 15, 120, '补缴社保', '高'),
    (18, '代课教师', '长期代课', 12, 0, '签合同或转正', '高'),
    (19, '代课教师', '工资偏低', 10, 100, '补足差额', '高'),
]
for row in diags: ws.append(row)
style(ws, 7); wb.save(os.path.join(d, '诊断结论汇总表.xlsx'))

fig, ax = plt.subplots(figsize=(8, 8), subplot_kw=dict(projection='polar'))
cats = ['职称评审', '岗位竞聘', '聘用合同', '绩效分配', '代课教师', '名册管理']
cur = [55, 58, 52, 55, 50, 60]; tgt = [90, 90, 92, 88, 90, 92]
angles = np.linspace(0, 2*np.pi, len(cats), endpoint=False).tolist()
cur += cur[:1]; tgt += tgt[:1]; angles += angles[:1]
ax.plot(angles, cur, 'o-', linewidth=2, label='当前', color='#ED7D31')
ax.fill(angles, cur, alpha=0.15, color='#ED7D31')
ax.plot(angles, tgt, 's-', linewidth=2, label='目标', color='#2F5496')
ax.fill(angles, tgt, alpha=0.15, color='#2F5496')
ax.set_xticks(angles[:-1]); ax.set_xticklabels(cats)
ax.set_ylim(0, 100); ax.set_title('人事合规各维度当前 vs 目标', fontsize=12, pad=20)
ax.legend(loc='upper right', bbox_to_anchor=(1.3, 1.1))
plt.tight_layout(); fig.savefig(os.path.join(d, '诊断雷达图.png'), dpi=150); plt.close()
mkzip(d, '01人事合规诊断报告.rar')

# ========== 答案2: 职称评审审查答案 ==========
print('=== 02 职称评审审查答案 ===')
d = os.path.join(ANS_DIR, '_a2'); os.makedirs(d)
doc = Document(); doc.add_heading('职称评审审查方案', level=0)
doc.add_heading('一、审查范围', level=1)
doc.add_paragraph('500名教师名册和80名职称评审记录。')
doc.add_heading('二、分类审查结论', level=1)
doc.add_paragraph('1. 职称不符15人（T0001至T0015），撤销职称。')
doc.add_paragraph('2. 学历造假12人（T0016至T0027），追责。')
doc.add_paragraph('3. 教龄造假8人（T0028至T0035），核实纠正。')
doc.add_paragraph('4. 职称达标5人（T0036至T0040，反直觉正常项）。')
doc.add_paragraph('5. 评审条件不符12人（ZC001至ZC012），取消结果。')
doc.add_paragraph('6. 论文造假10人（ZC013至ZC022），取消追责。')
doc.add_paragraph('7. 程序违规8人（ZC023至ZC030），重新评审。')
doc.save(os.path.join(d, '职称评审审查方案.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '职称审查明细'
ws.append(['序号', '类别', '编号范围', '数量', '问题', '措施', '依据'])
rows = [
    (1, '职称不符', 'T0001-0015', 15, '教龄不足', '撤销职称', '职称评审办法'),
    (2, '学历造假', 'T0016-0027', 12, '伪造学历', '追责', '职称评审办法'),
    (3, '教龄造假', 'T0028-0035', 8, '教龄不实', '核实纠正', '职称评审办法'),
    (4, '职称达标（保留）', 'T0036-0040', 5, '条件符合', '维持', '职称评审办法'),
    (5, '评审条件不符', 'ZC001-012', 12, '不达标', '取消结果', '职称评审办法'),
    (6, '论文造假', 'ZC013-022', 10, '抄袭代写', '取消追责', '职称评审办法'),
    (7, '程序违规', 'ZC023-030', 8, '未公示/委员会违规', '重新评审', '职称评审办法'),
    (8, '评审达标（保留）', 'ZC031-035', 5, '程序合规', '维持', '职称评审办法'),
]
for r in rows: ws.append(r)
style(ws, 7); wb.save(os.path.join(d, '职称审查明细表.xlsx'))

fig, ax = plt.subplots(figsize=(10, 6))
cats = ['职称不符(15)', '学历造假(12)', '教龄造假(8)', '评审不符(12)', '论文造假(10)', '程序违规(8)']
counts = [15, 12, 8, 12, 10, 8]
colors = ['#FF6B6B', '#FFC000', '#ED7D31', '#5B9BD5', '#9B59B6', '#A5A5A5']
ax.bar(cats, counts, color=colors, width=0.6)
ax.set_ylabel('人数'); ax.set_title('职称评审问题分类统计')
plt.xticks(rotation=15, ha='right'); plt.tight_layout()
fig.savefig(os.path.join(d, '问题分布图.png'), dpi=150); plt.close()
mkzip(d, '02职称评审审查答案.rar')

# ========== 答案3: 竞聘合同审查答案 ==========
print('=== 03 竞聘合同审查答案 ===')
d = os.path.join(ANS_DIR, '_a3'); os.makedirs(d)
doc = Document(); doc.add_heading('岗位竞聘与聘用合同审查方案', level=0)
doc.add_heading('一、岗位竞聘审查', level=1)
doc.add_paragraph('1. 条件不符10人（JP001至JP010），取消资格。')
doc.add_paragraph('2. 暗箱操作8人（JP011至JP018），调查追责。')
doc.add_paragraph('3. 未公示5人（JP019至JP023），补公示。')
doc.add_paragraph('4. 合规4人（JP024至JP027，反直觉正常项）。')
doc.add_heading('二、聘用合同审查', level=1)
doc.add_paragraph('1. 未签合同20份（HT001至HT020），补签+双倍工资。')
doc.add_paragraph('2. 条款违规15份（HT021至HT035），修正条款。')
doc.add_paragraph('3. 到期未续签10份（HT036至HT045），补续签。')
doc.add_paragraph('4. 合理合同5份（HT046至HT050，反直觉正常项）。')
doc.save(os.path.join(d, '竞聘合同审查方案.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '竞聘合同审查明细'
ws.append(['序号', '类别', '编号范围', '数量', '问题', '措施', '依据'])
rows = [
    (1, '竞聘条件不符', 'JP001-010', 10, '不达标', '取消资格', '岗位竞聘办法'),
    (2, '竞聘暗箱', 'JP011-018', 8, '面试异常', '调查追责', '岗位竞聘办法'),
    (3, '竞聘未公示', 'JP019-023', 5, '未公示', '补公示', '岗位竞聘办法'),
    (4, '竞聘合规（保留）', 'JP024-027', 4, '程序合规', '维持', '岗位竞聘办法'),
    (5, '未签合同', 'HT001-020', 20, '代课超6月', '补签+双倍工资', '劳动合同法'),
    (6, '条款违规', 'HT021-035', 15, '违法约定', '修正条款', '劳动合同法'),
    (7, '到期未续签', 'HT036-045', 10, '未续签', '补续签', '劳动合同法'),
    (8, '合理合同（保留）', 'HT046-050', 5, '以完成任务为期限', '维持', '劳动合同法'),
]
for r in rows: ws.append(r)
style(ws, 7); wb.save(os.path.join(d, '竞聘合同审查明细表.xlsx'))
mkzip(d, '03竞聘合同审查答案.rar')

# ========== 答案4: 绩效代课审查答案 ==========
print('=== 04 绩效代课审查答案 ===')
d = os.path.join(ANS_DIR, '_a4'); os.makedirs(d)
doc = Document(); doc.add_heading('绩效工资与代课教师审查方案', level=0)
doc.add_heading('一、绩效工资审查', level=1)
doc.add_paragraph('1. 未公示8套，应补公示。')
doc.add_paragraph('2. 分配不公6套，应重新分配。')
doc.add_paragraph('3. 依据不足5套，应完善考核。')
doc.add_paragraph('4. 合理3套（反直觉正常项）。')
doc.add_heading('二、代课教师审查', level=1)
doc.add_paragraph('1. 未缴社保15人，应补缴约120万元。')
doc.add_paragraph('2. 长期代课12人，应签合同或转正。')
doc.add_paragraph('3. 工资偏低10人，应补足差额约100万元。')
doc.add_paragraph('4. 正常短期5人（反直觉正常项）。')
doc.save(os.path.join(d, '绩效代课审查方案.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '绩效代课审查明细'
ws.append(['序号', '类别', '编号范围', '数量', '问题', '措施', '影响金额(万元)', '依据'])
rows = [
    (1, '绩效未公示', '8所学校', 8, '未公示', '补公示', 0, '绩效工资办法'),
    (2, '绩效分配不公', '6所学校', 6, '向管理层倾斜', '重新分配', 0, '绩效工资办法'),
    (3, '绩效依据不足', '5所学校', 5, '无量化考核', '完善考核', 0, '绩效工资办法'),
    (4, '绩效合理（保留）', '3所学校', 3, '经教代会通过', '维持', 0, '绩效工资办法'),
    (5, '未缴社保', 'D001-015', 15, '未缴社保', '补缴', 120, '社会保险法'),
    (6, '长期代课', 'D016-027', 12, '超6月未签合同', '签合同或转正', 0, '劳动合同法'),
    (7, '工资偏低', 'D028-037', 10, '违反同工同酬', '补足差额', 100, '劳动合同法'),
    (8, '正常短期（保留）', 'D038-042', 5, '短期代课合规', '维持', 0, '代课教师办法'),
]
for r in rows: ws.append(r)
style(ws, 8); wb.save(os.path.join(d, '绩效代课审查明细表.xlsx'))

fig, ax = plt.subplots(figsize=(10, 6))
cats = ['社保补缴', '工资差额', '未签合同双倍', '合计']
amounts = [120, 100, 80, 300]
colors = ['#FF6B6B', '#FFC000', '#ED7D31', '#4472C4']
ax.bar(cats, amounts, color=colors, width=0.5)
ax.set_ylabel('金额（万元）'); ax.set_title('代课教师违规整改金额')
for i, v in enumerate(amounts):
    ax.text(i, v+5, f'{v}万', ha='center', fontsize=9)
plt.tight_layout(); fig.savefig(os.path.join(d, '整改金额图.png'), dpi=150); plt.close()
mkzip(d, '04绩效代课审查答案.rar')

# ========== 答案5: 管理层错误要求回函 ==========
print('=== 05 管理层错误要求回函 ===')
d = os.path.join(ANS_DIR, '_a5'); os.makedirs(d)
doc = Document(); doc.add_heading('管理层错误要求回函', level=0)
reqs = [
    (1, '教育局长', '职称造假不追责', '职称评审管理办法规定造假应追责并撤销职称。'),
    (2, '人事科长', '竞聘暗箱不查处', '岗位竞聘管理办法规定暗箱操作应调查追责。'),
    (3, '教育局长', '代课教师不签合同', '劳动合同法规定用工超1个月应签劳动合同。'),
    (4, '人事科长', '社保不补缴', '社会保险法规定用人单位应足额缴纳社保。'),
    (5, '财务科长', '绩效不公不调整', '绩效工资管理办法规定分配应公平并公示。'),
    (6, '教育局长', '合同违规不整改', '劳动合同法规定合同条款应合法。'),
    (7, '人事科长', '学历造假不处理', '职称评审管理办法规定学历造假应处理。'),
    (8, '财务科长', '同工不同酬不纠正', '劳动合同法规定应同工同酬。'),
    (9, '教育局长', '竞聘不公示不补', '岗位竞聘管理办法规定结果应公示。'),
    (10, '人事科长', '长期代课不转正', '劳动合同法规定长期代课应签合同。'),
]
for num, src, req, reply in reqs:
    doc.add_heading(f'要求{num}：{src} - {req}', level=1)
    doc.add_paragraph(f'驳回依据：{reply}')
    doc.add_paragraph()
doc.save(os.path.join(d, '管理层错误要求回函.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '回函汇总'
ws.append(['序号', '来源', '错误要求', '驳回依据'])
for r in reqs:
    ws.append([r[0], r[1], r[2], r[3]])
style(ws, 4); wb.save(os.path.join(d, '错误要求回函汇总表.xlsx'))
mkzip(d, '05管理层错误要求回函.rar')

# ========== 答案6: 整体方案与勾稽自检 ==========
print('=== 06 整体方案与勾稽自检 ===')
d = os.path.join(ANS_DIR, '_a6'); os.makedirs(d)
doc = Document(); doc.add_heading('教师人事合规整改整体方案', level=0)
doc.add_heading('一、项目概述', level=1)
doc.add_paragraph('25所中小学500名教师人事管理审查，发现5大类15项问题，违规金额约300万元。'
    '通过整改可实现合规率从60%提升至90%。')
doc.add_heading('二、优化目标', level=1)
doc.add_paragraph('1. 职称评审规范化\n2. 岗位竞聘公开公平\n3. 聘用合同全签订\n'
    '4. 绩效分配公平透明\n5. 代课教师权益保障\n6. 违规金额整改300万元\n'
    '7. 合规率从60%提升至90%')
doc.add_heading('三、分阶段实施', level=1)
doc.add_paragraph('第一阶段（1-2月）：职称造假处理、竞聘违规调查、合同补签\n'
    '第二阶段（2-3月）：社保补缴、工资差额补足、绩效重分\n'
    '第三阶段（3-4月）：制度修订、长效机制')
doc.add_heading('四、预期效益', level=1)
doc.add_paragraph('1. 补缴社保120万元\n2. 工资差额100万元\n3. 双倍工资80万元\n'
    '4. 合规率60%→90%\n5. 教师权益保障\n6. 建立长效机制')
doc.save(os.path.join(d, '人事合规整改整体方案.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '数据勾稽自检'
ws.append(['勾稽项', '题目描述', '附件明细', '标准答案', '是否一致'])
rows = [
    ('教师总数', '500人', '名册500行', '500人', '一致'),
    ('职称不符', '15人', 'T0001-0015', '15人', '一致'),
    ('学历造假', '12人', 'T0016-0027', '12人', '一致'),
    ('教龄造假', '8人', 'T0028-0035', '8人', '一致'),
    ('职称达标保留', '5人', 'T0036-0040', '5人', '一致'),
    ('评审记录总数', '80人', '明细80行', '80人', '一致'),
    ('评审条件不符', '12人', 'ZC001-012', '12人', '一致'),
    ('论文造假', '10人', 'ZC013-022', '10人', '一致'),
    ('程序违规', '8人', 'ZC023-030', '8人', '一致'),
    ('竞聘记录总数', '60人', '明细60行', '60人', '一致'),
    ('竞聘条件不符', '10人', 'JP001-010', '10人', '一致'),
    ('竞聘暗箱', '8人', 'JP011-018', '8人', '一致'),
    ('竞聘未公示', '5人', 'JP019-023', '5人', '一致'),
    ('聘用合同总数', '120份', '明细120行', '120份', '一致'),
    ('未签合同', '20份', 'HT001-020', '20份', '一致'),
    ('条款违规', '15份', 'HT021-035', '15份', '一致'),
    ('到期未续签', '10份', 'HT036-045', '10份', '一致'),
    ('绩效方案总数', '30套', '明细30行', '30套', '一致'),
    ('绩效未公示', '8套', '8所学校', '8套', '一致'),
    ('绩效分配不公', '6套', '6所学校', '6套', '一致'),
    ('代课教师总数', '50人', '明细50行', '50人', '一致'),
    ('未缴社保', '15人', 'D001-015', '15人', '一致'),
    ('长期代课', '12人', 'D016-027', '12人', '一致'),
    ('工资偏低', '10人', 'D028-037', '10人', '一致'),
    ('管理制度', '5份', '5个Word', '5份', '一致'),
    ('管理层要求', '10条', '要求10条', '10条', '一致'),
    ('违规整改总额', '300万元', '各项加总', '300万元', '一致'),
    ('合规率目标', '60%→90%', '-', '90%', '一致'),
]
for r in rows: ws.append(r)
style(ws, 5); wb.save(os.path.join(d, '数据勾稽自检表.xlsx'))

fig, ax = plt.subplots(figsize=(12, 6))
metrics = ['职称违规(人)', '竞聘违规(人)', '合同违规(份)', '代课违规(人)', '违规金额(万)', '合规率(%)']
before = [30, 23, 45, 37, 300, 60]
after = [0, 0, 0, 0, 0, 90]
scales = [40, 30, 50, 40, 350, 100]
bn = [b/s for b, s in zip(before, scales)]
an = [a/s for a, s in zip(after, scales)]
x = range(len(metrics))
ax.bar([i-0.2 for i in x], bn, 0.4, label='整改前', color='#ED7D31', alpha=0.7)
ax.bar([i+0.2 for i in x], an, 0.4, label='整改后', color='#2F5496', alpha=0.7)
ax.set_ylabel('归一化值'); ax.set_title('人事合规整改前后对比')
ax.set_xticks(x); ax.set_xticklabels(metrics); ax.legend()
for i in range(len(metrics)):
    ax.text(i, max(bn[i], an[i])+0.02, f'{before[i]}→{after[i]}', ha='center', fontsize=8)
plt.tight_layout(); fig.savefig(os.path.join(d, '整改前后对比图.png'), dpi=150); plt.close()
mkzip(d, '06整体方案与勾稽自检.rar')

print('\n=== 标准答案生成完成 ===')
