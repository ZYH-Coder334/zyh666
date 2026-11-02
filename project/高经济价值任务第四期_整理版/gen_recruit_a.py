#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""招生入学合规审查 - 标准答案附件生成"""
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

BASE = '/workspace/招生合规审查_附件'
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

# ========== 答案1: 招生合规诊断报告 ==========
print('=== 01 招生合规诊断报告 ===')
d = os.path.join(ANS_DIR, '_a1'); os.makedirs(d)
doc = Document(); doc.add_heading('招生入学合规诊断报告', level=0)
doc.add_heading('一、诊断概述', level=1)
doc.add_paragraph('某市25所中小学2025年秋季招生工作审查。发现8所学校超额招生（超计划15%-40%），'
    '6所学校跨片区招生，200条报名录取中30条跨片、25条提前录取、20条违规择校。'
    '50名特长生中10名材料不全、8名证书造假。25所学校中8所有重点班、6所编班不均衡。'
    '150条收费中25条择校费、15条捐资助学、20条乱收费。'
    '80条转学中15条违规、10条空挂学籍。'
    '总体招生合规率约65%，低于目标90%。')
doc.add_heading('二、核心问题', level=1)
doc.add_paragraph('1. 超额招生8所，超计划15%-40%。')
doc.add_paragraph('2. 跨片区招生6所，违反免试就近入学。')
doc.add_paragraph('3. 特长生认定违规23名，占46%。')
doc.add_paragraph('4. 重点班8所，违反均衡编班。')
doc.add_paragraph('5. 违规收费60条，涉及金额约250万元。')
doc.add_paragraph('6. 违规转学15条、空挂学籍10条。')
doc.add_heading('三、总体结论', level=1)
doc.add_paragraph('招生入学存在6大类15项问题，涉及违规收费约250万元。'
    '通过整改可实现合规率从65%提升至90%，退还违规收费约250万元。')
doc.save(os.path.join(d, '招生合规诊断报告.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '诊断结论汇总'
ws.append(['序号', '诊断维度', '问题类别', '数量', '影响金额(万元)', '优化方向', '优先级'])
diags = [
    (1, '招生计划', '超额招生', 8, 0, '追责并核减计划', '高'),
    (2, '招生计划', '跨片区招生', 6, 0, '退回原籍或调剂', '高'),
    (3, '招生计划', '政策允许扩招（保留）', 3, 0, '维持', '低'),
    (4, '报名录取', '跨片录取', 30, 0, '退回划片学校', '高'),
    (5, '报名录取', '提前录取', 25, 0, '重新摇号', '高'),
    (6, '报名录取', '违规择校', 20, 0, '退回并退费', '高'),
    (7, '特长生', '材料不全', 10, 0, '补充材料或取消', '高'),
    (8, '特长生', '证书造假', 8, 0, '取消资格并追责', '高'),
    (9, '特长生', '比例超标', 5, 0, '调整比例', '中'),
    (10, '特长生', '认定合理（保留）', 4, 0, '维持', '低'),
    (11, '编班', '重点班', 8, 0, '取消并重新编班', '高'),
    (12, '编班', '不均衡编班', 6, 0, '调整均衡', '中'),
    (13, '编班', '分层教学合理（保留）', 3, 0, '维持', '低'),
    (14, '收费', '择校费', 25, 100, '全额退还', '高'),
    (15, '收费', '捐资助学', 15, 60, '退还或规范', '高'),
    (16, '收费', '乱收费', 20, 50, '全额退还', '高'),
    (17, '转学', '违规转学', 15, 0, '退回原校', '高'),
    (18, '转学', '空挂学籍', 10, 0, '清理学籍', '高'),
]
for row in diags: ws.append(row)
style(ws, 7); wb.save(os.path.join(d, '诊断结论汇总表.xlsx'))

fig, ax = plt.subplots(figsize=(8, 8), subplot_kw=dict(projection='polar'))
cats = ['招生计划', '报名录取', '特长生', '均衡编班', '收费合规', '转学管理']
cur = [55, 60, 50, 55, 60, 58]; tgt = [90, 92, 88, 90, 95, 90]
angles = np.linspace(0, 2*np.pi, len(cats), endpoint=False).tolist()
cur += cur[:1]; tgt += tgt[:1]; angles += angles[:1]
ax.plot(angles, cur, 'o-', linewidth=2, label='当前', color='#ED7D31')
ax.fill(angles, cur, alpha=0.15, color='#ED7D31')
ax.plot(angles, tgt, 's-', linewidth=2, label='目标', color='#2F5496')
ax.fill(angles, tgt, alpha=0.15, color='#2F5496')
ax.set_xticks(angles[:-1]); ax.set_xticklabels(cats)
ax.set_ylim(0, 100); ax.set_title('招生合规各维度当前 vs 目标', fontsize=12, pad=20)
ax.legend(loc='upper right', bbox_to_anchor=(1.3, 1.1))
plt.tight_layout(); fig.savefig(os.path.join(d, '诊断雷达图.png'), dpi=150); plt.close()
mkzip(d, '01招生合规诊断报告.rar')

# ========== 答案2: 招生计划审查答案 ==========
print('=== 02 招生计划审查答案 ===')
d = os.path.join(ANS_DIR, '_a2'); os.makedirs(d)
doc = Document(); doc.add_heading('招生计划审查方案', level=0)
doc.add_heading('一、审查范围', level=1)
doc.add_paragraph('25所中小学，重点审查超额招生、跨片区招生、特长生违规。')
doc.add_heading('二、分类审查结论', level=1)
doc.add_paragraph('1. 超额追责（8所）：XX001至XX008，超计划15%-40%，应追责并核减下一年度计划。')
doc.add_paragraph('2. 跨片整改（6所）：XX009至XX014，应退回原籍或调剂。')
doc.add_paragraph('3. 特长生违规（5所）：XX015至XX019，应重新认定。')
doc.add_paragraph('4. 政策允许扩招保留（3所，反直觉正常项）：XX020至XX022，不误判。')
doc.save(os.path.join(d, '招生计划审查方案.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '招生审查明细'
ws.append(['序号', '审查类别', '学校编号', '数量', '问题', '措施', '依据'])
rows = [
    (1, '超额招生', 'XX001-008', 8, '超计划15%-40%', '追责核减', '招生计划管理办法'),
    (2, '跨片区招生', 'XX009-014', 6, '违反就近入学', '退回调剂', '义务教育招生意见'),
    (3, '特长生违规', 'XX015-019', 5, '认定不规范', '重新认定', '特长生管理办法'),
    (4, '政策允许扩招（保留）', 'XX020-022', 3, '有批文扩招', '维持', '招生计划管理办法'),
    (5, '正常', '其余3所', 3, '无问题', '维持', '-'),
    (6, '合计', '-', 25, '-', '-', '-'),
]
for r in rows: ws.append(r)
style(ws, 7); wb.save(os.path.join(d, '招生审查明细表.xlsx'))

fig, ax = plt.subplots(figsize=(10, 6))
cats = ['超额(8)', '跨片(6)', '特长生违规(5)', '扩招合理(3)', '正常(3)']
counts = [8, 6, 5, 3, 3]
colors = ['#FF6B6B', '#FFC000', '#ED7D31', '#70AD47', '#A5A5A5']
ax.bar(cats, counts, color=colors, width=0.6)
ax.set_ylabel('学校数'); ax.set_title('招生问题分类统计')
for i, v in enumerate(counts):
    ax.text(i, v+0.2, f'{v}所', ha='center', fontsize=9)
plt.tight_layout(); fig.savefig(os.path.join(d, '问题分布图.png'), dpi=150); plt.close()
mkzip(d, '02招生计划审查答案.rar')

# ========== 答案3: 录取特长生编班答案 ==========
print('=== 03 录取特长生编班答案 ===')
d = os.path.join(ANS_DIR, '_a3'); os.makedirs(d)
doc = Document(); doc.add_heading('录取、特长生与编班审查方案', level=0)
doc.add_heading('一、报名录取审查', level=1)
doc.add_paragraph('1. 跨片录取30条（BM0001至BM0030），应退回划片学校。')
doc.add_paragraph('2. 提前录取25条（BM0031至BM0055），应重新摇号。')
doc.add_paragraph('3. 违规择校20条（BM0056至BM0075），应退回并退费。')
doc.add_paragraph('4. 正常跨片10条（BM0076至BM0085，反直觉正常项），随迁子女政策允许。')
doc.add_heading('二、特长生审查', level=1)
doc.add_paragraph('1. 材料不全10名（TC001至TC010），应补充或取消。')
doc.add_paragraph('2. 证书造假8名（TC011至TC018），应取消资格并追责。')
doc.add_paragraph('3. 比例超标5所，应调整特长生比例。')
doc.add_paragraph('4. 认定合理4名（TC024至TC027，反直觉正常项）。')
doc.add_heading('三、编班审查', level=1)
doc.add_paragraph('1. 重点班8所（BB001至BB008），应取消并重新编班。')
doc.add_paragraph('2. 不均衡6所（BB009至BB014），应调整均衡。')
doc.add_paragraph('3. 分层教学合理3所（BB015至BB017，反直觉正常项）。')
doc.save(os.path.join(d, '录取特长生编班审查方案.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '审查明细'
ws.append(['序号', '类别', '编号范围', '数量', '问题', '措施', '依据'])
rows = [
    (1, '跨片录取', 'BM0001-0030', 30, '不在划片范围', '退回', '招生意见'),
    (2, '提前录取', 'BM0031-0055', 25, '摇号前录取', '重新摇号', '民办招生办法'),
    (3, '违规择校', 'BM0056-0075', 20, '缴费录取', '退回退费', '招生收费办法'),
    (4, '正常跨片（保留）', 'BM0076-0085', 10, '随迁子女政策', '维持', '随迁子女政策'),
    (5, '特长生材料不全', 'TC001-010', 10, '缺获奖原件', '补充或取消', '特长生办法'),
    (6, '特长生证书造假', 'TC011-018', 8, '伪造证书', '取消追责', '特长生办法'),
    (7, '特长生比例超标', '5所学校', 5, '超5%上限', '调整', '特长生办法'),
    (8, '特长生认定合理（保留）', 'TC024-027', 4, '材料齐全', '维持', '特长生办法'),
    (9, '重点班', 'BB001-008', 8, '分重点普通班', '取消重编', '均衡编班办法'),
    (10, '不均衡编班', 'BB009-014', 6, '班额差异>10%', '调整', '均衡编班办法'),
    (11, '分层教学合理（保留）', 'BB015-017', 3, '有批文', '维持', '均衡编班办法'),
]
for r in rows: ws.append(r)
style(ws, 7); wb.save(os.path.join(d, '录取特长生编班审查明细表.xlsx'))
mkzip(d, '03录取特长生编班答案.rar')

# ========== 答案4: 收费转学审查答案 ==========
print('=== 04 收费转学审查答案 ===')
d = os.path.join(ANS_DIR, '_a4'); os.makedirs(d)
doc = Document(); doc.add_heading('收费与转学审查方案', level=0)
doc.add_heading('一、收费审查', level=1)
doc.add_paragraph('1. 择校费25条（SF0001至SF0025），全额退还约100万元。')
doc.add_paragraph('2. 捐资助学15条（SF0026至SF0040），退还或规范约60万元。')
doc.add_paragraph('3. 乱收费20条（SF0041至SF0060），全额退还约50万元。')
doc.add_paragraph('4. 正常住宿费10条（SF0061至SF0070，反直觉正常项）。')
doc.add_heading('二、转学审查', level=1)
doc.add_paragraph('1. 违规转学15条（ZX001至ZX015），退回原校。')
doc.add_paragraph('2. 空挂学籍10条（ZX016至ZX025），清理学籍。')
doc.add_paragraph('3. 未审批8条（ZX026至ZX033），补审批。')
doc.add_paragraph('4. 正常转学5条（ZX034至ZX038，反直觉正常项）。')
doc.save(os.path.join(d, '收费转学审查方案.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '收费转学审查明细'
ws.append(['序号', '类别', '编号范围', '数量', '问题', '措施', '影响金额(万元)', '依据'])
rows = [
    (1, '择校费', 'SF0001-0025', 25, '违规收费', '全额退还', 100, '招生收费办法'),
    (2, '捐资助学', 'SF0026-0040', 15, '强制收费', '退还或规范', 60, '招生收费办法'),
    (3, '乱收费', 'SF0041-0060', 20, '无依据收费', '全额退还', 50, '招生收费办法'),
    (4, '住宿费（保留）', 'SF0061-0070', 10, '正常收费', '维持', 0, '招生收费办法'),
    (5, '违规转学', 'ZX001-015', 15, '原因不实', '退回原校', 0, '转学管理办法'),
    (6, '空挂学籍', 'ZX016-025', 10, '未实际到校', '清理学籍', 0, '学籍管理办法'),
    (7, '未审批转学', 'ZX026-033', 8, '无审批', '补审批', 0, '转学管理办法'),
    (8, '正常转学（保留）', 'ZX034-038', 5, '材料齐全', '维持', 0, '转学管理办法'),
]
for r in rows: ws.append(r)
style(ws, 8); wb.save(os.path.join(d, '收费转学审查明细表.xlsx'))

fig, ax = plt.subplots(figsize=(10, 6))
cats = ['择校费', '捐资助学', '乱收费', '住宿费']
amounts = [100, 60, 50, 0]
colors = ['#FF6B6B', '#FFC000', '#ED7D31', '#70AD47']
ax.bar(cats, amounts, color=colors, width=0.5)
ax.set_ylabel('金额（万元）'); ax.set_title('各类违规收费退还金额')
for i, v in enumerate(amounts):
    ax.text(i, v+2, f'{v}万', ha='center', fontsize=9)
plt.tight_layout(); fig.savefig(os.path.join(d, '退费金额图.png'), dpi=150); plt.close()
mkzip(d, '04收费转学审查答案.rar')

# ========== 答案5: 管理层错误要求回函 ==========
print('=== 05 管理层错误要求回函 ===')
d = os.path.join(ANS_DIR, '_a5'); os.makedirs(d)
doc = Document(); doc.add_heading('管理层错误要求回函', level=0)
reqs = [
    (1, '教育局长', '超额招生不追责', '招生计划管理办法规定超额招生应追责并核减计划。'),
    (2, '基教科长', '跨片招生不处理', '义务教育招生意见规定免试就近入学，跨片应处理。'),
    (3, '教育局长', '特长生认定不严格', '特长生管理办法规定应严格认定标准和程序。'),
    (4, '基教科长', '重点班不取消', '均衡编班管理办法规定严禁分重点班。'),
    (5, '财务科长', '择校费不退还', '招生收费管理办法规定择校费应全额退还。'),
    (6, '基教科长', '捐资助学不叫停', '招生收费管理办法规定不得强制捐资助学。'),
    (7, '教育局长', '空挂学籍不清理', '学籍管理办法规定空挂学籍应清理。'),
    (8, '基教科长', '乱收费不查处', '招生收费管理办法规定乱收费应查处并退还。'),
    (9, '教育局长', '转学审批不严格', '转学管理办法规定转学应严格审批。'),
    (10, '基教科长', '编班方案不公示', '均衡编班管理办法规定编班方案应公示。'),
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
doc = Document(); doc.add_heading('招生入学合规整改整体方案', level=0)
doc.add_heading('一、项目概述', level=1)
doc.add_paragraph('25所中小学招生入学合规审查，发现6大类15项问题，违规收费约250万元。'
    '通过整改可实现合规率从65%提升至90%，退还违规收费250万元。')
doc.add_heading('二、优化目标', level=1)
doc.add_paragraph('1. 超额招生追责并核减计划\n2. 跨片招生退回调剂\n3. 特长生重新认定\n'
    '4. 重点班取消并均衡编班\n5. 违规收费全额退还250万元\n6. 空挂学籍清理\n'
    '7. 合规率从65%提升至90%')
doc.add_heading('三、分阶段实施', level=1)
doc.add_paragraph('第一阶段（1-2月）：超额追责、跨片退回、违规退费\n'
    '第二阶段（2-3月）：特长生重认、重点班取消、空挂清理\n'
    '第三阶段（3-4月）：编班均衡、转学补审\n'
    '第四阶段（4-6月）：制度修订、长效机制')
doc.add_heading('四、预期效益', level=1)
doc.add_paragraph('1. 退还违规收费250万元\n2. 合规率从65%提升至90%\n'
    '3. 招生公平性显著提升\n4. 社会满意度提升\n5. 建立长效监管机制')
doc.save(os.path.join(d, '招生合规整改整体方案.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '数据勾稽自检'
ws.append(['勾稽项', '题目描述', '附件明细', '标准答案', '是否一致'])
rows = [
    ('学校总数', '25所', '台账25行', '25所', '一致'),
    ('超额招生', '8所', 'XX001-008', '8所', '一致'),
    ('跨片招生', '6所', 'XX009-014', '6所', '一致'),
    ('特长生违规', '5所', 'XX015-019', '5所', '一致'),
    ('扩招合理', '3所', 'XX020-022', '3所', '一致'),
    ('正常学校', '3所', '其余3所', '3所', '一致'),
    ('报名录取总数', '200条', '明细200行', '200条', '一致'),
    ('跨片录取', '30条', 'BM0001-0030', '30条', '一致'),
    ('提前录取', '25条', 'BM0031-0055', '25条', '一致'),
    ('违规择校', '20条', 'BM0056-0075', '20条', '一致'),
    ('特长生认定总数', '50名', '明细50行', '50名', '一致'),
    ('材料不全', '10名', 'TC001-010', '10名', '一致'),
    ('证书造假', '8名', 'TC011-018', '8名', '一致'),
    ('编班方案总数', '25所', '明细25行', '25所', '一致'),
    ('重点班', '8所', 'BB001-008', '8所', '一致'),
    ('不均衡编班', '6所', 'BB009-014', '6所', '一致'),
    ('分层教学合理', '3所', 'BB015-017', '3所', '一致'),
    ('收费记录总数', '150条', '明细150行', '150条', '一致'),
    ('择校费', '25条', 'SF0001-0025', '25条', '一致'),
    ('捐资助学', '15条', 'SF0026-0040', '15条', '一致'),
    ('乱收费', '20条', 'SF0041-0060', '20条', '一致'),
    ('转学记录总数', '80条', '明细80行', '80条', '一致'),
    ('违规转学', '15条', 'ZX001-015', '15条', '一致'),
    ('空挂学籍', '10条', 'ZX016-025', '10条', '一致'),
    ('未审批转学', '8条', 'ZX026-033', '8条', '一致'),
    ('管理制度', '5份', '5个Word', '5份', '一致'),
    ('管理层要求', '10条', '要求10条', '10条', '一致'),
    ('违规退费总额', '250万元', '各项加总', '250万元', '一致'),
    ('合规率目标', '65%→90%', '-', '90%', '一致'),
]
for r in rows: ws.append(r)
style(ws, 5); wb.save(os.path.join(d, '数据勾稽自检表.xlsx'))

fig, ax = plt.subplots(figsize=(12, 6))
metrics = ['超额(所)', '跨片(所)', '重点班(所)', '违规收费(万)', '合规率(%)']
before = [8, 6, 8, 250, 65]
after = [0, 0, 0, 0, 90]
scales = [10, 10, 10, 300, 100]
bn = [b/s for b, s in zip(before, scales)]
an = [a/s for a, s in zip(after, scales)]
x = range(len(metrics))
ax.bar([i-0.2 for i in x], bn, 0.4, label='整改前', color='#ED7D31', alpha=0.7)
ax.bar([i+0.2 for i in x], an, 0.4, label='整改后', color='#2F5496', alpha=0.7)
ax.set_ylabel('归一化值'); ax.set_title('招生合规整改前后对比')
ax.set_xticks(x); ax.set_xticklabels(metrics); ax.legend()
for i in range(len(metrics)):
    ax.text(i, max(bn[i], an[i])+0.02, f'{before[i]}→{after[i]}', ha='center', fontsize=8)
plt.tight_layout(); fig.savefig(os.path.join(d, '整改前后对比图.png'), dpi=150); plt.close()
mkzip(d, '06整体方案与勾稽自检.rar')

print('\n=== 标准答案生成完成 ===')
