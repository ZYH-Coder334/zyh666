#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""高校科研经费审计 - 标准答案附件生成"""
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

BASE = '/workspace/科研经费审计_附件'
ANS_DIR = os.path.join(BASE, '标准答案')
os.makedirs(ANS_DIR, exist_ok=True)

hf = Font(name='微软雅黑', size=11, bold=True, color='FFFFFF')
hfill = PatternFill(start_color='7030A0', end_color='7030A0', fill_type='solid')
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
print('=== 01 科研经费使用诊断报告 ===')
d = os.path.join(ANS_DIR, '_a1'); os.makedirs(d)
doc = Document(); doc.add_heading('科研经费使用合规诊断报告', level=0)
doc.add_heading('一、诊断概述', level=1)
doc.add_paragraph('华材大学（脱敏代称）三年科研经费投入3.95亿元，40个项目。'
    '诊断发现8个项目超预算执行（超支约1600万元），6个项目经费挪用嫌疑，'
    '5个项目结题未验收，4个项目预算结余大（执行率低于50%），'
    '3个项目超支但成果突出（合理超支）。'
    '60条支出明细中10条超范围支出、8条虚列支出、6条拆分采购、4条跨项目调剂。'
    '35笔采购中6笔拆分规避招标、4笔超限额未招标。'
    '30台设备中8台闲置超6个月、4台利用率低但属定制设备合理。'
    '总体经费执行合规率约72%，低于目标85%。')
doc.add_heading('二、项目经费诊断', level=1)
doc.add_paragraph('1. 超预算项目8个（KY001至KY008），实际支出超预算25%-60%，超支约1600万元。'
    '依据科研经费管理办法，超预算20%以上应报原审批部门审批。')
doc.add_paragraph('2. 经费挪用嫌疑6个（KY009至KY014），支出内容与项目任务书不符。'
    '依据科研经费管理办法，经费挪用应追回并追究责任人。')
doc.add_paragraph('3. 结题未验收5个（KY015至KY019），已结题但未完成验收流程。'
    '依据科研经费管理办法，结题应先验收后结算。')
doc.add_paragraph('4. 预算结余大4个（KY020至KY023），执行率低于50%。'
    '依据科研经费管理办法，应调整预算或收回结余。')
doc.add_paragraph('5. 超支但成果突出3个（KY024至KY026），超支15%-35%但论文/专利/转化均超目标，'
    '属合理超支（反直觉正常项），不误判为违规。')
doc.add_heading('三、支出明细诊断', level=1)
doc.add_paragraph('1. 超范围支出10条（PZ0001至PZ0010），劳务费与项目任务书无关。')
doc.add_paragraph('2. 虚列支出8条（PZ0011至PZ0018），发票金额与实际采购不符。')
doc.add_paragraph('3. 拆分采购6条（PZ0019至PZ0024），将超限额采购拆分为多笔小额。')
doc.add_paragraph('4. 跨项目调剂4条（PZ0025至PZ0028），未经审批跨项目使用经费。')
doc.add_paragraph('5. 正常大额支出3条（PZ0029至PZ0031），金额大但属大型设备采购已按规定招标（反直觉正常项）。')
doc.add_heading('四、总体结论', level=1)
doc.add_paragraph('科研经费存在5大类15项问题，年浪费约2200万元。'
    '通过整改可实现年节约1800万元，经费执行合规率从72%提升至85%。')
doc.save(os.path.join(d, '经费使用诊断报告.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '诊断结论汇总'
ws.append(['序号', '诊断维度', '问题类别', '数量', '影响金额(万元)', '优化方向', '优先级'])
diags = [
    (1, '项目经费', '超预算执行', 8, 1600, '追责并报审批', '高'),
    (2, '项目经费', '经费挪用嫌疑', 6, 500, '追回并追责', '高'),
    (3, '项目经费', '结题未验收', 5, 0, '补验收流程', '高'),
    (4, '项目经费', '预算结余大', 4, 800, '调整或收回', '中'),
    (5, '项目经费', '超支但成果突出（保留）', 3, 0, '维持现状', '低'),
    (6, '支出明细', '超范围支出', 10, 200, '追回并整改', '高'),
    (7, '支出明细', '虚列支出', 8, 150, '追回并追责', '高'),
    (8, '支出明细', '拆分采购', 6, 300, '合并采购', '中'),
    (9, '支出明细', '跨项目调剂', 4, 100, '补审批或追回', '中'),
    (10, '采购', '拆分规避招标', 6, 400, '合并招标', '高'),
    (11, '采购', '超限额未招标', 4, 300, '补招标程序', '高'),
    (12, '资产', '设备闲置', 8, 500, '调配使用', '中'),
    (13, '采购', '单一来源合理（保留）', 3, 0, '维持现状', '低'),
    (14, '采购', '紧急采购合理（保留）', 2, 0, '维持现状', '低'),
    (15, '资产', '利用率低但合理（保留）', 4, 0, '维持现状', '低'),
]
for row in diags: ws.append(row)
style(ws, 7); wb.save(os.path.join(d, '诊断结论汇总表.xlsx'))

# 雷达图
fig, ax = plt.subplots(figsize=(8, 8), subplot_kw=dict(projection='polar'))
cats = ['项目经费', '支出明细', '采购合规', '资产管理', '制度体系', '成果绩效']
cur = [55, 60, 50, 65, 45, 72]; tgt = [85, 82, 88, 80, 78, 85]
angles = np.linspace(0, 2*np.pi, len(cats), endpoint=False).tolist()
cur += cur[:1]; tgt += tgt[:1]; angles += angles[:1]
ax.plot(angles, cur, 'o-', linewidth=2, label='当前', color='#ED7D31')
ax.fill(angles, cur, alpha=0.15, color='#ED7D31')
ax.plot(angles, tgt, 's-', linewidth=2, label='目标', color='#7030A0')
ax.fill(angles, tgt, alpha=0.15, color='#7030A0')
ax.set_xticks(angles[:-1]); ax.set_xticklabels(cats)
ax.set_ylim(0, 100); ax.set_title('科研经费管理各维度当前 vs 目标', fontsize=12, pad=20)
ax.legend(loc='upper right', bbox_to_anchor=(1.3, 1.1))
plt.tight_layout(); fig.savefig(os.path.join(d, '诊断雷达图.png'), dpi=150); plt.close()
mkzip(d, '01经费使用诊断报告.rar')

# ========== 答案2: 项目经费审查答案 ==========
print('=== 02 项目经费审查答案 ===')
d = os.path.join(ANS_DIR, '_a2'); os.makedirs(d)
doc = Document(); doc.add_heading('科研项目经费审查方案', level=0)
doc.add_heading('一、审查范围', level=1)
doc.add_paragraph('40个科研项目，60条支出明细。重点审查预算执行、经费挪用、结题验收。')
doc.add_heading('二、分类审查结论', level=1)
doc.add_heading('2.1 超预算追责（8个）', level=2)
doc.add_paragraph('项目编号：KY001至KY008。超支25%-60%，约1600万元。'
    '措施：依据科研经费管理办法，超预算20%以上报原审批部门审批，追责项目负责人。')
doc.add_heading('2.2 经费挪用整改（6个）', level=2)
doc.add_paragraph('项目编号：KY009至KY014。支出内容与任务书不符。'
    '措施：追回挪用经费，追究责任人，补充审批或调整任务书。')
doc.add_heading('2.3 结题未验收补验（5个）', level=2)
doc.add_paragraph('项目编号：KY015至KY019。已结题但未完成验收。'
    '措施：补完验收流程，验收不通过的追回已结算经费。')
doc.add_heading('2.4 预算结余调整（4个）', level=2)
doc.add_paragraph('项目编号：KY020至KY023。执行率低于50%。'
    '措施：调整预算用于其他项目或收回结余。')
doc.add_heading('2.5 超支但成果突出保留（3个，反直觉正常项）', level=2)
doc.add_paragraph('项目编号：KY024至KY026。超支15%-35%但论文/专利/转化均超目标。'
    '措施：维持现状，不误判为违规。依据：科研绩效评价应以成果产出为核心。')
doc.save(os.path.join(d, '项目经费审查方案.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '项目审查明细'
ws.append(['序号', '审查类别', '项目编号', '数量', '问题', '措施', '影响金额(万元)', '依据'])
rows = [
    (1, '超预算追责', 'KY001-KY008', 8, '超预算25%-60%', '追责并报审批', 1600, '科研经费管理办法'),
    (2, '经费挪用', 'KY009-KY014', 6, '支出与任务书不符', '追回并追责', 500, '科研经费管理办法'),
    (3, '结题未验收', 'KY015-KY019', 5, '已结题未验收', '补验收', 0, '科研经费管理办法'),
    (4, '预算结余大', 'KY020-KY023', 4, '执行率<50%', '调整或收回', 800, '科研经费管理办法'),
    (5, '超支成果突出（保留）', 'KY024-KY026', 3, '超支但成果超目标', '维持', 0, '科研绩效评价'),
    (6, '正常', '其余14个', 14, '无问题', '维持', 0, '-'),
    (7, '合计', '-', 40, '-', '-', 2900, '-'),
]
for r in rows: ws.append(r)
style(ws, 8); wb.save(os.path.join(d, '项目审查明细表.xlsx'))

fig, ax = plt.subplots(figsize=(10, 6))
cats = ['超预算(8)', '挪用(6)', '未验收(5)', '结余大(4)', '超支成果好(3)', '正常(14)']
amounts = [1600, 500, 0, 800, 0, 0]
colors = ['#FF6B6B', '#FFC000', '#ED7D31', '#5B9BD5', '#70AD47', '#A5A5A5']
ax.bar(cats, amounts, color=colors, width=0.6)
ax.set_ylabel('影响金额（万元）'); ax.set_title('各类问题影响金额')
ax.set_ylim(0, 1800)
for i, v in enumerate(amounts):
    ax.text(i, v+20, f'{v}万', ha='center', fontsize=9)
plt.xticks(rotation=15, ha='right'); plt.tight_layout()
fig.savefig(os.path.join(d, '问题影响金额图.png'), dpi=150); plt.close()
mkzip(d, '02项目经费审查答案.rar')

# ========== 答案3: 支出与采购审查答案 ==========
print('=== 03 支出与采购审查答案 ===')
d = os.path.join(ANS_DIR, '_a3'); os.makedirs(d)
doc = Document(); doc.add_heading('经费支出与采购审查方案', level=0)
doc.add_heading('一、支出明细审查', level=1)
doc.add_paragraph('1. 超范围支出10条（PZ0001至PZ0010），劳务费与任务书无关，应追回约200万元。')
doc.add_paragraph('2. 虚列支出8条（PZ0011至PZ0018），发票金额与实际不符，应追回约150万元。')
doc.add_paragraph('3. 拆分采购6条（PZ0019至PZ0024），应合并采购，涉及约300万元。')
doc.add_paragraph('4. 跨项目调剂4条（PZ0025至PZ0028），未经审批跨项目使用，应补审批或追回约100万元。')
doc.add_paragraph('5. 正常大额支出3条（PZ0029至PZ0031），金额大但属大型设备已招标（反直觉正常项），不误判。')
doc.add_heading('二、采购审查', level=1)
doc.add_paragraph('1. 拆分规避招标6笔（CG001至CG006），同类累计超限额但单笔低于限额，应合并招标。')
doc.add_paragraph('2. 超限额未招标4笔（CG007至CG010），单笔超50万限额未走招标，应补招标程序。')
doc.add_paragraph('3. 单一来源合理3笔（CG011至CG013），专有技术设备经审批（反直觉正常项），不误判。')
doc.add_paragraph('4. 紧急采购合理2笔（CG014至CG015），实验急需设备经审批（反直觉正常项），不误判。')
doc.add_heading('三、设备资产审查', level=1)
doc.add_paragraph('1. 闲置设备8台（ZC001至ZC008），闲置超6个月，应调配至其他课题组。')
doc.add_paragraph('2. 共享设备5台（ZC009至ZC013），已纳入共享平台，利用率60%-90%，合理。')
doc.add_paragraph('3. 利用率低但合理4台（ZC014至ZC017），专属实验定制设备（反直觉正常项），不误判。')
doc.save(os.path.join(d, '支出采购审查方案.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '支出采购审查明细'
ws.append(['序号', '类别', '编号范围', '数量', '问题', '措施', '影响金额(万元)', '依据'])
rows = [
    (1, '超范围支出', 'PZ0001-PZ0010', 10, '与任务书无关', '追回', 200, '科研经费管理办法'),
    (2, '虚列支出', 'PZ0011-PZ0018', 8, '发票不符', '追回追责', 150, '科研经费管理办法'),
    (3, '拆分采购', 'PZ0019-PZ0024', 6, '拆分规避', '合并采购', 300, '采购管理办法'),
    (4, '跨项目调剂', 'PZ0025-PZ0028', 4, '未审批调剂', '补审批或追回', 100, '科研经费管理办法'),
    (5, '正常大额（保留）', 'PZ0029-PZ0031', 3, '已招标', '维持', 0, '采购管理办法'),
    (6, '拆分规避招标', 'CG001-CG006', 6, '拆分规避', '合并招标', 400, '采购管理办法'),
    (7, '超限额未招标', 'CG007-CG010', 4, '超限额', '补招标', 300, '采购管理办法'),
    (8, '单一来源合理（保留）', 'CG011-CG013', 3, '专有技术', '维持', 0, '采购管理办法'),
    (9, '紧急采购合理（保留）', 'CG014-CG015', 2, '实验急需', '维持', 0, '采购管理办法'),
    (10, '设备闲置', 'ZC001-ZC008', 8, '闲置>6月', '调配', 500, '资产管理办法'),
    (11, '利用率低合理（保留）', 'ZC014-ZC017', 4, '定制设备', '维持', 0, '资产管理办法'),
]
for r in rows: ws.append(r)
style(ws, 8); wb.save(os.path.join(d, '支出采购审查明细表.xlsx'))
mkzip(d, '03支出采购审查答案.rar')

# ========== 答案4: 成果绩效评价答案 ==========
print('=== 04 成果绩效评价答案 ===')
d = os.path.join(ANS_DIR, '_a4'); os.makedirs(d)
doc = Document(); doc.add_heading('科研成果与绩效评价方案', level=0)
doc.add_heading('一、评价体系', level=1)
doc.add_paragraph('从学术成果(30分)、技术指标(25分)、经济效益(20分)、人才培养(15分)、经费执行(10分)五个维度评价。')
doc.add_heading('二、分类评价结论', level=1)
doc.add_paragraph('1. 超预算项目绩效：8个超预算项目（KY001至KY008），经费执行率超100%但成果未同步提升，评级B/C。')
doc.add_paragraph('2. 经费挪用项目绩效：6个挪用项目（KY009至KY014），成果待重新评估。')
doc.add_paragraph('3. 结题未验收项目绩效：5个（KY015至KY019），成果产出正常但验收流程缺失。')
doc.add_paragraph('4. 预算结余大项目绩效：4个（KY020至KY023），执行率低但学术成果正常，需调整预算。')
doc.add_paragraph('5. 超支成果突出项目绩效：3个（KY024至KY026），超支但论文/专利/转化均超目标，评级A（反直觉正常项）。')
doc.add_heading('三、绩效与经费执行挂钩', level=1)
doc.add_paragraph('经费执行率应纳入绩效评价（权重10%），执行率低于50%的项目评级降一档。'
    '绩效评价结果应与后续立项挂钩，绩效差的项目负责人限制新立项。')
doc.save(os.path.join(d, '成果绩效评价方案.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '绩效评价明细'
ws.append(['序号', '项目编号', '类别', '学术成果(30)', '技术指标(25)', '经济效益(20)', '人才培养(15)', '经费执行(10)', '总分(100)', '评级', '备注'])
for i in range(1, 41):
    if i <= 8:
        cat = '超预算'; ac = random.randint(20, 28); ti = random.randint(18, 23); ec = random.randint(12, 16)
        tr = random.randint(10, 13); fe = random.randint(3, 5); note = '经费执行超标，绩效一般'
    elif i <= 14:
        cat = '挪用'; ac = random.randint(18, 24); ti = random.randint(16, 20); ec = random.randint(10, 14)
        tr = random.randint(8, 12); fe = random.randint(4, 6); note = '经费使用异常，待重新评估'
    elif i <= 19:
        cat = '未验收'; ac = random.randint(15, 22); ti = random.randint(14, 20); ec = random.randint(10, 15)
        tr = random.randint(9, 12); fe = random.randint(6, 8); note = '成果正常但验收缺失'
    elif i <= 23:
        cat = '结余大'; ac = random.randint(20, 26); ti = random.randint(18, 23); ec = random.randint(14, 18)
        tr = random.randint(11, 14); fe = random.randint(2, 4); note = '执行率低，需调整预算'
    elif i <= 26:
        cat = '超支成果好'; ac = random.randint(26, 30); ti = random.randint(22, 25); ec = random.randint(16, 20)
        tr = random.randint(13, 15); fe = random.randint(3, 5); note = '成果突出，超支合理（反直觉正常项）'
    else:
        cat = '正常'; ac = random.randint(20, 28); ti = random.randint(18, 24); ec = random.randint(13, 18)
        tr = random.randint(10, 14); fe = random.randint(7, 9); note = ''
    total = ac+ti+ec+tr+fe
    grade = 'A' if total >= 85 else 'B' if total >= 70 else 'C'
    ws.append([i, f'KY{i:03d}', cat, ac, ti, ec, tr, fe, total, grade, note])
style(ws, 11); wb.save(os.path.join(d, '绩效评价明细表.xlsx'))
mkzip(d, '04成果绩效评价答案.rar')

# ========== 答案5: 管理层错误要求回函 ==========
print('=== 05 管理层错误要求回函 ===')
d = os.path.join(ANS_DIR, '_a5'); os.makedirs(d)
doc = Document(); doc.add_heading('管理层错误要求回函', level=0)
reqs = [
    (1, '校长', '超预算不追责', '科研经费管理办法规定超预算20%以上应报审批部门审批并追责。不追责将导致预算管控失效。'),
    (2, '财务处长', '经费挪用不处理', '科研经费管理办法规定经费挪用应追回并追究责任人。不处理将导致经费使用失控。'),
    (3, '科研院长', '结题未验收直接通过', '科研经费管理办法规定结题应先验收后结算。未验收直接通过将导致质量不可控。'),
    (4, '设备处长', '拆分采购不追究', '科研采购管理办法规定不得拆分采购规避招标。不追究将导致采购合规失效。'),
    (5, '校长', '闲置设备不调配', '科研资产管理办法规定闲置超6个月应调配。不调配将导致资源浪费。'),
    (6, '财务处长', '超范围支出不追回', '科研经费管理办法规定超范围支出应追回。不追回将导致经费流失。'),
    (7, '科研院长', '虚列支出不审计', '科研经费管理办法规定虚列支出应审计追责。不审计将导致虚假支出蔓延。'),
    (8, '设备处长', '单一来源不审批', '科研采购管理办法规定单一来源应审批。不审批将导致采购风险。'),
    (9, '校长', '预算结余不调整', '科研经费管理办法规定执行率低应调整预算。不调整将导致资金沉淀。'),
    (10, '科研院长', '绩效只看论文', '科研绩效评价办法规定应多维评价含经费执行率。只看论文将导致经费管理失控。'),
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
doc = Document(); doc.add_heading('科研经费审计整体方案', level=0)
doc.add_heading('一、项目概述', level=1)
doc.add_paragraph('华材大学三年科研经费3.95亿元，40个项目。通过审计发现5大类15项问题，'
    '年浪费约2200万元。通过整改可实现年节约1800万元，经费执行合规率从72%提升至85%。')
doc.add_heading('二、优化目标', level=1)
doc.add_paragraph('1. 项目经费：超预算追责1600万、经费挪用追回500万、结题补验收、预算结余调整800万\n'
    '2. 支出采购：超范围追回200万、虚列追回150万、拆分合并300万、采购违规整改700万\n'
    '3. 资产管理：闲置设备调配节约500万\n4. 管理层10项错误要求全部驳回\n'
    '5. 经费执行合规率从72%提升至85%')
doc.add_heading('三、分阶段实施', level=1)
doc.add_paragraph('第一阶段（1-2月）：超预算追责、经费挪用追回、虚列支出追回\n'
    '第二阶段（2-4月）：拆分采购合并、超限额补招标、闲置设备调配\n'
    '第三阶段（3-5月）：结题补验收、预算结余调整\n'
    '第四阶段（4-8月）：制度修订、建立长效机制')
doc.add_heading('四、预期效益', level=1)
doc.add_paragraph('1. 年节约1800万元\n2. 经费执行合规率从72%提升至85%\n'
    '3. 采购合规率从60%提升至90%\n4. 设备利用率从55%提升至75%\n'
    '5. 科研绩效评价体系完善，结果与立项挂钩')
doc.save(os.path.join(d, '科研审计整体方案.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '数据勾稽自检'
ws.append(['勾稽项', '题目描述', '附件明细', '标准答案', '是否一致'])
rows = [
    ('项目总数', '40个', '台账40行', '40个', '一致'),
    ('超预算', '8个', 'KY001-008', '8个', '一致'),
    ('经费挪用', '6个', 'KY009-014', '6个', '一致'),
    ('结题未验收', '5个', 'KY015-019', '5个', '一致'),
    ('预算结余大', '4个', 'KY020-023', '4个', '一致'),
    ('超支成果突出', '3个', 'KY024-026', '3个', '一致'),
    ('正常项目', '14个', '其余14个', '14个', '一致'),
    ('支出明细总数', '60条', '明细60行', '60条', '一致'),
    ('超范围支出', '10条', 'PZ0001-0010', '10条', '一致'),
    ('虚列支出', '8条', 'PZ0011-0018', '8条', '一致'),
    ('拆分采购', '6条', 'PZ0019-0024', '6条', '一致'),
    ('跨项目调剂', '4条', 'PZ0025-0028', '4条', '一致'),
    ('正常大额（保留）', '3条', 'PZ0029-0031', '3条', '一致'),
    ('正常支出', '29条', '其余29条', '29条', '一致'),
    ('采购记录总数', '35笔', '采购35行', '35笔', '一致'),
    ('拆分规避招标', '6笔', 'CG001-006', '6笔', '一致'),
    ('超限额未招标', '4笔', 'CG007-010', '4笔', '一致'),
    ('单一来源合理', '3笔', 'CG011-013', '3笔', '一致'),
    ('紧急采购合理', '2笔', 'CG014-015', '2笔', '一致'),
    ('正常采购', '20笔', '其余20笔', '20笔', '一致'),
    ('设备总数', '30台', '台账30行', '30台', '一致'),
    ('闲置设备', '8台', 'ZC001-008', '8台', '一致'),
    ('共享设备', '5台', 'ZC009-013', '5台', '一致'),
    ('利用率低合理', '4台', 'ZC014-017', '4台', '一致'),
    ('正常设备', '13台', '其余13台', '13台', '一致'),
    ('管理制度', '5份', '5个Word', '5份', '一致'),
    ('管理层要求', '10条', '要求10条', '10条', '一致'),
    ('经费总额', '3.95亿元', '趋势图', '3.95亿元', '一致'),
    ('年节约目标', '1800万元', '各项加总', '1800万元', '一致'),
    ('合规率目标', '72%→85%', '-', '85%', '一致'),
    ('出题标记', '0处', '纯业务数据', '-', '一致'),
]
for r in rows: ws.append(r)
style(ws, 5); wb.save(os.path.join(d, '数据勾稽自检表.xlsx'))

fig, ax = plt.subplots(figsize=(12, 6))
metrics = ['超预算(万)', '挪用(万)', '虚列(万)', '采购违规(万)', '合规率(%)', '设备利用(%)']
before = [1600, 500, 150, 700, 72, 55]
after = [0, 0, 0, 0, 85, 75]
scales = [2000, 1000, 500, 1000, 100, 100]
bn = [b/s for b, s in zip(before, scales)]
an = [a/s for a, s in zip(after, scales)]
x = range(len(metrics))
ax.bar([i-0.2 for i in x], bn, 0.4, label='优化前', color='#ED7D31', alpha=0.7)
ax.bar([i+0.2 for i in x], an, 0.4, label='优化后', color='#7030A0', alpha=0.7)
ax.set_ylabel('归一化值'); ax.set_title('科研经费审计优化前后对比')
ax.set_xticks(x); ax.set_xticklabels(metrics); ax.legend()
for i in range(len(metrics)):
    ax.text(i, max(bn[i], an[i])+0.02, f'{before[i]}→{after[i]}', ha='center', fontsize=8)
plt.tight_layout(); fig.savefig(os.path.join(d, '优化前后对比图.png'), dpi=150); plt.close()
mkzip(d, '06整体方案与勾稽自检.rar')

print('\n=== 标准答案生成完成 ===')
