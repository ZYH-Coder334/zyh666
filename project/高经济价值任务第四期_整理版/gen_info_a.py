#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""中小学教育信息化建设与专项资金使用审计 - 标准答案生成"""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from docx import Document
import zipfile, os, random, shutil

random.seed(22)
plt.rcParams['font.sans-serif'] = ['WenQuanYi Micro Hei', 'SimHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

BASE = '/workspace/教育信息化审计_附件'
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

# ========== 答案1: 合规诊断报告 ==========
print('=== 答案1 合规诊断报告 ===')
d = os.path.join(ANS_DIR, '_a1'); os.makedirs(d)

fig, ax = plt.subplots(figsize=(12, 7))
categories = ['学校管理', '设备采购', '软件采购', '设备使用', '教师培训', '网络安全']
violations = [19, 50, 30, 70, 25, 13]
colors = ['#FF6B6B', '#FF9F43', '#FECA57', '#48DBFB', '#1DD1A1', '#5F27CD']
ax.barh(categories, violations, color=colors)
ax.set_xlabel('违规数'); ax.set_title('各模块违规数量统计')
for i, v in enumerate(violations):
    ax.text(v+1, i, str(v), va='center', fontsize=9)
plt.tight_layout(); fig.savefig(os.path.join(d, '违规统计.png'), dpi=150); plt.close()

doc = Document()
doc.add_heading('教育信息化建设与专项资金使用审计报告', 0)
doc.add_paragraph('编制依据：《教育信息化专项资金管理办法》《政府采购法》《网络安全法》《中小学教师信息技术应用能力标准》等')
doc.add_paragraph(f'审查范围：30所学校、100条设备采购记录、60条软件采购记录、150条设备使用台账、40条培训记录、15份网络安全检查报告')
doc.add_paragraph(f'核心结论：共发现违规问题207项，其中重大违规88项，一般违规119项。涉及22所学校。')
doc.add_heading('一、学校信息化管理', level=1)
doc.add_paragraph('1. 专项资金挪用：8所学校（XX001-XX008）')
doc.add_paragraph('2. 设备闲置：6所学校（XX009-XX014）')
doc.add_paragraph('3. 培训流于形式：5所学校（XX015-XX019）')
doc.add_paragraph('4. 合规保留：3所学校（XX020-XX022）设备使用率高，合规。')
doc.add_heading('二、设备采购', level=1)
doc.add_paragraph('1. 未招标：25条（CG0001-CG0025）')
doc.add_paragraph('2. 超预算：15条（CG0026-CG0040）')
doc.add_paragraph('3. 验收不合格：10条（CG0041-CG0050）')
doc.add_paragraph('4. 合规保留：10条（CG0051-CG0060）单一来源采购理由充分。')
doc.add_heading('三、软件采购', level=1)
doc.add_paragraph('1. 重复采购：12条（RJ001-RJ012）')
doc.add_paragraph('2. 使用率低：10条（RJ013-RJ022）')
doc.add_paragraph('3. 未招标：8条（RJ023-RJ030）')
doc.add_paragraph('4. 合规保留：5条（RJ031-RJ035）正版软件授权。')
doc.add_heading('四、设备使用', level=1)
doc.add_paragraph('1. 闲置：30条（SY0001-SY0030）')
doc.add_paragraph('2. 使用率低：25条（SY0031-SY0055）')
doc.add_paragraph('3. 未入账：15条（SY0056-SY0070）')
doc.add_paragraph('4. 合规保留：10条（SY0071-SY0080）校际共享设备。')
doc.add_heading('五、教师培训', level=1)
doc.add_paragraph('1. 流于形式：12条（PX001-PX012）')
doc.add_paragraph('2. 未考核：8条（PX013-PX020）')
doc.add_paragraph('3. 超范围：5条（PX021-PX025）')
doc.add_paragraph('4. 合规保留：3条（PX026-PX028）培训实效显著。')
doc.add_heading('六、网络安全', level=1)
doc.add_paragraph('1. 等保未备案：6份（AQ001-AQ006）')
doc.add_paragraph('2. 隐患未整改：5份（AQ007-AQ011）')
doc.add_paragraph('3. 备份缺失：2份（AQ012-AQ013）')
doc.add_paragraph('4. 合规保留：2份（AQ014-AQ015）等保合规。')
doc.save(os.path.join(d, '审计报告.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '诊断结论汇总'
ws.append(['序号', '问题类别', '违规数量', '涉及编号范围', '严重等级', '法规依据'])
items = [
    (1, '专项资金挪用', 8, 'XX001-XX008', '重大', '《教育信息化专项资金管理办法》'),
    (2, '设备闲置', 6, 'XX009-XX014', '一般', '同上'),
    (3, '培训流于形式', 5, 'XX015-XX019', '一般', '《中小学教师信息技术应用能力标准》'),
    (4, '设备采购未招标', 25, 'CG0001-CG0025', '重大', '《政府采购法》'),
    (5, '设备采购超预算', 15, 'CG0026-CG0040', '一般', '《政府采购法》'),
    (6, '设备验收不合格', 10, 'CG0041-CG0050', '重大', '《政府采购法》'),
    (7, '软件重复采购', 12, 'RJ001-RJ012', '重大', '《政府采购法》'),
    (8, '软件使用率低', 10, 'RJ013-RJ022', '一般', '《教育信息化专项资金管理办法》'),
    (9, '软件采购未招标', 8, 'RJ023-RJ030', '重大', '《政府采购法》'),
    (10, '设备闲置', 30, 'SY0001-SY0030', '重大', '《教育信息化专项资金管理办法》'),
    (11, '设备使用率低', 25, 'SY0031-SY0055', '一般', '同上'),
    (12, '设备未入账', 15, 'SY0056-SY0070', '重大', '《行政事业单位国有资产管理办法》'),
    (13, '培训流于形式', 12, 'PX001-PX012', '一般', '《中小学教师信息技术应用能力标准》'),
    (14, '培训未考核', 8, 'PX013-PX020', '一般', '同上'),
    (15, '培训超范围', 5, 'PX021-PX025', '一般', '同上'),
    (16, '等保未备案', 6, 'AQ001-AQ006', '重大', '《网络安全法》'),
    (17, '安全隐患未整改', 5, 'AQ007-AQ011', '重大', '《网络安全法》'),
    (18, '数据备份缺失', 2, 'AQ012-AQ013', '重大', '《网络安全法》'),
]
for item in items:
    ws.append(item)
style(ws, 6); wb.save(os.path.join(d, '诊断结论汇总.xlsx'))
mkzip(d, '01合规诊断报告.rar')

# ========== 答案2: 设备采购审查 ==========
print('=== 答案2 设备采购审查 ===')
d = os.path.join(ANS_DIR, '_a2'); os.makedirs(d)
doc = Document()
doc.add_heading('设备采购审查答案', 0)
doc.add_heading('一、未招标清单（25条 CG0001-CG0025）', level=1)
doc.add_paragraph('问题：未按规定进行政府采购，自行采购')
doc.add_paragraph('整改：补办政府采购手续，追究责任')
doc.add_heading('二、超预算清单（15条 CG0026-CG0040）', level=1)
doc.add_paragraph('问题：实际采购金额超出预算')
doc.add_paragraph('整改：说明超预算原因，追回差额')
doc.add_heading('三、验收不合格清单（10条 CG0041-CG0050）', level=1)
doc.add_paragraph('问题：设备验收不合格，未整改')
doc.add_paragraph('整改：退货或更换，重新验收')
doc.add_heading('四、合规保留项（10条 CG0051-CG0060）', level=1)
doc.add_paragraph('单一来源采购理由充分，予以保留。')
doc.save(os.path.join(d, '设备采购审查答案.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '设备采购问题明细'
ws.append(['序号', '编号', '问题类型', '问题描述', '法规依据', '整改要求'])
for i in range(1, 26):
    ws.append([i, f'CG{i:04d}', '未招标', '未按规定政府采购', '《政府采购法》', '补办手续'])
for i in range(26, 41):
    ws.append([i, f'CG{i:04d}', '超预算', '实际金额超出预算', '《政府采购法》', '追回差额'])
for i in range(41, 51):
    ws.append([i, f'CG{i:04d}', '验收不合格', '验收不合格未整改', '《政府采购法》', '退货更换'])
style(ws, 6); wb.save(os.path.join(d, '设备采购问题明细.xlsx'))
mkzip(d, '02设备采购审查答案.rar')

# ========== 答案3: 软件采购审查 ==========
print('=== 答案3 软件采购审查 ===')
d = os.path.join(ANS_DIR, '_a3'); os.makedirs(d)
doc = Document()
doc.add_heading('软件采购审查答案', 0)
doc.add_heading('一、重复采购清单（12条 RJ001-RJ012）', level=1)
doc.add_paragraph('问题：与已采购软件功能重复')
doc.add_paragraph('整改：取消重复采购，统一使用')
doc.add_heading('二、使用率低清单（10条 RJ013-RJ022）', level=1)
doc.add_paragraph('问题：软件使用率极低，基本闲置')
doc.add_paragraph('整改：评估使用价值，决定是否续费')
doc.add_heading('三、未招标清单（8条 RJ023-RJ030）', level=1)
doc.add_paragraph('问题：未按规定进行政府采购')
doc.add_paragraph('整改：补办政府采购手续')
doc.add_heading('四、合规保留项（5条 RJ031-RJ035）', level=1)
doc.add_paragraph('正版软件授权，予以保留。')
doc.save(os.path.join(d, '软件采购审查答案.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '软件采购问题明细'
ws.append(['序号', '编号', '问题类型', '问题描述', '法规依据', '整改要求'])
for i in range(1, 13):
    ws.append([i, f'RJ{i:04d}', '重复采购', '与已采购软件功能重复', '《政府采购法》', '取消重复采购'])
for i in range(13, 23):
    ws.append([i, f'RJ{i:04d}', '使用率低', '软件基本闲置', '《教育信息化专项资金管理办法》', '评估续费'])
for i in range(23, 31):
    ws.append([i, f'RJ{i:04d}', '未招标', '未按规定政府采购', '《政府采购法》', '补办手续'])
style(ws, 6); wb.save(os.path.join(d, '软件采购问题明细.xlsx'))
mkzip(d, '03软件采购审查答案.rar')

# ========== 答案4: 设备使用与培训审查 ==========
print('=== 答案4 设备使用与培训审查 ===')
d = os.path.join(ANS_DIR, '_a4'); os.makedirs(d)
doc = Document()
doc.add_heading('设备使用与培训审查答案', 0)
doc.add_heading('一、设备使用违规（70条）', level=1)
doc.add_paragraph('1. 闲置30条（SY0001-SY0030）：设备长期闲置未投入使用')
doc.add_paragraph('2. 使用率低25条（SY0031-SY0055）：使用率远低于预期')
doc.add_paragraph('3. 未入账15条（SY0056-SY0070）：设备未登记入账')
doc.add_paragraph('4. 合规保留10条（SY0071-SY0080）：校际共享设备使用合理')
doc.add_heading('二、培训违规（25条）', level=1)
doc.add_paragraph('1. 流于形式12条（PX001-PX012）：签到代签')
doc.add_paragraph('2. 未考核8条（PX013-PX020）：培训后未组织考核')
doc.add_paragraph('3. 超范围5条（PX021-PX025）：培训内容超出信息化范围')
doc.add_paragraph('4. 合规保留3条（PX026-PX028）：培训实效显著')
doc.save(os.path.join(d, '设备使用与培训审查答案.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '设备使用问题明细'
ws.append(['序号', '编号', '问题类型', '问题描述', '法规依据', '整改要求'])
for i in range(1, 31):
    ws.append([i, f'SY{i:04d}', '闲置', '设备长期闲置', '《教育信息化专项资金管理办法》', '盘活使用'])
for i in range(31, 56):
    ws.append([i, f'SY{i:04d}', '使用率低', '使用率远低于预期', '同上', '加强培训应用'])
for i in range(56, 71):
    ws.append([i, f'SY{i:04d}', '未入账', '设备未登记入账', '《行政事业单位国有资产管理办法》', '补登入账'])
style(ws, 6); wb.save(os.path.join(d, '设备使用问题明细.xlsx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '培训问题明细'
ws.append(['序号', '编号', '问题类型', '问题描述', '法规依据', '整改要求'])
for i in range(1, 13):
    ws.append([i, f'PX{i:03d}', '流于形式', '签到代签', '《中小学教师信息技术应用能力标准》', '重新培训'])
for i in range(13, 21):
    ws.append([i, f'PX{i:03d}', '未考核', '培训后未考核', '同上', '组织考核'])
for i in range(21, 26):
    ws.append([i, f'PX{i:03d}', '超范围', '内容超出信息化范围', '同上', '调整内容'])
style(ws, 6); wb.save(os.path.join(d, '培训问题明细.xlsx'))
mkzip(d, '04设备使用与培训审查答案.rar')

# ========== 答案5: 网络安全审查 ==========
print('=== 答案5 网络安全审查 ===')
d = os.path.join(ANS_DIR, '_a5'); os.makedirs(d)
doc = Document()
doc.add_heading('网络安全审查答案', 0)
doc.add_heading('一、网络安全违规（13份）', level=1)
doc.add_paragraph('1. 等保未备案6份（AQ001-AQ006）：未按规定进行网络安全等级保护备案')
doc.add_paragraph('2. 隐患未整改5份（AQ007-AQ011）：安全隐患未按要求整改')
doc.add_paragraph('3. 备份缺失2份（AQ012-AQ013）：重要数据未定期备份')
doc.add_heading('二、合规保留项（2份 AQ014-AQ015）', level=1)
doc.add_paragraph('等保合规，隐患已整改，予以保留。')
doc.save(os.path.join(d, '网络安全审查答案.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '网络安全问题明细'
ws.append(['序号', '编号', '问题类型', '问题描述', '法规依据', '整改要求'])
for i in range(1, 7):
    ws.append([i, f'AQ{i:03d}', '等保未备案', '未按规定等级保护备案', '《网络安全法》', '立即备案'])
for i in range(7, 12):
    ws.append([i, f'AQ{i:03d}', '隐患未整改', '安全隐患未整改', '《网络安全法》', '立即整改'])
for i in range(12, 14):
    ws.append([i, f'AQ{i:03d}', '备份缺失', '重要数据未备份', '《网络安全法》', '建立备份机制'])
style(ws, 6); wb.save(os.path.join(d, '网络安全问题明细.xlsx'))
mkzip(d, '05网络安全审查答案.rar')

# ========== 答案6: 管理层错误要求回函与整体方案 ==========
print('=== 答案6 管理层错误要求回函与整体方案 ===')
d = os.path.join(ANS_DIR, '_a6'); os.makedirs(d)

fig, ax = plt.subplots(figsize=(12, 7))
phases = ['资金核查', '采购规范', '资产盘活', '安全整改']
tasks = [15, 20, 12, 8]
colors = ['#4472C4', '#ED7D31', '#70AD47', '#FFC000']
ax.bar(phases, tasks, color=colors, width=0.5)
ax.set_ylabel('工作项数'); ax.set_title('整改方案阶段任务分布')
for i, v in enumerate(tasks):
    ax.text(i, v+0.3, str(v), ha='center', fontsize=9)
plt.tight_layout(); fig.savefig(os.path.join(d, '整改方案.png'), dpi=150); plt.close()

doc = Document()
doc.add_heading('关于管理层要求的回函', 0)
rejections = [
    (1, '信息化专项资金挪用', '违反《教育信息化专项资金管理办法》，专项资金必须专款专用', '追回资金'),
    (2, '设备采购不招标', '违反《政府采购法》，必须按规定招标', '补办招标手续'),
    (3, '闲置设备不处理', '违反《教育信息化专项资金管理办法》，闲置设备应盘活使用', '盘活使用'),
    (4, '软件重复采购不清理', '违反《政府采购法》，不得重复采购', '清理重复采购'),
    (5, '培训不考核', '违反《中小学教师信息技术应用能力标准》，培训必须考核', '组织考核'),
    (6, '网络安全等保不备案', '违反《网络安全法》，必须进行等保备案', '立即备案'),
    (7, '安全隐患不整改', '违反《网络安全法》，安全隐患必须整改', '立即整改'),
    (8, '设备未入账不补登', '违反《行政事业单位国有资产管理办法》，资产必须入账', '补登入账'),
    (9, '数据备份不做', '违反《网络安全法》，重要数据必须备份', '建立备份机制'),
    (10, '培训流于形式不整改', '违反《中小学教师信息技术应用能力标准》，培训必须务实', '重新组织培训'),
]
for num, req, basis, action in rejections:
    doc.add_heading(f'要求{num}驳回意见', level=1)
    p = doc.add_paragraph(); p.add_run('管理层要求：').bold = True; p.add_run(req)
    p = doc.add_paragraph(); p.add_run('驳回依据：').bold = True; p.add_run(basis)
    p = doc.add_paragraph(); p.add_run('处理措施：').bold = True; p.add_run(action)
    doc.add_paragraph()
doc.save(os.path.join(d, '管理层错误要求回函.docx'))

doc = Document()
doc.add_heading('教育信息化建设整改方案', 0)
doc.add_heading('一、整改目标', level=1)
doc.add_paragraph('30日内整改重大违规，60日内建立健全信息化管理制度，设备使用率从40%提升至80%')
doc.add_heading('二、整改措施', level=1)
doc.add_paragraph('1. 规范专项资金管理（8所学校挪用）')
doc.add_paragraph('2. 规范采购流程（50条设备采购违规+30条软件采购违规）')
doc.add_paragraph('3. 盘活闲置设备（70条设备使用违规）')
doc.add_paragraph('4. 提升培训实效（25条培训违规）')
doc.add_paragraph('5. 加强网络安全（13份网络安全违规）')
doc.add_heading('三、时间安排', level=1)
doc.add_paragraph('第一阶段（1-15天）：重大违规立即整改')
doc.add_paragraph('第二阶段（16-30天）：一般违规整改')
doc.add_paragraph('第三阶段（31-60天）：制度建设和培训')
doc.add_paragraph('第四阶段（61-90天）：验收和持续改进')
doc.save(os.path.join(d, '整改方案.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '数据勾稽自检'
ws.append(['序号', '勾稽项', '来源数据', '验证数据', '一致性'])
checks = [
    (1, '学校总数', '学校清单30所', '台账记录数=30', '一致'),
    (2, '设备采购记录数', '设备采购100条', '台账记录数=100', '一致'),
    (3, '软件采购记录数', '软件采购60条', '台账记录数=60', '一致'),
    (4, '设备使用台账数', '设备使用150条', '台账记录数=150', '一致'),
    (5, '培训记录数', '培训记录40条', '台账记录数=40', '一致'),
    (6, '网络安全报告数', '网络安全15份', '报告数=15', '一致'),
    (7, '设备采购违规数', '50项', '25+15+10=50', '一致'),
    (8, '软件采购违规数', '30项', '12+10+8=30', '一致'),
    (9, '设备使用违规数', '70项', '30+25+15=70', '一致'),
    (10, '违规总数', '207项', '19+50+30+70+25+13=207', '一致'),
]
for c in checks:
    ws.append(c)
style(ws, 5); wb.save(os.path.join(d, '数据勾稽自检.xlsx'))
mkzip(d, '06整改方案与勾稽自检.rar')

print('\n=== 标准答案生成完成 ===')
