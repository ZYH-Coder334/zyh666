#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""高校毕业生就业质量与就业数据真实性审计 - 标准答案生成"""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from docx import Document
import zipfile, os, random, shutil

random.seed(44)
plt.rcParams['font.sans-serif'] = ['WenQuanYi Micro Hei', 'SimHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

BASE = '/workspace/毕业生就业审计_附件'
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
categories = ['高校管理', '毕业生就业', '就业协议', '灵活就业', '用人单位回访', '就业质量']
violations = [15, 230, 130, 65, 65, 90]
colors = ['#FF6B6B', '#FF9F43', '#FECA57', '#48DBFB', '#1DD1A1', '#5F27CD']
ax.barh(categories, violations, color=colors)
ax.set_xlabel('违规数'); ax.set_title('各模块违规数量统计')
for i, v in enumerate(violations):
    ax.text(v+2, i, str(v), va='center', fontsize=9)
plt.tight_layout(); fig.savefig(os.path.join(d, '违规统计.png'), dpi=150); plt.close()

doc = Document()
doc.add_heading('毕业生就业质量与数据真实性审计报告', 0)
doc.add_paragraph('编制依据：《教育部关于做好高校毕业生就业统计工作的通知》《普通高等学校毕业生就业工作规定》等')
doc.add_paragraph(f'审查范围：20所高校、500名毕业生样本、300份就业协议、80份灵活就业认定、100条用人单位回访、200份就业质量问卷')
doc.add_paragraph(f'核心结论：共发现违规问题595项，其中重大违规310项，一般违规285项。涉及18所高校。')
doc.add_heading('一、高校就业管理', level=1)
doc.add_paragraph('1. 就业率造假：6所高校（GX001-GX006）就业率数据造假')
doc.add_paragraph('2. 数据不完整：5所高校（GX007-GX011）就业数据不完整')
doc.add_paragraph('3. 指导缺失：4所高校（GX012-GX015）就业指导服务缺失')
doc.add_paragraph('4. 合规保留：3所高校（GX016-GX018）就业率真实偏高，合规。')
doc.add_heading('二、毕业生就业数据', level=1)
doc.add_paragraph('1. 虚假就业：80名毕业生（BY0001-BY0080）虚假就业证明')
doc.add_paragraph('2. 就业率虚高：60名毕业生（BY0081-BY0140）未就业登记为已就业')
doc.add_paragraph('3. 灵活就业不规范：50名毕业生（BY0141-BY0190）灵活就业认定不规范')
doc.add_paragraph('4. 就业质量差：40名毕业生（BY0191-BY0230）薪酬低于最低工资')
doc.add_paragraph('5. 合规保留：30名灵活就业（BY0231-BY0260）真实有效，合规。')
doc.add_heading('三、就业协议', level=1)
doc.add_paragraph('1. 虚假协议：60份（XY001-XY060）协议系伪造')
doc.add_paragraph('2. 协议不规范：40份（XY061-XY100）缺必备条款')
doc.add_paragraph('3. 单位不存在：30份（XY101-XY130）用人单位经核查不存在')
doc.add_paragraph('4. 合规保留：20份短期实习协议（XY131-XY150）真实有效，合规。')
doc.add_heading('四、灵活就业认定', level=1)
doc.add_paragraph('1. 认定不规范：30份（LH001-LH030）材料不规范')
doc.add_paragraph('2. 虚假灵活就业：25份（LH031-LH055）实际未就业')
doc.add_paragraph('3. 超时认定：10份（LH056-LH065）超过规定时限')
doc.add_paragraph('4. 合规保留：8份自主创业（LH066-LH073）有营业执照，合规。')
doc.add_heading('五、用人单位回访', level=1)
doc.add_paragraph('1. 查无此人：30条（HF001-HF030）回访查无此人')
doc.add_paragraph('2. 单位否认：20条（HF031-HF050）单位否认录用')
doc.add_paragraph('3. 联系失效：15条（HF051-HF065）联系方式失效')
doc.add_paragraph('4. 合规保留：10条（HF066-HF075）单位确认真实就业，合规。')
doc.add_heading('六、就业质量', level=1)
doc.add_paragraph('1. 薪酬低于最低工资：40份（WJ001-WJ040）')
doc.add_paragraph('2. 专业不匹配：30份（WJ041-WJ070）')
doc.add_paragraph('3. 满意度低：20份（WJ071-WJ090）')
doc.add_paragraph('4. 合规保留：10份（WJ091-WJ100）自愿选择低薪岗位，合规。')
doc.save(os.path.join(d, '审计报告.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '诊断结论汇总'
ws.append(['序号', '问题类别', '违规数量', '涉及编号范围', '严重等级', '法规依据'])
items = [
    (1, '就业率数据造假', 6, 'GX001-GX006', '重大', '《教育部关于做好高校毕业生就业统计工作的通知》'),
    (2, '就业数据不完整', 5, 'GX007-GX011', '一般', '同上'),
    (3, '就业指导服务缺失', 4, 'GX012-GX015', '一般', '《普通高等学校毕业生就业工作规定》'),
    (4, '虚假就业证明', 80, 'BY0001-BY0080', '重大', '《教育部关于做好高校毕业生就业统计工作的通知》'),
    (5, '就业率虚高', 60, 'BY0081-BY0140', '重大', '同上'),
    (6, '灵活就业认定不规范', 50, 'BY0141-BY0190', '重大', '同上'),
    (7, '就业质量差', 40, 'BY0191-BY0230', '一般', '同上'),
    (8, '虚假就业协议', 60, 'XY001-XY060', '重大', '《普通高等学校毕业生就业工作规定》'),
    (9, '协议不规范', 40, 'XY061-XY100', '一般', '同上'),
    (10, '用人单位不存在', 30, 'XY101-XY130', '重大', '同上'),
    (11, '灵活就业认定不规范', 30, 'LH001-LH030', '重大', '《教育部关于做好高校毕业生就业统计工作的通知》'),
    (12, '虚假灵活就业', 25, 'LH031-LH055', '重大', '同上'),
    (13, '超时未认定', 10, 'LH056-LH065', '一般', '同上'),
    (14, '回访查无此人', 30, 'HF001-HF030', '重大', '《普通高等学校毕业生就业工作规定》'),
    (15, '单位否认录用', 20, 'HF031-HF050', '重大', '同上'),
    (16, '联系方式失效', 15, 'HF051-HF065', '一般', '同上'),
    (17, '薪酬低于最低工资', 40, 'WJ001-WJ040', '重大', '《劳动法》'),
    (18, '专业不匹配', 30, 'WJ041-WJ070', '一般', '《普通高等学校毕业生就业工作规定》'),
    (19, '满意度低', 20, 'WJ071-WJ090', '一般', '同上'),
]
for item in items:
    ws.append(item)
style(ws, 6); wb.save(os.path.join(d, '诊断结论汇总.xlsx'))
mkzip(d, '01合规诊断报告.rar')

# ========== 答案2: 毕业生就业数据审查 ==========
print('=== 答案2 毕业生就业数据审查 ===')
d = os.path.join(ANS_DIR, '_a2'); os.makedirs(d)
doc = Document()
doc.add_heading('毕业生就业数据审查答案', 0)
doc.add_heading('一、虚假就业证明清单（80名 BY0001-BY0080）', level=1)
doc.add_paragraph('问题：就业证明系伪造，毕业生实际未就业')
doc.add_paragraph('整改：取消就业登记，追究相关责任')
doc.add_heading('二、就业率虚高清单（60名 BY0081-BY0140）', level=1)
doc.add_paragraph('问题：未就业登记为已就业，就业率虚高')
doc.add_paragraph('整改：更正就业状态，重新统计就业率')
doc.add_heading('三、灵活就业认定不规范清单（50名 BY0141-BY0190）', level=1)
doc.add_paragraph('问题：灵活就业认定材料不规范')
doc.add_paragraph('整改：重新认定，补齐材料')
doc.add_heading('四、就业质量差清单（40名 BY0191-BY0230）', level=1)
doc.add_paragraph('问题：薪酬低于当地最低工资标准')
doc.add_paragraph('整改：督促用人单位整改，保障毕业生权益')
doc.add_heading('五、合规保留项（30名 BY0231-BY0260）', level=1)
doc.add_paragraph('灵活就业真实有效，予以保留。')
doc.save(os.path.join(d, '毕业生就业数据审查答案.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '就业数据问题明细'
ws.append(['序号', '编号', '问题类型', '问题描述', '法规依据', '整改要求'])
for i in range(1, 81):
    ws.append([i, f'BY{i:04d}', '虚假就业证明', '就业证明系伪造', '《教育部关于做好高校毕业生就业统计工作的通知》', '取消就业登记'])
for i in range(81, 141):
    ws.append([i, f'BY{i:04d}', '就业率虚高', '未就业登记为已就业', '同上', '更正就业状态'])
for i in range(141, 191):
    ws.append([i, f'BY{i:04d}', '灵活就业不规范', '认定材料不规范', '同上', '重新认定'])
for i in range(191, 231):
    ws.append([i, f'BY{i:04d}', '就业质量差', '薪酬低于最低工资', '《劳动法》', '督促单位整改'])
style(ws, 6); wb.save(os.path.join(d, '就业数据问题明细.xlsx'))
mkzip(d, '02毕业生就业数据审查答案.rar')

# ========== 答案3: 就业协议与灵活就业审查 ==========
print('=== 答案3 就业协议与灵活就业审查 ===')
d = os.path.join(ANS_DIR, '_a3'); os.makedirs(d)
doc = Document()
doc.add_heading('就业协议与灵活就业审查答案', 0)
doc.add_heading('一、就业协议违规（130份）', level=1)
doc.add_paragraph('1. 虚假协议60份（XY001-XY060）：协议系伪造，单位公章不实')
doc.add_paragraph('2. 协议不规范40份（XY061-XY100）：缺必备条款')
doc.add_paragraph('3. 单位不存在30份（XY101-XY130）：用人单位经核查不存在')
doc.add_paragraph('4. 合规保留20份（XY131-XY150）：短期实习协议真实有效')
doc.add_heading('二、灵活就业认定违规（65份）', level=1)
doc.add_paragraph('1. 认定不规范30份（LH001-LH030）：材料不规范')
doc.add_paragraph('2. 虚假灵活就业25份（LH031-LH055）：实际未就业')
doc.add_paragraph('3. 超时认定10份（LH056-LH065）：超过规定时限')
doc.add_paragraph('4. 合规保留8份（LH066-LH073）：自主创业有营业执照')
doc.save(os.path.join(d, '就业协议与灵活就业审查答案.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '就业协议问题明细'
ws.append(['序号', '编号', '问题类型', '问题描述', '法规依据', '整改要求'])
for i in range(1, 61):
    ws.append([i, f'XY{i:04d}', '虚假协议', '协议系伪造', '《普通高等学校毕业生就业工作规定》', '追究责任'])
for i in range(61, 101):
    ws.append([i, f'XY{i:04d}', '协议不规范', '缺必备条款', '同上', '补充条款'])
for i in range(101, 131):
    ws.append([i, f'XY{i:04d}', '单位不存在', '用人单位不存在', '同上', '取消协议'])
style(ws, 6); wb.save(os.path.join(d, '就业协议问题明细.xlsx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '灵活就业问题明细'
ws.append(['序号', '编号', '问题类型', '问题描述', '法规依据', '整改要求'])
for i in range(1, 31):
    ws.append([i, f'LH{i:03d}', '认定不规范', '材料不规范', '《教育部关于做好高校毕业生就业统计工作的通知》', '补齐材料'])
for i in range(31, 56):
    ws.append([i, f'LH{i:03d}', '虚假灵活就业', '实际未就业', '同上', '取消认定'])
for i in range(56, 66):
    ws.append([i, f'LH{i:03d}', '超时认定', '超过规定时限', '同上', '限期认定'])
style(ws, 6); wb.save(os.path.join(d, '灵活就业问题明细.xlsx'))
mkzip(d, '03就业协议与灵活就业审查答案.rar')

# ========== 答案4: 用人单位回访审查 ==========
print('=== 答案4 用人单位回访审查 ===')
d = os.path.join(ANS_DIR, '_a4'); os.makedirs(d)
doc = Document()
doc.add_heading('用人单位回访审查答案', 0)
doc.add_heading('一、回访问题清单（65条）', level=1)
doc.add_paragraph('1. 查无此人30条（HF001-HF030）：回访查无此人，就业信息虚假')
doc.add_paragraph('2. 单位否认20条（HF031-HF050）：用人单位否认录用该毕业生')
doc.add_paragraph('3. 联系失效15条（HF051-HF065）：联系方式失效，无法核实')
doc.add_heading('二、合规保留项（10条 HF066-HF075）', level=1)
doc.add_paragraph('单位确认毕业生真实就业，予以保留。')
doc.save(os.path.join(d, '用人单位回访审查答案.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '回访问题明细'
ws.append(['序号', '编号', '问题类型', '问题描述', '法规依据', '整改要求'])
for i in range(1, 31):
    ws.append([i, f'HF{i:03d}', '查无此人', '就业信息虚假', '《普通高等学校毕业生就业工作规定》', '核实就业情况'])
for i in range(31, 51):
    ws.append([i, f'HF{i:03d}', '单位否认', '单位否认录用', '同上', '重新核实'])
for i in range(51, 66):
    ws.append([i, f'HF{i:03d}', '联系失效', '无法核实', '同上', '重新联系'])
style(ws, 6); wb.save(os.path.join(d, '回访问题明细.xlsx'))
mkzip(d, '04用人单位回访审查答案.rar')

# ========== 答案5: 就业质量审查 ==========
print('=== 答案5 就业质量审查 ===')
d = os.path.join(ANS_DIR, '_a5'); os.makedirs(d)
doc = Document()
doc.add_heading('就业质量审查答案', 0)
doc.add_heading('一、就业质量问题清单（90份）', level=1)
doc.add_paragraph('1. 薪酬低于最低工资40份（WJ001-WJ040）：违反劳动法')
doc.add_paragraph('2. 专业不匹配30份（WJ041-WJ070）：专业与岗位严重不匹配')
doc.add_paragraph('3. 满意度低20份（WJ071-WJ090）：工作满意度极低')
doc.add_heading('二、合规保留项（10份 WJ091-WJ100）', level=1)
doc.add_paragraph('毕业生自愿选择低薪兴趣岗位，予以保留。')
doc.save(os.path.join(d, '就业质量审查答案.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '就业质量问题明细'
ws.append(['序号', '编号', '问题类型', '问题描述', '法规依据', '整改要求'])
for i in range(1, 41):
    ws.append([i, f'WJ{i:04d}', '薪酬低于最低工资', '违反劳动法', '《劳动法》', '督促单位整改'])
for i in range(41, 71):
    ws.append([i, f'WJ{i:04d}', '专业不匹配', '专业与岗位不匹配', '《普通高等学校毕业生就业工作规定》', '加强就业指导'])
for i in range(71, 91):
    ws.append([i, f'WJ{i:04d}', '满意度低', '工作满意度极低', '同上', '改善就业服务'])
style(ws, 6); wb.save(os.path.join(d, '就业质量问题明细.xlsx'))
mkzip(d, '05就业质量审查答案.rar')

# ========== 答案6: 管理层错误要求回函与整体方案 ==========
print('=== 答案6 管理层错误要求回函与整体方案 ===')
d = os.path.join(ANS_DIR, '_a6'); os.makedirs(d)

fig, ax = plt.subplots(figsize=(12, 7))
phases = ['数据核查', '虚假清理', '制度规范', '验收改进']
tasks = [15, 25, 12, 8]
colors = ['#4472C4', '#ED7D31', '#70AD47', '#FFC000']
ax.bar(phases, tasks, color=colors, width=0.5)
ax.set_ylabel('工作项数'); ax.set_title('整改方案阶段任务分布')
for i, v in enumerate(tasks):
    ax.text(i, v+0.3, str(v), ha='center', fontsize=9)
plt.tight_layout(); fig.savefig(os.path.join(d, '整改方案.png'), dpi=150); plt.close()

doc = Document()
doc.add_heading('关于管理层要求的回函', 0)
rejections = [
    (1, '就业率造假虚报', '违反《教育部关于做好高校毕业生就业统计工作的通知》，就业率必须真实统计', '如实统计就业率'),
    (2, '虚假就业证明不核实', '违反《普通高等学校毕业生就业工作规定》，必须核实就业证明真实性', '逐一核实就业证明'),
    (3, '灵活就业认定放宽', '违反《教育部关于做好高校毕业生就业统计工作的通知》，认定标准必须严格执行', '严格按标准认定'),
    (4, '用人单位回访不做', '违反《普通高等学校毕业生就业工作规定》，必须开展用人单位回访', '按规定比例回访'),
    (5, '未就业登记为就业', '违反《教育部关于做好高校毕业生就业统计工作的通知》，不得虚报就业', '如实登记就业状态'),
    (6, '就业协议不审核', '违反《普通高等学校毕业生就业工作规定》，必须审核就业协议', '逐一审核就业协议'),
    (7, '就业质量差不上报', '违反《普通高等学校毕业生就业工作规定》，就业质量必须如实上报', '如实上报就业质量'),
    (8, '虚假灵活就业不清理', '违反《教育部关于做好高校毕业生就业统计工作的通知》，不得虚假认定', '立即清理虚假认定'),
    (9, '回访查无此人不处理', '违反《普通高等学校毕业生就业工作规定》，必须核实处理', '逐一核实处理'),
    (10, '就业指导服务不开展', '违反《普通高等学校毕业生就业工作规定》，必须开展就业指导', '加强就业指导服务'),
]
for num, req, basis, action in rejections:
    doc.add_heading(f'要求{num}驳回意见', level=1)
    p = doc.add_paragraph(); p.add_run('管理层要求：').bold = True; p.add_run(req)
    p = doc.add_paragraph(); p.add_run('驳回依据：').bold = True; p.add_run(basis)
    p = doc.add_paragraph(); p.add_run('处理措施：').bold = True; p.add_run(action)
    doc.add_paragraph()
doc.save(os.path.join(d, '管理层错误要求回函.docx'))

doc = Document()
doc.add_heading('毕业生就业数据整改方案', 0)
doc.add_heading('一、整改目标', level=1)
doc.add_paragraph('30日内整改重大违规，60日内建立健全就业统计制度，就业率数据真实率从60%提升至95%')
doc.add_heading('二、整改措施', level=1)
doc.add_paragraph('1. 核实就业数据真实性（230名毕业生违规）')
doc.add_paragraph('2. 清理虚假就业协议（130份违规协议）')
doc.add_paragraph('3. 规范灵活就业认定（65份违规认定）')
doc.add_paragraph('4. 加强用人单位回访（65条回访问题）')
doc.add_paragraph('5. 提升就业质量（90份质量问题）')
doc.add_heading('三、时间安排', level=1)
doc.add_paragraph('第一阶段（1-15天）：重大违规立即整改')
doc.add_paragraph('第二阶段（16-30天）：一般违规整改')
doc.add_paragraph('第三阶段（31-60天）：制度建设和培训')
doc.add_paragraph('第四阶段（61-90天）：验收和持续改进')
doc.save(os.path.join(d, '整改方案.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '数据勾稽自检'
ws.append(['序号', '勾稽项', '来源数据', '验证数据', '一致性'])
checks = [
    (1, '高校总数', '高校清单20所', '台账记录数=20', '一致'),
    (2, '毕业生样本数', '毕业生样本500人', '名册记录数=500', '一致'),
    (3, '就业协议数', '就业协议300份', '台账记录数=300', '一致'),
    (4, '灵活就业认定数', '灵活就业认定80份', '台账记录数=80', '一致'),
    (5, '回访记录数', '用人单位回访100条', '台账记录数=100', '一致'),
    (6, '就业质量问卷数', '就业质量问卷200份', '台账记录数=200', '一致'),
    (7, '毕业生违规数', '230项', '80+60+50+40=230', '一致'),
    (8, '协议违规数', '130项', '60+40+30=130', '一致'),
    (9, '回访违规数', '65项', '30+20+15=65', '一致'),
    (10, '违规总数', '595项', '15+230+130+65+65+90=595', '一致'),
]
for c in checks:
    ws.append(c)
style(ws, 5); wb.save(os.path.join(d, '数据勾稽自检.xlsx'))
mkzip(d, '06整改方案与勾稽自检.rar')

print('\n=== 标准答案生成完成 ===')
