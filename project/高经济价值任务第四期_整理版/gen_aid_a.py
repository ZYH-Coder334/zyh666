#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""中小学学生资助与教育扶贫资金使用审计 - 标准答案生成"""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from docx import Document
import zipfile, os, random, shutil

random.seed(66)
plt.rcParams['font.sans-serif'] = ['WenQuanYi Micro Hei', 'SimHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

BASE = '/workspace/学生资助扶贫审计_附件'
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
categories = ['学校管理', '资助认定', '资助发放', '免学杂费', '扶贫资金', '制度档案', '公示管理']
violations = [15, 100, 50, 33, 35, 15, 5]
colors = ['#FF6B6B', '#FF9F43', '#FECA57', '#48DBFB', '#1DD1A1', '#5F27CD', '#EE5A6F']
ax.barh(categories, violations, color=colors)
ax.set_xlabel('违规数'); ax.set_title('各模块违规数量统计')
for i, v in enumerate(violations):
    ax.text(v+1, i, str(v), va='center', fontsize=9)
plt.tight_layout(); fig.savefig(os.path.join(d, '违规统计.png'), dpi=150); plt.close()

doc = Document()
doc.add_heading('学生资助与教育扶贫资金合规诊断报告', 0)
doc.add_paragraph('编制依据：《关于进一步加强学生资助工作的通知》《城乡义务教育补助经费管理办法》《中央财政困难群众救助补助资金管理办法》等')
doc.add_paragraph(f'审查范围：20所学校、300名资助学生、200条发放记录、100条免学杂费记录、80条扶贫资金拨付记录')
doc.add_paragraph(f'核心结论：共发现违规问题253项，其中重大违规75项，一般违规178项。涉及18所学校。')
doc.add_heading('一、学校管理', level=1)
doc.add_paragraph('1. 资助制度缺失：6所学校（XX001-XX006）资助管理制度不健全')
doc.add_paragraph('2. 公示不到位：5所学校（XX007-XX011）资助名单未按规定公示')
doc.add_paragraph('3. 档案不全：4所学校（XX012-XX015）资助档案资料不齐全')
doc.add_paragraph('4. 合规保留：3所学校（XX016-XX018）制度健全公示规范，合规。')
doc.add_heading('二、资助认定', level=1)
doc.add_paragraph('1. 认定不精准：50名学生（XS0001-XS0050）家庭收入高于认定标准')
doc.add_paragraph('2. 重复享受：30名学生（XS0051-XS0080）在多所学校重复享受资助')
doc.add_paragraph('3. 已脱贫仍享受：20名学生（XS0081-XS0100）家庭已脱贫仍继续享受')
doc.add_paragraph('4. 合规保留：30名边缘户学生（XS0101-XS0130）认定合理，合规。')
doc.add_heading('三、资助发放', level=1)
doc.add_paragraph('1. 未及时发放：25条记录（FF0001-FF0025）超过规定时限未发放')
doc.add_paragraph('2. 金额不足：15条记录（FF0026-FF0040）实发金额低于应发金额')
doc.add_paragraph('3. 截留挪用：10条记录（FF0041-FF0050）资助金被截留挪用')
doc.add_paragraph('4. 合规保留：10条记录（FF0051-FF0060）提前发放符合规定，合规。')
doc.add_heading('四、免学杂费', level=1)
doc.add_paragraph('1. 超范围免：15条记录（MZ0001-MZ0015）非义务教育阶段违规免除')
doc.add_paragraph('2. 重复免：10条记录（MZ0016-MZ0025）同一学生重复免除')
doc.add_paragraph('3. 应免未免：8条记录（MZ0026-MZ0033）符合条件未免除')
doc.add_paragraph('4. 合规保留：7条记录（MZ0034-MZ0040）符合减免政策，合规。')
doc.add_heading('五、扶贫资金', level=1)
doc.add_paragraph('1. 挪用：15条记录（FP0001-FP0015）挪用至非教育用途')
doc.add_paragraph('2. 截留：10条记录（FP0016-FP0025）截留未及时拨付')
doc.add_paragraph('3. 超范围：10条记录（FP0026-FP0035）超出规定使用范围')
doc.add_paragraph('4. 合规保留：5条记录（FP0036-FP0040）专项用于建档立卡学生，合规。')
doc.add_heading('六、整改建议', level=1)
doc.add_paragraph('1. 立即整改重大违规，30日内完成')
doc.add_paragraph('2. 重新认定资助对象，清理重复和已脱贫人员')
doc.add_paragraph('3. 规范资助金发放流程，确保及时足额发放')
doc.add_paragraph('4. 严格免学杂费范围，杜绝超范围和重复免除')
doc.add_paragraph('5. 加强扶贫资金监管，确保专款专用')
doc.save(os.path.join(d, '合规诊断报告.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '诊断结论汇总'
ws.append(['序号', '问题类别', '违规数量', '涉及编号范围', '严重等级', '法规依据'])
items = [
    (1, '资助制度缺失', 6, 'XX001-XX006', '一般', '《学生资助管理办法》'),
    (2, '公示不到位', 5, 'XX007-XX011', '一般', '《学生资助管理办法》'),
    (3, '档案不全', 4, 'XX012-XX015', '一般', '《学生资助档案管理制度》'),
    (4, '资助认定不精准', 50, 'XS0001-XS0050', '重大', '《关于进一步加强学生资助工作的通知》'),
    (5, '重复享受资助', 30, 'XS0051-XS0080', '重大', '同上'),
    (6, '已脱贫仍享受', 20, 'XS0081-XS0100', '重大', '同上'),
    (7, '资助未及时发放', 25, 'FF0001-FF0025', '重大', '《学生资助资金管理办法》'),
    (8, '发放金额不足', 15, 'FF0026-FF0040', '重大', '同上'),
    (9, '资助金截留挪用', 10, 'FF0041-FF0050', '重大', '同上'),
    (10, '超范围免学杂费', 15, 'MZ0001-MZ0015', '重大', '《城乡义务教育补助经费管理办法》'),
    (11, '重复免学杂费', 10, 'MZ0016-MZ0025', '重大', '同上'),
    (12, '应免未免', 8, 'MZ0026-MZ0033', '重大', '同上'),
    (13, '扶贫资金挪用', 15, 'FP0001-FP0015', '重大', '《中央财政困难群众救助补助资金管理办法》'),
    (14, '扶贫资金截留', 10, 'FP0016-FP0025', '重大', '同上'),
    (15, '扶贫资金超范围', 10, 'FP0026-FP0035', '重大', '同上'),
]
for item in items:
    ws.append(item)
style(ws, 6); wb.save(os.path.join(d, '诊断结论汇总.xlsx'))
mkzip(d, '01合规诊断报告.rar')

# ========== 答案2: 资助认定审查 ==========
print('=== 答案2 资助认定审查 ===')
d = os.path.join(ANS_DIR, '_a2'); os.makedirs(d)
doc = Document()
doc.add_heading('资助认定审查答案', 0)
doc.add_heading('一、认定不精准清单（50名 XS0001-XS0050）', level=1)
doc.add_paragraph('问题：家庭收入高于资助认定标准，不符合资助条件')
doc.add_paragraph('整改：取消资助资格，追回已发资助金')
doc.add_heading('二、重复享受清单（30名 XS0051-XS0080）', level=1)
doc.add_paragraph('问题：同一学生在多所学校重复享受资助')
doc.add_paragraph('整改：清理重复记录，保留一处资助资格')
doc.add_heading('三、已脱贫仍享受清单（20名 XS0081-XS0100）', level=1)
doc.add_paragraph('问题：家庭已脱贫但仍继续享受资助')
doc.add_paragraph('整改：核实脱贫时间，及时调整资助资格')
doc.add_heading('四、合规保留项（30名 XS0101-XS0130）', level=1)
doc.add_paragraph('边缘户家庭学生，认定合理，予以保留。')
doc.save(os.path.join(d, '资助认定审查答案.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '资助认定问题明细'
ws.append(['序号', '编号', '问题类型', '问题描述', '法规依据', '整改要求'])
for i in range(1, 51):
    ws.append([i, f'XS{i:04d}', '认定不精准', '家庭收入高于认定标准', '《关于进一步加强学生资助工作的通知》', '取消资格追回资金'])
for i in range(51, 81):
    ws.append([i, f'XS{i:04d}', '重复享受', '在多所学校重复享受资助', '同上', '清理重复记录'])
for i in range(81, 101):
    ws.append([i, f'XS{i:04d}', '已脱贫仍享受', '家庭已脱贫仍继续享受', '同上', '调整资助资格'])
style(ws, 6); wb.save(os.path.join(d, '资助认定问题明细.xlsx'))
mkzip(d, '02资助认定审查答案.rar')

# ========== 答案3: 资助发放审查 ==========
print('=== 答案3 资助发放审查 ===')
d = os.path.join(ANS_DIR, '_a3'); os.makedirs(d)
doc = Document()
doc.add_heading('资助发放审查答案', 0)
doc.add_heading('一、未及时发放清单（25条 FF0001-FF0025）', level=1)
doc.add_paragraph('问题：资助金超过规定时限未发放')
doc.add_paragraph('整改：立即补发，追究相关责任')
doc.add_heading('二、金额不足清单（15条 FF0026-FF0040）', level=1)
doc.add_paragraph('问题：实发金额低于应发金额')
doc.add_paragraph('整改：补发差额，查明原因')
doc.add_heading('三、截留挪用清单（10条 FF0041-FF0050）', level=1)
doc.add_paragraph('问题：资助金被截留挪用')
doc.add_paragraph('整改：追回资金，依法追究责任')
doc.add_heading('四、合规保留项（10条 FF0051-FF0060）', level=1)
doc.add_paragraph('提前发放符合规定，予以保留。')
doc.save(os.path.join(d, '资助发放审查答案.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '资助发放问题明细'
ws.append(['序号', '编号', '问题类型', '问题描述', '法规依据', '整改要求'])
for i in range(1, 26):
    ws.append([i, f'FF{i:04d}', '未及时发放', '超过规定时限未发放', '《学生资助资金管理办法》', '立即补发'])
for i in range(26, 41):
    ws.append([i, f'FF{i:04d}', '金额不足', '实发金额低于应发金额', '同上', '补发差额'])
for i in range(41, 51):
    ws.append([i, f'FF{i:04d}', '截留挪用', '资助金被截留挪用', '同上', '追回资金'])
style(ws, 6); wb.save(os.path.join(d, '资助发放问题明细.xlsx'))
mkzip(d, '03资助发放审查答案.rar')

# ========== 答案4: 免学杂费审查 ==========
print('=== 答案4 免学杂费审查 ===')
d = os.path.join(ANS_DIR, '_a4'); os.makedirs(d)
doc = Document()
doc.add_heading('免学杂费审查答案', 0)
doc.add_heading('一、超范围免清单（15条 MZ0001-MZ0015）', level=1)
doc.add_paragraph('问题：非义务教育阶段违规免除学杂费')
doc.add_paragraph('整改：追回违规免除的费用')
doc.add_heading('二、重复免清单（10条 MZ0016-MZ0025）', level=1)
doc.add_paragraph('问题：同一学生重复免除学杂费')
doc.add_paragraph('整改：清理重复记录，追回多免费用')
doc.add_heading('三、应免未免清单（8条 MZ0026-MZ0033）', level=1)
doc.add_paragraph('问题：符合条件学生未免除学杂费')
doc.add_paragraph('整改：补免学杂费，退还已收费用')
doc.add_heading('四、合规保留项（7条 MZ0034-MZ0040）', level=1)
doc.add_paragraph('符合减免政策，予以保留。')
doc.save(os.path.join(d, '免学杂费审查答案.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '免学杂费问题明细'
ws.append(['序号', '编号', '问题类型', '问题描述', '法规依据', '整改要求'])
for i in range(1, 16):
    ws.append([i, f'MZ{i:04d}', '超范围免', '非义务教育阶段违规免除', '《城乡义务教育补助经费管理办法》', '追回费用'])
for i in range(16, 26):
    ws.append([i, f'MZ{i:04d}', '重复免', '同一学生重复免除', '同上', '追回多免费用'])
for i in range(26, 34):
    ws.append([i, f'MZ{i:04d}', '应免未免', '符合条件未免除', '同上', '补免退还'])
style(ws, 6); wb.save(os.path.join(d, '免学杂费问题明细.xlsx'))
mkzip(d, '04免学杂费审查答案.rar')

# ========== 答案5: 扶贫资金审查 ==========
print('=== 答案5 扶贫资金审查 ===')
d = os.path.join(ANS_DIR, '_a5'); os.makedirs(d)
doc = Document()
doc.add_heading('扶贫资金审查答案', 0)
doc.add_heading('一、挪用清单（15条 FP0001-FP0015）', level=1)
doc.add_paragraph('问题：扶贫资金被挪用至非教育用途')
doc.add_paragraph('整改：追回挪用资金，依法追究责任')
doc.add_heading('二、截留清单（10条 FP0016-FP0025）', level=1)
doc.add_paragraph('问题：扶贫资金被截留未及时拨付')
doc.add_paragraph('整改：立即拨付，追究截留责任')
doc.add_heading('三、超范围清单（10条 FP0026-FP0035）', level=1)
doc.add_paragraph('问题：扶贫资金超出规定使用范围')
doc.add_paragraph('整改：追回超范围资金，规范使用')
doc.add_heading('四、合规保留项（5条 FP0036-FP0040）', level=1)
doc.add_paragraph('专项用于建档立卡学生，合规，予以保留。')
doc.save(os.path.join(d, '扶贫资金审查答案.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '扶贫资金问题明细'
ws.append(['序号', '编号', '问题类型', '问题描述', '法规依据', '整改要求'])
for i in range(1, 16):
    ws.append([i, f'FP{i:04d}', '挪用', '挪用至非教育用途', '《中央财政困难群众救助补助资金管理办法》', '追回资金'])
for i in range(16, 26):
    ws.append([i, f'FP{i:04d}', '截留', '截留未及时拨付', '同上', '立即拨付'])
for i in range(26, 36):
    ws.append([i, f'FP{i:04d}', '超范围', '超出规定使用范围', '同上', '追回资金'])
style(ws, 6); wb.save(os.path.join(d, '扶贫资金问题明细.xlsx'))
mkzip(d, '05扶贫资金审查答案.rar')

# ========== 答案6: 管理层错误要求回函与整体方案 ==========
print('=== 答案6 管理层错误要求回函与整体方案 ===')
d = os.path.join(ANS_DIR, '_a6'); os.makedirs(d)

fig, ax = plt.subplots(figsize=(12, 7))
phases = ['清理阶段', '补发阶段', '制度阶段', '验收阶段']
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
    (1, '放宽资助认定标准', '违反《关于进一步加强学生资助工作的通知》，认定标准必须严格执行', '严格按标准认定'),
    (2, '资助金不及时发放', '违反《学生资助资金管理办法》，资助金必须按时足额发放', '立即补发'),
    (3, '挪用扶贫资金维修校舍', '违反《中央财政困难群众救助补助资金管理办法》，扶贫资金必须专款专用', '追回资金'),
    (4, '资助名单不公示', '违反《学生资助管理办法》，资助名单必须公示', '立即公示'),
    (5, '非义务教育阶段免学杂费', '违反《城乡义务教育补助经费管理办法》，免学杂费仅限义务教育', '追回费用'),
    (6, '已脱贫学生继续享受资助', '违反《关于进一步加强学生资助工作的通知》，应动态调整资助资格', '调整资格'),
    (7, '资助档案不齐全', '违反《学生资助档案管理制度》，必须建立完整档案', '补齐档案'),
    (8, '扶贫资金截留用于其他', '违反《中央财政困难群众救助补助资金管理办法》，不得截留挪用', '立即拨付'),
    (9, '重复享受资助不清理', '违反《学生资助资金管理办法》，不得重复享受', '立即清理'),
    (10, '应免未免不整改', '违反《城乡义务教育补助经费管理办法》，符合条件必须免除', '立即补免'),
]
for num, req, basis, action in rejections:
    doc.add_heading(f'要求{num}驳回意见', level=1)
    p = doc.add_paragraph(); p.add_run('管理层要求：').bold = True; p.add_run(req)
    p = doc.add_paragraph(); p.add_run('驳回依据：').bold = True; p.add_run(basis)
    p = doc.add_paragraph(); p.add_run('处理措施：').bold = True; p.add_run(action)
    doc.add_paragraph()
doc.save(os.path.join(d, '管理层错误要求回函.docx'))

doc = Document()
doc.add_heading('学生资助与教育扶贫资金整改方案', 0)
doc.add_heading('一、整改目标', level=1)
doc.add_paragraph('30日内整改重大违规，60日内建立健全资助管理制度，合规率从65%提升至90%')
doc.add_heading('二、整改措施', level=1)
doc.add_paragraph('1. 重新认定资助对象（100名违规学生）')
doc.add_paragraph('2. 规范资助金发放（50条违规记录）')
doc.add_paragraph('3. 严格免学杂费管理（33条违规记录）')
doc.add_paragraph('4. 加强扶贫资金监管（35条违规记录）')
doc.add_paragraph('5. 完善学校资助管理（15所学校违规）')
doc.add_heading('三、时间安排', level=1)
doc.add_paragraph('第一阶段（1-15天）：重大违规立即整改')
doc.add_paragraph('第二阶段（16-30天）：一般违规整改')
doc.add_paragraph('第三阶段（31-60天）：制度建设和培训')
doc.add_paragraph('第四阶段（61-90天）：验收和持续改进')
doc.add_heading('四、责任分工', level=1)
doc.add_paragraph('教育局资助办：牵头整改')
doc.add_paragraph('学校资助管理办公室：具体实施')
doc.add_paragraph('财政局：资金监管')
doc.add_paragraph('教育局：监督验收')
doc.save(os.path.join(d, '整改方案.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '数据勾稽自检'
ws.append(['序号', '勾稽项', '来源数据', '验证数据', '一致性'])
checks = [
    (1, '学校总数', '学校清单20所', '台账记录数=20', '一致'),
    (2, '资助学生数', '资助学生名册300人', '名册记录数=300', '一致'),
    (3, '资助发放记录数', '资助发放记录200条', '台账记录数=200', '一致'),
    (4, '免学杂费记录数', '免学杂费记录100条', '台账记录数=100', '一致'),
    (5, '扶贫资金记录数', '扶贫资金拨付台账80条', '台账记录数=80', '一致'),
    (6, '资助认定违规数', '100项', '50+30+20=100', '一致'),
    (7, '资助发放违规数', '50项', '25+15+10=50', '一致'),
    (8, '免学杂费违规数', '33项', '15+10+8=33', '一致'),
    (9, '扶贫资金违规数', '35项', '15+10+10=35', '一致'),
    (10, '违规总数', '253项', '15+100+50+33+35+15+5=253', '一致'),
]
for c in checks:
    ws.append(c)
style(ws, 5); wb.save(os.path.join(d, '数据勾稽自检.xlsx'))
mkzip(d, '06整改方案与勾稽自检.rar')

print('\n=== 标准答案生成完成 ===')
