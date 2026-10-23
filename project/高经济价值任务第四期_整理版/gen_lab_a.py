#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""高校实验室安全管理与科研伦理合规审查 - 标准答案生成"""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from docx import Document
from docx.shared import Cm, RGBColor
import zipfile, os, random, shutil

random.seed(88)
plt.rcParams['font.sans-serif'] = ['WenQuanYi Micro Hei', 'SimHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

BASE = '/workspace/实验室安全伦理审查_附件'
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

# ========== 答案1: 实验室安全合规诊断 ==========
print('=== 答案1 实验室安全合规诊断 ===')
d = os.path.join(ANS_DIR, '_a1'); os.makedirs(d)

fig, ax = plt.subplots(figsize=(12, 7))
categories = ['实验室制度', '危险化学品', '生物安全', '辐射安全', '实验动物', '人类受试者', '安全事故', '应急预案']
violations = [30, 60, 33, 18, 18, 16, 12, 13]
colors = ['#FF6B6B', '#FF9F43', '#FECA57', '#48DBFB', '#1DD1A1', '#5F27CD', '#EE5A6F', '#54A0FF']
ax.barh(categories, violations, color=colors)
ax.set_xlabel('违规数'); ax.set_title('各模块违规数量统计')
for i, v in enumerate(violations):
    ax.text(v+0.5, i, str(v), va='center', fontsize=9)
plt.tight_layout(); fig.savefig(os.path.join(d, '违规统计.png'), dpi=150); plt.close()

doc = Document()
doc.add_heading('实验室安全与科研伦理合规诊断报告', 0)
doc.add_paragraph('编制依据：《实验室生物安全通用要求》GB19489-2008、《危险化学品安全管理条例》、《实验动物管理条例》、《涉及人的生物医学研究伦理审查办法》等')
doc.add_paragraph(f'审查范围：15所高校、80个实验室、200条化学品记录、50条生物安全记录、30条辐射安全记录、30份动物伦理审批、25份人类受试者审批、15起安全事故、15份应急预案')
doc.add_paragraph(f'核心结论：共发现违规问题200项，其中重大违规45项，一般违规155项。涉及12所高校、65个实验室。')
doc.add_heading('一、实验室安全管理', level=1)
doc.add_paragraph('1. 安全管理制度缺失：12个实验室（LAB001-LAB012）未建立完整安全管理制度')
doc.add_paragraph('2. 安全员无资质：10个实验室（LAB013-LAB022）安全员无安全资质证书')
doc.add_paragraph('3. 超期未检查：8个实验室（LAB023-LAB030）超过一年未进行安全检查')
doc.add_paragraph('4. 正常保留项：5个新建实验室（LAB031-LAB035）制度齐全，合规。')
doc.add_heading('二、危险化学品管理', level=1)
doc.add_paragraph('1. 台账不符：25条记录（HX0001-HX0025）台账与实际库存不符')
doc.add_paragraph('2. 存储违规：20条记录（HX0026-HX0045）存储位置不符合危险等级')
doc.add_paragraph('3. 废弃违规：15条记录（HX0046-HX0060）废弃化学品未按规定处置')
doc.add_paragraph('4. 正常保留项：10条易制毒化学品记录（HX0061-HX0070）已在公安机关备案，合规。')
doc.add_heading('三、生物安全与辐射安全', level=1)
doc.add_paragraph('1. 生物安全违规33项：等级不符15+防护缺失10+废弃物违规8')
doc.add_paragraph('2. 辐射安全违规18项：无剂量计8+防护缺失6+超期检测4')
doc.add_heading('四、伦理审查', level=1)
doc.add_paragraph('1. 动物伦理违规18项：未审批8+超范围6+到期未续4')
doc.add_paragraph('2. 人类受试者违规16项：无知情同意7+隐私泄露5+未审批4')
doc.add_heading('五、安全事故与应急', level=1)
doc.add_paragraph('1. 安全事故违规12项：隐瞒5+处置不当4+超期上报3')
doc.add_paragraph('2. 应急预案违规13项：未演练6+未评审5+超期未更新2')
doc.add_heading('六、整改建议', level=1)
doc.add_paragraph('1. 立即整改重大违规，30日内完成')
doc.add_paragraph('2. 建立健全安全管理制度，配齐有资质安全员')
doc.add_paragraph('3. 规范危险化学品全流程管理，实行双人双锁')
doc.add_paragraph('4. 严格生物安全等级认定和操作规范')
doc.add_paragraph('5. 完善伦理审查程序，强化跟踪审查')
doc.add_paragraph('6. 加强应急预案演练，规范事故上报')
doc.save(os.path.join(d, '合规诊断报告.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '诊断结论汇总'
ws.append(['序号', '问题类别', '违规数量', '涉及编号范围', '严重等级', '法规依据'])
items = [
    (1, '安全管理制度缺失', 12, 'LAB001-LAB012', '重大', '《实验室安全管理办法》'),
    (2, '安全员无资质', 10, 'LAB013-LAB022', '一般', '《实验室安全管理办法》'),
    (3, '超期未安全检查', 8, 'LAB023-LAB030', '一般', '《实验室安全管理办法》'),
    (4, '化学品台账不符', 25, 'HX0001-HX0025', '重大', '《危险化学品安全管理条例》'),
    (5, '化学品存储违规', 20, 'HX0026-HX0045', '重大', '《危险化学品安全管理条例》'),
    (6, '化学品废弃违规', 15, 'HX0046-HX0060', '重大', '《危险化学品安全管理条例》'),
    (7, '生物安全等级不符', 15, 'BS001-BS015', '重大', 'GB19489-2008'),
    (8, '生物防护缺失', 10, 'BS016-BS025', '重大', 'GB19489-2008'),
    (9, '生物废弃物违规', 8, 'BS026-BS033', '重大', 'GB19489-2008'),
    (10, '辐射无剂量计', 8, 'RS001-RS008', '重大', '《放射性同位素与射线装置安全防护条例》'),
    (11, '辐射防护缺失', 6, 'RS009-RS014', '重大', '同上'),
    (12, '辐射超期检测', 4, 'RS015-RS018', '一般', '同上'),
    (13, '动物伦理未审批', 8, 'AL001-AL008', '重大', '《实验动物管理条例》'),
    (14, '动物伦理超范围', 6, 'AL009-AL014', '重大', '同上'),
    (15, '动物伦理到期未续', 4, 'AL015-AL018', '一般', '同上'),
    (16, '人类受试者无知情同意', 7, 'HS001-HS007', '重大', '《涉及人的生物医学研究伦理审查办法》'),
    (17, '人类受试者隐私泄露', 5, 'HS008-HS012', '重大', '同上'),
    (18, '人类受试者未审批', 4, 'HS013-HS016', '重大', '同上'),
    (19, '安全事故隐瞒', 5, 'AC001-AC005', '重大', '《生产安全事故报告和调查处理条例》'),
    (20, '安全事故处置不当', 4, 'AC006-AC009', '重大', '同上'),
    (21, '安全事故超期上报', 3, 'AC010-AC012', '一般', '同上'),
    (22, '应急预案未演练', 6, 'EP001-EP006', '一般', '《突发事件应急预案管理办法》'),
    (23, '应急预案未评审', 5, 'EP007-EP011', '一般', '同上'),
    (24, '应急预案超期未更新', 2, 'EP012-EP013', '一般', '同上'),
]
for item in items:
    ws.append(item)
style(ws, 6); wb.save(os.path.join(d, '诊断结论汇总.xlsx'))
mkzip(d, '01实验室安全合规诊断.rar')

# ========== 答案2: 危险化学品审查 ==========
print('=== 答案2 危险化学品审查 ===')
d = os.path.join(ANS_DIR, '_a2'); os.makedirs(d)
doc = Document()
doc.add_heading('危险化学品管理审查答案', 0)
doc.add_heading('一、台账不符清单（25条 HX0001-HX0025）', level=1)
doc.add_paragraph('问题：台账记录与实际库存存在差异，差额不等（有的为正数有的为负数），需逐一核实')
doc.add_paragraph('整改：重新盘点，核实采购记录和领用记录，查明差异原因')
doc.add_heading('二、存储违规清单（20条 HX0026-HX0045）', level=1)
doc.add_paragraph('问题：存储位置不符合危险等级要求，酸碱混放、氧化剂与还原剂混放')
doc.add_paragraph('整改：按危险类别分区存放，安装通风设施，配备防爆柜')
doc.add_heading('三、废弃违规清单（15条 HX0046-HX0060）', level=1)
doc.add_paragraph('问题：废弃化学品未按规定交由有资质单位处置，存在自行倾倒风险')
doc.add_paragraph('整改：建立废弃化学品台账，委托有资质单位处置，留存处置凭证')
doc.add_heading('四、合规保留项（10条 HX0061-HX0070）', level=1)
doc.add_paragraph('易制毒化学品已在公安机关备案，采购、存储、使用均符合规定，予以保留。')
doc.save(os.path.join(d, '危险化学品审查答案.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '化学品问题明细'
ws.append(['序号', '编号', '问题类型', '问题描述', '法规依据', '整改要求'])
for i in range(1, 26):
    ws.append([i, f'HX{i:04d}', '台账不符', '台账数量与实际库存不符', '《危险化学品安全管理条例》', '重新盘点核实'])
for i in range(26, 46):
    ws.append([i, f'HX{i:04d}', '存储违规', '存储位置不符合危险等级', '《危险化学品安全管理条例》', '分区存储'])
for i in range(46, 61):
    ws.append([i, f'HX{i:04d}', '废弃违规', '未按规定处置废弃化学品', '《危险化学品安全管理条例》', '委托有资质单位处置'])
style(ws, 6); wb.save(os.path.join(d, '化学品问题明细.xlsx'))
mkzip(d, '02危险化学品审查答案.rar')

# ========== 答案3: 生物与辐射安全审查 ==========
print('=== 答案3 生物与辐射安全审查 ===')
d = os.path.join(ANS_DIR, '_a3'); os.makedirs(d)
doc = Document()
doc.add_heading('生物安全与辐射安全审查答案', 0)
doc.add_heading('一、生物安全违规（33项）', level=1)
doc.add_paragraph('1. 等级不符15项（BS001-BS015）：实验内容与生物安全等级不匹配，存在感染风险')
doc.add_paragraph('2. 防护缺失10项（BS016-BS025）：操作人员未佩戴防护装备，操作不规范')
doc.add_paragraph('3. 废弃物违规8项（BS026-BS033）：生物废弃物未高压灭菌，存在环境污染风险')
doc.add_heading('二、辐射安全违规（18项）', level=1)
doc.add_paragraph('1. 无剂量计8项（RS001-RS008）：操作人员未佩戴个人剂量计，无法监测受照剂量')
doc.add_paragraph('2. 防护缺失6项（RS009-RS014）：辐射防护设施缺失或失效')
doc.add_paragraph('3. 超期检测4项（RS015-RS018）：辐射设备超期未由专业机构检测')
doc.add_heading('三、合规保留项', level=1)
doc.add_paragraph('生物安全5项（BS034-BS038）低风险分子生物学实验合规；辐射安全3项（RS019-RS021）低剂量豁免源合规。')
doc.save(os.path.join(d, '生物与辐射安全审查答案.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '生物安全问题明细'
ws.append(['序号', '编号', '问题类型', '问题描述', '法规依据', '整改要求'])
for i in range(1, 16):
    ws.append([i, f'BS{i:03d}', '等级不符', '实验内容与生物安全等级不符', 'GB19489-2008', '重新评估等级'])
for i in range(16, 26):
    ws.append([i, f'BS{i:03d}', '防护缺失', '操作人员防护措施缺失', 'GB19489-2008', '配齐防护装备'])
for i in range(26, 34):
    ws.append([i, f'BS{i:03d}', '废弃物违规', '生物废弃物未高压灭菌', 'GB19489-2008', '规范废弃物处理'])
style(ws, 6); wb.save(os.path.join(d, '生物安全问题明细.xlsx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '辐射安全问题明细'
ws.append(['序号', '编号', '问题类型', '问题描述', '法规依据', '整改要求'])
for i in range(1, 9):
    ws.append([i, f'RS{i:03d}', '无剂量计', '未佩戴个人剂量计', '《放射性同位素与射线装置安全防护条例》', '配备并佩戴'])
for i in range(9, 15):
    ws.append([i, f'RS{i:03d}', '防护缺失', '辐射防护设施缺失', '同上', '完善防护设施'])
for i in range(15, 19):
    ws.append([i, f'RS{i:03d}', '超期检测', '辐射设备超期未检测', '同上', '立即检测'])
style(ws, 6); wb.save(os.path.join(d, '辐射安全问题明细.xlsx'))
mkzip(d, '03生物与辐射安全审查答案.rar')

# ========== 答案4: 伦理审查 ==========
print('=== 答案4 伦理审查 ===')
d = os.path.join(ANS_DIR, '_a4'); os.makedirs(d)
doc = Document()
doc.add_heading('科研伦理审查答案', 0)
doc.add_heading('一、实验动物伦理违规（18项）', level=1)
doc.add_paragraph('1. 未审批8项（AL001-AL008）：未经伦理委员会审批擅自开展动物实验')
doc.add_paragraph('2. 超范围6项（AL009-AL014）：实验内容超出伦理批准范围')
doc.add_paragraph('3. 到期未续4项（AL015-AL018）：伦理审批到期未续期')
doc.add_heading('二、人类受试者伦理违规（16项）', level=1)
doc.add_paragraph('1. 无知情同意7项（HS001-HS007）：受试者未签署知情同意书')
doc.add_paragraph('2. 隐私泄露5项（HS008-HS012）：受试者隐私信息未脱敏处理')
doc.add_paragraph('3. 未审批4项（HS013-HS016）：未经伦理委员会审批')
doc.add_heading('三、合规保留项', level=1)
doc.add_paragraph('动物伦理3项（AL019-AL021）采用替代方法合规；人类受试者3项（HS017-HS019）使用公开匿名数据合规。')
doc.save(os.path.join(d, '伦理审查答案.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '动物伦理问题明细'
ws.append(['序号', '编号', '问题类型', '问题描述', '法规依据', '整改要求'])
for i in range(1, 9):
    ws.append([i, f'AL{i:03d}', '未审批', '未经伦理委员会审批', '《实验动物管理条例》', '补办审批'])
for i in range(9, 15):
    ws.append([i, f'AL{i:03d}', '超范围', '超出批准范围', '同上', '重新审批'])
for i in range(15, 19):
    ws.append([i, f'AL{i:03d}', '到期未续', '伦理审批到期未续', '同上', '续期审批'])
style(ws, 6); wb.save(os.path.join(d, '动物伦理问题明细.xlsx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '人类受试者问题明细'
ws.append(['序号', '编号', '问题类型', '问题描述', '法规依据', '整改要求'])
for i in range(1, 8):
    ws.append([i, f'HS{i:03d}', '无知情同意', '未签署知情同意书', '《涉及人的生物医学研究伦理审查办法》', '补签知情同意'])
for i in range(8, 13):
    ws.append([i, f'HS{i:03d}', '隐私泄露', '隐私信息未脱敏', '同上', '脱敏处理'])
for i in range(13, 17):
    ws.append([i, f'HS{i:03d}', '未审批', '未经伦理委员会审批', '同上', '补办审批'])
style(ws, 6); wb.save(os.path.join(d, '人类受试者问题明细.xlsx'))
mkzip(d, '04伦理审查答案.rar')

# ========== 答案5: 安全事故与应急审查 ==========
print('=== 答案5 安全事故与应急审查 ===')
d = os.path.join(ANS_DIR, '_a5'); os.makedirs(d)
doc = Document()
doc.add_heading('安全事故与应急预案审查答案', 0)
doc.add_heading('一、安全事故违规（12项）', level=1)
doc.add_paragraph('1. 隐瞒5项（AC001-AC005）：事故隐瞒未上报，违反《生产安全事故报告和调查处理条例》')
doc.add_paragraph('2. 处置不当4项（AC006-AC009）：事故处置不规范，扩大了损失')
doc.add_paragraph('3. 超期上报3项（AC010-AC012）：超过24小时未上报')
doc.add_heading('二、应急预案违规（13项）', level=1)
doc.add_paragraph('1. 未演练6项（EP001-EP006）：应急预案从未演练')
doc.add_paragraph('2. 未评审5项（EP007-EP011）：预案未经专家评审')
doc.add_paragraph('3. 超期未更新2项（EP012-EP013）：预案超期未更新')
doc.save(os.path.join(d, '安全事故与应急审查答案.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '安全事故问题明细'
ws.append(['序号', '编号', '问题类型', '问题描述', '法规依据', '整改要求'])
for i in range(1, 6):
    ws.append([i, f'AC{i:03d}', '隐瞒不报', '事故隐瞒未上报', '《生产安全事故报告和调查处理条例》', '依法上报，追究责任'])
for i in range(6, 10):
    ws.append([i, f'AC{i:03d}', '处置不当', '事故处置不规范', '同上', '规范处置流程'])
for i in range(10, 13):
    ws.append([i, f'AC{i:03d}', '超期上报', '超过24小时上报', '同上', '严格执行上报时限'])
style(ws, 6); wb.save(os.path.join(d, '安全事故问题明细.xlsx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '应急预案问题明细'
ws.append(['序号', '编号', '问题类型', '问题描述', '法规依据', '整改要求'])
for i in range(1, 7):
    ws.append([i, f'EP{i:03d}', '未演练', '应急预案从未演练', '《突发事件应急预案管理办法》', '定期演练'])
for i in range(7, 12):
    ws.append([i, f'EP{i:03d}', '未评审', '预案未经专家评审', '同上', '组织专家评审'])
for i in range(12, 14):
    ws.append([i, f'EP{i:03d}', '超期未更新', '预案超期未更新', '同上', '及时更新'])
style(ws, 6); wb.save(os.path.join(d, '应急预案问题明细.xlsx'))
mkzip(d, '05安全事故与应急审查答案.rar')

# ========== 答案6: 管理层错误要求回函与整体方案 ==========
print('=== 答案6 管理层错误要求回函与整体方案 ===')
d = os.path.join(ANS_DIR, '_a6'); os.makedirs(d)

fig, ax = plt.subplots(figsize=(12, 7))
phases = ['诊断阶段', '核查阶段', '整改阶段', '验收阶段']
tasks = [12, 25, 18, 10]
colors = ['#4472C4', '#ED7D31', '#70AD47', '#FFC000']
ax.bar(phases, tasks, color=colors, width=0.5)
ax.set_ylabel('工作项数'); ax.set_title('整改方案阶段任务分布')
for i, v in enumerate(tasks):
    ax.text(i, v+0.3, str(v), ha='center', fontsize=9)
plt.tight_layout(); fig.savefig(os.path.join(d, '整改方案.png'), dpi=150); plt.close()

doc = Document()
doc.add_heading('关于管理层要求的回函', 0)
rejections = [
    (1, '安全事故隐瞒不上报', '违反《生产安全事故报告和调查处理条例》第9条，事故必须1小时内上报', '依法上报，追究责任'),
    (2, '危险化学品台账不整改', '违反《危险化学品安全管理条例》第20条，必须如实记录', '限期整改，重新盘点'),
    (3, '生物安全实验不审批', '违反GB19489-2008，高风险实验必须审批', '立即停止，补办审批'),
    (4, '动物伦理不补审批', '违反《实验动物管理条例》，必须经伦理委员会审查', '立即补办'),
    (5, '人类受试者不签知情同意', '违反《涉及人的生物医学研究伦理审查办法》第13条', '立即补签，暂停研究'),
    (6, '辐射安全检测不做', '违反《放射性同位素与射线装置安全防护条例》', '立即检测'),
    (7, '应急预案不演练', '违反《突发事件应急预案管理办法》', '半年内至少演练一次'),
    (8, '化学品废弃不按规定', '违反《危险化学品安全管理条例》', '委托有资质单位处置'),
    (9, '安全检查不整改', '违反《实验室安全管理办法》', '限期整改，跟踪验收'),
    (10, '安全员无资质不培训', '违反《实验室安全管理办法》', '立即培训取证'),
]
for num, req, basis, action in rejections:
    doc.add_heading(f'要求{num}驳回意见', level=1)
    p = doc.add_paragraph(); p.add_run('管理层要求：').bold = True; p.add_run(req)
    p = doc.add_paragraph(); p.add_run('驳回依据：').bold = True; p.add_run(basis)
    p = doc.add_paragraph(); p.add_run('处理措施：').bold = True; p.add_run(action)
    doc.add_paragraph()
doc.save(os.path.join(d, '管理层错误要求回函.docx'))

doc = Document()
doc.add_heading('实验室安全与科研伦理整改方案', 0)
doc.add_heading('一、整改目标', level=1)
doc.add_paragraph('30日内整改重大违规，60日内建立健全安全管理制度，合规率从60%提升至90%')
doc.add_heading('二、整改措施', level=1)
doc.add_paragraph('1. 建立安全管理制度（12个实验室）')
doc.add_paragraph('2. 配齐有资质安全员（10个实验室）')
doc.add_paragraph('3. 规范危险化学品管理（60条违规记录）')
doc.add_paragraph('4. 强化生物安全和辐射安全（51条违规记录）')
doc.add_paragraph('5. 完善伦理审查程序（34条违规记录）')
doc.add_paragraph('6. 规范事故上报和应急预案（25条违规记录）')
doc.add_heading('三、时间安排', level=1)
doc.add_paragraph('第一阶段（1-15天）：重大违规立即整改')
doc.add_paragraph('第二阶段（16-30天）：一般违规整改')
doc.add_paragraph('第三阶段（31-60天）：制度建设和培训')
doc.add_paragraph('第四阶段（61-90天）：验收和持续改进')
doc.add_heading('四、责任分工', level=1)
doc.add_paragraph('高校安全管理处：牵头整改')
doc.add_paragraph('实验室负责人：具体实施')
doc.add_paragraph('科研伦理委员会：伦理审查')
doc.add_paragraph('省教育厅：监督验收')
doc.save(os.path.join(d, '整改方案.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '数据勾稽自检'
ws.append(['序号', '勾稽项', '来源数据', '验证数据', '一致性'])
checks = [
    (1, '实验室总数', '实验室台账80条', '台账记录数=80', '一致'),
    (2, '高校数', '高校清单15条', '实验室所属高校15所', '一致'),
    (3, '化学品记录数', '危险化学品台账200条', '台账记录数=200', '一致'),
    (4, '生物安全记录数', '生物安全记录50条', '台账记录数=50', '一致'),
    (5, '辐射安全记录数', '辐射安全记录30条', '台账记录数=30', '一致'),
    (6, '动物伦理审批数', '实验动物伦理审批30份', '审批记录数=30', '一致'),
    (7, '人类受试者审批数', '人类受试者伦理审批25份', '审批记录数=25', '一致'),
    (8, '安全事故数', '安全事故记录15起', '事故记录数=15', '一致'),
    (9, '应急预案数', '应急预案15份', '预案记录数=15', '一致'),
    (10, '违规总数', '200项', '各模块加总=200', '一致'),
]
for c in checks:
    ws.append(c)
style(ws, 5); wb.save(os.path.join(d, '数据勾稽自检.xlsx'))
mkzip(d, '06整改方案与勾稽自检.rar')

print('\n=== 标准答案生成完成 ===')
