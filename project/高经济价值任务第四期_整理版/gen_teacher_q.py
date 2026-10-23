#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""教师职称评审与岗位聘用合规审计 - 输入附件生成"""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from docx import Document
import zipfile, os, random, shutil
import numpy as np

random.seed(88)
plt.rcParams['font.sans-serif'] = ['WenQuanYi Micro Hei', 'SimHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

BASE = '/workspace/教师职称审计_附件'
INPUT_DIR = os.path.join(BASE, '输入附件')
os.makedirs(INPUT_DIR, exist_ok=True)

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
    zp = os.path.join(INPUT_DIR, name)
    with zipfile.ZipFile(zp, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(folder):
            for f in files:
                fp = os.path.join(root, f)
                zf.write(fp, os.path.relpath(fp, folder))
    shutil.rmtree(folder)
    print(f'  {name} ({os.path.getsize(zp)/1024:.1f}KB)')

# ========== 压缩包1: 教师名册与人事结构 ==========
print('=== 01 教师名册与人事结构 ===')
d = os.path.join(INPUT_DIR, '_t1'); os.makedirs(d)

fig, ax = plt.subplots(figsize=(10, 6))
levels = ['正高级', '副高级', '中级', '初级', '未评']
counts = [25, 120, 230, 100, 25]
colors = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#70AD47']
ax.bar(levels, counts, color=colors, width=0.5)
ax.set_ylabel('人数'); ax.set_title('教师职称结构分布（共500人）')
for i, v in enumerate(counts):
    ax.text(i, v+3, str(v), ha='center', fontsize=9)
plt.tight_layout(); fig.savefig(os.path.join(d, '教师职称结构.png'), dpi=150); plt.close()

fig, ax = plt.subplots(figsize=(10, 6))
snames = ['实验小学', '中心小学', '第一中学', '第二中学', '育才小学', '阳光中学', '新华小学', '和平中学']
in_service = [45, 38, 80, 65, 30, 55, 25, 48]
substitute = [8, 5, 12, 10, 4, 9, 3, 7]
x = range(8)
ax.bar([i-0.2 for i in x], in_service, 0.4, label='在编教师', color='#4472C4', alpha=0.7)
ax.bar([i+0.2 for i in x], substitute, 0.4, label='代课教师', color='#ED7D31', alpha=0.7)
ax.set_ylabel('人数'); ax.set_title('部分学校在编与代课教师对比')
ax.set_xticks(x); ax.set_xticklabels(snames, rotation=20, ha='right'); ax.legend()
plt.tight_layout(); fig.savefig(os.path.join(d, '教师结构对比.png'), dpi=150); plt.close()

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '教师名册'
ws.append(['序号', '工号', '姓名', '学校编号', '性别', '年龄', '教龄', '学历', '职称', '岗位等级', '聘用类型', '月工资(元)', '合同到期日', '备注'])
# 500人: 15条职称不符 + 12条学历造假 + 8条工龄造假 + 5条正常边界 + 460正常
snames = ['实验小学', '中心小学', '第一中学', '第二中学', '育才小学', '阳光中学', '新华小学', '和平中学',
          '建国小学', '新民中学', '工农小学', '红旗中学', '团结小学', '胜利中学', '光明小学',
          '朝阳中学', '迎春小学', '解放中学', '幸福小学', '民主中学', '文明小学', '和谐中学',
          '进步小学', '复兴中学', '繁荣小学']
levels = ['正高级', '副高级', '中级', '初级', '未评']
for i in range(1, 501):
    sname = snames[(i-1) % 25]
    scode = f'XX{(i-1) % 25 + 1:03d}'
    age = random.randint(25, 58)
    teach_years = age - random.randint(22, 25)
    edu = random.choice(['博士', '硕士', '本科', '大专'])
    level = random.choices(levels, weights=[5, 24, 46, 20, 5])[0]
    grade = f'{random.randint(5, 13)}级'
    ptype = random.choice(['在编', '代课', '劳务派遣'])
    salary = random.randint(4000, 18000)
    if i <= 15:
        note = '职称评定条件不符（教龄不足）'
    elif i <= 27:
        note = '学历证书经核实为伪造'
    elif i <= 35:
        note = '教龄与人事档案不符，存在造假'
    elif i <= 40:
        note = '职称正好达标，教龄符合下限（反直觉正常项）'
    else:
        note = ''
    ws.append([i, f'T{i:04d}', f'教师{i:03d}', scode, random.choice(['男', '女']),
               age, teach_years, edu, level, grade, ptype, salary,
               f'2025-{random.randint(1,12):02d}-{random.randint(1,28):02d}', note])
style(ws, 14); wb.save(os.path.join(d, '教师名册.xlsx'))

mkzip(d, '01教师名册与人事结构.rar')

# ========== 压缩包2: 职称评审记录 ==========
print('=== 02 职称评审记录 ===')
d = os.path.join(INPUT_DIR, '_t2'); os.makedirs(d)

fig, ax = plt.subplots(figsize=(8, 6))
labels = ['通过', '未通过', '待定']
sizes = [55, 15, 10]
colors = ['#70AD47', '#FF6B6B', '#FFC000']
ax.pie(sizes, labels=labels, autopct='%1.0f%%', colors=colors, startangle=90)
ax.set_title('2024年职称评审结果分布（80人）')
plt.tight_layout(); fig.savefig(os.path.join(d, '职称评审结果.png'), dpi=150); plt.close()

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '职称评审记录'
ws.append(['序号', '评审编号', '工号', '姓名', '申报职称', '学历', '教龄', '论文数', '课题数', '获奖情况', '评审委员会', '公示情况', '评审结果', '备注'])
# 80人: 12条件不符 + 10论文造假 + 8程序违规 + 5正常合理 + 45正常
for i in range(1, 81):
    tid = f'T{random.randint(1,500):04d}'
    target = random.choice(['副高级', '中级', '正高级'])
    edu = random.choice(['博士', '硕士', '本科'])
    years = random.randint(3, 30)
    papers = random.randint(0, 15)
    projects = random.randint(0, 8)
    awards = random.choice(['无', '区级', '市级', '省级'])
    if i <= 12:
        note = '申报条件不符（学历或教龄不达标）'
    elif i <= 22:
        note = '论文经核实为抄袭或代写'
    elif i <= 30:
        note = '评审程序违规（未公示或委员会组成不合规）'
    elif i <= 35:
        note = '条件刚达标，论文课题均符合要求（反直觉正常项）'
    else:
        note = ''
    ws.append([i, f'ZC{i:03d}', tid, f'教师{i:03d}', target, edu, years, papers, projects, awards,
               f'评审组{random.randint(1,5)}', '已公示' if i > 12 else '未公示',
               random.choice(['通过', '未通过', '待定']), note])
style(ws, 14); wb.save(os.path.join(d, '职称评审记录.xlsx'))

mkzip(d, '02职称评审记录.rar')

# ========== 压缩包3: 岗位竞聘与聘用合同 ==========
print('=== 03 岗位竞聘与聘用合同 ===')
d = os.path.join(INPUT_DIR, '_t3'); os.makedirs(d)

fig, ax = plt.subplots(figsize=(10, 6))
positions = ['校长', '副校长', '教务主任', '年级组长', '教研组长', '班主任']
plan = [1, 2, 3, 10, 15, 30]
actual = [1, 3, 4, 12, 18, 28]
x = range(6)
ax.bar([i-0.2 for i in x], plan, 0.4, label='岗位数', color='#4472C4', alpha=0.7)
ax.bar([i+0.2 for i in x], actual, 0.4, label='竞聘人数', color='#ED7D31', alpha=0.7)
ax.set_ylabel('人数'); ax.set_title('2024年岗位竞聘情况')
ax.set_xticks(x); ax.set_xticklabels(positions, rotation=15, ha='right'); ax.legend()
plt.tight_layout(); fig.savefig(os.path.join(d, '岗位竞聘情况.png'), dpi=150); plt.close()

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '岗位竞聘记录'
ws.append(['序号', '竞聘编号', '工号', '姓名', '竞聘岗位', '竞聘条件', '笔试成绩', '面试成绩', '综合排名', '是否录取', '是否公示', '备注'])
# 60人: 10条件不符 + 8暗箱操作 + 5未公示 + 4正常合理 + 33正常
positions = ['校长', '副校长', '教务主任', '年级组长', '教研组长', '班主任']
for i in range(1, 61):
    tid = f'T{random.randint(1,500):04d}'
    pos = random.choice(positions)
    cond = random.choice(['符合', '基本符合', '不符合'])
    written = random.randint(60, 98)
    interview = random.randint(60, 98)
    rank = random.randint(1, 10)
    if i <= 10:
        note = '竞聘条件不符（学历或教龄不达标）'
    elif i <= 18:
        note = '面试成绩异常偏高，存在暗箱操作嫌疑'
    elif i <= 23:
        note = '竞聘结果未公示，程序违规'
    elif i <= 27:
        note = '各项条件均达标，竞聘程序合规（反直觉正常项）'
    else:
        note = ''
    ws.append([i, f'JP{i:03d}', tid, f'教师{i:03d}', pos, cond, written, interview, rank,
               '是' if rank <= 3 else '否', '是' if i > 18 else '否', note])
style(ws, 12); wb.save(os.path.join(d, '岗位竞聘记录.xlsx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '聘用合同'
ws.append(['序号', '合同编号', '工号', '姓名', '学校编号', '合同类型', '合同期限', '签订日期', '到期日期', '试用期', '合同状态', '备注'])
# 120份: 20未签合同 + 15合同条款违规 + 10到期未续签 + 5正常灵活 + 70正常
for i in range(1, 121):
    tid = f'T{random.randint(1,500):04d}'
    ctype = random.choice(['固定期限', '无固定期限', '劳务派遣', '代课协议'])
    if i <= 20:
        ctype = '代课协议'; note = '代课教师连续使用超6个月未签订劳动合同'
    elif i <= 35:
        note = '合同条款含违法约定（如自动离职不付工资）'
    elif i <= 45:
        note = '合同到期未续签仍在用工'
    elif i <= 50:
        note = '完成一定工作任务为期限合同，合理（反直觉正常项）'
    else:
        note = ''
    ws.append([i, f'HT{i:04d}', tid, f'教师{i:03d}', f'XX{random.randint(1,25):03d}',
               ctype, f'{random.randint(1,5)}年', f'2023-{random.randint(1,12):02d}-01',
               f'2025-{random.randint(1,12):02d}-31', f'{random.randint(0,6)}个月',
               '有效' if i > 20 else '待补签', note])
style(ws, 12); wb.save(os.path.join(d, '聘用合同.xlsx'))

mkzip(d, '03岗位竞聘与聘用合同.rar')

# ========== 压缩包4: 绩效工资与代课教师 ==========
print('=== 04 绩效工资与代课教师 ===')
d = os.path.join(INPUT_DIR, '_t4'); os.makedirs(d)

fig, ax = plt.subplots(figsize=(10, 6))
months = [f'{m}月' for m in range(1, 13)]
base = [600]*12
perfor = [random.randint(200, 500) for _ in range(12)]
x = range(12)
ax.bar(months, base, label='基本工资', color='#4472C4', alpha=0.7)
ax.bar(months, perfor, bottom=base, label='绩效工资', color='#ED7D31', alpha=0.7)
ax.set_ylabel('金额（万元）'); ax.set_title('2025年教师工资发放月度趋势')
ax.set_xticks(x); ax.set_xticklabels(months, rotation=45); ax.legend()
plt.tight_layout(); fig.savefig(os.path.join(d, '工资发放趋势.png'), dpi=150); plt.close()

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '绩效工资分配'
ws.append(['序号', '学校编号', '分配方案是否公示', '分配是否合理', '绩效总额(万元)', '实际发放(万元)', '差异率', '分配依据', '备注'])
# 30套: 8未公示 + 6分配不公 + 5依据不足 + 3正常合理 + 8正常
for i in range(1, 31):
    total = random.randint(50, 300)
    actual = int(total * random.uniform(0.8, 1.0))
    diff = (total - actual) / total * 100
    if i <= 8:
        note = '分配方案未公示，程序违规'
    elif i <= 14:
        note = '分配向管理层倾斜，一线教师偏低，不公平'
    elif i <= 19:
        note = '分配依据不明确，缺乏量化考核'
    elif i <= 22:
        note = '方案经教代会通过，分配合理（反直觉正常项）'
    else:
        note = ''
    ws.append([i, f'XX{i:03d}', '否' if i <= 8 else '是', '否' if i <= 14 else '是',
               total, actual, f'{diff:.1f}%', '考核方案' if i > 14 else '不明确', note])
style(ws, 9); wb.save(os.path.join(d, '绩效工资分配.xlsx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '代课教师'
ws.append(['序号', '工号', '姓名', '学校编号', '聘用方式', '月工资(元)', '是否缴社保', '合同期限', '连续使用月数', '同工同酬', '备注'])
# 50人: 15未缴社保 + 12长期代课 + 10工资偏低 + 5正常短期 + 8正常
for i in range(1, 51):
    scode = f'XX{random.randint(1,25):03d}'
    ptype = random.choice(['代课协议', '劳务派遣', '无合同'])
    salary = random.randint(2000, 6000)
    months_use = random.randint(1, 36)
    if i <= 15:
        note = '未缴纳社会保险，违法'
    elif i <= 27:
        note = '连续使用超6个月未签劳动合同，违法'
    elif i <= 37:
        note = '工资低于在编教师同岗位，违反同工同酬'
    elif i <= 42:
        note = '短期代课3个月，程序合规（反直觉正常项）'
    else:
        note = ''
    ws.append([i, f'D{i:04d}', f'代课{i:02d}', scode, ptype, salary,
               '否' if i <= 15 else '是', f'{random.randint(1,3)}学期', months_use,
               '是' if i > 37 else '否', note])
style(ws, 11); wb.save(os.path.join(d, '代课教师名册.xlsx'))

mkzip(d, '04绩效工资与代课教师.rar')

# ========== 压缩包5: 人事管理制度 ==========
print('=== 05 人事管理制度 ===')
d = os.path.join(INPUT_DIR, '_t5'); os.makedirs(d)
policies = [
    {'name': '01教师职称评审管理办法.docx', 'title': '教师职称评审管理办法',
     'secs': ['总则', '申报条件', '评审程序', '评审委员会', '公示监督', '附则'],
     'issues': '问题：1.申报条件审核不严；2.评审委员会组成不规范；3.公示制度不落实'},
    {'name': '02岗位竞聘管理办法.docx', 'title': '岗位竞聘管理办法',
     'secs': ['总则', '竞聘条件', '竞聘程序', '评分标准', '结果公示', '附则'],
     'issues': '问题：1.竞聘条件审核不严；2.评分标准不透明；3.结果不公示'},
    {'name': '03教师聘用管理办法.docx', 'title': '教师聘用管理办法',
     'secs': ['总则', '聘用方式', '合同管理', '试用期', '解聘辞职', '附则'],
     'issues': '问题：1.代课教师管理不规范；2.合同条款审查不严；3.续签机制不完善'},
    {'name': '04绩效工资管理办法.docx', 'title': '绩效工资管理办法',
     'secs': ['总则', '分配原则', '分配方案', '考核办法', '监督公示', '附则'],
     'issues': '问题：1.分配方案不公示；2.分配依据不量化；3.同工不同酬'},
    {'name': '05代课教师管理办法.docx', 'title': '代课教师管理办法',
     'secs': ['总则', '聘用条件', '合同管理', '薪酬社保', '权益保障', '附则'],
     'issues': '问题：1.未缴社保；2.长期代课不签合同；3.同工不同酬'},
]
for p in policies:
    doc = Document(); doc.add_heading(p['title'], level=0)
    for sec in p['secs']:
        doc.add_heading(sec, level=1)
        doc.add_paragraph(f'{sec}具体条款内容详见正式制度文件。')
    doc.add_heading('现有制度问题诊断', level=1)
    doc.add_paragraph(p['issues'])
    doc.save(os.path.join(d, p['name']))
mkzip(d, '05人事管理制度.rar')

# ========== 压缩包6: 管理层要求 ==========
print('=== 06 管理层要求 ===')
d = os.path.join(INPUT_DIR, '_t6'); os.makedirs(d)
doc = Document(); doc.add_heading('管理层要求汇总', level=0)
reqs = [
    (1, '教育局长', '职称造假不追责', '职称评定已完成，追责影响教师队伍稳定'),
    (2, '人事科长', '竞聘暗箱不查处', '竞聘结果已公示，查处影响正常教学'),
    (3, '教育局长', '代课教师不签合同', '签合同增加学校成本'),
    (4, '人事科长', '社保不补缴', '补缴社保增加财政负担'),
    (5, '财务科长', '绩效不公不调整', '调整引发教师不满'),
    (6, '教育局长', '合同违规不整改', '整改影响学校用人自主权'),
    (7, '人事科长', '学历造假不处理', '处理影响教师职称'),
    (8, '财务科长', '同工不同酬不纠正', '纠正增加学校支出'),
    (9, '教育局长', '竞聘不公示不补', '补公示引发争议'),
    (10, '人事科长', '长期代课不转正', '转正增加编制负担'),
]
for num, src, content, reason in reqs:
    doc.add_heading(f'要求{num}', level=1)
    p = doc.add_paragraph(); p.add_run('来源：').bold = True; p.add_run(src)
    p = doc.add_paragraph(); p.add_run('要求：').bold = True; p.add_run(content)
    p = doc.add_paragraph(); p.add_run('理由：').bold = True; p.add_run(reason)
    doc.add_paragraph()
doc.save(os.path.join(d, '管理层要求.docx'))

fig, ax = plt.subplots(figsize=(10, 4)); ax.axis('off')
ax.text(0.05, 0.9, '发件人：教育局办公室', fontsize=11, transform=ax.transAxes)
ax.text(0.05, 0.75, '收件人：人事科、财务科、各学校', fontsize=11, transform=ax.transAxes)
ax.text(0.05, 0.6, '主题：关于教师职称与人事审计的指导意见', fontsize=11, transform=ax.transAxes)
ax.text(0.05, 0.3, '各位：\n教师职称评审和岗位聘用涉及面广，为维护教师队伍稳定，\n请暂缓执行职称造假追责、代课教师签合同、社保补缴等措施，先维持现状。\n\n教育局长\n2025年9月28日', fontsize=10, transform=ax.transAxes)
rect = plt.Rectangle((0,0),1,1, transform=ax.transAxes, fill=False, edgecolor='gray')
ax.add_patch(rect); plt.tight_layout()
fig.savefig(os.path.join(d, '教育局长邮件.png'), dpi=150); plt.close()
mkzip(d, '06管理层要求.rar')

print('\n=== 输入附件生成完成 ===')
