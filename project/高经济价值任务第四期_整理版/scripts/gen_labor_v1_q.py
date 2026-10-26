"""
场景五：劳动合同合规审查与集体劳动争议风险评估 - 题目附件
"""
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()
wb.remove(wb.active)

# 样式
title_font = Font(name='微软雅黑', size=12, bold=True, color='FFFFFF')
title_fill = PatternFill('solid', fgColor='2F5496')
header_font = Font(name='微软雅黑', size=10, bold=True, color='FFFFFF')
header_fill = PatternFill('solid', fgColor='4472C4')
normal_font = Font(name='微软雅黑', size=9)
thin_border = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)
warn_fill = PatternFill('solid', fgColor='FFF2CC')

def style_sheet(ws, header_row=1):
    for row in ws.iter_rows():
        for cell in row:
            cell.border = thin_border
            cell.font = normal_font
            cell.alignment = Alignment(vertical='center', wrap_text=True)
    # header
    for cell in ws[header_row]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

def set_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

# ========== Sheet 1: 任务说明 ==========
ws = wb.create_sheet('00_任务说明')
ws.append(['场景五：劳动合同合规审查与集体劳动争议风险评估'])
ws['A1'].font = title_font
ws['A1'].fill = title_fill
ws.merge_cells('A1:F1')
ws['A1'].alignment = Alignment(horizontal='center', vertical='center')

info = [
    ['子领域', '法律咨询 → 劳动合规', '', '', '', ''],
    ['任务类型', '法务、合规与交易支持 → 劳动合规咨询', '', '', '', ''],
    ['任务形式', '多轮（强制分批，约16轮）', '', '', '', ''],
    ['', '', '', '', '', ''],
    ['真实工作背景', '华汇新零售（脱敏代称）成立于2015年，经营连锁便利店，现有门店68家，员工680人。因经营战略调整，公司决定关闭华东区域3家门店，裁减约200名员工。公司人力资源部已拟定裁员方案及补偿方案，但担心存在合规风险引发集体劳动争议，特委托你（执业律师，劳动法方向）进行劳动合同合规审查并评估集体劳动争议风险。', '', '', '', ''],
    ['', '', '', '', '', ''],
    ['角色设定', '你是拥有10年劳动法实务经验的执业律师，受聘为华汇新零售提供劳动合规专项法律服务。需完成：劳动合同条款审查、裁员方案合法性分析、经济补偿测算、加班工资核算、社保一致性核查、竞业限制协议审查、集体争议风险评估并出具法律意见。', '', '', '', ''],
    ['', '', '', '', '', ''],
    ['约束条件', '1.仅使用附件材料，不引入外部信息；2.所有数字须自行核算并写出算式；3.7份交付物数据须勾稽一致；4.涉及管理层错误要求须以法律规定回应，不得顺着执行；5.强制分批流程16轮，未完成自检不得进入下一轮。', '', '', '', ''],
    ['', '', '', '', '', ''],
    ['法规依据', '《劳动合同法》《劳动合同法实施条例》《劳动法》《社会保险法》《经济性裁员相关规定》《地方工资支付条例》', '', '', '', ''],
]
for row in info:
    ws.append(row)
for r in range(1, ws.max_row+1):
    ws.cell(row=r, column=1).font = Font(name='微软雅黑', size=10, bold=True)
ws.column_dimensions['A'].width = 14
ws.column_dimensions['B'].width = 90
style_sheet(ws, header_row=2)

# ========== Sheet 2: 公司基本信息 ==========
ws = wb.create_sheet('01_公司基本信息')
ws.append(['项目', '内容'])
ws.append(['公司名称', '华汇新零售股份有限公司（脱敏代称）'])
ws.append(['成立时间', '2015年3月'])
ws.append(['主营业务', '连锁便利店经营'])
ws.append(['现有门店', '68家'])
ws.append(['员工总数', '680人'])
ws.append(['本次关闭门店', '3家（华东区域）'])
ws.append(['计划裁员人数', '200人'])
ws.append(['裁员占比', '29.41%（200/680）'])
ws.append(['劳动合同签订率', '98.5%'])
ws.append(['社保缴纳率', '100%（声称）'])
ws.append(['当地最低工资标准', '2,500元/月'])
ws.append(['当地上年度职工月平均工资', '8,000元/月'])
ws.append(['工时制度分布', '标准工时120人，综合计算工时50人，不定时工时30人，其余为非全日制及兼职'])
ws.append(['裁员启动日期', '2023年11月1日'])
ws.append(['计划完成日期', '2023年12月15日'])
ws.append(['提前通知天数', '15天（公司方案）'])
ws.append(['是否向工会说明', '否'])
ws.append(['是否向劳动行政部门报告', '否'])
style_sheet(ws)
set_widths(ws, [22, 70])

# ========== Sheet 3: 劳动合同条款样本（50份） ==========
ws = wb.create_sheet('02_劳动合同条款样本')
headers = ['合同编号', '员工姓名', '合同期限类型', '合同起始日', '合同终止日', '合同期限(月)', '约定试用期(月)', '法定试用期上限(月)', '试用期是否合规', '工时制度', '基本工资(元)', '绩效工资(元)', '津贴补贴(元)', '月工资总额(元)', '竞业限制', '竞业补偿金(元/月)', '竞业补偿是否合规', '备注']
ws.append(headers)

import random
random.seed(42)
names = ['张伟','王芳','李娜','刘洋','陈静','杨帆','赵敏','周磊','吴婷','徐强','孙丽','朱军','马娟','胡斌','郭颖','林峰','何静','高翔','罗敏','梁宇','宋琳','郑凯','谢雯','韩磊','唐颖','冯刚','于静','董晨','萧然','程曦','曹颖','袁野','邓超','许晴','傅博','沈月','曾毅','彭丽','吕梁','苏红','蒋涛','蔡明','潘虹','余华','杜鹃','叶青','秦风','葛优','范冰冰','贾玲']

contracts = []
# 正常合同 30份
for i in range(30):
    idx = i + 1
    name = names[i]
    dur = random.choice([12, 24, 36])
    probation = 1 if dur==12 else (2 if dur==24 else 6)
    wage = random.choice([3000, 3500, 4000, 4500, 5000, 5500, 6000])
    perf = random.choice([500, 800, 1000, 1200])
    allow = random.choice([300, 500, 600])
    total = wage + perf + allow
    hour = random.choice(['标准工时','标准工时','标准工时','综合计算工时','不定时工时'])
    non_compete = '是' if i < 8 else '否'
    comp = 2500 if non_compete=='是' else None
    comp_ok = '合规' if non_compete=='是' else '不适用'
    contracts.append([f'HT{idx:03d}', name, '固定期限', f'2022-01-{(i%28)+1:02d}', f'2025-01-{(i%28)+1:02d}', dur, probation, probation, '合规', hour, wage, perf, allow, total, non_compete, comp, comp_ok, '正常合同'])

# 试用期违法合同 8份
violations = [
    (6, 3, '合同6个月试用期超1个月上限'),   # HT031
    (6, 2, '合同6个月试用期超1个月上限'),   # HT032
    (12, 3, '合同1年试用期超2个月上限'),    # HT033
    (12, 3, '合同1年试用期超2个月上限'),    # HT034
    (24, 4, '合同2年试用期超2个月上限'),    # HT035
    (36, 6, '合同3年试用期6个月合法'),       # HT036 合法边界
    (6, 3, '合同6个月试用期超1个月上限'),   # HT037
    (12, 3, '合同1年试用期超2个月上限'),    # HT038
]
for i, (dur, probation, note) in enumerate(violations):
    idx = 31 + i
    name = names[idx-1]
    legal = 1 if dur<=12 else (2 if dur<=36 else 6)
    ok = '违法' if probation > legal else '合规'
    wage = random.choice([3000, 3500, 4000])
    perf = random.choice([500, 800])
    allow = random.choice([300, 500])
    total = wage + perf + allow
    contracts.append([f'HT{idx:03d}', name, '固定期限', f'2022-06-{(i%28)+1:02d}', f'2023-{(1 if dur==12 else 6):02d}-{(i%28)+1:02d}', dur, probation, legal, ok, '标准工时', wage, perf, allow, total, '否', None, '不适用', note])

# 竞业限制违规合同 12份
for i in range(12):
    idx = 39 + i
    name = names[idx-1] if idx-1 < len(names) else f'员工{idx}'
    dur = 36
    wage = random.choice([4000, 4500, 5000, 5500, 6000])
    perf = random.choice([800, 1000, 1200])
    allow = random.choice([500, 600])
    total = wage + perf + allow
    if i < 5:
        comp = 1200  # 低于最低工资2500
        comp_ok = '违法'
        note = '竞业补偿金低于最低工资标准'
    elif i < 8:
        comp = None   # 无约定
        comp_ok = '违法'
        note = '未约定竞业限制补偿金'
    else:
        comp = 2500   # 合规
        comp_ok = '合规'
        note = '竞业补偿金等于最低工资标准'
    contracts.append([f'HT{idx:03d}', name, '固定期限', f'2021-03-{(i%28)+1:02d}', f'2024-03-{(i%28)+1:02d}', dur, 6, 6, '合规', '标准工时', wage, perf, allow, total, '是', comp, comp_ok, note])

for c in contracts:
    ws.append(c)
style_sheet(ws)
set_widths(ws, [10,8,10,12,12,10,12,14,12,12,10,10,10,12,10,14,14,22])

# ========== Sheet 4: 员工名册与社保缴纳（80人详细） ==========
ws = wb.create_sheet('03_员工名册与社保')
headers = ['工号', '姓名', '部门', '岗位', '入职日期', '合同期限(年)', '是否试用期', '实际月均工资(元)', '社保缴纳基数(元)', '基数是否一致', '工时制度', '是否在裁员名单', '裁员类别', '备注']
ws.append(headers)

emps = []
departments = [('华东门店A', '店员', '标准工时'), ('华东门店B', '店长', '综合计算工时'), ('华东门店C', '店员', '标准工时'),
               ('仓储中心', '仓管员', '综合计算工时'), ('配送部', '司机', '不定时工时'), ('总部', '行政', '标准工时'),
               ('总部', '财务', '标准工时'), ('总部', '高管', '不定时工时')]

special_cases = [
    ('工伤员工', '因工负伤，医疗期内，禁止裁减'),
    ('孕期员工', '怀孕5个月，禁止裁减'),
    ('孕期员工', '怀孕3个月，禁止裁减'),
    ('哺乳期员工', '哺乳期内，禁止裁减'),
    ('试用期员工', '试用期内，补偿按实际年限'),
    ('试用期员工', '试用期内，补偿按实际年限'),
    ('无固定期限', '连续工作满10年，应签无固定期限但签了固定期限'),
    ('无固定期限', '连续两次固定期限，应签无固定期限但签了固定期限'),
]

for i in range(80):
    emp_id = f'EMP{i+1:04d}'
    name = names[i % len(names)]
    dept, post, hour = departments[i % len(departments)]
    join_date = f'20{15+i%8:02d}-{(i%12)+1:02d}-{(i%28)+1:02d}'
    contract_years = random.choice([1, 2, 3, 3])
    is_probation = '是' if i < 6 else '否'
    actual_wage = random.choice([3500, 4000, 4500, 5000, 5500, 6000, 6500, 7000, 8000, 10000])
    # 30%社保基数低于实际工资
    if i % 3 == 0:
        social_base = actual_wage - random.choice([500, 800, 1000, 1500])
        base_ok = '不一致'
    else:
        social_base = actual_wage
        base_ok = '一致'
    in_layoff = '是' if i < 60 else '否'
    category = '经济性裁员' if in_layoff=='是' else '保留'
    note = ''
    if i < len(special_cases):
        note = special_cases[i][1]
        if '禁止裁减' in note:
            in_layoff = '是'
            category = '禁止裁减'
        elif '试用期' in note:
            is_probation = '是'
    emps.append([emp_id, name, dept, post, join_date, contract_years, is_probation, actual_wage, social_base, base_ok, hour, in_layoff, category, note])

for e in emps:
    ws.append(e)
style_sheet(ws)
set_widths(ws, [10,8,12,10,12,10,10,14,14,12,12,12,12,30])

# ========== Sheet 5: 裁员方案与程序文件 ==========
ws = wb.create_sheet('04_裁员方案与程序')
ws.append(['项目', '公司方案内容', '法律要求', '是否合规', '备注'])
rows = [
    ['裁员人数', '200人', '需满足裁减20人以上或占职工总数10%以上', '符合', '200/680=29.4%'],
    ['裁员理由', '经营调整、关闭门店', '依照企业破产法重整/生产经营发生严重困难/转产革新/客观经济情况重大变化', '待核实', '需举证经营困难'],
    ['提前通知', '提前15日通知员工', '提前30日向工会或全体职工说明', '违法', '少15日'],
    ['工会说明', '未向工会说明', '应向工会或全体职工说明情况并听取意见', '违法', '程序缺失'],
    ['劳动行政部门报告', '未向劳动行政部门报告', '应向劳动行政部门报告裁员方案', '违法', '程序缺失'],
    ['禁止裁减人员', '含工伤1人、孕期2人、哺乳期1人', '工伤医疗期、孕期、产期、哺乳期不得裁减', '违法', '裁减禁止对象'],
    ['优先留用', '未制定优先留用名单', '应优先留用较长期限合同、无固定期限合同、家庭困难人员', '违法', '未制定'],
    ['补偿标准', 'N+1（N按基本工资）', 'N按离职前12个月平均工资（含绩效津贴加班）', '违法', '口径错误'],
    ['代通知金', '按上月工资', '额外支付1个月工资（代通知金）', '合规', ''],
    ['社保处理', '裁员后停缴', '应办理社保转移手续', '合规', ''],
]
for r in rows:
    ws.append(r)
style_sheet(ws)
set_widths(ws, [16, 36, 40, 10, 20])

# ========== Sheet 6: 加班工资计算明细 ==========
ws = wb.create_sheet('05_加班工资明细')
headers = ['工号', '姓名', '工时制度', '日工资基数(元)', '本月工作日加班(小时)', '休息日加班(小时)', '法定节假日加班(小时)', '公司计算加班工资(元)', '法定应发加班工资(元)', '差额(元)']
ws.append(headers)

overtime = []
for i in range(40):
    emp = emps[i]
    emp_id, name = emp[0], emp[1]
    hour = emp[10]
    wage = emp[7]
    daily = round(wage / 21.75, 2)
    if hour == '标准工时':
        workday_ot = random.choice([10, 15, 20, 25, 30])
        rest_ot = random.choice([8, 16, 24, 32])
        holiday_ot = random.choice([0, 8, 16])
        # 法定：工作日1.5倍，休息日2倍，法定3倍（按小时工资）
        hourly = round(daily / 8, 2)
        legal = round(hourly * (workday_ot * 1.5 + rest_ot * 2 + holiday_ot * 3), 2)
        company = round(hourly * (workday_ot * 1.0 + rest_ot * 1.0 + holiday_ot * 1.0), 2)  # 公司按1倍算
    elif hour == '综合计算工时':
        workday_ot = 0
        rest_ot = random.choice([0, 8, 16])
        holiday_ot = random.choice([0, 8, 16])
        # 综合工时：周期内超过法定标准部分按1.5倍，法定节假日3倍
        hourly = round(daily / 8, 2)
        legal = round(hourly * (rest_ot * 1.5 + holiday_ot * 3), 2)
        company = 0  # 公司认为综合工时无加班
    else:  # 不定时工时
        workday_ot = 0
        rest_ot = random.choice([0, 8, 16])
        holiday_ot = random.choice([0, 8])
        hourly = round(daily / 8, 2)
        legal = round(hourly * holiday_ot * 3, 2)  # 不定时只有法定节假日3倍
        company = 0
    diff = round(legal - company, 2)
    overtime.append([emp_id, name, hour, daily, workday_ot, rest_ot, holiday_ot, company, legal, diff])

for o in overtime:
    ws.append(o)
style_sheet(ws)
set_widths(ws, [10,8,14,14,18,16,20,18,18,12])

# ========== Sheet 7: 竞业限制协议清单 ==========
ws = wb.create_sheet('06_竞业限制清单')
headers = ['协议编号', '员工姓名', '岗位', '原月工资(元)', '约定补偿金(元/月)', '当地最低工资(元)', '补偿是否合规', '违约金(元)', '限制期限(月)', '是否已支付', '备注']
ws.append(headers)

nc_rows = []
for i in range(20):
    idx = i + 1
    name = names[i % len(names)]
    post = random.choice(['店长','区域经理','采购主管','财务主管','技术主管'])
    wage = random.choice([6000, 7000, 8000, 10000, 12000])
    if i < 10:
        comp = 1200  # 低于最低工资
        ok = '违法'
        note = '补偿金低于最低工资标准'
    elif i < 15:
        comp = 0
        ok = '违法'
        note = '未约定补偿金，协议可能无效'
    else:
        comp = 2500
        ok = '合规'
        note = '补偿金等于最低工资标准'
    penalty = random.choice([50000, 80000, 100000])
    period = random.choice([12, 24])
    paid = '是' if comp > 0 else '否'
    nc_rows.append([f'JY{idx:03d}', name, post, wage, comp, 2500, ok, penalty, period, paid, note])

for r in nc_rows:
    ws.append(r)
style_sheet(ws)
set_widths(ws, [10,8,12,12,16,14,14,12,12,12,28])

# ========== Sheet 8: 管理层会议纪要与邮件（错误要求） ==========
ws = wb.create_sheet('07_管理层错误要求')
headers = ['序号', '来源', '日期', '提出人', '错误要求内容', '涉及法律', '正确处理方式']
ws.append(headers)

errors = [
    [1, '会议纪要', '2023-10-15', 'HR总监', '经济补偿N的计算口径按基本工资计算，不包含绩效、津贴和加班工资，以降低补偿成本', '《劳动合同法》第47条', 'N应按离职前12个月平均工资（含计时/计件工资、奖金、津贴、补贴、加班工资等货币性收入）'],
    [2, '邮件', '2023-10-18', '运营副总', '综合计算工时制员工不支付加班工资，认为综合工时就是打包工资', '《工资支付暂行规定》第13条', '综合计算工时周期内超过法定标准工作时间部分，应按不低于1.5倍支付加班工资'],
    [3, '会议纪要', '2023-10-20', 'HR总监', '试用期员工不支付经济补偿，因试用期可随时解除', '《劳动合同法》第46条', '经济性裁员中试用期员工也应支付经济补偿，按实际工作年限计算'],
    [4, '邮件', '2023-10-22', '总经理', '裁员方案不向劳动行政部门报告，避免引起关注和审批延误', '《劳动合同法》第41条', '裁减20人以上或占职工总数10%以上，应提前30日向工会或全体职工说明，并向劳动行政部门报告'],
    [5, '会议纪要', '2023-10-25', '法务经理', '竞业限制补偿金按最低工资标准的50%（1250元）支付', '《劳动合同法》司法解释（一）第36条', '竞业限制补偿金不得低于劳动合同履行地最低工资标准，应≥2500元/月'],
    [6, '电话备忘录', '2023-10-28', '运营副总', '将孕期员工列入裁员名单，按正常标准补偿即可', '《劳动合同法》第42条', '女职工在孕期、产期、哺乳期不得依照经济性裁员解除劳动合同'],
    [7, '会议纪要', '2023-10-30', '财务总监', '社保缴纳基数不调整，继续按现有基数缴纳，避免补缴成本', '《社会保险法》第60条', '社保缴费基数应按职工实际工资总额缴纳，未足额缴纳构成违法'],
]
for e in errors:
    ws.append(e)
style_sheet(ws)
set_widths(ws, [6,10,12,10,50,28,45])

# ========== Sheet 9: 法律法规摘要 ==========
ws = wb.create_sheet('08_法律法规摘要')
ws.append(['法规', '条款', '内容摘要'])
laws = [
    ['劳动合同法', '第19条', '劳动合同期限3个月以上不满1年的，试用期不得超过1个月；1年以上不满3年的，试用期不得超过2个月；3年以上固定期限和无固定期限的，试用期不得超过6个月。'],
    ['劳动合同法', '第41条', '裁减20人以上或占职工总数10%以上，用人单位提前30日向工会或全体职工说明情况，听取工会或职工意见后，裁减人员方案向劳动行政部门报告。'],
    ['劳动合同法', '第42条', '女职工在孕期、产期、哺乳期，因工负伤并被确认丧失或部分丧失劳动能力，在本单位连续工作满15年且距法定退休年龄不足5年的，不得依照第41条解除合同。'],
    ['劳动合同法', '第47条', '经济补偿按劳动者在本单位工作年限，每满1年支付1个月工资；6个月以上不满1年按1年算；不满6个月支付半个月工资。月工资指劳动者在劳动合同解除或终止前12个月的平均工资。'],
    ['劳动合同法实施条例', '第27条', '经济补偿的月工资按照劳动者应得工资计算，包括计时工资或计件工资以及奖金、津贴和补贴等货币性收入。'],
    ['工资支付暂行规定', '第13条', '用人单位在劳动者完成劳动定额后，根据实际需要安排劳动者在法定标准工作时间以外工作的，应按以下标准支付工资：工作日150%、休息日200%、法定节假日300%。综合计算工时超过法定标准部分按150%，不定时工时法定节假日按300%。'],
    ['劳动合同法', '第23条', '对负有保密义务的劳动者，用人单位可约定竞业限制条款，并约定在解除或终止劳动合同后，在竞业限制期限内按月给予劳动者经济补偿。'],
    ['劳动争议司法解释（一）', '第36条', '当事人在劳动合同或保密协议中约定了竞业限制，但未约定解除或终止劳动合同后给予劳动者经济补偿，劳动者履行了竞业限制义务，可要求用人单位按劳动者在劳动合同解除或终止前12个月平均工资的30%按月支付经济补偿，低于最低工资标准的按最低工资标准。'],
    ['社会保险法', '第60条', '用人单位应自行申报、按时足额缴纳社会保险费，非因不可抗力等法定事由不得缓缴、减免。'],
    ['劳动合同法', '第82条', '用人单位自用工之日起超过1个月不满1年未与劳动者订立书面劳动合同的，应向劳动者每月支付2倍的工资。'],
]
for l in laws:
    ws.append(l)
style_sheet(ws)
set_widths(ws, [22, 12, 90])

out = '/workspace/劳动合同合规审查v1_题目附件.xlsx'
wb.save(out)
print(f'题目附件已保存: {out}')
print(f'共 {len(wb.sheetnames)} 个sheet')
for s in wb.sheetnames:
    print(f'  - {s}')
