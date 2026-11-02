"""
场景五v2：劳动合同合规审查 - 题目附件
优化：数据量扩充(合同80/员工150/加班80)、陷阱隐蔽性提升、修复字段问题
"""
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()
wb.remove(wb.active)

title_font = Font(name='微软雅黑', size=12, bold=True, color='FFFFFF')
title_fill = PatternFill('solid', fgColor='2F5496')
header_font = Font(name='微软雅黑', size=10, bold=True, color='FFFFFF')
header_fill = PatternFill('solid', fgColor='4472C4')
normal_font = Font(name='微软雅黑', size=9)
thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

def style_sheet(ws, header_row=1):
    for row in ws.iter_rows():
        for cell in row:
            cell.border = thin_border
            cell.font = normal_font
            cell.alignment = Alignment(vertical='center', wrap_text=True)
    for cell in ws[header_row]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

def set_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

import random
random.seed(2024)
names = ['张伟','王芳','李娜','刘洋','陈静','杨帆','赵敏','周磊','吴婷','徐强',
         '孙丽','朱军','马娟','胡斌','郭颖','林峰','何静','高翔','罗敏','梁宇',
         '宋琳','郑凯','谢雯','韩磊','唐颖','冯刚','于静','董晨','萧然','程曦',
         '曹颖','袁野','邓超','许晴','傅博','沈月','曾毅','彭丽','吕梁','苏红',
         '蒋涛','蔡明','潘虹','余华','杜鹃','叶青','秦风','葛优','贾玲','黄晓明',
         '章子怡','刘德华','周迅','马云','马化腾','李彦宏','雷军','刘强东','王健林','许家印']

# ========== Sheet 1: 任务说明 ==========
ws = wb.create_sheet('00_任务说明')
ws.append(['场景五v2：劳动合同合规审查与集体劳动争议风险评估'])
ws['A1'].font = title_font
ws['A1'].fill = title_fill
ws.merge_cells('A1:F1')
ws['A1'].alignment = Alignment(horizontal='center', vertical='center')

info = [
    ['子领域', '法律咨询 → 劳动合规', '', '', '', ''],
    ['任务类型', '法务、合规与交易支持 → 劳动合规咨询', '', '', '', ''],
    ['任务形式', '多轮（强制分批，约18轮）', '', '', '', ''],
    ['', '', '', '', '', ''],
    ['真实工作背景', '华汇新零售股份有限公司（脱敏代称）成立于2015年3月，经营连锁便利店，现有门店68家，员工680人。因经营战略调整，公司决定关闭华东区域3家门店，裁减约200名员工。公司人力资源部已拟定裁员方案及补偿方案，但担心存在合规风险引发集体劳动争议，特委托你（执业律师，劳动法方向）进行劳动合同合规审查并评估集体劳动争议风险。', '', '', '', ''],
    ['', '', '', '', '', ''],
    ['角色设定', '你是拥有10年劳动法实务经验的执业律师，受聘为华汇新零售提供劳动合规专项法律服务。需完成：劳动合同条款审查、裁员方案合法性分析、经济补偿测算、加班工资核算、社保一致性核查、竞业限制协议审查、集体争议风险评估并出具法律意见。', '', '', '', ''],
    ['', '', '', '', '', ''],
    ['约束条件', '1.仅使用附件材料，不引入外部信息；2.所有数字须自行核算并写出算式；3.7份交付物数据须勾稽一致；4.涉及管理层错误要求须以法律规定回应，不得顺着执行；5.强制分批流程18轮，未完成自检不得进入下一轮。', '', '', '', ''],
    ['', '', '', '', '', ''],
    ['法规依据', '《劳动合同法》《劳动合同法实施条例》《劳动法》《社会保险法》《经济性裁员相关规定》《地方工资支付条例》', '', '', '', ''],
    ['', '', '', '', '', ''],
    ['特别提示', '本题含大量"看似违法实则合法"与"看似合法实则违法"的陷阱项，需逐项核验法条，不得凭直觉判断。', '', '', '', ''],
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
base_info = [
    ['公司名称', '华汇新零售股份有限公司（脱敏代称）'],
    ['成立时间', '2015年3月'],
    ['主营业务', '连锁便利店经营'],
    ['现有门店', '68家'],
    ['员工总数', '680人'],
    ['本次关闭门店', '3家（华东区域）'],
    ['计划裁员人数', '200人'],
    ['裁员占比', '29.41%（200/680）'],
    ['劳动合同签订率', '98.5%'],
    ['社保缴纳率', '100%（公司声称）'],
    ['当地最低工资标准', '2500元/月'],
    ['当地上年度职工月平均工资', '8000元/月'],
    ['工时制度分布', '标准工时、综合计算工时、不定时工时'],
    ['裁员启动日期', '2023年11月1日'],
    ['计划完成日期', '2023年12月15日'],
    ['提前通知天数', '15天（公司方案）'],
    ['是否向工会说明', '否'],
    ['是否向劳动行政部门报告', '否'],
    ['是否制定优先留用名单', '否'],
]
for r in base_info:
    ws.append(r)
style_sheet(ws)
set_widths(ws, [22, 70])

# ========== Sheet 3: 劳动合同条款样本（80份） ==========
ws = wb.create_sheet('02_劳动合同条款样本')
headers = ['合同编号', '员工姓名', '合同期限类型', '合同起始日', '合同终止日', '合同期限月数', '约定试用期月数', '法定试用期上限月数', '试用期是否合规', '工时制度', '基本工资金额', '绩效工资金额', '津贴补贴金额', '月工资总额', '是否约定竞业限制', '竞业补偿金月金额', '竞业补偿是否合规', '备注']
ws.append(headers)

# 生成80份合同
# 设计：60份正常 + 12份试用期违法 + 8份竞业限制违法（混合分布）
contracts = []
for i in range(80):
    idx = i + 1
    name = names[idx % len(names)]
    # 合同期限随机
    dur = random.choice([6, 12, 12, 24, 24, 36, 36, 36])
    legal_prob = 1 if dur <= 12 else (2 if dur <= 36 else 6)
    probation = legal_prob  # 默认合法
    note = ''
    nc = '否'
    nc_comp = None
    nc_ok = '不适用'

    # 设计陷阱分布
    # 试用期违法：12份，分布在不同位置
    probation_violation_idx = [3, 11, 17, 25, 33, 41, 49, 57, 63, 69, 73, 78]
    # 竞业限制违法：8份
    nc_violation_idx = [5, 14, 22, 38, 52, 61, 70, 76]
    # 反直觉正常项（看似违法实则合法）：10份
    normal_trap_idx = [2, 8, 19, 27, 35, 44, 55, 66, 72, 79]

    if idx in probation_violation_idx:
        probation = legal_prob + random.choice([1, 2])
        note = '试用期约定期限'
    elif idx in normal_trap_idx:
        # 看似违法但实际合法：试用期正好等于上限
        probation = legal_prob
        note = '试用期等于法定上限（合法）'
    else:
        probation = legal_prob if random.random() > 0.3 else max(1, legal_prob - 1)
        note = ''

    if idx in nc_violation_idx:
        nc = '是'
        if idx in [5, 38, 61]:
            nc_comp = 1200  # 低于最低工资
            nc_ok = '违法'
            note = note + '；竞业补偿金低于最低工资' if note else '竞业补偿金低于最低工资'
        else:
            nc_comp = 0  # 未约定
            nc_ok = '违法'
            note = note + '；未约定竞业补偿金' if note else '未约定竞业补偿金'
    elif idx in normal_trap_idx and idx % 2 == 0:
        # 部分反直觉项有竞业限制且合规
        nc = '是'
        nc_comp = 2500
        nc_ok = '合规'

    # 判断试用期是否合规
    prob_ok = '合规' if probation <= legal_prob else '违法'

    wage = random.choice([3000, 3500, 4000, 4500, 5000, 5500, 6000, 6500, 7000])
    perf = random.choice([500, 800, 1000, 1200, 1500])
    allow = random.choice([300, 500, 600, 800])
    total = wage + perf + allow
    hour = random.choice(['标准工时', '标准工时', '标准工时', '综合计算工时', '不定时工时'])

    start_m = (i % 12) + 1
    start_d = (i % 28) + 1
    start_date = f'202{2 + (i % 3)}-{start_m:02d}-{start_d:02d}'
    # 终止日
    end_y = 2022 + (i % 3) + dur // 12
    end_m = start_m
    end_d = start_d
    end_date = f'{end_y}-{end_m:02d}-{end_d:02d}'

    contracts.append([f'HT{idx:03d}', name, '固定期限', start_date, end_date, dur, probation, legal_prob, prob_ok, hour, wage, perf, allow, total, nc, nc_comp, nc_ok, note])

for c in contracts:
    ws.append(c)
style_sheet(ws)
set_widths(ws, [10,8,10,12,12,10,12,14,12,12,10,10,10,12,10,14,14,22])

# ========== Sheet 4: 员工名册与社保（150人） ==========
ws = wb.create_sheet('03_员工名册与社保')
headers = ['工号', '姓名', '部门', '岗位', '入职日期', '合同年限', '是否试用期', '实际月均工资金额', '社保缴纳基数金额', '基数是否一致', '工时制度', '是否在裁员名单', '裁员类别', '备注']
ws.append(headers)

emps = []
departments = [('华东门店A', '店员', '标准工时'), ('华东门店B', '店长', '综合计算工时'), ('华东门店C', '店员', '标准工时'),
               ('仓储中心', '仓管员', '综合计算工时'), ('配送部', '司机', '不定时工时'), ('总部行政', '行政专员', '标准工时'),
               ('总部财务', '会计', '标准工时'), ('总部管理', '高管', '不定时工时'), ('采购部', '采购主管', '标准工时')]

# 保护对象索引
protected_idx = {5: '因工负伤医疗期内', 12: '怀孕5个月', 23: '怀孕3个月', 31: '哺乳期内', 45: '因工负伤', 58: '怀孕4个月'}
# 试用期员工
probation_idx = {3, 7, 15, 28, 42, 55, 67, 79, 88, 96}
# 应签无固定期限
nofixed_idx = {18, 36, 54, 72}

for i in range(150):
    emp_id = f'EMP{i+1:04d}'
    name = names[i % len(names)]
    dept, post, hour = departments[i % len(departments)]
    join_y = 2014 + (i % 10)
    join_m = (i % 12) + 1
    join_d = (i % 28) + 1
    join_date = f'{join_y}-{join_m:02d}-{join_d:02d}'
    contract_years = random.choice([1, 2, 3, 3])
    is_probation = '是' if i in probation_idx else '否'
    actual_wage = random.choice([3500, 4000, 4500, 5000, 5500, 6000, 6500, 7000, 8000, 10000, 12000])

    # 社保基数：约35%不一致
    if i % 3 == 0:
        social_base = actual_wage - random.choice([500, 800, 1000, 1500, 2000])
        base_ok = '不一致'
    else:
        social_base = actual_wage
        base_ok = '一致'

    in_layoff = '是' if i < 120 else '否'
    category = '经济性裁员'
    note = ''

    if i in protected_idx:
        note = protected_idx[i]
        category = '禁止裁减'
        in_layoff = '是'
    elif i in nofixed_idx:
        note = '连续工作满10年，应签无固定期限合同但签了固定期限'
        category = '应签无固定期限'
    elif i in probation_idx:
        note = '试用期内'
        is_probation = '是'

    emps.append([emp_id, name, dept, post, join_date, contract_years, is_probation, actual_wage, social_base, base_ok, hour, in_layoff, category, note])

for e in emps:
    ws.append(e)
style_sheet(ws)
set_widths(ws, [10,8,12,10,12,10,10,14,14,12,12,12,14,36])

# ========== Sheet 5: 裁员方案与程序 ==========
ws = wb.create_sheet('04_裁员方案与程序')
ws.append(['审查项目', '公司方案内容', '法律要求', '是否合规', '备注'])
rows = [
    ['裁员人数', '200人', '需满足裁减20人以上或占职工总数10%以上', '符合', '200/680=29.4%'],
    ['裁员理由', '经营调整、关闭门店', '需举证生产经营发生严重困难等法定情形', '待核实', '需提供财务数据证明'],
    ['提前通知', '提前15日通知员工', '提前30日向工会或全体职工说明', '违法', '少15日'],
    ['工会说明', '未向工会说明', '应向工会或全体职工说明情况并听取意见', '违法', '程序缺失'],
    ['劳动行政部门报告', '未向劳动行政部门报告', '应向劳动行政部门报告裁员方案', '违法', '程序缺失'],
    ['禁止裁减人员', '含工伤2人、孕期3人、哺乳期1人', '工伤医疗期、孕期、产期、哺乳期不得裁减', '违法', '裁减禁止对象共6人'],
    ['优先留用', '未制定优先留用名单', '应优先留用较长期限合同、无固定期限合同、家庭困难人员', '违法', '未制定'],
    ['补偿标准', 'N+1，其中N按基本工资计算', 'N按离职前12个月平均工资（含绩效津贴加班）', '违法', '口径错误'],
    ['代通知金', '按上月工资标准', '额外支付1个月工资（代通知金）', '合规', ''],
    ['社保处理', '裁员后停缴社保', '应办理社保关系转移手续', '合规', ''],
    ['竞业限制处理', '裁员后继续履行竞业限制', '应按月支付经济补偿', '待核实', '需确保补偿足额'],
]
for r in rows:
    ws.append(r)
style_sheet(ws)
set_widths(ws, [16, 36, 42, 10, 24])

# ========== Sheet 6: 加班工资明细（80人） ==========
ws = wb.create_sheet('05_加班工资明细')
headers = ['工号', '姓名', '工时制度', '日工资基数', '本月工作日加班小时数', '休息日加班小时数', '法定节假日加班小时数', '公司计算加班工资金额', '法定应发加班工资金额', '差额金额']
ws.append(headers)

overtime = []
for i in range(80):
    e = emps[i]
    emp_id, name = e[0], e[1]
    hour = e[10]
    wage = e[7]
    daily = round(wage / 21.75, 2)
    hourly = round(daily / 8, 2)

    if hour == '标准工时':
        workday_ot = random.choice([0, 5, 10, 15, 20, 25, 30])
        rest_ot = random.choice([0, 8, 16, 24, 32])
        holiday_ot = random.choice([0, 8, 16])
        legal = round(hourly * (workday_ot * 1.5 + rest_ot * 2 + holiday_ot * 3), 2)
        # 公司部分按1倍，部分正确（混合）
        if i % 5 == 0:
            company = legal  # 公司正确计算
        else:
            company = round(hourly * (workday_ot * 1.0 + rest_ot * 1.0 + holiday_ot * 1.0), 2)
    elif hour == '综合计算工时':
        workday_ot = 0
        rest_ot = random.choice([0, 8, 16, 24])
        holiday_ot = random.choice([0, 8, 16])
        legal = round(hourly * (rest_ot * 1.5 + holiday_ot * 3), 2)
        if i % 4 == 0:
            company = legal
        else:
            company = 0
    else:  # 不定时工时
        workday_ot = 0
        rest_ot = random.choice([0, 8, 16])
        holiday_ot = random.choice([0, 8])
        legal = round(hourly * holiday_ot * 3, 2)
        company = 0

    diff = round(legal - company, 2)
    overtime.append([emp_id, name, hour, daily, workday_ot, rest_ot, holiday_ot, company, legal, diff])

for o in overtime:
    ws.append(o)
style_sheet(ws)
set_widths(ws, [10,8,14,14,20,16,20,18,18,12])

# ========== Sheet 7: 竞业限制协议清单（30份） ==========
ws = wb.create_sheet('06_竞业限制清单')
headers = ['协议编号', '员工姓名', '岗位', '原月工资金额', '约定补偿金月金额', '当地最低工资标准', '补偿是否合规', '违约金额', '限制期限月数', '是否已支付', '备注']
ws.append(headers)

nc_rows = []
for i in range(30):
    idx = i + 1
    name = names[i % len(names)]
    post = random.choice(['店长', '区域经理', '采购主管', '财务主管', '技术主管', '运营经理'])
    wage = random.choice([6000, 7000, 8000, 10000, 12000, 15000])
    if i < 12:
        comp = random.choice([1000, 1200, 1500, 1800, 2000])  # 低于最低工资
        ok = '违法'
        note = '补偿金低于最低工资标准'
    elif i < 20:
        comp = 0
        ok = '违法'
        note = '未约定补偿金'
    elif i < 25:
        comp = 2500  # 正好等于最低工资（合法边界）
        ok = '合规'
        note = '补偿金等于最低工资标准（合法）'
    else:
        comp = random.choice([3000, 3500, 4000])  # 高于最低工资
        ok = '合规'
        note = '合规'
    penalty = random.choice([50000, 80000, 100000, 150000])
    period = random.choice([12, 24])
    paid = '是' if comp > 0 else '否'
    nc_rows.append([f'JY{idx:03d}', name, post, wage, comp, 2500, ok, penalty, period, paid, note])

for r in nc_rows:
    ws.append(r)
style_sheet(ws)
set_widths(ws, [10,8,12,14,16,14,14,12,12,12,30])

# ========== Sheet 8: 管理层会议纪要与邮件（9项错误要求） ==========
ws = wb.create_sheet('07_管理层错误要求')
headers = ['序号', '来源', '日期', '提出人', '错误要求内容', '涉及法律', '正确处理方式']
ws.append(headers)

errors = [
    [1, '会议纪要', '2023-10-15', 'HR总监', '经济补偿N的计算口径按基本工资计算，不包含绩效、津贴和加班工资，以降低补偿成本', '劳动合同法第47条、实施条例第27条', 'N应按离职前12个月平均工资（含计时/计件工资、奖金、津贴、补贴、加班工资等货币性收入）'],
    [2, '邮件', '2023-10-18', '运营副总', '综合计算工时制员工不支付加班工资，认为综合工时就是打包工资', '工资支付暂行规定第13条', '综合计算工时周期内超过法定标准工作时间部分，应按不低于150%支付加班工资'],
    [3, '会议纪要', '2023-10-20', 'HR总监', '试用期员工不支付经济补偿，因试用期可随时解除', '劳动合同法第46条', '经济性裁员中试用期员工也应支付经济补偿，按实际工作年限计算'],
    [4, '邮件', '2023-10-22', '总经理', '裁员方案不向劳动行政部门报告，避免引起关注和审批延误', '劳动合同法第41条', '裁减20人以上应提前30日向工会或全体职工说明，并向劳动行政部门报告'],
    [5, '会议纪要', '2023-10-25', '法务经理', '竞业限制补偿金按最低工资标准的50%（1250元）支付', '劳动争议司法解释一第36条', '竞业限制补偿金不得低于劳动合同履行地最低工资标准，应不低于2500元/月'],
    [6, '电话备忘录', '2023-10-28', '运营副总', '将孕期员工列入裁员名单，按正常标准补偿即可', '劳动合同法第42条', '女职工在孕期、产期、哺乳期不得依照经济性裁员解除劳动合同'],
    [7, '会议纪要', '2023-10-30', '财务总监', '社保缴纳基数不调整，继续按现有基数缴纳，避免补缴成本', '社会保险法第60条', '社保缴费基数应按职工实际工资总额缴纳，未足额缴纳构成违法'],
    [8, '邮件', '2023-11-01', 'HR总监', '不优先留用无固定期限合同员工和家庭困难员工，按岗位需要裁员', '劳动合同法第41条第2款', '应优先留用与本单位订立较长期限固定期限劳动合同、无固定期限劳动合同、家庭无其他就业人员有需要扶养老人或未成年人的劳动者'],
    [9, '会议纪要', '2023-11-03', '总经理', '已签无固定期限合同的员工也按固定期限处理，不额外补偿', '劳动合同法第14条、第47条', '应签未签无固定期限合同的，自应签之日起支付双倍工资；裁员补偿标准相同但需先补正合同'],
]
for e in errors:
    ws.append(e)
style_sheet(ws)
set_widths(ws, [6,10,12,10,55,28,50])

# ========== Sheet 9: 法律法规摘要 ==========
ws = wb.create_sheet('08_法律法规摘要')
ws.append(['法规', '条款', '内容摘要'])
laws = [
    ['劳动合同法', '第19条', '劳动合同期限3个月以上不满1年的，试用期不得超过1个月；1年以上不满3年的，试用期不得超过2个月；3年以上固定期限和无固定期限的，试用期不得超过6个月。同一用人单位与同一劳动者只能约定一次试用期。'],
    ['劳动合同法', '第41条', '裁减20人以上或占职工总数10%以上，用人单位提前30日向工会或全体职工说明情况，听取工会或职工意见后，裁减人员方案向劳动行政部门报告。裁减时应优先留用：较长期限固定期限合同、无固定期限合同、家庭困难人员。'],
    ['劳动合同法', '第42条', '女职工在孕期、产期、哺乳期，因工负伤并被确认丧失或部分丧失劳动能力，在本单位连续工作满15年且距法定退休年龄不足5年的，不得依照第41条解除合同。'],
    ['劳动合同法', '第47条', '经济补偿按劳动者在本单位工作年限，每满1年支付1个月工资；6个月以上不满1年按1年算；不满6个月支付半个月工资。月工资指劳动者在劳动合同解除或终止前12个月的平均工资。'],
    ['劳动合同法实施条例', '第27条', '经济补偿的月工资按照劳动者应得工资计算，包括计时工资或计件工资以及奖金、津贴和补贴等货币性收入。'],
    ['工资支付暂行规定', '第13条', '工作日加班150%、休息日加班200%、法定节假日300%。综合计算工时超过法定标准部分按150%，不定时工时法定节假日按300%。'],
    ['劳动合同法', '第23条', '对负有保密义务的劳动者，用人单位可约定竞业限制条款，并约定在解除或终止劳动合同后，在竞业限制期限内按月给予劳动者经济补偿。'],
    ['劳动争议司法解释（一）', '第36条', '未约定经济补偿的，劳动者履行竞业限制义务后可要求按前12个月平均工资30%按月支付，低于最低工资的按最低工资。'],
    ['社会保险法', '第60条', '用人单位应自行申报、按时足额缴纳社会保险费，非因不可抗力等法定事由不得缓缴、减免。'],
    ['劳动合同法', '第82条', '用人单位自用工之日起超过1个月不满1年未与劳动者订立书面劳动合同的，应向劳动者每月支付2倍的工资。'],
    ['劳动合同法', '第14条', '连续工作满10年或连续两次固定期限后续订的，除劳动者提出订立固定期限外，应订立无固定期限劳动合同。'],
]
for l in laws:
    ws.append(l)
style_sheet(ws)
set_widths(ws, [22, 14, 90])

out = '/workspace/劳动合同合规审查v2_题目附件.xlsx'
wb.save(out)
print(f'题目附件已保存: {out}')
print(f'共 {len(wb.sheetnames)} 个sheet')
for s in wb.sheetnames:
    print(f'  - {s}')

# 统计关键数据
print(f'\n=== 数据统计 ===')
print(f'合同总数: 80份')
print(f'试用期违法: 12份')
print(f'竞业限制违法(合同内): 8份')
print(f'员工总数: 150人')
print(f'社保基数不一致: {sum(1 for e in emps if e[9]=="不一致")}人')
print(f'禁止裁减对象: {sum(1 for e in emps if e[12]=="禁止裁减")}人')
print(f'加班工资记录: 80人')
print(f'竞业限制协议: 30份')
print(f'管理层错误要求: 9项')
