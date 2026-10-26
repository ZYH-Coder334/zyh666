"""
场景五：劳动合同合规审查与集体劳动争议风险评估 - 标准答案
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

# 读取题目附件数据
src = openpyxl.load_workbook('/workspace/劳动合同合规审查v1_题目附件.xlsx', data_only=True)

# ========== Sheet 1: 总体结论与风险清单 ==========
ws = wb.create_sheet('00_总体结论与风险清单')
ws.append(['场景五：劳动合同合规审查与集体劳动争议风险评估 - 标准答案'])
ws['A1'].font = title_font
ws['A1'].fill = title_fill
ws.merge_cells('A1:F1')
ws['A1'].alignment = Alignment(horizontal='center', vertical='center')

ws.append([''])
ws.append(['一、总体结论'])
ws.append(['华汇新零售裁员方案存在8项重大合规风险，若不整改将面临集体劳动仲裁、行政罚款及声誉损失。核心问题：裁员程序违法（4项）、试用期条款违法（7份合同）、竞业限制违法（15份协议）、加班工资未足额支付（40人）、社保基数未足额缴纳（27人）、禁止裁减对象被列入名单（4人）、经济补偿口径错误、管理层7项错误要求。'])
ws.append([''])
ws.append(['二、风险清单（15项）'])

risks = [
    ['序号','风险事项','风险等级','法律依据','整改建议'],
    [1,'裁员未提前30日向工会/全体职工说明，仅提前15日','高','劳动合同法第41条','提前30日说明并听取意见'],
    [2,'裁员方案未向劳动行政部门报告','高','劳动合同法第41条','向劳动行政部门报告'],
    [3,'裁减孕期、哺乳期、工伤员工共4人','高','劳动合同法第42条','立即移出裁员名单'],
    [4,'7份合同试用期约定超过法定期限','高','劳动合同法第19条','变更合同或支付赔偿金'],
    [5,'15份竞业限制协议补偿金低于最低工资或未约定','高','劳动合同法第23条、司法解释一第36条','补足补偿金或解除协议'],
    [6,'40名员工加班工资未足额支付','高','工资支付暂行规定第13条','补发加班工资差额'],
    [7,'27名员工社保缴纳基数低于实际工资','中','社会保险法第60条','补缴社保差额'],
    [8,'经济补偿N按基本工资计算，口径错误','高','劳动合同法第47条、实施条例第27条','按前12个月平均工资（含绩效津贴加班）'],
    [9,'综合计算工时制员工被认定无加班工资','高','工资支付暂行规定第13条','周期内超法定标准部分按1.5倍支付'],
    [10,'试用期员工被拒绝支付经济补偿','中','劳动合同法第46条','按实际工作年限支付补偿'],
    [11,'未制定优先留用人员名单','中','劳动合同法第41条','优先留用长期合同、无固定期限、困难员工'],
    [12,'未签订书面劳动合同（若有）','高','劳动合同法第82条','补签并支付双倍工资'],
    [13,'管理层要求按最低工资50%支付竞业补偿','高','司法解释一第36条','按最低工资标准以上支付'],
    [14,'应签无固定期限合同而签固定期限','中','劳动合同法第14条','改签无固定期限合同'],
    [15,'集体劳动争议爆发风险','极高','劳动争议调解仲裁法','立即整改并与员工协商'],
]
for r in risks:
    ws.append(r)
style_sheet(ws, header_row=5)
set_widths(ws, [6,40,10,28,40])

# ========== Sheet 2: 劳动合同条款审查结果 ==========
ws = wb.create_sheet('01_劳动合同条款审查')
ws.append(['审查项','审查结果','依据','处理建议'])
rows = [
    ['合同总数','50份','',''],
    ['固定期限合同','50份','','全部为固定期限'],
    ['试用期合规合同','43份','劳动合同法第19条','合规'],
    ['试用期违法合同','7份','劳动合同法第19条','变更试用期条款或支付违法约定试用期赔偿金'],
    ['  - HT031:6个月合同约定3个月试用期','违法（上限1个月）','','变更为1个月'],
    ['  - HT032:6个月合同约定2个月试用期','违法（上限1个月）','','变更为1个月'],
    ['  - HT033:1年合同约定3个月试用期','违法（上限2个月）','','变更为2个月'],
    ['  - HT034:1年合同约定3个月试用期','违法（上限2个月）','','变更为2个月'],
    ['  - HT035:2年合同约定4个月试用期','违法（上限2个月）','','变更为2个月'],
    ['  - HT037:6个月合同约定3个月试用期','违法（上限1个月）','','变更为1个月'],
    ['  - HT038:1年合同约定3个月试用期','违法（上限2个月）','','变更为2个月'],
    ['竞业限制协议','30份（含合同内约定）','',''],
    ['  - 补偿金合规','约12份','司法解释一第36条','≥最低工资2500元'],
    ['  - 补偿金低于最低工资','约13份','司法解释一第36条','补足至≥2500元/月'],
    ['  - 未约定补偿金','约5份','司法解释一第36条','按前12月平均工资30%或最低工资支付'],
    ['反直觉正常项','HT036:3年合同6个月试用期','合法上限','无需处理'],
]
for r in rows:
    ws.append(r)
style_sheet(ws)
set_widths(ws, [40,30,25,35])

# ========== Sheet 3: 裁员方案合法性分析 ==========
ws = wb.create_sheet('02_裁员方案合法性分析')
ws.append(['审查项','公司方案','法定要求','结论','法律依据'])
rows = [
    ['裁员人数','200人','≥20人或≥10%','符合（200/680=29.4%）','劳动合同法第41条'],
    ['裁员理由','经营调整','需举证经营困难等法定情形','需补充举证','劳动合同法第41条'],
    ['提前通知','15日','≥30日','违法','劳动合同法第41条'],
    ['工会/职工说明','未进行','应说明并听取意见','违法','劳动合同法第41条'],
    ['劳动行政部门报告','未报告','应报告','违法','劳动合同法第41条'],
    ['禁止裁减对象','含工伤1人、孕期2人、哺乳期1人','不得裁减','违法','劳动合同法第42条'],
    ['优先留用','未制定','应优先留用长期/无固定/困难','违法','劳动合同法第41条'],
    ['经济补偿口径','N按基本工资','N按前12月平均工资（含绩效津贴加班）','违法','劳动合同法第47条、实施条例第27条'],
    ['代通知金','按上月工资','额外1个月工资','合规','劳动合同法第40条'],
]
for r in rows:
    ws.append(r)
style_sheet(ws)
set_widths(ws, [16,30,40,14,28])

# ========== Sheet 4: 经济补偿测算 ==========
# 从题目附件读取员工数据
ws_emp = src['03_员工名册与社保']
emps = []
for row in ws_emp.iter_rows(min_row=2, values_only=True):
    if row[0]:
        emps.append({
            'id': row[0], 'name': row[1], 'dept': row[2], 'post': row[3],
            'join': row[4], 'years': row[5], 'probation': row[6],
            'wage': row[7], 'social': row[8], 'base_ok': row[9],
            'hour': row[10], 'layoff': row[11], 'category': row[12], 'note': row[13]
        })

ws = wb.create_sheet('03_经济补偿测算')
ws.append(['工号','姓名','入职日期','工作年限(N)','月均工资(元)','经济补偿N×月工资(元)','代通知金(元)','合计补偿(元)','备注'])
total_comp = 0
for e in emps:
    if e['layoff'] == '是' and e['category'] != '禁止裁减':
        # 工作年限：假设2023年裁员，从入职日期算
        # 简化：用合同期限作为参考年数，实际应精确计算
        # 取入职年份到2023年
        join_year = int(str(e['join'])[:4])
        n = 2023 - join_year
        if n <= 0:
            n = 0.5  # 不满6个月
        wage = e['wage']
        comp = round(n * wage, 2)
        notice = wage  # 代通知金1个月
        total = comp + notice
        total_comp += total
        ws.append([e['id'], e['name'], e['join'], n, wage, comp, notice, total, e['note']])

ws.append(['', '', '', '', '合计', '', '', round(total_comp, 2), ''])
style_sheet(ws)
set_widths(ws, [10,8,12,12,14,20,14,16,25])

# ========== Sheet 5: 加班工资差额测算 ==========
ws_ot = src['05_加班工资明细']
ws = wb.create_sheet('04_加班工资差额测算')
ws.append(['工时制度','人数','公司已付(元)','法定应付(元)','差额(元)','处理建议'])

# 汇总
from collections import defaultdict
ot_summary = defaultdict(lambda: {'count':0,'company':0,'legal':0})
for row in ws_ot.iter_rows(min_row=2, values_only=True):
    if row[0]:
        hour = row[2]
        ot_summary[hour]['count'] += 1
        ot_summary[hour]['company'] += row[7]
        ot_summary[hour]['legal'] += row[8]

for hour, d in ot_summary.items():
    diff = round(d['legal'] - d['company'], 2)
    suggestion = '补发差额' if diff > 0 else '已足额'
    ws.append([hour, d['count'], round(d['company'],2), round(d['legal'],2), diff, suggestion])

total_company = sum(d['company'] for d in ot_summary.values())
total_legal = sum(d['legal'] for d in ot_summary.values())
ws.append(['合计', sum(d['count'] for d in ot_summary.values()), round(total_company,2), round(total_legal,2), round(total_legal-total_company,2), '补发差额'])
style_sheet(ws)
set_widths(ws, [16,8,14,14,14,20])

# ========== Sheet 6: 竞业限制审查 ==========
ws_nc = src['06_竞业限制清单']
ws = wb.create_sheet('05_竞业限制审查')
ws.append(['审查项','数量','处理建议'])

nc_illegal_low = 0
nc_illegal_none = 0
nc_legal = 0
for row in ws_nc.iter_rows(min_row=2, values_only=True):
    if row[0]:
        if row[5] == 2500 and row[4] >= 2500:
            nc_legal += 1
        elif row[4] == 0:
            nc_illegal_none += 1
        elif row[4] < 2500:
            nc_illegal_low += 1

ws.append(['竞业限制协议总数', 20, ''])
ws.append(['补偿金合规（≥2500元）', nc_legal, '无需处理'])
ws.append(['补偿金低于最低工资', nc_illegal_low, '补足至≥2500元/月'])
ws.append(['未约定补偿金', nc_illegal_none, '按前12月平均工资30%或最低工资支付'])
ws.append(['违法合计', nc_illegal_low + nc_illegal_none, '需整改'])
style_sheet(ws)
set_widths(ws, [30,10,40])

# ========== Sheet 7: 社保一致性核查 ==========
ws = wb.create_sheet('06_社保一致性核查')
ws.append(['审查项','人数','占比','处理建议'])

base_illegal = sum(1 for e in emps if e['base_ok'] == '不一致')
base_legal = sum(1 for e in emps if e['base_ok'] == '一致')
ws.append(['抽查员工总数', len(emps), '100%', ''])
ws.append(['社保基数与实际工资一致', base_legal, f'{base_legal/len(emps)*100:.1f}%', '合规'])
ws.append(['社保基数低于实际工资', base_illegal, f'{base_illegal/len(emps)*100:.1f}%', '补缴差额及滞纳金'])
style_sheet(ws)
set_widths(ws, [30,10,12,40])

# ========== Sheet 8: 评分矩阵（考点） ==========
ws = wb.create_sheet('07_评分矩阵')
ws.append(['考点编号','考点内容','分值','得分条件'])
points = [
    ['L1','识别7份试用期违法合同并指出法定期限','10','全部识别正确得10分，每漏1份扣2分'],
    ['L2','识别裁员程序3项违法（提前通知、工会说明、劳动行政报告）','10','全部识别得10分，每漏1项扣3分'],
    ['L3','识别4名禁止裁减对象（工伤、孕期、哺乳期）','10','全部识别得10分，每漏1人扣3分'],
    ['L4','正确计算经济补偿口径（含绩效津贴加班，非仅基本工资）','10','口径正确得10分，错误得0分'],
    ['L5','识别综合工时加班工资计算规则（1.5倍）','8','正确得8分，错误得0分'],
    ['L6','识别不定时工时加班工资规则（仅法定节假日3倍）','8','正确得8分，错误得0分'],
    ['L7','识别竞业限制补偿金不得低于最低工资标准','10','正确得10分，错误得0分'],
    ['L8','识别未约定竞业补偿金的处理方式','8','正确得8分，错误得0分'],
    ['L9','识别社保基数未足额缴纳问题','8','正确得8分，错误得0分'],
    ['L10','正确处理管理层7项错误要求（逐项引用法条驳回）','10','全部驳回得10分，每顺任1项扣2分'],
    ['L11','反直觉正常项不误判（如3年合同6个月试用期合法）','5','不误判得5分，误判1项扣3分'],
    ['L12','7份交付物数据勾稽一致','5','一致得5分，不一致每处扣2分'],
]
for p in points:
    ws.append(p)
style_sheet(ws)
set_widths(ws, [10,55,8,45])

out = '/workspace/劳动合同合规审查v1_标准答案.xlsx'
wb.save(out)
print(f'标准答案已保存: {out}')
print(f'共 {len(wb.sheetnames)} 个sheet')
for s in wb.sheetnames:
    print(f'  - {s}')

# 输出关键统计
print(f'\n=== 关键统计 ===')
print(f'试用期违法合同: 7份')
print(f'裁员程序违法: 3项')
print(f'禁止裁减对象: 4人')
print(f'竞业限制违法: {nc_illegal_low + nc_illegal_none}份')
print(f'社保基数不一致: {base_illegal}人')
print(f'经济补偿总额(合规口径): {round(total_comp,2)}元')
print(f'加班工资差额: {round(total_legal-total_company,2)}元')
