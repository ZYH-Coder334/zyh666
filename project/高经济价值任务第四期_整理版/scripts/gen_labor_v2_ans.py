"""
场景五v2：劳动合同合规审查 - 标准答案
"""
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from collections import defaultdict

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

src = openpyxl.load_workbook('/workspace/劳动合同合规审查v2_题目附件.xlsx', data_only=True)

# 读取合同数据
ws_ct = src['02_劳动合同条款样本']
contracts = []
for row in ws_ct.iter_rows(min_row=2, values_only=True):
    if row[0]:
        contracts.append({
            'id': row[0], 'name': row[1], 'dur': row[5],
            'probation': row[6], 'legal_prob': row[7], 'prob_ok': row[8],
            'nc': row[14], 'nc_comp': row[15], 'nc_ok': row[16], 'note': row[17]
        })

# 读取员工数据
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

# 读取加班数据
ws_ot = src['05_加班工资明细']
overtime = []
for row in ws_ot.iter_rows(min_row=2, values_only=True):
    if row[0]:
        overtime.append({
            'id': row[0], 'name': row[1], 'hour': row[2],
            'daily': row[3], 'workday': row[4], 'rest': row[5],
            'holiday': row[6], 'company': row[7], 'legal': row[8], 'diff': row[9]
        })

# ========== Sheet 1: 总体结论与风险清单 ==========
ws = wb.create_sheet('00_总体结论与风险清单')
ws.append(['场景五v2：劳动合同合规审查与集体劳动争议风险评估 - 标准答案'])
ws['A1'].font = title_font
ws['A1'].fill = title_fill
ws.merge_cells('A1:F1')
ws['A1'].alignment = Alignment(horizontal='center', vertical='center')

ws.append([''])
ws.append(['一、总体结论'])
ws.append(['华汇新零售裁员方案存在10项重大合规风险：裁员程序3项违法、试用期12份合同违法、竞业限制23份违法（合同内8份+协议15份）、6名禁止裁减对象被列入名单、加班工资未足额支付、50人社保基数未足额缴纳、经济补偿口径错误、优先留用缺失、应签未签无固定期限、管理层9项错误要求。若不整改将面临集体劳动仲裁、行政罚款及声誉损失。'])
ws.append([''])
ws.append(['二、风险清单（15项）'])

risks = [
    ['序号', '风险事项', '风险等级', '法律依据', '整改建议'],
    [1, '裁员未提前30日向工会说明，仅提前15日', '高', '劳动合同法第41条', '提前30日说明并听取意见'],
    [2, '裁员方案未向劳动行政部门报告', '高', '劳动合同法第41条', '向劳动行政部门报告'],
    [3, '裁减工伤2人、孕期3人、哺乳期1人共6人', '高', '劳动合同法第42条', '立即移出裁员名单'],
    [4, '12份合同试用期超过法定期限', '高', '劳动合同法第19条', '变更合同或支付赔偿金'],
    [5, '23份竞业限制补偿金违法（低于最低工资或未约定）', '高', '司法解释一第36条', '补足补偿金或解除协议'],
    [6, '80人中加班工资未足额支付', '高', '工资支付暂行规定第13条', '补发加班工资差额'],
    [7, '50名员工社保缴纳基数低于实际工资', '中', '社会保险法第60条', '补缴社保差额'],
    [8, '经济补偿N按基本工资计算，口径错误', '高', '劳动合同法第47条、实施条例第27条', '按前12个月平均工资'],
    [9, '未制定优先留用人员名单', '中', '劳动合同法第41条第2款', '优先留用长期/无固定/困难员工'],
    [10, '应签未签无固定期限合同4人', '中', '劳动合同法第14条', '改签无固定期限并支付双倍工资'],
    [11, '管理层9项错误要求', '高', '多项法律', '逐项驳回并依法执行'],
    [12, '综合工时员工加班工资未支付', '高', '工资支付暂行规定第13条', '周期内超标准部分1.5倍'],
    [13, '不定时工时员工法定节假日加班未支付', '中', '工资支付暂行规定第13条', '法定节假日3倍'],
    [14, '试用期员工被拒付经济补偿', '中', '劳动合同法第46条', '按实际工作年限支付'],
    [15, '集体劳动争议爆发风险', '极高', '劳动争议调解仲裁法', '立即整改并协商'],
]
for r in risks:
    ws.append(r)
style_sheet(ws, header_row=5)
set_widths(ws, [6, 45, 10, 28, 40])

# ========== Sheet 2: 劳动合同条款审查 ==========
ws = wb.create_sheet('01_劳动合同条款审查')
ws.append(['审查项', '数量', '依据', '处理建议'])

prob_illegal = [c for c in contracts if c['prob_ok'] == '违法']
prob_legal = [c for c in contracts if c['prob_ok'] == '合规']
nc_illegal = [c for c in contracts if c['nc_ok'] == '违法']
nc_legal_nc = [c for c in contracts if c['nc_ok'] == '合规']

rows = [
    ['合同总数', 80, '', ''],
    ['试用期合规合同', len(prob_legal), '劳动合同法第19条', '合规'],
    ['试用期违法合同', len(prob_illegal), '劳动合同法第19条', '变更试用期条款或支付赔偿金'],
]
for c in prob_illegal:
    rows.append([f'  - {c["id"]}:合同{c["dur"]}个月试用期{c["probation"]}个月', f'违法(上限{c["legal_prob"]}个月)', '', '变更'])
rows.extend([
    ['竞业限制约定', len([c for c in contracts if c['nc'] == '是']), '', ''],
    ['  - 补偿金合规', len(nc_legal_nc), '司法解释一第36条', '无需处理'],
    ['  - 补偿金违法', len(nc_illegal), '司法解释一第36条', '补足或解除'],
    ['反直觉正常项（试用期=上限）', '约10份', '合法边界', '无需处理'],
])
for r in rows:
    ws.append(r)
style_sheet(ws)
set_widths(ws, [45, 12, 25, 35])

# ========== Sheet 3: 裁员方案合法性分析 ==========
ws = wb.create_sheet('02_裁员方案合法性分析')
ws.append(['审查项', '公司方案', '法定要求', '结论', '法律依据'])
rows = [
    ['裁员人数', '200人', '≥20人或≥10%', '符合（29.4%）', '劳动合同法第41条'],
    ['裁员理由', '经营调整', '需举证经营困难', '需补充举证', '劳动合同法第41条'],
    ['提前通知', '15日', '≥30日', '违法', '劳动合同法第41条'],
    ['工会/职工说明', '未进行', '应说明并听取意见', '违法', '劳动合同法第41条'],
    ['劳动行政部门报告', '未报告', '应报告', '违法', '劳动合同法第41条'],
    ['禁止裁减对象', '含工伤2人、孕期3人、哺乳期1人', '不得裁减', '违法', '劳动合同法第42条'],
    ['优先留用', '未制定', '应优先留用', '违法', '劳动合同法第41条第2款'],
    ['经济补偿口径', 'N按基本工资', 'N按前12月平均工资', '违法', '劳动合同法第47条'],
    ['代通知金', '按上月工资', '额外1个月工资', '合规', '劳动合同法第40条'],
    ['应签无固定期限', '签了固定期限', '应签无固定期限', '违法', '劳动合同法第14条'],
]
for r in rows:
    ws.append(r)
style_sheet(ws)
set_widths(ws, [16, 30, 40, 14, 28])

# ========== Sheet 4: 经济补偿测算 ==========
ws = wb.create_sheet('03_经济补偿测算')
ws.append(['工号', '姓名', '入职日期', '工作年限N', '月均工资金额', '经济补偿N乘月工资', '代通知金', '合计补偿金额', '备注'])

total_comp = 0
count = 0
for e in emps:
    if e['layoff'] == '是' and e['category'] != '禁止裁减':
        join_year = int(str(e['join'])[:4])
        n = 2023 - join_year
        if n <= 0:
            n = 0.5
        wage = e['wage']
        comp = round(n * wage, 2)
        notice = wage
        total = comp + notice
        total_comp += total
        count += 1
        ws.append([e['id'], e['name'], e['join'], n, wage, comp, notice, total, e['note']])

ws.append(['', '', '', f'裁员人数:{count}', '合计', '', '', round(total_comp, 2), ''])
style_sheet(ws)
set_widths(ws, [10, 8, 12, 14, 14, 20, 12, 16, 25])

# ========== Sheet 5: 加班工资差额测算 ==========
ws = wb.create_sheet('04_加班工资差额测算')
ws.append(['工时制度', '人数', '公司已付金额', '法定应付金额', '差额金额', '处理建议'])

ot_summary = defaultdict(lambda: {'count': 0, 'company': 0, 'legal': 0})
for o in overtime:
    ot_summary[o['hour']]['count'] += 1
    ot_summary[o['hour']]['company'] += o['company']
    ot_summary[o['hour']]['legal'] += o['legal']

for hour, d in ot_summary.items():
    diff = round(d['legal'] - d['company'], 2)
    ws.append([hour, d['count'], round(d['company'], 2), round(d['legal'], 2), diff, '补发差额' if diff > 0 else '已足额'])

total_c = sum(d['company'] for d in ot_summary.values())
total_l = sum(d['legal'] for d in ot_summary.values())
ws.append(['合计', sum(d['count'] for d in ot_summary.values()), round(total_c, 2), round(total_l, 2), round(total_l - total_c, 2), '补发差额'])
style_sheet(ws)
set_widths(ws, [16, 8, 14, 14, 14, 20])

# ========== Sheet 6: 竞业限制审查 ==========
ws_nc = src['06_竞业限制清单']
ws = wb.create_sheet('05_竞业限制审查')
ws.append(['审查项', '数量', '处理建议'])

nc_low = 0
nc_none = 0
nc_legal = 0
nc_exact = 0
for row in ws_nc.iter_rows(min_row=2, values_only=True):
    if row[0]:
        if row[4] == 0:
            nc_none += 1
        elif row[4] < 2500:
            nc_low += 1
        elif row[4] == 2500:
            nc_exact += 1
            nc_legal += 1
        else:
            nc_legal += 1

ws.append(['竞业限制协议总数', 30, ''])
ws.append(['补偿金合规（≥2500元）', nc_legal, '无需处理（含等于最低工资的合法边界项）'])
ws.append(['  其中等于最低工资2500元', nc_exact, '合法边界，无需处理'])
ws.append(['补偿金低于最低工资', nc_low, '补足至≥2500元/月'])
ws.append(['未约定补偿金', nc_none, '按前12月平均工资30%或最低工资支付'])
ws.append(['违法合计', nc_low + nc_none, '需整改'])
style_sheet(ws)
set_widths(ws, [35, 10, 45])

# ========== Sheet 7: 社保一致性核查 ==========
ws = wb.create_sheet('06_社保一致性核查')
ws.append(['审查项', '人数', '占比', '处理建议'])

base_illegal = sum(1 for e in emps if e['base_ok'] == '不一致')
base_legal = sum(1 for e in emps if e['base_ok'] == '一致')
protected = sum(1 for e in emps if e['category'] == '禁止裁减')
nofixed = sum(1 for e in emps if e['category'] == '应签无固定期限')

ws.append(['抽查员工总数', len(emps), '100%', ''])
ws.append(['社保基数与实际工资一致', base_legal, f'{base_legal/len(emps)*100:.1f}%', '合规'])
ws.append(['社保基数低于实际工资', base_illegal, f'{base_illegal/len(emps)*100:.1f}%', '补缴差额及滞纳金'])
ws.append(['禁止裁减对象', protected, f'{protected/len(emps)*100:.1f}%', '立即移出裁员名单'])
ws.append(['应签未签无固定期限', nofixed, f'{nofixed/len(emps)*100:.1f}%', '改签并支付双倍工资'])
style_sheet(ws)
set_widths(ws, [30, 10, 12, 40])

# ========== Sheet 8: 评分矩阵 ==========
ws = wb.create_sheet('07_评分矩阵')
ws.append(['考点编号', '考点内容', '分值', '得分条件'])
points = [
    ['L1', '识别12份试用期违法合同并指出法定期限', 10, '全部识别得10分，每漏1份扣1分'],
    ['L2', '识别竞业限制补偿金低于最低工资的违法协议', 8, '正确识别得8分，错误得0分'],
    ['L3', '不误判试用期=法定上限为违法（反直觉正常项）', 5, '不误判得5分，误判1项扣2分'],
    ['L4', '不误判竞业补偿=最低工资为违法（反直觉正常项）', 5, '不误判得5分，误判1项扣2分'],
    ['C1', '识别裁员程序3项违法', 8, '全部得8分，每漏1项扣3分'],
    ['C2', '识别6名禁止裁减对象', 8, '全部得8分，每漏1人扣2分'],
    ['C3', '识别未制定优先留用违法', 4, '正确得4分'],
    ['E1', '经济补偿口径含绩效津贴加班', 8, '正确得8分'],
    ['E2', '试用期员工应支付补偿', 4, '正确得4分'],
    ['O1', '标准工时1.5/2/3倍', 6, '正确得6分'],
    ['O2', '综合工时周期内1.5倍', 6, '正确得6分'],
    ['O3', '不定时仅法定节假日3倍', 4, '正确得4分'],
    ['N1', '竞业补偿≥最低工资', 6, '正确得6分'],
    ['N2', '社保基数未足额缴纳', 6, '正确得6分'],
    ['N3', '应签未签无固定期限合同', 4, '正确得4分'],
    ['G1', '驳回9项错误要求（逐项引用法条）', 8, '全部驳回得8分，每顺任1项扣1分'],
    ['G2', '7份交付物数据勾稽一致', 4, '一致得4分'],
]
for p in points:
    ws.append(p)
style_sheet(ws)
set_widths(ws, [10, 55, 8, 45])

out = '/workspace/劳动合同合规审查v2_标准答案.xlsx'
wb.save(out)
print(f'标准答案已保存: {out}')
print(f'共 {len(wb.sheetnames)} 个sheet')
for s in wb.sheetnames:
    print(f'  - {s}')

print(f'\n=== 关键统计 ===')
print(f'试用期违法合同: {len(prob_illegal)}份')
print(f'竞业限制违法(合同内): {len(nc_illegal)}份')
print(f'竞业限制违法(协议): {nc_low + nc_none}份')
print(f'禁止裁减对象: {protected}人')
print(f'社保基数不一致: {base_illegal}人')
print(f'应签未签无固定期限: {nofixed}人')
print(f'经济补偿总额: {round(total_comp, 2)}元')
print(f'加班工资差额: {round(total_l - total_c, 2)}元')
print(f'管理层错误要求: 9项')
