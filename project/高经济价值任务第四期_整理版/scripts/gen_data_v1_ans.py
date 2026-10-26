"""
场景六：互联网平台数据合规与PIA - 标准答案
"""
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from collections import Counter

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

src = openpyxl.load_workbook('/workspace/数据合规与PIA_题目附件.xlsx', data_only=True)

# 读取数据
ws_da = src['02_数据处理活动清单']
activities = []
for row in ws_da.iter_rows(min_row=2, values_only=True):
    if row[0]:
        activities.append({
            'id': row[0], 'name': row[1], 'cat': row[2], 'field_count': row[3],
            'sensitive': row[4], 'basis': row[5], 'purpose': row[6],
            'source': row[7], 'storage': row[8], 'cross_border': row[9],
            'notified': row[10], 'separate_consent': row[11], 'retention': row[12], 'status': row[13]
        })

ws_sdk = src['03_SDK清单']
sdks = []
for row in ws_sdk.iter_rows(min_row=2, values_only=True):
    if row[0]:
        sdks.append({
            'id': row[0], 'name': row[1], 'provider': row[2], 'type': row[3],
            'permissions': row[4], 'has_pi': row[5], 'disclosed': row[6],
            'consented': row[7], 'has_report': row[8], 'direction': row[9],
            'status': row[10], 'note': row[11]
        })

ws_field = src['04_个人信息字段清单']
fields = []
for row in ws_field.iter_rows(min_row=2, values_only=True):
    if row[0]:
        fields.append({
            'id': row[0], 'name': row[1], 'cat': row[2], 'sensitive': row[3],
            'scene': row[4], 'basis': row[5], 'necessary': row[6],
            'consent': row[7], 'separate': row[8], 'retention': row[9],
            'status': row[10], 'note': row[11]
        })

ws_share = src['06_数据共享协议清单']
sharing = []
for row in ws_share.iter_rows(min_row=2, values_only=True):
    if row[0]:
        sharing.append({
            'id': row[0], 'receiver': row[1], 'data_type': row[2], 'purpose': row[3],
            'consent': row[4], 'agreement': row[5], 'export_assess': row[6],
            'status': row[7], 'note': row[8]
        })

# ========== Sheet 1: 总体结论与风险清单 ==========
ws = wb.create_sheet('00_总体结论与风险清单')
ws.append(['场景六：互联网平台数据合规与PIA - 标准答案'])
ws['A1'].font = title_font
ws['A1'].fill = title_fill
ws.merge_cells('A1:F1')
ws['A1'].alignment = Alignment(horizontal='center', vertical='center')

ws.append([''])
ws.append(['一、总体结论'])
ws.append(['华数科技数据合规存在10项重大风险：数据处理活动合法性基础错误约30项、SDK未披露及出境未评估12个、敏感个人信息未获单独同意约25项、隐私政策与实际采集行为不一致12处、数据出境未评估15项、数据共享协议缺失4份、个人信息保留期限超期18项、非必要信息错误标记为必要信息若干、关联公司间数据共享未做出境评估、管理层9项错误要求。若不整改将面临监管处罚、APP下架及融资失败。'])
ws.append([''])
ws.append(['二、风险清单（15项）'])

da_illegal = sum(1 for a in activities if a['status'] == '违法')
sdk_illegal = sum(1 for s in sdks if s['status'] == '违法')
sdk_pending = sum(1 for s in sdks if s['status'] == '待整改')
field_illegal = sum(1 for f in fields if f['status'] == '违法')
share_illegal = sum(1 for s in sharing if s['status'] == '违法')

risks = [
    ['序号', '风险事项', '风险等级', '法律依据', '整改建议'],
    [1, f'数据处理活动合法性基础错误约{da_illegal}项', '高', '个人信息保护法第13条', '按场景正确设定合法性基础'],
    [2, f'{sdk_illegal}个SDK未在隐私政策中披露', '高', '个人信息保护法第17条', '补充披露并获用户同意'],
    [3, f'{sdk_illegal}个境外SDK数据出境未评估', '高', '个人信息保护法第38条', '做数据出境安全评估或停止使用'],
    [4, f'敏感个人信息未获单独同意约25项', '高', '个人信息保护法第29条', '拆分同意流程获单独同意'],
    [5, '隐私政策与APP实际采集行为不一致12处', '高', '个人信息保护法第17条', '修改隐私政策或修改APP行为'],
    [6, f'数据出境未评估约15项', '高', '个人信息保护法第38条', '做数据出境安全评估'],
    [7, f'数据共享协议缺失{share_illegal}份', '高', '个人信息保护法第23条', '签订数据共享协议并获单独同意'],
    [8, f'个人信息保留期限超期约18项', '中', '个人信息保护法第19条', '设定差异化最短保留期限'],
    [9, f'非必要信息错误标记为必要信息若干项', '高', 'App违法违规收集认定方法第4条', '区分必要与非必要信息'],
    [10, '关联公司间数据共享未做出境评估', '高', '个人信息保护法第38条', '关联公司间出境同样适用'],
    [11, f'{sdk_pending}个SDK缺独立合规报告', '中', 'App违法违规收集认定方法', '要求SDK提供方提供合规报告'],
    [12, '用户权利行使入口隐藏', '中', '个人信息保护法第45条', '提供便捷的删除权行使入口'],
    [13, '隐私政策版本未及时更新', '中', '个人信息保护法第17条', '随处理活动变化及时更新'],
    [14, '管理层9项错误要求', '高', '多项法律', '逐项驳回并依法执行'],
    [15, '监管处罚与APP下架风险', '极高', '个人信息保护法第64条', '立即整改并主动报告'],
]
for r in risks:
    ws.append(r)
style_sheet(ws, header_row=5)
set_widths(ws, [6, 45, 10, 28, 40])

# ========== Sheet 2: 数据处理活动审查 ==========
ws = wb.create_sheet('01_数据处理活动审查')
ws.append(['审查项', '数量', '依据', '处理建议'])

legal_basis_errors = sum(1 for a in activities if '合法性基础' in str(a.get('note', '')) or '非必要信息' in str(a.get('note', '')))
unnotified = sum(1 for a in activities if a['notified'] == '否')
no_consent = sum(1 for a in activities if a['separate_consent'] == '否' and a['sensitive'] == '是')
retention_issues = sum(1 for a in activities if '保留期限' in str(a.get('note', '')))
cross_border_issues = sum(1 for a in activities if '出境' in str(a.get('note', '')))
normal_traps = sum(1 for a in activities if '合法边界' in str(a.get('note', '')) or '合法' in str(a.get('note', '')))

rows = [
    ['处理活动总数', 280, '', ''],
    ['合法性基础错误', legal_basis_errors, '个人信息保护法第13条', '按场景正确设定'],
    ['未告知用户', unnotified, '个人信息保护法第17条', '补充告知'],
    ['敏感信息未获单独同意', no_consent, '个人信息保护法第29条', '拆分同意流程'],
    ['保留期限超期', retention_issues, '个人信息保护法第19条', '设定最短保留期限'],
    ['数据出境未评估', cross_border_issues, '个人信息保护法第38条', '做出境安全评估'],
    ['反直觉正常项（合法边界）', normal_traps, '合法', '无需处理'],
    ['违法合计', da_illegal, '', '需整改'],
]
for r in rows:
    ws.append(r)
style_sheet(ws)
set_widths(ws, [35, 12, 25, 35])

# ========== Sheet 3: SDK审查 ==========
ws = wb.create_sheet('02_SDK审查')
ws.append(['审查项', '数量', '依据', '处理建议'])

sdk_undisclosed = sum(1 for s in sdks if s['disclosed'] == '否')
sdk_overseas = sum(1 for s in sdks if s['direction'] and '境外' in str(s['direction']))
sdk_no_report = sum(1 for s in sdks if s['has_report'] == '否' and s['has_pi'] == '是')
sdk_legal = sum(1 for s in sdks if s['status'] == '合规')

rows = [
    ['SDK总数', 48, '', ''],
    ['已披露合规', sdk_legal, '个人信息保护法第17条', '无需处理'],
    ['未披露', sdk_undisclosed, '个人信息保护法第17条', '补充披露'],
    ['境外SDK出境未评估', sdk_undisclosed, '个人信息保护法第38条', '做出境评估或停止使用'],
    ['缺独立合规报告', sdk_no_report, 'App违法违规收集认定方法', '要求提供方提供合规报告'],
    ['违法合计', sdk_illegal, '', '需整改'],
    ['待整改合计', sdk_pending, '', '需补充合规报告'],
]
for r in rows:
    ws.append(r)
style_sheet(ws)
set_widths(ws, [35, 12, 25, 35])

# ========== Sheet 4: 个人信息字段审查 ==========
ws = wb.create_sheet('03_个人信息字段审查')
ws.append(['审查项', '数量', '依据', '处理建议'])

field_basis_error = sum(1 for f in fields if '合法性基础' in str(f.get('note', '')))
field_no_consent = sum(1 for f in fields if f['separate'] == '否' and f['sensitive'] == '是')
field_retention = sum(1 for f in fields if '保留期限' in str(f.get('note', '')))
field_necessary = sum(1 for f in fields if '非必要' in str(f.get('note', '')))
field_normal = sum(1 for f in fields if '合法边界' in str(f.get('note', '')))

rows = [
    ['字段总数', 150, '', ''],
    ['合法性基础错误', field_basis_error, '个人信息保护法第13条', '按场景设定'],
    ['敏感信息未获单独同意', field_no_consent, '个人信息保护法第29条', '获单独同意'],
    ['保留期限超期', field_retention, '个人信息保护法第19条', '设定最短期限'],
    ['非必要信息标记为必要', field_necessary, 'App认定方法第4条', '区分必要与非必要'],
    ['反直觉正常项（合法边界）', field_normal, '合法', '无需处理'],
    ['违法合计', field_illegal, '', '需整改'],
]
for r in rows:
    ws.append(r)
style_sheet(ws)
set_widths(ws, [35, 12, 25, 35])

# ========== Sheet 5: 隐私政策比对结果 ==========
ws = wb.create_sheet('04_隐私政策比对结果')
ws.append(['比对项', '不一致原因', '合规状态', '整改建议'])

ws_cmp = src['05_隐私政策比对']
for row in ws_cmp.iter_rows(min_row=2, values_only=True):
    if row[0]:
        ws.append([row[0], row[3], row[4], row[5]])
style_sheet(ws)
set_widths(ws, [25, 20, 12, 40])

# ========== Sheet 6: 数据共享审查 ==========
ws = wb.create_sheet('05_数据共享审查')
ws.append(['审查项', '数量', '依据', '处理建议'])

share_legal = sum(1 for s in sharing if s['status'] == '合规')
share_illegal_count = sum(1 for s in sharing if s['status'] == '违法')

rows = [
    ['共享协议总数', 12, '', ''],
    ['合规共享', share_legal, '', '无需处理'],
    ['违法共享', share_illegal_count, '个人信息保护法第23、38条', '签协议+做出境评估+获同意'],
    ['未签协议', sum(1 for s in sharing if s['agreement'] == '否'), '个人信息保护法第23条', '签订数据共享协议'],
    ['未获同意', sum(1 for s in sharing if s['consent'] == '否'), '个人信息保护法第23条', '获用户单独同意'],
    ['未做出境评估', sum(1 for s in sharing if s['export_assess'] == '否'), '个人信息保护法第38条', '做出境安全评估'],
]
for r in rows:
    ws.append(r)
style_sheet(ws)
set_widths(ws, [35, 12, 25, 35])

# ========== Sheet 7: PIA评估结果 ==========
ws = wb.create_sheet('06_PIA评估结果')
ws.append(['评估项', '风险等级', '风险描述', '缓解措施'])

pia_items = [
    ['敏感个人信息处理', '高', f'约{no_consent}项敏感信息未获单独同意，位置信息、人脸特征值、通讯录等敏感信息处理活动未做PIA', '拆分同意流程，逐项开展PIA'],
    ['数据出境', '高', f'约{cross_border_issues}项数据出境未做安全评估，含SDK出境和关联公司间共享', '做数据出境安全评估或停止出境'],
    ['SDK合规', '高', f'{sdk_illegal}个SDK未披露，{sdk_undisclosed}个境外SDK未评估', '补充披露、做出境评估或停止使用'],
    ['隐私政策一致性', '高', '隐私政策与APP实际采集行为存在12处不一致', '修改隐私政策或修改APP行为'],
    ['保留期限', '中', f'约{retention_issues}项保留期限超期或永久保留', '设定差异化最短保留期限'],
    ['合法性基础', '高', f'约{legal_basis_errors}项合法性基础设定错误', '按场景正确设定合法性基础'],
    ['数据共享', '高', f'{share_illegal_count}份数据共享协议缺失或未评估', '签订协议、获同意、做出境评估'],
    ['用户权利保障', '中', '删除权入口隐藏在5级菜单，难以行使', '提供便捷的权利行使入口'],
    ['自动化决策', '中', '营销推送使用用户画像进行自动化决策，未提供退出机制', '提供便捷的退出自动化决策机制'],
    ['数据安全', '中', '数据安全等级保护三级，但部分敏感信息未加密存储', '对敏感信息实施加密存储'],
]
for p in pia_items:
    ws.append(p)
style_sheet(ws)
set_widths(ws, [25, 10, 50, 40])

# ========== Sheet 8: 评分矩阵 ==========
ws = wb.create_sheet('07_评分矩阵')
ws.append(['考点编号', '考点内容', '分值', '得分条件'])
points = [
    ['D1', '识别合法性基础错误（非必要信息用履行合同所必需）', 8, '正确识别得8分'],
    ['D2', '识别SDK未披露及出境未评估', 8, '正确识别得8分'],
    ['D3', '识别敏感个人信息未获单独同意', 8, '正确识别得8分'],
    ['D4', '识别隐私政策与实际采集行为不一致', 8, '正确识别得8分'],
    ['D5', '识别数据出境未评估', 8, '正确识别得8分'],
    ['D6', '识别数据共享协议缺失', 6, '正确识别得6分'],
    ['D7', '识别保留期限超期', 6, '正确识别得6分'],
    ['D8', '识别非必要信息标记为必要', 6, '正确识别得6分'],
    ['D9', '不误判履行合同所必需的敏感信息不需单独同意（反直觉正常项）', 5, '不误判得5分'],
    ['D10', '不误判员工个人数据出境不适用安全评估（反直觉正常项）', 5, '不误判得5分'],
    ['D11', '不误判法定义务共享不需同意（反直觉正常项）', 5, '不误判得5分'],
    ['D12', '正确识别关联公司间出境同样适用安全评估', 5, '正确得5分'],
    ['D13', '驳回管理层9项错误要求（逐项引用法条）', 8, '全部驳回得8分，每顺任1项扣1分'],
    ['D14', '7份交付物数据勾稽一致', 5, '一致得5分'],
    ['D15', 'PIA报告覆盖全部280项处理活动', 5, '全覆盖得5分'],
]
for p in points:
    ws.append(p)
style_sheet(ws)
set_widths(ws, [10, 55, 8, 45])

out = '/workspace/数据合规与PIA_标准答案.xlsx'
wb.save(out)
print(f'标准答案已保存: {out}')
print(f'共 {len(wb.sheetnames)} 个sheet')
for s in wb.sheetnames:
    print(f'  - {s}')

print(f'\n=== 关键统计 ===')
print(f'处理活动违法: {da_illegal}项')
print(f'SDK违法: {sdk_illegal}个')
print(f'SDK待整改: {sdk_pending}个')
print(f'字段违法: {field_illegal}项')
print(f'共享违法: {share_illegal_count}份')
print(f'反直觉正常项(处理活动): {normal_traps}项')
print(f'反直觉正常项(字段): {field_normal}项')
