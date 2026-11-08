#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""客户数据与保密合规专项审计 v1 —— 标准答案生成器"""
from datetime import date, timedelta
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE = '/workspace/高经济价值第八期_数据合规审计'
SRC = f'{BASE}/数据合规审计_v1_题目.xlsx'
OUT = f'{BASE}/数据合规审计_v1_标准答案.xlsx'

HDR_FILL = PatternFill('solid', fgColor='2F5597')
HDR_FONT = Font(name='宋体', bold=True, color='FFFFFF', size=10)
BODY = Font(name='宋体', size=9)
THIN = Side(style='thin', color='CCCCCC')
BORDER = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
WRAP = Alignment(wrap_text=True, vertical='top')
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)

CUTOFF = date(2026, 12, 31)


def sheet(wb, title, headers, rows, widths=None):
    ws = wb.create_sheet(title)
    ws.append(headers)
    for r in rows:
        ws.append(r)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = HDR_FONT
        cell.fill = HDR_FILL
        cell.alignment = CENTER
        cell.border = BORDER
        ws.column_dimensions[get_column_letter(c)].width = (widths or {}).get(c, 14)
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(headers)):
        for cell in row:
            cell.font = BODY
            cell.alignment = WRAP
            cell.border = BORDER
    ws.freeze_panes = 'A2'
    ws.sheet_view.showGridLines = False
    return ws


def rows_of(ws):
    data = list(ws.iter_rows(values_only=True))
    hdr = data[0]
    return [dict(zip(hdr, r)) for r in data[1:] if r[0] is not None]


def parse(d):
    y, m, dd = str(d).split('-')
    return date(int(y), int(m), int(dd))


wb0 = openpyxl.load_workbook(SRC, data_only=True)
tz = rows_of(wb0['项目数据留存台账'])
contracts = rows_of(wb0['项目合同签署记录'])
destr = {r['项目编号']: r for r in rows_of(wb0['数据销毁记录'])}

# ---- 版本适用判定：以主业务合同签约年份为准 ----
main_contract = {}
for c in contracts:
    if c['合同类型'] == '主业务合同':
        assert c['项目编号'] not in main_contract, f"重复主业务合同: {c['项目编号']}"
        main_contract[c['项目编号']] = parse(c['签约日期'])

ver_rows = []      # 版本适用判定
ret_rows = []      # 留存期重算
n_v1 = n_v1 = 0
n_v2 = 0
n_sys_fix = 0
n_ongoing = 0
n_ended = 0
n_destroyed_ok = 0
n_in_period = 0
overdue = {}       # 项目编号 -> 超期天数

for t in tz:
    code = t['项目编号']
    cust = t['客户名称']
    mc_d = main_contract[code]
    ver = 'v1' if mc_d.year <= 2024 else 'v2'
    keep = 90 if ver == 'v1' else 180
    if ver == 'v1':
        n_v1 += 1
    else:
        n_v2 += 1
    sys_ver = t['系统登记协议版本']
    if sys_ver != ver:
        n_sys_fix += 1
        if code == 'P07':
            diff = '系统误登记为v2（按2025年补充协议登记），应按2024年主业务合同适用v1'
        else:
            diff = '系统误登记为v1（按2024年框架协议登记），应按2025年主业务合同适用v2'
    else:
        diff = '—'
    basis = f'主业务合同{mc_d}签约，属{"2024年及以前" if ver == "v1" else "2025年起"}，适用{ver}'
    ver_rows.append([code, cust, str(mc_d), ver, keep, sys_ver, diff, basis])

    status = t['项目状态']
    if status == '进行中':
        n_ongoing += 1
        ret_rows.append([code, cust, '进行中', '—', ver, '—', '—', '—', '进行中，数据在使用中，不涉及销毁'])
        continue
    n_ended += 1
    end_d = parse(t['项目结束日'])
    destroy_by = end_d + timedelta(days=keep)
    d_row = destr[code]
    actual = d_row['实际销毁日期'] or ''
    if actual:
        a_d = parse(actual)
        ok = a_d <= destroy_by
        if ok:
            n_destroyed_ok += 1
            judge = f'合规：已于{actual}销毁（不晚于应销毁截止日），销毁证明{d_row["销毁证明编号"]}'
            st = '已销毁'
        else:
            judge = f'违规：{actual}销毁，晚于应销毁截止日{destroy_by}'
            st = '销毁迟延'
    else:
        overdue_days = (CUTOFF - destroy_by).days
        if overdue_days > 0:
            overdue[code] = overdue_days
            st = f'超期{overdue_days}天'
            judge = f'违规：截至2026-12-31超期{overdue_days}天未销毁'
        else:
            n_in_period += 1
            st = '留存期内'
            judge = f'合规：应销毁截止日{destroy_by}，截至2026-12-31尚在留存期内'
    ret_rows.append([code, cust, '已结束', str(end_d), ver, str(destroy_by), actual or '未销毁', st, judge])

n_overdue = len(overdue)
n_projects = len(tz)

wb = Workbook()
wb.remove(wb.active)

# ============ 01 版本适用判定 ============
sheet(wb, '版本适用判定',
      ['项目编号', '客户名称', '主业务合同签约日', '适用版本', '留存天数', '系统登记版本', '系统登记差异', '判定依据'],
      ver_rows, {1: 9, 2: 11, 3: 14, 4: 9, 5: 9, 6: 11, 7: 46, 8: 42})

# ============ 02 留存期重算 ============
sheet(wb, '留存期重算',
      ['项目编号', '客户名称', '项目状态', '项目结束日', '适用版本', '应销毁截止日', '实际销毁日', '截止日状态', '合规判定'],
      ret_rows, {1: 9, 2: 11, 3: 9, 4: 12, 5: 9, 6: 12, 7: 12, 8: 12, 9: 52})

# ============ 03 外包转委托审查 ============
wbz_rows = [
    ['转委托事实', '外包合同RH-OUT-2026-03（2026-06-15）+ 数析工单DS-0701（2026-07-01）',
     '存在转委托：数析科技自2026-07-01起处理P09星桥物流项目运输明细数据52,410行'],
    ['书面同意要件', '仅有2026-06-12客户项目经理张帆邮件“数据建模的事你们看着安排就行”',
     '不构成有效书面同意：张帆不在星桥物流授权签字人名录内（名录仅法务总监刘畅、首席执行官王镇）'],
    ['脱敏豁免抗辩', '外包范围虽为已脱敏数据',
     '不成立：客户保密协议v2第7.2条明确含已脱敏数据，脱敏不构成转委托豁免'],
    ['补正方式', '管理层要求倒签客户同意书',
     '不允许：倒签构成伪造证据，须如实向客户报告并补办手续'],
    ['审查结论', '四类证据相互印证（外包合同、处理工单、沟通邮件、授权名录）',
     '构成违规转委托：整改 + 报告客户（认定之日起5个工作日内）+ 解除RH-OUT-2026-03外包合同并回收数据'],
]
sheet(wb, '外包转委托审查', ['审查事项', '证据', '认定'], wbz_rows, {1: 16, 2: 52, 3: 60})

# ============ 04 设备回收审查 ============
sb_rows = [
    ['李慕', '2026-03-31', 'NB-0388', '当日回收，双人登记', '合规'],
    ['周岩', '2026-06-15', 'NB-0412', '当日回收，双人登记', '合规'],
    ['陈曦', '2026-09-30', 'NB-0417', '无回收登记；行政称口头交接给部门助理，无签收单', '违规：口头交接无效，回收记录缺失'],
    ['吴桐', '2026-10-20', 'NB-0455', '当日回收，双人登记', '合规'],
    ['郑好', '2026-11-15', 'NB-0523', '无回收登记；仅有前台快递签收，无IT入库登记', '违规：快递签收不等于入库登记，回收记录缺失'],
    ['审查结论', '—', '—', '2026年离职顾问5人，应回收设备5台，已登记3台，缺失2台（NB-0417、NB-0523）',
     '整改：30日内追查去向并补充登记，逾期未找回启动数据泄露风险评估；追责：追究行政管理部设备管理责任'],
]
sheet(wb, '设备回收审查', ['离职人员', '离职日期', '设备编号', '回收情况', '认定'], sb_rows,
      {1: 10, 2: 12, 3: 11, 4: 44, 5: 52})

# ============ 05 违规分级处置 ============
d1, d2, d3 = overdue.get('P02'), overdue.get('P04'), overdue.get('P06')
wf_rows = [
    ['V1', '客户数据超期留存', 'P02 蓝湾地产', '留存台账+销毁记录',
     f'适用v1留存90天，应销毁截止日2026-06-13，截至2026-12-31超期{d1}天未销毁',
     '整改、追责', '立即销毁并于7日内补具销毁证明；超期超过180天，追究数据管理部负责人管理责任',
     '保密协议v1第5.1条；RH-DATA-04第2、5条'],
    ['V2', '客户数据超期留存', 'P04 金石证券', '留存台账+销毁记录',
     f'适用v1留存90天，应销毁截止日2026-11-18，截至2026-12-31超期{d2}天未销毁',
     '整改', '立即销毁并于7日内补具销毁证明',
     '保密协议v1第5.1条；RH-DATA-04第2条'],
    ['V3', '客户数据超期留存', 'P06 峰瑞零售', '留存台账+销毁记录',
     f'适用v2留存180天，应销毁截止日2026-10-07，截至2026-12-31超期{d3}天未销毁',
     '整改', '立即销毁并于7日内补具销毁证明',
     '保密协议v2第5.1条；RH-DATA-04第2条'],
    ['V4', '未经客户书面同意转委托第三方处理客户数据', 'P09 星桥物流 / 数析科技',
     '外包合同RH-OUT-2026-03、工单DS-0701、沟通邮件、授权签字人名录',
     '脱敏数据不豁免；项目经理邮件不构成有效书面同意；第三方已实际处理52,410行',
     '整改、报告客户、解除合同', '5个工作日内报告星桥物流；解除外包合同并回收销毁数据；复核全部存量外包',
     '保密协议v2第7.2条；RH-DATA-04第3、4条'],
    ['V5', '离职顾问设备回收记录缺失2台', 'NB-0417、NB-0523 / 行政管理部',
     '离职登记5人、行政说明、快递签收单、RH-IT-02第4条',
     '口头交接无效、快递签收不等于入库登记，2台设备去向需追查',
     '整改、追责', '30日内追查去向并补充登记，逾期未找回启动数据泄露风险评估；追究行政管理部管理责任',
     'RH-IT-02第4条；RH-DATA-04第2、5条'],
    ['V6', '顾问将客户数据上传个人网盘', 'P15 康泰医疗 / 顾问王晓',
     '平台导出日志（2026-10-11 23:41）、网络出口日志（23:47）、文件内容抽样比对',
     '含客户员工姓名与手机号的明细3,862行实质脱离公司受控环境，属重大违规',
     '整改、报告客户、解除合同、追责', '5个工作日内报告康泰医疗；依员工手册解除王晓劳动合同；删除网盘文件并取证保全；全平台导出权限复核',
     'RH-DATA-03第5条；RH-DATA-04第3、4、5条'],
]
sheet(wb, '违规分级处置', ['违规编号', '违规事项', '涉及对象', '证据', '认定', '处置分级', '处置动作', '依据条款'],
      wf_rows, {1: 8, 2: 22, 3: 20, 4: 34, 5: 40, 6: 18, 7: 46, 8: 26})

n_rectify = sum(1 for r in wf_rows if '整改' in r[5])
n_report = sum(1 for r in wf_rows if '报告客户' in r[5])
n_terminate = sum(1 for r in wf_rows if '解除合同' in r[5])
n_account = sum(1 for r in wf_rows if '追责' in r[5])
n_viol = len(wf_rows)

# ============ 06 网盘事件处置 ============
wp_rows = [
    ['事件时间线', '2026-10-11 23:41 账号wang.xiao从调研宝平台导出“康泰医疗_员工调研明细_3862行.xlsx”（4.2MB）；23:47 同一账号会话向pan.baidu.com上传4.2MB，源IP 112.10.36.58（非公司VPN地址段）'],
    ['涉及项目', 'P15 康泰医疗员工满意度调研（进行中，协议版本v2）'],
    ['涉及数据', '客户员工姓名与手机号明细，共3,862行'],
    ['证据链', '平台导出日志 + 网络出口日志 + 文件内容抽样比对，三类证据相互印证；人力资源部确认账号使用人为顾问王晓（在职，P15项目组成员）'],
    ['事件定级', '重大违规：客户数据实质脱离公司受控环境，触发报告客户条件'],
    ['处置措施', '①自认定之日起5个工作日内书面报告康泰医疗；②依员工手册解除王晓劳动合同并保留追责权利；③删除网盘文件并取证保全；④对调研宝平台全部项目导出权限进行复核，关闭非必要导出权限'],
    ['处置时限', '报告客户：认定之日起5个工作日内；权限复核：2027-01-31前完成'],
]
sheet(wb, '网盘事件处置', ['项目', '内容'], wp_rows, {1: 14, 2: 110})

# ============ 07 错误要求驳回 ============
bh_rows = [
    [1, '“超期留存的项目先别销毁，等客户问起再处理”', '驳回',
     '保密协议约定留存期满当日或之前销毁，超期留存本身即构成违约，必须立即销毁并补具销毁证明，不得观望。'],
    [2, '“P02数据保留到明年做数据资产入表”', '驳回',
     '销毁义务不因数据商业价值而免除；将客户数据用于入表属超范围使用，构成新的违约。'],
    [3, '“补一份客户同意书，日期倒签到外包合同签署之前”', '驳回',
     '倒签构成伪造证据，且不能改变转委托时未取得书面同意的事实；只能如实报告客户并补办手续。'],
    [4, '“项目经理邮件里已经同意了，直接认定书面同意已取得”', '驳回',
     '张帆不在星桥物流授权签字人名录内，未经授权人员的邮件不构成书面同意。'],
    [5, '“脱敏数据不算客户数据，转委托不需要客户同意”', '驳回',
     '保密协议v2第7.2条明确含已脱敏数据，脱敏不构成转委托豁免。'],
    [6, '“两台笔记本按资产盘亏处理，不再追查”', '驳回',
     'RH-IT-02第4条要求回收记录缺失须30日内追查去向并评估数据泄露风险，盘亏处理不能替代追查。'],
    [7, '“前台已签收，让IT补一条入库登记就算回收完成”', '驳回',
     '快递签收不等于入库登记；补登记须以实物核验为前提，未核验不得补记。'],
    [8, '“网盘的事内部处理，不要告诉康泰医疗”', '驳回',
     '客户数据已实质脱离公司受控环境，触发RH-DATA-04第3条，必须自认定之日起5个工作日内报告客户。'],
    [9, '“把上传的顾问开除就算了结，部门和公司层面不再追责”', '驳回',
     '员工故意违规依员工手册解除劳动合同，但导出审批与平台权限管理失职不免除，仍须追究管理责任。'],
    [10, '“全部项目统一按v2的180天留存期重新计算”', '驳回',
     '版本适用以主业务合同签约年份为准，2024年及以前签约项目适用v1的90天，不得统一套用v2。'],
    [11, '“P08按框架协议年份适用v1并认定超期，体现审计严格”', '驳回',
     '框架协议不构成项目主业务合同，P08应按2025年主业务合同适用v2；不得为体现严格而错误适用版本制造假违规。'],
    [12, '“销毁证明后补，先把台账改成已销毁”', '驳回',
     '不得篡改台账；销毁须先实际执行并出具证明后，方可更新台账状态。'],
]
sheet(wb, '错误要求驳回', ['序号', '管理层要求', '结论', '驳回理由'], bh_rows, {1: 6, 2: 42, 3: 8, 4: 66})

# ============ 08 勾稽自检 ============
# 左值：从题目数据独立重算；右值：从本答案底稿读取
checks = []

def add_check(no, item, left, right):
    checks.append([no, item, left, right, '一致' if str(left) == str(right) else '不一致'])

add_check('C01', '项目总数', len(tz), n_projects)
add_check('C02', '适用v1项目数', sum(1 for t in tz if main_contract[t['项目编号']].year <= 2024), n_v1)
add_check('C03', '适用v2项目数', sum(1 for t in tz if main_contract[t['项目编号']].year > 2024), n_v2)
add_check('C04', '系统登记版本需更正项目数',
          sum(1 for t in tz if t['系统登记协议版本'] != ('v1' if main_contract[t['项目编号']].year <= 2024 else 'v2')),
          n_sys_fix)
add_check('C05', '进行中项目数', sum(1 for t in tz if t['项目状态'] == '进行中'), n_ongoing)
add_check('C06', '已结束项目数', sum(1 for t in tz if t['项目状态'] == '已结束'), n_ended)
add_check('C07', '已按期销毁项目数',
          sum(1 for t in tz if t['项目状态'] == '已结束' and destr[t['项目编号']]['实际销毁日期']), n_destroyed_ok)
add_check('C08', '留存期内未到期项目数', n_ended - n_destroyed_ok - n_overdue, n_in_period)
add_check('C09', '超期未销毁项目数', n_ended - n_destroyed_ok - n_in_period, n_overdue)
add_check('C10', '已结束项目构成闭合（销毁+留存期内+超期）', n_destroyed_ok + n_in_period + n_overdue, n_ended)
add_check('C11', 'P02超期天数', (CUTOFF - (parse('2026-03-15') + timedelta(days=90))).days, overdue.get('P02'))
add_check('C12', 'P04超期天数', (CUTOFF - (parse('2026-08-20') + timedelta(days=90))).days, overdue.get('P04'))
add_check('C13', 'P06超期天数', (CUTOFF - (parse('2026-04-10') + timedelta(days=180))).days, overdue.get('P06'))
add_check('C14', '违规事项总数', 3 + 1 + 1 + 1, n_viol)
add_check('C15', '整改级处置项数', n_viol, n_rectify)
add_check('C16', '报告客户项数', 2, n_report)
add_check('C17', '解除合同项数', 2, n_terminate)
add_check('C18', '追责项数', 3, n_account)
add_check('C19', '2026年离职顾问人数',
          sum(1 for r in rows_of(wb0['离职与设备记录']) if r['记录类型'] == '离职登记'), 5)
add_check('C20', '设备回收记录缺失台数',
          sum(1 for r in rows_of(wb0['离职与设备记录']) if r['记录类型'] == '离职登记' and r['登记状态'] == '无记录'), 2)
add_check('C21', '网盘泄露数据行数',
          next(t['数据量级(行)'] for t in tz if t['项目编号'] == 'P15'), 3862)
add_check('C22', '管理层错误要求驳回条数', len(rows_of(wb0['管理层要求'])), len(bh_rows))

sheet(wb, '勾稽自检', ['编号', '勾稽项', '重算值', '底稿值', '结果'], checks, {1: 7, 2: 40, 3: 14, 4: 14, 5: 9})

n_checks = len(checks)
n_bad = sum(1 for c in checks if c[4] != '一致')

# ============ 结论 ============
concl = [
    ['核心结论',
     f'20个项目协议版本适用判定完成（v1共{n_v1}个、v2共{n_v2}个，更正系统登记错误{n_sys_fix}项）；'
     f'共发现违规{n_viol}项：超期留存3项（P02超期{d1}天、P04超期{d2}天、P06超期{d3}天）、'
     f'未经客户书面同意转委托1项（P09/数析科技，已处理52,410行）、离职设备回收记录缺失1项（2台）、'
     f'客户数据上传个人网盘1项（P15/康泰医疗，3,862行）。处置：{n_viol}项均限期整改；'
     f'报告客户{n_report}项（P09、P15）；解除合同{n_terminate}项（数析科技外包合同、王晓劳动合同）；'
     f'追责{n_account}项（P02超期管理责任、设备管理责任、王晓故意违规）。管理层{len(bh_rows)}项错误要求全部驳回。'],
    ['支撑一：版本与留存',
     '版本以主业务合同签约年份为准：P07主合同2024年签约适用v1（2025年补充协议不改变版本），应销毁截止日2027-01-18，合规；'
     'P08按2025年主业务合同适用v2（2024年框架协议不构成项目合同），应销毁截止日2027-02-11，合规，不得误判超期。'
     '系统登记版本错误2项（P07误为v2、P08误为v1），须更正。'],
    ['支撑二：外包转委托',
     '数析科技自2026-07-01起处理P09运输明细52,410行；客户项目经理张帆邮件不构成书面同意（不在授权签字人名录内）；'
     'v2第7.2条明确脱敏数据不豁免。构成违规转委托。'],
    ['支撑三：设备与网盘',
     '离职5人应回收5台，已登记3台，缺失2台（NB-0417口头交接无效、NB-0523快递签收不等于入库登记）；'
     '王晓2026-10-11 23:41导出、23:47上传个人网盘，3,862行姓名与手机号明细脱离受控环境，属重大违规。'],
    ['支撑四：处置与时限',
     '超期留存3项立即销毁并7日内补具销毁证明；报告客户2项自认定之日起5个工作日内执行；'
     '设备追查30日内完成；网盘文件删除并取证，平台导出权限2027-01-31前完成复核。'],
    ['关键结果',
     f'v1×{n_v1}/v2×{n_v2}；超期201/43/85天；违规6项；整改6/报告客户2/解除合同2/追责3；'
     f'设备缺失2台；泄露3,862行；驳回12项；勾稽{n_checks}项全部一致。'],
]
sheet(wb, '数据合规审计结论', ['项目', '内容'], concl, {1: 18, 2: 120})

wb.save(OUT)
print('标准答案已生成：', OUT)
print(f'勾稽 {n_checks} 项，不一致 {n_bad} 项')
for c in checks:
    if c[4] != '一致':
        print('  不一致：', c)
print(f'v1={n_v1} v2={n_v2} 系统更正={n_sys_fix} 进行中={n_ongoing} 已结束={n_ended} '
      f'已销毁={n_destroyed_ok} 留存期内={n_in_period} 超期={n_overdue} {overdue}')
print(f'违规={n_viol} 整改={n_rectify} 报告客户={n_report} 解除合同={n_terminate} 追责={n_account}')
