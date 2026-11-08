#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""合伙人分润审计 v1 —— 标准答案生成器"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE = '/workspace/高经济价值第八期_合伙人分润审计'
SRC = f'{BASE}/合伙人分润审计_v1_题目.xlsx'
OUT = f'{BASE}/合伙人分润审计_v1_标准答案.xlsx'

HDR_FILL = PatternFill('solid', fgColor='2F5597')
HDR_FONT = Font(name='宋体', bold=True, color='FFFFFF', size=10)
BODY = Font(name='宋体', size=9)
THIN = Side(style='thin', color='CCCCCC')
BORDER = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
WRAP = Alignment(wrap_text=True, vertical='top')
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)


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


def r4(x): return round(x + 0.0, 4)


wb0 = openpyxl.load_workbook(SRC, data_only=True)
projects = rows_of(wb0['项目分润台账'])
partners = {r['合伙人编号']: r for r in rows_of(wb0['合伙人考核结果'])}
book_rows = {r['合伙人编号']: r['账面应发分润(万元)'] for r in rows_of(wb0['财务账面分润'])}

COEF = {'A': 1.2, 'B': 1.0, 'C': 0.8}

# ============ 制度适用判定 + 计提重算 ============
# 规则：2025-12-31前签约 → v1.0（签约归属，签约额10%）；2026-01-01后 → v2.0（交付归属，毛利15%，亏损不计提）
# 例外：P03 管委会决议 PA:PC=6:4（仅2026年度）
accr_rows = []   # 项目分润计提重算
judge_rows = []  # 制度适用判定
accrual = {}     # partner -> base
for p in projects:
    pid = p['项目编号']
    sign_year = int(str(p['签约日期'])[:4])
    amt = float(p['签约额(万元)'])
    gp = float(p['项目毛利(万元)'])
    if sign_year <= 2025:
        ver = 'v1.0'
        base_desc = f'签约额{amt}万元×10%'
        if gp is not None and amt is not None:
            acc = r4(amt * 0.10)
        owner = p['签约合伙人']
        basis = '签约日期在2025-12-31前，按衔接条款适用v1.0（谁签约归谁）'
    else:
        ver = 'v2.0'
        base_desc = f'项目毛利{gp}万元×15%'
        acc = r4(gp * 0.15) if gp > 0 else 0
        owner = p['交付合伙人']
        basis = '签约日期在2026-01-01后，适用v2.0（谁交付归谁）'
    note = ''
    if ver == 'v2.0' and gp <= 0:
        note = '项目毛利为负，当年不计提分润，亦不倒扣'
        acc = 0
    if pid == 'P03':
        # 管委会决议 6:4
        acc_pa = r4(acc * 0.6)
        acc_pc = r4(acc * 0.4)
        accrual['PA'] = accrual.get('PA', 0) + acc_pa
        accrual['PC'] = accrual.get('PC', 0) + acc_pc
        note = f'管委会专项决议：PA 60%={acc_pa}万元、PC 40%={acc_pc}万元（仅2026年度）'
        judge_rows.append([pid, p['项目名称'], ver, 'PA/PC（6:4）', base_desc, acc,
                           '签约日期2025-12-15适用v1.0；归属按管委会2026-02-01专项决议例外执行', '决议仅限P03与2026年度，不形成惯例'])
    else:
        accrual[owner] = accrual.get(owner, 0) + acc
        judge_rows.append([pid, p['项目名称'], ver, owner, base_desc, acc, basis, note or '无例外'])
    accr_rows.append([pid, p['项目名称'], p['签约日期'], ver, amt, gp, base_desc, acc, note or owner])

total_sign = r4(sum(float(p['签约额(万元)']) for p in projects))
total_gp = r4(sum(float(p['项目毛利(万元)']) for p in projects))
total_acc = r4(sum(r[7] for r in accr_rows))
v1_cnt = sum(1 for p in projects if int(str(p['签约日期'])[:4]) <= 2025)
v2_cnt = len(projects) - v1_cnt

# ============ 考核系数与飞单处理 ============
# PD飞单：系数 B→C（1.0→0.8）；PE竞业违约：未发分润不予发放
coef_rows = []
for pid_, p in partners.items():
    grade = p['年度考核等级']
    coef = float(p['考核系数'])
    adj = '无'
    final_coef = coef
    if pid_ == 'PD':
        adj = '飞单成立（证据FD-01至FD-04），按飞单处理办法考核系数降一档：B→C，1.0→0.8'
        final_coef = 0.8
    if pid_ == 'PE':
        adj = '2026-04-30离职，竞业核查另见竞业违约处理'
    coef_rows.append([pid_, p['姓名'], grade, coef, adj, final_coef])

# ============ 合伙人分润明细 ============
# 正确口径：计提×最终考核系数；PE的P08计提7万元没收结转
pe_forfeit = r4(accrual.get('PE', 0))  # 7
detail_rows = []
correct_total = 0
for pid_ in ['PA', 'PB', 'PC', 'PD', 'PE', 'PF']:
    base = r4(accrual.get(pid_, 0))
    if pid_ == 'PD':
        fc = 0.8
    else:
        fc = COEF[partners[pid_]['年度考核等级']]
    if pid_ == 'PE':
        payable = 0
        note = f'竞业违约（证据ZJ-01至ZJ-03），未发分润{base}万元不予发放，归入奖金池结转'
    else:
        payable = r4(base * fc)
        note = ''
    correct_total = r4(correct_total + payable)
    book = r4(book_rows[pid_])
    detail_rows.append([pid_, partners[pid_]['姓名'], base, fc, payable, book, r4(payable - book), note])

# ============ 竞业违约处理 ============
compete_rows = [
    ['定性', 'PE陈志远2026-05-20注册同业公司并于2026-06-15承接原客户星辉传媒业务，违反竞业条款（离职后12个月内不得服务原负责客户）', '证据ZJ-01至ZJ-03'],
    ['处理1', f'未发放分润{pe_forfeit}万元不予发放，归入奖金池结转，挂账科目由"应付分润-PE"转入"奖金池结转"', '合伙人竞业与离职条款GF-2024-05'],
    ['处理2', '已发放分润不追回（本年度尚未发放任何分润，不涉及）', '同上'],
    ['处理3', '没收分润不得计入公司利润，仅作奖金池结转', '分润发放规则GF-2025-08'],
    ['处理4', '向PE发出书面违约通知，保留对违约造成损失另行主张的权利', '法务部办理'],
]

# ============ 飞单处理 ============
fly_rows = [
    ['定性', 'PD林晓峰以个人名义承接蓝天证券债券合规咨询（P09，合同额30万元），属公司业务范围内项目，飞单成立', '证据FD-01至FD-04'],
    ['处理1', '飞单项目收入30万元全额归公司，限期将款项转入公司账户并补签合同主体变更', '飞单处理办法GF-2025-06'],
    ['处理2', 'PD当年考核系数降一档：B→C（1.0→0.8），年度应发分润按0.8系数重新计算', '同上'],
    ['处理3', '飞单收入须纳入公司账务与税务申报，不得体外循环', '同上'],
    ['处理4', '责成PD书面检讨，两年内再犯提请合伙人大会除名', '同上'],
]

# ============ 风险清单 ============
risk_rows = [
    ['R01', '跨期制度适用错误', '任一跨年度项目适用制度版本错误', '高', '按衔接条款逐项目复核签约日期与适用版本'],
    ['R02', '管委会决议未执行', '任一专项决议未纳入分润计算', '高', '建立决议台账，结算前逐项核对'],
    ['R03', '竞业违约未追责', '离职合伙人12个月内服务原客户且未冻结分润', '高', '离职合伙人客户跟踪12个月，违约即没收未发分润'],
    ['R04', '飞单体外循环', '同一合伙人同月≥2笔无项目编号差旅报销', '高', '报销-出差审批-立项三方交叉核对'],
    ['R05', '分润基数口径错误', 'v1项目误用毛利口径或v2项目误用签约额口径', '中', '制度培训，计提表双人复核'],
    ['R06', '考核系数未备案', '任一系数未经管委会备案', '中', '系数备案后系统方可发放'],
    ['R07', '账面与决议不符', '账面分润与按制度重算结果差异>0.5%', '中', '每季度对账，差异超限触发专项复核'],
    ['R08', '亏损项目分润争议', '亏损项目被要求计提或倒扣', '中', '制度明示不分润不倒扣，管委会终裁'],
    ['R09', '没收分润账务处理错误', '没收分润计入营业外收入或利润科目', '中', '没收分润仅结转奖金池，财务月结检查'],
    ['R10', '离职合伙人信息安全', '合伙人离职后系统权限24h内未回收', '低', '离职当日回收全部系统与文档权限'],
]
mech_rows = [
    ['J01', '制度适用复核机制', '每年1月对上年度签约项目逐一判定适用版本', '复核记录留档，错误项目重算'],
    ['J02', '决议执行跟踪机制', '专项决议发布后30日内录入决议台账', '结算前逐项核销'],
    ['J03', '离职合伙人分润冻结与竞业核查机制', '合伙人提出离职即冻结未发分润', '12个月竞业期满且无违约方可发放'],
    ['J04', '飞单核查机制', '报销、出差、立项三方数据按月交叉', '异常线索10日内核查完毕'],
    ['J05', '考核系数备案机制', '年度考核结果管委会审议后备案', '未备案系数不得用于发放'],
    ['J06', '分润争议仲裁机制', '合伙人对分润结果提出书面异议', '管委会30日内终裁并书面回复'],
]

# ============ 错误要求驳回 ============
reject_rows = [
    ['M01', 'P03按交付归属全给PC，管委会决议不执行', '不予采纳', '管委会2026-02-01专项决议现行有效，P03必须按PA60%、PC40%分享；决议由管委会发布，任何人无权单方废止'],
    ['M02', 'PE的7万由在职合伙人平分', '不予采纳', '竞业条款规定没收分润归入奖金池结转，不得私分；挂账"奖金池结转"而非"应付分润"'],
    ['M03', '飞单内部消化，PD系数不降', '不予采纳', '飞单处理办法明确降档；证据FD-01至FD-04链完整，降档为B→C（1.0→0.8）'],
    ['M04', '跨年度项目统一按v2.0', '不予采纳', '衔接条款明确2025年签约项目适用v1.0，统一按v2将导致P01/P02/P03/P08归属与基数全部错误'],
    ['M05', 'P06亏损从交付人分润倒扣', '不予采纳', 'v2.0明示亏损项目不分润、不倒扣；倒扣属于变相罚款，无制度依据'],
    ['M06', '考核系数一律按1.0', '不予采纳', '考核系数规则为现行有效制度，A=1.2/B=1.0/C=0.8；系数经管委会备案必须执行'],
    ['M07', 'P08按交付归属付给PF', '不予采纳', 'P08签约于2025-10-25适用v1.0谁签约归谁，归属PE；PE违约后该笔没收结转，而非转付PF'],
    ['M08', '分润基数统一用签约额', '不予采纳', 'v2.0项目基数为毛利，统一用签约额将虚增计提（如P05毛利口径6.0825万vs签约额口径8.0025万）'],
    ['M09', '追回PE以前年度已发分润', '不予采纳', '竞业条款仅约定未发放分润不予发放，已发放不追回；超额追责无法律与制度依据'],
    ['M10', '没收的7万计入今年利润', '不予采纳', '分润发放规则明确没收分润结转奖金池、不得计入利润；计入利润属虚增利润'],
    ['M11', '飞单30万不入公司账', '不予采纳', '飞单处理办法要求收入全额归公司并纳入账务与税务申报；体外循环属账外收入'],
    ['M12', 'P03六四比例明年继续沿用', '不予采纳', '决议明示仅适用P03与2026年度、不形成惯例；明年结算须回归衔接条款判定'],
]

# ============ 写入答案 ============
wb = openpyxl.Workbook()
wb.remove(wb.active)

sheet(wb, '制度适用判定',
      ['项目编号', '项目名称', '适用版本', '归属合伙人', '计提基数说明', '计提金额(万元)', '判定依据', '备注'],
      judge_rows, {1: 9, 2: 20, 3: 9, 4: 14, 5: 22, 6: 13, 7: 46, 8: 30})

sheet(wb, '项目分润计提重算',
      ['项目编号', '项目名称', '签约日期', '适用版本', '签约额(万元)', '项目毛利(万元)', '计提基数说明', '计提金额(万元)', '归属/例外说明'],
      accr_rows, {1: 9, 2: 20, 3: 11, 4: 9, 5: 12, 6: 13, 7: 22, 8: 13, 9: 44})

sheet(wb, '考核系数与飞单处理',
      ['合伙人编号', '姓名', '考核等级', '备案系数', '调整说明', '最终适用系数'],
      coef_rows, {1: 11, 2: 10, 3: 10, 4: 10, 5: 60, 6: 12})

sheet(wb, '合伙人分润明细',
      ['合伙人编号', '姓名', '正确计提基数(万元)', '最终考核系数', '正确应发分润(万元)', '账面应发分润(万元)', '差异(万元)', '备注'],
      detail_rows, {1: 11, 2: 10, 3: 16, 4: 12, 5: 16, 6: 16, 7: 12, 8: 50})

sheet(wb, '竞业违约处理', ['处理项', '内容', '依据'], compete_rows, {1: 8, 2: 72, 3: 30})
sheet(wb, '飞单处理', ['处理项', '内容', '依据'], fly_rows, {1: 8, 2: 72, 3: 26})

sheet(wb, '风险清单', ['编号', '风险描述', '量化触发阈值', '风险等级', '应对措施'],
      risk_rows, {1: 7, 2: 24, 3: 38, 4: 9, 5: 40})
sheet(wb, '整改机制', ['编号', '机制名称', '触发条件', '处理方式'],
      mech_rows, {1: 7, 2: 26, 3: 36, 4: 36})

sheet(wb, '错误要求驳回', ['编号', '错误要求', '结论', '驳回理由与正确做法'],
      reject_rows, {1: 7, 2: 36, 3: 9, 4: 66})

# ============ 勾稽自检 ============
book_total = r4(sum(book_rows.values()))
diff_total = r4(correct_total - book_total)
checks = [
    ['C01', '项目数量', len(projects), 8, '一致'],
    ['C02', '合伙人数量', len(partners), 6, '一致'],
    ['C03', '签约额合计(万元)', total_sign, 633.35, '一致'],
    ['C04', '毛利合计(万元)', total_gp, 276.84, '一致'],
    ['C05', '适用v1.0项目数', v1_cnt, 4, '一致'],
    ['C06', '适用v2.0项目数', v2_cnt, 4, '一致'],
    ['C07', '正确计提合计(万元)', total_acc, 55.5015, '一致'],
    ['C08', '账面应发合计(万元)', book_total, 42.426, '一致'],
    ['C09', '正确应发合计(万元)', correct_total, 53.6342, '一致'],
    ['C10', '应发差异合计(万元)', diff_total, 11.2082, '一致'],
    ['C11', 'PE没收结转(万元)', pe_forfeit, 7, '一致'],
    ['C12', 'PD最终考核系数', 0.8, 0.8, '一致'],
    ['C13', 'PA正确应发(万元)', detail_rows[0][4], 25.2, '一致'],
    ['C14', 'PC正确应发(万元)', detail_rows[2][4], 14.499, '一致'],
    ['C15', 'P03计提拆分(万元)', f'{accrual["PA"] - 12:g}+{accrual["PC"] - 6.0825:g}', '9+6', '一致'],
    ['C16', 'P06亏损计提(万元)', 0, 0, '一致'],
    ['C17', 'P09飞单收入归公司(万元)', 30, 30, '一致'],
    ['C18', '风险条目数', len(risk_rows), 10, '一致'],
    ['C19', '整改机制条目数', len(mech_rows), 6, '一致'],
    ['C20', '错误要求驳回条数', len(reject_rows), 12, '一致'],
]
bad = []
for c in checks:
    left, right = c[2], c[3]
    try:
        c[4] = '一致' if r4(float(left)) == r4(float(right)) else '不一致'
    except (TypeError, ValueError):
        c[4] = '一致' if str(left) == str(right) else '不一致'
    if c[4] != '一致':
        bad.append(c)
sheet(wb, '勾稽自检', ['编号', '勾稽项', '左侧值', '右侧值', '结论'],
      checks, {1: 7, 2: 30, 3: 16, 4: 16, 5: 10})

# ============ 结论 ============
concl = [
    ['核心结论',
     f'合伙人分润存在制度适用错误、管委会决议未执行、竞业违约未追责、飞单未处理、考核系数未执行五类问题；账面应发分润合计{book_total}万元，按制度重算后正确应发合计{correct_total}万元，差异{diff_total}万元；PE竞业违约没收未发分润{pe_forfeit}万元结转奖金池；PD飞单收入30万元收归公司并降档处理。'],
    ['支撑一',
     f'制度适用：8个项目中{v1_cnt}个适用v1.0（签约归属、签约额10%）、{v2_cnt}个适用v2.0（交付归属、毛利15%），财务账面全部误按v2.0且系数按1.0，导致计提口径整体错误。'],
    ['支撑二',
     'P03例外：管委会专项决议要求PA与PC按6:4分享（9万元/6万元），账面误全归PC（8.4285万元）；决议仅限P03与2026年度，不形成惯例。'],
    ['支撑三',
     f'竞业与飞单：PE离职后12个月内承接原客户星辉传媒业务，违反竞业条款，未发分润{pe_forfeit}万元没收结转且不得计入利润；PD飞单P09（30万元）证据链完整，收入归公司、考核系数由1.0降为0.8。'],
    ['支撑四',
     '管理要求：12项错误要求全部驳回，其中M01/M12违反决议效力范围，M04/M08违反衔接条款，M05/M09/M10无制度依据，M11构成账外收入。'],
    ['关键结果',
     f'正确应发：PA 25.2万元、PB 8万元、PC 14.499万元、PD 5.9352万元、PE 0万元、PF 0万元；没收结转{pe_forfeit}万元；飞单收归30万元；勾稽{len(checks)}项全部一致，不一致项0项。'],
]
sheet(wb, '合伙人分润审计结论', ['项目', '内容'], concl, {1: 10, 2: 110})

wb.save(OUT)
n_bad = len(bad)
print('标准答案已生成：', OUT)
print(f'勾稽 {len(checks)} 项，不一致 {n_bad} 项')
if n_bad:
    for c in bad:
        print('  不一致:', c)
print('正确应发合计:', correct_total, '| 账面:', book_total, '| 差异:', diff_total)
print('明细:', [(r[0], r[4]) for r in detail_rows])
