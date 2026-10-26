#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""加盟商准入与分级续约 v1 附件拆分与打包：8输入 + 8答案，并生成题目 docx"""
import os
import io
import shutil
import zipfile
import time
from copy import copy

import openpyxl
from openpyxl import Workbook
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

BASE = '/workspace/高经济价值第六期_加盟商准入续约'
SRC_Q = f'{BASE}/加盟商准入续约_v1_题目.xlsx'
SRC_A = f'{BASE}/加盟商准入续约_v1_标准答案.xlsx'
SPLIT = '/tmp/storeC_split'
OUT = f'{BASE}/加盟商准入续约_v1_附件'
shutil.rmtree(SPLIT, ignore_errors=True)
shutil.rmtree(OUT, ignore_errors=True)
os.makedirs(f'{SPLIT}/input', exist_ok=True)
os.makedirs(f'{SPLIT}/ans', exist_ok=True)
os.makedirs(f'{OUT}/输入附件', exist_ok=True)
os.makedirs(f'{OUT}/标准答案', exist_ok=True)

EXCEL_APP = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
             '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
             'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
             '<Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>'
             '<Company>麦淳餐饮管理有限公司</Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc>'
             '<HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0300</AppVersion></Properties>')
WORD_APP = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
            'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
            '<Application>Microsoft Office Word</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>'
            '<Company>麦淳餐饮管理有限公司</Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc>'
            '<HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0000</AppVersion></Properties>')


def r2(x): return round(x + 0.0, 2)


def core_xml(author, created, modified):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            f'<dc:creator>{author}</dc:creator><cp:lastModifiedBy>{author}</cp:lastModifiedBy>'
            f'<dcterms:created xsi:type="dcterms:W3CDTF">{created}</dcterms:created>'
            f'<dcterms:modified xsi:type="dcterms:W3CDTF">{modified}</dcterms:modified></cp:coreProperties>')


def scrub(path, app_xml, author, created, modified, zip_dt):
    with zipfile.ZipFile(path) as zf:
        items = [(i, zf.read(i.filename)) for i in zf.infolist()]
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for info, data in items:
            if info.filename == 'docProps/app.xml':
                data = app_xml.encode('utf-8')
            elif info.filename == 'docProps/core.xml':
                data = core_xml(author, created, modified).encode('utf-8')
            zi = zipfile.ZipInfo(info.filename, date_time=zip_dt)
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = info.external_attr
            zf.writestr(zi, data)
    with open(path, 'wb') as f:
        f.write(buf.getvalue())
    ts = time.mktime((zip_dt[0], zip_dt[1], zip_dt[2], zip_dt[3], zip_dt[4], zip_dt[5], 0, 0, -1))
    os.utime(path, (ts, ts))


def copy_sheet(src, wb, title=None):
    dst = wb.create_sheet(title or src.title)
    for row in src.iter_rows():
        for cell in row:
            d = dst.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                d.font = copy(cell.font)
                d.fill = copy(cell.fill)
                d.border = copy(cell.border)
                d.alignment = copy(cell.alignment)
                d.protection = copy(cell.protection)
                d.number_format = cell.number_format
    for col, dim in src.column_dimensions.items():
        dst.column_dimensions[col].width = dim.width
    for row, dim in src.row_dimensions.items():
        dst.row_dimensions[row].height = dim.height
    dst.freeze_panes = src.freeze_panes
    dst.sheet_view.showGridLines = False
    for rng in src.merged_cells.ranges:
        dst.merge_cells(str(rng))
    return dst


def make_rar(files, rar_path, zip_dt):
    with zipfile.ZipFile(rar_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for f in files:
            data = open(f, 'rb').read()
            zi = zipfile.ZipInfo(os.path.basename(f), date_time=zip_dt)
            zi.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(zi, data)


def split_workbook(src_path, packs, out_dir, prefix, created, modified, zip_dt):
    wb0 = openpyxl.load_workbook(src_path)
    for name, sheets in packs:
        wb = Workbook()
        wb.remove(wb.active)
        for sh in sheets:
            copy_sheet(wb0[sh], wb, sh)
        tmp = f'{SPLIT}/{prefix}_{name}.xlsx'
        wb.save(tmp)
        scrub(tmp, EXCEL_APP, '运营分析部', created, modified, zip_dt)
        make_rar([tmp], f'{out_dir}/{name}.rar', zip_dt)
        print('packed', name, len(sheets))


INPUT_PACKS = [
    ('01新加盟商申请', ['新加盟商申请']),
    ('02现有加盟门店', ['现有加盟门店']),
    ('03准入评分标准', ['准入评分标准']),
    ('04续约评级标准', ['续约评级标准']),
    ('05加盟合同条款', ['加盟合同条款']),
    ('06城市与商圈系数', ['城市分级系数', '商圈等级系数']),
    ('07财务确认规则', ['财务确认规则']),
    ('08管理层要求', ['管理层要求']),
]

ANSWER_PACKS = [
    ('01准入评分与排序', ['准入评分与排序']),
    ('02现有门店评级', ['现有门店评级']),
    ('03授权费与保证金', ['授权费与保证金']),
    ('04续约与整改建议', ['续约与整改建议']),
    ('05公司供应链毛利', ['公司供应链毛利']),
    ('06风险与退出机制', ['风险清单', '退出机制']),
    ('07错误要求驳回', ['错误要求驳回']),
    ('08勾稽与结论', ['勾稽自检', '加盟准入续约结论']),
]

IN_DT = (2026, 1, 5, 9, 0, 0)
AN_DT = (2026, 1, 15, 14, 0, 0)
split_workbook(SRC_Q, INPUT_PACKS, f'{OUT}/输入附件', 'in', '2026-01-04T01:00:00Z', '2026-01-05T01:00:00Z', IN_DT)
split_workbook(SRC_A, ANSWER_PACKS, f'{OUT}/标准答案', 'an', '2026-01-13T02:00:00Z', '2026-01-15T06:00:00Z', AN_DT)

# 读取答案指标
ans = openpyxl.load_workbook(SRC_A, data_only=True)
concl = {r[0]: r[1] for r in ans['加盟准入续约结论'].iter_rows(values_only=True) if r[0]}
check_map = {r[0]: r[1] for r in ans['勾稽自检'].iter_rows(values_only=True) if r[0] != '勾稽项'}

app_rows = list(ans['准入评分与排序'].iter_rows(min_row=2, values_only=True))
store_rows = list(ans['现有门店评级'].iter_rows(min_row=2, values_only=True))
on_list = [r[0] for r in app_rows if r[11] == '准入']
ref_list = [r[0] for r in app_rows if r[11] == '整改后准入']
rej_list = [r[0] for r in app_rows if r[11] == '拒绝']
grade_counts = {}
for r in store_rows:
    grade_counts[r[8]] = grade_counts.get(r[8], 0) + 1
supply_total = sum(r[5] for r in ans['公司供应链毛利'].iter_rows(min_row=2, values_only=True))
mgmt_total = sum(r[4] for r in ans['公司供应链毛利'].iter_rows(min_row=2, values_only=True))
check_n = len(check_map)

# 题目 docx
def new_doc():
    d = Document()
    st = d.styles['Normal']
    st.font.name = '宋体'
    st.font.size = Pt(10.5)
    st.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    for sec in d.sections:
        sec.top_margin = sec.bottom_margin = Pt(56)
        sec.left_margin = sec.right_margin = Pt(64)
    return d


def H(d, text, level=1):
    p = d.add_heading(text, level if level else 0)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    for run in p.runs:
        run.font.name = '宋体'
        run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
        run.font.color.rgb = None
    return p


def P(d, text, bold=False, indent=True):
    p = d.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.25
    if indent:
        p.paragraph_format.first_line_indent = Pt(21)
    run = p.add_run(text)
    run.font.name = '宋体'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    run.font.size = Pt(10.5)
    run.bold = bold
    return p


doc = new_doc()
H(doc, '加盟商准入与分级续约决策', 0)

H(doc, '一、编号')
P(doc, 'G6-ZX-006（v1）。', indent=False)
P(doc, '【职业定位】二级行业：餐饮连锁运营管理 / 企业管理与运营咨询；L1职业簇：咨询与战略、财务与审计、运营管理；L2职业：管理咨询顾问、财务分析师、审计专员、加盟运营经理、门店经营分析师。', indent=False)

H(doc, '二、任务背景')
P(doc, '麦淳餐饮管理有限公司采用直营与加盟并行的扩张模式，截至2025年底已在苏州、无锡、宁波、杭州、合肥、南昌六城开设多家加盟门店。2026年初，公司收到5个新加盟商/点位申请，运营部倾向于快速扩张、尽快落店；财务部则担心部分申请人资金实力和合规记录不足，且现有8家加盟门店中已有2家连续两年未达营收目标、客诉率偏高。经运营总监与财务总监共同商议，决定委托加盟管理岗对5个新申请做准入评分，对8家现有门店做年度续约评级，并同步测算授权费、保证金、管理费与供应链毛利。')
P(doc, '本次分析存在三个现实张力：一是运营希望简化准入流程、财务部坚持四维度评分；二是部分加盟商认为授权费应按统一价收取、不应区分城市商圈；三是M05、M07两家门店合同即将到期，运营主张再给一次机会、财务部要求按评级结果处置。所有结论须基于附件数据逐条计算，不得凭印象或关系判断，也不得只给汇总数。')
P(doc, '数据截止日为2025年12月31日，分析基准日为2026年1月10日。新加盟商准入须按资金实力、选址评分、运营经验、历史合规四维度加权评分；现有门店评级须按营收达成率、客诉率、食安扣分、巡店评分四项标准就低判定。')

H(doc, '三、任务分解')
P(doc, '本项委托按八轮执行，每轮只完成一个任务，并只提交当轮对应交付包。上一轮自检通过后，方可进入下一轮；未完成或自检不通过，不得进入下一轮，也不得提前输出后续轮次结果。')
P(doc, '第1轮为准入评分与排序。对5个新加盟商申请按四维度加权评分，计算总分、排名并给出准入/整改/拒绝结论。本轮只输出01准入评分与排序.rar。本轮自检要求5个申请均有加权总分和排名，权重合计100%。')
P(doc, '第2轮为现有门店评级。对8家现有加盟门店按营收达成率、客诉率、食安扣分、巡店评分判定A/B/C/D级。本轮只输出02现有门店评级.rar。本轮自检要求8家门店均有评级，标准判定采用就低原则。')
P(doc, '第3轮为授权费与保证金。按城市系数、商圈系数、门店面积计算5个新加盟商的授权费，并列出保证金。本轮只输出03授权费与保证金.rar。本轮自检要求授权费计算与合同条款一致。')
P(doc, '第4轮为续约与整改建议。按A/B/C/D级给出8家现有门店的续约、整改或收回建议，并明确时间节点。本轮只输出04续约与整改建议.rar。本轮自检要求D级门店有收回路径，C级有整改期。')
P(doc, '第5轮为公司供应链毛利。按食材成本×20%÷120%计算公司从8家门店获得的供应链毛利，并汇总管理费。本轮只输出05公司供应链毛利.rar。本轮自检要求毛利口径与合同一致。')
P(doc, '第6轮为风险与退出机制。列出10项风险、量化触发阈值、应对措施和退出动作。本轮只输出06风险与退出机制.rar。本轮自检要求10项风险均有阈值，6类退出机制均有回收与损失处理。')
P(doc, '第7轮为错误要求驳回。逐条驳回管理层12项错误要求并说明理由。本轮只输出07错误要求驳回.rar。本轮自检要求12项均有驳回结论。')
P(doc, '第8轮为勾稽与结论。完成勾稽自检并给出核心结论。本轮只输出08勾稽与结论.rar。本轮自检要求勾稽全部闭合、不一致项为0；未闭合不得出具结论。')

H(doc, '四、输入附件')
P(doc, '输入附件共8包。')
P(doc, '01新加盟商申请.rar，包含新加盟商申请。')
P(doc, '02现有加盟门店.rar，包含现有加盟门店年度经营数据。')
P(doc, '03准入评分标准.rar，包含准入评分标准。')
P(doc, '04续约评级标准.rar，包含续约评级标准。')
P(doc, '05加盟合同条款.rar，包含加盟合同条款。')
P(doc, '06城市与商圈系数.rar，包含城市分级系数、商圈等级系数。')
P(doc, '07财务确认规则.rar，包含财务确认规则。')
P(doc, '08管理层要求.rar，包含管理层要求。')

H(doc, '五、制度与判断规则')
P(doc, '新加盟商准入评分采用四维度加权：资金实力权重25%、选址评分权重30%、运营经验权重25%、历史合规权重20%，满分100分。加权总分≥80分直接准入，60-80分整改后准入，<60分拒绝。')
P(doc, '现有门店年度评级采用就低原则，须同时满足该等级全部标准方可定级；任一指标不达标即降至下一等级。A级标准为营收达成率≥100%、客诉率≤0.5%、食安扣分=0、巡店评分≥90；B级为营收达成率≥90%、客诉率≤1%、食安扣分≤2、巡店评分≥80；C级为营收达成率≥80%、客诉率≤2%、食安扣分≤5、巡店评分≥70；低于C级任一标准为D级。')
P(doc, 'A级门店正常续约并享有优先续约权；B级门店正常续约；C级门店限期3个月整改，到期复核后续约；D级门店合同到期后收回加盟权或转直营。')
P(doc, '授权费等于门店面积×500元/㎡×城市系数×商圈系数。城市系数：一线（苏州、杭州）1.20，二线（无锡、宁波、合肥）1.00，三线（南昌）0.80。商圈系数：A级1.15、B级1.00、C级0.85。')
P(doc, '保证金为10万元/店，一次性缴纳，合同终止无违约则退还；不得充抵货款或挪作他用。')
P(doc, '年度管理费按门店年度营业收入的3%计提。公司供应链毛利按门店食材成本×20%÷120%计算，即公司在食材采购额中获得20%的毛利份额。')
P(doc, '续约谈判须在合同到期前3个月启动，C级整改期3个月，D级门店不得直接收回须按合同到期后执行。')

H(doc, '六、交付物')
for t in ['01 准入评分与排序.rar', '02 现有门店评级.rar', '03 授权费与保证金.rar',
          '04 续约与整改建议.rar', '05 公司供应链毛利.rar', '06 风险与退出机制.rar',
          '07 错误要求驳回.rar', '08 勾稽与结论.rar']:
    P(doc, t, indent=False)

H(doc, '七、约束条件')
constraints = [
    '1. 仅使用附件材料，不得引入外部数据，不得联网检索。',
    '2. 金额以万元或元列示并保留两位小数，比例保留两位小数，不得混用单位。',
    '3. 逐申请、逐门店给出结论，不得抽样或估算；每条记录必须有明确状态，不得留空。',
    '4. 5个新加盟商编号唯一，8家现有门店编号唯一，不得重复或遗漏。',
    '5. 准入评分权重合计必须等于100%，四维度缺一不可。',
    '6. 现有门店评级采用就低原则，任一指标不达标即降级。',
    '7. 授权费须按城市系数和商圈系数差异化计算，不得固定价。',
    '8. 保证金10万元/店，不得免除或充抵货款。',
    '9. 年度管理费按门店营业收入3%计提，不得按利润或其他口径。',
    '10. 公司供应链毛利按食材成本×20%÷120%计算，不得按营业额计提。',
    '11. 续约窗口期为合同到期前3个月，不得到期后才通知。',
    '12. C级门店整改期为3个月，到期复核；D级门店合同到期后收回，不得提前强制收回。',
    '13. 有违规记录的加盟商不得直接拒绝，1次轻微违规可整改后准入。',
    '14. 加盟门店独立核算，公司不承担亏损，不得补贴。',
    '15. 巡店评分须由公司评定，不得由加盟商自评。',
    '16. 客诉率须直接采用门店年度客诉率，不得按投诉次数简单估算。',
    '17. 必须列出10项风险和6类退出机制，每项风险须有量化触发阈值。',
    '18. 勾稽全部闭合后方可出具核心结论；任一项不闭合视为未完成。',
    '19. 必须按八轮顺序逐轮交付，每轮只完成一个任务，不得一次性输出全部结果。',
    '20. 每轮结束必须提交本轮自检记录，列明检查项、结果、证据和未通过原因；上一轮未通过不得进入下一轮。',
    '21. 每轮只提交当轮对应交付包，不得跨轮合并或提前输出；任何一轮无法完成时必须停止并说明缺口；八轮全部通过方视为完成。',
    '22. 管理层12项错误要求须逐条驳回并说明理由，不得采纳。',
    '23. 所有续约与整改建议必须有明确触发条件、时间节点和责任人，不得只给方向性意见。',
    '24. 供应链毛利和管理费须逐店计算并汇总，不得只给总数。',
]
for t in constraints:
    P(doc, t, indent=False)

H(doc, '八、完成标准')
P(doc, f'八轮任务逐轮完成且逐轮自检通过。5个新加盟商中{len(on_list)}家直接准入（{"、".join(on_list)}）、{len(ref_list)}家整改后准入（{"、".join(ref_list)}）、{len(rej_list)}家拒绝；8家现有门店A级{grade_counts.get("A",0)}家、B级{grade_counts.get("B",0)}家、C级{grade_counts.get("C",0)}家、D级{grade_counts.get("D",0)}家；公司年度供应链毛利{r2(supply_total)}万元、管理费{r2(mgmt_total)}万元；{check_n}项勾稽全部闭合、不一致项为0。任一轮未完成或自检不通过，不得进入下一轮，整体不得视为完成。')

H(doc, '九、考点')
P(doc, '以下评分依据按数据治理、准入评分、续约评级、授权费测算、供应链毛利、续约建议、风险退出、错误驳回、勾稽结论九组展开，共50项，用于逐条评定答题底稿的完整性与正确性。', indent=False)
groups = [
    ('（一）数据治理（评分依据1—5）', [
        '评分依据1：底稿中应当列示5个新加盟商申请，编号J01-J05，并分别标明所在城市与商圈等级。',
        '评分依据2：底稿中应当列示8家现有加盟门店年度经营数据，编号M01-M08，覆盖营收、成本、客诉、食安、巡店等字段。',
        '评分依据3：底稿中应当明确准入评分四维度，即资金实力、选址评分、运营经验、历史合规。',
        '评分依据4：底稿中应当明确续约评级四指标，即营收达成率、客诉率、食安扣分、巡店评分。',
        '评分依据5：底稿中应当列示合同关键条款，含授权费单价、保证金、年度管理费率、供应链加价率、续约窗口期、整改期。',
    ]),
    ('（二）准入评分（评分依据6—15）', [
        '评分依据6：底稿中应当按加权总分=Σ（维度得分×维度权重）计算，满分100分。',
        '评分依据7：底稿中应当列明资金实力权重25%、选址评分30%、运营经验25%、历史合规20%。',
        '评分依据8：底稿中应当给出评分结论，≥80分直接准入，60-80分整改后准入，<60分拒绝。',
        f'评分依据9：底稿中应当计算并给出5个申请的准入结论，明确{len(on_list)}家直接准入（{"、".join(on_list)}）、{len(ref_list)}家整改后准入（{"、".join(ref_list)}）、{len(rej_list)}家拒绝。',
        '评分依据10：底稿中应当按加权总分对5个申请进行高低排序。',
        '评分依据11：底稿中应当说明历史合规得分规则，无违规100分、1次轻微违规85分、2次违规70分。',
        '评分依据12：底稿中应当指出有违规记录不得直接拒绝，须结合违规次数与性质判定。',
        '评分依据13：底稿中应当将选址评分与商圈等级系数交叉核对。',
        '评分依据14：底稿中应当区分运营经验评分中的已有门店数量与从业经验。',
        '评分依据15：底稿中应当将资金实力评分与计划投资额进行交叉核对。',
    ]),
    ('（三）续约评级（评分依据16—25）', [
        '评分依据16：底稿中应当按营收达成率=实际营收÷预算营收×100%计算。',
        '评分依据17：底稿中应当直接采用门店年度客诉率（%）指标。',
        '评分依据18：底稿中应当采用就低原则判定评级，四项指标须同时满足。',
        '评分依据19：底稿中应当明确A级标准：营收达成率≥100%、客诉率≤0.5%、食安扣分=0、巡店评分≥90。',
        '评分依据20：底稿中应当明确B级标准：营收达成率≥90%、客诉率≤1%、食安扣分≤2、巡店评分≥80。',
        '评分依据21：底稿中应当明确C级标准：营收达成率≥80%、客诉率≤2%、食安扣分≤5、巡店评分≥70。',
        '评分依据22：底稿中应当明确低于C级任一标准即判定为D级。',
        f'评分依据23：底稿中应当给出8家门店评级结果，A级{grade_counts.get("A",0)}家、B级{grade_counts.get("B",0)}家、C级{grade_counts.get("C",0)}家、D级{grade_counts.get("D",0)}家。',
        '评分依据24：底稿中应当指出D级门店须合同到期后收回，不得提前强制收回。',
        '评分依据25：底稿中应当明确C级门店整改期为3个月，到期复核后续约。',
    ]),
    ('（四）授权费与保证金（评分依据26—30）', [
        '评分依据26：底稿中应当按授权费=面积×500×城市系数×商圈系数计算。',
        '评分依据27：底稿中应当列明城市系数，一线1.20、二线1.00、三线0.80。',
        '评分依据28：底稿中应当列明商圈系数，A级1.15、B级1.00、C级0.85。',
        '评分依据29：底稿中应当列示保证金10万元/店，并指出不得免除或充抵货款。',
        '评分依据30：底稿中应当体现授权费按不同城市、商圈差异化计算。',
    ]),
    ('（五）公司供应链毛利（评分依据31—35）', [
        '评分依据31：底稿中应当按公司供应链毛利=食材成本×20%÷120%计算。',
        '评分依据32：底稿中应当按年度管理费=年度营业收入×3%计算。',
        f'评分依据33：底稿中应当汇总8家门店年度供应链毛利合计{r2(supply_total)}万元。',
        f'评分依据34：底稿中应当汇总8家门店年度管理费合计{r2(mgmt_total)}万元。',
        f'评分依据35：底稿中应当汇总公司从加盟业务获得供应链毛利与管理费合计{r2(supply_total+mgmt_total)}万元。',
    ]),
    ('（六）续约与整改建议（评分依据36—40）', [
        '评分依据36：底稿中应当对A级门店给出正常续约并享有优先续约权的建议。',
        '评分依据37：底稿中应当对B级门店给出正常续约的建议。',
        '评分依据38：底稿中应当对C级门店给出限期3个月整改、到期复核后续约的建议。',
        '评分依据39：底稿中应当对D级门店给出合同到期收回加盟权或转直营的建议。',
        '评分依据40：底稿中应当明确续约建议的时间节点和触发条件。',
    ]),
    ('（七）风险与退出（评分依据41—45）', [
        '评分依据41：底稿中应当列示10项风险，含加盟商资金链断裂、选址失误、食安事故、客诉率超标、违规采购、合同到期拒不续约、城市商圈降级、管理费拖欠、数据造假、竞品恶性挖角。',
        '评分依据42：底稿中应当对每项风险给出量化触发阈值。',
        '评分依据43：底稿中应当对每项风险给出具体可执行的应对措施。',
        '评分依据44：底稿中应当列示6类退出机制，含收回加盟权、停售整改、引入新加盟商、暂停供货、转址、合同终止。',
        '评分依据45：底稿中应当对每类退出机制给出回收与损失处理方案。',
    ]),
    ('（八）错误要求驳回（评分依据46—48）', [
        '评分依据46：底稿中应当逐条识别管理层12项错误要求，并给出不予采纳结论。',
        '评分依据47：底稿中应当识别以下错误要求：只看资金、直接拒绝违规者、只看营收、固定授权费、保证金充抵货款、D级直接收回、毛利按营业额、客诉按次数、自评巡店、到期不通知、免保证金、公司补贴亏损。',
        '评分依据48：底稿中应当对每项错误要求说明正确做法。',
    ]),
    ('（九）勾稽与结论（评分依据49—50）', [
        f'评分依据49：底稿中应当完成{check_n}项勾稽，覆盖申请数、门店数、维度、权重、准入分类、评级分布、合同条款、毛利口径，所有勾稽结论应为一致。',
        '评分依据50：底稿中应当给出核心结论，包括5个新加盟商准入情况、8家现有门店评级分布、公司供应链毛利与管理费合计。',
    ]),
]
for title, items in groups:
    H(doc, title, 2)
    for t in items:
        P(doc, t, indent=False)

H(doc, '十、答题思路')
P(doc, '第一步是读取新加盟商申请、现有门店数据、准入评分标准、续约评级标准、合同条款、城市与商圈系数。')
P(doc, '第二步是按四维度加权计算5个新加盟商的准入总分，排序后给出准入/整改/拒绝结论。')
P(doc, '第三步是按就低原则对8家现有门店进行A/B/C/D评级，识别各级门店名单。')
P(doc, '第四步是按城市系数、商圈系数、面积计算新加盟商授权费，并列出保证金。')
P(doc, '第五步是按A/B/C/D级给出续约、整改、收回建议，并明确时间节点。')
P(doc, '第六步是按食材成本×20%÷120%计算公司供应链毛利，按营收3%计算管理费，逐店汇总。')
P(doc, '第七步是列出10项风险、6类退出机制，量化触发阈值与回收处理。')
P(doc, '第八步是逐条驳回12项管理层错误要求，完成勾稽自检并给出核心结论。')

H(doc, '十一、标准答案')
P(doc, '标准答案共8个包，首表为加盟准入续约结论，逐包覆盖准入评分、门店评级、授权费保证金、续约建议、供应链毛利、风险退出、错误驳回和勾稽自检。', indent=False)
P(doc, concl['核心结论'], bold=True)
for k in ['支撑一', '支撑二', '支撑三', '支撑四']:
    if k in concl:
        P(doc, concl[k])
P(doc, f'关键结果：5个新加盟商中{len(on_list)}家直接准入（{"、".join(on_list)}）、{len(ref_list)}家整改后准入（{"、".join(ref_list)}）、{len(rej_list)}家拒绝；8家现有门店A级{grade_counts.get("A",0)}家、B级{grade_counts.get("B",0)}家、C级{grade_counts.get("C",0)}家、D级{grade_counts.get("D",0)}家；公司年度供应链毛利{r2(supply_total)}万元、管理费{r2(mgmt_total)}万元；{check_n}项勾稽全部闭合，不一致项为0。')

doc.core_properties.author = '运营分析部'
doc.save(f'{OUT}/加盟商准入续约_v1_题目.docx')
scrub(f'{OUT}/加盟商准入续约_v1_题目.docx', WORD_APP, '运营分析部', '2026-01-04T02:00:00Z', '2026-01-10T01:00:00Z', (2026, 1, 10, 10, 0, 0))

print('done')
print('on', on_list, 'reform', ref_list, 'rej', rej_list)
print('grades', grade_counts)
print('checks', check_n, 'supply', r2(supply_total), 'mgmt', r2(mgmt_total))
print('packs', len(INPUT_PACKS), len(ANSWER_PACKS))
