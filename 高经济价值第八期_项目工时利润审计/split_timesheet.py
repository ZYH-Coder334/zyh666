#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""咨询项目工时与项目利润核算审计 v1 附件拆分与打包：8输入 + 8答案，并生成题目 docx"""
import os
import shutil
import zipfile

import openpyxl
from openpyxl import Workbook
from copy import copy
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

BASE = '/workspace/高经济价值第八期_项目工时利润审计'
SRC_Q = f'{BASE}/咨询项目工时利润审计_v1_题目.xlsx'
SRC_A = f'{BASE}/咨询项目工时利润审计_v1_标准答案.xlsx'
SPLIT = '/tmp/timesheet_split'
OUT = f'{BASE}/咨询项目工时利润审计_v1_附件'
shutil.rmtree(SPLIT, ignore_errors=True)
shutil.rmtree(OUT, ignore_errors=True)
os.makedirs(f'{SPLIT}/input', exist_ok=True)
os.makedirs(f'{SPLIT}/ans', exist_ok=True)
os.makedirs(f'{OUT}/输入附件', exist_ok=True)
os.makedirs(f'{OUT}/标准答案', exist_ok=True)

COMPANY = '睿衡企业管理咨询有限公司'
IN_DT = (2026, 8, 31, 9, 0, 0)
AN_DT = (2026, 9, 10, 14, 0, 0)
DOC_DT = (2026, 9, 5, 10, 0, 0)

EXCEL_APP = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
             '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
             'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
             '<Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>'
             f'<Company>{COMPANY}</Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc>'
             '<HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0300</AppVersion></Properties>')
WORD_APP = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
            'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
            '<Application>Microsoft Office Word</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>'
            f'<Company>{COMPANY}</Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc>'
            '<HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0000</AppVersion></Properties>')


def core_xml(author, created, modified):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            f'<dc:creator>{author}</dc:creator><cp:lastModifiedBy>{author}</cp:lastModifiedBy>'
            f'<dcterms:created xsi:type="dcterms:W3CDTF">{created}</dcterms:created>'
            f'<dcterms:modified xsi:type="dcterms:W3CDTF">{modified}</dcterms:modified></cp:coreProperties>')


def scrub_inner(data, app_xml, core):
    """在内存中重写 xlsx/docx 的 docProps"""
    import io
    src = zipfile.ZipFile(io.BytesIO(data), 'r')
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as dst:
        for info in src.infolist():
            if info.filename in ('docProps/app.xml', 'docProps/core.xml'):
                continue
            zi = zipfile.ZipInfo(info.filename, date_time=info.date_time)
            zi.compress_type = zipfile.ZIP_DEFLATED
            dst.writestr(zi, src.read(info.filename))
        dst.writestr('docProps/app.xml', app_xml)
        dst.writestr('docProps/core.xml', core)
    src.close()
    return buf.getvalue()


def pack_rar(files, rar_path, app_xml, author, created, modified, zip_dt):
    """将若干 xlsx/docx 伪装后打入 .rar（zip deflate），时间戳统一，无顶层残留"""
    import datetime
    core = core_xml(author, created, modified)
    with zipfile.ZipFile(rar_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for arcname, fpath in files:
            with open(fpath, 'rb') as f:
                data = scrub_inner(f.read(), app_xml, core)
            zi = zipfile.ZipInfo(arcname, date_time=zip_dt)
            zi.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(zi, data)
    ts = datetime.datetime(*zip_dt).timestamp()
    os.utime(rar_path, (ts, ts))


def copy_sheet(ws_src, ws_dst):
    for row in ws_src.iter_rows():
        for cell in row:
            nc = ws_dst[cell.coordinate]
            nc.value = cell.value
            if cell.has_style:
                nc.font = copy(cell.font)
                nc.fill = copy(cell.fill)
                nc.border = copy(cell.border)
                nc.alignment = copy(cell.alignment)
    for col, dim in ws_src.column_dimensions.items():
        ws_dst.column_dimensions[col].width = dim.width
    ws_dst.freeze_panes = ws_src.freeze_panes
    ws_dst.sheet_view.showGridLines = ws_src.sheet_view.showGridLines
    for mr in ws_src.merged_cells.ranges:
        ws_dst.merge_cells(str(mr))


def split_pack(src, sheet_names, out_xlsx):
    wb_src = openpyxl.load_workbook(src)
    wb = Workbook()
    wb.remove(wb.active)
    for name in sheet_names:
        ws = wb.create_sheet(name)
        copy_sheet(wb_src[name], ws)
    wb.save(out_xlsx)
    return out_xlsx


INPUT_PACKS = [
    ('01项目主数据', ['项目主数据']),
    ('02工时系统导出', ['工时系统导出']),
    ('03项目管理系统数据', ['项目管理系统数据']),
    ('04费率制度文件', ['费率制度文件']),
    ('05顾问成本费率表', ['顾问成本费率表']),
    ('06财务账面收入成本', ['财务账面收入成本']),
    ('07停滞项目沟通记录', ['停滞项目沟通记录']),
    ('08管理层要求', ['管理层要求']),
]

ANSWER_PACKS = [
    ('01费率适用判定', ['费率适用判定']),
    ('02工时勾稽调节', ['工时勾稽调节']),
    ('03项目收入重算', ['项目收入重算']),
    ('04项目成本与利润', ['项目成本重算', '项目利润表']),
    ('05停滞项目减值', ['停滞项目减值处理']),
    ('06风险与整改机制', ['风险清单与整改机制']),
    ('07错误要求驳回', ['错误要求驳回']),
    ('08勾稽与结论', ['勾稽自检', '项目利润核算审计结论']),
]


def r4(x): return round(x + 0.0, 4)


def main():
    in_created = '2026-08-31T01:00:00Z'
    in_modified = '2026-08-31T01:00:00Z'
    an_created = '2026-09-10T06:00:00Z'
    an_modified = '2026-09-10T06:00:00Z'

    for name, sheets in INPUT_PACKS:
        p = split_pack(SRC_Q, sheets, f'{SPLIT}/input/{name}.xlsx')
        pack_rar([(f'{name}.xlsx', p)], f'{OUT}/输入附件/{name}.rar',
                 EXCEL_APP, '项目经济管理部', in_created, in_modified, IN_DT)

    for name, sheets in ANSWER_PACKS:
        p = split_pack(SRC_A, sheets, f'{SPLIT}/ans/{name}.xlsx')
        pack_rar([(f'{name}.xlsx', p)], f'{OUT}/标准答案/{name}.rar',
                 EXCEL_APP, '审计与合规项目组', an_created, an_modified, AN_DT)

    build_docx(f'{OUT}/咨询项目工时利润审计_v1_题目.docx')

    print('packs done:', len(os.listdir(f"{OUT}/输入附件")), len(os.listdir(f"{OUT}/标准答案")))


def build_docx(path):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = '宋体'
    style.font.size = Pt(10.5)
    style.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    def H(text, size=14):
        p = doc.add_paragraph()
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(size)
        run.font.name = '宋体'
        run.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    def P(text):
        p = doc.add_paragraph()
        p.paragraph_format.first_line_indent = Pt(21)
        run = p.add_run(text)
        run.font.size = Pt(10.5)
        run.font.name = '宋体'
        run.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    # 从答案xlsx动态读取关键结果
    wba = openpyxl.load_workbook(SRC_A, data_only=True)
    checks = list(wba['勾稽自检'].iter_rows(min_row=2, values_only=True))
    n_checks = len(checks)
    n_bad = sum(1 for c in checks if c[4] != '一致')
    fin_rev_total = r4(sum(r[2] for r in wba['项目收入重算'].iter_rows(min_row=2, values_only=True)))
    cor_rev_total = r4(sum(r[3] for r in wba['项目收入重算'].iter_rows(min_row=2, values_only=True)))
    fin_cost_total = r4(sum(r[2] for r in wba['项目成本重算'].iter_rows(min_row=2, values_only=True)))
    cor_cost_total = r4(sum(r[3] for r in wba['项目成本重算'].iter_rows(min_row=2, values_only=True)))
    cor_profit = r4(cor_rev_total - cor_cost_total)

    H('咨询项目工时与项目利润核算审计', 16)

    H('一、编号')
    P('G8-ZX-002（v1）。')
    P('【职业定位】二级行业：专业咨询服务/项目经济管理；L1职业簇：财务与审计、运营管理、项目管理；'
      'L2职业：项目财务经理、交付经理、工时审核专员、内控审计师、管理会计。')

    H('二、任务背景')
    P('睿衡企业管理咨询有限公司按项目制交付咨询服务，收入确认分为固定价里程碑、固定价完工百分比、计时T&M三种模式，'
      '项目成本按实际工时乘以成本费率核算。截至2026年8月31日，公司在执行项目8个，顾问10人。')
    P('半年度项目经济复盘发现：工时系统填报工时与项目管理系统审批工时不一致；计费费率存在v2025与v2026两个版本，'
      '一份销售总监折扣邮件已经备案、另一份客户经理折扣申请被合规部驳回；P06项目表面微利但工时系统标注了大量返工；'
      'P07项目停滞超60天，客户已书面确认只按已完成工作的80%结算，财务尚未计提减值。'
      '公司要求出具一份覆盖费率适用判定、工时勾稽调节、收入与成本重算、项目利润表、停滞项目减值的完整审计底稿。')
    P('本次任务存在三个现实张力：一是交付部门希望多计费、多确认收入，财务要求按审批口径从严；'
      '二是费率版本与折扣文件并存，适用正确与否直接改变收入；三是P06亏损与P07减值都会影响半年度利润，管理层存在粉饰倾向。')
    P('数据截止日为2026年8月31日，分析基准日为2026年9月10日。'
      '所有结论须基于项目主数据、工时系统导出、项目管理系统数据、费率制度文件、顾问成本费率表、财务账面收入成本、停滞项目沟通记录计算，不得只给汇总数。')

    H('三、任务分解')
    P('本项委托按八轮执行，每轮只完成一个任务，并只提交当轮对应交付包。'
      '上一轮自检通过后，方可进入下一轮；未完成或自检不通过，不得进入下一轮，也不得提前输出后续轮次结果。')
    P('第1轮为费率适用判定。对8个项目逐一判定适用费率版本与折扣效力，建立证据、权限、生效时间、适用对象、当前状态的判定链。本轮只输出01费率适用判定.rar。')
    P('第2轮为工时勾稽调节。将工时系统填报工时调节至项目管理系统审批工时，识别未审批加班与返工工时。本轮只输出02工时勾稽调节.rar。')
    P('第3轮为项目收入重算。按正确确认方法与适用费率逐项目重算收入，与账面核对。本轮只输出03项目收入重算.rar。')
    P('第4轮为项目成本与利润。按实际工时重算成本，编制逐项目利润表并标注亏损项目。本轮只输出04项目成本与利润.rar。')
    P('第5轮为停滞项目减值。评估P07可回收性并按客户书面确认比例计提减值。本轮只输出05停滞项目减值.rar。')
    P('第6轮为风险与整改机制。列出10项风险、6类整改机制。本轮只输出06风险与整改机制.rar。')
    P('第7轮为错误要求驳回。逐条驳回管理层12项错误要求。本轮只输出07错误要求驳回.rar。')
    P('第8轮为勾稽与结论。完成勾稽自检并给出核心结论。本轮只输出08勾稽与结论.rar。')

    H('四、输入附件')
    P('输入附件共8包。')
    for name, sheets in INPUT_PACKS:
        P(f'{name}.rar，包含{"、".join(sheets)}。')

    H('五、制度与判断规则')
    P('计费费率：v2025费率（高级2000元/h、顾问1400元/h、初级900元/h、专家2600元/h）自2025-01-01生效；'
      'v2026费率（上浮10%）自2026-01-01生效。费率版本以合同约定为准，P01合同签订于2025-11-20并锁定v2025；'
      '折扣须经合规部备案方为有效，P03的9折已备案、P05的85折已被驳回。')
    P('收入确认：固定价有验收条款按里程碑验收确认；固定价无验收条款按完工百分比（审批工时÷计划工时）确认；'
      '计时项目按有效工时×适用费率确认。')
    P('工时口径：未审批加班工时不得对外计费；返工工时不得向客户计费，但必须全额计入项目实际成本。')
    P('成本核算：项目成本=实际发生工时×成本费率，含返工与未审批加班工时；成本费率不随计费费率版本变化。')
    P('停滞项目：停滞超60天须评估可回收性，客户书面确认的结算比例作为收入上限，差额计提减值。')
    P('适用判定链：证据、权限、生效时间、适用对象、当前状态五要素齐备方可适用，不得仅凭"最新日期"或"领导邮件"判断。')

    H('六、交付物')
    P('01 费率适用判定.rar')
    P('02 工时勾稽调节.rar')
    P('03 项目收入重算.rar')
    P('04 项目成本与利润.rar')
    P('05 停滞项目减值.rar')
    P('06 风险与整改机制.rar')
    P('07 错误要求驳回.rar')
    P('08 勾稽与结论.rar')

    H('七、约束条件')
    for i, t in enumerate([
        '仅使用附件材料，不得引入外部数据，不得联网检索。',
        '金额以万元列示并保留四位小数，工时以小时列示为整数，比例保留两位小数。',
        '逐项目、逐条工时给出计算结果，不得抽样或估算。',
        '8个项目、10名顾问、18条工时记录编号唯一，不得重复或遗漏。',
        '费率版本必须以合同约定与审批备案为准，不得仅凭最新日期。',
        'P01必须锁定v2025费率并按里程碑确认收入80万元。',
        'P03的9折因合规备案方有效，P05的85折因被驳回而不得适用。',
        'P02未审批加班40h不得计费，但其成本必须计入项目。',
        'P06返工工时150h不得计费，但9万元成本必须补计。',
        'P07必须按客户书面确认的80%结算并计提减值2.52万元。',
        '项目利润必须逐项目披露，不得合并抵销亏损项目。',
        '完工百分比只能按审批工时÷计划工时计算，不得人为调高。',
        '内部培训工时不得分摊到项目成本，应入期间费用。',
        '成本只能按实际工时×成本费率计算，不得用估算系数。',
        '必须列出10项风险和6类整改机制，每项风险须有量化触发阈值。',
        '勾稽全部闭合后方可出具核心结论。',
        '必须按八轮顺序逐轮交付。',
        '每轮结束必须提交本轮自检记录。',
        '每轮只提交当轮对应交付包，不得跨轮合并或提前输出。',
        '管理层12项错误要求须逐条驳回并说明理由。',
        '收入重算、成本重算、利润表必须同源同口径，表间不得互相矛盾。',
        '工时勾稽必须区分未审批加班与返工两类差异并分别说明处理。',
        '减值计提必须以客户书面确认为依据，不得按主观预计。',
        '风险、整改机制、驳回要求之间不得重复计数同一事项。',
    ], 1):
        P(f'{i}. {t}')

    H('八、完成标准')
    P(f'八轮任务逐轮完成且逐轮自检通过。8个项目收入、成本、利润全部重算完成；'
      f'工时勾稽差异190h全部识别（未审批加班40h、返工150h）；P07减值2.52万元计提完毕；'
      f'{n_checks}项勾稽全部一致、不一致项为{n_bad}。'
      f'任一轮未完成或自检不通过，不得进入下一轮，整体不得视为完成。')

    H('九、考点')
    P('以下考点按费率判定、工时勾稽、收入重算、成本利润、停滞减值、风险整改、错误驳回、勾稽结论八组展开，共50项，用于逐条评定答题底稿的完整性与正确性。')
    groups = [
        ('（一）费率适用判定（考点1—7）', [
            '底稿中应当对8个项目逐一给出适用费率版本或不适用的判定。',
            '底稿中应当指出P01锁定v2025费率，依据是合同签订日期2025-11-20与合同锁定条款。',
            '底稿中应当指出v2026费率自2026-01-01起适用于新签项目。',
            '底稿中应当指出P03的9折有效，依据是销售总监邮件加合规部备案，证据链完整。',
            '底稿中应当指出P05的85折无效，依据是合规部2026-03-10驳回。',
            '底稿中应当明确"最新日期"与"领导邮件"都不是费率适用判据，须按证据、权限、生效时间、适用对象、当前状态判定。',
            '底稿中应当指出固定价项目收入与计费费率无关，费率判定只影响计时项目。',
        ]),
        ('（二）工时勾稽调节（考点8—14）', [
            '底稿中应当将系统工时合计3960h调节至审批工时合计3770h。',
            '底稿中应当识别P02差异40h为未审批周末加班，不得对外计费。',
            '底稿中应当识别P06差异150h为返工工时，不得对外计费。',
            '底稿中应当指出未审批加班与返工工时的成本仍须计入项目成本。',
            '底稿中应当逐项目列示系统工时、审批工时、差异三列。',
            '底稿中应当说明工时差异合计190h全部有来源，无无法解释的差异。',
            '底稿中应当以项目管理系统审批口径作为计费依据。',
        ]),
        ('（三）收入重算（考点15—24）', [
            '底稿中应当按里程碑确认P01收入80万元（容差±0.01万元），并指出账面误按完工百分比多确认13.3333万元。',
            '底稿中应当计算P02正确收入92.62万元（容差±0.01万元）。',
            '底稿中应当计算P03正确收入80.19万元（容差±0.01万元），即450h×v2026费率×0.9。',
            '底稿中应当计算P04收入70.2万元（容差±0.01万元），并确认账面正确。',
            '底稿中应当计算P05正确收入53.35万元（容差±0.01万元），即按v2026全价。',
            '底稿中应当确认P06收入25万元，账面正确。',
            '底稿中应当计算P08收入27.72万元（容差±0.01万元），并确认账面正确。',
            '底稿中应当计算正确收入合计439.16万元（容差±0.01万元）。',
            '底稿中应当计算收入净调减20.7208万元（容差±0.01万元）。',
            '底稿中应当逐项目列示账面收入、正确收入、调整额三列。',
        ]),
        ('（四）成本与利润（考点25—32）', [
            '底稿中应当按实际工时×成本费率重算8个项目成本。',
            '底稿中应当补计P06返工成本9万元（容差±0.01万元），即150h×600元/h。',
            '底稿中应当计算正确成本合计170.34万元（容差±0.01万元）。',
            '底稿中应当指出P06账面成本22万元系按计费口径漏记返工。',
            '底稿中应当编制逐项目利润表，含收入、成本、利润、利润率四列。',
            '底稿中应当揭示P06显著性反转：账面微利3万元、实际亏损6万元。',
            '底稿中应当计算正确利润合计268.82万元（容差±0.01万元）。',
            '底稿中应当计算利润虚增29.7208万元（容差±0.01万元），并分解为收入多确认20.7208万元与成本漏记9万元。',
        ]),
        ('（五）停滞项目减值（考点33—36）', [
            '底稿中应当指出P07停滞超60天触发减值评估。',
            '底稿中应当以客户财务总监书面确认的80%作为结算上限。',
            '底稿中应当计算P07正确收入10.08万元并计提减值2.52万元（容差±0.01万元）。',
            '底稿中应当指出账面按完工百分比确认12.6万元未计提减值属于差错。',
        ]),
        ('（六）风险与整改（考点37—40）', [
            '底稿中应当列示10项风险，每项含量化触发阈值与应对措施。',
            '底稿中应当列示6类整改机制，每类含触发条件与处理方式。',
            '底稿中风险与整改机制应当一一对应、不重复计数。',
            '底稿中应当包含计费工时与审批工时差异超2%触发复核的机制。',
        ]),
        ('（七）错误要求驳回（考点41—45）', [
            '底稿中应当逐条识别管理层12项错误要求。',
            '底稿中应当识别未审批计费、误用完工百分比、返工不计成本、适用被驳回折扣、合并隐匿亏损、停滞全额确认、费率一刀切、折扣不备案、培训工时分摊、成本估算、调高完工百分比、不计提减值等错误。',
            '底稿中应当对每项错误要求给出不予采纳结论并说明正确做法。',
            '底稿中驳回理由应当引用具体制度条款或数据，不得只做原则性表态。',
            '底稿中应当指出M07与M02、M08与M04互为对照，考察折扣与费率的证据链完整性。',
        ]),
        ('（八）勾稽与结论（考点46—50）', [
            '底稿中应当完成不少于20项勾稽，覆盖项目数、顾问数、工时合计、收入合计、成本合计、利润、单项目收入、减值、返工成本等维度。',
            '底稿中勾稽结论应当全部为一致，不一致项为0。',
            '底稿中应当给出核心结论，包含收入调减、成本补计、利润虚增、亏损项目、减值计提五项量化结果。',
            '底稿中核心结论须由四条支撑组成。',
            '底稿中应当列出关键结果，包括正确收入439.16万元、正确成本170.34万元、正确利润268.82万元、勾稽状态。',
        ]),
    ]
    n = 1
    for gt, items in groups:
        P(gt)
        for it in items:
            P(f'考点{n}：{it}')
            n += 1

    H('十、答题思路')
    P('第一步是读取项目主数据、工时系统导出、项目管理系统数据、费率制度文件、顾问成本费率表、财务账面收入成本与停滞项目沟通记录。')
    P('第二步是建立费率适用判定链，逐项目判定费率版本与折扣效力，特别注意P01锁定、P03备案、P05驳回三类情形。')
    P('第三步是将工时系统填报工时调节至审批工时，识别未审批加班与返工工时两类差异。')
    P('第四步是按合同类型选择正确确认方法逐项目重算收入，与账面核对形成调整表。')
    P('第五步是按实际工时重算成本并编制逐项目利润表，揭示P06亏损反转。')
    P('第六步是评估P07停滞项目，按客户书面确认比例计提减值。')
    P('第七步是列出10项风险与6类整改机制，逐条驳回12项管理层错误要求。')
    P('第八步是完成勾稽自检，确认收入、成本、利润三表同源闭合后给出核心结论。')

    H('十一、标准答案')
    P('标准答案共8个包，首表为项目利润核算审计结论，逐包覆盖费率适用判定、工时勾稽调节、收入重算、'
      '成本与利润、停滞项目减值、风险整改、错误驳回和勾稽自检。')
    P(f'核心结论：8个咨询项目账面收入{fin_rev_total}万元、账面成本{fin_cost_total}万元、账面利润298.5408万元；'
      f'经重算，正确收入{cor_rev_total}万元、正确成本{cor_cost_total}万元、正确利润{cor_profit}万元，'
      f'利润虚增29.7208万元，来源为收入多确认20.7208万元与返工成本漏记9万元。')
    P('支撑一：收入差错四项：P01有验收条款误按完工百分比多确认13.3333万元；P02含未审批加班多确认3.96万元；'
      'P03漏用已备案9折多确认8.91万元；P05误用被驳回折扣少确认8.0025万元；P07未按客户书面80%确认少提减值2.52万元。')
    P('支撑二：费率适用以合同约定与审批备案为准：P01锁定v2025、P03的9折因合规备案有效、P05的85折被驳回无效；'
      '最新费率表不是适用判据。')
    P('支撑三：P06显著性反转：账面微利3万元，补记返工工时150h×600元/h=9万元成本后实际亏损6万元，'
      '须单独披露并复盘返工原因、更新报价模型。')
    P('支撑四：工时勾稽：系统填报3960h、审批3770h，差异190h全部为已识别事项（未审批加班40h、返工150h），'
      '无无法解释的差异；停滞项目P07已按制度计提减值2.52万元。')
    P(f'关键结果：正确收入{cor_rev_total}万元（调减20.7208万元）；正确成本{cor_cost_total}万元（补计9万元）；'
      f'正确利润{cor_profit}万元；亏损项目1个（P06，-6万元）；减值计提2.52万元；'
      f'勾稽{n_checks}项全部一致，不一致项{n_bad}项；管理层12项错误要求全部驳回。')

    doc.save(path)
    # docx scrub
    core = core_xml('审计与合规项目组', '2026-09-05T02:00:00Z', '2026-09-05T02:00:00Z')
    with open(path, 'rb') as f:
        data = f.read()
    data = scrub_inner(data, WORD_APP, core)
    with open(path, 'wb') as f:
        f.write(data)
    import datetime
    ts = datetime.datetime(*DOC_DT).timestamp()
    os.utime(path, (ts, ts))
    print('docx done:', path)


main()
