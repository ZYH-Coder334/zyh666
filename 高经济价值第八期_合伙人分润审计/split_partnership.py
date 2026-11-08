#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""合伙人分润审计 v1 附件拆分与打包：8输入 + 8答案，并生成题目 docx"""
import os
import io
import shutil
import zipfile

import openpyxl
from openpyxl import Workbook
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

BASE = '/workspace/高经济价值第八期_合伙人分润审计'
SRC_Q = f'{BASE}/合伙人分润审计_v1_题目.xlsx'
SRC_A = f'{BASE}/合伙人分润审计_v1_标准答案.xlsx'
SPLIT = '/tmp/partnership_split'
OUT = f'{BASE}/合伙人分润审计_v1_附件'
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
             '<Company>睿衡企业管理咨询有限公司</Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc>'
             '<HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0300</AppVersion></Properties>')
WORD_APP = EXCEL_APP.replace('Microsoft Excel', 'Microsoft Office Word').replace('16.0300', '16.0000')

CORE_TPL = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            '<dc:creator>{creator}</dc:creator><cp:lastModifiedBy>{creator}</cp:lastModifiedBy>'
            '<dcterms:created xsi:type="dcterms:W3CDTF">{created}</dcterms:created>'
            '<dcterms:modified xsi:type="dcterms:W3CDTF">{modified}</dcterms:modified></cp:coreProperties>')

IN_DT = (2026, 10, 15, 9, 0, 0)
AN_DT = (2026, 10, 25, 14, 0, 0)
DOC_DT = (2026, 10, 20, 10, 0, 0)


def copy_sheet(src_ws, dst_wb):
    dst_ws = dst_wb.create_sheet(src_ws.title)
    for row in src_ws.iter_rows():
        for cell in row:
            nc = dst_ws.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                nc.font = cell.font.copy()
                nc.fill = cell.fill.copy()
                nc.border = cell.border.copy()
                nc.alignment = cell.alignment.copy()
                nc.number_format = cell.number_format
    for k, v in src_ws.column_dimensions.items():
        dst_ws.column_dimensions[k].width = v.width
    dst_ws.freeze_panes = src_ws.freeze_panes
    dst_ws.sheet_view.showGridLines = False
    for mc in src_ws.merged_cells.ranges:
        dst_ws.merge_cells(str(mc))


def scrub(path, app_xml, created, modified, dt, creator):
    tmp = path + '.tmp'
    with zipfile.ZipFile(path, 'r') as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == 'docProps/app.xml':
                data = app_xml.encode('utf-8')
            elif item.filename == 'docProps/core.xml':
                data = CORE_TPL.format(creator=creator, created=created, modified=modified).encode('utf-8')
            zi = zipfile.ZipInfo(item.filename, date_time=dt)
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = 0o600 << 16
            zout.writestr(zi, data)
    os.replace(tmp, path)
    ts = int(__import__('datetime').datetime(*dt).timestamp())
    os.utime(path, (ts, ts))


def pack_rar(xlsx_path, rar_path, arcname, dt):
    with zipfile.ZipFile(rar_path, 'w', zipfile.ZIP_DEFLATED) as z:
        zi = zipfile.ZipInfo(arcname, date_time=dt)
        zi.compress_type = zipfile.ZIP_DEFLATED
        zi.external_attr = 0o600 << 16
        z.writestr(zi, open(xlsx_path, 'rb').read())
    ts = int(__import__('datetime').datetime(*dt).timestamp())
    os.utime(rar_path, (ts, ts))


INPUT_PACKS = [
    ('01项目分润台账', ['项目分润台账']),
    ('02合伙人考核结果', ['合伙人考核结果']),
    ('03制度文件汇编', ['制度文件汇编']),
    ('04财务账面分润', ['财务账面分润']),
    ('05竞业与离职证据', ['竞业与离职证据']),
    ('06飞单证据', ['飞单证据']),
    ('07分润发放记录', ['分润发放记录']),
    ('08管理层要求', ['管理层要求']),
]
ANSWER_PACKS = [
    ('01制度适用判定', ['制度适用判定']),
    ('02项目分润计提', ['项目分润计提重算']),
    ('03合伙人分润明细', ['合伙人分润明细', '考核系数与飞单处理']),
    ('04竞业违约处理', ['竞业违约处理']),
    ('05飞单处理', ['飞单处理']),
    ('06风险与整改机制', ['风险清单', '整改机制']),
    ('07错误要求驳回', ['错误要求驳回']),
    ('08勾稽与结论', ['勾稽自检', '合伙人分润审计结论']),
]


def r4(x): return round(x + 0.0, 4)


def main():
    in_created = '2026-10-15T01:00:00Z'
    in_modified = '2026-10-15T01:00:00Z'
    an_created = '2026-10-25T06:00:00Z'
    an_modified = '2026-10-25T06:00:00Z'

    # ---- 输入包 ----
    wb_q = openpyxl.load_workbook(SRC_Q)
    for name, sheets in INPUT_PACKS:
        nb = Workbook()
        nb.remove(nb.active)
        for s in sheets:
            copy_sheet(wb_q[s], nb)
        xp = f'{SPLIT}/input/{name}.xlsx'
        nb.save(xp)
        scrub(xp, EXCEL_APP, in_created, in_modified, IN_DT, '审计部-王梓')
        pack_rar(xp, f'{OUT}/输入附件/{name}.rar', f'{name}.xlsx', IN_DT)

    # ---- 答案包 ----
    wb_a = openpyxl.load_workbook(SRC_A)
    for name, sheets in ANSWER_PACKS:
        nb = Workbook()
        nb.remove(nb.active)
        for s in sheets:
            copy_sheet(wb_a[s], nb)
        xp = f'{SPLIT}/ans/{name}.xlsx'
        nb.save(xp)
        scrub(xp, EXCEL_APP, an_created, an_modified, AN_DT, '审计部-王梓')
        pack_rar(xp, f'{OUT}/标准答案/{name}.rar', f'{name}.xlsx', AN_DT)

    # ---- 从答案读结论数据（动态） ----
    wb_v = openpyxl.load_workbook(SRC_A, data_only=True)
    detail = {r[0]: r for r in wb_v['合伙人分润明细'].iter_rows(min_row=2, values_only=True)}
    checks = [r for r in wb_v['勾稽自检'].iter_rows(min_row=2, values_only=True)]
    n_checks = len(checks)
    n_bad = sum(1 for c in checks if c[4] != '一致')
    book_total = r4(sum(r[5] for r in detail.values()))
    correct_total = r4(sum(r[4] for r in detail.values()))
    diff_total = r4(correct_total - book_total)
    pe_forfeit = r4(detail['PE'][2])
    pa_pay = detail['PA'][4]
    pc_pay = detail['PC'][4]
    pd_pay = detail['PD'][4]

    # ---- 题目 docx ----
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = '宋体'
    style.font.size = Pt(10.5)
    style.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    def H(txt, level=1):
        h = doc.add_heading(txt, level=level)
        for run in h.runs:
            run.font.name = '宋体'
            run.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
        return h

    def P(txt, indent=True):
        p = doc.add_paragraph(txt)
        if indent:
            p.paragraph_format.first_line_indent = Pt(21)
        return p

    doc.add_heading('合伙人绩效与项目分润审计（G8-ZX-003 v1）', 0)
    for run in doc.paragraphs[0].runs:
        run.font.name = '宋体'
        run.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    H('一、编号')
    P('G8-ZX-003（v1）。')
    P('【职业定位】二级行业：专业咨询服务/合伙人治理与激励；L1职业簇：人力资源、财务与审计、合规与风控；L2职业：薪酬绩效经理、合伙事务经理、内控审计师、劳动关系合规专员、财务经理。')

    H('二、任务背景')
    P('睿衡企业管理咨询有限公司实行合伙人制，项目分润按制度版本、项目归属与年度考核系数计提发放。截至2026年10月15日，公司2026年度涉及分润的项目8个、合伙人6名。')
    P('年度结算前，审计部发现：分润制度存在v1.0（谁签约归谁、签约额10%）与v2.0（谁交付归谁、毛利15%）两个版本，跨年度项目须按衔接条款判定；管委会对P03项目有六四分享的专项决议，财务账面未执行；合伙人PE离职后涉嫌违反竞业条款，名下7万元分润未发放；合伙人PD涉嫌飞单；财务账面分润全部按v2.0与系数1.0计提，与制度要求不符。公司要求出具一份覆盖制度适用判定、计提重算、竞业与飞单处理、分润明细核对的完整审计底稿。')
    P('本次任务存在三个现实张力：一是签约合伙人与交付合伙人对归属各执一词，管委会决议的效力范围须准确界定；二是PE已离职，没收分润涉及劳资敏感处理；三是管理层出于维稳与利润考虑，提出多项与制度冲突的口头要求。')
    P('数据截止日为2026年10月15日，分析基准日为2026年10月25日。所有结论须基于项目分润台账、合伙人考核结果、制度文件、财务账面分润、竞业与离职证据、飞单证据、分润发放记录计算，不得只给汇总数。')

    H('三、任务分解')
    rounds = [
        ('第1轮为制度适用判定。对8个项目逐一判定适用制度版本与归属合伙人，建立证据、权限、生效时间、适用对象、当前状态的判定链。', '01制度适用判定.rar'),
        ('第2轮为项目分润计提重算。按正确版本与基数逐项目重算计提金额，识别P03例外拆分。', '02项目分润计提.rar'),
        ('第3轮为合伙人分润明细。结合考核系数与飞单降档，编制逐合伙人应发分润明细并与账面核对。', '03合伙人分润明细.rar'),
        ('第4轮为竞业违约处理。对PE竞业违约定性并处理其未发分润。', '04竞业违约处理.rar'),
        ('第5轮为飞单处理。对PD飞单定性并处理飞单收入与考核降档。', '05飞单处理.rar'),
        ('第6轮为风险与整改机制。列出10项风险、6类整改机制。', '06风险与整改机制.rar'),
        ('第7轮为错误要求驳回。逐条驳回管理层12项错误要求。', '07错误要求驳回.rar'),
        ('第8轮为勾稽与结论。完成勾稽自检并给出核心结论。', '08勾稽与结论.rar'),
    ]
    P('本项委托按八轮执行，每轮只完成一个任务，并只提交当轮对应交付包。上一轮自检通过后，方可进入下一轮；未完成或自检不通过，不得进入下一轮，也不得提前输出后续轮次结果。')
    for txt, pkg in rounds:
        P(f'{txt}本轮只输出{pkg}。')

    H('四、输入附件')
    P('输入附件共8包。')
    for name, sheets in INPUT_PACKS:
        P(f'{name}.rar，包含{"、".join(sheets)}。')

    H('五、制度与判断规则')
    P('制度适用：2025年12月31日前签约的项目适用v1.0（分润归属按谁签约归谁，基数为签约额、比例10%）；2026年1月1日后签约的项目适用v2.0（归属按谁交付归谁，基数为项目毛利、比例15%）；跨年度项目按签约时点制度执行，不因结算年度变更。')
    P('管委会决议：P03项目按签约合伙人PA与交付合伙人PC六四分享，仅适用P03与2026年度结算，不形成惯例。')
    P('考核系数：A=1.2、B=1.0、C=0.8，须管委会备案；应发分润=计提金额×最终考核系数。')
    P('亏损项目：项目毛利为负的当年不计提分润，亦不从合伙人其他分润中倒扣。')
    P('竞业条款：合伙人离职后12个月内不得服务其原负责客户；违反者未发放分润不予发放、归入奖金池结转，已发放分润不追回；没收分润不得计入公司利润。')
    P('飞单处理：合伙人私接公司业务范围内项目属飞单；飞单收入全额归公司并纳入账务与税务申报；责任合伙人当年考核系数降一档。')
    P('适用判定链：证据、权限、生效时间、适用对象、当前状态五要素齐备方可适用，不得仅凭最新日期或领导口头意见判断。')

    H('六、交付物')
    for name, _ in ANSWER_PACKS:
        P(f'{name}.rar'.replace('0', '0 '), indent=False) if False else P(f'{name.split(".")[0]}.rar', indent=False)

    H('七、约束条件')
    conds = [
        '仅使用附件材料，不得引入外部数据，不得联网检索。',
        '金额以万元列示并保留四位小数，系数保留一位小数。',
        '逐项目、逐合伙人给出计算结果，不得抽样或估算。',
        '8个项目、6名合伙人编号唯一，不得重复或遗漏。',
        '制度版本必须按签约日期与衔接条款判定，不得统一按最新版本。',
        '管委会决议仅适用P03与2026年度，不得扩大或沿用。',
        '亏损项目不分润、不倒扣。',
        '考核系数必须按备案结果执行，不得一律按1.0。',
        'PE未发分润7万元必须没收结转，不得私分或转付他人。',
        '没收分润归入奖金池结转，不得计入公司利润。',
        '飞单收入30万元必须全额归入公司并纳入税务申报。',
        'PD考核系数必须降档至0.8。',
        'PE已发放分润不追回，无制度依据不得扩大追责。',
        '分润年末计提、次年一季度发放，未发放部分挂账应付分润。',
        '必须列出10项风险和6类整改机制，每项风险须有量化触发阈值。',
        '勾稽全部闭合后方可出具核心结论。',
        '必须按八轮顺序逐轮交付。',
        '每轮结束必须提交本轮自检记录。',
        '每轮只提交当轮对应交付包，不得跨轮合并或提前输出。',
        '管理层12项错误要求须逐条驳回并说明理由。',
        '计提重算与分润明细必须同源同口径，两表不得互相矛盾。',
        '制度适用判定须覆盖全部8个项目，每个项目注明判定依据。',
        '竞业违约定性必须引用证据编号，不得只给结论。',
        '风险、整改机制、驳回要求之间不得重复计数同一事项。',
    ]
    for i, c in enumerate(conds, 1):
        P(f'{i}. {c}', indent=False)

    H('八、完成标准')
    P(f'八轮任务逐轮完成且逐轮自检通过。8个项目制度适用全部判定完成；正确应发分润合计{correct_total}万元、账面{book_total}万元、差异{diff_total}万元全部核定；PE没收结转{pe_forfeit}万元、飞单收归30万元处理完毕；{n_checks}项勾稽全部一致、不一致项为{n_bad}。任一轮未完成或自检不通过，不得进入下一轮，整体不得视为完成。')

    H('九、考点')
    P('以下考点按制度适用、计提重算、系数与飞单、分润明细、竞业处理、风险整改、错误驳回、勾稽结论八组展开，共50项，用于逐条评定答题底稿的完整性与正确性。')
    kao = [
        '（一）制度适用判定（考点1—7）',
        '考点1：底稿中应当对8个项目逐一给出适用版本与归属合伙人的判定。',
        '考点2：底稿中应当指出P01、P02、P03、P08共4个项目适用v1.0。',
        '考点3：底稿中应当指出P04、P05、P06、P07共4个项目适用v2.0。',
        '考点4：底稿中应当以签约日期2025-12-31前后作为版本分界线，并引用衔接条款。',
        '考点5：底稿中应当指出管委会决议仅适用P03与2026年度，不形成惯例。',
        '考点6：底稿中应当明确最新日期与领导口头意见都不是适用判据，须按证据、权限、生效时间、适用对象、当前状态判定。',
        '考点7：底稿中应当识别财务账面全部误按v2.0且系数按1.0的系统性错误。',
        '（二）计提重算（考点8—15）',
        '考点8：底稿中应当逐项目列示计提基数、比例与计提金额。',
        '考点9：底稿中应当计算P01计提12万元（容差±0.01万元）。',
        '考点10：底稿中应当计算P03计提15万元并按6:4拆分为PA9万元、PC6万元。',
        '考点11：底稿中应当计算P05计提6.0825万元（容差±0.01万元），基数为毛利40.55万元而非签约额。',
        '考点12：底稿中应当计算P04计提5.574万元、P07计提1.845万元（容差±0.01万元）。',
        '考点13：底稿中应当计算P08计提7万元并归属签约合伙人PE。',
        '考点14：底稿中应当指出P06毛利为负，计提0元且不倒扣。',
        '考点15：底稿中应当计算正确计提合计55.5015万元（容差±0.01万元）。',
        '（三）考核系数与飞单（考点16—21）',
        '考点16：底稿中应当按备案考核等级确定系数：PA1.2、PB1.0、PC1.2、PD1.0（调整前）、PE1.0、PF1.0。',
        '考点17：底稿中应当认定PD飞单成立并引用FD-01至FD-04证据。',
        '考点18：底稿中应当将PD考核系数由1.0降为0.8。',
        '考点19：底稿中应当将飞单收入30万元全额归入公司并纳入税务申报。',
        '考点20：底稿中应当指出飞单证据链包含报销单、出差审批、客户证言、合同摘要四类。',
        '考点21：底稿中应当指出飞单收入不得体外循环、不得账外处理。',
        '（四）分润明细（考点22—27）',
        '考点22：底稿中应当计算PA应发25.2万元（容差±0.01万元）。',
        '考点23：底稿中应当计算PB应发8万元、PC应发14.499万元（容差±0.01万元）。',
        '考点24：底稿中应当计算PD应发5.9352万元（容差±0.01万元）。',
        '考点25：底稿中应当计算PE、PF应发均为0元，并说明原因。',
        '考点26：底稿中应当计算正确应发合计53.6342万元、账面42.426万元、差异11.2082万元（容差±0.01万元）。',
        '考点27：底稿中应当逐合伙人列示计提基数、系数、正确应发、账面应发、差异五列。',
        '（五）竞业处理（考点28—32）',
        '考点28：底稿中应当认定PE违反竞业条款并引用ZJ-01至ZJ-03证据。',
        '考点29：底稿中应当将PE未发分润7万元没收并归入奖金池结转。',
        '考点30：底稿中应当指出没收分润不得计入公司利润。',
        '考点31：底稿中应当指出已发放分润不追回，且本年尚未发放任何分润。',
        '考点32：底稿中应当指出没收分润不得由在职合伙人平分、不得转付PF。',
        '（六）风险与整改（考点33—37）',
        '考点33：底稿中应当列示10项风险，每项含量化触发阈值与应对措施。',
        '考点34：底稿中应当列示6类整改机制，每类含触发条件与处理方式。',
        '考点35：底稿中风险与整改机制应当一一对应、不重复计数。',
        '考点36：底稿中应当包含离职合伙人分润冻结与12个月竞业核查机制。',
        '考点37：底稿中应当包含报销、出差、立项三方交叉的飞单核查机制。',
        '（七）错误要求驳回（考点38—43）',
        '考点38：底稿中应当逐条识别管理层12项错误要求。',
        '考点39：底稿中应当识别不执行决议、私分没收款、不降档、统一新版本、亏损倒扣、系数一刀切、错付P08、基数用签约额、追回已发、没收款入利润、飞单不入账、决议沿用等错误。',
        '考点40：底稿中应当对每项错误要求给出不予采纳结论并说明正确做法。',
        '考点41：底稿中驳回理由应当引用具体制度编号或数据，不得只做原则性表态。',
        '考点42：底稿中应当指出M04与M08互为口径错误、M07与M02同属错付没收款。',
        '考点43：底稿中应当指出M11构成账外收入的合规后果。',
        '（八）勾稽与结论（考点44—50）',
        '考点44：底稿中应当完成不少于20项勾稽，覆盖项目数、合伙人数、签约额、毛利、版本分布、计提合计、账面合计、应发合计、差异、没收结转、系数、单人应发等维度。',
        '考点45：底稿中勾稽结论应当全部为一致，不一致项为0。',
        '考点46：底稿中应当给出核心结论，包含五类问题定性、应发差异、没收结转、飞单处理四项量化结果。',
        '考点47：底稿中核心结论须由四条支撑组成。',
        '考点48：底稿中应当列出关键结果，包括逐合伙人应发、没收结转7万元、飞单收归30万元、勾稽状态。',
        '考点49：底稿中结论数据应当与明细表同源，不得出现两口径并存。',
        '考点50：底稿中应当说明P03决议下年度不再沿用，回归衔接条款判定。',
    ]
    for k in kao:
        P(k, indent=False)

    H('十、答题思路')
    steps = [
        '第一步是读取项目分润台账、合伙人考核结果、制度文件、财务账面分润、竞业与离职证据、飞单证据与分润发放记录。',
        '第二步是建立制度适用判定链，按签约日期与衔接条款逐项目判定版本与归属，识别P03管委会例外。',
        '第三步是按正确版本与基数逐项目重算计提金额，亏损项目不计提不倒扣。',
        '第四步是结合考核系数与飞单降档编制逐合伙人应发分润明细，与账面核对形成差异表。',
        '第五步是对PE竞业违约定性，没收未发分润7万元并结转奖金池。',
        '第六步是对PD飞单定性，飞单收入30万元收归公司并降档处理。',
        '第七步是列出10项风险与6类整改机制，逐条驳回12项管理层错误要求。',
        '第八步是完成勾稽自检，确认计提表与明细表同源闭合后给出核心结论。',
    ]
    for s in steps:
        P(s)

    H('十一、标准答案')
    P('标准答案共8个包，首表为合伙人分润审计结论，逐包覆盖制度适用判定、项目分润计提、合伙人分润明细、竞业违约处理、飞单处理、风险整改、错误驳回和勾稽自检。')
    P(f'核心结论：合伙人分润存在制度适用错误、管委会决议未执行、竞业违约未追责、飞单未处理、考核系数未执行五类问题；账面应发分润合计{book_total}万元，按制度重算后正确应发合计{correct_total}万元，差异{diff_total}万元；PE竞业违约没收未发分润{pe_forfeit}万元结转奖金池；PD飞单收入30万元收归公司并降档处理。')
    P('支撑一：制度适用：8个项目中4个适用v1.0（签约归属、签约额10%）、4个适用v2.0（交付归属、毛利15%），财务账面全部误按v2.0且系数按1.0，计提口径整体错误。')
    P('支撑二：P03例外：管委会专项决议要求PA与PC按6:4分享（9万元/6万元），账面误全归PC（8.4285万元）；决议仅限P03与2026年度，不形成惯例。')
    P(f'支撑三：竞业与飞单：PE离职后12个月内承接原客户星辉传媒业务，违反竞业条款，未发分润{pe_forfeit}万元没收结转且不得计入利润；PD飞单P09证据链完整，收入归公司、系数由1.0降为0.8。')
    P('支撑四：管理要求：12项错误要求全部驳回，其中M01/M12违反决议效力范围，M04/M08违反衔接条款，M05/M09/M10无制度依据，M11构成账外收入。')
    P(f'关键结果：正确应发：PA {pa_pay}万元、PB 8万元、PC {pc_pay}万元、PD {pd_pay}万元、PE 0万元、PF 0万元；没收结转{pe_forfeit}万元；飞单收归30万元；勾稽{n_checks}项全部一致，不一致项{n_bad}项；管理层12项错误要求全部驳回。')

    docx_path = f'{OUT}/合伙人分润审计_v1_题目.docx'
    doc.save(docx_path)
    scrub(docx_path, WORD_APP, '2026-10-20T02:00:00Z', '2026-10-20T02:00:00Z', DOC_DT, '审计部-王梓')
    print('docx done:', docx_path)
    print('packs done:', len(INPUT_PACKS), len(ANSWER_PACKS))


if __name__ == '__main__':
    main()
