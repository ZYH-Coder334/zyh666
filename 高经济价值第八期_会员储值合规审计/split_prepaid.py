#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""会员储值与预收账款合规管理审计 v1 附件拆分与打包：8输入 + 8答案，并生成题目 docx"""
import os
import shutil
import zipfile

import openpyxl
from openpyxl import Workbook
from copy import copy
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

BASE = '/workspace/高经济价值第八期_会员储值合规审计'
SRC_Q = f'{BASE}/会员储值合规审计_v1_题目.xlsx'
SRC_A = f'{BASE}/会员储值合规审计_v1_标准答案.xlsx'
SPLIT = '/tmp/prepaid_split'
OUT = f'{BASE}/会员储值合规审计_v1_附件'
shutil.rmtree(SPLIT, ignore_errors=True)
shutil.rmtree(OUT, ignore_errors=True)
os.makedirs(f'{SPLIT}/input', exist_ok=True)
os.makedirs(f'{SPLIT}/ans', exist_ok=True)
os.makedirs(f'{OUT}/输入附件', exist_ok=True)
os.makedirs(f'{OUT}/标准答案', exist_ok=True)

COMPANY = '睿衡企业管理咨询有限公司'
IN_DT = (2026, 6, 30, 9, 0, 0)
AN_DT = (2026, 7, 10, 14, 0, 0)
DOC_DT = (2026, 7, 5, 10, 0, 0)

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
    ('01客户储值台账', ['客户储值台账']),
    ('02储值消费明细', ['储值消费明细']),
    ('03储值系统余额', ['储值系统余额表']),
    ('04财务预收账款', ['财务预收账款明细']),
    ('05银行专户流水', ['银行监管专户流水']),
    ('06制度文件汇编', ['制度文件汇编']),
    ('07退卡申请记录', ['退卡申请记录']),
    ('08管理层要求', ['管理层要求']),
]

ANSWER_PACKS = [
    ('01三方勾稽调节', ['三方勾稽调节表']),
    ('02余额与收入重算', ['储值余额重算表', '收入确认重算表']),
    ('03退卡与沉睡卡处理', ['退卡处理方案', '沉睡卡处理']),
    ('04监管资金缺口定性', ['监管资金缺口定性']),
    ('05制度适用判定', ['制度适用判定']),
    ('06风险与整改机制', ['风险清单与整改机制']),
    ('07错误要求驳回', ['错误要求驳回']),
    ('08勾稽与结论', ['勾稽自检', '预收账款合规审计结论']),
]


def main():
    in_created = '2026-06-30T01:00:00Z'
    in_modified = '2026-06-30T01:00:00Z'
    an_created = '2026-07-10T06:00:00Z'
    an_modified = '2026-07-10T06:00:00Z'

    for name, sheets in INPUT_PACKS:
        p = split_pack(SRC_Q, sheets, f'{SPLIT}/input/{name}.xlsx')
        pack_rar([(f'{name}.xlsx', p)], f'{OUT}/输入附件/{name}.rar',
                 EXCEL_APP, '财务管理部', in_created, in_modified, IN_DT)

    for name, sheets in ANSWER_PACKS:
        p = split_pack(SRC_A, sheets, f'{SPLIT}/ans/{name}.xlsx')
        pack_rar([(f'{name}.xlsx', p)], f'{OUT}/标准答案/{name}.rar',
                 EXCEL_APP, '审计与合规项目组', an_created, an_modified, AN_DT)

    build_docx(f'{OUT}/会员储值合规审计_v1_题目.docx')

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
    gap = {r[0]: r[1] for r in wba['监管资金缺口定性'].iter_rows(min_row=2, values_only=True)}
    should = gap['应存金额']
    bank_end = gap['专户期末余额（2026-06-30）']
    recon = list(wba['三方勾稽调节表'].iter_rows(min_row=2, values_only=True))
    fin_total = recon[0][2]
    correct_total = recon[-1][2]

    H('会员储值与预收账款合规管理审计', 16)

    H('一、编号')
    P('G8-ZX-001（v1）。')
    P('【职业定位】二级行业：专业咨询服务/商务服务与咨询；L1职业簇：财务与审计、合规与风控、运营管理；'
      'L2职业：审计经理、收入会计、合规专员、内控经理、客户成功经理。')

    H('二、任务背景')
    P('睿衡企业管理咨询有限公司是一家面向企业客户提供战略、流程、数字化与合规咨询的专业服务机构。'
      '公司自2025年起发行会员制顾问储值卡，客户预先充值本金并获赠一定比例赠送金额，咨询费按工时消耗从储值中扣减。'
      '截至2026年6月30日，在售储值卡客户12家，储值本金合计860万元、赠送金额合计98万元。')
    P('审计部在半年度检查中发现：储值系统余额、财务预收账款（合同负债）、银行监管专户三方数据互不相等；'
      '消费抵扣规则存在v2.1与v3.0两个版本并行，且有一封业务部门的豁免邮件与一封合规部门的撤销邮件；'
      '两笔退卡处理口径存疑；一张卡已沉睡超期；监管专户存在83万元缺口。'
      '公司要求出具一份覆盖三方勾稽调节、规则适用判定、余额与收入重算、退卡与沉睡卡处理、资金缺口定性的完整审计底稿。')
    P('本次任务存在三个现实张力：一是业务部门希望维护大客户、倾向从宽适用规则，财务与合规部要求从严；'
      '二是系统按旧规则实时记账，财务按新规则确认收入，两边都"看起来正确"；'
      '三是管理层担心专户缺口影响融资，存在隐瞒倾向。')
    P('数据截止日为2026年6月30日，分析基准日为2026年7月10日。'
      '所有结论须基于客户储值台账、储值消费明细、储值系统余额、财务预收账款、银行专户流水、制度文件、退卡申请记录计算，不得只给汇总数。')

    H('三、任务分解')
    P('本项委托按八轮执行，每轮只完成一个任务，并只提交当轮对应交付包。'
      '上一轮自检通过后，方可进入下一轮；未完成或自检不通过，不得进入下一轮，也不得提前输出后续轮次结果。')
    P('第1轮为三方勾稽调节。以财务账面合同负债为起点，逐项调节至正确口径，说明每项差异的来源单据。本轮只输出01三方勾稽调节.rar。')
    P('第2轮为余额与收入重算。按正确规则逐客户重算本金消耗、赠送消耗、余额与收入，与账面核对。本轮只输出02余额与收入重算.rar。')
    P('第3轮为退卡与沉睡卡处理。重算K07、K11退卡金额与手续费，判定K05沉睡卡处理。本轮只输出03退卡与沉睡卡处理.rar。')
    P('第4轮为监管资金缺口定性。计算应存金额，核对专户流水，区分时间性差异与实质性挪用。本轮只输出04监管资金缺口定性.rar。')
    P('第5轮为制度适用判定。对每条消费记录判定适用规则版本，建立证据、权限、生效时间、适用对象、当前状态的判定链。本轮只输出05制度适用判定.rar。')
    P('第6轮为风险与整改机制。列出10项风险、6类整改机制。本轮只输出06风险与整改机制.rar。')
    P('第7轮为错误要求驳回。逐条驳回管理层12项错误要求。本轮只输出07错误要求驳回.rar。')
    P('第8轮为勾稽与结论。完成勾稽自检并给出核心结论。本轮只输出08勾稽与结论.rar。')

    H('四、输入附件')
    P('输入附件共8包。')
    for name, sheets in INPUT_PACKS:
        P(f'{name}.rar，包含{"、".join(sheets)}。')

    H('五、制度与判断规则')
    P('储值消费抵扣：制度v2.1（2025-01-01生效）先扣本金、本金扣完后扣赠送；制度v3.0（2026-03-01生效）按储值本金与赠送金额的原始比例同步消耗。'
      '运营总监豁免邮件（2026-04-10）仅对K09中泰能源在窗口期内适用v2.1；合规部撤销邮件（2026-05-20）自撤销日起统一按v3.0执行，窗口期已发生消费不追溯。')
    P('收入确认：预收储值款计入合同负债，按工时消耗进度确认收入；赠送部分属商业折扣不计收入；开票与确认时点差异属时间性差异，须在次月调整。')
    P('退卡：退剩余本金，收取剩余本金5%手续费，赠送余额作废不退，退款30日内到账。')
    P('沉睡卡：超过卡有效期且连续12个月无消费且无法联系客户的，本金余额转营业外收入，赠送余额作废。')
    P('资金监管：合同负债余额的40%须存入银行监管专户，专户资金不得挪作他用；时间性差异须有期后银行流水佐证，超出承诺期限未归还的转出定性为挪用。')
    P('规则适用判定链：证据、权限、生效时间、适用对象、当前状态五要素齐备方可适用，不得仅凭"最新日期"判断有效性。')

    H('六、交付物')
    P('01 三方勾稽调节.rar')
    P('02 余额与收入重算.rar')
    P('03 退卡与沉睡卡处理.rar')
    P('04 监管资金缺口定性.rar')
    P('05 制度适用判定.rar')
    P('06 风险与整改机制.rar')
    P('07 错误要求驳回.rar')
    P('08 勾稽与结论.rar')

    H('七、约束条件')
    for i, t in enumerate([
        '仅使用附件材料，不得引入外部数据，不得联网检索。',
        '金额以万元列示并保留四位小数，比例保留两位小数。',
        '逐客户、逐笔消费给出计算结果，不得抽样或估算。',
        '12家客户、15笔消费编号唯一，不得重复或遗漏。',
        '赠送部分属商业折扣，不得计入收入，退卡时不得退还。',
        '抵扣规则必须按证据、权限、生效时间、适用对象、当前状态五要素判定，不得仅凭最新日期。',
        'K09豁免仅适用于2026-04-10至2026-05-20窗口内的消费，窗口外必须按v3.0。',
        'K03在2026-06-28的20万元消费须补确认收入，不得跨期隐藏。',
        '退卡手续费必须为剩余本金的5%，退款30日内到账。',
        '沉睡卡三要件（超有效期、连续12个月无消费、失联）齐备方可转营业外收入。',
        '专户应存金额按正确口径合同负债的40%计算。',
        '时间性差异与实质性挪用必须分开定性，并各自给出佐证。',
        '63万元期后转回不得用于冲抵20万元挪用定性。',
        '不得隐瞒专户缺口，审计报告必须如实披露。',
        '必须列出10项风险和6类整改机制，每项风险须有量化触发阈值。',
        '勾稽全部闭合后方可出具核心结论。',
        '必须按八轮顺序逐轮交付。',
        '每轮结束必须提交本轮自检记录。',
        '每轮只提交当轮对应交付包，不得跨轮合并或提前输出。',
        '管理层12项错误要求须逐条驳回并说明理由。',
        '三方勾稽调节必须从财务账面余额出发，逐项调节至正确口径，每步注明依据。',
        '退卡补退金额、手续费补确认金额须单独列示，不得与正常退款混同。',
        '收入重算与余额重算必须同源同口径，两表不得互相矛盾。',
        '风险、整改机制、驳回要求之间不得重复计数同一事项。',
    ], 1):
        P(f'{i}. {t}')

    H('八、完成标准')
    P(f'八轮任务逐轮完成且逐轮自检通过。12家客户全部重算完成；三方勾稽调节四步闭合；'
      f'退卡2笔（K07补退、K11待审批）、沉睡卡1张（K05）处理完毕；专户缺口83万元定性完毕'
      f'（63万元时间性差异已消除、20万元实质性挪用未归还）；{n_checks}项勾稽全部一致、不一致项为{n_bad}。'
      f'任一轮未完成或自检不通过，不得进入下一轮，整体不得视为完成。')

    H('九、考点')
    P('以下考点按三方勾稽、规则适用、余额重算、收入确认、退卡沉睡卡、资金定性、风险整改、错误驳回、勾稽结论九组展开，共50项，用于逐条评定答题底稿的完整性与正确性。')
    groups = [
        ('（一）三方勾稽（考点1—6）', [
            '底稿中应当以财务账面合同负债467.4536万元（容差±0.01万元）为调节起点。',
            '底稿中应当逐项列示调节事项，且每步注明依据单据或制度。',
            '底稿中应当将正确口径合同负债调节至425.8798万元（容差±0.01万元）。',
            '底稿中应当说明储值系统总余额535万元与财务账面的差异来源，包含K07退卡系统未核销、v2.1与v3.0口径差异两项。',
            '底稿中应当列明调节四步的净额为调减41.5738万元（容差±0.01万元）。',
            '底稿中三方差异应当全部可追溯至具体单据，不得出现无来源的差异项。',
        ]),
        ('（二）规则适用判定（考点7—13）', [
            '底稿中应当对15笔消费逐笔判定适用规则版本。',
            '底稿中应当指出v3.0自2026-03-01起现行有效、适用全部客户。',
            '底稿中应当指出2026-03-01之前的消费按v2.1执行。',
            '底稿中应当指出K09豁免仅覆盖2026-04-10至2026-05-20窗口，窗口内消费按v2.1。',
            '底稿中应当指出合规部撤销邮件（2026-05-20）自撤销日起生效且不追溯。',
            '底稿中应当明确"最新日期"不是有效性判据，须按证据、权限、生效时间、适用对象、当前状态判定。',
            '底稿中应当识别K09在2026-05-08的30万元碳管理咨询适用v2.1、其余窗口外消费适用v3.0。',
        ]),
        ('（三）余额重算（考点14—20）', [
            '底稿中应当逐客户重算本金消耗、赠送消耗、本金余额、赠送余额四个字段。',
            '底稿中应当计算K09正确本金余额为76.5217万元（容差±0.01万元）。',
            '底稿中应当计算K09正确赠送余额为15.9783万元（容差±0.01万元）。',
            '底稿中应当给出系统余额与正确余额的逐客户差异列。',
            '底稿中应当识别系统按v2.1记账导致赠送余额普遍未被消耗（合计98万元未动）。',
            '底稿中应当识别K07系统未核销退卡导致系统余额虚高39万元（30万本金+9万赠送）。',
            '底稿中余额重算与收入重算应当同源同口径。',
        ]),
        ('（四）收入确认（考点21—26）', [
            '底稿中应当按本金消耗进度确认收入，赠送部分不计收入。',
            '底稿中应当补确认K03收入17.3913万元（容差±0.01万元）。',
            '底稿中应当补确认K09收入3.9131万元（容差±0.01万元）。',
            '底稿中应当列示逐客户"账面收入、正确收入、调整额"三列。',
            '底稿中应当指出K03差异属开票与确认的时点差异，不得定性为隐瞒收入以外的处理。',
            '底稿中收入补确认合计应当为21.3044万元（容差±0.01万元）。',
        ]),
        ('（五）退卡与沉睡卡（考点27—33）', [
            '底稿中应当计算K07正确剩余本金35.4545万元（容差±0.01万元）。',
            '底稿中应当计算K07应退33.6818万元、已退28.5万元、补退5.1818万元（容差±0.01万元）。',
            '底稿中应当补确认K07手续费收入0.2727万元（容差±0.01万元）。',
            '底稿中应当计算K11应退8.7963万元、手续费0.4630万元（容差±0.01万元），并标注其状态为待审批。',
            '底稿中应当指出K07退款错误根源是按系统v2.1余额办理。',
            '底稿中应当判定K05满足沉睡卡三要件，本金余额14.8148万元转营业外收入（容差±0.01万元）。',
            '底稿中应当将K05、K07赠送余额合计4.7307万元作废（容差±0.01万元）。',
        ]),
        ('（六）资金缺口定性（考点34—39）', [
            '底稿中应当计算专户应存金额为170.3519万元（容差±0.01万元）。',
            '底稿中应当核对专户期末余额87.3519万元并得出缺口83万元。',
            '底稿中应当将63万元定性为时间性差异，并以2026-07-01期后转回流水为佐证。',
            '底稿中应当将20万元定性为实质性挪用，并指出其超出CFO承诺的30天期限。',
            '底稿中不得以期后转回的63万元冲抵20万元挪用的定性。',
            '底稿中应当要求20万元限期归还计息、冻结审批权限并追责。',
        ]),
        ('（七）风险与整改（考点40—43）', [
            '底稿中应当列示10项风险，每项含量化触发阈值与应对措施。',
            '底稿中应当列示6类整改机制，每类含触发条件与处理方式。',
            '底稿中风险与整改机制应当一一对应、不重复计数。',
            '底稿中应当包含每月末三方对账机制（系统、财务、银行两两差异超0.5%触发）。',
        ]),
        ('（八）错误要求驳回（考点44—47）', [
            '底稿中应当逐条识别管理层12项错误要求。',
            '底稿中应当识别赠送计收入、退卡退赠送、沉睡卡不处理、挪用专户、规则挑有利的用、跨期藏收入、手续费加码、按系统余额为准、隐瞒缺口、K09全程老规则、不存专户、办卡一次性确认收入等错误。',
            '底稿中应当对每项错误要求给出不予采纳结论并说明正确做法。',
            '底稿中驳回理由应当引用具体制度条款或数据，不得只做原则性表态。',
        ]),
        ('（九）勾稽与结论（考点48—50）', [
            f'底稿中应当完成不少于20项勾稽，覆盖客户数、消费笔数、本金合计、赠送合计、系统余额、财务账面、调节净额、专户应存、流水勾稽、退卡金额、赠送作废、收入补确认等维度。',
            f'底稿中勾稽结论应当全部为一致，不一致项为0。',
            '底稿中应当给出核心结论，包含五类问题定性、合同负债调减额、缺口定性、整改要求，并由四条支撑组成。',
        ]),
    ]
    n = 1
    for gt, items in groups:
        P(gt)
        for it in items:
            P(f'考点{n}：{it}')
            n += 1

    H('十、答题思路')
    P('第一步是读取客户储值台账、储值消费明细、储值系统余额、财务预收账款、银行专户流水、制度文件与退卡申请记录。')
    P('第二步是建立规则适用判定链，对15笔消费逐笔判定适用v2.1还是v3.0，特别注意K09豁免窗口。')
    P('第三步是按正确口径逐客户重算本金消耗、赠送消耗与余额，与系统余额逐户核对。')
    P('第四步是重算收入确认，识别K03未入账与K09口径差异，形成收入调整表。')
    P('第五步是处理退卡与沉睡卡，重算K07补退、K11待退与K05转营业外收入。')
    P('第六步是核对专户流水，计算应存金额，将83万元缺口区分为时间性差异与实质性挪用。')
    P('第七步是列出10项风险与6类整改机制，逐条驳回12项管理层错误要求。')
    P('第八步是从财务账面出发完成三方勾稽调节四步闭合，自检全部勾稽项后给出核心结论。')

    H('十一、标准答案')
    P('标准答案共8个包，首表为预收账款合规审计结论，逐包覆盖三方勾稽调节、余额与收入重算、退卡与沉睡卡处理、'
      '资金缺口定性、制度适用判定、风险整改、错误驳回和勾稽自检。')
    P(f'核心结论：睿衡咨询预收账款管理存在收入确认滞后、抵扣规则适用错误、退卡口径错误、沉睡卡未清理、监管资金挪用五类问题；'
      f'财务账面合同负债{fin_total}万元，经四步调节后正确余额为{correct_total}万元；'
      f'专户应存{should}万元、期末实存{bank_end}万元，缺口83万元中63万元为已消除的时间性差异、20万元定性为挪用且超期未还。')
    P(f'支撑一：三方勾稽：储值系统总余额535万元（v2.1口径且退卡未核销）、财务账面{fin_total}万元、银行专户{bank_end}万元，'
      f'差异全部可追溯至K03未入账、K09口径差异、K05沉睡卡、K07退卡口径错误与两笔专户转出。')
    P('支撑二：规则适用：v3.0自2026-03-01现行有效，K09豁免仅覆盖2026-04-10至2026-05-20窗口，'
      '窗口外按v3.0；据此K09正确收入73.4783万元，财务少确认3.9131万元。')
    P('支撑三：退卡与沉睡卡：K07应退33.6818万元（已退28.5万元系按系统v2.1余额误算）须补退5.1818万元；'
      'K11应退8.7963万元待审批；K05本金余额14.8148万元转营业外收入、赠送1.1852万元作废。')
    P('支撑四：资金定性：63万元月末调拨已于7月1日转回属时间性差异；20万元转出发放奖金已超CFO承诺的30天期限未归还，'
      '定性为挪用预收资金，须3个工作日内归还计息、冻结审批权限、追责并披露。')
    P(f'关键结果：正确合同负债{correct_total}万元；补确认收入21.3044万元；补确认手续费0.2727万元；'
      f'转营业外收入14.8148万元；补退客户5.1818万元；挪用资金20万元；'
      f'勾稽{n_checks}项全部一致，不一致项{n_bad}项；管理层12项错误要求全部驳回。')

    doc.save(path)
    # docx scrub
    core = core_xml('审计与合规项目组', '2026-07-05T02:00:00Z', '2026-07-05T02:00:00Z')
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
