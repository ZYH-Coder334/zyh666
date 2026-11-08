#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""应收账款审计 v1 附件拆分与打包：8输入 + 8答案，并生成题目 docx"""
import os
import shutil
import zipfile

import openpyxl
from openpyxl import Workbook
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

BASE = '/workspace/高经济价值第八期_应收账款审计'
SRC_Q = f'{BASE}/应收账款审计_v1_题目.xlsx'
SRC_A = f'{BASE}/应收账款审计_v1_标准答案.xlsx'
SPLIT = '/tmp/ar_split'
OUT = f'{BASE}/应收账款审计_v1_附件'
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

IN_DT = (2026, 11, 30, 9, 0, 0)
AN_DT = (2026, 12, 10, 14, 0, 0)
DOC_DT = (2026, 12, 5, 10, 0, 0)


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
    ('01应收账款明细', ['应收账款明细']),
    ('02客户信用档案', ['客户信用档案']),
    ('03制度文件汇编', ['制度文件汇编']),
    ('04财务账面坏账准备', ['财务账面坏账准备']),
    ('05债务重组协议', ['债务重组协议']),
    ('06破产公告摘要', ['破产公告摘要']),
    ('07催收记录', ['催收记录']),
    ('08管理层要求', ['管理层要求']),
]
ANSWER_PACKS = [
    ('01账龄口径重算', ['账龄口径重算']),
    ('02坏账准备重算', ['坏账准备重算']),
    ('03个别认定处理', ['个别认定处理']),
    ('04账面差异调节', ['账面差异调节']),
    ('05催收与信用政策', ['催收与信用政策']),
    ('06风险与整改机制', ['风险清单', '整改机制']),
    ('07错误要求驳回', ['错误要求驳回']),
    ('08勾稽与结论', ['勾稽自检', '应收账款审计结论']),
]


def r4(x): return round(x + 0.0, 4)


def main():
    in_created = '2026-11-30T01:00:00Z'
    an_created = '2026-12-10T06:00:00Z'

    wb_q = openpyxl.load_workbook(SRC_Q)
    for name, sheets in INPUT_PACKS:
        nb = Workbook()
        nb.remove(nb.active)
        for s in sheets:
            copy_sheet(wb_q[s], nb)
        xp = f'{SPLIT}/input/{name}.xlsx'
        nb.save(xp)
        scrub(xp, EXCEL_APP, in_created, in_created, IN_DT, '审计部-王梓')
        pack_rar(xp, f'{OUT}/输入附件/{name}.rar', f'{name}.xlsx', IN_DT)

    wb_a = openpyxl.load_workbook(SRC_A)
    for name, sheets in ANSWER_PACKS:
        nb = Workbook()
        nb.remove(nb.active)
        for s in sheets:
            copy_sheet(wb_a[s], nb)
        xp = f'{SPLIT}/ans/{name}.xlsx'
        nb.save(xp)
        scrub(xp, EXCEL_APP, an_created, an_created, AN_DT, '审计部-王梓')
        pack_rar(xp, f'{OUT}/标准答案/{name}.rar', f'{name}.xlsx', AN_DT)

    # ---- 从答案读结论数据（动态） ----
    wb_v = openpyxl.load_workbook(SRC_A, data_only=True)
    checks = [r for r in wb_v['勾稽自检'].iter_rows(min_row=2, values_only=True)]
    n_checks = len(checks)
    n_bad = sum(1 for c in checks if c[4] != '一致')
    prov = {r[0]: r for r in wb_v['坏账准备重算'].iter_rows(min_row=2, values_only=True)}
    correct_prov = r4(sum(r[4] for r in prov.values()))
    book_prov = r4(sum(r[5] for r in prov.values()))
    prov_diff = r4(book_prov - correct_prov)
    correct_bal = r4(sum(r[2] for r in prov.values()))

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

    doc.add_heading('应收账款与客户信用管理审计（G8-ZX-004 v1）', 0)
    for run in doc.paragraphs[0].runs:
        run.font.name = '宋体'
        run.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    H('一、编号')
    P('G8-ZX-004（v1）。')
    P('【职业定位】二级行业：专业咨询服务/信用与应收管理；L1职业簇：财务与审计、合规与风控、运营管理；L2职业：应收会计、信用分析师、风控专员、法务专员、财务经理。')

    H('二、任务背景')
    P('睿衡企业管理咨询有限公司按项目验收结算咨询费，合同约定付款期为验收后60天。截至2026年11月30日，公司应收账款14笔、客户14家，其中3笔逾期超180天，1家客户进入破产程序，1笔已签债务重组协议，1笔为质保金。')
    P('年末审计发现：应收系统仍按已废止的开票日起算账龄，而现行制度v2.0以验收日后第61天为起算日；星辉传媒已进入破产程序，财务仍按账龄比例计提而非全额个别认定；青云文旅债务重组协议已生效并正常履约两期，财务未确认重组损失、仍按原余额计提；金陵商贸质保金被并入账龄按开票日起算导致全额误提。公司要求出具一份覆盖账龄口径重算、坏账准备重算、个别认定处理、账面差异调节、催收与信用政策的完整审计底稿。')
    P('本次任务存在三个现实张力：一是系统参数与现行制度不一致，口径更正涉及全部应收；二是破产、重组、质保金三笔特殊业务各有专门规则，均不能套用账龄比例法；三是管理层担心利润与考核，提出多项与制度冲突的口头要求。')
    P('数据截止日为2026年11月30日，分析基准日为2026年12月10日。所有结论须基于应收账款明细、客户信用档案、制度文件、财务账面坏账准备、债务重组协议、破产公告摘要、催收记录计算，不得只给汇总数。')

    H('三、任务分解')
    P('本项委托按八轮执行，每轮只完成一个任务，并只提交当轮对应交付包。上一轮自检通过后，方可进入下一轮；未完成或自检不通过，不得进入下一轮，也不得提前输出后续轮次结果。')
    rounds = [
        ('第1轮为账龄口径重算。以验收日后第61天为起算日逐笔重算逾期天数，与系统开票日口径对照。', '01账龄口径重算.rar'),
        ('第2轮为坏账准备重算。按制度账龄比例逐笔重算计提金额。', '02坏账准备重算.rar'),
        ('第3轮为个别认定处理。处理星辉破产、青云重组、金陵质保金三笔特殊业务。', '03个别认定处理.rar'),
        ('第4轮为账面差异调节。将账面计提逐笔调节至正确口径并说明原因。', '04账面差异调节.rar'),
        ('第5轮为催收与信用政策。对逾期客户给出处置方案并调整信用政策。', '05催收与信用政策.rar'),
        ('第6轮为风险与整改机制。列出10项风险、6类整改机制。', '06风险与整改机制.rar'),
        ('第7轮为错误要求驳回。逐条驳回管理层12项错误要求。', '07错误要求驳回.rar'),
        ('第8轮为勾稽与结论。完成勾稽自检并给出核心结论。', '08勾稽与结论.rar'),
    ]
    for txt, pkg in rounds:
        P(f'{txt}本轮只输出{pkg}。')

    H('四、输入附件')
    P('输入附件共8包。')
    for name, sheets in INPUT_PACKS:
        P(f'{name}.rar，包含{"、".join(sheets)}。')

    H('五、制度与判断规则')
    P('账龄起算：现行制度v2.0以验收日后第61天为账龄起算日（合同约定付款期为验收后60天）；旧制度按开票日起算，自2026-01-01废止。')
    P('比例计提：未逾期0.5%；逾期1-90天5%；逾期91-180天20%；逾期181-365天50%；逾期超365天100%。')
    P('个别认定：客户进入破产程序、被吊销营业执照或明确表示无力偿还的，不受账龄比例限制，全额计提。')
    P('债务重组：减免部分确认债务重组损失；重组后债权按新协议余额核算，正常履约期间按10%个别计提；履约逾期一期即恢复全额计提。')
    P('质保金：质保期内不并入应收账龄，自质保期届满之日起转为应收款并起算账龄。')
    P('催收分级：逾期90天邮件催收、180天律师函、365天诉讼；破产公告后须在申报期限内完成债权申报。')
    P('适用判定链：证据、权限、生效时间、适用对象、当前状态五要素齐备方可适用，不得仅凭系统现状或领导口头意见判断。')

    H('六、交付物')
    for name, _ in ANSWER_PACKS:
        P(f'{name}.rar', indent=False)

    H('七、约束条件')
    conds = [
        '仅使用附件材料，不得引入外部数据，不得联网检索。',
        '金额以万元列示并保留四位小数，比例保留一位小数，天数为整数。',
        '逐笔应收给出计算结果，不得抽样或估算。',
        '14笔应收、14家客户编号唯一，不得重复或遗漏。',
        '账龄必须以验收日后第61天起算，不得沿用开票日口径。',
        '破产客户必须全额个别计提，不得套用账龄比例。',
        '债务重组减免12.6万元必须确认损失，余额按10%个别计提。',
        '质保金必须自届满日起算，质保期内不并入账龄。',
        '未逾期应收必须按0.5%计提，不得免提。',
        '逾期应收必须按阶梯比例计提，不得统一按100%。',
        '破产债权必须在申报期限内申报，不得直接核销。',
        '催收外包也必须保留全部书面记录。',
        '坏账准备必须按季滚动调整，不得年底一次性调。',
        '信用政策必须按客户评级收紧，C级预付款50%、D级款到服务。',
        '必须列出10项风险和6类整改机制，每项风险须有量化触发阈值。',
        '勾稽全部闭合后方可出具核心结论。',
        '必须按八轮顺序逐轮交付。',
        '每轮结束必须提交本轮自检记录。',
        '每轮只提交当轮对应交付包，不得跨轮合并或提前输出。',
        '管理层12项错误要求须逐条驳回并说明理由。',
        '账龄重算、计提重算、差异调节必须同源同口径，表间不得互相矛盾。',
        '账面差异调节必须逐笔说明原因，不得只给汇总数。',
        '个别认定处理必须引用制度编号与证据编号。',
        '风险、整改机制、驳回要求之间不得重复计数同一事项。',
    ]
    for i, c in enumerate(conds, 1):
        P(f'{i}. {c}', indent=False)

    H('八、完成标准')
    P(f'八轮任务逐轮完成且逐轮自检通过。14笔应收账款龄与计提全部重算完成；账面计提{book_prov}万元调节至正确计提{correct_prov}万元，差异{prov_diff}万元；重组损失12.6万元确认完毕；破产债权申报、诉讼、分期、停止新单处置全部落地；{n_checks}项勾稽全部一致、不一致项为{n_bad}。任一轮未完成或自检不通过，不得进入下一轮，整体不得视为完成。')

    H('九、考点')
    P('以下考点按账龄口径、计提重算、个别认定、差异调节、催收信用、风险整改、错误驳回、勾稽结论八组展开，共50项，用于逐条评定答题底稿的完整性与正确性。')
    kao = [
        '（一）账龄口径（考点1—7）',
        '考点1：底稿中应当以验收日后第61天为账龄起算日逐笔重算。',
        '考点2：底稿中应当指出系统按开票日起算属已废止口径。',
        '考点3：底稿中应当逐笔列示开票日口径与验收日口径两套逾期天数。',
        '考点4：底稿中应当识别两口径计提比例不同的应收笔数。',
        '考点5：底稿中应当指出质保金不适用于开票日或验收日起算，自届满日起算。',
        '考点6：底稿中应当指出账龄制度v2.0自2026-01-01现行有效、发布主体为财务部。',
        '考点7：底稿中应当明确系统现状不是口径判据，须按制度版本与生效时间判定。',
        '（二）计提重算（考点8—14）',
        '考点8：底稿中应当逐笔按制度账龄比例重算计提金额。',
        '考点9：底稿中应当计算安捷物流按50%计提32.5万元（容差±0.01万元）。',
        '考点10：底稿中应当计算博瑞汽车按50%计提27.5万元（容差±0.01万元）。',
        '考点11：底稿中应当计算天启科技按20%计提10万元（余额50万元）。',
        '考点12：底稿中应当计算华宇制造未逾期按0.5%计提0.3万元（容差±0.01万元）。',
        '考点13：底稿中应当计算正确计提合计225.775万元（容差±0.01万元）。',
        '考点14：底稿中应当指出未逾期应收也须按0.5%计提。',
        '（三）个别认定（考点15—22）',
        '考点15：底稿中应当对星辉传媒全额计提48万元并引用破产公告BP-01。',
        '考点16：底稿中应当指出星辉须在2026-12-17前完成破产债权申报。',
        '考点17：底稿中应当确认青云重组损失12.6万元（容差±0.01万元）。',
        '考点18：底稿中应当将青云应收余额调整为24.5万元并按10%计提2.45万元（容差±0.01万元）。',
        '考点19：底稿中应当指出青云已正常履约2期共4.9万元。',
        '考点20：底稿中应当指出重组履约逾期一期即恢复全额计提。',
        '考点21：底稿中应当将金陵质保金按届满日起算，逾期10天按5%计提0.4万元（容差±0.01万元）。',
        '考点22：底稿中应当指出个别认定共3笔，分别为破产、重组、质保金口径更正。',
        '（四）差异调节（考点23—28）',
        '考点23：底稿中应当从账面计提254.625万元逐笔调节至225.775万元（容差±0.01万元）。',
        '考点24：底稿中应当识别账面多提28.85万元（容差±0.01万元）。',
        '考点25：底稿中应当指出金陵质保金账面误全额计提8万元为最大单笔差异之一。',
        '考点26：底稿中应当指出青云账面误按50%计提21万元且未确认重组损失。',
        '考点27：底稿中应当逐笔说明调节原因，不得只给汇总数。',
        '考点28：底稿中应当将正确应收余额核定为648.5万元（容差±0.01万元）。',
        '（五）催收与信用（考点29—33）',
        '考点29：底稿中应当对安捷物流、宏图建筑给出诉讼处置。',
        '考点30：底稿中应当对博瑞汽车给出分期协议处置、对旭日科技给出停止新单处置。',
        '考点31：底稿中应当提出C级客户预付款50%、D级款到服务的信用政策。',
        '考点32：底稿中应当指出破产客户处置为债权申报而非直接核销。',
        '考点33：底稿中处置方案应当与账龄分级和信用评级联动。',
        '（六）风险与整改（考点34—37）',
        '考点34：底稿中应当列示10项风险，每项含量化触发阈值与应对措施。',
        '考点35：底稿中应当列示6类整改机制，每类含触发条件与处理方式。',
        '考点36：底稿中风险与整改机制应当一一对应、不重复计数。',
        '考点37：底稿中应当包含账龄口径锁定与重组台账跟踪两项机制。',
        '（七）错误要求驳回（考点38—43）',
        '考点38：底稿中应当逐条识别管理层12项错误要求。',
        '考点39：底稿中应当识别开票日起算、破产按比例提、不确认重组损失、质保金按开票日、统一全额提、未逾期免提、重组履约免提、长账龄不提、年底一次性调、催收不留记录、信用政策不改、直接核销等错误。',
        '考点40：底稿中应当对每项错误要求给出不予采纳结论并说明正确做法。',
        '考点41：底稿中驳回理由应当引用具体制度编号或数据，不得只做原则性表态。',
        '考点42：底稿中应当指出M05与M06方向相反但同属违反计提办法。',
        '考点43：底稿中应当指出M12直接核销造成破产分配权益丧失的合规后果。',
        '（八）勾稽与结论（考点44—50）',
        '考点44：底稿中应当完成不少于20项勾稽，覆盖笔数、客户数、账面余额、正确余额、重组损失、账面计提、正确计提、差异、逾期分级、个别认定、单笔计提等维度。',
        '考点45：底稿中勾稽结论应当全部为一致，不一致项为0。',
        '考点46：底稿中应当给出核心结论，包含四类问题定性、计提差异、重组损失、处置方案四项量化结果。',
        '考点47：底稿中核心结论须由四条支撑组成。',
        '考点48：底稿中应当列出关键结果，包括正确余额648.5万元、正确计提225.775万元、重组损失12.6万元、勾稽状态。',
        '考点49：底稿中结论数据应当与明细表同源，不得出现两口径并存。',
        '考点50：底稿中应当说明账龄口径更正后系统参数须锁定，防止再次回退。',
    ]
    for k in kao:
        P(k, indent=False)

    H('十、答题思路')
    steps = [
        '第一步是读取应收账款明细、客户信用档案、制度文件、财务账面坏账准备、债务重组协议、破产公告摘要与催收记录。',
        '第二步是按验收日后第61天逐笔重算逾期天数，与系统开票日口径对照形成两套账龄。',
        '第三步是按制度账龄比例逐笔重算坏账准备，识别破产、重组、质保金三笔须个别认定的业务。',
        '第四步是对星辉全额计提、青云确认重组损失并按10%计提、金陵按届满日起算更正。',
        '第五步是从账面计提逐笔调节至正确口径，说明每笔差异原因。',
        '第六步是按账龄分级与信用评级给出诉讼、申报、分期、停止新单等处置并调整信用政策。',
        '第七步是列出10项风险与6类整改机制，逐条驳回12项管理层错误要求。',
        '第八步是完成勾稽自检，确认账龄、计提、调节三表同源闭合后给出核心结论。',
    ]
    for s in steps:
        P(s)

    H('十一、标准答案')
    P('标准答案共8个包，首表为应收账款审计结论，逐包覆盖账龄口径重算、坏账准备重算、个别认定处理、账面差异调节、催收与信用政策、风险整改、错误驳回和勾稽自检。')
    P(f'核心结论：应收账款存在账龄口径错误、个别认定缺失、债务重组未入账、质保金混入账龄四类问题；账面应收余额666万元经重组调整后正确余额{correct_bal}万元；账面坏账准备{book_prov}万元，正确计提{correct_prov}万元，账面多提{prov_diff}万元；另须补确认债务重组损失12.6万元。')
    P('支撑一：账龄口径：账龄管理制度v2.0以验收日后第61天为起算日，系统仍按已废止的开票日口径运行，14笔应收中多笔两口径计提比例不同。')
    P('支撑二：个别认定：星辉传媒已进入破产程序须全额计提48万元并在申报期限内申报债权；青云文旅重组减免12.6万元确认损失、余额24.5万元按10%计提；金陵商贸质保金届满日起算仅计提0.4万元，账面误全额计提8万元。')
    P('支撑三：催收与信用：安捷物流、宏图建筑进入诉讼程序，博瑞汽车签分期协议，旭日科技停止新单；C级客户预付款提至50%、D级款到服务。')
    P('支撑四：管理要求：12项错误要求全部驳回，其中M02/M12违反个别认定与破产程序，M03/M07违反重组处理规则，M01/M04违反账龄与质保金制度，M05/M06违反计提办法。')
    P(f'关键结果：正确应收余额{correct_bal}万元；正确坏账准备{correct_prov}万元（账面多提{prov_diff}万元）；重组损失12.6万元；破产申报1笔（48万元）；诉讼2笔（135万元）；勾稽{n_checks}项全部一致，不一致项{n_bad}项；管理层12项错误要求全部驳回。')

    docx_path = f'{OUT}/应收账款审计_v1_题目.docx'
    doc.save(docx_path)
    scrub(docx_path, WORD_APP, '2026-12-05T02:00:00Z', '2026-12-05T02:00:00Z', DOC_DT, '审计部-王梓')
    print('docx done:', docx_path)
    print('packs done:', len(INPUT_PACKS), len(ANSWER_PACKS))


if __name__ == '__main__':
    main()
