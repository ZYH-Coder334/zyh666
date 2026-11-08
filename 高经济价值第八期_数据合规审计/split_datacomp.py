#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""数据合规审计 v1 附件拆分与打包：8输入 + 8答案，并生成题目 docx"""
import os
import shutil
import zipfile

import openpyxl
from openpyxl import Workbook
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

BASE = '/workspace/高经济价值第八期_数据合规审计'
SRC_Q = f'{BASE}/数据合规审计_v1_题目.xlsx'
SRC_A = f'{BASE}/数据合规审计_v1_标准答案.xlsx'
SPLIT = '/tmp/datacomp_split'
OUT = f'{BASE}/数据合规审计_v1_附件'
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

IN_DT = (2026, 12, 31, 9, 0, 0)
AN_DT = (2027, 1, 10, 14, 0, 0)
DOC_DT = (2027, 1, 5, 10, 0, 0)


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
    ('01项目数据留存台账', ['项目数据留存台账']),
    ('02制度文件汇编', ['制度文件汇编']),
    ('03项目合同签署记录', ['项目合同签署记录']),
    ('04外包合同与沟通记录', ['外包合同与沟通记录']),
    ('05离职与设备记录', ['离职与设备记录']),
    ('06安全巡检日志', ['安全巡检日志']),
    ('07数据销毁记录', ['数据销毁记录']),
    ('08管理层要求', ['管理层要求']),
]
ANSWER_PACKS = [
    ('01版本适用判定', ['版本适用判定']),
    ('02留存期重算', ['留存期重算']),
    ('03外包转委托审查', ['外包转委托审查']),
    ('04设备回收审查', ['设备回收审查']),
    ('05违规分级处置', ['违规分级处置']),
    ('06网盘事件处置', ['网盘事件处置']),
    ('07错误要求驳回', ['错误要求驳回']),
    ('08勾稽与结论', ['勾稽自检', '数据合规审计结论']),
]


def main():
    in_created = '2026-12-31T01:00:00Z'
    an_created = '2027-01-10T06:00:00Z'

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
    vers = [r for r in wb_v['版本适用判定'].iter_rows(min_row=2, values_only=True)]
    n_v1 = sum(1 for r in vers if r[3] == 'v1')
    n_v2 = sum(1 for r in vers if r[3] == 'v2')
    n_fix = sum(1 for r in vers if r[6] != '—')
    rets = [r for r in wb_v['留存期重算'].iter_rows(min_row=2, values_only=True)]
    od = {r[0]: r[7] for r in rets if r[7] and str(r[7]).startswith('超期')}
    wfs = [r for r in wb_v['违规分级处置'].iter_rows(min_row=2, values_only=True)]
    n_viol = len(wfs)
    n_rectify = sum(1 for r in wfs if '整改' in r[5])
    n_report = sum(1 for r in wfs if '报告客户' in r[5])
    n_terminate = sum(1 for r in wfs if '解除合同' in r[5])
    n_account = sum(1 for r in wfs if '追责' in r[5])

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

    doc.add_heading('客户数据与保密合规专项审计（G8-ZX-005 v1）', 0)
    for run in doc.paragraphs[0].runs:
        run.font.name = '宋体'
        run.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    H('一、编号')
    P('G8-ZX-005（v1）。')
    P('【职业定位】二级行业：专业咨询服务/数据合规与信息安全；L1职业簇：合规与风控、信息技术、运营管理；L2职业：数据合规专员、信息安全工程师、内控审计师、法务专员、IT资产管理员。')

    H('二、任务背景')
    P('睿衡企业管理咨询有限公司为客户提供管理咨询服务，服务过程中接触大量客户数据。截至2026年12月31日，公司年度数据合规检查覆盖20个咨询项目的客户数据使用、外发、留存与销毁情况，涉及保密协议v1（留存90天）与v2（留存180天）两个并行版本。')
    P('检查发现多项线索：留存台账登记的协议版本与项目合同签署记录不一致；P09项目的数据分析环节外包给数析科技，疑似未经客户书面同意转委托；2026年离职5名顾问的设备回收记录不完整；IT安全巡检发现一起将客户明细数据上传个人网盘的历史行为。')
    P('本次任务存在三个现实张力：一是保密协议v1与v2并行，系统登记口径与合同签署口径冲突，版本适用直接决定超期认定结果；二是数据销毁、客户报告与商业利益冲突，管理层提出多项与制度冲突的要求；三是违规证据分散在合同、邮件、平台日志、快递单等多个载体，须交叉印证方可定性。')
    P('数据截止日为2026年12月31日，分析基准日为2027年1月10日。所有结论须基于项目数据留存台账、制度文件、合同签署记录、外包与沟通记录、离职与设备记录、安全巡检日志、数据销毁记录逐项核对得出，不得只给汇总数。')

    H('三、任务分解')
    P('本项委托按八轮执行，每轮只完成一个任务，并只提交当轮对应交付包。上一轮自检通过后，方可进入下一轮；未完成或自检不通过，不得进入下一轮，也不得提前输出后续轮次结果。')
    rounds = [
        ('第1轮为版本适用判定。合并台账与合同签署记录，逐项目判定适用保密协议版本并指出系统登记错误。', '01版本适用判定.rar'),
        ('第2轮为留存期重算。逐项目重算应销毁截止日，与销毁记录比对，认定超期留存项目及超期天数。', '02留存期重算.rar'),
        ('第3轮为外包转委托审查。核对书面同意要件、授权签字人名录与脱敏条款，认定是否构成违规转委托。', '03外包转委托审查.rar'),
        ('第4轮为设备回收审查。核对离职登记、设备回收登记与签收凭证，认定缺失台数。', '04设备回收审查.rar'),
        ('第5轮为违规分级处置。对全部违规逐项按整改、报告客户、解除合同、追责分级定性并给出处置动作与时限。', '05违规分级处置.rar'),
        ('第6轮为网盘事件处置。还原事件时间线、影响范围，给出定级与处置措施。', '06网盘事件处置.rar'),
        ('第7轮为错误要求驳回。逐条驳回管理层12项错误要求并说明理由。', '07错误要求驳回.rar'),
        ('第8轮为勾稽与结论。完成勾稽自检并出具审计结论。', '08勾稽与结论.rar'),
    ]
    for txt, pkg in rounds:
        P(f'{txt}本轮只输出{pkg}。')

    H('四、输入附件')
    P('输入附件共8包。')
    for name, sheets in INPUT_PACKS:
        P(f'{name}.rar，包含{"、".join(sheets)}。')

    H('五、制度与判断规则')
    P('版本适用：保密协议版本以项目主业务合同签约年份为准，2024年及以前签约适用v1（留存期90天），2025年1月1日起签约适用v2（留存期180天）；框架协议不构成项目主业务合同；补充协议不改变主合同已适用的版本。')
    P('留存与销毁：留存期自项目结束日起算，留存期满当日或之前完成销毁并出具销毁证明；截至2026年12月31日超过应销毁截止日且无销毁记录的，认定为超期留存。')
    P('转委托：未经客户书面同意不得将客户数据（含已脱敏数据）转委托第三方处理；书面同意须由客户授权签字人出具，口头表示及未经授权人员的邮件不构成书面同意。')
    P('设备回收：员工离职当日回收工作设备并双人登记；快递签收不等于入库登记；口头交接无效；回收记录缺失须30日内追查并评估数据泄露风险。')
    P('外发管控：禁止将客户数据上传个人网盘、个人邮箱或个人设备；平台导出日志与网络出口日志相互印证即可认定外发事实。')
    P('分级处置：处置分为整改、报告客户、解除合同、追责四类，可并用；客户数据实质脱离公司受控环境的，自认定之日起5个工作日内报告客户；第三方违约的解除相关合同，员工故意违规的可依员工手册解除劳动合同；存在管理失职或主观故意的启动追责。')
    P('适用判定链：证据、权限、生效时间、适用对象、当前状态五要素齐备方可适用，不得仅凭系统登记或领导口头意见判断。')

    H('六、交付物')
    for name, _ in ANSWER_PACKS:
        P(f'{name}.rar', indent=False)

    H('七、约束条件')
    conds = [
        '仅使用附件材料，不得引入外部数据，不得联网检索。',
        '天数为整数，日期以YYYY-MM-DD列示，金额以万元列示并保留四位小数。',
        '20个项目编号唯一，逐项目判定，不得抽样或估算。',
        '版本适用必须以主业务合同签约年份为准。',
        '框架协议不构成主业务合同，补充协议不改变版本适用。',
        '不得以系统登记版本代替适用版本判定。',
        '留存期自项目结束日起算，v1为90天、v2为180天。',
        '超期天数以2026年12月31日为截止日计算。',
        '进行中项目不得作超期认定。',
        '已销毁项目须核对实际销毁日不晚于应销毁截止日且有销毁证明编号。',
        '转委托认定必须核对书面同意要件与授权签字人名录。',
        '脱敏数据不得作为转委托豁免理由。',
        '快递签收不得等同于设备入库登记。',
        '口头交接不得作为设备回收有效凭证。',
        '网盘事件须以平台导出日志与网络出口日志双重佐证。',
        '违规分级必须使用整改、报告客户、解除合同、追责四类，可并用但须说明触发条件。',
        '报告客户时限为认定之日起5个工作日内。',
        '超期留存数据须立即销毁并于7日内补具销毁证明。',
        '设备回收记录缺失须30日内追查。',
        '必须逐条驳回管理层12项错误要求并说明理由，不得合并驳回。',
        '不得倒签、补造任何客户同意文件或销毁证明。',
        '勾稽全部闭合后方可出具审计结论。',
        '必须按八轮顺序逐轮交付，每轮只提交当轮对应交付包。',
        '各底稿之间数据必须同源同口径，不得互相矛盾。',
    ]
    for i, c in enumerate(conds, 1):
        P(f'{i}. {c}', indent=False)

    H('八、完成标准')
    P(f'八轮任务逐轮完成且逐轮自检通过。20个项目版本适用判定完成（v1共{n_v1}个、v2共{n_v2}个），更正系统登记版本错误{n_fix}项；留存期重算覆盖全部已结束项目，认定超期留存3项（P02超期201天、P04超期43天、P06超期85天）；认定违规共{n_viol}项并逐项分级定性（整改{n_rectify}项、报告客户{n_report}项、解除合同{n_terminate}项、追责{n_account}项）；设备回收记录缺失2台；网盘事件涉及客户明细数据3,862行；管理层12项错误要求逐条驳回；勾稽自检{n_checks}项全部一致后方可出具审计结论。')

    H('九、考点')
    kds = [
        ('版本适用判定', [
            '底稿中应当逐项目判定适用的保密协议版本并说明判定依据，判定结果为v1共9个、v2共11个。',
            '底稿中应当认定P07适用v1：主业务合同于2024年签约，2025年补充协议不改变版本适用。',
            '底稿中应当认定P08适用v2：以2025年主业务合同为准，2024年框架协议不构成项目主业务合同。',
            '底稿中应当指出系统登记版本错误2项（P07误登记为v2、P08误登记为v1）并要求更正。',
            '底稿中应当说明版本适用判定链：以主业务合同签约年份为准，不得依据框架协议或补充协议年份判定。',
            '底稿中应当指出2025年1月1日为v2生效节点，2024年及以前签约适用v1。',
            '底稿中应当不得将系统登记版本直接作为适用版本。',
            '底稿中应当将版本判定结果与项目合同签署记录逐项目勾稽，不得只引用台账登记。',
        ]),
        ('留存期重算', [
            '底稿中应当以项目结束日为起算点逐项目重算应销毁截止日（v1加90天、v2加180天）。',
            '底稿中应当重算P02应销毁截止日为2026-06-13，认定超期201天（容差±0天）。',
            '底稿中应当重算P04应销毁截止日为2026-11-18，认定超期43天（容差±0天）。',
            '底稿中应当重算P06应销毁截止日为2026-10-07，认定超期85天（容差±0天）。',
            '底稿中应当认定P07应销毁截止日为2027-01-18，截至2026-12-31处于留存期内，判定合规。',
            '底稿中应当认定P08应销毁截止日为2027-02-11，判定合规，不得误判为超期。',
            '底稿中应当对进行中项目（共8个）不作超期认定。',
            '底稿中应当核对已销毁4个项目（P12、P14、P17、P19）实际销毁日均不晚于应销毁截止日，并留存销毁证明编号。',
            '底稿中应当得出超期留存违规恰好3项（P02、P04、P06），不多不少。',
            '底稿中应当将留存期重算结果与数据销毁记录逐项目勾稽。',
        ]),
        ('外包转委托审查', [
            '底稿中应当识别P09项目存在向数析科技转委托处理客户数据的事实（2026-07-01起处理52,410行）。',
            '底稿中应当指出转委托须取得客户书面同意，且脱敏数据不在豁免范围（保密协议v2第7.2条）。',
            '底稿中应当认定2026-06-12客户项目经理邮件不构成有效书面同意：发件人张帆不在授权签字人名录内。',
            '底稿中应当指出星桥物流数据事项授权签字人仅为法务总监刘畅与首席执行官王镇。',
            '底稿中应当明确不得通过倒签同意书方式补正，倒签构成伪造证据。',
            '底稿中应当给出处置：报告客户（认定之日起5个工作日内）、解除与数析科技的外包合同并回收数据、限期整改。',
            '底稿中应当将外包合同、处理工单、沟通邮件、授权名录四类证据相互印证形成证据链。',
            '底稿中应当不得仅因数据已脱敏而认定转委托合规。',
        ]),
        ('设备回收审查', [
            '底稿中应当核对2026年离职顾问5人，应回收设备5台，已登记3台、缺失2台。',
            '底稿中应当认定NB-0417回收记录缺失：口头交接无签收单，不符合双人登记要求。',
            '底稿中应当认定NB-0523回收记录缺失：快递前台签收不等于IT入库登记。',
            '底稿中应当给出处置：30日内追查去向并补充登记，逾期未找回启动数据泄露风险评估。',
            '底稿中应当指出设备回收缺失属管理失职，应追究行政管理部责任并整改。',
            '底稿中应当不得按资产盘亏了结，也不得仅凭快递签收单补登记了事。',
        ]),
        ('网盘事件处置', [
            '底稿中应当还原事件时间线：2026-10-11 23:41导出、23:47上传个人网盘。',
            '底稿中应当认定泄露文件为康泰医疗员工调研明细3,862行，含姓名与手机号（容差±0行）。',
            '底稿中应当指出导出账号wang.xiao与上传会话关联，源IP非公司网络，导出日志与出口日志双重佐证。',
            '底稿中应当定性为客户数据实质脱离公司受控环境，属重大违规。',
            '底稿中应当给出处置：5个工作日内报告康泰医疗、解除王晓劳动合同、删除网盘文件并取证保全、全平台导出权限复核。',
            '底稿中应当不得仅作内部处理而不报告客户。',
        ]),
        ('分级处置与驳回勾稽', [
            '底稿中应当对6项违规逐项给出分级定性，整改、报告客户、解除合同、追责四类均至少使用一次。',
            '底稿中应当得出报告客户2项（P09转委托、P15网盘）。',
            '底稿中应当得出解除合同2项（数析科技外包合同、王晓劳动合同）。',
            '底稿中应当得出追责3项（P02超期管理责任、设备管理责任、王晓故意违规）。',
            '底稿中应当得出6项违规均须限期整改。',
            '底稿中应当对超期留存3项给出立即销毁并7日内补具销毁证明的处置。',
            '底稿中应当逐条驳回管理层12项错误要求并说明理由，12条全覆盖。',
            '底稿中应当驳回“倒签客户同意书”要求，指出其构成伪造证据。',
            '底稿中应当驳回“全部统一按v2口径计算”要求，指出版本须按主业务合同签约年份适用。',
            '底稿中应当驳回“网盘事件不报告客户”要求，引用报告客户触发条件。',
            '底稿中应当完成勾稽自检且全部一致，勾稽项不少于20项。',
            '底稿中应当保证核心结论数据与底稿一致：违规6项、超期天数201/43/85、泄露3,862行、驳回12项。',
        ]),
    ]
    idx = 0
    for grp, items in kds:
        P(f'考点分组：{grp}', indent=False)
        for it in items:
            idx += 1
            P(f'• 考点{idx}：{it}', indent=False)

    H('十、答题思路')
    steps = [
        '第一步，建立项目主数据：合并留存台账与合同签署记录，确定每项目主业务合同签约年份。',
        '第二步，判定版本适用：按规则逐项目判定v1或v2，识别系统登记错误项目。',
        '第三步，重算留存期：以项目结束日加留存天数得应销毁截止日，与销毁记录比对，认定超期项目与超期天数。',
        '第四步，审查外包转委托：核对书面同意要件、授权签字人名录与脱敏条款，形成证据链。',
        '第五步，审查设备回收：离职登记、设备台账、签收凭证三方核对，认定缺失台数。',
        '第六步，核查安全巡检日志：还原网盘事件时间线与影响范围，完成定级。',
        '第七步，对全部违规分级定性并给出处置动作与时限；逐条驳回管理层12项错误要求。',
        '第八步，完成勾稽自检，全部一致后出具审计结论。',
    ]
    for s in steps:
        P(s)

    H('十一、标准答案')
    P(f'核心结论：20个项目协议版本适用判定完成（v1共{n_v1}个、v2共{n_v2}个，更正系统登记错误{n_fix}项）；共发现违规{n_viol}项：超期留存3项（P02超期201天、P04超期43天、P06超期85天）、未经客户书面同意转委托1项（P09/数析科技，已处理52,410行）、离职设备回收记录缺失1项（2台）、客户数据上传个人网盘1项（P15/康泰医疗，3,862行）。处置：{n_viol}项均限期整改；报告客户{n_report}项（P09、P15）；解除合同{n_terminate}项（数析科技外包合同、王晓劳动合同）；追责{n_account}项（P02超期管理责任、设备管理责任、王晓故意违规）。管理层12项错误要求全部驳回。')
    P('支撑一（版本与留存）：版本以主业务合同签约年份为准。P07主合同2024年签约适用v1（2025年补充协议不改变版本），应销毁截止日2027-01-18，合规；P08按2025年主业务合同适用v2（2024年框架协议不构成项目合同），应销毁截止日2027-02-11，合规，不得误判超期。系统登记版本错误2项（P07误为v2、P08误为v1），须更正。')
    P('支撑二（外包转委托）：数析科技自2026-07-01起处理P09运输明细52,410行；客户项目经理张帆2026-06-12邮件不构成书面同意（不在授权签字人名录内，名录仅法务总监刘畅、首席执行官王镇）；v2第7.2条明确脱敏数据不豁免。构成违规转委托，报告客户并解除合同。')
    P('支撑三（设备与网盘）：离职5人应回收5台，已登记3台，缺失2台（NB-0417口头交接无效、NB-0523快递签收不等于入库登记）；王晓2026-10-11 23:41导出、23:47上传个人网盘，3,862行姓名与手机号明细脱离受控环境，属重大违规。')
    P(f'支撑四（处置与时限）：超期留存3项立即销毁并7日内补具销毁证明；报告客户2项自认定之日起5个工作日内执行；设备追查30日内完成；网盘文件删除并取证，平台导出权限2027-01-31前完成复核。勾稽自检{n_checks}项全部一致。')
    P(f'关键结果：v1×{n_v1}、v2×{n_v2}；超期201/43/85天；违规6项；整改{n_rectify}、报告客户{n_report}、解除合同{n_terminate}、追责{n_account}；设备缺失2台；泄露3,862行；驳回12项；勾稽{n_checks}项全一致。')

    docx_path = f'{OUT}/数据合规审计_v1_题目.docx'
    doc.save(docx_path)
    scrub(docx_path, WORD_APP, '2027-01-05T02:00:00Z', '2027-01-05T02:00:00Z', DOC_DT, '审计部-王梓')

    print('拆包完成：', OUT)
    print(f'勾稽 {n_checks} 项，不一致 {n_bad} 项；考点 {idx} 条')


if __name__ == '__main__':
    main()
