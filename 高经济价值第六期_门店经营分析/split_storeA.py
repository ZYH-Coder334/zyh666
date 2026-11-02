#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""门店经营分析 v1 附件拆分与打包：8输入 + 8答案，并生成题目 docx"""
import os
import io
import re
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

BASE = '/workspace/高经济价值第六期_门店经营分析'
SRC_Q = f'{BASE}/门店经营分析_v1_题目.xlsx'
SRC_A = f'{BASE}/门店经营分析_v1_标准答案.xlsx'
SPLIT = '/tmp/storeA_split'
OUT = f'{BASE}/门店经营分析_v1_附件'
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
    ('01门店基础与组织架构', ['门店基础信息', '城市与区域']),
    ('02门店Q1分月P&L', ['门店Q1分月P&L']),
    ('03POS流水与外卖结算', ['POS堂食流水', '外卖平台结算单']),
    ('04排班工时与工资', ['排班工时与工资']),
    ('05租约明细', ['租约明细']),
    ('06盘点损耗与总部分摊', ['月度盘点', '总部分摊规则']),
    ('07店长绩效考核办法', ['店长绩效考核办法']),
    ('08财务口径与管理层要求', ['财务确认规则', '管理层要求']),
]

ANSWER_PACKS = [
    ('01门店盈利分级', ['门店盈利分级', '分级标准']),
    ('02单店差异归因', ['单店差异归因']),
    ('03外卖对账差异', ['外卖对账差异']),
    ('04盘点损耗与责任', ['盘点损耗与责任']),
    ('05店长奖金核算', ['店长奖金核算']),
    ('06亏损店处置建议', ['亏损店处置建议']),
    ('07风险清单与退出机制', ['风险清单', '退出机制']),
    ('08勾稽自检与经营分析结论', ['错误要求驳回', '勾稽自检', '经营分析结论']),
]

IN_DT = (2026, 4, 2, 9, 0, 0)
AN_DT = (2026, 4, 10, 14, 0, 0)
split_workbook(SRC_Q, INPUT_PACKS, f'{OUT}/输入附件', 'in', '2026-04-01T01:00:00Z', '2026-04-02T01:00:00Z', IN_DT)
split_workbook(SRC_A, ANSWER_PACKS, f'{OUT}/标准答案', 'an', '2026-04-08T02:00:00Z', '2026-04-10T06:00:00Z', AN_DT)

# 读取答案指标
ans = openpyxl.load_workbook(SRC_A, data_only=True)
concl = {r[0]: r[1] for r in ans['经营分析结论'].iter_rows(values_only=True) if r[0]}
check_map = {r[0]: r[1] for r in ans['勾稽自检'].iter_rows(values_only=True) if r[0] != '勾稽项'}

total_rev = check_map['Q1营业收入合计(万元)']
total_ebitda = check_map['Q1 EBITDA合计(万元)']
total_net = check_map['Q1税后利润合计(万元)']
a_n = check_map['A级门店']
b_n = check_map['B级门店']
c_n = check_map['C级门店']
d_n = check_map['D级门店']
bonus = check_map['店长奖金合计(万元)']
deliv_diff = check_map['外卖对账差异门店数']
inv_resp = check_map['盘点超标门店数']

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
H(doc, '连锁门店2026年Q1经营分析与亏损店处置', 0)

H(doc, '一、编号')
P(doc, 'G6-ZX-004（v1）。', indent=False)

H(doc, '二、任务背景')
P(doc, '麦淳餐饮管理有限公司是一家连锁烘焙与轻食品牌，截至2026年3月共运营20家门店，分布在苏州、无锡、宁波、杭州、合肥、南昌六个城市，其中直营18家、加盟2家。2026年4月，公司召开一季度经营分析会，发现杭州、合肥、南昌新店爬坡不及预期，无锡惠山万达、南昌青山湖万象汇等门店连续亏损。现受运营总监与财务总监委托，以运营分析岗身份，对20家门店Q1经营数据完成盈利分级、亏损归因、外卖对账、盘点责任划分、店长奖金核算和亏损店处置建议。')
P(doc, '数据截止日为2026年3月31日，分析基准日为2026年4月5日。所有结论须可回溯到附件逐条记录，不得只给汇总数。凡涉及分级、归因、责任、扣减、处置的记录，必须逐条说明触发条件与依据。')

H(doc, '三、任务分解')
P(doc, '本项委托按八轮执行，每轮只完成一个任务，并只提交当轮对应交付包。上一轮自检通过后，方可进入下一轮；未完成或自检不通过，不得进入下一轮，也不得提前输出后续轮次结果。')
P(doc, '第1轮为门店盈利分级。完成20家门店Q1营业收入、EBITDA、EBITDA率计算，按A级≥12%、B级8%-12%、C级4%-8%、D级<4%分级。本轮只输出01门店盈利分级.rar。本轮自检要求20家门店均有分级，分级标准与制度一致。')
P(doc, '第2轮为单店差异归因。对比预算与实际营收、EBITDA，逐店归因租金、外卖佣金、损耗等差异。本轮只输出02单店差异归因.rar。本轮自检要求20家门店均有差异归因，C/D级门店有明确主因。')
P(doc, '第3轮为外卖对账差异。逐店核对外卖GMV、平台佣金、平台实收与门店入账，识别少计、多计收入。本轮只输出03外卖对账差异.rar。本轮自检要求差异门店均有原因说明。')
P(doc, '第4轮为盘点损耗与责任。逐店计算Q1平均差异率，超2%认定门店责任。本轮只输出04盘点损耗与责任.rar。本轮自检要求超标门店均有责任认定。')
P(doc, '第5轮为店长奖金核算。按基础奖金、EBITDA达成率、营收达成率、损耗达标计算奖金系数，D级扣减50%。本轮只输出05店长奖金核算.rar。本轮自检要求20家店长奖金均有计算，D级门店标注追回。')
P(doc, '第6轮为亏损店处置建议。对C/D级门店给出整改、转址、关店建议。本轮只输出06亏损店处置建议.rar。本轮自检要求D级门店均有处置路径。')
P(doc, '第7轮为风险清单与退出机制。列出10项风险、量化触发阈值、应对措施和退出动作。本轮只输出07风险清单与退出机制.rar。本轮自检要求10项风险均有阈值，6类退出机制均有回收与损失处理。')
P(doc, '第8轮为勾稽自检与经营分析结论。完成24项勾稽并给出一句话核心结论。本轮只输出08勾稽自检与经营分析结论.rar。本轮自检要求24项全部闭合、不一致项为0；未闭合不得出具结论。')

H(doc, '四、输入附件')
P(doc, '输入附件共8包。')
P(doc, '01门店基础与组织架构.rar，包含门店基础信息、城市与区域。')
P(doc, '02门店Q1分月P&L.rar，包含门店Q1分月P&L。')
P(doc, '03POS流水与外卖结算.rar，包含POS堂食流水、外卖平台结算单。')
P(doc, '04排班工时与工资.rar，包含排班工时与工资。')
P(doc, '05租约明细.rar，包含租约明细。')
P(doc, '06盘点损耗与总部分摊.rar，包含月度盘点、总部分摊规则。')
P(doc, '07店长绩效考核办法.rar，包含店长绩效考核办法。')
P(doc, '08财务口径与管理层要求.rar，包含财务确认规则、管理层要求。')

H(doc, '五、制度与判断规则')
P(doc, '门店共20家，直营18家、加盟2家。加盟门店P&L按门店口径核算，公司仅确认授权费与供应链毛利，不并入直营利润。')
P(doc, '营业收入等于堂食净额加外卖GMV，外卖GMV含平台补贴，退款冲减收入。堂食净额等于堂食应收减堂食退款。')
P(doc, '外卖平台佣金仅对外卖GMV按18%计提，不得对堂食收入计提。平台实收等于外卖GMV减平台佣金减配送费减退款加平台补贴。门店入账应等于外卖GMV。')
P(doc, '食材成本率36%、人工成本率21%、水电能耗费率3%、营销推广费率4%、管理与中央厨房分摊率6%，均按营业收入计提。人工含店长、烘焙、服务、兼职与社保公积金。')
P(doc, '租金含物业费，按租约月租金乘面积计算，Q1无免租期剩余。折旧按建店投资5年直线计提，残值为0，月折旧等于建店投资除以60。')
P(doc, '所得税按正税前利润的25%计提，亏损年度所得税为0。税前利润等于EBITDA减折旧。')
P(doc, '门店分级：A级EBITDA率不低于12%，B级8%至12%，C级4%至8%，D级低于4%或连续亏损。D级门店启动整改，连续2季D级启动关店或转址。')
P(doc, '盘点差异率不超过2%为正常损耗，超过2%由门店承担并从店长奖金扣除。差异金额计入食材成本。')
P(doc, '店长奖金等于基础奖金乘奖金系数。基础奖金A级1.2万元、B级1万元、C级0.8万元、D级0.5万元，D级扣减50%。奖金系数等于EBITDA达成率乘0.5加营收达成率乘0.3加损耗达标得分乘0.2，各项上限1.5。损耗达标得分为1减超过2%部分乘以10，最低为0。')
P(doc, '预算营收等于稳定月营收乘3，预算EBITDA按标准费率测算。')

H(doc, '六、交付物')
for t in ['01 门店盈利分级.rar', '02 单店差异归因.rar', '03 外卖对账差异.rar',
          '04 盘点损耗与责任.rar', '05 店长奖金核算.rar', '06 亏损店处置建议.rar',
          '07 风险清单与退出机制.rar', '08 勾稽自检与经营分析结论.rar']:
    P(doc, t, indent=False)

H(doc, '七、约束条件')
constraints = [
    '1. 仅使用附件材料，不得引入外部数据，不得联网检索。',
    '2. 对外报表金额以万元列示并保留两位小数，比例保留两位；底层计算以元为单位，不得混用。',
    '3. 逐门店给出结论，不得抽样或估算；每条记录必须有明确状态，不得留空。',
    '4. 20家门店编号唯一，不得跨城市重复计入；六城门店数为4/4/4/3/3/2。',
    '5. 加盟门店P&L不并入公司利润，公司仅确认授权费与供应链毛利。',
    '6. 营业收入等于堂食净额加外卖GMV，外卖GMV含平台补贴，退款冲减收入。',
    '7. 外卖平台佣金只对外卖GMV按18%计提，不得对堂食收入或全部收入计提。',
    '8. 门店入账外卖收入应等于外卖GMV，不得按平台实收入账。',
    '9. 食材、人工、水电、营销、管理分摊比例固定，不得逐店调整。',
    '10. 租金含物业费，Q1无免租期剩余，不得忽略物业费。',
    '11. 折旧按5年直线计提，月折旧等于建店投资除以60。',
    '12. 所得税按正税前利润的25%计提，亏损年度所得税为0。',
    '13. 门店分级按EBITDA率，A级≥12%、B级8%-12%、C级4%-8%、D级<4%。',
    '14. D级门店启动整改，连续2季D级才关店或转址，不得立即关店。',
    '15. 盘点差异率超过2%认定门店责任，从店长奖金扣除。',
    '16. 店长奖金系数为EBITDA达成×0.5+营收达成×0.3+损耗达标×0.2，各项上限1.5。',
    '17. D级门店基础奖金扣减50%，并追回已发预提。',
    '18. 必须列出10项风险和6类退出机制，每项风险须有量化触发阈值。',
    '19. 24项勾稽全部闭合后方可出具核心结论；任一项不闭合视为未完成。',
    '20. 必须按八轮顺序逐轮交付，每轮只完成一个任务，不得一次性输出全部结果。',
    '21. 每轮结束必须提交本轮自检记录，列明检查项、结果、证据和未通过原因；上一轮未通过不得进入下一轮。',
    '22. 每轮只提交当轮对应交付包，不得跨轮合并或提前输出；任何一轮无法完成时必须停止并说明缺口；八轮全部通过方视为完成。',
    '23. 所有外卖对账差异必须逐条说明原因，不得合并为差异总额。',
    '24. 所有处置建议必须有明确触发条件与时间节点，不得只给方向性意见。',
]
for t in constraints:
    P(doc, t, indent=False)

H(doc, '八、完成标准')
P(doc, f'八轮任务逐轮完成且逐轮自检通过。20家门店全部分级，A级{a_n}家、B级{b_n}家、C级{c_n}家、D级{d_n}家；Q1合计营收{total_rev}万元、EBITDA {total_ebitda}万元、税后利润{total_net}万元；外卖对账差异{deliv_diff}家、盘点超标{inv_resp}家；店长奖金合计{bonus}万元；10项风险和6类退出机制完整；24项勾稽全部闭合、不一致项为0。任一轮未完成或自检不通过，不得进入下一轮，整体不得视为完成。')

H(doc, '九、考点')
P(doc, '考点按数据治理、收入确认、成本费用、外卖对账、盘点责任、门店分级、差异归因、店长奖金、亏损处置、风险退出、勾稽结论十一组展开，共50项。', indent=False)
groups = [
    ('（一）数据治理与门店清单（考点1—5）', [
        '考点1：门店共20家，苏州4家、无锡4家、宁波4家、杭州3家、合肥3家、南昌2家，编号唯一。',
        '考点2：直营18家、加盟2家，加盟门店为HF03和NC03，利润不并入公司。',
        '考点3：门店基础信息含面积、开业日、店长、稳定月营收、外卖占比、租金占比、建店投资。',
        '考点4：租约明细含月租金、物业费、免租期剩余、年递增、押金月数、合同到期日。',
        '考点5：Q1为2026年1月至3月，12月开业门店有爬坡系数0.95。',
    ]),
    ('（二）收入确认（考点6—10）', [
        '考点6：营业收入等于堂食净额加外卖GMV，不得用堂食应收或平台实收。',
        '考点7：堂食净额等于堂食应收减堂食退款，退款冲减收入。',
        '考点8：外卖GMV含平台补贴，平台补贴应计入收入。',
        '考点9：门店入账外卖收入应等于外卖GMV，SZ04按实收入账少计收入。',
        '考点10：NC03退款未冲减收入，导致多计收入。',
    ]),
    ('（三）成本与费用（考点11—15）', [
        '考点11：食材成本率36%、人工21%、水电3%、营销4%、管理分摊6%，按营业收入计提。',
        '考点12：人工成本含店长、烘焙、服务、兼职与社保公积金，社保按应发工资28%计提。',
        '考点13：租金含物业费，按月租金加物业费乘面积计算。',
        '考点14：折旧按建店投资5年直线计提，月折旧等于建店投资除以60。',
        '考点15：所得税按正税前利润的25%计提，亏损年度所得税为0。',
    ]),
    ('（四）外卖佣金与对账（考点16—20）', [
        '考点16：外卖平台佣金只对外卖GMV按18%计提，不得对堂食收入计提。',
        '考点17：平台实收等于外卖GMV减佣金减配送费减退款加平台补贴。',
        '考点18：门店入账应等于外卖GMV，与平台实收的差额为佣金、配送费、退款和补贴。',
        '考点19：WX04平台补贴未计入收入，导致少计收入。',
        '考点20：外卖对账差异共3家，分别为SZ04、WX04、NC03。',
    ]),
    ('（五）盘点损耗与责任（考点21—25）', [
        '考点21：盘点差异率等于差异金额除以账面库存。',
        '考点22：差异率不超过2%为正常损耗，超过2%认定门店责任。',
        '考点23：WX04和NB04盘点差异率超标，认定门店责任。',
        '考点24：差异金额计入食材成本，从店长奖金扣除。',
        '考点25：损耗达标得分等于1减超过2%部分乘以10，最低为0。',
    ]),
    ('（六）门店分级（考点26—30）', [
        '考点26：A级EBITDA率不低于12%，B级8%至12%，C级4%至8%，D级低于4%。',
        f'考点27：20家门店分级为A级{a_n}家、B级{b_n}家、C级{c_n}家、D级{d_n}家。',
        '考点28：D级门店为HZ04、WX04、NC03，启动整改。',
        '考点29：EBITDA率等于Q1 EBITDA除以Q1营业收入。',
        '考点30：加盟门店按门店口径分级，不影响公司利润并表。',
    ]),
    ('（七）预算与差异归因（考点31—35）', [
        '考点31：预算营收等于稳定月营收乘3，预算EBITDA按标准费率测算。',
        '考点32：营收差异等于实际营收减预算营收，EBITDA差异同理。',
        '考点33：C/D级门店主要归因为租金占比过高、客流下滑或损耗超标。',
        '考点34：WX04租金占收入超过15%且客流下滑，是亏损主因。',
        '考点35：差异归因须逐店给出，不得合并说明。',
    ]),
    ('（八）店长奖金（考点36—40）', [
        '考点36：基础奖金A级1.2万、B级1万、C级0.8万、D级0.5万。',
        '考点37：奖金系数等于EBITDA达成率乘0.5加营收达成率乘0.3加损耗达标得分乘0.2。',
        '考点38：EBITDA达成率和营收达成率上限1.5，最低为0。',
        '考点39：D级门店基础奖金扣减50%，并追回已发预提。',
        f'考点40：20家店长奖金合计{bonus}万元。',
    ]),
    ('（九）亏损店处置（考点41—44）', [
        '考点41：D级门店启动整改，连续2季D级才关店或转址。',
        '考点42：C级门店限期整改，Q2复核。',
        '考点43：A级门店保持经营，可作为标杆复制。',
        '考点44：处置建议须有明确触发条件与时间节点。',
    ]),
    ('（十）风险与退出（考点45—47）', [
        '考点45：列出客流下滑、租金占比过高、外卖佣金上调、盘点损耗超标、新店爬坡不及预期、店长流失、食安证照、加盟商违约、总部分摊争议、竞品进入共10项风险。',
        '考点46：连续2季EBITDA率低于4%，或租金占收入连续2季超过18%，启动关店或转址。',
        '考点47：退出机制含设备回收、装修损失、押金退还、保证金抵扣等处理。',
    ]),
    ('（十一）勾稽与核心结论（考点48—50）', [
        '考点48：24项勾稽覆盖门店数、城市分布、分级、营收、EBITDA、利润、外卖差异、盘点超标、奖金，全部闭合。',
        '考点49：加盟门店利润不并表，公司仅确认授权费与供应链毛利。',
        f'考点50：核心结论为20家门店Q1合计营收{total_rev}万元、EBITDA {total_ebitda}万元、税后利润{total_net}万元，D级门店启动整改。',
    ]),
]
for title, items in groups:
    H(doc, title, 2)
    for t in items:
        P(doc, t, indent=False)

H(doc, '十、答题思路')
P(doc, '第一步是读取门店基础、租约、P&L、POS、外卖结算、排班、盘点和考核办法，锁定20家门店及六城分布。')
P(doc, '第二步是按制度费率重新计算每家门店Q1的营收、成本、EBITDA和税后利润，与P&L核对。')
P(doc, '第三步是按EBITDA率对20家门店分级，识别A级、B级、C级、D级门店。')
P(doc, '第四步是对比预算与实际，逐店做营收和EBITDA差异归因，重点分析C/D级门店。')
P(doc, '第五步是逐店核对外卖GMV与门店入账，识别少计、多计收入的门店并说明原因。')
P(doc, '第六步是计算每家门店Q1平均盘点差异率，认定门店责任并计算对奖金的影响。')
P(doc, '第七步是按基础奖金、达成率和损耗得分计算每家店长奖金，D级扣减50%并标注追回。')
P(doc, '第八步是对C/D级门店给出处置建议，明确整改、转址、关店的触发条件与时间节点。')
P(doc, '第九步是完成10项风险、6类退出机制和24项勾稽，最后给出一句话核心结论。')

H(doc, '十一、标准答案')
P(doc, '标准答案共8个包，首表为经营分析结论，逐包覆盖盈利分级、差异归因、外卖对账、盘点责任、店长奖金、处置建议、风险退出和勾稽自检。', indent=False)
P(doc, concl['核心结论'], bold=True)
for k in ['支撑一', '支撑二', '支撑三', '支撑四']:
    if k in concl:
        P(doc, concl[k])
P(doc, f'关键结果：20家门店Q1合计营收{total_rev}万元、EBITDA {total_ebitda}万元、税后利润{total_net}万元；A级{a_n}家、B级{b_n}家、C级{c_n}家、D级{d_n}家；外卖对账差异{deliv_diff}家、盘点超标{inv_resp}家；店长奖金合计{bonus}万元；24项勾稽全部闭合，不一致项为0。')

doc.core_properties.author = '运营分析部'
doc.save(f'{OUT}/门店经营分析_v1_题目.docx')
scrub(f'{OUT}/门店经营分析_v1_题目.docx', WORD_APP, '运营分析部', '2026-04-03T02:00:00Z', '2026-04-05T01:00:00Z', (2026, 4, 5, 10, 0, 0))

print('done')
print('total_rev', total_rev, 'ebitda', total_ebitda, 'net', total_net)
print('grades', a_n, b_n, c_n, d_n, 'bonus', bonus)
print('packs', len(INPUT_PACKS), len(ANSWER_PACKS))
