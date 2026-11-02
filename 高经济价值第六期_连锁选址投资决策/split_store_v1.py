#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""连锁选址投资决策 v1 附件拆分与打包：8输入 + 8答案，并生成题目 docx"""
import os
import io
import re
import shutil
import zipfile
import time
from copy import copy
from datetime import date

import openpyxl
from openpyxl import Workbook
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

BASE = '/workspace/高经济价值第六期_连锁选址投资决策'
SRC_Q = f'{BASE}/连锁选址投资决策_v1_题目.xlsx'
SRC_A = f'{BASE}/连锁选址投资决策_v1_标准答案.xlsx'
SPLIT = '/tmp/store_v1_split'
OUT = f'{BASE}/连锁选址投资决策_v1_附件'
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
        scrub(tmp, EXCEL_APP, '投资发展部', created, modified, zip_dt)
        make_rar([tmp], f'{out_dir}/{name}.rar', zip_dt)
        print('packed', name, len(sheets))


INPUT_PACKS = [
    ('01城市与商圈数据', ['城市基础数据', '候选点位清单']),
    ('02租约与物业条件', ['租约关键条款']),
    ('03销售预测与外卖', ['点位销售预测']),
    ('04建店投资与成本', ['建店投资标准', '城市投资系数', '运营成本参数']),
    ('05直营与加盟合同', ['直营管理要求', '加盟合同摘要']),
    ('06区域仓配与现有门店', ['区域仓配方案', '现有门店P&L']),
    ('07开业计划与团队', ['开业产能与团队']),
    ('08财务口径与管理要求', ['财务确认规则', '管理层要求']),
]

ANSWER_PACKS = [
    ('01商圈与点位评分', ['点位评分', '评分权重']),
    ('02单店P&L与投资回收', ['单店P&L汇总', '单店P&L逐年']),
    ('03直营与加盟对比', ['直营与加盟对比']),
    ('04首批8店清单与开业节奏', ['首批8店清单与开业节奏']),
    ('05区域仓配与投资总额', ['仓配方案比选', '投资总额汇总', '分城市仓配']),
    ('06三情景敏感性与回报', ['三情景敏感性']),
    ('07风险清单与退出机制', ['风险清单', '退出机制']),
    ('08勾稽自检与投资决策结论', ['错误要求驳回', '勾稽自检', '投资决策结论']),
]

IN_DT = (2025, 10, 15, 8, 30, 0)
AN_DT = (2025, 10, 25, 9, 10, 0)
split_workbook(SRC_Q, INPUT_PACKS, f'{OUT}/输入附件', 'in', '2025-10-08T01:20:00Z', '2025-10-15T00:30:00Z', IN_DT)
split_workbook(SRC_A, ANSWER_PACKS, f'{OUT}/标准答案', 'an', '2025-10-20T02:00:00Z', '2025-10-25T01:10:00Z', AN_DT)


# ---------------- 读取答案指标 ----------------
ans = openpyxl.load_workbook(SRC_A, data_only=True)
concl = {r[0]: r[1] for r in ans['投资决策结论'].iter_rows(values_only=True) if r[0]}
check_map = {r[0]: r[1] for r in ans['勾稽自检'].iter_rows(values_only=True) if r[0] != '勾稽项'}
sens = [r for r in ans['三情景敏感性'].iter_rows(values_only=True) if r[0] != '情景编号']
wh_rows = [r for r in ans['分城市仓配'].iter_rows(values_only=True) if r[0] != '城市']
total_invest = check_map['首批总投资(万元)']
first_rev = check_map['首年营业收入(万元)']
first_ebitda = check_map['首年EBITDA(万元)']
five_ebitda = check_map['5年累计EBITDA(万元)']
port_npv = check_map['组合NPV(万元)']
port_irr = check_map['组合IRR(%)']
payback = re.search(r'回收期([\d.]+)个月', concl['核心结论']).group(1)
sens_min = min(r[5] for r in sens)
sens_max = max(r[5] for r in sens)
wh_text = '、'.join(f'{r[0]}{r[2]}' for r in wh_rows)


# ---------------- 题目 docx ----------------
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
H(doc, '连锁品牌新区域扩张的选址与单店投资决策', 0)

H(doc, '一、编号')
P(doc, 'G6-ZX-003（v1）。', indent=False)

H(doc, '二、任务背景')
P(doc, '麦淳餐饮管理有限公司是一家连锁烘焙与轻食品牌，现有12家直营门店，集中在苏州、无锡、宁波。2025年10月，公司计划进入杭州、合肥、南昌三个新区域，首批开设8家门店。现受投资委员会委托，以外部战略与投资顾问身份，在18个候选点位中完成首批8店的选址、单店投资测算、直营与加盟模式比较、开业节奏、仓配方案、风险与退出机制设计。')
P(doc, '数据截止日为2025年9月30日，投资决策基准日为2025年10月1日。所有结论须可回溯到附件逐条记录，不得只给汇总数。凡涉及剔除、备选、退出、不投、转址的记录，必须逐条说明触发条件与依据。')

H(doc, '三、任务分解')
P(doc, '本项委托按八轮执行，每轮只完成一个任务，并只提交当轮对应交付包。上一轮自检通过后，方可进入下一轮；未完成或自检不通过，不得进入下一轮，也不得提前输出后续轮次结果。')
P(doc, '第1轮为商圈与点位评分。完成18个候选点位的资格筛查、七因子评分和城市内排序，按杭州3家、合肥3家、南昌2家形成首批8店建议。本轮只输出01商圈与点位评分.rar。本轮自检要求18个点位均有评分，权重合计100分，入选8家且城市分布为3/3/2。')
P(doc, '第2轮为单店P&L与投资回收。完成8家门店的建店投资、首年与稳定期收入、五年P&L、回收期、NPV和IRR测算。本轮只输出02单店P&L与投资回收.rar。本轮自检要求8家门店均有逐年P&L，直接门店回收期不超过42个月，加盟门店单列公司口径。')
P(doc, '第3轮为直营与加盟对比。对8家门店分别测算直营和加盟两种模式的公司现金流与NPV，给出推荐模式与理由。本轮只输出03直营与加盟对比.rar。本轮自检要求直营6家、加盟2家，加盟比例25%，不超过30%上限。')
P(doc, '第4轮为首批8店清单与开业节奏。按租约起租日加2个月形成计划开业日，并按季度产能排布。本轮只输出04首批8店清单与开业节奏.rar。本轮自检要求Q1 3家、Q2 3家、Q3 2家，单季不超过3家。')
P(doc, '第5轮为区域仓配与投资总额。比较杭州、合肥、南昌的自建城市仓、第三方仓配和区域中心仓三种方案，汇总首批总投资。本轮只输出05区域仓配与投资总额.rar。本轮自检要求三城均选出5年成本现值最低方案，总投资不超过1500万元。')
P(doc, '第6轮为三情景敏感性与回报。完成销售变动±10%与租金变动±5%的9组组合测算。本轮只输出06三情景敏感性与回报.rar。本轮自检要求9组情景齐全，基准情景与第2轮、第5轮结果一致。')
P(doc, '第7轮为风险清单与退出机制。列出10项主要风险、量化触发阈值、应对措施和退出动作。本轮只输出07风险清单与退出机制.rar。本轮自检要求10项风险均有阈值，6类退出机制均有回收与损失处理。')
P(doc, '第8轮为勾稽自检与投资决策结论。完成24项勾稽并给出一句话核心结论。本轮只输出08勾稽自检与投资决策结论.rar。本轮自检要求24项全部闭合、不一致项为0；未闭合不得出具结论。')

H(doc, '四、输入附件')
P(doc, '输入附件共8包。')
P(doc, '01城市与商圈数据.rar，包含城市基础数据、候选点位清单。')
P(doc, '02租约与物业条件.rar，包含租约关键条款。')
P(doc, '03销售预测与外卖.rar，包含点位销售预测。')
P(doc, '04建店投资与成本.rar，包含建店投资标准、城市投资系数、运营成本参数。')
P(doc, '05直营与加盟合同.rar，包含直营管理要求、加盟合同摘要。')
P(doc, '06区域仓配与现有门店.rar，包含区域仓配方案、现有门店P&L。')
P(doc, '07开业计划与团队.rar，包含开业产能与团队。')
P(doc, '08财务口径与管理要求.rar，包含财务确认规则、管理层要求。')

H(doc, '五、制度与判断规则')
P(doc, '候选点位共18个，杭州、合肥、南昌各6个。点位按商圈日均客流25%、周边竞品数15%、外卖平台月单量15%、租金效率15%、人均可支配收入10%、可见性10%、物业条件10%评分，权重合计100分。各指标在18个点位间做线性归一化，竞品数按反向归一化。')
P(doc, '租金效率按年化稳定收入除以年租金计算。年化稳定收入为堂食与外卖稳定月收入乘以12。堂食收入为堂食日均单量乘堂食客单价乘30天，外卖收入为外卖日均单量乘外卖客单价乘30天。')
P(doc, '首年收入按爬坡计算，前3个月60%、第4至6个月80%、第7个月起100%，等效10.2个满月。第2年起按3%增长。外卖平台佣金仅对外卖收入按18%计提。')
P(doc, '建店投资包括装修、厨房与冷藏设备、家具与陈列、POS与信息系统、首批物料、证照与开业、设计与管理费、押金和流动资金。装修按3200元/㎡乘城市系数，设备按55万元乘城市系数，家具11万元，POS 5.5万元，首批物料7.5万元，证照3.5万元，设计费为装修的5%，押金为3个月租金，流动资金23万元。')
P(doc, '单店P&L中，食材成本率38%、人工成本率23%、水电能耗费率3%、营销推广费率5%、管理与中央厨房分摊率6%。租金含物业费，首年扣减免租期，之后每年递增5%。折旧按扣除押金与流动资金后的建店投资分5年直线计提。所得税按正税前利润的25%计提。')
P(doc, '回收期为累计未折现现金流首次转正的月份，超过60个月按“>60个月”列示。NPV按10%折现，期末收回押金与流动资金。IRR按5年现金流计算。直接门店回收期超过42个月须专项审批。')
P(doc, '直营模式由公司承担全部建店投资并取得全部收入。加盟模式由加盟商承担85%建店投资，公司收取15万元加盟费、6%持续授权费、18%供应链毛利，承担3%督导培训成本和每店每年8万元固定督导费，公司初期支持投入10万元。')
P(doc, '每城首店必须直营，首批加盟比例不超过30%。计划开业日按租约起租日加2个月计算，项目组每季最多开业3家。')
P(doc, '仓配方案按5年成本现值比较，包含资本支出、年固定成本和按营业收入计提的变动成本。首批总投资包括直营建店投资、加盟支持投入、仓配资本支出和5%预备费，总额不超过1500万元。')

H(doc, '六、交付物')
for t in ['01 商圈与点位评分.rar', '02 单店P&L与投资回收.rar', '03 直营与加盟对比.rar',
          '04 首批8店清单与开业节奏.rar', '05 区域仓配与投资总额.rar', '06 三情景敏感性与回报.rar',
          '07 风险清单与退出机制.rar', '08 勾稽自检与投资决策结论.rar']:
    P(doc, t, indent=False)

H(doc, '七、约束条件')
constraints = [
    '1. 仅使用附件材料，不得引入外部数据，不得联网检索。',
    '2. 对外报表金额以万元列示并保留两位小数，比例保留两位；底层计算以元为单位，不得混用。',
    '3. 逐点位、逐门店、逐城市给出结论，不得抽样或估算；每条记录必须有明确状态，不得留空。',
    '4. 18个候选点位编号唯一，不得跨城市重复计入；三城各6个候选点位。',
    '5. 餐饮条件为“否”或可签年限不足5年的点位，不得进入首批名单，但须保留评分与不符合说明。',
    '6. 评分权重固定为25/15/15/15/10/10/10，合计100分；不得临时增删指标。',
    '7. 所有评分指标在18个点位间做线性归一化，竞品数按反向归一化，不得按城市内部分别归一化。',
    '8. 销售预测必须拆分堂食与外卖，不得用总单量乘单一客单价。',
    '9. 外卖平台佣金只对外卖收入按18%计提，不得对全部收入计提。',
    '10. 建店投资必须包含押金和流动资金，不得只算装修与设备。',
    '11. 单店P&L的食材、人工、水电、营销、管理分摊比例固定，不得逐店调整。',
    '12. 租金首年扣减免租期，之后按5%递增；不得忽略免租期或递增。',
    '13. 折旧按5年直线计提，折旧基数为扣除押金与流动资金后的建店投资。',
    '14. 回收期、NPV、IRR按统一口径计算；NPV折现率10%，期末收回押金与流动资金。',
    '15. 每城首店必须直营；首批加盟门店不超过2家，加盟比例不超过30%。',
    '16. 计划开业日按租约起租日加2个月计算，单季开业不超过3家。',
    '17. 仓配方案按5年成本现值比较，三城分别选择成本最低方案。',
    '18. 首批总投资包括5%预备费，总额不得超过1500万元。',
    '19. 必须完成9组“销售±10%×租金±5%”敏感性测算，不得只做单一情景。',
    '20. 必须列出10项风险和6类退出机制，每项风险须有量化触发阈值。',
    '21. 24项勾稽全部闭合后方可出具核心结论；任一项不闭合视为未完成。',
    '22. 必须按八轮顺序逐轮交付，每轮只完成一个任务，不得一次性输出全部结果。',
    '23. 每轮结束必须提交本轮自检记录，列明检查项、结果、证据和未通过原因；上一轮未通过不得进入下一轮。',
    '24. 每轮只提交当轮对应交付包，不得跨轮合并或提前输出；任何一轮无法完成时必须停止并说明缺口；八轮全部通过方视为完成。',
]
for t in constraints:
    P(doc, t, indent=False)

H(doc, '八、完成标准')
P(doc, f'八轮任务逐轮完成且逐轮自检通过。18个候选点位全部评分，首批8家门店按杭州3家、合肥3家、南昌2家选定；直营6家、加盟2家；总投资{total_invest}万元，不超过1500万元；首年收入{first_rev}万元、首年EBITDA {first_ebitda}万元；组合5年NPV {port_npv}万元、IRR {port_irr}%、回收期{payback}个月；9组敏感性齐全；10项风险和6类退出机制完整；24项勾稽全部闭合、不一致项为0。任一轮未完成或自检不通过，不得进入下一轮，整体不得视为完成。')

H(doc, '九、考点')
P(doc, '考点按数据治理、评分模型、销售预测、建店投资、单店P&L、投资回收、直营加盟、开业节奏、仓配、投资总额、敏感性、风险退出和勾稽结论十三组展开，共50项。', indent=False)
groups = [
    ('（一）数据治理与候选点位（考点1—5）', [
        '考点1：候选点位共18个，杭州、合肥、南昌各6个，点位编号唯一且不得跨城市重复计入。',
        '考点2：面积、月租金、物业费、免租期、递增率、押金月数和可签年限必须按同一租约口径读取，不得混用含税与不含税口径。',
        '考点3：可起租日用于排定计划开业日，计划开业日等于起租日加2个月，不得直接使用决策日或装修完成日。',
        '考点4：餐饮条件为“否”的点位不得进入首批，HZ05和NC02因排烟或证照条件不符被剔除。',
        '考点5：可签年限不足5年的点位不得进入首批，HF04因仅3年租约、HF05因仅4年租约被剔除。',
    ]),
    ('（二）评分模型（考点6—10）', [
        '考点6：评分权重为客流25%、竞品15%、外卖15%、租金效率15%、收入10%、可见性10%、物业10%，合计100分。',
        '考点7：所有指标在18个候选点位之间做线性归一化，不得按城市内部分别归一化。',
        '考点8：周边竞品数采用反向归一化，竞品越多得分越低，HZ03因15家竞品被显著扣分。',
        '考点9：租金效率按年化稳定收入除以年租金计算，租金效率越高得分越高。',
        '考点10：按城市内符合资格点位排序，杭州取前3、合肥取前3、南昌取前2，形成首批8家。',
    ]),
    ('（三）销售预测与外卖（考点11—15）', [
        '考点11：堂食收入等于堂食日均单量乘堂食客单价乘30天，不得与外卖单量合并。',
        '考点12：外卖收入等于外卖日均单量乘外卖客单价乘30天，外卖客单价与堂食客单价分开。',
        '考点13：首年按前3月60%、第4至6月80%、第7月起100%爬坡，等效10.2个满月。',
        '考点14：稳定期第2年起按3%增长，不得在首年就按满月12个月计算。',
        '考点15：外卖平台佣金只对外卖收入按18%计提，不得对堂食收入计提。',
    ]),
    ('（四）建店投资（考点16—20）', [
        '考点16：装修按3200元/㎡乘城市系数，杭州1.10、合肥1.00、南昌0.92。',
        '考点17：厨房与冷藏设备按55万元乘城市系数，杭州1.08、合肥1.00、南昌0.95。',
        '考点18：家具与陈列11万元、POS与信息系统5.5万元、首批物料7.5万元、证照与开业3.5万元。',
        '考点19：设计与管理费按装修工程的5%计提，不得按总投资计提。',
        '考点20：押金按3个月租金计入初始投资，流动资金23万元，两者期末可收回。',
    ]),
    ('（五）单店P&L（考点21—25）', [
        '考点21：单店P&L按营业收入减食材、人工、租金、水电、营销、外卖佣金、管理分摊和加盟授权费计算EBITDA。',
        '考点22：食材成本率38%、人工23%、水电3%、营销5%、管理分摊6%，不得逐店调整。',
        '考点23：租金含物业费，首年扣减免租期，第2年起按5%递增。',
        '考点24：折旧按扣除押金与流动资金后的建店投资分5年直线计提，残值为0。',
        '考点25：所得税按正税前利润的25%计提，亏损年度所得税为0。',
    ]),
    ('（六）回收期、NPV与IRR（考点26—30）', [
        '考点26：回收期按累计未折现现金流首次转正的月份计算，超过60个月按“>60个月”列示。',
        '考点27：NPV按10%折现，期末收回押金与流动资金，不得计入其他残值。',
        '考点28：IRR按5年公司口径现金流计算，加盟店按加盟费、授权费、供应链毛利和督导成本计算。',
        '考点29：直营店与加盟店必须分别计算公司口径现金流，不得用门店EBITDA直接替代公司现金流。',
        '考点30：直接门店回收期超过42个月须专项审批，HZ02回收期41.0个月，未触发审批线。',
    ]),
    ('（七）直营与加盟（考点31—35）', [
        '考点31：直营模式由公司承担全部建店投资并取得全部营业收入。',
        '考点32：加盟模式收取15万元加盟费、6%持续授权费、18%供应链毛利，承担3%督导培训成本和每店每年8万元固定督导费。',
        '考点33：加盟模式公司初期支持投入10万元，加盟商承担85%建店投资。',
        '考点34：直营与加盟比较必须使用公司口径NPV和回收期，不得只比较门店EBITDA。',
        '考点35：首批直营6家、加盟2家，加盟比例25%，不超过30%上限，且每城首店均为直营。',
    ]),
    ('（八）首批清单与开业节奏（考点36—39）', [
        '考点36：首批8家为HZ01、HZ02、HZ04、HF01、HF02、HF03、NC01、NC03，城市分布3/3/2。',
        '考点37：杭州3家全部直营，合肥HF01和HF02直营、HF03加盟，南昌NC01直营、NC03加盟。',
        '考点38：计划开业日按租约起租日加2个月计算，形成2026年2月至8月的开业序列。',
        '考点39：开业节奏为2026Q1 3家、Q2 3家、Q3 2家，单季不超过项目组3家产能。',
    ]),
    ('（九）区域仓配（考点40—42）', [
        '考点40：每城比较自建城市仓、第三方仓配、区域中心仓三种方案，按5年成本现值选择最低方案。',
        '考点41：杭州和合肥选择自建城市仓，南昌选择第三方仓配，兼顾时效与成本。',
        f'考点42：仓配资本支出合计{check_map["仓配资本支出(万元)"]}万元，未重复计入门店建店投资。',
    ]),
    ('（十）投资总额（考点43—45）', [
        '考点43：首批总投资等于直营建店投资、加盟支持投入、仓配资本支出和5%预备费之和。',
        f'考点44：首批总投资{total_invest}万元，低于1500万元预算上限。',
        f'考点45：预算余量{round(1500 - total_invest, 1)}万元，可用于开业期营销或应急支出。',
    ]),
    ('（十一）敏感性与回报（考点46—47）', [
        '考点46：完成销售变动±10%与租金变动±5%的9组组合测算，逐店重算后汇总，不得只对总额乘系数。',
        f'考点47：9组情景组合NPV区间为{sens_min}万元至{sens_max}万元，基准情景IRR为{port_irr}%。',
    ]),
    ('（十二）风险与退出（考点48—49）', [
        '考点48：列出销售未达预测、租金上涨、竞品进入、外卖佣金上调、证照食安、加盟商违约、仓配不足、开业延期、投资超预算、商圈客流下滑共10项风险。',
        '考点49：销售连续2季低于基准15%，或租金占收入连续2季超过15%，触发关店、转址或解约评估。',
    ]),
    ('（十三）勾稽与核心结论（考点50）', [
        '考点50：完成24项勾稽，覆盖点位数、城市分布、直营加盟、投资、收入、EBITDA、NPV、IRR、敏感性和预算，全部闭合后方可出具一句话核心结论。',
    ]),
]
for title, items in groups:
    H(doc, title, 2)
    for t in items:
        P(doc, t, indent=False)

H(doc, '十、答题思路')
P(doc, '第一步是读取城市、候选点位、租约和销售预测，锁定18个候选点位及三城各6个的分布，先不做计算。')
P(doc, '第二步是资格筛查，剔除餐饮条件不符和租期不足5年的点位，保留评分和不符合说明。')
P(doc, '第三步是按七因子评分，所有指标在18个点位间做线性归一化，竞品数反向处理，得到总分和城市内排序。')
P(doc, '第四步是按每城3/3/2取前序点位，形成首批8家，并校验预算和首店直营约束。')
P(doc, '第五步是预测销售，拆分堂食与外卖，按爬坡系数计算首年收入，按3%计算稳定期增长。')
P(doc, '第六步是测算建店投资，逐项计算装修、设备、家具、POS、物料、证照、设计、押金和流动资金。')
P(doc, '第七步是编制五年P&L，计算EBITDA、折旧、所得税、税后现金流、回收期、NPV和IRR。')
P(doc, '第八步是做直营与加盟对比，按公司口径分别计算NPV和回收期，确定6家直营、2家加盟。')
P(doc, '第九步是完成仓配比选、投资总额、9组敏感性、10项风险、退出机制和24项勾稽，最后给出一句话核心结论。')

H(doc, '十一、标准答案')
P(doc, '标准答案共8个包，首表为投资决策结论，逐包覆盖点位评分、单店P&L、直营加盟、开业节奏、仓配投资、敏感性、风险退出和勾稽自检。', indent=False)
P(doc, concl['核心结论'], bold=True)
for k in ['支撑一', '支撑二', '支撑三', '支撑四']:
    if k in concl:
        P(doc, concl[k])
P(doc, f'关键结果：首批8店首年收入{first_rev}万元、首年EBITDA {first_ebitda}万元、5年累计EBITDA {five_ebitda}万元；组合5年NPV {port_npv}万元、IRR {port_irr}%、回收期{payback}个月；9组敏感性组合NPV区间{sens_min}万元至{sens_max}万元；24项勾稽全部闭合，不一致项为0。')

doc.core_properties.author = '投资委员会'
doc.save(f'{OUT}/连锁选址投资决策_v1_题目.docx')
scrub(f'{OUT}/连锁选址投资决策_v1_题目.docx', WORD_APP, '投资委员会', '2025-10-18T02:00:00Z', '2025-10-20T01:00:00Z', (2025, 10, 20, 9, 0, 0))

print('done')
print('total_invest', total_invest, 'npv', port_npv, 'irr', port_irr, 'payback', payback)
print('warehouse', wh_text)
print('packs', len(INPUT_PACKS), len(ANSWER_PACKS))
