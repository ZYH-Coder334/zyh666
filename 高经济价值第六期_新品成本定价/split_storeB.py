#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""新品成本核算与双渠道定价 v1 附件拆分与打包：8输入 + 8答案，并生成题目 docx"""
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

BASE = '/workspace/高经济价值第六期_新品成本定价'
SRC_Q = f'{BASE}/新品成本定价_v1_题目.xlsx'
SRC_A = f'{BASE}/新品成本定价_v1_标准答案.xlsx'
SPLIT = '/tmp/storeB_split'
OUT = f'{BASE}/新品成本定价_v1_附件'
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
    ('01新品清单与BOM配方', ['新品清单', 'BOM配方']),
    ('02原料采购价目表', ['原料采购价目表']),
    ('03原料涨价记录', ['原料涨价记录']),
    ('04包装与外卖规则', ['包装成本', '外卖佣金与活动规则']),
    ('05价格带与弹性', ['价格带', '价格弹性']),
    ('06门店损耗与工序', ['门店操作损耗', '中央厨房工序']),
    ('07财务确认规则', ['财务确认规则']),
    ('08管理层要求', ['管理层要求']),
]

ANSWER_PACKS = [
    ('01标准成本卡', ['SKU标准成本卡', 'BOM成本明细']),
    ('02双渠道毛利率', ['双渠道毛利率', '毛利率红线']),
    ('03定价与促销底线', ['定价与促销底线']),
    ('04上新与下架建议', ['上新与下架建议']),
    ('05涨价传导分析', ['原料涨价传导', '原料涨价明细']),
    ('06弹性与销量预测', ['价格弹性与销量预测']),
    ('07风险与退出机制', ['风险清单', '退出机制']),
    ('08勾稽与结论', ['错误要求驳回', '勾稽自检', '成本定价结论']),
]

IN_DT = (2026, 3, 1, 9, 0, 0)
AN_DT = (2026, 3, 10, 14, 0, 0)
split_workbook(SRC_Q, INPUT_PACKS, f'{OUT}/输入附件', 'in', '2026-02-28T01:00:00Z', '2026-03-01T01:00:00Z', IN_DT)
split_workbook(SRC_A, ANSWER_PACKS, f'{OUT}/标准答案', 'an', '2026-03-08T02:00:00Z', '2026-03-10T06:00:00Z', AN_DT)

# 读取答案指标
ans = openpyxl.load_workbook(SRC_A, data_only=True)
concl = {r[0]: r[1] for r in ans['成本定价结论'].iter_rows(values_only=True) if r[0]}
check_map = {r[0]: r[1] for r in ans['勾稽自检'].iter_rows(values_only=True) if r[0] != '勾稽项'}
decision_rows = list(ans['上新与下架建议'].iter_rows(values_only=True))[1:]
on_list = [r[0] for r in decision_rows if r[5] == '上新']
ref_list = [r[0] for r in decision_rows if r[5] == '改配方']
off_list = [r[0] for r in decision_rows if r[5] == '下架']
on_n = len(on_list)
ref_n = len(ref_list)
off_n = len(off_list)
check_n = len(check_map)
hall_pass = sum(1 for r in decision_rows if r[3] >= 65)
deliv_pass = sum(1 for r in decision_rows if r[4] >= 55)

# 各SKU毛利率
sku_margin = {r[0]: (r[3], r[4]) for r in decision_rows}

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
H(doc, '新品成本核算与双渠道定价决策', 0)

H(doc, '一、编号')
P(doc, 'G6-ZX-005（v1）。', indent=False)

H(doc, '二、任务背景')
P(doc, '麦淳餐饮管理有限公司是一家连锁烘焙、饮品与轻食品牌，2026年3月计划推出6款春季新品，涵盖烘焙、饮品、轻食品类。研发部已提交BOM配方，采购部提供原料采购价目表及三轮涨价记录，运营部提供价格带、竞品价格与历史价格弹性，财务部提供包装成本、外卖佣金规则和毛利率红线。现受财务总监与运营总监委托，以成本核算岗身份，完成6款新品的标准成本核算、双渠道毛利率测算、促销底线价反推、上新与下架建议、原料涨价传导分析和价格弹性销量预测。')
P(doc, '数据截止日为2026年2月28日，计划上市日为2026年3月1日。标准成本须按第3轮最新采购价计算，含门店操作损耗；堂食毛利率红线65%、外卖毛利率红线55%，双渠道均达标方可上新；外卖佣金自2026年3月起由18%上调至20%。所有结论须可回溯到附件逐条记录，不得只给汇总数。')

H(doc, '三、任务分解')
P(doc, '本项委托按八轮执行，每轮只完成一个任务，并只提交当轮对应交付包。上一轮自检通过后，方可进入下一轮；未完成或自检不通过，不得进入下一轮，也不得提前输出后续轮次结果。')
P(doc, '第1轮为标准成本卡。按BOM配方、第3轮采购价和出成率，计算6款新品的单份食材成本、门店损耗、堂食与外卖包装成本。本轮只输出01标准成本卡.rar。本轮自检要求6款SKU均有完整BOM成本明细，单位换算正确，出成率已计入。')
P(doc, '第2轮为双渠道毛利率。按堂食和外卖两种渠道，分别计算6款新品的毛利额和毛利率，对照红线判定是否达标。本轮只输出02双渠道毛利率.rar。本轮自检要求6款SKU堂食与外卖毛利率均有计算，红线判定准确。')
P(doc, '第3轮为定价与促销底线。按促销后毛利率不低于红线90%的要求，反推堂食与外卖促销底线价，计算最大促销空间。本轮只输出03定价与促销底线.rar。本轮自检要求6款SKU均有底线价，底线毛利率为58.5%和49.5%。')
P(doc, '第4轮为上新与下架建议。按双渠道达标情况，给出上新、改配方、下架三类结论及理由。本轮只输出04上新与下架建议.rar。本轮自检要求6款SKU均有结论，双达标上新、单达标改配方、双不达标下架。')
P(doc, '第5轮为涨价传导分析。对比第1轮与第3轮食材成本，计算各SKU成本上涨金额与涨幅，列出主要涨价原料。本轮只输出05涨价传导分析.rar。本轮自检要求6款SKU均有涨价传导，13种涨价原料明细完整。')
P(doc, '第6轮为弹性与销量预测。按品类价格弹性，测算建议售价下的基础销量和降价10%促销后的销量与收入。本轮只输出06弹性与销量预测.rar。本轮自检要求6款SKU均有弹性系数、促销价和预测销量。')
P(doc, '第7轮为风险与退出机制。列出10项风险、量化触发阈值、应对措施和退出动作。本轮只输出07风险与退出机制.rar。本轮自检要求10项风险均有阈值，6类退出机制均有回收与损失处理。')
P(doc, '第8轮为勾稽与结论。驳回管理层12项错误要求，完成勾稽自检并给出核心结论。本轮只输出08勾稽与结论.rar。本轮自检要求勾稽全部闭合、不一致项为0；未闭合不得出具结论。')

H(doc, '四、输入附件')
P(doc, '输入附件共8包。')
P(doc, '01新品清单与BOM配方.rar，包含新品清单、BOM配方。')
P(doc, '02原料采购价目表.rar，包含原料采购价目表。')
P(doc, '03原料涨价记录.rar，包含原料涨价记录。')
P(doc, '04包装与外卖规则.rar，包含包装成本、外卖佣金与活动规则。')
P(doc, '05价格带与弹性.rar，包含价格带、价格弹性。')
P(doc, '06门店损耗与工序.rar，包含门店操作损耗、中央厨房工序。')
P(doc, '07财务确认规则.rar，包含财务确认规则。')
P(doc, '08管理层要求.rar，包含管理层要求。')

H(doc, '五、制度与判断规则')
P(doc, '6款新品分为烘焙、饮品、轻食三个品类，各品类有独立的门店操作损耗率（烘焙4%、饮品2%、轻食5%）和价格弹性系数（烘焙-0.8、饮品-1.2、轻食-0.6）。')
P(doc, '单份食材成本等于Σ（单耗÷出成率×采购单价），单位换算规则为g转kg除以1000、ml转L除以1000、个按个计。出成率为原料加工后可用部分比例，单耗须除以出成率，不得按100%计算。')
P(doc, '标准成本按第3轮最新采购价计算，第3轮单价等于第1轮单价乘（1+第2轮涨幅）乘（1+第3轮涨幅）。未列入涨价记录的原料按采购价目表单价计算。')
P(doc, '门店操作损耗须计入食材成本，含损耗食材成本等于食材成本乘（1+品类损耗率）。中央厨房到门店的半成品结转按工序结转，不得重复计料。')
P(doc, '包装成本按品类与渠道区分：烘焙堂食0.30元、外卖0.80元；饮品堂食0.50元、外卖1.20元；轻食堂食0.10元、外卖1.50元。')
P(doc, '堂食毛利等于堂食售价减食材成本（含损耗）减堂食包装成本，堂食毛利率等于堂食毛利除以堂食售价。')
P(doc, '外卖毛利等于外卖售价减食材成本（含损耗）减外卖包装减外卖佣金减配送费加平台补贴减退款。外卖佣金按外卖售价的20%计提，配送费4%、退款2%、平台补贴3%，均按外卖售价计提。')
P(doc, '毛利率红线：堂食不低于65%、外卖不低于55%，双渠道均达标方可上新；单渠道不达标须优化配方或调整包装；双渠道均不达标下架。')
P(doc, '促销底线：促销后毛利率不得低于红线的90%，即堂食不低于58.5%、外卖不低于49.5%。促销底线价按底线毛利率反推，低于底线的促销须审批。')
P(doc, '价格弹性用于销量预测：促销降价10%时，销量变化率等于弹性系数乘（-10%）。基础日均销量为建议售价下的预测销量。')

H(doc, '六、交付物')
for t in ['01 标准成本卡.rar', '02 双渠道毛利率.rar', '03 定价与促销底线.rar',
          '04 上新与下架建议.rar', '05 涨价传导分析.rar', '06 弹性与销量预测.rar',
          '07 风险与退出机制.rar', '08 勾稽与结论.rar']:
    P(doc, t, indent=False)

H(doc, '七、约束条件')
constraints = [
    '1. 仅使用附件材料，不得引入外部数据，不得联网检索。',
    '2. 金额以元列示并保留两位小数，比例保留两位小数，不得混用单位。',
    '3. 逐SKU给出结论，不得抽样或估算；每条记录必须有明确状态，不得留空。',
    '4. 6款新品编号唯一，烘焙3款、饮品2款、轻食1款，不得跨品类重复计入。',
    '5. 标准成本按第3轮采购价计算，不得使用第1轮或第2轮价格。',
    '6. 单耗须除以出成率，出成率不得按100%简化计算。',
    '7. 单位换算：g转kg除以1000，ml转L除以1000，个按个计。',
    '8. 门店操作损耗须计入食材成本，烘焙4%、饮品2%、轻食5%。',
    '9. 包装成本按品类与渠道区分，不得忽略或统一取值。',
    '10. 外卖佣金按外卖售价的20%计提，不得按原18%或堂食售价计提。',
    '11. 外卖毛利须扣配送费4%、退款2%，加平台补贴3%，不得只扣佣金。',
    '12. 堂食毛利率红线65%、外卖55%，双渠道均达标方可上新。',
    '13. 促销后毛利率不得低于红线的90%，堂食58.5%、外卖49.5%。',
    '14. 促销底线价按底线毛利率反推，低于底线的促销须审批。',
    '15. 半成品结转按工序结转，不得重复计料或遗漏。',
    '16. 价格弹性按品类区分，烘焙-0.8、饮品-1.2、轻食-0.6，不得统一按-1.0。',
    '17. 双渠道均不达标须下架，不得直接降价促销。',
    '18. 必须列出10项风险和6类退出机制，每项风险须有量化触发阈值。',
    '19. 勾稽全部闭合后方可出具核心结论；任一项不闭合视为未完成。',
    '20. 必须按八轮顺序逐轮交付，每轮只完成一个任务，不得一次性输出全部结果。',
    '21. 每轮结束必须提交本轮自检记录，列明检查项、结果、证据和未通过原因；上一轮未通过不得进入下一轮。',
    '22. 每轮只提交当轮对应交付包，不得跨轮合并或提前输出；任何一轮无法完成时必须停止并说明缺口；八轮全部通过方视为完成。',
    '23. 管理层12项错误要求须逐条驳回并说明理由，不得采纳。',
    '24. 所有定价与促销建议必须有明确触发条件与审批流程，不得只给方向性意见。',
]
for t in constraints:
    P(doc, t, indent=False)

H(doc, '八、完成标准')
P(doc, f'八轮任务逐轮完成且逐轮自检通过。6款新品中上{on_n}款（{"、".join(on_list)}）、改配方{ref_n}款（{"、".join(ref_list)}）、下架{off_n}款（{"、".join(off_list)}）；堂食达标{hall_pass}款、外卖达标{deliv_pass}款；标准成本按第3轮采购价计算，含门店损耗；促销底线价已反推；{check_n}项勾稽全部闭合、不一致项为0。任一轮未完成或自检不通过，不得进入下一轮，整体不得视为完成。')

H(doc, '九、考点')
P(doc, '考点按数据治理、BOM与成本、出成率与损耗、包装成本、外卖佣金、双渠道毛利、促销底线、上新判定、涨价传导、价格弹性、风险退出、勾稽结论十二组展开，共50项。', indent=False)
groups = [
    ('（一）数据治理与新品清单（考点1—5）', [
        '考点1：新品共6款，烘焙3款（N01、N02、N05）、饮品2款（N03、N06）、轻食1款（N04），编号唯一。',
        '考点2：BOM配方共35行，覆盖6款SKU的全部原料单耗与出成率。',
        '考点3：原料采购价目表共26种原料，含规格、采购单价、单位、供应商。',
        '考点4：原料涨价记录共13种，含第1轮单价、第2轮涨幅、第3轮涨幅。',
        '考点5：计划上市日为2026年3月1日，数据截止日为2026年2月28日。',
    ]),
    ('（二）BOM与标准成本（考点6—10）', [
        '考点6：单份食材成本等于Σ（单耗÷出成率×采购单价）。',
        '考点7：单位换算规则为g转kg除以1000、ml转L除以1000、个按个计。',
        '考点8：标准成本按第3轮采购价计算，第3轮单价等于第1轮乘（1+第2轮涨幅）乘（1+第3轮涨幅）。',
        '考点9：未列入涨价记录的原料按采购价目表单价计算。',
        '考点10：BOM成本明细须逐原料列出单耗、出成率、第3轮单价和成本。',
    ]),
    ('（三）出成率与门店损耗（考点11—15）', [
        '考点11：出成率为原料加工后可用部分比例，单耗须除以出成率。',
        '考点12：红豆出成率0.85、芋头0.80、鸡胸肉0.85、牛油果0.75、芒果0.60、西柚0.50。',
        '考点13：门店操作损耗须计入食材成本，烘焙4%、饮品2%、轻食5%。',
        '考点14：含损耗食材成本等于食材成本乘（1+品类损耗率）。',
        '考点15：中央厨房到门店的半成品结转按工序结转，不得重复计料。',
    ]),
    ('（四）包装成本（考点16—18）', [
        '考点16：包装成本按品类与渠道区分，共6类。',
        '考点17：烘焙堂食0.30元、外卖0.80元；饮品堂食0.50元、外卖1.20元；轻食堂食0.10元、外卖1.50元。',
        '考点18：堂食毛利减堂食包装，外卖毛利减外卖包装，不得混用。',
    ]),
    ('（五）外卖佣金与补贴（考点19—23）', [
        '考点19：外卖佣金自2026年3月起由18%上调至20%，按外卖售价计提。',
        '考点20：配送费按外卖售价的4%由门店承担。',
        '考点21：平台补贴按外卖售价的3%计入外卖GMV。',
        '考点22：退款率按外卖售价的2%冲减收入。',
        '考点23：外卖毛利等于售价减食材减外卖包装减佣金减配送加补贴减退款。',
    ]),
    ('（六）双渠道毛利率（考点24—28）', [
        '考点24：堂食毛利等于售价减食材（含损耗）减堂食包装。',
        '考点25：堂食毛利率等于堂食毛利除以堂食售价。',
        '考点26：外卖毛利率等于外卖毛利除以外卖售价。',
        '考点27：堂食毛利率红线65%、外卖55%，双渠道均达标方可上新。',
        f'考点28：堂食达标{hall_pass}款、外卖达标{deliv_pass}款。',
    ]),
    ('（七）促销底线（考点29—32）', [
        '考点29：促销后毛利率不得低于红线的90%，堂食58.5%、外卖49.5%。',
        '考点30：促销底线价按底线毛利率反推，堂食底线价等于（食材+堂食包装）除以（1-58.5%）。',
        '考点31：外卖底线价等于（食材+外卖包装）除以（1-佣金-配送+补贴-退款-49.5%）。',
        '考点32：最大促销空间等于建议售价减促销底线价。',
    ]),
    ('（八）上新判定（考点33—36）', [
        '考点33：双渠道均达标判定为上新。',
        '考点34：单渠道不达标判定为改配方，须优化配方或调整包装。',
        '考点35：双渠道均不达标判定为下架，不得直接降价促销。',
        f'考点36：上新{on_n}款（{"、".join(on_list)}）、改配方{ref_n}款（{"、".join(ref_list)}）、下架{off_n}款（{"、".join(off_list)}）。',
    ]),
    ('（九）涨价传导（考点37—40）', [
        '考点37：涨价传导等于第3轮食材成本减第1轮食材成本。',
        '考点38：涨幅等于成本上涨除以第1轮成本。',
        '考点39：主要涨价原料为咸蛋黄、牛油果、芒果、抹茶粉，累计涨幅居前。',
        '考点40：13种涨价原料须逐种列出第1轮单价、第2轮涨幅、第3轮涨幅、第3轮单价和累计涨价。',
    ]),
    ('（十）价格弹性（考点41—44）', [
        '考点41：价格弹性按品类区分，烘焙-0.8、饮品-1.2、轻食-0.6。',
        '考点42：促销降价10%时，销量变化率等于弹性系数乘（-10%）。',
        '考点43：促销销量等于基础日均销量乘（1+销量变化率）。',
        '考点44：促销收入等于促销价乘促销销量，须与基础收入对比。',
    ]),
    ('（十一）风险与退出（考点45—47）', [
        '考点45：列出原料持续涨价、外卖佣金上调、出成率低于标准、价格弹性误判、包装成本上涨、竞品降价、食安与品质、促销过度、半成品结转错误、门店操作损耗超标共10项风险。',
        '考点46：每项风险须有量化触发阈值、影响、应对措施和退出动作。',
        '考点47：退出机制含下架、停售整改、临时下架、收缩渠道、停止促销、调整定价共6类，均有回收与损失处理。',
    ]),
    ('（十二）勾稽与结论（考点48—50）', [
        f'考点48：{check_n}项勾稽覆盖新品数、BOM行数、原料数、涨价原料数、上新/改配方/下架数、达标数、红线、佣金率、计算公式、取价轮次、出成率、损耗、促销底线，全部闭合。',
        '考点49：管理层12项错误要求须逐条驳回，包括佣金按18%、按第1轮价、出成率100%、忽略包装、只扣佣金、只看堂食、促销不影响毛利、按竞品最低价、忽略损耗、重复计料、弹性统一-1.0、不达标直接降价。',
        f'考点50：核心结论为6款新品中{on_n}款上新、{ref_n}款改配方、{off_n}款下架，标准成本按第3轮价含损耗，堂食红线65%外卖55%，外卖佣金20%。',
    ]),
]
for title, items in groups:
    H(doc, title, 2)
    for t in items:
        P(doc, t, indent=False)

H(doc, '十、答题思路')
P(doc, '第一步是读取新品清单、BOM配方、原料采购价目表和涨价记录，锁定6款SKU及BOM明细。')
P(doc, '第二步是按第3轮采购价、出成率和单位换算，逐原料计算BOM成本，汇总为单份食材成本。')
P(doc, '第三步是加门店操作损耗，得到含损耗食材成本，并匹配堂食与外卖包装成本。')
P(doc, '第四步是按堂食和外卖两种渠道，分别计算毛利额和毛利率，对照65%和55%红线判定达标。')
P(doc, '第五步是按促销底线毛利率58.5%和49.5%，反推堂食与外卖促销底线价，计算最大促销空间。')
P(doc, '第六步是按双渠道达标情况，给出上新、改配方、下架结论，逐SKU说明理由。')
P(doc, '第七步是对比第1轮与第3轮食材成本，计算涨价传导金额与涨幅，列出主要涨价原料。')
P(doc, '第八步是按品类价格弹性，测算降价10%促销后的销量与收入变化。')
P(doc, '第九步是完成10项风险、6类退出机制、12项错误要求驳回和勾稽自检，最后给出核心结论。')

H(doc, '十一、标准答案')
P(doc, '标准答案共8个包，首表为成本定价结论，逐包覆盖标准成本卡、双渠道毛利率、定价与促销底线、上新与下架建议、涨价传导、弹性预测、风险退出和勾稽结论。', indent=False)
P(doc, concl['核心结论'], bold=True)
for k in ['支撑一', '支撑二', '支撑三', '支撑四']:
    if k in concl:
        P(doc, concl[k])
P(doc, f'关键结果：6款新品中上{on_n}款（{"、".join(on_list)}）、改配方{ref_n}款（{"、".join(ref_list)}）、下架{off_n}款（{"、".join(off_list)}）；堂食达标{hall_pass}款、外卖达标{deliv_pass}款；{check_n}项勾稽全部闭合，不一致项为0。')

doc.core_properties.author = '运营分析部'
doc.save(f'{OUT}/新品成本定价_v1_题目.docx')
scrub(f'{OUT}/新品成本定价_v1_题目.docx', WORD_APP, '运营分析部', '2026-03-01T02:00:00Z', '2026-03-06T01:00:00Z', (2026, 3, 6, 10, 0, 0))

print('done')
print('on', on_list, 'reform', ref_list, 'off', off_list)
print('checks', check_n, 'hall_pass', hall_pass, 'deliv_pass', deliv_pass)
print('packs', len(INPUT_PACKS), len(ANSWER_PACKS))
