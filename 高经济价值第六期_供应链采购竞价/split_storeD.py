#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""供应链采购竞价与供应商评级 v1 —— 拆包+题目docx生成器"""
import os
import shutil
import zipfile
import tempfile
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from docx import Document
from docx.shared import Pt, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn

BASE = '/workspace/高经济价值第六期_供应链采购竞价'
QFILE = f'{BASE}/供应链采购竞价_v1_题目.xlsx'
AFILE = f'{BASE}/供应链采购竞价_v1_标准答案.xlsx'
OUT = f'{BASE}/供应链采购竞价_v1_附件'
IN_DT = (2026, 2, 5, 9, 0, 0)
AN_DT = (2026, 2, 15, 14, 0, 0)
DOC_DT = (2026, 2, 10, 10, 0, 0)

INPUT_PACKS = [
    ('01原料年度需求', ['原料年度需求']),
    ('02新供应商投标', ['新供应商投标']),
    ('03现有供应商年度数据', ['现有供应商年度数据']),
    ('04评分权重', ['评分权重']),
    ('05TCO参数', ['TCO参数']),
    ('06ABC分级与份额规则', ['ABC分级标准', '份额分配规则']),
    ('07财务确认规则', ['财务确认规则']),
    ('08管理层要求', ['管理层要求']),
]

ANSWER_PACKS = [
    ('01综合评分表', ['综合评分表']),
    ('02TCO总成本', ['TCO总成本']),
    ('03中标结果', ['中标结果']),
    ('04推荐中标方案', ['推荐中标方案']),
    ('05现有供应商ABC分级与份额', ['现有供应商ABC分级', '份额分配与调整建议']),
    ('06风险与退出机制', ['风险清单', '退出机制']),
    ('07错误要求驳回', ['错误要求驳回']),
    ('08勾稽与结论', ['勾稽自检', '供应链采购竞价结论']),
]

HDR_FILL = PatternFill('solid', fgColor='2F5597')
HDR_FONT = Font(name='宋体', bold=True, color='FFFFFF', size=10)
BODY = Font(name='宋体', size=9)
THIN = Side(style='thin', color='CCCCCC')
BORDER = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
WRAP = Alignment(wrap_text=True, vertical='top')
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)


def copy_sheet(src_ws, dst_wb, title):
    dst_ws = dst_wb.create_sheet(title)
    for row in src_ws.iter_rows():
        for cell in row:
            new_cell = dst_ws.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                new_cell.font = Font(
                    name=cell.font.name or '宋体',
                    size=cell.font.size or 10,
                    bold=cell.font.bold,
                    color=cell.font.color
                )
                new_cell.fill = PatternFill(
                    fill_type=cell.fill.fill_type or 'solid',
                    fgColor=cell.fill.fgColor.rgb if cell.fill.fgColor else 'FFFFFF'
                )
                new_cell.alignment = Alignment(
                    horizontal=cell.alignment.horizontal,
                    vertical=cell.alignment.vertical,
                    wrap_text=cell.alignment.wrap_text
                )
                new_cell.border = BORDER
    for col in src_ws.column_dimensions:
        dst_ws.column_dimensions[col].width = src_ws.column_dimensions[col].width
    for rng in src_ws.merged_cells.ranges:
        dst_ws.merge_cells(str(rng))
    dst_ws.freeze_panes = src_ws.freeze_panes
    dst_ws.sheet_view.showGridLines = False
    return dst_ws


def split_workbook(src, dst_dir, packs, date_time, is_answer=False):
    wb_src = openpyxl.load_workbook(src, data_only=not is_answer)
    for pack_name, sheets in packs:
        wb_dst = openpyxl.Workbook()
        wb_dst.remove(wb_dst.active)
        for sh in sheets:
            if sh in wb_src.sheetnames:
                copy_sheet(wb_src[sh], wb_dst, sh)
        tmp_xlsx = f'{dst_dir}/{pack_name}.xlsx'
        wb_dst.save(tmp_xlsx)
        make_rar(tmp_xlsx, f'{dst_dir}/{pack_name}.rar', date_time)
        os.remove(tmp_xlsx)


def make_rar(xlsx_path, rar_path, date_time):
    with zipfile.ZipFile(rar_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        info = zipfile.ZipInfo(os.path.basename(xlsx_path), date_time=date_time)
        with open(xlsx_path, 'rb') as f:
            zf.writestr(info, f.read())
    ts = datetime(*date_time).timestamp()
    os.utime(rar_path, (ts, ts))


def r2(x): return round(x + 0.0, 2)


def scrub(rar_path, date_time, is_word=False):
    temp = tempfile.mkdtemp()
    with zipfile.ZipFile(rar_path, 'r') as zf:
        zf.extractall(temp)
    xml_dir = os.path.join(temp, 'docProps')
    os.makedirs(xml_dir, exist_ok=True)
    app_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>Microsoft Excel</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <Company>麦淳餐饮管理有限公司</Company>
  <LinksUpToDate>false</LinksUpToDate>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>16.0300</AppVersion>
</Properties>''' if not is_word else '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>Microsoft Office Word</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <Company>麦淳餐饮管理有限公司</Company>
  <LinksUpToDate>false</LinksUpToDate>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>16.0000</AppVersion>
</Properties>'''
    with open(os.path.join(xml_dir, 'app.xml'), 'w', encoding='utf-8') as f:
        f.write(app_xml)
    dt_str = datetime(*date_time).strftime('%Y-%m-%dT%H:%M:%SZ')
    core_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>供应链采购竞价_v1</dc:title>
  <dc:subject>高经济价值第六期</dc:subject>
  <dc:creator>运营分析部</dc:creator>
  <cp:lastModifiedBy>运营分析部</cp:lastModifiedBy>
  <cp:revision>1</cp:revision>
  <dcterms:created xsi:type="dcterms:W3CDTF">{dt_str}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{dt_str}</dcterms:modified>
</cp:coreProperties>'''
    with open(os.path.join(xml_dir, 'core.xml'), 'w', encoding='utf-8') as f:
        f.write(core_xml)
    with zipfile.ZipFile(rar_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(temp):
            for file in files:
                full = os.path.join(root, file)
                arc = os.path.relpath(full, temp)
                info = zipfile.ZipInfo(arc, date_time=date_time)
                with open(full, 'rb') as f:
                    zf.writestr(info, f.read())
    ts = datetime(*date_time).timestamp()
    os.utime(rar_path, (ts, ts))
    shutil.rmtree(temp)


def set_run_font(run, name='宋体', size=10.5, bold=False):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    run._element.rPr.rFonts.set(qn('w:eastAsia'), name)


def H(doc, text, level=1):
    p = doc.add_heading(level=level)
    r = p.add_run(text)
    set_run_font(r, '黑体', 14 if level == 1 else 12, bold=True)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)


def P(doc, text, bold=False, indent=True):
    p = doc.add_paragraph()
    r = p.add_run(text)
    set_run_font(r, '宋体', 10.5, bold)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p.paragraph_format.line_spacing = Pt(20)
    if indent:
        p.paragraph_format.first_line_indent = Inches(0.3)
    p.paragraph_format.space_after = Pt(3)


def main():
    if os.path.exists(OUT):
        shutil.rmtree(OUT)
    os.makedirs(f'{OUT}/输入附件')
    os.makedirs(f'{OUT}/标准答案')

    split_workbook(QFILE, f'{OUT}/输入附件', INPUT_PACKS, IN_DT, is_answer=False)
    split_workbook(AFILE, f'{OUT}/标准答案', ANSWER_PACKS, AN_DT, is_answer=True)

    for d in [f'{OUT}/输入附件', f'{OUT}/标准答案']:
        for f in os.listdir(d):
            scrub(f'{d}/{f}', IN_DT if '输入' in d else AN_DT, is_word=False)

    wb_ans = openpyxl.load_workbook(AFILE, data_only=True)
    ws_concl = wb_ans['供应链采购竞价结论']
    concl = {ws_concl.cell(row=r, column=1).value: ws_concl.cell(row=r, column=2).value for r in range(2, ws_concl.max_row + 1)}

    ws_check = wb_ans['勾稽自检']
    check_map = {ws_check.cell(row=r, column=1).value: ws_check.cell(row=r, column=5).value for r in range(2, ws_check.max_row + 1)}
    check_n = ws_check.max_row - 1
    inconsistency = sum(1 for r in range(2, ws_check.max_row + 1) if ws_check.cell(row=r, column=5).value != '一致')

    ws_score = wb_ans['综合评分表']
    score_winners = {}
    for r in range(2, ws_score.max_row + 1):
        if ws_score.cell(row=r, column=13).value == '中标':
            mid = ws_score.cell(row=r, column=1).value
            sid = ws_score.cell(row=r, column=3).value
            score_winners[mid] = sid

    doc = Document()
    doc.styles['Normal'].font.name = '宋体'
    doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    H(doc, '供应链采购竞价与供应商评级')

    H(doc, '一、编号', 2)
    P(doc, 'G6-ZX-007（v1）。', indent=False)
    P(doc, '【职业定位】二级行业：餐饮连锁供应链管理 / 采购与供应管理；L1职业簇：供应链管理、财务与审计、运营管理；L2职业：采购经理、供应链分析师、成本会计师、质量审计专员、运营分析专员。', indent=False)

    H(doc, '二、任务背景', 2)
    P(doc, '麦淳餐饮管理有限公司旗下烘焙、饮品、轻食三类门店对高筋面粉、动物黄油、牛油果、抹茶粉、红豆、芋头、鸡胸肉、鲜牛奶等8种核心原料年采购额超过300万元。2026年初，公司计划对8种原料进行新一轮年度招标，收到3家新供应商投标；同时需要对现有6家供应商进行年度ABC分级，确定下一年份额分配与淘汰名单。')
    P(doc, '本次采购决策存在三个现实张力：一是采购部倾向于选择报价最低的金谷粮油以降低食材成本；二是财务部与质量部担心低价供应商的质检合格率、交货及时率不足会带来食安与断供风险；三是运营部希望保留账期长、质检高的高端供应商作为牛奶、黄油等高质原料的备选。')
    P(doc, '数据截止日为2026年2月28日，分析基准日为2026年3月10日。所有结论须同时计算综合评分与TCO总成本，并给出推荐中标方案，不得只给汇总数。')

    H(doc, '三、任务分解', 2)
    P(doc, '本项委托按八轮执行，每轮只完成一个任务，并只提交当轮对应交付包。上一轮自检通过后，方可进入下一轮；未完成或自检不通过，不得进入下一轮，也不得提前输出后续轮次结果。')
    P(doc, '第1轮为综合评分。对3家新供应商在8种原料上的投标按单价60%、账期15%、交货及时率15%、质检合格率5%、MOQ灵活性5%进行阶梯式加权评分，输出每类原料的综合评分与排名。本轮只输出01综合评分表.rar。本轮自检要求24条投标记录均有综合评分，权重合计100%。')
    P(doc, '第2轮为TCO总成本。按采购金额、资金占用成本、仓储成本、质检损失、断货风险成本计算每类原料每个供应商的TCO。本轮只输出02TCO总成本.rar。本轮自检要求TCO口径与财务确认规则一致。')
    P(doc, '第3轮为中标结果。输出综合评分最高者与TCO最低者对比。本轮只输出03中标结果.rar。本轮自检要求8种原料均有推荐中标供应商。')
    P(doc, '第4轮为推荐中标方案。结合综合评分与TCO差异，给出8种原料的推荐中标供应商及理由。本轮只输出04推荐中标方案.rar。本轮自检要求推荐理由具体、可追溯。')
    P(doc, '第5轮为现有供应商ABC分级与份额。按年度采购额对6家现有供应商排序，按累计占比0-70%、70-90%、90-100%划分A/B/C级，并按A级60%、B级30%、C级10%分配下一年份额。本轮只输出05现有供应商ABC分级与份额.rar。')
    P(doc, '第6轮为风险与退出机制。列出10项风险、量化触发阈值、应对措施和退出动作；列出6类退出机制及回收处理。本轮只输出06风险与退出机制.rar。')
    P(doc, '第7轮为错误要求驳回。逐条驳回管理层12项错误要求并说明理由。本轮只输出07错误要求驳回.rar。')
    P(doc, '第8轮为勾稽与结论。完成勾稽自检并给出核心结论。本轮只输出08勾稽与结论.rar。本轮自检要求所有勾稽结论为一致，不一致项为0。')

    H(doc, '四、输入附件', 2)
    P(doc, '输入附件共8包。')
    P(doc, '01原料年度需求.rar，包含原料年度需求。')
    P(doc, '02新供应商投标.rar，包含新供应商投标。')
    P(doc, '03现有供应商年度数据.rar，包含现有供应商年度数据。')
    P(doc, '04评分权重.rar，包含评分权重。')
    P(doc, '05TCO参数.rar，包含TCO参数。')
    P(doc, '06ABC分级与份额规则.rar，包含ABC分级标准、份额分配规则。')
    P(doc, '07财务确认规则.rar，包含财务确认规则。')
    P(doc, '08管理层要求.rar，包含管理层要求。')

    H(doc, '五、制度与判断规则', 2)
    P(doc, '新供应商准入评分采用五维度加权：单价60%、账期15%、交货及时率15%、质检合格率5%、MOQ灵活性5%，满分100分。单价得分按最低报价除以当前报价折算百分制；账期采用阶梯得分，≥60天100分、≥45天80分、≥30天60分、<30天40分；交货及时率采用阶梯得分，≥98%得100分、≥95%得85分、≥93%得75分、≥90%得65分、<90%得50分；质检合格率采用阶梯得分，≥99%得100分、≥98.5%得90分、≥98%得85分、≥97%得75分、<97%得60分；MOQ灵活性按同原料组内最大最小MOQ归一化得分。')
    P(doc, 'TCO总成本=采购金额+资金占用成本+仓储成本+质检损失+断货风险成本。采购金额=报价×年度需求量；资金占用成本=采购金额×账期÷365×8%；仓储成本=报价×安全库存量×3%，其中安全库存量=年度需求量÷365×安全库存天数；质检损失=采购金额×(1-质检合格率)×50%；断货风险成本=采购金额×(1-交货及时率)×10%。')
    P(doc, '推荐中标方案以综合评分最高为原则，同时列示TCO最低者作为对比；若推荐中标供应商与TCO最低供应商不一致，须量化说明TCO差异及质量、账期、供应稳定性方面的权衡理由。')
    P(doc, '现有供应商ABC分级按年度采购额从大到小排序，累计占比0-70%为A级、70-90%为B级、90-100%为C级。A级供应商下一年份额60%，B级30%，C级10%。连续2年C级或质检合格率<90%或交货及时率<80%的供应商应退出或整改。')
    P(doc, '续约谈判须提前3个月启动；关键原料须保留至少2家合格供应商，单一供应商占比不得超过70%。')

    H(doc, '六、交付物', 2)
    P(doc, '01 综合评分表.rar')
    P(doc, '02 TCO总成本.rar')
    P(doc, '03 中标结果.rar')
    P(doc, '04 推荐中标方案.rar')
    P(doc, '05 现有供应商ABC分级与份额.rar')
    P(doc, '06 风险与退出机制.rar')
    P(doc, '07 错误要求驳回.rar')
    P(doc, '08 勾稽与结论.rar')

    H(doc, '七、约束条件', 2)
    P(doc, '1. 仅使用附件材料，不得引入外部数据，不得联网检索。')
    P(doc, '2. 金额以元列示并保留两位小数，比例保留两位小数，不得混用单位。')
    P(doc, '3. 逐原料、逐供应商给出评分与TCO，不得抽样或估算；每条记录必须有明确结果，不得留空。')
    P(doc, '4. 3家新供应商编号唯一，6家现有供应商编号唯一，8种原料编号唯一，不得重复或遗漏。')
    P(doc, '5. 评分权重合计必须等于100%，五维度缺一不可。')
    P(doc, '6. 单价、账期、交货、质检、MOQ须采用阶梯式评分规则，不得改为线性连续评分。')
    P(doc, '7. TCO必须包含采购金额、资金占用、仓储、质检损失、断货风险五项，缺一不可。')
    P(doc, '8. 安全库存量=年度需求量÷365×安全库存天数，不得直接采用MOQ或其他口径。')
    P(doc, '9. 推荐中标方案必须同时列示综合评分最高者和TCO最低者，并给出明确理由。')
    P(doc, '10. ABC分级须按年度采购额累计占比划分，不得按主观印象。')
    P(doc, '11. 份额分配按A级60%、B级30%、C级10%执行，不得随意调整。')
    P(doc, '12. 关键原料单一供应商占比不得超过70%，须保留备选。')
    P(doc, '13. 续约谈判须提前3个月启动，不得到期后处理。')
    P(doc, '14. 退出机制须有回收与损失处理方案，不得只给方向性意见。')
    P(doc, '15. 质检扣款比例50%、断货损失比例10%为合同约定，不得调整。')
    P(doc, '16. 资金成本率8%、仓储费率3%为年化口径，须按实际占用天数折算。')
    P(doc, '17. 必须列出10项风险和6类退出机制，每项风险须有量化触发阈值。')
    P(doc, '18. 勾稽全部闭合后方可出具核心结论；任一项不闭合视为未完成。')
    P(doc, '19. 必须按八轮顺序逐轮交付，每轮只完成一个任务，不得一次性输出全部结果。')
    P(doc, '20. 每轮结束必须提交本轮自检记录，列明检查项、结果、证据和未通过原因；上一轮未通过不得进入下一轮。')
    P(doc, '21. 每轮只提交当轮对应交付包，不得跨轮合并或提前输出；任何一轮无法完成时必须停止并说明缺口；八轮全部通过方视为完成。')
    P(doc, '22. 管理层12项错误要求须逐条驳回并说明理由，不得采纳。')
    P(doc, '23. 现有供应商分级与份额分配建议必须有明确触发条件和时间节点，不得只给方向性意见。')
    P(doc, '24. 综合评分、TCO、分级结果须逐条计算并汇总，不得只给总数。')

    H(doc, '八、完成标准', 2)
    core_text = concl.get('核心结论', '')
    if '新招标年度采购额' in core_text and '推荐方案TCO合计' in core_text:
        purchase_amt = r2(float(core_text.split('新招标年度采购额')[1].split('元')[0]))
        tco_amt = r2(float(core_text.split('推荐方案TCO合计')[1].split('元')[0]))
    else:
        purchase_amt = '待计算'
        tco_amt = '待计算'
    P(doc, f'八轮任务逐轮完成且逐轮自检通过。8种原料推荐中标方案为：高筋面粉-S1绿源食品、动物黄油-S1绿源食品、牛油果-S2金谷粮油、抹茶粉-S1绿源食品、红豆-S2金谷粮油、芋头-S2金谷粮油、鸡胸肉-S2金谷粮油、鲜牛奶-S3优品农贸；6家现有供应商A级2家、B级2家、C级2家；推荐方案年度采购额约{purchase_amt}元、TCO合计约{tco_amt}元；{check_n}项勾稽全部一致、不一致项为0。任一轮未完成或自检不通过，不得进入下一轮，整体不得视为完成。')

    H(doc, '九、考点', 2)
    P(doc, '以下考点按评分模型、TCO测算、中标决策、供应商分级、风险退出、错误驳回、勾稽结论七组展开，共50项，用于逐条评定答题底稿的完整性与正确性。', indent=False)

    groups = [
        ('（一）评分模型（考点1—10）', [
            '考点1：底稿中应当列示8种核心原料的年度需求量、单位、安全库存天数。',
            '考点2：底稿中应当列示3家新供应商对8种原料的24条投标记录。',
            '考点3：底稿中应当列明五维度评分权重为单价60%、账期15%、交货及时率15%、质检合格率5%、MOQ灵活性5%。',
            '考点4：底稿中应当按阶梯规则计算账期得分。',
            '考点5：底稿中应当按阶梯规则计算交货及时率得分。',
            '考点6：底稿中应当按阶梯规则计算质检合格率得分。',
            '考点7：底稿中应当按最低报价折算单价得分。',
            '考点8：底稿中应当按组内最大最小MOQ计算MOQ灵活性得分。',
            '考点9：底稿中应当对24条投标记录给出综合评分并排序。',
            '考点10：底稿中应当明确每类原料综合评分最高者。',
        ]),
        ('（二）TCO测算（考点11—18）', [
            '考点11：底稿中应当按采购金额=报价×年度需求量计算。',
            '考点12：底稿中应当按资金占用成本=采购金额×账期÷365×8%计算。',
            '考点13：底稿中应当按安全库存量=年度需求量÷365×安全库存天数计算。',
            '考点14：底稿中应当按仓储成本=报价×安全库存量×3%计算。',
            '考点15：底稿中应当按质检损失=采购金额×(1-质检合格率)×50%计算。',
            '考点16：底稿中应当按断货风险成本=采购金额×(1-交货及时率)×10%计算。',
            '考点17：底稿中应当汇总TCO总成本并排序。',
            '考点18：底稿中应当列示每类原料TCO最低者。',
        ]),
        ('（三）中标决策（考点19—25）', [
            '考点19：底稿中应当给出8种原料的推荐中标供应商。',
            '考点20：底稿中应当同时列示推荐中标供应商和TCO最低供应商。',
            '考点21：底稿中应当对推荐中标与TCO最低不一致的原料量化TCO差异。',
            '考点22：底稿中应当说明推荐中标方案中质量、账期、供应稳定性的权衡理由。',
            '考点23：底稿中应当计算推荐方案年度采购额。',
            '考点24：底稿中应当计算推荐方案TCO合计。',
            '考点25：底稿中应当避免单一供应商包揽全部8种原料。',
        ]),
        ('（四）供应商分级（考点26—33）', [
            '考点26：底稿中应当列示6家现有供应商年度采购额。',
            '考点27：底稿中应当按采购额从大到小排序。',
            '考点28：底稿中应当按累计占比0-70%划分A级。',
            '考点29：底稿中应当按累计占比70-90%划分B级。',
            '考点30：底稿中应当按累计占比90-100%划分C级。',
            '考点31：底稿中应当给出A/B/C各级供应商数量。',
            '考点32：底稿中应当按A级60%、B级30%、C级10%分配下一年份额。',
            '考点33：底稿中应当对各级供应商给出续约、监控或退出建议。',
        ]),
        ('（五）风险与退出（考点34—40）', [
            '考点34：底稿中应当列示10项风险。',
            '考点35：底稿中应当对每项风险给出量化触发阈值。',
            '考点36：底稿中应当对每项风险给出应对措施。',
            '考点37：底稿中应当列示6类退出机制。',
            '考点38：底稿中应当对每类退出机制给出触发条件。',
            '考点39：底稿中应当对每类退出机制给出回收与损失处理方案。',
            '考点40：底稿中应当明确关键原料单一供应商占比不超过70%。',
        ]),
        ('（六）错误要求驳回（考点41—44）', [
            '考点41：底稿中应当逐条识别管理层12项错误要求。',
            '考点42：底稿中应当识别只看最低价、忽略账期、忽略交货、忽略质检、忽略MOQ、只选一家、主观分级、直接砍C级、盲目加库存、不合格不扣款、只算报价不算TCO、一年只评一次等错误要求。',
            '考点43：底稿中应当对每项错误要求给出不予采纳结论。',
            '考点44：底稿中应当对每项错误要求说明正确做法。',
        ]),
        ('（七）勾稽与结论（考点45—50）', [
            '考点45：底稿中应当完成20项以上勾稽，覆盖原料数、供应商数、投标记录数、维度、权重、中标数、分级、份额、TCO参数。',
            '考点46：底稿中勾稽结论应当全部为一致。',
            '考点47：底稿中应当列示勾稽不一致项数量为0。',
            '考点48：底稿中应当给出核心结论，包含8种原料推荐中标方案和6家供应商分级。',
            '考点49：底稿中核心结论须由四条支撑组成。',
            '考点50：底稿中应当列出关键结果，包括推荐方案年度采购额、TCO合计、分级分布、勾稽状态。',
        ]),
    ]
    for title, items in groups:
        P(doc, title, bold=True, indent=False)
        for item in items:
            P(doc, f'• {item}')

    H(doc, '十、答题思路', 2)
    P(doc, '第一步是读取原料年度需求、新供应商投标、现有供应商年度数据、评分权重、TCO参数、ABC分级标准、份额分配规则、财务确认规则。')
    P(doc, '第二步是按五维度阶梯评分规则计算24条投标记录的综合评分，并对每类原料排序。')
    P(doc, '第三步是按五项成本公式计算每类原料每个供应商的TCO总成本，并找出TCO最低者。')
    P(doc, '第四步是结合综合评分与TCO差异，给出8种原料的推荐中标方案及理由。')
    P(doc, '第五步是按年度采购额对6家现有供应商排序并划分A/B/C级，按60%/30%/10%分配份额。')
    P(doc, '第六步是列出10项风险、6类退出机制，量化触发阈值与回收处理。')
    P(doc, '第七步是逐条驳回12项管理层错误要求，说明正确做法。')
    P(doc, '第八步是完成勾稽自检并给出核心结论。')

    H(doc, '十一、标准答案', 2)
    P(doc, '标准答案共8个包，首表为供应链采购竞价结论，逐包覆盖综合评分、TCO总成本、中标结果、推荐中标方案、现有供应商分级与份额、风险退出、错误驳回和勾稽自检。')

    core_text = concl.get('核心结论', '')
    P(doc, f'核心结论：{core_text}', bold=True)
    for k in ['支撑一', '支撑二', '支撑三', '支撑四']:
        if k in concl:
            P(doc, f'{k}：{concl[k]}')

    P(doc, '关键结果：8种原料推荐中标方案为高筋面粉-S1绿源食品、动物黄油-S1绿源食品、牛油果-S2金谷粮油、抹茶粉-S1绿源食品、红豆-S2金谷粮油、芋头-S2金谷粮油、鸡胸肉-S2金谷粮油、鲜牛奶-S3优品农贸；6家现有供应商A级2家、B级2家、C级2家；所有勾稽结论一致，不一致项为0。', bold=True)

    doc_path = f'{OUT}/供应链采购竞价_v1_题目.docx'
    doc.save(doc_path)
    scrub(doc_path, DOC_DT, is_word=True)

    print('saved', OUT)


if __name__ == '__main__':
    main()
