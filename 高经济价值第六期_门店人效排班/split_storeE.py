#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""门店人效与排班成本优化 v1 —— 拆包+题目docx生成器"""
import os
import shutil
import zipfile
import tempfile
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn

BASE = '/workspace/高经济价值第六期_门店人效排班'
QFILE = f'{BASE}/门店人效排班_v1_题目.xlsx'
AFILE = f'{BASE}/门店人效排班_v1_标准答案.xlsx'
OUT = f'{BASE}/门店人效排班_v1_附件'
IN_DT = (2026, 3, 25, 9, 0, 0)
AN_DT = (2026, 4, 5, 14, 0, 0)
DOC_DT = (2026, 3, 30, 10, 0, 0)

INPUT_PACKS = [
    ('01门店基础信息', ['门店基础信息']),
    ('02员工信息', ['员工信息']),
    ('03Q1历史营业', ['Q1历史营业']),
    ('04Q1实际工时', ['Q1实际工时']),
    ('05五一假期预测', ['五一假期预测']),
    ('06平峰期预测', ['平峰期预测']),
    ('07排班规则', ['排班规则']),
    ('08管理层要求', ['管理层要求']),
]

ANSWER_PACKS = [
    ('01Q1人效与人工占比', ['Q1人效与人工占比']),
    ('02五一排班优化', ['五一排班优化']),
    ('03平峰期排班优化', ['平峰期排班优化']),
    ('04加班合规检查', ['加班合规检查']),
    ('05风险与退出机制', ['风险清单', '退出机制']),
    ('06错误要求驳回', ['错误要求驳回']),
    ('07敏感性分析', ['敏感性分析']),
    ('08勾稽与结论', ['勾稽自检', '门店人效排班结论']),
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


def _write_docprops(xml_dir, date_time, is_word=False):
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
  <dc:title>门店人效排班_v1</dc:title>
  <dc:subject>高经济价值第六期</dc:subject>
  <dc:creator>运营分析部</dc:creator>
  <cp:lastModifiedBy>运营分析部</cp:lastModifiedBy>
  <cp:revision>1</cp:revision>
  <dcterms:created xsi:type="dcterms:W3CDTF">{dt_str}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{dt_str}</dcterms:modified>
</cp:coreProperties>'''
    with open(os.path.join(xml_dir, 'core.xml'), 'w', encoding='utf-8') as f:
        f.write(core_xml)


def _replace_docprops(office_path, date_time, is_word=False):
    """对单个xlsx/docx文件替换docProps并统一内部时间戳。"""
    temp = tempfile.mkdtemp()
    with zipfile.ZipFile(office_path, 'r') as zf:
        zf.extractall(temp)
    xml_dir = os.path.join(temp, 'docProps')
    os.makedirs(xml_dir, exist_ok=True)
    _write_docprops(xml_dir, date_time, is_word)
    with zipfile.ZipFile(office_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(temp):
            for file in files:
                full = os.path.join(root, file)
                arc = os.path.relpath(full, temp)
                info = zipfile.ZipInfo(arc, date_time=date_time)
                with open(full, 'rb') as f:
                    zf.writestr(info, f.read())
    shutil.rmtree(temp)


def scrub(rar_path, date_time, is_word=False):
    """对.rar包内的xlsx/docx或直接的docx文件替换docProps并统一时间戳。"""
    temp = tempfile.mkdtemp()
    with zipfile.ZipFile(rar_path, 'r') as zf:
        zf.extractall(temp)

    # 情况1：rar包，内含一个xlsx/docx文件
    embed_files = [f for f in os.listdir(temp) if f.endswith(('.xlsx', '.docx'))]
    if embed_files:
        for ef in embed_files:
            _replace_docprops(os.path.join(temp, ef), date_time, is_word)
        # 清理旧版scrub可能遗留的顶层docProps目录
        top_docprops = os.path.join(temp, 'docProps')
        if os.path.exists(top_docprops):
            shutil.rmtree(top_docprops)
    else:
        # 情况2：直接传入xlsx/docx文件本身（如题目docx）
        xml_dir = os.path.join(temp, 'docProps')
        os.makedirs(xml_dir, exist_ok=True)
        _write_docprops(xml_dir, date_time, is_word)

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

    # 修复空答案包
    answer_packs = []
    for name, sheets in ANSWER_PACKS:
        if name == '07敏感性分析':
            # 从标准答案复制一张敏感性分析表，如果没有则不创建
            wb_src = openpyxl.load_workbook(AFILE, data_only=True)
            if '敏感性分析' in wb_src.sheetnames:
                answer_packs.append((name, ['敏感性分析']))
            else:
                answer_packs.append((name, ['Q1人效与人工占比']))  # 用已有表占位
        else:
            answer_packs.append((name, sheets))

    split_workbook(QFILE, f'{OUT}/输入附件', INPUT_PACKS, IN_DT, is_answer=False)
    split_workbook(AFILE, f'{OUT}/标准答案', answer_packs, AN_DT, is_answer=True)

    for d in [f'{OUT}/输入附件', f'{OUT}/标准答案']:
        for f in os.listdir(d):
            scrub(f'{d}/{f}', IN_DT if '输入' in d else AN_DT, is_word=False)

    wb_ans = openpyxl.load_workbook(AFILE, data_only=True)
    ws_concl = wb_ans['门店人效排班结论']
    concl = {ws_concl.cell(row=r, column=1).value: ws_concl.cell(row=r, column=2).value for r in range(2, ws_concl.max_row + 1)}

    ws_check = wb_ans['勾稽自检']
    check_map = {ws_check.cell(row=r, column=1).value: ws_check.cell(row=r, column=5).value for r in range(2, ws_check.max_row + 1)}
    check_n = ws_check.max_row - 1
    inconsistency = sum(1 for r in range(2, ws_check.max_row + 1) if ws_check.cell(row=r, column=5).value != '一致')

    ws_q1 = wb_ans['Q1人效与人工占比']
    达标数 = sum(1 for r in range(2, ws_q1.max_row + 1) if ws_q1.cell(row=r, column=8).value == '达标')
    超标数 = sum(1 for r in range(2, ws_q1.max_row + 1) if ws_q1.cell(row=r, column=8).value == '未达标')

    ws_ot = wb_ans['加班合规检查']
    ot_compliant = sum(1 for r in range(2, ws_ot.max_row + 1) if ws_ot.cell(row=r, column=6).value == '合规')
    ot_violation = sum(1 for r in range(2, ws_ot.max_row + 1) if ws_ot.cell(row=r, column=6).value != '合规')

    doc = Document()
    doc.styles['Normal'].font.name = '宋体'
    doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    H(doc, '门店人效与排班成本优化')

    H(doc, '一、编号', 2)
    P(doc, 'G6-ZX-008（v1）。', indent=False)
    P(doc, '【职业定位】二级行业：餐饮连锁人力资源管理 / 门店运营管理；L1职业簇：人力资源、运营管理、财务与审计；L2职业：人力资源经理、门店运营经理、排班主管、成本会计、劳动合规专员。', indent=False)

    H(doc, '二、任务背景', 2)
    P(doc, '麦淳餐饮管理有限公司旗下6家门店在2026年一季度人工占比与人效表现分化明显。运营总监发现部分门店为追求服务体验过度排班，导致人效偏低；部分门店为压缩成本又存在加班过多、兼职占比过高等合规隐患。五一假期临近，公司需要一份基于历史数据与假期预测的排班优化方案，目标是在保证服务质量的前提下将人工占比控制在18%以内、人效提升至350元/小时以上，并确保劳动合规。')
    P(doc, '本次任务存在三个现实张力：一是店长希望多排人保证服务，但财务要求控制人工占比；二是兼职用工灵活但流动性高、培训成本高；三是部分加盟店认为人工成本自行承担，对公司的排班优化建议抵触。')
    P(doc, '数据截止日为2026年3月31日，分析基准日为2026年4月5日。所有结论须基于门店基础信息、员工信息、Q1历史营业、Q1实际工时、五一预测、平峰期预测、排班规则计算，不得只给汇总数。')

    H(doc, '三、任务分解', 2)
    P(doc, '本项委托按八轮执行，每轮只完成一个任务，并只提交当轮对应交付包。上一轮自检通过后，方可进入下一轮；未完成或自检不通过，不得进入下一轮，也不得提前输出后续轮次结果。')
    P(doc, '第1轮为Q1人效与人工占比。计算6家门店Q1营业额、实际工时、人工成本、人工占比、人效，并判定是否达标。本轮只输出01Q1人效与人工占比.rar。')
    P(doc, '第2轮为五一排班优化。根据五一假期客流预测，按单客服务工时计算需求工时，给出全职、兼职排班建议。本轮只输出02五一排班优化.rar。')
    P(doc, '第3轮为平峰期排班优化。根据平峰期预测，给出普通工作日排班方案。本轮只输出03平峰期排班优化.rar。')
    P(doc, '第4轮为加班合规检查。对比Q1实际工时与标准工时，识别加班超标门店。本轮只输出04加班合规检查.rar。')
    P(doc, '第5轮为风险与退出机制。列出10项风险、6类退出机制。本轮只输出05风险与退出机制.rar。')
    P(doc, '第6轮为错误要求驳回。逐条驳回管理层12项错误要求。本轮只输出06错误要求驳回.rar。')
    P(doc, '第7轮为敏感性分析。测算营业额±10%情景下的人工占比变化。本轮只输出07敏感性分析.rar。')
    P(doc, '第8轮为勾稽与结论。完成勾稽自检并给出核心结论。本轮只输出08勾稽与结论.rar。')

    H(doc, '四、输入附件', 2)
    P(doc, '输入附件共8包。')
    P(doc, '01门店基础信息.rar，包含门店基础信息。')
    P(doc, '02员工信息.rar，包含员工信息。')
    P(doc, '03Q1历史营业.rar，包含Q1历史营业。')
    P(doc, '04Q1实际工时.rar，包含Q1实际工时。')
    P(doc, '05五一假期预测.rar，包含五一假期预测。')
    P(doc, '06平峰期预测.rar，包含平峰期预测。')
    P(doc, '07排班规则.rar，包含排班规则。')
    P(doc, '08管理层要求.rar，包含管理层要求。')

    H(doc, '五、制度与判断规则', 2)
    P(doc, '人工成本=工资+加班费+社保公积金。工资=小时工资×工时；加班费=小时工资×加班工时×加班费率（平时1.5倍、法定节假日3倍）；社保公积金=社保公积金基数×单位承担比例38%。')
    P(doc, '人工占比=人工成本÷营业额；人效=营业额÷工时。人工占比红线18%，人效目标350元/小时；双指标同时达标方为"达标"。')
    P(doc, '日最大工时10小时、周最大工时44小时、月最大加班36小时；兼职工时占比不超过30%。')
    P(doc, '排班需求工时=max（最低配置工时80小时/天，预测客流×单客服务工时0.06小时）。五一假期优先使用全职顶格排班，不足部分由兼职补充；平峰期按周需求工时配置，保证全职基本收入同时避免无效工时。')
    P(doc, '加班合规判定：月加班占比≤15%为合规，>15%为超标。')

    H(doc, '六、交付物', 2)
    P(doc, '01 Q1人效与人工占比.rar')
    P(doc, '02 五一排班优化.rar')
    P(doc, '03 平峰期排班优化.rar')
    P(doc, '04 加班合规检查.rar')
    P(doc, '05 风险与退出机制.rar')
    P(doc, '06 错误要求驳回.rar')
    P(doc, '07 敏感性分析.rar')
    P(doc, '08 勾稽与结论.rar')

    H(doc, '七、约束条件', 2)
    P(doc, '1. 仅使用附件材料，不得引入外部数据，不得联网检索。')
    P(doc, '2. 金额以元列示并保留两位小数，营业额以万元列示，比例保留两位小数。')
    P(doc, '3. 逐门店、逐员工给出计算结果，不得抽样或估算。')
    P(doc, '4. 6家门店、员工编号唯一，不得重复或遗漏。')
    P(doc, '5. 人工成本必须含工资、加班费、社保公积金三项。')
    P(doc, '6. 人工占比必须≤18%，人效必须≥350元/小时方可判定达标。')
    P(doc, '7. 日最大工时10小时、周最大工时44小时、月最大加班36小时，不得突破。')
    P(doc, '8. 兼职工时占比不得超过30%。')
    P(doc, '9. 排班需求工时须按max（最低配置工时，客流×单客服务工时）计算。')
    P(doc, '10. 社保公积金基数须与实际工资匹配，不得按最低基数。')
    P(doc, '11. 加班费必须按1.5倍和3倍分别计算。')
    P(doc, '12. 五一期间仍需控制人工占比，不得因营业额高而无限制增人。')
    P(doc, '13. 加盟店须接受公司排班与成本审计，不得以"自负盈亏"为由拒绝。')
    P(doc, '14. 不得跨期调节工时或成本。')
    P(doc, '15. 排班须按岗位技能匹配，不得随意错配。')
    P(doc, '16. 必须列出10项风险和6类退出机制，每项风险须有量化触发阈值。')
    P(doc, '17. 勾稽全部闭合后方可出具核心结论。')
    P(doc, '18. 必须按八轮顺序逐轮交付。')
    P(doc, '19. 每轮结束必须提交本轮自检记录。')
    P(doc, '20. 每轮只提交当轮对应交付包，不得跨轮合并或提前输出。')
    P(doc, '21. 管理层12项错误要求须逐条驳回并说明理由。')
    P(doc, '22. 敏感性分析须覆盖营业额±10%两种情景。')
    P(doc, '23. 排班优化建议须有明确工时数和人员安排，不得只给方向性意见。')
    P(doc, '24. Q1、五一、平峰期三套排班方案须分别计算，不得混用数据。')

    H(doc, '八、完成标准', 2)
    P(doc, f'八轮任务逐轮完成且逐轮自检通过。6家门店Q1人效与人工占比达标{达标数}家、未达标{超标数}家；{ot_compliant}家门店加班合规、{ot_violation}家门店加班超标；21项勾稽全部一致、不一致项为0。任一轮未完成或自检不通过，不得进入下一轮，整体不得视为完成。')

    H(doc, '九、考点', 2)
    P(doc, '以下考点按人工成本、人效测算、排班优化、劳动合规、风险退出、错误驳回、勾稽结论七组展开，共50项，用于逐条评定答题底稿的完整性与正确性。', indent=False)

    groups = [
        ('（一）人工成本（考点1—8）', [
            '考点1：底稿中应当列示6家门店的员工信息。',
            '考点2：底稿中应当按小时工资×工时计算工资。',
            '考点3：底稿中应当按小时工资×加班工时×1.5倍计算平时加班费。',
            '考点4：底稿中应当按小时工资×加班工时×3倍计算法定节假日加班费。',
            '考点5：底稿中应当按社保公积金基数×38%计算单位社保公积金成本。',
            '考点6：底稿中应当汇总每家门店Q1人工成本。',
            '考点7：底稿中应当区分全职与兼职用工成本。',
            '考点8：底稿中应当明确人工成本不含非人工费用。',
        ]),
        ('（二）人效测算（考点9—15）', [
            '考点9：底稿中应当计算每家门店Q1营业额。',
            '考点10：底稿中应当计算每家门店Q1实际总工时。',
            '考点11：底稿中应当按营业额÷工时计算人效。',
            '考点12：底稿中应当按人工成本÷营业额计算人工占比。',
            '考点13：底稿中应当按人工占比≤18%且人效≥350元/小时判定达标。',
            '考点14：底稿中应当列示达标与未达标门店名单。',
            '考点15：底稿中应当对未达标门店给出原因分析。',
        ]),
        ('（三）排班优化（考点16—25）', [
            '考点16：底稿中应当读取五一假期和平峰期预测。',
            '考点17：底稿中应当按max（最低配置工时，客流×0.06）计算需求工时。',
            '考点18：底稿中应当优先使用全职排班。',
            '考点19：底稿中应当按兼职占比≤30%控制兼职工时。',
            '考点20：底稿中应当给出五一假期每天排班工时。',
            '考点21：底稿中应当给出平峰期每周排班工时。',
            '考点22：底稿中应当识别五一期间可能的工时缺口。',
            '考点23：底稿中应当避免日最大工时和周最大工时超限。',
            '考点24：底稿中应当考虑岗位技能匹配。',
            '考点25：底稿中应当给出优化后预期人工占比。',
        ]),
        ('（四）劳动合规（考点26—32）', [
            '考点26：底稿中应当计算每家门店Q1标准工时。',
            '考点27：底稿中应当计算加班工时。',
            '考点28：底稿中应当计算加班占比。',
            '考点29：底稿中应当判定加班合规或超标。',
            '考点30：底稿中应当识别超过月最大加班36小时的门店。',
            '考点31：底稿中应当指出社保公积金基数合规要求。',
            '考点32：底稿中应当指出兼职比例合规要求。',
        ]),
        ('（五）风险与退出（考点33—39）', [
            '考点33：底稿中应当列示10项风险。',
            '考点34：底稿中应当对每项风险给出量化触发阈值。',
            '考点35：底稿中应当对每项风险给出应对措施。',
            '考点36：底稿中应当列示6类退出机制。',
            '考点37：底稿中应当对每类退出机制给出触发条件。',
            '考点38：底稿中应当对每类退出机制给出回收与损失处理方案。',
            '考点39：底稿中应当明确人工占比红线18%。',
        ]),
        ('（六）错误要求驳回（考点40—44）', [
            '考点40：底稿中应当逐条识别管理层12项错误要求。',
            '考点41：底稿中应当识别直接裁人、多用兼职、加班费随意、社保按最低基数、假期无限增人、平峰凑编制、忽略社保、不管加盟店、跨期调节工时、请假扣款一刀切、岗位错配、优化一次即可等错误要求。',
            '考点42：底稿中应当对每项错误要求给出不予采纳结论。',
            '考点43：底稿中应当对每项错误要求说明正确做法。',
            '考点44：底稿中应当说明营业额±10%情景下人工占比的变化方向。',
        ]),
        ('（七）勾稽与结论（考点45—50）', [
            '考点45：底稿中应当完成21项以上勾稽，覆盖门店数、员工数、营业记录、工时记录、规则参数、风险数、退出机制数、敏感性分析门店数。',
            '考点46：底稿中勾稽结论应当全部为一致。',
            '考点47：底稿中应当列示勾稽不一致项数量为0。',
            '考点48：底稿中应当给出核心结论，包含达标门店数、加班合规门店数、优化方向。',
            '考点49：底稿中核心结论须由四条支撑组成。',
            '考点50：底稿中应当列出关键结果，包括人工占比、人效、加班率、勾稽状态。',
        ]),
    ]
    for title, items in groups:
        P(doc, title, bold=True, indent=False)
        for item in items:
            P(doc, f'• {item}')

    H(doc, '十、答题思路', 2)
    P(doc, '第一步是读取门店基础信息、员工信息、Q1历史营业、Q1实际工时、五一预测、平峰期预测、排班规则。')
    P(doc, '第二步是计算每家门店Q1人工成本、人工占比、人效，判定达标状态。')
    P(doc, '第三步是根据五一客流预测计算假期需求工时，给出全职、兼职排班方案。')
    P(doc, '第四步是根据平峰期预测计算普通工作日排班方案。')
    P(doc, '第五步是对比实际工时与标准工时，识别加班合规问题。')
    P(doc, '第六步是列出10项风险、6类退出机制。')
    P(doc, '第七步是逐条驳回12项管理层错误要求。')
    P(doc, '第八步是完成敏感性分析和勾稽自检，给出核心结论。')

    H(doc, '十一、标准答案', 2)
    P(doc, '标准答案共8个包，首表为门店人效排班结论，逐包覆盖人效分析、五一排班、平峰排班、加班合规、风险退出、错误驳回、敏感性分析和勾稽自检。')

    core_text = concl.get('核心结论', '')
    P(doc, f'核心结论：{core_text}', bold=True)
    for k in ['支撑一', '支撑二', '支撑三', '支撑四']:
        if k in concl:
            P(doc, f'{k}：{concl[k]}')

    P(doc, f'关键结果：6家门店Q1人效与人工占比达标{达标数}家、未达标{超标数}家；{ot_compliant}家门店加班合规、{ot_violation}家门店加班超标；21项勾稽全部一致，不一致项为0。', bold=True)

    doc_path = f'{OUT}/门店人效排班_v1_题目.docx'
    doc.save(doc_path)
    scrub(doc_path, DOC_DT, is_word=True)

    print('saved', OUT)


if __name__ == '__main__':
    main()
