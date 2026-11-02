#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""销售佣金清算 v1 附件拆分与打包：8输入 + 8答案（含 docx/png）"""
import openpyxl, os, io, zipfile, shutil
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

BASE = '/workspace/高经济价值第六期_销售佣金清算'
SPLIT = '/tmp/comm_v1_split'
OUT = f'{BASE}/销售佣金清算_v1_附件'
shutil.rmtree(SPLIT, ignore_errors=True)
shutil.rmtree(OUT, ignore_errors=True)
os.makedirs(f'{SPLIT}/input', exist_ok=True)
os.makedirs(f'{SPLIT}/ans', exist_ok=True)
os.makedirs(f'{OUT}/输入附件', exist_ok=True)
os.makedirs(f'{OUT}/标准答案', exist_ok=True)

HDR_FILL = PatternFill('solid', fgColor='2F5597')
HDR_FONT = Font(name='宋体', bold=True, color='FFFFFF', size=10)
BODY = Font(name='宋体', size=9)
THIN = Side(style='thin', color='CCCCCC')
BORDER = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
WRAP = Alignment(wrap_text=True, vertical='top')
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)

EXCEL_APP = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
             '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
             'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
             '<Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>'
             '<Company>睿联工业设备有限公司</Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc>'
             '<HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0300</AppVersion></Properties>')
WORD_APP = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
            'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
            '<Application>Microsoft Office Word</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>'
            '<Company>睿联工业设备有限公司</Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc>'
            '<HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0000</AppVersion></Properties>')


def core_xml(author, created, modified):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            f'<dc:creator>{author}</dc:creator><cp:lastModifiedBy>{author}</cp:lastModifiedBy>'
            f'<dcterms:created xsi:type="dcterms:W3CDTF">{created}</dcterms:created>'
            f'<dcterms:modified xsi:type="dcterms:W3CDTF">{modified}</dcterms:modified></cp:coreProperties>')


def scrub(path, app_xml, author, created, modified, zip_dt=(2025, 9, 5, 9, 5, 0)):
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
    ts = __import__('time').mktime((zip_dt[0], zip_dt[1], zip_dt[2], zip_dt[3], zip_dt[4], zip_dt[5], 0, 0, -1))
    os.utime(path, (ts, ts))


def strip_png(path, stamp=(2025, 9, 5, 9, 0, 0)):
    """移除 PNG 的 tEXt/zTXt/iTXt/tIME 等文本块，仅保留图像必需块"""
    keep = {b'IHDR', b'PLTE', b'IDAT', b'IEND', b'tRNS', b'gAMA', b'pHYs', b'sRGB'}
    with open(path, 'rb') as f:
        data = f.read()
    if data[:8] != b'\x89PNG\r\n\x1a\n':
        return
    out = bytearray(data[:8])
    pos = 8
    while pos + 8 <= len(data):
        ln = int.from_bytes(data[pos:pos + 4], 'big')
        typ = data[pos + 4:pos + 8]
        end = pos + 12 + ln
        if typ in keep:
            out += data[pos:end]
        pos = end
        if typ == b'IEND':
            break
    with open(path, 'wb') as f:
        f.write(bytes(out))
    ts = __import__('time').mktime((stamp[0], stamp[1], stamp[2], stamp[3], stamp[4], stamp[5], 0, 0, -1))
    os.utime(path, (ts, ts))


def dump(ws, path):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    nws = wb.create_sheet(ws.title)
    for r_idx, row in enumerate(ws.iter_rows(values_only=True), 1):
        for c_idx, val in enumerate(row, 1):
            c = nws.cell(row=r_idx, column=c_idx, value=val)
            if r_idx == 1:
                c.font = HDR_FONT; c.fill = HDR_FILL; c.alignment = CENTER; c.border = BORDER
            else:
                c.font = BODY; c.alignment = WRAP; c.border = BORDER
    for i in range(1, ws.max_column + 1):
        nws.column_dimensions[get_column_letter(i)].width = ws.column_dimensions[get_column_letter(i)].width or 14
    nws.freeze_panes = 'A2'
    wb.save(path)
    return nws.max_row - 1


def make_rar(files, rar_path):
    with zipfile.ZipFile(rar_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for f in files:
            zf.write(f, arcname=os.path.basename(f))
    ts = max(os.path.getmtime(f) for f in files)
    os.utime(rar_path, (ts, ts))


src_q = openpyxl.load_workbook(f'{BASE}/销售佣金清算_v1_题目.xlsx', data_only=True)
src_a = openpyxl.load_workbook(f'{BASE}/销售佣金清算_v1_标准答案.xlsx', data_only=True)
IN, AN = f'{SPLIT}/input', f'{SPLIT}/ans'
rows = {}
for sn in src_q.sheetnames:
    p = f'{IN}/{sn}.xlsx'
    rows[sn] = dump(src_q[sn], p)
    scrub(p, EXCEL_APP, '销售运营部', '2025-08-25T02:10:00Z', '2025-08-27T07:35:00Z', (2025, 8, 27, 7, 35, 0))
for sn in src_a.sheetnames:
    p = f'{AN}/{sn}.xlsx'
    rows[sn] = dump(src_a[sn], p)
    scrub(p, EXCEL_APP, '薪酬与激励项目组', '2025-09-03T01:20:00Z', '2025-09-05T09:05:00Z', (2025, 9, 5, 9, 5, 0))

# ---------------- docx ----------------
def new_doc():
    doc = Document()
    st = doc.styles['Normal']
    st.font.name = '宋体'
    st.font.size = Pt(10.5)
    from docx.oxml.ns import qn
    rpr = st.element.get_or_add_rPr()
    rpr.get_or_add_rFonts().set(qn('w:eastAsia'), '宋体')
    for s in doc.sections:
        s.top_margin = Cm(2.2); s.bottom_margin = Cm(2.2)
        s.left_margin = Cm(2.6); s.right_margin = Cm(2.6)
    return doc


def H(doc, text, level=1):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.bold = True
    r.font.size = Pt(14 if level == 1 else 12)
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(6)
    return p


def P(doc, text, bold=False, indent=True):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.35
    if indent:
        p.paragraph_format.first_line_indent = Cm(0.74)
    r = p.add_run(text)
    r.bold = bold
    return p


def table(doc, headers, data, widths=None):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(headers):
        c = t.rows[0].cells[i]
        c.text = str(h)
        for p in c.paragraphs:
            for r in p.runs:
                r.bold = True
                r.font.size = Pt(9)
    for row in data:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = str(v)
            for p in cells[i].paragraphs:
                for r in p.runs:
                    r.font.size = Pt(9)
    if widths:
        for i, w in enumerate(widths):
            for row in t.rows:
                row.cells[i].width = Cm(w)
    return t


# 1 争议裁决说明
doc = new_doc()
H(doc, '销售佣金争议裁决说明')
P(doc, '一、裁决范围')
P(doc, f'本次共受理佣金争议 {rows["佣金争议台账"]} 件，涉及60个订单。裁决以《佣金政策V1》《V2》《V3》《财务确认规则》及正式发文为依据，逐单复核，不采用申诉人自报口径。')
H(doc, '二、裁决口径', 2)
P(doc, '1. 政策适用以订单发货日期所在年度为准；2023年1月1日起执行V2，以睿联销字〔2023〕07号正式发文为准，2023年3月10日经营会纪要属讨论稿，不作为追溯依据。')
P(doc, '2. 计提基数按各版政策执行：V1为含税发货额扣退货；V2为含税回款扣退款；V3为不含税净回款（含税回款减退款、减坏账核销后除以1.13）。')
P(doc, '3. 渠道单佣金按直销费率减半，不得与渠道返利重复计提。')
P(doc, '4. 合作单自2024年起按主谈70%、协助30%分配。')
P(doc, '5. 退货、坏账按制度冲减；离职人员自离职次月起回款不再计提；月度封顶按政策执行。')
H(doc, '三、裁决结果', 2)
P(doc, '60 件争议中，支持 2 件，部分支持 53 件，不支持 5 件；裁决金额合计 660.51 万元，申诉金额合计 1575.07 万元，核减 914.56 万元。')
P(doc, '主要核减原因：申诉人按订单合同额直接乘以3%—5%主张，未扣除退货、坏账、封顶及合作单拆分；部分订单在离职后回款仍主张全额；渠道单主张按3%计提。')
H(doc, '四、典型争议处理', 2)
table(doc, ['争议类型', '处理结论', '依据'],
      [['按合同额×3%主张', '部分支持，按实际计提基数重算', 'V1-V3计提基数条款'],
       ['离职后回款主张', '不支持或部分支持，次月起不计提', 'V2-6、V3-7'],
       ['渠道单按3%主张', '部分支持，按1.5%裁决', 'V3-3、睿联销字〔2024〕15号'],
       ['合作单全额归属', '部分支持，2024年起按7:3拆分', 'V3-5'],
       ['退货/坏账不冲减', '不支持，按净额重算', '财务确认规则第5、6条'],
       ['封顶条款无效主张', '不支持，制度已公示', 'V1-6、V3-6']],
      [4.2, 5.6, 5.6])
P(doc, '附：逐单裁决明细见《争议裁决》工作表，含申诉金额、裁决结论、裁决金额与依据条款。')
doc.core_properties.author = '薪酬与激励项目组'
doc.save(f'{AN}/争议裁决说明.docx')
scrub(f'{AN}/争议裁决说明.docx', WORD_APP, '薪酬与激励项目组', '2025-09-04T02:00:00Z', '2025-09-05T08:30:00Z', (2025, 9, 5, 8, 30, 0))

# 2 销售激励方案
doc = new_doc()
H(doc, '2026年销售佣金与渠道激励方案')
P(doc, '一、设计目标')
P(doc, '在佣金总成本较2024年旧方案增幅不超过15%的前提下，解决历史政策中“重个人、轻团队”“渠道重复激励”“回款质量无人负责”三个问题，把销售行为从“签单导向”转到“净回款导向”。')
H(doc, '二、方案结构', 2)
table(doc, ['要素', '2026年方案', '与旧方案差异'],
      [['计提基数', '不含税净回款（扣退款、扣坏账）', '延续V3，明确坏账全额扣减'],
       ['直销费率', '2.8%（达成率110%以内）', '较V3的3.0%下调0.2个百分点'],
       ['加速档', '110%-130%部分3.5%，130%以上4.2%', '新增两档加速，替代V2的单笔500万档'],
       ['渠道费率', '1.4%，不参与加速', '较V3的1.5%下调，避免与返利重复'],
       ['分配结构', '个人70%、团队池30%', '个人比例由80%降至70%'],
       ['团队池发放', '团队年度达成率≥100%方可发放', '延续V3，未达标留存'],
       ['月度封顶', '5倍月基本工资', '较V3的12倍大幅收紧'],
       ['新人保底', '10000元/月，入职前3个自然月', '较V3的8000元提高'],
       ['离职规则', '次月起停计，退货坏账全额追回', '延续V3并明确追回'],
       ['清算节奏', '季度预发70%，年度清算，追溯12个月', '新增追溯期']],
      [3.2, 7.2, 5.2])
H(doc, '三、回测结果', 2)
P(doc, '以2024年全年数据回测：新方案佣金总成本 1430.78 万元，旧方案 1295.47 万元，增加 135.31 万元，增幅 10.44%，控制在15%以内；其中个人部分占70%，团队池占30%。')
P(doc, '敏感性测算显示：回款下降10%且费率下调0.2个百分点时，成本降至1196.80万元；回款增长10%且费率上调0.2个百分点时，成本升至1682.31万元，仍在可承受区间。')
H(doc, '四、配套机制', 2)
P(doc, '1. 数据口径：统一以财务确认的净回款为准，订单去重、退货、坏账、离职状态每月5日前同步。')
P(doc, '2. 争议处理：季度预发前公示个人清算表，争议在5个工作日内书面答复。')
P(doc, '3. 过渡安排：2022-2025年8月历史佣金按原政策清算，不溯及既往；2026年1月1日起执行新方案。')
P(doc, '4. 风险控制：设置团队池留存与追回条款，防止冲量后退货、坏账造成超发。')
doc.core_properties.author = '薪酬与激励项目组'
doc.save(f'{AN}/销售激励方案.docx')
scrub(f'{AN}/销售激励方案.docx', WORD_APP, '薪酬与激励项目组', '2025-09-02T03:10:00Z', '2025-09-05T06:20:00Z', (2025, 9, 5, 6, 20, 0))

# 3 清算结论
doc = new_doc()
H(doc, '销售佣金与渠道返利历史清算结论')
P(doc, '一、核心结论')
P(doc, '历史佣金整体超发 273.13 万元、渠道返利少付 450.59 万元，两项相抵公司净少支出 177.46 万元；问题根源是手工台账按订单额粗算，未执行去重、封顶、退货、坏账与离职规则，建议逐人补追并在2025年四季度一次性清理。')
P(doc, '二、总结论')
P(doc, '经对2022年1月1日至2025年8月31日全部720个有效订单逐单复核，历史佣金应发 4402.42 万元，财务已发 4675.54 万元，整体超发 273.13 万元；其中应补 545.57 万元、应追回 818.70 万元，涉及 117 人。渠道返利应付 1268.23 万元、已付 817.65 万元，应补 450.59 万元。')
H(doc, '三、主要发现', 2)
P(doc, '1. 订单主表存在12条重复导入记录，涉及金额约 1.1 亿元，已按首次导入记录去重，避免重复计提。')
P(doc, '2. V1政策下38个订单因回款率不足90%应挂账，已发台账却已预发，形成超发。')
P(doc, '3. 财务预发按“订单额×3%×70%”粗算，未执行月度封顶、未扣退货与坏账，是超发主因；封顶扣减合计605.21万元。')
P(doc, '4. 2023年V2政策被误按4月1日执行的说法存在于会议纪要，按正式发文应自1月1日追溯，影响金额已重算。')
P(doc, '5. 2024年起30个合作单未按7:3拆分，全部计入签约人，涉及佣金已调整。')
P(doc, '6. 渠道返利按含税进货额与阶梯比例重算，5家保底渠道按“阶梯与保底孰高”执行；2025年度未到结算期，未计入应付。')
H(doc, '四、争议与制度', 2)
P(doc, '60件争议中支持2件、部分支持53件、不支持5件，裁决金额660.51万元，核减914.56万元。管理层提出的12项要求均与正式制度冲突，已逐条驳回（详见《勾稽自检》与《错误要求驳回说明》）。')
H(doc, '五、新方案建议', 2)
P(doc, '建议2026年执行“个人70%+团队池30%、净回款计提、两档加速、5倍封顶”的新方案，2024年回测成本增幅10.44%，在15%控制线内；同时将历史清算差额在2025年四季度一次性处理，避免跨年争议。')
H(doc, '六、勾稽结论', 2)
P(doc, '24项勾稽全部通过：订单数、三版政策订单数、逐单与人员汇总、已发台账、应补应追回人数、已发异常、争议三类、渠道返利、新方案个人与团队拆分均闭合。')
doc.core_properties.author = '薪酬与激励项目组'
doc.save(f'{AN}/清算结论.docx')
scrub(f'{AN}/清算结论.docx', WORD_APP, '薪酬与激励项目组', '2025-09-05T01:30:00Z', '2025-09-05T10:00:00Z', (2025, 9, 5, 10, 0, 0))

# ---------------- charts ----------------
plt.rcParams['font.sans-serif'] = ['Noto Sans CJK SC', 'Noto Sans CJK JP', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

fig, ax = plt.subplots(figsize=(6.4, 3.6), dpi=200)
vals = [1295.47, 1430.78]
bars = ax.bar(['旧方案（V3）', '新方案（2026）'], vals, color=['#8EA9DB', '#2F5597'], width=0.5)
for b, v in zip(bars, vals):
    ax.text(b.get_x() + b.get_width() / 2, v + 15, f'{v:,.2f}', ha='center', fontsize=9)
ax.set_ylabel('佣金成本（万元）')
ax.set_title('2024年回测：新旧方案佣金成本对比（+10.44%）', fontsize=11)
ax.set_ylim(0, 1650)
ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
fig.tight_layout()
fig.savefig(f'{AN}/新旧成本对比.png')
plt.close(fig)
strip_png(f'{AN}/新旧成本对比.png')

fig, ax = plt.subplots(figsize=(6.4, 3.6), dpi=200)
labels = ['V1\n2022', 'V2\n2023', 'V3个人\n2024-2025.8', '新人保底', '团队池']
vals = [1063.88, 1909.13, 1407.22, 22.18, 358.08]
bars = ax.bar(labels, vals, color=['#A9D08E', '#FFC000', '#5B9BD5', '#ED7D31', '#70AD47'], width=0.55)
for b, v in zip(bars, vals):
    ax.text(b.get_x() + b.get_width() / 2, v + 20, f'{v:,.2f}', ha='center', fontsize=8.5)
ax.set_ylabel('金额（万元）')
ax.set_title('历史佣金应发结构（合计 4,760.49 万元）', fontsize=11)
ax.set_ylim(0, 2200)
ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
fig.tight_layout()
fig.savefig(f'{AN}/佣金结构.png')
plt.close(fig)
strip_png(f'{AN}/佣金结构.png')

# ---------------- 题目 docx ----------------
doc = new_doc()
doc.add_heading('销售佣金与渠道返利的历史清算及2026年新激励方案设计', 0)
H(doc, '一、编号')
P(doc, 'G6-ZX-002（v1）。', indent=False)
H(doc, '二、任务背景')
P(doc, '睿联工业设备有限公司是一家工业泵阀制造企业，销售采用直销与渠道并行模式。2025年7月，公司发现历史佣金发放与渠道返利长期依赖手工台账，销售人员对金额争议集中爆发。现受董事会审计委员会委托，以外部薪酬与销售运营顾问身份，对2022年1月1日至2025年8月31日的销售佣金和渠道返利进行全量清算，并设计2026年激励方案。')
P(doc, '数据截止日为2025年8月31日，清算基准日为2025年9月1日。所有结论须可回溯到附件逐条记录，不得只给汇总数。凡涉及剔除、挂账、追回、留存、不结的记录，必须逐条说明触发条件与条款依据。')
H(doc, '三、任务分解')
P(doc, '本项委托按八轮执行，每轮只完成一个任务，并只提交当轮对应交付包。上一轮自检通过后，方可进入下一轮；未完成或自检不通过，不得进入下一轮，也不得提前输出后续轮次结果。')
P(doc, '第1轮为订单佣金清算。完成订单去重、政策版本归属、含税与不含税换算，按V1、V2、V3逐单计提，处理退货、坏账、离职、封顶和合作单拆分。本轮只输出01订单佣金清算.rar。本轮自检要求全部720个有效订单均有清算结论，V1、V2、V3订单数分别为168、197、355单，12条重复导入行已逐条留痕。')
P(doc, '第2轮为人员汇总与异常核查。逐人汇总V1、V2、V3应发，叠加新人保底，并与已发台账比对，识别异常预发。本轮只输出02人员汇总与异常核查.rar。本轮自检要求120名销售均有应发、已发和差额，13条异常预发记录均单列。')
P(doc, '第3轮为争议裁决。对60件佣金争议逐单裁决，并写明依据条款。本轮只输出03争议裁决.rar。本轮自检要求60件争议均有裁决，支持、部分支持和不支持三类数量闭合。')
P(doc, '第4轮为渠道返利清算。按渠道和年度套用阶梯并与保底孰高，扣除已付金额，标注2025年度未到结算期。本轮只输出04渠道返利台账.rar。本轮自检要求40家渠道乘4个年度共160行齐全，2025年度不计入本期应付。')
P(doc, '第5轮为新激励方案。设计2026年计提基数、个人与团队池比例、费率、封顶、保底、渠道返利和结算周期。本轮只输出05新激励方案.rar。本轮自检要求10项参数完整，新旧成本对比可复算。')
P(doc, '第6轮为模拟测算与敏感性。使用2024年全年订单回测新方案，并完成9组“回款增长率×直销费率”敏感性测算。本轮只输出06模拟测算与敏感性.rar。本轮自检要求120人逐人对比齐全，9组敏感性均有结果，成本增幅不超过15%。')
P(doc, '第7轮为勾稽自检。完成订单数、人数、金额、封顶、团队池、渠道返利和新旧成本共24项勾稽。本轮只输出07勾稽自检.rar。本轮自检要求24项全部闭合且不一致项为0；未闭合不得进入第8轮。')
P(doc, '第8轮为结论与驳回。给出一句话核心结论、结论摘要、12项错误管理要求驳回和清算结论。本轮只输出08结论与驳回.rar。本轮自检要求核心结论可由明细与勾稽逐项回溯，12项驳回均引用依据条款。')
H(doc, '四、输入附件')
P(doc, '输入附件共8包。')
P(doc, '01组织与人员薪酬.rar，包含销售花名册、团队目标与薪酬、区域划分。')
P(doc, '02订单主表.rar，包含订单主表。')
P(doc, '03回款退货与坏账.rar，包含回款流水、退货记录、坏账核销。')
P(doc, '04佣金政策三版.rar，包含佣金政策V1、V2、V3。')
P(doc, '05政策发文与会议纪要.rar，包含政策变更通知、经营会会议纪要。')
P(doc, '06渠道协议与返利.rar，包含渠道商清单、渠道返利阶梯、渠道协议摘要、渠道订单明细、渠道返利已付台账。')
P(doc, '07争议与申诉.rar，包含佣金争议台账、申诉沟通记录。')
P(doc, '08财务口径与管理要求.rar，包含财务确认规则、税务与换算口径、佣金已发台账、管理层要求。')
H(doc, '五、制度与判断规则')
P(doc, 'V1适用于2022年发货订单。计提基数为含税发货额扣减退货，标准费率为3%，新客户首单为4%，累计回款率达到90%方可发放，月度封顶为10倍月基本工资。')
P(doc, 'V2适用于2023年发货订单。计提基数为含税回款扣减退款，标准费率为3.5%，单笔累计回款超过500万元的部分适用5%，不设封顶。')
P(doc, 'V3适用于2024年1月1日至2025年8月31日发货订单。计提基数为不含税净回款，直销费率为3%，渠道单费率为1.5%，个人分配80%，团队池分配20%，合作单按主谈70%、协助30%拆分，月度封顶为12倍月基本工资，新人前3个月保底8000元。')
P(doc, '渠道返利按年度含税进货额套用2%至5%阶梯，进货额按发货额扣减退货计算。设有保底条款的渠道按阶梯返利与保底金额孰高确认，次年1月结算。')
P(doc, '判断规则如下。', bold=True, indent=False)
for t in ['1. 同一订单号多次出现时，以首次导入记录为准，其余作为重复行剔除，但不得整单删除。',
          '2. 政策版本按订单发货日期归属，不按订单日期或回款日期；2025年发货仍适用V3。',
          '3. 会议纪要和讨论稿不作为计算依据，以正式发文为准，并注意追溯条款。',
          '4. V1和V2为含税口径，V3须除以1.13换算为不含税净回款。',
          '5. 退货冲减计提基数和已发佣金，坏账从V3净回款中扣除。',
          '6. 封顶按“人×年月”累计，超出部分不结转以后月份。',
          '7. V1和V2合作单全部归签约人，V3按主谈70%、协助30%拆分。',
          '8. 团队池按团队年度目标达成率不低于100%发放，未达标部分留存。',
          '9. 佣金已发台账为财务实付数，不作为应发依据。',
          '10. 渠道返利与销售佣金是两套激励，不得重复计入。']:
    P(doc, t, indent=False)
H(doc, '六、交付物')
for t in ['01 订单佣金清算.rar',
          '02 人员汇总与异常核查.rar',
          '03 争议裁决.rar',
          '04 渠道返利台账.rar',
          '05 新激励方案.rar',
          '06 模拟测算与敏感性.rar',
          '07 勾稽自检.rar',
          '08 结论与驳回.rar']:
    P(doc, t, indent=False)
H(doc, '七、约束条件')
for t in ['1. 仅使用附件材料，不得引入外部数据，不得联网检索。',
          '2. 所有金额以人民币元为单位并保留两位小数，比例保留两位；明细中不得用“万元”替代“元”。',
          '3. 逐单、逐人给出结论，不得抽样或估算；每条记录必须有明确状态，不得留空。',
          '4. 订单去重、政策归属、文件效力和金额口径以《财务确认规则》为准。',
          '5. 政策适用以订单发货日期为准，不按订单日期或回款日期归属。',
          '6. 退货、坏账、离职、封顶和合作单拆分必须逐单体现，不得只做总额调整。',
          '7. 每个争议裁决、每条管理层要求驳回均须给出依据条款编号。',
          '8. 八个交付包之间的订单数、人数和金额必须相互勾稽，24项自检全部闭合；任一项不闭合视为未完成。',
          '9. 不得修改原始附件；交付文件按包名和工作表命名，表头和列序不得调整。',
          '10. 新方案须在旧方案基础上测算成本，并给出9组敏感性结果。',
          '11. 已发台账中的13条异常预发记录须单列核查，注明异常类型、应补、应追回或一致，并给出处理建议。',
          '12. 结论中必须给出一句话核心结论，且能被明细数据支撑；不得只写“存在若干问题”。',
          '13. 重复导入行、统计截止日后的回款或退款、负净回款订单和无回款订单，必须逐条列出处理结论，不得静默丢弃。',
          '14. 计算底稿须保留“政策版本、事件日期、计提基数、费率、封顶扣减”五要素，便于逐行复核。',
          '15. 不得使用合并单元格、空表或空列；金额差异为正写“应补”，为负写“应追回”。',
          '16. 必须按八轮顺序逐轮交付，每轮只完成一个任务，不得一次性输出全部结果。',
          '17. 每轮结束必须提交本轮自检记录，列明检查项、结果、证据和未通过原因。',
          '18. 上一轮自检未通过或交付物不完整，不得进入下一轮。',
          '19. 每轮只提交当轮对应交付包，不得跨轮合并、提前输出或补交后续轮次内容。',
          '20. 任何一轮无法完成时，必须停止在该轮并说明缺口，不得跳过或降级处理。',
          '21. 最终交付以八轮全部通过为前提；任一轮未通过，整体视为未完成。']:
    P(doc, t, indent=False)
H(doc, '八、完成标准')
P(doc, '八轮任务逐轮完成且逐轮自检通过。全部720个有效订单均有清算结论，120名销售均有应发、已发和差额，60件争议均有裁决，40家渠道的2022年至2024年度返利均清算完毕，13条已发异常记录均单列，24项勾稽全部通过，新方案回测成本增幅不超过15%，核心结论与明细数据一致。任一轮未完成或自检不通过，不得进入下一轮，整体不得视为完成。')

H(doc, '九、考点')
P(doc, '考点按数据治理、金额口径、政策效力、三版清算、人员汇总、团队池、已发异常、渠道返利、争议驳回、新方案和勾稽结论十三组展开，共50项。', indent=False)
H(doc, '（一）数据治理与订单去重（考点1—5）', 2)
for t in ['考点1：订单号重复出现时，以首次导入记录为准，其余作为重复行剔除，不得整单删除。',
          '考点2：逐条列明12条重复导入行的保留对象和剔除对象，并在订单清算底稿中保留可追溯说明。',
          '考点3：回款、退货和坏账的统计截止日统一锁定为2025年8月31日，基准日之后的流水不得计入本期清算。',
          '考点4：政策版本按发货日期而不是订单日期或回款日期确定，避免跨期订单误用费率。',
          '考点5：识别17个V3无回款订单、12个负净回款订单以及统计截止日后的回款退款，并逐条给出处理结论。']:
    P(doc, t, indent=False)
H(doc, '（二）金额口径与换算（考点6—8）', 2)
for t in ['考点6：V1和V2按含税金额计算，V3按含税金额除以1.13换算为不含税净回款。',
          '考点7：退货同时冲减计提基数和已发佣金，跨期退货回到原订单事件流处理，不得只做总额调整。',
          '考点8：坏账核销作为负回款事件冲减V3净回款，并在逐单明细中保留核销日期、金额和依据。']:
    P(doc, t, indent=False)
H(doc, '（三）政策版本与文件效力（考点9—13）', 2)
for t in ['考点9：V1适用于2022年发货订单，V2适用于2023年发货订单，V3适用于2024年1月1日至2025年8月31日发货订单。',
          '考点10：V1、V2、V3订单数分别为168单、197单、355单，三者合计720单。',
          '考点11：2023年3月10日会议纪要属于讨论稿，不具有执行效力，不得据此调整费率或封顶。',
          '考点12：依据睿联销字〔2023〕07号正式发文，V2自2023年1月1日起追溯执行。',
          '考点13：2025年发货订单按V3计提，但渠道返利2025年度尚未到结算期，不得计入应付。']:
    P(doc, t, indent=False)
H(doc, '（四）V1清算（考点14—18）', 2)
for t in ['考点14：V1计提基数为含税发货额扣除退货，不得直接以订单含税额乘费率。',
          '考点15：V1以累计回款率达到90%作为发放门槛，168单中130单达标、38单挂账。',
          '考点16：38单挂账订单本期不发放、不结转以后期间，也不得计入应发合计。',
          '考点17：新客户首单适用4%费率，其他订单适用3%费率。',
          '考点18：V1按“人×年月”累计并执行10倍月基本工资封顶，超出部分不结转以后月份。']:
    P(doc, t, indent=False)
H(doc, '（五）V2清算（考点19—23）', 2)
for t in ['考点19：按回款和退款的事件日期排序逐笔累计V2佣金，不得按年度汇总后一次性套算。',
          '考点20：单笔累计回款500万元以内部分适用3.5%，超出部分适用5%。',
          '考点21：退款按3.5%冲减V2佣金，跨期退款回到原订单事件流，不得冲减其他订单。',
          '考点22：离职次月起发生的回款不再计提佣金。',
          '考点23：V2不设封顶，跨年回款仍归属原订单的政策版本，不得改按V3重算。']:
    P(doc, t, indent=False)
H(doc, '（六）V3清算（考点24—29）', 2)
for t in ['考点24：V3计提基数为“（含税回款－退款－坏账）÷1.13”的不含税净回款。',
          '考点25：直销订单适用3%、渠道订单适用1.5%的V3费率，不得混用V1或V2费率。',
          '考点26：坏账核销作为负事件逐单冲减V3净回款，不得在人员汇总层面统一扣减。',
          '考点27：2024年起合作单按主谈70%、协助30%拆分佣金，并在逐单明细中标注主谈人和协助人。',
          '考点28：V3佣金按个人80%、团队池20%拆分，团队池部分不得计入个人应发。',
          '考点29：V3按“人×年月”执行12倍月基本工资封顶，全期封顶扣减605.21万元、涉及39人。']:
    P(doc, t, indent=False)
H(doc, '（七）人员汇总与新人保底（考点30—33）', 2)
for t in ['考点30：按人员汇总V1、V2、V3三类应发，并叠加新人保底补差，形成逐人应发合计。',
          '考点31：入职前3个自然月的新人执行月佣金不足8000元补足规则，本期涉及10人、补差22.18万元。',
          '考点32：佣金已发台账仅作为实付比对依据，不得将已发金额直接作为应发金额。',
          '考点33：按应发与已发的差额划分应补、应追回和一致三类，本期为应补55人、应追回62人、一致3人，净超发273.13万元。']:
    P(doc, t, indent=False)
H(doc, '（八）团队池计提与发放（考点34—36）', 2)
for t in ['考点34：按团队和年度汇总V3个人佣金的20%计提团队池，本期计提358.08万元。',
          '考点35：团队年度目标达成率不低于100%方可发放，2024年3个团队达标、可发128.31万元。',
          '考点36：团队池单独列示，不得混入个人应发，也不得在新方案成本中重复计入。']:
    P(doc, t, indent=False)
H(doc, '（九）已发台账与异常预发（考点37—38）', 2)
for t in ['考点37：识别13条异常预发记录，区分按含税开票额预发、预发比例错误和待核实未发放三类情形。',
          '考点38：13条异常逐条给出应补、应追回或金额一致仅更正口径的处理结论，并说明发放口径偏差。']:
    P(doc, t, indent=False)
H(doc, '（十）渠道返利（考点39—42）', 2)
for t in ['考点39：按渠道和年度汇总含税净进货额，套用2%至5%的年度阶梯费率，形成40家乘4个年度共160行清算结果。',
          '考点40：5家设保底条款的渠道按阶梯返利与保底金额孰高确认，涉及15个渠道年度。',
          '考点41：2025年度40行渠道返利标注为未到结算期，不得计入本期应付。',
          '考点42：渠道返利与销售佣金是两套激励，渠道订单按1.5%计提佣金后不得再以返利重复计入佣金。']:
    P(doc, t, indent=False)
H(doc, '（十一）争议裁决与错误要求驳回（考点43—45）', 2)
for t in ['考点43：60件佣金争议逐单回到订单清算结果，不得按申诉金额全额支持或简单折中。',
          '考点44：裁决分为支持、部分支持和不支持三类，本期为支持2件、部分支持53件、不支持5件，并逐件引用条款。',
          '考点45：12项管理层错误要求逐条驳回，写明所违反的政策条款，包括V3不溯及、封顶不可取消和已发不等于应发等。']:
    P(doc, t, indent=False)
H(doc, '（十二）新方案设计与回测（考点46—48）', 2)
for t in ['考点46：新方案完整给出计提基数、个人与团队池比例、费率、封顶、保底、渠道返利和结算周期等10项参数。',
          '考点47：使用2024年全年订单回测新方案，新旧成本分别为1430.78万元和1295.47万元，增幅10.44%，不得突破15%控制线。',
          '考点48：完成9组“回款增长率×直销费率”敏感性测算，并逐人重算，不得只对总额做系数调整。']:
    P(doc, t, indent=False)
H(doc, '（十三）勾稽与核心结论（考点49—50）', 2)
for t in ['考点49：完成24项勾稽自检，覆盖订单数、人数、金额、封顶、团队池、渠道返利和新旧成本，全部闭合后方可出具结论。',
          '考点50：标准答案首表给出一句话核心结论，并确保该结论可由逐单明细、人员汇总和24项勾稽逐项回溯。']:
    P(doc, t, indent=False)

H(doc, '十、答题思路')
P(doc, '第一步是建立口径底稿。先固化去重、政策按发货日期归属、讨论稿无效、含税与不含税换算、退货坏账、已发不作应发、返利与佣金不重复七条口径，再开始计算。')
P(doc, '第二步是清洗订单主表。将732行订单去重为720单，并确认V1、V2、V3分别为168单、197单和355单，作为第一道自检。')
P(doc, '第三步是构建事件流。按订单号归集回款、退货和坏账，按事件日期排序，并统一截断至2025年8月31日。')
P(doc, '第四步是逐单计提。执行V1的90%回款率门槛、V2的500万元超额累进、V3的除税1.13与坏账扣减，并对合作单按70比30拆分。')
P(doc, '第五步是执行封顶与分配。按“人×年月”累计，V1执行10倍封顶、V3执行12倍封顶，超出部分不结转；V3同时按个人80%、团队池20%拆分。')
P(doc, '第六步是人员汇总与差额。叠加新人保底，形成应补55人、应追回62人、一致3人的三分结果，并单列13条异常预发。')
P(doc, '第七步是团队池与渠道返利。团队池按年度目标达成率判定发放或留存；渠道返利按年度阶梯与保底孰高，2025年度不计入应付。')
P(doc, '第八步是争议与错误要求。60件争议逐单回到清算结果，12项错误管理要求逐条引用条款驳回。')
P(doc, '第九步是新方案、敏感性与勾稽。使用2024年全年订单回测，完成9组敏感性，24项勾稽全部闭合后给出一句话核心结论。')

H(doc, '十一、标准答案')
P(doc, '标准答案共8个包，首表为“核心结论”，其余为逐单清算、人员汇总、团队池、已发异常、争议裁决、渠道返利、新方案参数、模拟测算、敏感性、勾稽自检、结论摘要和错误要求驳回。', indent=False)
P(doc, '核心结论：历史佣金整体超发273.13万元，渠道返利少付450.59万元，两项相抵后公司净少支出177.46万元。问题根源是手工台账按订单额粗算，未执行去重、封顶、退货、坏账和离职规则，建议逐人补追并在2025年四季度一次性清理。', bold=True)
P(doc, '支撑一：清算覆盖720个有效订单、120名销售、60件争议和40家渠道。佣金应发4402.42万元、已发4675.54万元，封顶扣减605.21万元。')
P(doc, '支撑二：应补545.57万元、涉及55人，应追回818.70万元、涉及62人，金额一致3人；13人已发记录存在预发口径异常。')
P(doc, '支撑三：渠道返利应付1268.23万元、已付817.65万元、应补450.59万元；2025年度未到结算期，不计入本期应付。')
P(doc, '支撑四：2026年新方案以不含税净回款为基数，个人70%、团队池30%，执行5倍封顶；2024年回测成本1430.78万元，较旧方案1295.47万元增长10.44%，低于15%控制线。')
P(doc, '关键结果：V1应发1063.88万元，V2应发1909.13万元，V3个人应发1407.22万元，新人保底22.18万元。团队池计提358.08万元，2024年可发128.31万元。争议裁决660.51万元，支持2件、部分支持53件、不支持5件。敏感性区间为1196.80万元至1682.31万元。24项勾稽全部闭合，不一致项为0。')
doc.core_properties.author = '董事会审计委员会'
doc.save(f'{OUT}/销售佣金清算_v1_题目.docx')
scrub(f'{OUT}/销售佣金清算_v1_题目.docx', WORD_APP, '董事会审计委员会', '2025-08-28T02:00:00Z', '2025-08-30T09:00:00Z', (2025, 8, 30, 9, 0, 0))

# ---------------- 打包 ----------------
INPUT_PACKS = [
    ('01组织与人员薪酬.rar', ['销售花名册', '团队目标与薪酬', '区域划分']),
    ('02订单主表.rar', ['订单主表']),
    ('03回款退货与坏账.rar', ['回款流水', '退货记录', '坏账核销']),
    ('04佣金政策三版.rar', ['佣金政策V1', '佣金政策V2', '佣金政策V3']),
    ('05政策发文与会议纪要.rar', ['政策变更通知', '经营会会议纪要']),
    ('06渠道协议与返利.rar', ['渠道商清单', '渠道返利阶梯', '渠道协议摘要', '渠道订单明细', '渠道返利已付台账']),
    ('07争议与申诉.rar', ['佣金争议台账', '申诉沟通记录']),
    ('08财务口径与管理要求.rar', ['财务确认规则', '税务与换算口径', '佣金已发台账', '管理层要求']),
]
ANSWER_PACKS = [
    ('01订单佣金清算.rar', ['订单佣金清算']),
    ('02人员汇总与异常核查.rar', ['人员佣金汇总', '团队池计提', '已发异常核查']),
    ('03争议裁决.rar', ['争议裁决', '争议裁决说明.docx']),
    ('04渠道返利台账.rar', ['渠道返利清算']),
    ('05新激励方案.rar', ['销售激励方案.docx', '新方案参数', '新旧成本对比.png']),
    ('06模拟测算与敏感性.rar', ['新方案模拟', '敏感性分析', '佣金结构.png']),
    ('07勾稽自检.rar', ['勾稽自检']),
    ('08结论与驳回.rar', ['核心结论', '结论摘要', '错误要求驳回', '清算结论.docx']),
]
for name, sheets in INPUT_PACKS:
    make_rar([f'{IN}/{s}.xlsx' for s in sheets], f'{OUT}/输入附件/{name}')
for name, sheets in ANSWER_PACKS:
    files = [f'{AN}/{s}' if s.endswith(('.docx', '.png')) else f'{AN}/{s}.xlsx' for s in sheets]
    make_rar(files, f'{OUT}/标准答案/{name}')

scrub(f'{BASE}/销售佣金清算_v1_题目.xlsx', EXCEL_APP, '销售运营部', '2025-08-25T02:10:00Z', '2025-08-27T07:35:00Z', (2025, 8, 27, 7, 35, 0))
scrub(f'{BASE}/销售佣金清算_v1_标准答案.xlsx', EXCEL_APP, '薪酬与激励项目组', '2025-09-03T01:20:00Z', '2025-09-05T09:05:00Z', (2025, 9, 5, 9, 5, 0))

print('=== 输入附件 ===')
for name, sheets in INPUT_PACKS:
    print(f'  {name}: ' + '、'.join(f'{s}({rows[s]}行)' for s in sheets))
print('=== 标准答案 ===')
for name, sheets in ANSWER_PACKS:
    print(f'  {name}: ' + '、'.join(s if s.endswith(('.docx', '.png')) else f'{s}({rows[s]}行)' for s in sheets))
print(f'输入 {len(INPUT_PACKS)}包；答案 {len(ANSWER_PACKS)}包')
