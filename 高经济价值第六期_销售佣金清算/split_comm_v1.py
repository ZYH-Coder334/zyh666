#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""销售佣金清算 v1 附件拆分与打包：6输入 + 6答案（含 docx/png）"""
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


def scrub(path, app_xml, author, created, modified):
    with zipfile.ZipFile(path) as zf:
        items = [(i, zf.read(i.filename)) for i in zf.infolist()]
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for info, data in items:
            if info.filename == 'docProps/app.xml':
                data = app_xml.encode('utf-8')
            elif info.filename == 'docProps/core.xml':
                data = core_xml(author, created, modified).encode('utf-8')
            zf.writestr(info, data)
    with open(path, 'wb') as f:
        f.write(buf.getvalue())


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


src_q = openpyxl.load_workbook(f'{BASE}/销售佣金清算_v1_题目.xlsx', data_only=True)
src_a = openpyxl.load_workbook(f'{BASE}/销售佣金清算_v1_标准答案.xlsx', data_only=True)
IN, AN = f'{SPLIT}/input', f'{SPLIT}/ans'
rows = {}
for sn in src_q.sheetnames:
    p = f'{IN}/{sn}.xlsx'
    rows[sn] = dump(src_q[sn], p)
    scrub(p, EXCEL_APP, '销售运营部', '2025-08-25T02:10:00Z', '2025-08-27T07:35:00Z')
for sn in src_a.sheetnames:
    p = f'{AN}/{sn}.xlsx'
    rows[sn] = dump(src_a[sn], p)
    scrub(p, EXCEL_APP, '薪酬与激励项目组', '2025-09-03T01:20:00Z', '2025-09-05T09:05:00Z')

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
scrub(f'{AN}/争议裁决说明.docx', WORD_APP, '薪酬与激励项目组', '2025-09-04T02:00:00Z', '2025-09-05T08:30:00Z')

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
scrub(f'{AN}/销售激励方案.docx', WORD_APP, '薪酬与激励项目组', '2025-09-02T03:10:00Z', '2025-09-05T06:20:00Z')

# 3 清算结论
doc = new_doc()
H(doc, '销售佣金与渠道返利历史清算结论')
P(doc, '一、总结论')
P(doc, '经对2022年1月1日至2025年8月31日全部720个有效订单逐单复核，历史佣金应发 4402.42 万元，财务已发 4675.54 万元，整体超发 273.13 万元；其中应补 545.57 万元、应追回 818.70 万元，涉及 117 人。渠道返利应付 1268.23 万元、已付 817.65 万元，应补 450.59 万元。')
H(doc, '二、主要发现', 2)
P(doc, '1. 订单主表存在12条重复导入记录，涉及金额约 1.1 亿元，已按首次导入记录去重，避免重复计提。')
P(doc, '2. V1政策下38个订单因回款率不足90%应挂账，已发台账却已预发，形成超发。')
P(doc, '3. 财务预发按“订单额×3%×70%”粗算，未执行月度封顶、未扣退货与坏账，是超发主因；封顶扣减合计605.21万元。')
P(doc, '4. 2023年V2政策被误按4月1日执行的说法存在于会议纪要，按正式发文应自1月1日追溯，影响金额已重算。')
P(doc, '5. 2024年起30个合作单未按7:3拆分，全部计入签约人，涉及佣金已调整。')
P(doc, '6. 渠道返利按含税进货额与阶梯比例重算，5家保底渠道按“阶梯与保底孰高”执行；2025年度未到结算期，未计入应付。')
H(doc, '三、争议与制度', 2)
P(doc, '60件争议中支持2件、部分支持53件、不支持5件，裁决金额660.51万元，核减914.56万元。管理层提出的12项要求均与正式制度冲突，已逐条驳回（详见《勾稽自检》与《错误要求驳回说明》）。')
H(doc, '四、新方案建议', 2)
P(doc, '建议2026年执行“个人70%+团队池30%、净回款计提、两档加速、5倍封顶”的新方案，2024年回测成本增幅10.44%，在15%控制线内；同时将历史清算差额在2025年四季度一次性处理，避免跨年争议。')
H(doc, '五、勾稽结论', 2)
P(doc, '20项勾稽全部通过：订单数、三版政策订单数、逐单与人员汇总、已发台账、争议三类、渠道返利、新方案个人与团队拆分均闭合。')
doc.core_properties.author = '薪酬与激励项目组'
doc.save(f'{AN}/清算结论.docx')
scrub(f'{AN}/清算结论.docx', WORD_APP, '薪酬与激励项目组', '2025-09-05T01:30:00Z', '2025-09-05T10:00:00Z')

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

# ---------------- 题目 docx ----------------
doc = new_doc()
H(doc, '睿联工业设备有限公司销售佣金与渠道返利清算')
P(doc, '一、任务背景')
P(doc, '睿联工业设备有限公司（虚构）是一家工业泵阀制造企业，销售采用直销与渠道并行模式。2025年7月，公司发现历史佣金发放与渠道返利长期依赖手工台账，销售人员对金额争议集中爆发。你受董事会审计委员会委托，作为外部薪酬与销售运营顾问，对2022年1月1日至2025年8月31日的销售佣金、渠道返利进行全量清算，并设计2026年激励方案。')
P(doc, '数据截止日为2025年8月31日，清算基准日为2025年9月1日。所有结论须可回溯到附件逐条记录。')
H(doc, '二、输入附件')
table(doc, ['包名', '文件', '内容要点'],
      [['01组织与人员薪酬.rar', '销售花名册、团队目标与薪酬、区域划分', '120名销售、6个团队、入职离职与月度目标'],
       ['02订单与回款流水.rar', '订单主表、回款流水、退货记录、坏账核销', '732行订单、1327笔回款、70笔退货、49笔坏账'],
       ['03佣金政策三版.rar', 'V1/V2/V3、政策变更通知、经营会纪要', '三版政策条款、生效文件与讨论稿'],
       ['04渠道协议与返利.rar', '渠道商清单、返利阶梯、协议摘要、渠道订单、返利已付', '40家渠道、189个渠道订单'],
       ['05争议与申诉台账.rar', '佣金争议台账、申诉沟通记录', '60件争议及申诉理由'],
       ['06财务口径与管理层要求.rar', '财务确认规则、税务口径、佣金已发台账、管理层要求', '去重与口径规则、480条已发记录、12项要求']],
      [4.6, 6.4, 5.2])
H(doc, '三、制度要点')
P(doc, 'V1（2022年发货订单）：含税发货额扣退货为基数，标准3%、新客户首单4%，累计回款率≥90%方可发放，月度封顶10倍月基本工资。')
P(doc, 'V2（2023年发货订单）：含税回款扣退款为基数，3.5%，单笔累计回款超500万元部分5%，无封顶。')
P(doc, 'V3（2024年1月1日至2025年8月31日发货订单）：不含税净回款为基数，直销3%、渠道单1.5%，个人80%、团队池20%，合作单7:3，月度封顶12倍月基本工资，新人前3个月保底8000元。')
P(doc, '渠道返利：按年度含税进货额（发货额扣退货）套阶梯，2%-5%，有保底条款的按阶梯与保底孰高；次年1月结算。')
H(doc, '四、交付物')
table(doc, ['序号', '交付包', '内容与格式'],
      [['01', '逐单清算与人员汇总.rar', '订单佣金清算（720行）、人员佣金汇总（120行）、团队池计提（12行）'],
       ['02', '争议裁决.rar', '争议裁决（60行）+ 裁决说明.docx'],
       ['03', '渠道返利台账.rar', '渠道返利清算（160行，40家×4个年度）'],
       ['04', '新激励方案.rar', '销售激励方案.docx + 新方案参数表'],
       ['05', '模拟测算与敏感性.rar', '新方案模拟（120行）、敏感性分析（9组）、成本对比图'],
       ['06', '勾稽自检与结论.rar', '勾稽自检（20项）、结论摘要、清算结论.docx']],
      [1.6, 5.4, 8.8])
H(doc, '五、约束条件')
for t in ['1. 仅使用附件材料，不得引入外部数据、不得联网检索。',
          '2. 所有金额为人民币元，保留两位小数；比例保留两位。',
          '3. 逐单、逐人给出结论，不得抽样或估算；每条记录必须有状态。',
          '4. 订单去重、政策归属、文件效力、金额口径以《财务确认规则》为准。',
          '5. 政策适用以订单发货日期为准，不按订单日期或回款日期。',
          '6. 退货、坏账、离职、封顶、合作单拆分必须逐单体现。',
          '7. 每个争议裁决、每条管理层要求驳回均须给出依据条款。',
          '8. 六个交付包之间的订单数、人数、金额必须相互勾稽。',
          '9. 不得修改原始附件；交付文件按包名与工作表命名。',
          '10. 新方案须在旧方案基础上测算成本，并给出9组敏感性结果。']:
    P(doc, t, indent=False)
H(doc, '六、完成标准')
P(doc, '全部720个有效订单均有清算结论；120名销售均有应发、已发、差额；60件争议均有裁决；40家渠道的2022-2024年度返利均清算完毕；20项勾稽全部通过；新方案回测成本增幅不超过15%。')
doc.core_properties.author = '董事会审计委员会'
doc.save(f'{OUT}/销售佣金清算_v1_题目.docx')
scrub(f'{OUT}/销售佣金清算_v1_题目.docx', WORD_APP, '董事会审计委员会', '2025-08-28T02:00:00Z', '2025-08-30T09:00:00Z')

# ---------------- 打包 ----------------
INPUT_PACKS = [
    ('01组织与人员薪酬.rar', ['销售花名册', '团队目标与薪酬', '区域划分']),
    ('02订单与回款流水.rar', ['订单主表', '回款流水', '退货记录', '坏账核销']),
    ('03佣金政策三版.rar', ['佣金政策V1', '佣金政策V2', '佣金政策V3', '政策变更通知', '经营会会议纪要']),
    ('04渠道协议与返利.rar', ['渠道商清单', '渠道返利阶梯', '渠道协议摘要', '渠道订单明细', '渠道返利已付台账']),
    ('05争议与申诉台账.rar', ['佣金争议台账', '申诉沟通记录']),
    ('06财务口径与管理层要求.rar', ['财务确认规则', '税务与换算口径', '佣金已发台账', '管理层要求']),
]
ANSWER_PACKS = [
    ('01逐单清算与人员汇总.rar', ['订单佣金清算', '人员佣金汇总', '团队池计提']),
    ('02争议裁决.rar', ['争议裁决', '争议裁决说明.docx']),
    ('03渠道返利台账.rar', ['渠道返利清算']),
    ('04新激励方案.rar', ['新方案参数', '销售激励方案.docx', '新旧成本对比.png']),
    ('05模拟测算与敏感性.rar', ['新方案模拟', '敏感性分析', '佣金结构.png']),
    ('06勾稽自检与结论.rar', ['勾稽自检', '错误要求驳回', '结论摘要', '清算结论.docx']),
]
for name, sheets in INPUT_PACKS:
    make_rar([f'{IN}/{s}.xlsx' for s in sheets], f'{OUT}/输入附件/{name}')
for name, sheets in ANSWER_PACKS:
    files = [f'{AN}/{s}' if s.endswith(('.docx', '.png')) else f'{AN}/{s}.xlsx' for s in sheets]
    make_rar(files, f'{OUT}/标准答案/{name}')

print('=== 输入附件 ===')
for name, sheets in INPUT_PACKS:
    print(f'  {name}: ' + '、'.join(f'{s}({rows[s]}行)' for s in sheets))
print('=== 标准答案 ===')
for name, sheets in ANSWER_PACKS:
    print(f'  {name}: ' + '、'.join(s if s.endswith(('.docx', '.png')) else f'{s}({rows[s]}行)' for s in sheets))
print(f'输入 {len(INPUT_PACKS)}包；答案 {len(ANSWER_PACKS)}包')
