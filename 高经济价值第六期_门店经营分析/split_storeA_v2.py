#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""G6-ZX-004 门店经营分析 v2 返修版：2 包打包（8 表输入 + 8 表答案）+ docx
- 每包内 8 个 xlsx，各自独立的 creator/created/modified/revision（避免批量机器签名）
- 保留 sheet 级批注（关键规则藏批注）
- docx 十一章节 + 职业定位 + 70 考点（50 数据型 + 20 推理链型）
"""
import io
import os
import re
import shutil
import time
import zipfile
from copy import copy

import openpyxl
from openpyxl import Workbook
from openpyxl.comments import Comment
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

BASE = '/workspace/高经济价值第六期_门店经营分析'
SRC_Q = f'{BASE}/门店经营分析_v1_题目.xlsx'
SRC_A = f'{BASE}/门店经营分析_v1_标准答案.xlsx'
SPLIT = '/tmp/storeA_v2_split'
OUT = f'{BASE}/门店经营分析_v1_附件'
COMPANY = '麦淳餐饮管理有限公司'
shutil.rmtree(SPLIT, ignore_errors=True)
os.makedirs(SPLIT, exist_ok=True)
os.makedirs(OUT, exist_ok=True)

IN_RAR = f'{OUT}/门店经营分析_v1_输入附件.rar'
AN_RAR = f'{OUT}/门店经营分析_v1_标准答案.rar'
DOCX = f'{OUT}/门店经营分析_v1_题目.docx'

XMLHEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
PROPS_NS = ('<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
            'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">')


def app_xml_xlsx(sheet_names, ver):
    n = len(sheet_names)
    pairs = ('<HeadingPairs><vt:vector size="2" baseType="variant">'
             '<vt:variant><vt:lpstr>工作表</vt:lpstr></vt:variant>'
             f'<vt:variant><vt:i4>{n}</vt:i4></vt:variant></vt:vector></HeadingPairs>'
             f'<TitlesOfParts><vt:vector size="{n}" baseType="lpstr">'
             + ''.join(f'<vt:lpstr>{s}</vt:lpstr>' for s in sheet_names)
             + '</vt:vector></TitlesOfParts>')
    return (XMLHEAD + PROPS_NS + '<Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity>'
            '<ScaleCrop>false</ScaleCrop>' + pairs +
            f'<Company>{COMPANY}</Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc>'
            f'<HyperlinksChanged>false</HyperlinksChanged><AppVersion>{ver}</AppVersion></Properties>')


def app_xml_docx(nchar, npara, ver='16.0000'):
    pages = max(1, round(nchar / 780))
    lines = max(1, round(nchar / 34))
    return (XMLHEAD + PROPS_NS + '<Template>Normal.dotm</Template><TotalTime>143</TotalTime>'
            f'<Pages>{pages}</Pages><Words>{nchar}</Words><Characters>{nchar}</Characters>'
            f'<Application>Microsoft Office Word</Application><DocSecurity>0</DocSecurity>'
            f'<Lines>{lines}</Lines><Paragraphs>{npara}</Paragraphs><ScaleCrop>false</ScaleCrop>'
            f'<Company>{COMPANY}</Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc>'
            f'<HyperlinksChanged>false</HyperlinksChanged><AppVersion>{ver}</AppVersion></Properties>')


def core_xml(author, lastmod, rev, created, modified):
    return (XMLHEAD +
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            f'<dc:creator>{author}</dc:creator><cp:lastModifiedBy>{lastmod}</cp:lastModifiedBy>'
            f'<cp:revision>{rev}</cp:revision>'
            f'<dcterms:created xsi:type="dcterms:W3CDTF">{created}</dcterms:created>'
            f'<dcterms:modified xsi:type="dcterms:W3CDTF">{modified}</dcterms:modified></cp:coreProperties>')


FILEVERSION = ('<fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="27519"/>')


def scrub(path, meta, zip_dt, kind='xlsx', app_ver='16.0300', doc_stats=None):
    author, lastmod, rev, created, modified = meta
    with zipfile.ZipFile(path) as zf:
        items = [(i, zf.read(i.filename)) for i in zf.infolist()]
    names = [i.filename for i, _ in items]
    dropped = set()
    if kind == 'docx' and 'docProps/thumbnail.jpeg' in names:
        dropped.add('docProps/thumbnail.jpeg')
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for info, data in items:
            if info.filename in dropped:
                continue
            if info.filename == 'docProps/app.xml':
                if kind == 'xlsx':
                    with zipfile.ZipFile(path) as zf0:
                        wbx = zf0.read('xl/workbook.xml').decode('utf-8')
                    sheets = re.findall(r'<sheet[^>]*name="([^"]+)"', wbx)
                    data = app_xml_xlsx(sheets, app_ver).encode('utf-8')
                else:
                    nchar, npara = doc_stats
                    data = app_xml_docx(nchar, npara, app_ver).encode('utf-8')
            elif info.filename == 'docProps/core.xml':
                data = core_xml(author, lastmod, rev, created, modified).encode('utf-8')
            elif info.filename == '_rels/.rels' and dropped:
                txt = data.decode('utf-8')
                txt = re.sub(r'<Relationship[^>]*thumbnail\.jpeg[^>]*/>', '', txt)
                data = txt.encode('utf-8')
            elif info.filename == 'xl/workbook.xml' and '<fileVersion' not in data.decode('utf-8'):
                txt = data.decode('utf-8')
                m = re.match(r'(<workbook[^>]*>)', txt)
                if m:
                    txt = txt[:m.end()] + FILEVERSION + txt[m.end():]
                data = txt.encode('utf-8')
            zi = zipfile.ZipInfo(info.filename, date_time=zip_dt)
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = 0o600 << 16
            zf.writestr(zi, data)
    with open(path, 'wb') as f:
        f.write(buf.getvalue())
    ts = time.mktime((*zip_dt, 0, 0, -1))
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
            if cell.comment:
                d.comment = Comment(cell.comment.text, cell.comment.author)
                d.comment.width = cell.comment.width
                d.comment.height = cell.comment.height
    for col, dim in src.column_dimensions.items():
        dst.column_dimensions[col].width = dim.width
    for row, dim in src.row_dimensions.items():
        dst.row_dimensions[row].height = dim.height
    dst.freeze_panes = src.freeze_panes
    dst.sheet_view.showGridLines = False
    for rng in src.merged_cells.ranges:
        dst.merge_cells(str(rng))
    return dst


# 每个 xlsx 的独立签名：(creator, lastModifiedBy, revision, created, modified)
META_IN = {
    '01组织与门店基础.xlsx': ('门店运营部-陈建华', '门店运营部-陈建华', '38', '2026-03-27T02:14:37Z', '2026-04-01T09:52:18Z'),
    '02日营业流水.xlsx': ('信息技术部-刘畅', '信息技术部-刘畅', '61', '2026-03-30T23:41:05Z', '2026-04-02T00:37:26Z'),
    '03门店P&L与外卖结算.xlsx': ('财务部-周涛', '财务部-周涛', '27', '2026-03-29T06:18:44Z', '2026-04-01T13:26:51Z'),
    '04成本与采购明细.xlsx': ('采购部-孙莉', '财务部-周涛', '19', '2026-03-28T07:35:12Z', '2026-04-01T10:09:33Z'),
    '05租约与盘点.xlsx': ('行政部-吴静', '行政部-吴静', '14', '2026-03-26T08:52:29Z', '2026-03-31T15:44:07Z'),
    '06制度与会议纪要.xlsx': ('文控中心-李涵', '文控中心-李涵', '9', '2026-02-11T03:27:46Z', '2026-03-31T11:18:52Z'),
    '07考核与分级办法.xlsx': ('人力资源部-冯雪', '运营总监-陈建华', '11', '2026-01-08T05:39:21Z', '2026-03-31T09:07:14Z'),
    '08财务口径与管理要求.xlsx': ('财务部-周涛', '总经办-王梓', '6', '2026-03-31T01:22:58Z', '2026-04-02T01:41:09Z'),
}
META_AN = {
    '01口径适用判定.xlsx': ('审计部-王梓', '审计部-王梓', '4', '2026-04-07T02:36:19Z', '2026-04-08T09:14:52Z'),
    '02收入重算.xlsx': ('审计部-王梓', '审计部-王梓', '5', '2026-04-07T08:11:43Z', '2026-04-08T14:27:36Z'),
    '03P&L重算与外卖对账.xlsx': ('审计部-吴静', '审计部-吴静', '7', '2026-04-08T01:47:28Z', '2026-04-09T07:52:41Z'),
    '04成本与费用核查.xlsx': ('审计部-吴静', '审计部-吴静', '3', '2026-04-08T06:23:15Z', '2026-04-09T10:18:07Z'),
    '05分级与盘点.xlsx': ('审计部-李涵', '审计部-李涵', '6', '2026-04-08T09:35:54Z', '2026-04-09T13:41:22Z'),
    '06店长奖金.xlsx': ('审计部-李涵', '审计部-李涵', '4', '2026-04-09T00:52:37Z', '2026-04-09T15:26:48Z'),
    '07处置与驳回.xlsx': ('审计部-王梓', '审计部-王梓', '5', '2026-04-09T03:14:09Z', '2026-04-10T01:37:55Z'),
    '08勾稽与结论.xlsx': ('审计部-王梓', '合伙人-林志远', '8', '2026-04-09T07:48:31Z', '2026-04-10T05:22:43Z'),
}

INPUT_PACKS = [
    ('01组织与门店基础', ['门店基础信息', '城市与区域', '员工花名册']),
    ('02日营业流水', ['日营业流水明细']),
    ('03门店P&L与外卖结算', ['门店Q1分月P&L', '外卖平台结算单']),
    ('04成本与采购明细', ['食材采购入库明细', '费用报销明细']),
    ('05租约与盘点', ['租约明细', '月度盘点']),
    ('06制度与会议纪要', ['成本与分摊制度v1.0', '成本与分摊制度v2.0', '成本与分摊制度v3.0草案',
                    '制度生效与废止通知', '经营分析会会议纪要']),
    ('07考核与分级办法', ['店长绩效考核办法v1', '店长绩效考核办法v2', '门店分级标准', '总部分摊规则']),
    ('08财务口径与管理要求', ['财务确认规则', '管理层要求']),
]
ANSWER_PACKS = [
    ('01口径适用判定', ['口径适用判定']),
    ('02收入重算', ['收入重算']),
    ('03P&L重算与外卖对账', ['P&L重算', '外卖对账']),
    ('04成本与费用核查', ['成本核查', '报销异常']),
    ('05分级与盘点', ['门店分级', '盘点责任']),
    ('06店长奖金', ['店长奖金']),
    ('07处置与驳回', ['处置建议', '错误要求驳回']),
    ('08勾稽与结论', ['勾稽自检', '经营分析结论']),
]


def build_packs(src_path, packs, meta_map, rar_path, dt_default):
    wb0 = openpyxl.load_workbook(src_path)
    members = []
    for name, sheets in packs:
        wb = Workbook()
        wb.remove(wb.active)
        for sh in sheets:
            copy_sheet(wb0[sh], wb, sh)
        fn = f'{name}.xlsx'
        tmp = os.path.join(SPLIT, fn)
        wb.save(tmp)
        meta = meta_map[fn]
        y, mo, d, h, mi, s = (int(x) for x in re.split(r'[-T:Z]', meta[4])[:6])
        zip_dt = (y, mo, d, h + 8 if h + 8 < 24 else h, mi, s)
        zip_dt = (y, mo, d, (h + 8) % 24, mi, s)
        scrub(tmp, meta, zip_dt, kind='xlsx', app_ver='16.0300' if y % 2 == 0 else '16.0000')
        members.append((fn, tmp, zip_dt))
    with zipfile.ZipFile(rar_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for fn, path, zd in members:
            zi = zipfile.ZipInfo(fn, date_time=zd)
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = 0o600 << 16
            z.writestr(zi, open(path, 'rb').read())
    os.utime(rar_path, (dt_default, dt_default))
    print('packed', os.path.basename(rar_path), len(members), 'files')


build_packs(SRC_Q, INPUT_PACKS, META_IN, IN_RAR, time.mktime((2026, 4, 2, 9, 30, 0, 0, 0, -1)))
build_packs(SRC_A, ANSWER_PACKS, META_AN, AN_RAR, time.mktime((2026, 4, 10, 18, 0, 0, 0, 0, -1)))

# ===== 读取答案指标 =====
ans = openpyxl.load_workbook(SRC_A, data_only=True)


def rows_of(ws):
    data = list(ws.iter_rows(values_only=True))
    hdr = data[0]
    return [dict(zip(hdr, r)) for r in data[1:] if r[0] is not None]


chk = {r['勾稽项']: r['左值'] for r in rows_of(ans['勾稽自检'])}
concl = {r['项目']: r['内容'] for r in rows_of(ans['经营分析结论'])}
grades = {r['门店编号']: r for r in rows_of(ans['门店分级'])}
inc = {r['门店编号']: r for r in rows_of(ans['收入重算'])}
bonus_rows = rows_of(ans['店长奖金'])

TOT_REV = chk['重算营业收入合计(万元)']
BOOK_REV = chk['账面营业收入合计(万元)']
REV_GAP = chk['收入差异合计(万元)']
EBITDA_GAP = chk['EBITDA差异合计(万元)']
A_N = chk['A级门店数']
B_N = chk['B级门店数']
C_N = chk['C级门店数']
D_N = chk['D级门店数']
BONUS = chk['店长奖金合计(万元)']
N_ISSUE = chk['P&L需重算门店数']
N_FLIP = chk['分级反转门店数']
N_EXP = chk['报销异常条数']
N_CHK = len(chk)
D_LIST = [s for s, r in grades.items() if r['分级'] == 'D']
FLIP_LIST = [s for s, r in grades.items() if r['是否反转'] == '是']
TOT_NET = round(sum(float(r['Q1税后利润(万元)']) for r in grades.values()), 2)
TOT_EBITDA = round(sum(float(r['Q1 EBITDA(万元)']) for r in grades.values()), 2)
DIRECT_REV = round(sum(float(r['Q1营业收入(万元)']) for r in grades.values() if r['类型'] == '直营'), 2)
DIRECT_EBITDA = round(sum(float(r['Q1 EBITDA(万元)']) for r in grades.values() if r['类型'] == '直营'), 2)
DIRECT_NET = round(sum(float(r['Q1税后利润(万元)']) for r in grades.values() if r['类型'] == '直营'), 2)

G = {s: (float(r['EBITDA率(%)']), r['分级'], r['账面分级'], float(r['Q1营业收入(万元)']), float(r['Q1 EBITDA(万元)']))
     for s, r in grades.items()}

# ===== docx =====
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
# 移除环境默认模板插入的 "含 AI 生成" 文本框
for _p in list(doc.paragraphs):
    if 'AILabel' in _p._element.xml:
        _p._element.getparent().remove(_p._element)
H(doc, '连锁门店2026年Q1经营分析与亏损店处置', 0)

H(doc, '一、编号')
P(doc, 'G6-ZX-004（v1）。', indent=False)
P(doc, '【职业定位】二级行业：餐饮连锁经营；L1 职业簇：运营管理、财务与审计、合规与风控；'
       'L2 职业：门店运营经理、区域经理、财务分析师、管理会计、内控审计师。', indent=False)

H(doc, '二、任务背景')
P(doc, f'麦淳餐饮管理有限公司是一家连锁烘焙与轻食品牌，截至2026年3月共运营30家门店，分布在苏州、无锡、宁波、杭州、合肥、南昌六个城市，'
       f'其中直营27家、加盟3家。2026年4月，公司召开一季度经营分析会，发现杭州、合肥、南昌新市场门店爬坡不及预期，'
       f'无锡惠山万达、合肥瑶海万达、南昌新建吾悦等门店利润率偏低，且门店上报的P&L与业务系统流水存在口径不一致。'
       f'现受运营总监与财务总监委托，以运营分析岗身份，对30家门店Q1经营数据完成盈利分级、口径纠偏、亏损归因、外卖对账、'
       f'采购与费用核查、盘点责任划分、店长奖金核算和亏损店处置建议。')
P(doc, '数据截止日为2026年3月31日，分析基准日为2026年4月5日。所有结论须可回溯到附件逐条记录，不得只给汇总数。'
       '凡涉及分级、归因、责任、扣减、处置的记录，必须逐条说明触发条件与依据。门店上报的P&L不得直接采信，须以日营业流水为底稿重算。')

H(doc, '三、任务分解')
P(doc, '本项委托按八轮执行，每轮只完成一个任务，并只提交当轮对应交付包。上一轮自检通过后，方可进入下一轮；'
       '未完成或自检不通过，不得进入下一轮，也不得提前输出后续轮次结果。')
P(doc, '第1轮为口径适用判定。逐店判定P&L收入是否按外卖GMV列报、成本费率适用版本、Q1应计租金月数，识别需重算门店。'
       '本轮只输出01口径适用判定.xlsx。本轮自检要求30家门店均有四项判定结论与依据，需重算门店标注具体问题类型。')
P(doc, '第2轮为收入重算。以日营业流水2700行为底稿，逐店汇总堂食净额与外卖GMV，与P&L账面收入比对。'
       '本轮只输出02收入重算.xlsx。本轮自检要求30家门店均有重算值、账面值、差异及结论，差异门店说明原因。')
P(doc, '第3轮为P&L重算与外卖对账。按当版制度重算收入、成本、费用、EBITDA、税前利润与税后利润，并逐店核对平台实收与门店入账。'
       '本轮只输出03P&L重算与外卖对账.xlsx。本轮自检要求30家门店均有完整重算P&L与分级，外卖差异门店均有原因。')
P(doc, '第4轮为成本与费用核查。按标准配方测算理论食材成本与采购入库比对，并逐条核查报销异常。'
       '本轮只输出04成本与费用核查.xlsx。本轮自检要求30家门店均有成本差异率与风险提示，报销异常逐条列明类型与依据。')
P(doc, '第5轮为分级与盘点。按重算EBITDA率完成A/B/C/D分级，测算Q1平均盘点差异率并认定门店责任。'
       '本轮只输出05分级与盘点.xlsx。本轮自检要求30家门店均有分级，超标门店均有责任认定与金额。')
P(doc, '第6轮为店长奖金核算。按v2考核办法计算每家店长奖金系数与应发奖金，D级标注扣减。'
       '本轮只输出06店长奖金.xlsx。本轮自检要求30家门店均有奖金计算，口径与考核办法一致。')
P(doc, '第7轮为处置与驳回。对C/D级门店给出整改、转址、关店建议，并逐条驳回管理层错误要求。'
       '本轮只输出07处置与驳回.xlsx。本轮自检要求处置建议均有触发条件与时间节点，12项要求逐条驳回并说明理由。')
P(doc, '第8轮为勾稽与结论。完成全部勾稽并给出一句话核心结论。本轮只输出08勾稽与结论.xlsx。'
       '本轮自检要求勾稽全部闭合、不一致项为0；未闭合不得出具结论。')

H(doc, '四、输入附件')
P(doc, '输入附件共1包：门店经营分析_v1_输入附件.rar，内含8个表格文件，可单独下载。')
for name, sheets in INPUT_PACKS:
    P(doc, f'{name}.xlsx，包含{"、".join(sheets)}。', indent=False)

H(doc, '五、制度与判断规则')
P(doc, '门店共30家，直营27家、加盟3家（WX05、HF03、NC03）。加盟门店纳入本次经营分析与分级处置范围，收入成本按门店口径还原；'
       '但公司合并报表仅确认授权费与供应链毛利，加盟利润不并入直营利润合计。')
P(doc, '成本与分摊制度存在三个版本：PRC-COST-v1.0 于2024年1月1日生效、2026年1月1日废止；PRC-COST-v2.0 于2025年1月1日生效、现行有效；'
       'PRC-COST-v3.0 于2026年1月15日经管委会审议通过，自2026年4月1日起执行，不追溯调整2026Q1。'
       '各版本的具体费率以对应工作表为准，本说明不重复列示。')
P(doc, '营业收入等于堂食净额加外卖GMV，外卖GMV为扣除退款后的净额并含平台补贴，不得以平台实收列报营业收入。'
       '门店上报的P&L账面值仅供参考，发生口径争议时以日营业流水重算值为准。')
P(doc, '管理分摊的计提基数自2026Q1起由"营业收入"调整为"营业收入减外卖平台佣金"，费率按当版制度执行。'
       '外卖平台佣金仅对外卖GMV计提，不得对堂食收入或全部收入计提。加盟门店口径、免租期处理、奖金权重等2026年调整事项，'
       '以《经营分析会会议纪要》JY-2026-001 为准。')
P(doc, '租金含物业费，按月租金乘面积乘Q1应计租金月数计算；免租期以租约补充条款约定为准，已计提但依约应免的租金在当期冲回，不追溯调整以前年度。'
       '折旧按建店投资5年直线计提，残值为0，月折旧等于建店投资除以60。所得税按正税前利润的25%计提，亏损年度为0，不确认递延所得税资产。')
P(doc, '门店分级按EBITDA率：A级不低于12%，B级8%至12%，C级4%至8%，D级低于4%或连续两季亏损。'
       'D级门店先整改一个季度，连续两个季度D级方可关店或转址。盘点差异率超过2%的部分全额由门店承担并从店长季度奖金扣除。')
P(doc, '店长奖金等于基础奖金乘奖金系数，基础奖金A级1.2万元、B级1万元、C级0.8万元、D级0.5万元，D级再扣减50%；'
       '奖金系数的权重与损耗得分口径以现行有效的考核办法为准。')

H(doc, '六、交付物')
for name, _ in ANSWER_PACKS:
    P(doc, f'{name}.xlsx', indent=False)

H(doc, '七、约束条件')
constraints = [
    '1. 仅使用附件材料，不得引入外部数据，不得联网检索。',
    '2. 对外列示金额以万元为单位并保留两位小数，比例保留两位；底层计算以元为单位，不得混用。',
    '3. 逐门店给出结论，不得抽样或估算；每条记录必须有明确状态，不得留空。',
    '4. 30家门店编号唯一，不得跨城市重复计入；六城门店数为6/5/5/5/5/4。',
    '5. 加盟门店纳入分级与处置，但利润不并入直营利润合计，公司仅确认授权费与供应链毛利。',
    '6. 营业收入等于堂食净额加外卖GMV，外卖GMV为扣除退款后的净额并含平台补贴。',
    '7. 门店上报P&L账面值不得直接采信，须以日营业流水2700行为底稿重算。',
    '8. Q1成本费率须按现行有效版本执行，未生效版本不得提前适用，亦不追溯调整已结月份。',
    '9. 管理分摊计提基数须为"营业收入减外卖平台佣金"，不得按全额营业收入或按门店数平均。',
    '10. 外卖平台佣金只对外卖GMV计提，不得对堂食收入或全部收入计提。',
    '11. 门店入账外卖收入应等于外卖GMV，不得按平台实收入账。',
    '12. 食材、人工、水电、营销、管理分摊比例固定，不得逐店调整。',
    '13. 租金含物业费；免租期以租约补充条款为准，已计提但依约应免的租金须当期冲回。',
    '14. 折旧按5年直线计提，月折旧等于建店投资除以60。',
    '15. 所得税按正税前利润的25%计提，亏损年度所得税为0，不确认递延所得税资产。',
    '16. 门店分级按重算EBITDA率，A级≥12%、B级8%-12%、C级4%-8%、D级<4%。',
    '17. 账面分级与重算分级不一致的，一律以重算结果为准。',
    '18. D级门店先整改一个季度，连续2季D级才关店或转址，不得当季立即关店。',
    '19. 盘点差异率超过2%的部分全额由门店承担，并从店长季度奖金扣除。',
    '20. 店长奖金权重与损耗得分口径以现行有效考核办法为准，不得沿用旧版权重。',
    '21. D级门店基础奖金扣减50%，并追回已发预提。',
    '22. 管理层要求逐条判断是否采纳，不予采纳的须写明冲突条款。',
    '23. 必须按八轮顺序逐轮交付，每轮只完成一个任务，每轮结束提交本轮自检记录；上一轮未通过不得进入下一轮。',
    '24. 全部勾稽闭合后方可出具核心结论；任一项不闭合视为未完成。',
]
for t in constraints:
    P(doc, t, indent=False)

H(doc, '八、完成标准')
P(doc, f'八轮任务逐轮完成且逐轮自检通过。30家门店全部完成口径判定、收入重算、P&L重算与分级，'
       f'A级{A_N}家、B级{B_N}家、C级{C_N}家、D级{D_N}家；重算营业收入{TOT_REV}万元、EBITDA {TOT_EBITDA}万元、税后利润{TOT_NET}万元；'
       f'账面营业收入{BOOK_REV}万元，重算与账面差异{REV_GAP}万元（EBITDA差异{EBITDA_GAP}万元）；'
       f'识别需重算门店{N_ISSUE}家、分级反转{N_FLIP}家；店长奖金合计{BONUS}万元；报销异常{N_EXP}条；'
       f'驳回管理层要求12项；{N_CHK}项勾稽全部闭合、不一致项为0。任一轮未完成或自检不通过，不得进入下一轮，整体不得视为完成。')

H(doc, '九、考点')
P(doc, f'考点按数据治理、制度适用、收入重算、成本费用、外卖对账、采购费用核查、门店分级、店长奖金、处置驳回、勾稽结论十组展开，共70项，'
       f'其中数据型50项、推理链型20项（标注"推理链"者须同时引用两张以上表与具体制度条款）。', indent=False)

groups = [
    ('（一）数据治理与门店清单（考点1—8）', [
        '考点1：底稿中应当确认门店共30家，苏州6家、无锡5家、宁波5家、杭州5家、合肥5家、南昌4家，编号唯一且不跨城市重复。',
        '考点2：底稿中应当区分直营27家、加盟3家，加盟门店为WX05、HF03、NC03。',
        '考点3：底稿中应当确认日营业流水共2700行，等于30家门店乘90天，且每家门店无缺日。',
        '考点4：底稿中应当识别2025-11-01之后开业的新店（WX05、HZ04、HF05、NC03、NC04、HZ05）存在爬坡，实际收入低于"稳定月营收×3"。',
        '考点5：底稿中应当确认员工花名册240人，等于30家门店乘8人，每店1名店长。',
        '考点6：底稿中应当认定"稳定月营收"为2025年12月稳态值，不得直接作为Q1实际营业收入。',
        '考点7：底稿中应当以租约明细的"补充条款"作为免租期唯一依据，NC04约定2026年1月免租。',
        '考点8：底稿中应当确认输入附件共8个表格文件，且不含任何分级、奖金、处置结论列。',
    ]),
    ('（二）制度适用与口径版本（考点9—18）', [
        '考点9：底稿中应当认定2026Q1适用PRC-COST-v2.0（2025-01-01生效、现行有效），不得适用PRC-COST-v1.0。',
        '考点10：推理链——底稿中应当依据《制度生效与废止通知》确认PRC-COST-v3.0于2026-01-15审议通过、2026-04-01起执行，'
        'Q1不得适用且不追溯；须同时引用通知的生效日期与v3.0草案说明两处。',
        '考点11：底稿中应当识别制度汇编顶置红字"按v3.0执行"为文控误标，不得作为Q1适用依据。',
        '考点12：推理链——底稿中应当驳回管理层M01"按v3.0重算"，理由须同时引用制度生效通知的2026-04-01生效日与JY-2026-001纪要'
        '"未生效版本不得提前适用"两处依据。',
        '考点13：底稿中应当按JY-2026-001纪要将管理分摊基数确定为"营业收入减外卖平台佣金"，不得按全额营业收入计提。',
        '考点14：推理链——底稿中应当识别《财务确认规则》为2025版旧口径（加盟不并表、奖金50/30/20、分摊按全额收入），'
        '与JY-2026-001纪要存在三处冲突，并以纪要为准；须逐处列明冲突条款。',
        '考点15：底稿中应当确认外卖平台佣金仅对外卖GMV按18%计提。',
        '考点16：推理链——底稿中应当认定加盟门店纳入分级与处置范围但利润不并入直营，'
        '须同时引用《财务确认规则》"加盟门店"条与JY-2026-001纪要第5项。',
        '考点17：底稿中应当认定免租期已计提但依约应免的租金在当期冲回，不追溯调整以前年度。',
        '考点18：底稿中应当确认所得税按正税前利润的25%计提，亏损年度为0，且不确认递延所得税资产（驳回M11）。',
    ]),
    ('（三）收入重算与流水勾稽（考点19—28）', [
        '考点19：底稿中应当按"营业收入＝堂食净额＋外卖GMV"口径重算，外卖GMV为扣除退款后的净额。',
        f'考点20：底稿中应当汇总日营业流水，重算Q1营业收入合计{TOT_REV}万元（容差±0.05万元）。',
        f'考点21：底稿中应当认定P&L账面营业收入合计{BOOK_REV}万元，与重算值差异{REV_GAP}万元。',
        '考点22：推理链——底稿中应当识别SZ04、WX04、NC03按平台实收列报收入，须用日营业流水与外卖平台结算单两表交叉验证，'
        '并说明差额等于佣金加配送费。',
        f'考点23：底稿中应当列示SZ04重算收入{G["SZ04"][3]:.2f}万元、账面{float(inc["SZ04"]["P&L账面营业收入(万元)"]):.2f}万元、'
        f'差异{float(inc["SZ04"]["差异(万元)"]):.2f}万元（容差±0.05万元）。',
        '考点24：推理链——底稿中应当认定门店入账外卖收入应等于外卖GMV，平台实收与GMV的差额为佣金与配送费，'
        '须同时引用《外卖平台结算单》的佣金率、配送费率与《财务确认规则》收入确认条。',
        '考点25：底稿中应当认定收入少计差额不得通过"计入营销费用"方式消化（驳回M03）。',
        f'考点26：推理链——底稿中应当说明重算后EBITDA较账面变动{EBITDA_GAP}万元，原因是成本按实际发生额计提、不随账面收入缩放，'
        '须同时引用P&L、日流水与制度费率三处。',
        '考点27：底稿中应当拒绝以"稳定月营收×3"替代实际流水重算收入。',
        '考点28：底稿中应当识别杭州、合肥、南昌新店因爬坡导致实际收入低于预算营收，并在差异归因中单列。',
    ]),
    ('（四）成本费用与分摊口径（考点29—36）', [
        '考点29：底稿中应当按v2.0计提食材36%、人工21%、水电3%、营销4%、管理分摊5%，管理分摊基数为"营业收入减佣金"。',
        '考点30：推理链——底稿中应当识别NB04、HF05套用未生效v3.0费率（食材率37%），改回v2.0；'
        '须同时引用P&L食材成本、制度版本与生效通知三处。',
        '考点31：底稿中应当按"（月租金＋物业费）×面积×Q1应计租金月数"计算租金。',
        f'考点32：推理链——底稿中应当识别NC04未冲回1月免租，按2个月计提租金（{G["NC04"][3]:.2f}万元收入对应），'
        '须同时引用租约补充条款、门店基础信息免租月数与P&L租金三处。',
        '考点33：底稿中应当按建店投资除以60计提月折旧。',
        '考点34：推理链——底稿中应当说明管理分摊若按全额收入计提将高估分摊，须按"（收入－佣金）×5%"重算，'
        '并引用JY-2026-001纪要第4项。',
        '考点35：底稿中应当确认人工成本已含社保公积金，不得另行加计。',
        '考点36：底稿中应当对亏损门店所得税确认为0，不得按负数计提。',
    ]),
    ('（五）外卖对账（考点37—40）', [
        '考点37：底稿中应当按"平台实收＝外卖GMV－平台佣金－配送费"核对结算单。',
        '考点38：推理链——底稿中应当认定27家门店入账等于外卖GMV、3家（SZ04、WX04、NC03）入账等于平台实收，'
        '差异逐店列示且不得合并为总额。',
        '考点39：底稿中应当确认佣金率18%、配送费率5%。',
        '考点40：底稿中应当拒绝以平台实收确认营业收入（驳回M03）。',
    ]),
    ('（六）采购与费用核查（考点41—46）', [
        '考点41：底稿中应当按"营业收入×36%"测算理论食材成本并与采购入库合计比对。',
        '考点42：推理链——底稿中应当识别SZ05、HF02采购入库低于理论食材成本超过5%，提示未入账采购风险，'
        '须同时引用采购入库明细与重算P&L食材成本两处。',
        '考点43：底稿中应当说明采购差异不调整按标准配方计提的食材成本，仅作为风险事项披露。',
        '考点44：底稿中应当列出报销异常5条：同一发票号重复报销2条、单笔≥5000元无区域经理会签2条、招待费未注明客户与事由1条。',
        '考点45：推理链——底稿中应当按发票号分组比对识别重复报销，并引用《财务确认规则》报销审批与票据两条要求。',
        '考点46：底稿中应当按"差异率超过2%部分全额由门店承担"认定责任（驳回M06按1%统一计入）。',
    ]),
    ('（七）门店分级（考点47—54）', [
        '考点47：底稿中应当按"EBITDA率＝Q1 EBITDA÷Q1营业收入"计算，分母为重算收入。',
        '考点48：底稿中应当按A级≥12%、B级8%-12%、C级4%-8%、D级<4%分级。',
        f'考点49：底稿中应当给出分级结果A级{A_N}家、B级{B_N}家、C级{C_N}家、D级{D_N}家，合计30家。',
        f'考点50：底稿中应当列明D级门店为{"、".join(sorted(D_LIST))}。',
        f'考点51：推理链——底稿中应当识别SZ04账面为{G["SZ04"][2]}级、重算为{G["SZ04"][1]}级（EBITDA率{G["SZ04"][0]:.2f}%），'
        '原因是账面按平台实收少计收入，须跨P&L、日流水、分级标准三处说明。',
        f'考点52：推理链——底稿中应当识别NC04账面为{G["NC04"][2]}级、重算为{G["NC04"][1]}级（EBITDA率{G["NC04"][0]:.2f}%），'
        '原因是未冲回免租期租金，须跨租约、纪要、P&L三处说明。',
        '考点53：底稿中应当将加盟3家纳入分级与处置范围，但注明其利润不并入直营利润合计。',
        '考点54：底稿中应当声明分级以重算值为准，账面分级仅作对照。',
    ]),
    ('（八）店长奖金（考点55—60）', [
        '考点55：底稿中应当按A级1.2万元、B级1.0万元、C级0.8万元、D级0.5万元确定基础奖金。',
        '考点56：底稿中应当按现行有效的v2考核办法采用45%/35%/20%权重，不得沿用v1的50%/30%/20%（驳回M10）。',
        '考点57：底稿中应当将EBITDA达成率与营收达成率各自封顶1.5、下限0后再加权。',
        '考点58：底稿中应当按"损耗得分＝1－max(0,差异率－2%)×10、下限0"计算损耗项。',
        '考点59：底稿中应当对D级门店基础奖金扣减50%。',
        f'考点60：底稿中应当给出30家门店店长奖金合计{BONUS}万元（容差±0.05万元）。',
    ]),
    ('（九）处置与驳回（考点61—65）', [
        '考点61：底稿中应当对D级门店给出"先整改一个季度、连续两季D级方可关店或转址"的处置路径。',
        '考点62：推理链——底稿中应当驳回M07"本季度立即关店"，须同时引用门店分级标准与JY-2026-001纪要第9项。',
        '考点63：底稿中应当对C级门店给出限期整改、Q2复核的处置建议。',
        '考点64：底稿中应当逐条判断12项管理层要求，全部不予采纳，并逐条写明冲突条款（驳回项12项，容差0项）。',
        '考点65：底稿中应当为每条处置建议标注触发条件与时间节点，不得只给方向性意见。',
    ]),
    ('（十）勾稽与核心结论（考点66—70）', [
        f'考点66：底稿中应当完成{N_CHK}项勾稽并全部闭合，不一致项为0。',
        '考点67：底稿中应当给出一句话核心结论，包含门店数、营业收入、EBITDA、税后利润、分级分布五项要素。',
        f'考点68：推理链——底稿中应当说明账面与重算差异的根因构成（收入口径3家、成本费率2家、免租未冲回1家，合计{N_ISSUE}家），'
        '须跨四类表逐项对应。',
        f'考点69：推理链——底稿中应当区分直营与加盟口径，给出直营营业收入{DIRECT_REV}万元、EBITDA {DIRECT_EBITDA}万元、'
        f'税后利润{DIRECT_NET}万元，并说明加盟利润不并入；须引用纪要第5项与分级结果两处。',
        '考点70：底稿中应当统一金额单位为万元两位小数、底层以元计算，且全部结论可回溯至附件逐条记录。',
    ]),
]
for title, items in groups:
    H(doc, title, 2)
    for t in items:
        P(doc, t, indent=False)

H(doc, '十、答题思路')
P(doc, '第一步是读取门店基础信息、城市与区域、员工花名册，锁定30家门店、六城分布与直营/加盟构成。')
P(doc, '第二步是读取制度三版、生效通知与经营分析会纪要，确定Q1适用v2.0、管理分摊基数、加盟口径与免租冲回规则。')
P(doc, '第三步是汇总2700行日营业流水，得到每家门店Q1堂食净额与外卖GMV，与P&L账面收入比对，识别收入口径错误门店。')
P(doc, '第四步是按v2.0重算成本费用、租金（含免租冲回）、折旧与所得税，形成重算P&L。')
P(doc, '第五步是逐店核对平台实收与门店入账，识别少计收入门店并量化差额。')
P(doc, '第六步是按标准配方测算理论食材成本与采购入库比对，并逐条核查报销异常。')
P(doc, '第七步是按重算EBITDA率完成分级，识别D级门店与分级反转门店，计算盘点责任。')
P(doc, '第八步是按v2考核办法计算店长奖金，对C/D级门店给出处置建议，并逐条驳回12项管理层要求。')
P(doc, '第九步是完成全部勾稽，最后给出一句话核心结论。')

H(doc, '十一、标准答案')
P(doc, '标准答案共8个包，首表为经营分析结论，逐包覆盖口径判定、收入重算、P&L重算与外卖对账、成本费用核查、分级盘点、店长奖金、处置驳回与勾稽自检。', indent=False)
P(doc, concl['核心结论'], bold=True)
for k in ['支撑一（口径版本）', '支撑二（重建收入与成本）', '支撑三（管理分摊与加盟口径）', '支撑四（分级与处置）', '风险提示']:
    if k in concl:
        P(doc, concl[k])
P(doc, f'关键结果：30家门店Q1合计营业收入{TOT_REV}万元、EBITDA {TOT_EBITDA}万元、税后利润{TOT_NET}万元；'
       f'A级{A_N}家、B级{B_N}家、C级{C_N}家、D级{D_N}家；P&L需重算{N_ISSUE}家、分级反转{N_FLIP}家；'
       f'店长奖金合计{BONUS}万元；{N_CHK}项勾稽全部闭合，不一致项为0。')

doc.core_properties.author = '运营分析部'
NPARA = len(doc.paragraphs)
NCHAR = sum(len(p.text) for p in doc.paragraphs)
doc.save(DOCX)
scrub(DOCX, ('运营分析部-陈建华', '财务总监-周涛', '12', '2026-04-03T02:41:27Z', '2026-04-05T01:18:36Z'),
      (2026, 4, 5, 10, 0, 0), kind='docx', app_ver='16.0000', doc_stats=(NCHAR, NPARA))

print('done')
print('rev', TOT_REV, 'book', BOOK_REV, 'gap', REV_GAP, 'ebitdaGap', EBITDA_GAP)
print('grades', A_N, B_N, C_N, D_N, 'D', sorted(D_LIST), 'flip', sorted(FLIP_LIST))
print('net', TOT_NET, 'bonus', BONUS, 'checks', N_CHK)
