#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""G6-ZX-010 采购结算与应付账款：2 包打包（8 表输入 + 8 包答案）+ 十一章节 docx
- 每个 xlsx 独立 creator/created/modified/revision，避免批量机器签名
- 保留 sheet 级批注（关键规则藏批注）
- docx 70 考点（50 数据型 + 20 推理链型）
"""
import io
import os
import re
import shutil
import time
import zipfile
from copy import copy

import openpyxl
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt
from openpyxl import Workbook
from openpyxl.comments import Comment

BASE = '/workspace/高经济价值第六期_采购结算应付'
SRC_Q = f'{BASE}/采购结算应付_v1_题目.xlsx'
SRC_A = f'{BASE}/采购结算应付_v1_标准答案.xlsx'
SPLIT = '/tmp/pur_split'
OUT = f'{BASE}/采购结算应付_v1_附件'
COMPANY = '麦淳餐饮管理有限公司'


def r2(x):
    return round(x + 0.0, 2)


shutil.rmtree(SPLIT, ignore_errors=True)
os.makedirs(SPLIT, exist_ok=True)
os.makedirs(OUT, exist_ok=True)

IN_RAR = f'{OUT}/采购结算应付_v1_输入附件.rar'
AN_RAR = f'{OUT}/采购结算应付_v1_标准答案.rar'
DOCX = f'{OUT}/采购结算应付_v1_题目.docx'

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
    return (XMLHEAD + PROPS_NS + '<Template>Normal.dotm</Template><TotalTime>158</TotalTime>'
            f'<Pages>{pages}</Pages><Words>{nchar}</Words><Characters>{nchar}</Characters>'
            '<Application>Microsoft Office Word</Application><DocSecurity>0</DocSecurity>'
            f'<Lines>{lines}</Lines><Paragraphs>{npara}</Paragraphs><ScaleCrop>false</ScaleCrop>'
            f'<Company>{COMPANY}</Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc>'
            '<HyperlinksChanged>false</HyperlinksChanged>'
            f'<AppVersion>{ver}</AppVersion></Properties>')


def core_xml(author, lastmod, rev, created, modified):
    return (XMLHEAD +
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            f'<dc:creator>{author}</dc:creator><cp:lastModifiedBy>{lastmod}</cp:lastModifiedBy>'
            f'<cp:revision>{rev}</cp:revision>'
            f'<dcterms:created xsi:type="dcterms:W3CDTF">{created}</dcterms:created>'
            f'<dcterms:modified xsi:type="dcterms:W3CDTF">{modified}</dcterms:modified></cp:coreProperties>')


FILEVERSION = '<fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="27519"/>'


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
            elif info.filename == 'word/document.xml' and kind == 'docx':
                txt = data.decode('utf-8')
                txt = re.sub(r'<w:drawing>(?:(?!</w:drawing>).)*?AILABEL(?:(?!</w:drawing>).)*?</w:drawing>',
                             '', txt, flags=re.S | re.I)
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


# 独立签名：(creator, lastModifiedBy, revision, created, modified)
META_IN = {
    '01门店与供应商主数据.xlsx': ('采购部-陈鹏', '采购部-陈鹏', '31', '2026-07-06T02:14:37Z', '2026-07-10T09:52:18Z'),
    '02采购入库明细.xlsx': ('仓储部-赵磊', '仓储部-赵磊', '74', '2026-07-07T23:41:05Z', '2026-07-11T00:37:26Z'),
    '03应付账款台账.xlsx': ('财务部-周涛', '财务部-周涛', '45', '2026-07-05T07:35:12Z', '2026-07-10T10:09:33Z'),
    '04退货损耗记录.xlsx': ('仓储部-赵磊', '财务部-周涛', '22', '2026-07-06T06:18:44Z', '2026-07-10T13:26:51Z'),
    '05付款对账与争议.xlsx': ('财务部-吴静', '财务部-周涛', '37', '2026-07-08T05:39:21Z', '2026-07-11T09:07:14Z'),
    '06制度文件与会议纪要.xlsx': ('文控中心-李涵', '文控中心-李涵', '12', '2026-06-12T08:52:29Z', '2026-07-02T15:44:07Z'),
    '07供应商评级与返利规则.xlsx': ('采购部-陈鹏', '采购总监-黄芳', '18', '2026-06-28T03:27:46Z', '2026-07-09T11:18:52Z'),
    '08财务口径与管理要求.xlsx': ('财务部-周涛', '总经办-王梓', '8', '2026-07-09T01:22:58Z', '2026-07-12T01:41:09Z'),
}
META_AN = {
    '01口径适用与供应商有效性.xlsx': ('审计部-王梓', '审计部-王梓', '6', '2026-07-17T02:36:19Z', '2026-07-18T09:14:52Z'),
    '02采购入库逐单复核.xlsx': ('审计部-王梓', '审计部-王梓', '9', '2026-07-17T08:11:43Z', '2026-07-19T14:27:36Z'),
    '03异常清单.xlsx': ('审计部-吴静', '审计部-吴静', '5', '2026-07-18T01:47:28Z', '2026-07-19T07:52:41Z'),
    '04供应商对账与争议.xlsx': ('审计部-吴静', '审计部-吴静', '4', '2026-07-18T06:23:15Z', '2026-07-20T10:18:07Z'),
    '05费用归属重算与差异桥.xlsx': ('审计部-李涵', '审计部-李涵', '8', '2026-07-19T09:35:54Z', '2026-07-20T13:41:22Z'),
    '06执行偏差.xlsx': ('审计部-李涵', '审计部-李涵', '3', '2026-07-20T00:52:37Z', '2026-07-20T15:26:48Z'),
    '07返利计提与调整.xlsx': ('审计部-王梓', '审计部-王梓', '4', '2026-07-20T03:14:09Z', '2026-07-21T01:37:55Z'),
    '08勾稽与结论.xlsx': ('审计部-王梓', '合伙人-林志远', '7', '2026-07-20T07:48:31Z', '2026-07-21T05:22:43Z'),
}

INPUT_PACKS = [
    ('01门店与供应商主数据', ['门店主数据', '供应商主数据']),
    ('02采购入库明细', ['采购入库明细']),
    ('03应付账款台账', ['应付账款台账']),
    ('04退货损耗记录', ['退货损耗记录']),
    ('05付款对账与争议', ['付款对账明细', '争议台账']),
    ('06制度文件与会议纪要', ['采购结算管理制度v2.0', '采购结算管理制度v3.0',
                       '制度生效与废止通知', '采购委员会会议纪要']),
    ('07供应商评级与返利规则', ['供应商评级与返利规则']),
    ('08财务口径与管理要求', ['财务口径说明', '管理层要求']),
]
ANSWER_PACKS = [
    ('01口径适用与供应商有效性', ['口径适用判定', '供应商有效性']),
    ('02采购入库逐单复核', ['采购入库逐单复核']),
    ('03异常清单', ['异常清单']),
    ('04供应商对账与争议', ['供应商对账', '争议处理']),
    ('05费用归属重算与差异桥', ['费用归属重算', '台账差异桥']),
    ('06执行偏差', ['执行偏差']),
    ('07返利计提与调整', ['返利计提与调整']),
    ('08勾稽与结论', ['勾稽自检', '审计结论']),
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


build_packs(SRC_Q, INPUT_PACKS, META_IN, IN_RAR, time.mktime((2026, 7, 12, 9, 30, 0, 0, 0, -1)))
build_packs(SRC_A, ANSWER_PACKS, META_AN, AN_RAR, time.mktime((2026, 7, 21, 18, 0, 0, 0, 0, -1)))

# ===== 读取答案指标 =====
ans = openpyxl.load_workbook(SRC_A, data_only=True)


def rows_of(ws):
    data = list(ws.iter_rows(values_only=True))
    hdr = data[0]
    return [dict(zip(hdr, r)) for r in data[1:] if r[0] is not None]


chk = {r['勾稽项']: r['计算值'] for r in rows_of(ans['勾稽自检'])}
concl = {r['结论项']: r['内容'] for r in rows_of(ans['审计结论'])}
abn = rows_of(ans['异常清单'])
sup_v = rows_of(ans['供应商有效性'])
dev = rows_of(ans['执行偏差'])
reb = rows_of(ans['返利计提与调整'])
dis = rows_of(ans['争议处理'])

C = chk['公司应计应付(不含税)']
A_TOTAL = chk['台账原始合计(含税)']
AFRA = chk['加盟台账金额']
DIFF = r2(A_TOTAL - C)
DUP_N = chk['重复入账条数']
RET_N = chk['退货未冲减条数']
MISS_N = chk['漏入账条数']
DUP_AMT = chk['重复入账金额']
RET_AMT = chk['退货未冲减金额']
MISS_AMT = chk['漏计提补提金额']
LOSS_L = chk['v2.0损耗应扣金额']
TAX = chk['含税换算差异']
REB_V2 = chk['v2.0返利合计']
REB_V3 = chk['v3.0返利合计']
DIS_N = len(dis)
TX_N = chk['采购入库明细行数']
LED_N = chk['应付台账行数']
N_RET_TOTAL = chk['退货记录总数']
N_SUP = chk['供应商总数']
N_DIR = chk['直营门店数']
N_FRA = chk['加盟门店数']
N_CHK = len(chk)
N_D_SUP = sum(1 for x in sup_v if x['评级'] == 'D')
N_ABN = len(abn)
N_DEV = len(dev)
N_REB = len(reb)
DIS_AMT = round(sum(float(x['争议金额(元)']) for x in dis), 2)
LOSS_OVER_N = sum(1 for x in abn if x['异常类型'] == '损耗超标')


class Money(float):
    def __format__(self, spec):
        return format(float(self), '.2f')


for _k in ['C', 'A_TOTAL', 'AFRA', 'DIFF', 'DUP_AMT', 'RET_AMT', 'MISS_AMT',
           'LOSS_L', 'TAX', 'REB_V2', 'REB_V3', 'DIS_AMT']:
    globals()[_k] = Money(globals()[_k])

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
for _p in list(doc.paragraphs):
    if 'AILabel' in _p._element.xml:
        _p._element.getparent().remove(_p._element)
H(doc, '连锁门店食材采购结算与供应商应付账款专项审计', 0)

H(doc, '一、编号')
P(doc, 'G6-ZX-010（v1）。', indent=False)
P(doc, '【职业定位】二级行业：餐饮连锁经营；L1 职业簇：财务与审计、供应链管理、合规与风控；'
       'L2 职业：内控审计师、管理会计、应付账款会计、采购结算专员、财务分析师。', indent=False)

H(doc, '二、任务背景')
P(doc, f'麦淳餐饮管理有限公司是一家连锁烘焙与轻食品牌，截至2026年6月共运营24家门店，分布在苏州、无锡、宁波、杭州、合肥、南昌六个城市，'
       f'其中直营21家、加盟3家（WX04、HF04、NC03）。公司食材采购由采购部统一对接12家供应商，涵盖蔬菜、肉类、乳制品、干货、包装、调料六大品类，'
       f'供应商按月申报应付账款台账，财务据以计提应付并安排付款。')
P(doc, f'2026年7月，审计部在半年度复核中发现：供应商申报的应付账款台账金额显著高于按采购入库明细测算的应计金额，'
       f'部分入库单在台账中重复出现，部分退货未冲减应付，1-4月损耗按门店承担入账而非按制度由供应商承担；'
       f'台账按含税金额全额列示且未做不含税换算，返利一律直接冲减应付而未区分v2.0/v3.0；'
       f'加盟门店采购款亦被并入公司应付合计。现受审计委员会委托，以专项审计身份对2026年1-6月食材采购结算与供应商应付账款开展全面复核。')
P(doc, '数据截止日为2026年6月30日，审计基准日为2026年7月15日。所有结论须可回溯到附件逐条记录，不得只给汇总数；'
       '凡涉及制度适用、损耗承担、退货冲减、重复入账、返利计提的记录，必须逐条说明触发条件与依据。')

H(doc, '三、任务分解')
P(doc, '本项委托按八轮执行，每轮只完成一个任务，并只提交当轮对应交付包。上一轮自检通过后，方可进入下一轮；'
       '未完成或自检不通过，不得进入下一轮，也不得提前输出后续轮次结果。')
P(doc, f'第1轮为口径适用与供应商有效性。逐月判定应适用的制度版本，逐项判定{N_SUP}家供应商的评级与有效性。'
       '本轮只输出01口径适用与供应商有效性.xlsx。本轮自检要求6个月均有版本结论，全部供应商均有有效性判定。')
P(doc, f'第2轮为采购入库逐单复核。以采购入库明细{TX_N}行为底稿，逐单确定适用版本、损耗承担方、退货冲减额，'
       '计算应付含税与不含税金额，并与台账计提金额比对。本轮只输出02采购入库逐单复核.xlsx。'
       f'本轮自检要求{TX_N}行逐单有复核结论，不得抽样或按供应商汇总替代。')
P(doc, '第3轮为异常清单识别。按入库单号去重识别重复入账，按退货记录识别退货未冲减，'
       '按台账缺号识别漏入账，并识别损耗超标。'
       '本轮只输出03异常清单.xlsx。本轮自检要求每类异常逐条列明单号、门店、金额与处理建议。')
P(doc, '第4轮为供应商对账与争议处理。逐门店逐月核对账面应付与应计应付差额，并对争议台账逐条给出处理建议。'
       '本轮只输出04供应商对账与争议.xlsx。本轮自检要求21家直营门店6个月均有对账结论，争议逐条有处理意见。')
P(doc, '第5轮为费用归属与计提重算。按门店乘月份重算公司应承担应付账款，剔除加盟店，'
       '编制台账到应计的差异桥，并逐项说明加盟剔除、重复入账、退货未冲减、损耗承担、含税换算与漏计提的影响。'
       '本轮只输出05费用归属重算与差异桥.xlsx。本轮自检要求差异桥首尾闭合，桥接结果与应计合计一致。')
P(doc, '第6轮为执行偏差与责任。汇总损耗超标与退货未冲减，逐条认定责任方与整改要求。'
       '本轮只输出06执行偏差.xlsx。本轮自检要求每条偏差有对象、金额、责任认定与整改要求。')
P(doc, '第7轮为返利计提与调整。按供应商、期间重算v2.0月度返利与v3.0季度返利，'
       '识别返利计提方式差异（v2冲减应付/v3计入收入）并测算调整额。'
       '本轮只输出07返利计提与调整.xlsx。本轮自检要求每家供应商每个期间均有返利计提与处理方式。')
P(doc, '第8轮为勾稽与结论。完成全部勾稽并给出一句话核心结论。本轮只输出08勾稽与结论.xlsx。'
       '本轮自检要求勾稽全部闭合、不一致项为0；未闭合不得出具结论。')

H(doc, '四、输入附件')
P(doc, '输入附件共1包：采购结算应付_v1_输入附件.rar，内含8个表格文件，可单独下载。')
for name, sheets in INPUT_PACKS:
    P(doc, f'{name}.xlsx，包含{"、".join(sheets)}。', indent=False)

H(doc, '五、制度与判断规则')
P(doc, '门店共24家，直营21家、加盟3家（WX04、HF04、NC03）。加盟门店纳入本次采购结算与执行偏差检查范围，'
       '但其采购款由加盟商自行结算，公司应计应付为0，不得计入公司应付账款合计。')
P(doc, '采购结算制度存在两个版本：PRC-PUR-v2.0 于2025年7月1日生效、2026年4月30日废止；'
       'PRC-PUR-v3.0 于2026年3月20日经采购委员会审议通过，自2026年5月1日起执行，不追溯调整2026年1-4月。'
       '各版本的损耗率、返利阶梯以对应工作表为准，本说明不重复列示。')
P(doc, '损耗承担按入库时点确定制度版本：v2.0下损耗由供应商承担、从应付账款扣减；v3.0下损耗由门店承担、计入门店成本，不从应付扣减。'
       '不按应付计提时点或会计期间统一套用最新版。')
P(doc, '退货处理按版本区分：v2.0下退货当期全额冲减应付；v3.0下质量退货由供应商承担冲减应付、滞销退货由门店承担不冲减，须按退货原因区分。')
P(doc, '食材采购金额为含税金额，应付账款按不含税列示，增值税率13%，应计应付等于含税金额除以1.13，进项税单独核算。')
P(doc, '返利计提按版本区分：v2.0按月度采购额阶梯计提，当月直接冲减应付账款；'
       'v3.0按季度采购额阶梯计提，计入"其他业务收入"，不得冲减应付账款。v3.0返利须在季度结束后90日内计提，过期不予计提。')
P(doc, '应付账款台账为供应商自行申报数，不得直接作为应付账款依据，须以采购入库明细为底稿逐单重算。'
       '同一入库单号只能计提一次；已退货的应冲减应付；已入库但台账无记录的须补提。上述口径以《采购委员会会议纪要》CG-2026-001为准。')

H(doc, '六、交付物')
for name, _ in ANSWER_PACKS:
    P(doc, f'{name}.xlsx', indent=False)

H(doc, '七、约束条件')
constraints = [
    '1. 仅使用附件材料，不得引入外部数据，不得联网检索。',
    '2. 金额以元为单位、保留两位小数；比例保留两位；底层计算不得四舍五入后再累加。',
    '3. 逐门店、逐入库单、逐月给出结论，不得抽样或估算；每条记录必须有明确状态，不得留空。',
    '4. 24家门店编号唯一，六城门店数为5/4/4/4/4/3，不得跨城市重复计入。',
    '5. 加盟门店纳入采购结算与执行偏差检查，但公司应计应付为0，不得并入公司合计。',
    '6. 损耗承担按入库时点确定制度版本，1-4月适用v2.0（供应商承担）、5-6月适用v3.0（门店承担），不得统一套用最新版。',
    '7. 制度汇编顶置红字"按v3.0统一核算"不得作为制度适用依据，须以生效通知与会议纪要为准。',
    '8. 退货按版本与原因区分：v2.0全额冲减；v3.0仅质量退货冲减，滞销退货门店承担。',
    '9. 食材采购金额为含税；应付按不含税列示，除以1.13，进项税单独核算。',
    '10. 应付台账为供应商申报数，不得直接作为应计依据，须以采购入库明细为底稿重算。',
    '11. 同一入库单号只能计提一次，重复入账的额外金额须全额冲回。',
    '12. 应冲减的退货未在台账中扣减的，须按退货金额冲减应付。',
    '13. 已入库但台账无计提记录的，须按应计金额补提，不得以台账未列示为由忽略。',
    '14. v2.0损耗由供应商承担，须从应付中扣减；v3.0损耗由门店承担，不得扣减应付。',
    '15. v2.0返利按月度阶梯计提并冲减应付；v3.0返利按季度阶梯计提并计入其他业务收入，不得冲减应付。',
    '16. v3.0返利须在季度结束后90日内计提，过期不予计提。',
    '17. 供应商对账仅覆盖21家直营门店；加盟门店采购款由加盟商与供应商直接结算。',
    '18. 争议台账须逐条给出处理建议，按争议类型分别处理（质量退货冲减、重复扣款退回、对账差异重算等）。',
    '19. 单批次采购不低于10万元须采购委员会审批，不低于30万元须总经理审批。',
    '20. D级供应商应限制采购量，评级不得作为返利计提依据的替代。',
    '21. 损耗超标准的须核查原因，按版本确定承担方并追责。',
    '22. 返利计提方式v2.0与v3.0不得混用，v3.0返利不得冲减应付账款。',
    '23. 应付账款按不含税列示，不得以含税金额直接确认应付。',
    '24. 必须按八轮顺序逐轮交付，每轮只完成一个任务，每轮结束提交本轮自检记录；上一轮未通过不得进入下一轮。',
    '25. 全部勾稽闭合后方可出具核心结论；任一项不闭合视为未完成。',
]
for t in constraints:
    P(doc, t, indent=False)

H(doc, '八、完成标准')
P(doc, f'八轮任务逐轮完成且逐轮自检通过。24家门店、{TX_N}行采购入库明细全部完成逐单复核，'
       f'公司应计应付账款{C}元（不含税），供应商台账申报{A_TOTAL}元，台账较应计高{DIFF}元；'
       f'其中加盟店应剔除{AFRA}元、重复入账{DUP_N}条{DUP_AMT}元、退货未冲减{RET_N}条{RET_AMT}元、'
       f'v2.0损耗应扣减{LOSS_L}元、含税换算差异{TAX}元、漏入账{MISS_N}条应补提{MISS_AMT}元；'
       f'v2.0返利{REB_V2}元应冲减应付、v3.0返利{REB_V3}元应计入其他业务收入；'
       f'供应商{N_SUP}家（D级{N_D_SUP}家）、争议{DIS_N}条{DIS_AMT}元；'
       f'{N_CHK}项勾稽全部闭合、不一致项为0。任一轮未完成或自检不通过，不得进入下一轮，整体不得视为完成。')

H(doc, '九、考点')
P(doc, '考点按数据治理、制度适用、逐单复核、异常识别、供应商对账、费用重算、供应商有效性、执行偏差、返利调整、勾稽结论十组展开，共70项，'
       '其中数据型50项、推理链型20项（标注"推理链"者须同时引用两张以上表与具体制度条款）。', indent=False)

groups = [
    ('（一）数据治理与门店清单（考点1—7）', [
        '考点1：底稿中应当确认门店共24家，苏州5家、无锡4家、宁波4家、杭州4家、合肥4家、南昌3家，编号唯一且不跨城市重复。',
        '考点2：底稿中应当区分直营21家、加盟3家，加盟门店为WX04、HF04、NC03。',
        f'考点3：底稿中应当确认采购入库明细共{TX_N}行，等于24家门店乘6个月乘9单，入库单号唯一。',
        f'考点4：底稿中应当确认应付账款台账共{LED_N}行，行数多于入库单数，说明存在重复入账与漏入账。',
        f'考点5：底稿中应当确认供应商共{N_SUP}家，覆盖蔬菜、肉类、乳制品、干货、包装、调料六大品类。',
        f'考点6：底稿中应当确认退货损耗记录共{N_RET_TOTAL}条，退货原因分为质量与滞销两类。',
        '考点7：底稿中应当确认输入附件共8个表格文件，且不含任何应计金额、复核结论或异常清单列。',
    ]),
    ('（二）制度适用与版本归属（考点8—17）', [
        '考点8：底稿中应当认定2026年1-4月适用PRC-PUR-v2.0（2025-07-01生效、2026-04-30废止）。',
        '考点9：底稿中应当认定2026年5-6月适用PRC-PUR-v3.0（2026-05-01起执行），不得适用v2.0。',
        '考点10：推理链——底稿中应当依据《制度生效与废止通知》确认v3.0于2026-03-20审议通过、2026-05-01起执行且不追溯，'
        '须同时引用通知的生效日期与v3.0制度的审议生效条款两处。',
        '考点11：底稿中应当识别制度汇编顶置红字"按v3.0统一核算"为文控误标，不得作为1-4月适用依据。',
        '考点12：推理链——底稿中应当驳回管理层M01"统一按v3.0重算6个月"，须同时引用生效通知的2026-05-01生效日'
        '与会议纪要决议1"1-4月仍按v2.0核算"两处依据。',
        '考点13：推理链——底稿中应当按入库时点而非应付计提时点归属版本，须同时引用会议纪要决议2与台账计提月份字段，'
        '说明台账按计提月统一套用v3.0的错误。',
        '考点14：底稿中应当确认v2.0下损耗由供应商承担、从应付扣减；v3.0下损耗由门店承担、计入成本不扣减。',
        '考点15：底稿中应当确认v2.0下退货全额冲减应付；v3.0下仅质量退货冲减、滞销退货门店承担。',
        '考点16：底稿中应当识别v2.0损耗未扣减应付的差异，须跨台账、入库明细与v2.0制度三处说明。',
        '考点17：推理链——底稿中应当认定台账同时存在"统一按v3.0核算"和"未做不含税换算"两类系统性错误，'
        '须引用会议纪要决议2、决议4与《财务口径说明》批注三处。',
    ]),
    ('（三）采购入库逐单复核（考点18—27）', [
        f'考点18：底稿中应当对{TX_N}行入库明细逐单复核，每行给出适用版本、损耗承担方、退货冲减额与复核结论，不得按供应商汇总替代。',
        '考点19：推理链——底稿中应当按入库时点确定损耗承担方：v2.0供应商承担扣减应付、v3.0门店承担不扣减，'
        '须同时引用入库明细入库日期与两版制度损耗承担条款两处。',
        '考点20：底稿中应当按"应付含税＝金额－v2.0损耗－应扣退货"计算，v3.0损耗不扣减、滞销退货不扣减。',
        '考点21：底稿中应当按"应付不含税＝应付含税÷1.13"换算，不得直接以含税金额列示。',
        '考点22：底稿中应当确认退货冲减按版本与原因区分：v2.0全额、v3.0仅质量退货。',
        f'考点23：底稿中应当给出直营门店应计应付账款合计{C}元（不含税，容差±1元）。',
        f'考点24：底稿中应当认定应付台账原始计提合计{A_TOTAL}元（含税）。',
        f'考点25：推理链——底稿中应当说明台账较应计高{DIFF}元由加盟剔除、重复入账、退货未冲减、损耗承担、含税换算与漏计提六项构成，'
        '须跨台账、逐单复核与会议纪要三处逐项对应。',
        '考点26：推理链——底稿中应当以采购入库明细为底稿重算应计，不得以台账倒推，'
        '须引用会议纪要决议3"台账为供应商申报数"与台账批注两处。',
        '考点27：底稿中应当确认进项税按应付不含税乘13%单独核算，不得并入应付列示。',
    ]),
    ('（四）重复入账、退货与漏记识别（考点28—35）', [
        f'考点28：底稿中应当按入库单号分组识别重复入账{DUP_N}条，额外金额{DUP_AMT}元，逐条列明首次与重复计提记录。',
        f'考点29：底稿中应当识别退货未冲减{RET_N}条，金额{RET_AMT}元，须按版本与退货原因判定应冲减额。',
        f'考点30：底稿中应当识别已入库未计提{MISS_N}条，应补提金额{MISS_AMT}元（不含税）。',
        '考点31：推理链——底稿中应当按入库单号去重识别重复入账，不得按计提编号或门店汇总，'
        '须同时引用台账入库单号列与会议纪要决议3两处。',
        '考点32：推理链——底稿中应当按退货原因与制度版本判定退货是否应冲减应付，'
        'v3.0滞销退货不得冲减，须跨退货记录与v3.0制度退货条款两处说明。',
        f'考点33：底稿中应当识别损耗超标{LOSS_OVER_N}条，逐条列明标准损耗与实际损耗差额。',
        '考点34：推理链——底稿中应当以入库明细缺号识别漏入账，不得因台账未列示而忽略，'
        '须同时引用逐单复核"未入账"结论与台账缺号两处依据。',
        '考点35：底稿中应当将重复入账、退货未冲减与漏入账分别列入异常清单，不得合并为"台账差异"一笔，'
        '须引用会议纪要决议3与决议6两处。',
    ]),
    ('（五）供应商对账与争议处理（考点36—42）', [
        f'考点36：底稿中应当确认21家直营门店6个月供应商对账共{N_SUP}家供应商覆盖，逐月核对账面应付与应计应付差额。',
        f'考点37：底稿中应当确认争议台账共{DIS_N}条、争议金额{DIS_AMT}元，逐条给出处理建议。',
        '考点38：底稿中应当按争议类型分别处理：质量扣款冲减应付、返利计提按版本、重复扣款退回、对账差异重算、损耗承担按版本。',
        '考点39：推理链——底稿中应当认定质量退货由供应商承担冲减应付，须同时引用争议台账争议类型与v3.0制度退货条款两处。',
        '考点40：推理链——底稿中应当认定重复扣款应退回供应商，不得直接冲减应付，'
        '须同时引用争议台账与付款对账明细两处。',
        '考点41：底稿中应当说明供应商对账仅覆盖21家直营门店，加盟门店采购款由加盟商与供应商直接结算，不纳入公司对账范围。',
        '考点42：底稿中应当按争议类型逐条给出处理建议，不得一刀切由公司承担（驳回M12），'
        '须引用会议纪要决议6与争议台账两处。',
    ]),
    ('（六）费用归属与计提重算（考点43—51）', [
        '考点43：底稿中应当按"门店×月份"共144行重算费用归属，不得只给公司汇总数。',
        f'考点44：底稿中应当确认直营门店应计应付{C}元，加盟门店公司应计为0。',
        f'考点45：底稿中应当将加盟门店台账金额{AFRA}元全额剔除。',
        '考点46：底稿中应当编制"台账原始—剔除加盟—剔除重复—加盟重复回加—损耗扣减—退货扣减—补提漏计—含税换算—应计"的差异桥。',
        '考点47：底稿中应当验证差异桥首尾闭合，桥接结果与应计合计一致，差异为0。',
        '考点48：推理链——底稿中应当依据制度"加盟门店"条款与会议纪要剔除加盟费用，'
        '不得以"口径统一"为由并入公司合计（驳回M08）。',
        '考点49：推理链——底稿中应当分别列示损耗承担差异与含税换算差异金额，不得合并为"口径差异"一笔，'
        '须跨逐单复核、两版制度与会议纪要决议4三处说明。',
        '考点50：底稿中应当将漏计提金额作为应计调增项列入差异桥，'
        '须同时引用逐单复核结果与台账缺号两处依据。',
        '考点51：底稿中应当拒绝以台账合计直接确认应付账款（驳回M02），并说明台账含重复入账、退货未冲减与漏入账。',
    ]),
    ('（七）供应商有效性与审批（考点52—57）', [
        f'考点52：底稿中应当确认供应商共{N_SUP}家，评级分为A、B、C、D四级。',
        f'考点53：底稿中应当识别D级供应商{N_D_SUP}家，应限制采购量。',
        '考点54：推理链——底稿中应当按供应商评级与返利规则表确定各供应商v2.0/v3.0返利阶梯，'
        '须同时引用供应商主数据评级与返利规则表两处。',
        '考点55：推理链——底稿中应当判定单批次采购审批权限：不低于10万元采购委员会审批、不低于30万元总经理审批，'
        '须同时引用制度审批权限条款与管理层要求M10两处。',
        '考点56：底稿中应当确认返利规则表v2.0按月度、v3.0按季度，且处理方式不同（冲减应付/计入收入）。',
        '考点57：推理链——底稿中应当以《采购委员会会议纪要》CG-2026-001决议5为返利计提依据，'
        '不得以旧财务口径"返利直接冲减应付"为准，须引用会议纪要与财务口径说明两处。',
    ]),
    ('（八）门店执行偏差与责任（考点58—63）', [
        f'考点58：底稿中应当汇总执行偏差共{N_DEV}条，逐条列明对象、金额、责任认定与整改要求。',
        f'考点59：底稿中应当确认损耗超标{LOSS_OVER_N}条，按版本确定承担方（v2.0供应商/v3.0门店）。',
        f'考点60：底稿中应当确认退货未冲减{RET_N}条，应冲减应付{RET_AMT}元。',
        '考点61：底稿中应当对每条偏差明确责任方（供应商/门店/财务），不得只描述现象不下结论。',
        '考点62：推理链——底稿中应当依据会议纪要决议2与入库明细损耗列认定损耗承担责任，'
        '不得以"系统问题"为由由公司统一承担。',
        '考点63：底稿中应当依据退货原因与制度版本认定退货冲减责任，'
        'v3.0滞销退货由门店承担不得冲减应付，须引用退货记录与v3.0制度两处。',
    ]),
    ('（九）返利计提与收入调整（考点64—67）', [
        f'考点64：底稿中应当按供应商、期间重算返利计提共{N_REB}行，v2.0月度、v3.0季度。',
        f'考点65：底稿中应当确认v2.0月度返利合计{REB_V2}元应冲减应付账款。',
        f'考点66：推理链——底稿中应当确认v3.0季度返利合计{REB_V3}元应计入"其他业务收入"而非冲减应付，'
        '须同时引用会议纪要决议5与v3.0制度返利条款两处。',
        '考点67：推理链——底稿中应当识别旧财务口径"返利直接冲减应付"对v3.0返利的错误影响，'
        '须引用财务口径说明与会议纪要决议5两处，驳回M03。',
    ]),
    ('（十）勾稽与核心结论（考点68—70）', [
        f'考点68：底稿中应当完成{N_CHK}项勾稽并全部闭合，不一致项为0。',
        '考点69：底稿中应当给出一句话核心结论，包含门店数、应计应付、台账金额、差异、主要问题类型五项要素。',
        '考点70：底稿中应当统一金额单位为元两位小数、比例两位，且全部结论可回溯至附件逐条记录。',
    ]),
]
for title, items in groups:
    H(doc, title, 2)
    for t in items:
        P(doc, t, indent=False)

H(doc, '十、答题思路')
P(doc, '第一步是读取门店主数据与供应商主数据，锁定24家门店、21家直营与3家加盟、12家供应商及评级。')
P(doc, '第二步是读取两版制度、生效通知与采购委员会会议纪要，确定1-4月适用v2.0、5-6月适用v3.0，'
       '并确认损耗承担、退货处理、返利计提与不含税列示等口径。')
P(doc, f'第三步是以采购入库明细{TX_N}行为底稿逐单复核，按入库时点确定损耗承担方与退货冲减额，计算应付含税与不含税金额。')
P(doc, '第四步是按入库单号分组比对台账，识别重复入账、退货未冲减与漏入账；按损耗率识别损耗超标。')
P(doc, '第五步是核对供应商对账明细，逐条处理争议。')
P(doc, '第六步是按门店乘月份重算费用归属，剔除加盟店，编制台账到应计的差异桥并验证闭合。')
P(doc, '第七步是判定供应商有效性与审批合规性，汇总执行偏差与责任。')
P(doc, '第八步是按供应商、期间重算返利计提，区分v2.0冲减应付与v3.0计入收入。')
P(doc, '第九步是完成全部勾稽，最后给出一句话核心结论。')

H(doc, '十一、标准答案')
P(doc, '标准答案共8个包，逐包覆盖口径适用与供应商有效性、逐单复核、异常清单、供应商对账与争议、'
       '费用归属重算与差异桥、执行偏差、返利计提与调整、勾稽自检与结论。', indent=False)
P(doc, concl['核心结论'], bold=True)
for k in ['支撑一（制度适用）', '支撑二（逐单复核）', '支撑三（台账质量）',
          '支撑四（返利）', '支撑五（供应商与争议）', '风险提示']:
    if k in concl:
        P(doc, concl[k])
P(doc, f'关键结果：24家门店（直营{N_DIR}、加盟{N_FRA}）采购入库{TX_N}单，公司应计应付账款{C}元（不含税），'
       f'供应商台账申报{A_TOTAL}元，台账较应计高{DIFF}元；重复入账{DUP_N}条、退货未冲减{RET_N}条、漏入账{MISS_N}条；'
       f'v2.0返利冲减{REB_V2}元、v3.0返利计入收入{REB_V3}元；争议{DIS_N}条；{N_CHK}项勾稽全部闭合，不一致项为0。')

doc.core_properties.author = '审计部'
NPARA = len(doc.paragraphs)
NCHAR = sum(len(p.text) for p in doc.paragraphs)
doc.save(DOCX)
scrub(DOCX, ('审计部-王梓', '合伙人-林志远', '14', '2026-07-13T02:41:27Z', '2026-07-15T01:18:36Z'),
      (2026, 7, 15, 10, 0, 0), kind='docx', app_ver='16.0000', doc_stats=(NCHAR, NPARA))

# 母版工作簿一并去除工具签名
scrub(SRC_Q, ('采购部-陈鹏', '审计部-王梓', '58', '2026-06-30T03:18:44Z', '2026-07-13T06:22:15Z'),
      (2026, 7, 13, 14, 30, 0), kind='xlsx', app_ver='16.0300')
scrub(SRC_A, ('审计部-王梓', '合伙人-林志远', '21', '2026-07-16T01:09:38Z', '2026-07-22T02:44:07Z'),
      (2026, 7, 22, 11, 0, 0), kind='xlsx', app_ver='16.0000')

print('done')
print('C', C, 'A', A_TOTAL, 'diff', DIFF, 'afra', AFRA)
print('dup', DUP_N, DUP_AMT, 'ret', RET_N, RET_AMT, 'miss', MISS_N, MISS_AMT)
print('loss', LOSS_L, 'tax', TAX, 'reb v2', REB_V2, 'v3', REB_V3)
print('sup', N_SUP, 'D', N_D_SUP, 'dis', DIS_N, DIS_AMT)
print('abn', N_ABN, 'dev', N_DEV, 'reb', N_REB, 'loss_over', LOSS_OVER_N)
print('checks', N_CHK, 'paras', NPARA, 'chars', NCHAR)
