#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""G6-ZX-009 营销费用核销：2 包打包（8 表输入 + 8 包答案）+ 十一章节 docx
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

BASE = '/workspace/高经济价值第六期_营销费用核销'
SRC_Q = f'{BASE}/营销费用核销_v1_题目.xlsx'
SRC_A = f'{BASE}/营销费用核销_v1_标准答案.xlsx'
SPLIT = '/tmp/mkt_split'
OUT = f'{BASE}/营销费用核销_v1_附件'
COMPANY = '麦淳餐饮管理有限公司'


def r2(x):
    return round(x + 0.0, 2)


shutil.rmtree(SPLIT, ignore_errors=True)
os.makedirs(SPLIT, exist_ok=True)
os.makedirs(OUT, exist_ok=True)

IN_RAR = f'{OUT}/营销费用核销_v1_输入附件.rar'
AN_RAR = f'{OUT}/营销费用核销_v1_标准答案.rar'
DOCX = f'{OUT}/营销费用核销_v1_题目.docx'

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
    '01门店与活动主数据.xlsx': ('营销部-林晓', '营销部-林晓', '31', '2026-07-06T02:14:37Z', '2026-07-10T09:52:18Z'),
    '02订单核销明细.xlsx': ('信息技术部-刘畅', '信息技术部-刘畅', '74', '2026-07-07T23:41:05Z', '2026-07-11T00:37:26Z'),
    '03平台结算与争议.xlsx': ('营销部-孙莉', '财务部-周涛', '22', '2026-07-06T06:18:44Z', '2026-07-10T13:26:51Z'),
    '04营销费用台账.xlsx': ('财务部-周涛', '财务部-周涛', '45', '2026-07-05T07:35:12Z', '2026-07-10T10:09:33Z'),
    '05制度文件与会议纪要.xlsx': ('文控中心-李涵', '文控中心-李涵', '12', '2026-06-12T08:52:29Z', '2026-07-02T15:44:07Z'),
    '06券规则与价格.xlsx': ('营销部-林晓', '营销总监-黄芳', '18', '2026-06-28T03:27:46Z', '2026-07-09T11:18:52Z'),
    '07门店销售汇总与冲正.xlsx': ('财务部-吴静', '财务部-周涛', '37', '2026-07-08T05:39:21Z', '2026-07-11T09:07:14Z'),
    '08财务口径与管理要求.xlsx': ('财务部-周涛', '总经办-王梓', '8', '2026-07-09T01:22:58Z', '2026-07-12T01:41:09Z'),
}
META_AN = {
    '01制度适用与活动有效性.xlsx': ('审计部-王梓', '审计部-王梓', '6', '2026-07-17T02:36:19Z', '2026-07-18T09:14:52Z'),
    '02券核销逐单复核.xlsx': ('审计部-王梓', '审计部-王梓', '9', '2026-07-17T08:11:43Z', '2026-07-19T14:27:36Z'),
    '03重复超发过期券识别.xlsx': ('审计部-吴静', '审计部-吴静', '5', '2026-07-18T01:47:28Z', '2026-07-19T07:52:41Z'),
    '04平台对账与争议处理.xlsx': ('审计部-吴静', '审计部-吴静', '4', '2026-07-18T06:23:15Z', '2026-07-20T10:18:07Z'),
    '05费用归属与计提重算.xlsx': ('审计部-李涵', '审计部-李涵', '8', '2026-07-19T09:35:54Z', '2026-07-20T13:41:22Z'),
    '06门店执行偏差与责任.xlsx': ('审计部-李涵', '审计部-李涵', '3', '2026-07-20T00:52:37Z', '2026-07-20T15:26:48Z'),
    '07收入确认调整.xlsx': ('审计部-王梓', '审计部-王梓', '4', '2026-07-20T03:14:09Z', '2026-07-21T01:37:55Z'),
    '08勾稽与结论.xlsx': ('审计部-王梓', '合伙人-林志远', '7', '2026-07-20T07:48:31Z', '2026-07-21T05:22:43Z'),
}

INPUT_PACKS = [
    ('01门店与活动主数据', ['门店主数据', '活动主数据']),
    ('02订单核销明细', ['订单核销明细']),
    ('03平台结算与争议', ['平台结算', '争议台账']),
    ('04营销费用台账', ['营销费用台账']),
    ('05制度文件与会议纪要', ['营销费用管理制度v2.0', '营销费用管理制度v3.0',
                       '制度生效与废止通知', '营销委员会会议纪要']),
    ('06券规则与价格', ['券规则']),
    ('07门店销售汇总与冲正', ['门店销售汇总', '冲正台账']),
    ('08财务口径与管理要求', ['财务口径说明', '管理层要求']),
]
ANSWER_PACKS = [
    ('01制度适用与活动有效性', ['口径适用判定', '活动有效性']),
    ('02券核销逐单复核', ['券核销逐单复核']),
    ('03重复超发过期券识别', ['券异常清单']),
    ('04平台对账与争议处理', ['平台对账', '争议处理']),
    ('05费用归属与计提重算', ['费用归属重算', '台账差异桥']),
    ('06门店执行偏差与责任', ['门店执行偏差']),
    ('07收入确认调整', ['收入确认调整']),
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


chk = {r['勾稽项']: r['左值'] for r in rows_of(ans['勾稽自检'])}
concl = {r['项目']: r['内容'] for r in rows_of(ans['审计结论'])}
abn = rows_of(ans['券异常清单'])
acts = rows_of(ans['活动有效性'])
dev = rows_of(ans['门店执行偏差'])
inc = rows_of(ans['收入确认调整'])
dis = rows_of(ans['争议处理'])

TOT_C = chk['公司应计营销费用(元)']
TOT_A = chk['台账原始计提合计(元)']
AFRA = chk['加盟店台账金额(元)']
DIFF = chk['台账与应计差异(元)']
DUP_N = chk['重复计提条数']
EXP_N = chk['过期未核销条数']
MISS_N = chk['漏计提条数']
DUP_AMT = chk['重复计提金额(元)']
EXP_AMT = chk['过期未核销金额(元)']
MISS_AMT = chk['漏计提金额(元)']
VER = chk['版本比例调整(元)']
TAX = chk['含税换算调整(元)']
P_ASSUME = chk['平台应承担合计(元)']
P_ACTUAL = chk['平台实结算合计(元)']
P_DIFF = chk['平台差额合计(元)']
DIS_N = chk['争议条数']
APPROVE_BAD = chk['审批级别不足活动数']
OVER_N = chk['超预算活动数']
CROSS_N = chk['跨店核销条数']
INC_N = chk['收入调整门店月数']
TX_N = chk['核销明细行数']
LED_N = chk['营销费用台账行数']
N_CHK = len(chk)

DIS_AMT = round(sum(float(x['争议金额(元)']) for x in dis), 2)
INC_AMT = round(sum(float(x['调整额(元)']) for x in inc), 2)
N_ABN = len(abn)
OVER_CP = sum(1 for x in abn if x['异常类型'] == '券超发')
MEM_BAD = sum(1 for x in abn if x['异常类型'] == '会员日违规核销')
N_VALID = sum(1 for x in acts if x['结论'] == '有效')
N_AFTER = sum(1 for x in acts if '先执行' in str(x['结论']))
N_CROSSACT = sum(1 for x in acts if '跨期' in str(x['结论']))
N_DEV = len(dev)


class Money(float):
    def __format__(self, spec):
        return format(float(self), '.2f')


for _k in ['TOT_C', 'TOT_A', 'AFRA', 'DIFF', 'DUP_AMT', 'EXP_AMT', 'MISS_AMT',
           'VER', 'TAX', 'P_ASSUME', 'P_ACTUAL', 'P_DIFF', 'DIS_AMT', 'INC_AMT']:
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
H(doc, '连锁门店营销活动费用与优惠券核销专项审计', 0)

H(doc, '一、编号')
P(doc, 'G6-ZX-009（v1）。', indent=False)
P(doc, '【职业定位】二级行业：餐饮连锁经营；L1 职业簇：财务与审计、营销管理、合规与风控；'
       'L2 职业：内控审计师、管理会计、营销费用专员、平台运营、财务分析师。', indent=False)

H(doc, '二、任务背景')
P(doc, f'麦淳餐饮管理有限公司是一家连锁烘焙与轻食品牌，截至2026年6月共运营24家门店，分布在苏州、无锡、宁波、杭州、合肥、南昌六个城市，'
       f'其中直营21家、加盟3家。2026年上半年公司通过自有小程序、美团、饿了么、抖音团购、点评五个渠道投放满减券、折扣券、代金券、'
       f'买赠、会员日折扣与平台补贴券，营销费用由门店按月申报计入营销费用台账。')
P(doc, f'2026年7月，审计部在季度复核中发现：营销费用台账金额显著高于按核销明细测算的应计金额，部分券码重复出现，'
       f'部分已过期未核销的券仍在计提；1-4月与5-6月分属两个制度版本，但门店统一按最新版比例申报且未做不含税换算；'
       f'加盟门店费用亦被计入公司合计。现受审计委员会委托，以专项审计身份对2026年1-6月营销活动费用与优惠券核销开展全面复核。')
P(doc, '数据截止日为2026年6月30日，审计基准日为2026年7月15日。所有结论须可回溯到附件逐条记录，不得只给汇总数；'
       '凡涉及制度适用、费用归属、重复与过期、责任认定的记录，必须逐条说明触发条件与依据。')

H(doc, '三、任务分解')
P(doc, '本项委托按八轮执行，每轮只完成一个任务，并只提交当轮对应交付包。上一轮自检通过后，方可进入下一轮；'
       '未完成或自检不通过，不得进入下一轮，也不得提前输出后续轮次结果。')
P(doc, f'第1轮为制度适用与活动有效性。逐月判定应适用的制度版本，逐项判定{len(acts)}项活动的审批与执行有效性。'
       '本轮只输出01制度适用与活动有效性.xlsx。本轮自检要求6个月均有版本结论，全部活动均有有效性判定与依据。')
P(doc, f'第2轮为券核销逐单复核。以订单核销明细{TX_N}行为底稿，逐单确定券种、适用版本、商家承担比例，'
       '计算应计含税与不含税金额，并与台账计提金额比对。本轮只输出02券核销逐单复核.xlsx。'
       f'本轮自检要求{TX_N}行逐单有复核结论，不得抽样或按券编号汇总替代。')
P(doc, '第3轮为重复、超发、过期券识别。按券码去重识别重复计提，核对发放数量识别超发，'
       '按有效期与核销记录识别过期未核销及漏计提，并识别跨店核销与会员日违规。'
       '本轮只输出03重复超发过期券识别.xlsx。本轮自检要求每类异常逐条列明对象、门店、金额与问题。')
P(doc, '第4轮为平台对账与争议处理。逐门店逐月核对平台应承担、实际结算与差额，并对争议台账逐条给出处理建议。'
       '本轮只输出04平台对账与争议处理.xlsx。本轮自检要求21家直营门店6个月均有对账结论，争议逐条有处理意见。')
P(doc, '第5轮为费用归属与计提重算。按门店乘月份重算公司应承担营销费用，剔除加盟店，'
       '编制台账到应计的差异桥，并逐项说明重复、过期、漏计提、版本比例与含税换算的影响。'
       '本轮只输出05费用归属与计提重算.xlsx。本轮自检要求差异桥首尾闭合，桥接结果与应计合计一致。')
P(doc, '第6轮为门店执行偏差与责任。汇总券超发、跨店核销、会员日违规与收入确认错误，逐条认定责任门店与整改要求。'
       '本轮只输出06门店执行偏差与责任.xlsx。本轮自检要求每条偏差有对象、金额、责任人与整改要求。')
P(doc, '第7轮为收入确认调整。逐门店逐月核对账面净收入与按实收确认的正确净收入，识别券抵扣重复冲减收入的情形并测算调整额。'
       '本轮只输出07收入确认调整.xlsx。本轮自检要求差异门店月逐条列示原价、券抵扣、实收、退款、账面与调整额。')
P(doc, '第8轮为勾稽与结论。完成全部勾稽并给出一句话核心结论。本轮只输出08勾稽与结论.xlsx。'
       '本轮自检要求勾稽全部闭合、不一致项为0；未闭合不得出具结论。')

H(doc, '四、输入附件')
P(doc, '输入附件共1包：营销费用核销_v1_输入附件.rar，内含8个表格文件，可单独下载。')
for name, sheets in INPUT_PACKS:
    P(doc, f'{name}.xlsx，包含{"、".join(sheets)}。', indent=False)

H(doc, '五、制度与判断规则')
P(doc, '门店共24家，直营21家、加盟3家（WX04、HF04、NC03）。加盟门店纳入本次核销与执行偏差检查范围，'
       '但其营销费用由加盟商自行承担，公司应计为0，不得计入公司营销费用合计。')
P(doc, '营销费用制度存在两个版本：PRC-MKT-v2.0 于2025年7月1日生效、2026年4月30日废止；'
       'PRC-MKT-v3.0 于2026年3月20日经营销委员会审议通过，自2026年5月1日起执行，不追溯调整2026年1-4月。'
       '各版本的具体承担比例以对应工作表为准，本说明不重复列示。')
P(doc, '费用归属按券核销时点确定制度版本，不按发放时点、计提时点或会计期间统一套用最新版。'
       '券面值、折扣额、赠品成本均为含税金额，营销费用按不含税列示，增值税率6%，应计金额等于含税金额除以1.06。')
P(doc, '营销费用台账为门店申报数，不得直接作为应计依据。同一券码只能计提一次；未核销且已过有效期的券不得计提，已计提的须冲回；'
       '已核销但台账无记录的须补提。上述口径以《营销委员会会议纪要》MKT-2026-001为准。')
P(doc, '营业收入按订单实收确认，券抵扣不得作为销售折让重复冲减收入；平台少结算或重复扣款由营销部在次月结算中追补，'
       '不得直接冲减营销费用；跨店核销的费用由实际核销门店所在区域承担。')
P(doc, '活动审批门槛：单场活动预算不低于5万元须营销委员会审批，不低于10万元须总经理审批；先执行后审批的须补审批并说明。'
       '活动预算与实际费用以活动主数据为准，跨期活动的费用须按期间拆分。')

H(doc, '六、交付物')
for name, _ in ANSWER_PACKS:
    P(doc, f'{name}.xlsx', indent=False)

H(doc, '七、约束条件')
constraints = [
    '1. 仅使用附件材料，不得引入外部数据，不得联网检索。',
    '2. 金额以元为单位、保留两位小数；比例保留两位；底层计算不得四舍五入后再累加。',
    '3. 逐门店、逐券码、逐月给出结论，不得抽样或估算；每条记录必须有明确状态，不得留空。',
    '4. 24家门店编号唯一，六城门店数为5/4/4/4/4/3，不得跨城市重复计入。',
    '5. 加盟门店纳入核销与执行偏差检查，但公司应计营销费用为0，不得并入公司合计。',
    '6. 费用归属按券核销时点确定制度版本，1-4月适用v2.0、5-6月适用v3.0，不得统一套用最新版。',
    '7. 制度汇编顶置红字“按v3.0统一核算”不得作为制度适用依据，须以生效通知与会议纪要为准。',
    '8. 券面值、折扣额、赠品成本均为含税；营销费用按不含税列示，除以1.06。',
    '9. 平台补贴券的商家承担比例按券种、平台与核销时点适用版本共同确定，不得一刀切。',
    '10. 折扣券、会员日折扣由商家全额承担；买赠按赠品成本价计入，不得按券面值或订单原价计算。',
    '11. 同一券码只能计提一次，重复计提的额外金额须全额冲回。',
    '12. 未核销且已过有效期的券不得计提，已计提金额须全额冲回。',
    '13. 已核销但台账无计提记录的，须按应计金额补提，不得以台账未列示为由忽略。',
    '14. 营销费用台账为门店申报数，不得直接作为应计依据，须以订单核销明细为底稿重算。',
    '15. 平台对账仅覆盖21家直营门店；加盟门店的平台补贴由加盟商与平台结算。',
    '16. 平台少结算或重复扣款应在次月结算中追补，不得直接冲减营销费用。',
    '17. 争议台账须逐条给出处理建议，跨店核销争议由实际核销门店所在区域承担。',
    '18. 活动预算不低于5万元须营销委员会审批，不低于10万元须总经理审批；先执行后审批须补审批说明。',
    '19. 跨期活动（结束日晚于2026-06-30）须按期间拆分费用，不得全额计入本期。',
    '20. 券核销量超过发放数量的，须核查超发来源并停止继续发放。',
    '21. 券限定门店与实际核销门店不一致的，费用由实际核销门店所在区域承担并追责。',
    '22. 会员日折扣券仅限每月8日、18日使用，非会员日核销的须单独列示。',
    '23. 营业收入按订单实收确认，券抵扣不得重复冲减收入；发现重复冲减的须重述收入。',
    '24. 必须按八轮顺序逐轮交付，每轮只完成一个任务，每轮结束提交本轮自检记录；上一轮未通过不得进入下一轮。',
    '25. 全部勾稽闭合后方可出具核心结论；任一项不闭合视为未完成。',
]
for t in constraints:
    P(doc, t, indent=False)

H(doc, '八、完成标准')
P(doc, f'八轮任务逐轮完成且逐轮自检通过。24家门店、{TX_N}行核销明细全部完成逐单复核，'
       f'公司应计营销费用{TOT_C}元（不含税），台账原始计提{TOT_A}元，台账较应计高{DIFF}元；'
       f'其中加盟店应剔除{AFRA}元、重复计提{DUP_N}条{DUP_AMT}元、过期未核销{EXP_N}条{EXP_AMT}元、'
       f'漏计提{MISS_N}条{MISS_AMT}元；制度版本比例调整{VER}元、含税换算调整{TAX}元；'
       f'平台应承担{P_ASSUME}元、实结算{P_ACTUAL}元、差额{P_DIFF}元；争议{DIS_N}条{DIS_AMT}元；'
       f'活动48项中有效{N_VALID}项、审批级别不足{APPROVE_BAD}项、超预算{OVER_N}项、跨期{N_CROSSACT}项；'
       f'券超发{OVER_CP}款、跨店核销{CROSS_N}条、会员日违规{MEM_BAD}条；'
       f'收入确认调整{INC_N}个门店月、合计调减{r2(-INC_AMT) if INC_AMT < 0 else INC_AMT}元；'
       f'{N_CHK}项勾稽全部闭合、不一致项为0。任一轮未完成或自检不通过，不得进入下一轮，整体不得视为完成。')

H(doc, '九、考点')
P(doc, '考点按数据治理、制度适用、逐单复核、异常券识别、平台对账、费用重算、活动有效性、执行偏差、收入调整、勾稽结论十组展开，共70项，'
       '其中数据型50项、推理链型20项（标注“推理链”者须同时引用两张以上表与具体制度条款）。', indent=False)

groups = [
    ('（一）数据治理与门店清单（考点1—7）', [
        '考点1：底稿中应当确认门店共24家，苏州5家、无锡4家、宁波4家、杭州4家、合肥4家、南昌3家，编号唯一且不跨城市重复。',
        '考点2：底稿中应当区分直营21家、加盟3家，加盟门店为WX04、HF04、NC03。',
        f'考点3：底稿中应当确认订单核销明细共{TX_N}行，等于24家门店乘6个月乘9单，且券码唯一、无重复核销单号。',
        f'考点4：底稿中应当确认营销费用台账共{LED_N}行，行数多于券码数，说明存在重复计提、过期券及漏计提。',
        '考点5：底稿中应当确认券规则共48款，覆盖美团、饿了么、抖音团购、点评、自有小程序五个平台。',
        '考点6：底稿中应当认定加盟门店的营销费用由加盟商承担，公司应计为0，但仍须纳入核销与执行偏差检查。',
        '考点7：底稿中应当确认输入附件共8个表格文件，且不含任何应计金额、复核结论或异常清单列。',
    ]),
    ('（二）制度适用与版本归属（考点8—17）', [
        '考点8：底稿中应当认定2026年1-4月适用PRC-MKT-v2.0（2025-07-01生效、2026-04-30废止）。',
        '考点9：底稿中应当认定2026年5-6月适用PRC-MKT-v3.0（2026-05-01起执行），不得适用v2.0。',
        '考点10：推理链——底稿中应当依据《制度生效与废止通知》确认v3.0于2026-03-20审议通过、2026-05-01起执行且不追溯，'
        '须同时引用通知的生效日期与v3.0制度的审议生效条款两处。',
        '考点11：底稿中应当识别制度汇编顶置红字“按v3.0统一核算”为文控误标，不得作为1-4月适用依据。',
        '考点12：推理链——底稿中应当驳回管理层M01“统一按v3.0重算6个月”，须同时引用生效通知的2026-05-01生效日'
        '与会议纪要“按核销时点归属版本”两处依据。',
        '考点13：推理链——底稿中应当按券核销时点而非计提月份归属版本，须同时引用会议纪要第2项与台账计提月份字段，'
        '说明台账按计提月统一套用v3.0的错误。',
        '考点14：底稿中应当确认v2.0下平台补贴券商家承担比例为美团50%、饿了么50%、点评50%、抖音团购70%、自有小程序100%。',
        '考点15：底稿中应当确认v3.0下平台补贴券商家承担比例为美团60%、饿了么60%、点评60%、抖音团购50%、自有小程序100%。',
        '考点16：推理链——底稿中应当识别1-4月美团、饿了么、点评因按v3.0的60%申报而多计，'
        '抖音团购因按v3.0的50%申报而少计，须跨台账、券规则与两版制度三处说明。',
        '考点17：推理链——底稿中应当认定台账同时存在“统一按v3.0比例”和“未做不含税换算”两类系统性错误，'
        '须引用会议纪要第2项、第4项与《财务口径说明》批注三处。',
    ]),
    ('（三）券核销逐单复核（考点18—27）', [
        f'考点18：底稿中应当对{TX_N}行核销明细逐单复核，每行给出券种、适用版本、商家承担比例与复核结论，不得按券编号汇总替代。',
        '考点19：推理链——底稿中应当按券种分类确定承担规则：满减券、代金券、平台补贴券按平台比例，折扣券与会员日折扣由商家全额承担，'
        '买赠按赠品成本价，须同时引用券规则“券种”列与制度承担条款两处。',
        '考点20：底稿中应当按“应计含税＝券抵扣×商家承担比例”计算，买赠按赠品成本价计入。',
        '考点21：底稿中应当按“应计不含税＝应计含税÷1.06”换算，不得直接以含税金额列示。',
        '考点22：底稿中应当确认平台补贴券适用平台比例，折扣券与会员日券比例恒为100%，不随制度版本变化。',
        '考点23：推理链——底稿中应当识别买赠券“券抵扣”为0但赠品成本大于0，不得因抵扣为0而漏计费用，'
        '须同时引用核销明细赠品成本列与制度买赠条款两处。',
        f'考点24：底稿中应当给出直营门店应计营销费用合计{TOT_C}元（不含税，容差±1元）。',
        f'考点25：底稿中应当认定营销费用台账原始计提合计{TOT_A}元。',
        f'考点26：推理链——底稿中应当说明台账较应计高{DIFF}元由加盟剔除、重复计提、过期未核销、版本比例、含税换算与漏计提六项构成，'
        '须跨台账、逐单复核与会议纪要三处逐项对应。',
        '考点27：推理链——底稿中应当以订单核销明细为底稿重算应计，不得以台账倒推，'
        '须引用会议纪要第3项“台账为门店申报数”与台账批注两处。',
    ]),
    ('（四）重复、超发与过期券识别（考点28—35）', [
        f'考点28：底稿中应当按券码分组识别重复计提{DUP_N}条，额外金额{DUP_AMT}元，逐条列明首次与重复计提记录。',
        f'考点29：底稿中应当识别过期未核销仍计提{EXP_N}条，金额{EXP_AMT}元，须全额冲回。',
        f'考点30：底稿中应当识别已核销未计提{MISS_N}条，应补提金额{MISS_AMT}元。',
        '考点31：推理链——底稿中应当按“券编号核销数量大于发放数量”识别券超发，逐款列明超发数量，'
        '须同时引用券规则发放数量与核销明细计数两处。',
        f'考点32：底稿中应当识别跨店核销{CROSS_N}条，逐条列明券限定门店与实际核销门店。',
        f'考点33：底稿中应当识别会员日违规核销{MEM_BAD}条，即非每月8日、18日使用会员日折扣券。',
        '考点34：推理链——底稿中应当按券码去重识别重复计提，不得按计提编号或门店汇总，'
        '须同时引用台账券码列与会议纪要第3项两处。',
        '考点35：推理链——底稿中应当以券规则有效期止与核销明细为准判定过期，'
        '有效期止早于2026-06-30且无核销记录的须全额冲回，须跨券规则与台账两处说明。',
    ]),
    ('（五）平台对账与争议处理（考点36—42）', [
        f'考点36：底稿中应当确认21家直营门店6个月平台应承担合计{P_ASSUME}元。',
        f'考点37：底稿中应当确认平台实际结算合计{P_ACTUAL}元。',
        f'考点38：底稿中应当确认平台实结算与应承担差额{P_DIFF}元，并逐门店逐月列示少结算与多结算。',
        f'考点39：底稿中应当确认争议台账共{DIS_N}条、争议金额{DIS_AMT}元，逐条给出处理建议。',
        '考点40：推理链——底稿中应当认定平台少结算须由营销部在次月结算中追补，不得直接冲减营销费用，'
        '须同时引用会议纪要第8项与平台结算单两处。',
        '考点41：推理链——底稿中应当按争议类型分别处理：平台驳回补结算、重复扣款退回、跨店核销由实际门店区域承担、'
        '补贴未到账次月追补，须引用会议纪要第8项与第9项两处。',
        '考点42：底稿中应当说明平台对账仅覆盖21家直营门店，加盟门店的平台补贴由加盟商与平台直接结算，不纳入公司对账范围。',
    ]),
    ('（六）费用归属与计提重算（考点43—51）', [
        '考点43：底稿中应当按“门店×月份”共144行重算费用归属，不得只给公司汇总数。',
        f'考点44：底稿中应当确认直营门店应计营销费用{TOT_C}元，加盟门店公司应计为0。',
        f'考点45：底稿中应当将加盟门店台账金额{AFRA}元全额剔除。',
        '考点46：底稿中应当编制“台账原始—剔除加盟—剔除重复—剔除过期—版本比例调整—含税换算调整—补提漏计提—应计”的差异桥。',
        '考点47：底稿中应当验证差异桥首尾闭合，桥接结果与应计合计一致，差异为0。',
        '考点48：推理链——底稿中应当依据制度“加盟门店”条款与会议纪要第5项剔除加盟费用，'
        '不得以“口径统一”为由并入公司合计（驳回M04）。',
        '考点49：推理链——底稿中应当分别列示版本比例调整与含税换算调整金额，不得合并为“口径差异”一笔，'
        '须跨逐单复核、两版制度与会议纪要第4项三处说明。',
        '考点50：推理链——底稿中应当将漏计提金额作为应计调增项列入差异桥，'
        '须同时引用逐单复核结果与台账缺号两处依据。',
        '考点51：底稿中应当拒绝以台账合计直接确认营销费用（驳回M02），并说明台账含重复、过期与漏计提。',
    ]),
    ('（七）活动有效性与审批（考点52—57）', [
        f'考点52：底稿中应当确认活动共48项，其中有效{N_VALID}项。',
        f'考点53：底稿中应当识别审批级别不足{APPROVE_BAD}项，即预算不低于5万元仅由区域经理审批，或预算不低于10万元未经总经理审批。',
        f'考点54：底稿中应当识别先执行后审批{N_AFTER}项，即审批日期晚于活动开始日。',
        f'考点55：底稿中应当识别超预算执行{OVER_N}项，实际费用高于预算。',
        f'考点56：底稿中应当识别跨期活动{N_CROSSACT}项，结束日晚于2026-06-30，费用须按期间拆分。',
        '考点57：推理链——底稿中应当按“预算不低于5万元营销委员会审批、不低于10万元总经理审批”判定审批有效性，'
        '须同时引用制度审批权限条款与会议纪要第7项两处。',
    ]),
    ('（八）门店执行偏差与责任（考点58—63）', [
        f'考点58：底稿中应当汇总门店执行偏差共{N_DEV}条，逐条列明对象、金额、责任认定与整改要求。',
        f'考点59：底稿中应当确认券超发{OVER_CP}款，并核查超发来源。',
        f'考点60：底稿中应当确认跨店核销{CROSS_N}条，费用由实际核销门店所在区域承担。',
        f'考点61：底稿中应当确认会员日违规核销{MEM_BAD}条，须按券规则限制核销时段。',
        '考点62：底稿中应当对每条偏差明确责任门店或责任部门，不得只描述现象不下结论（驳回M10）。',
        '考点63：推理链——底稿中应当依据会议纪要第9项与券规则“适用门店”列认定跨店核销责任，'
        '不得以“系统问题”为由由公司统一承担。',
    ]),
    ('（九）收入确认调整（考点64—67）', [
        '考点64：底稿中应当按“营业收入＝订单实收－退款”确认收入，不得按订单原价确认。',
        f'考点65：底稿中应当识别{INC_N}个门店月按原价确认收入、未冲减券抵扣，合计应调减收入{r2(-INC_AMT) if INC_AMT < 0 else INC_AMT}元。',
        '考点66：推理链——底稿中应当认定券抵扣不得作为销售折让重复冲减收入，'
        '须同时引用会议纪要第6项与门店销售汇总“账面净收入”列两处。',
        '考点67：底稿中应当识别冲正台账中“收入冲减-券抵扣”类型处理错误，该类型不得冲减收入。',
    ]),
    ('（十）勾稽与核心结论（考点68—70）', [
        f'考点68：底稿中应当完成{N_CHK}项勾稽并全部闭合，不一致项为0。',
        '考点69：底稿中应当给出一句话核心结论，包含门店数、应计营销费用、台账金额、差异、主要问题类型五项要素。',
        '考点70：底稿中应当统一金额单位为元两位小数、比例两位，且全部结论可回溯至附件逐条记录。',
    ]),
]
for title, items in groups:
    H(doc, title, 2)
    for t in items:
        P(doc, t, indent=False)

H(doc, '十、答题思路')
P(doc, '第一步是读取门店主数据、活动主数据与券规则，锁定24家门店、21家直营与3家加盟、48项活动与48款券。')
P(doc, '第二步是读取两版制度、生效通知与营销委员会会议纪要，确定1-4月适用v2.0、5-6月适用v3.0，'
       '并确认券码去重、过期冲回、加盟剔除、不含税列示与收入按实收确认等口径。')
P(doc, '第三步是以订单核销明细1296行为底稿逐单复核，按券种与核销时点确定商家承担比例，计算应计含税与不含税金额。')
P(doc, '第四步是按券码分组比对台账，识别重复计提、过期未核销与漏计提；按发放数量识别超发，按适用门店识别跨店核销，'
       '按核销日期识别会员日违规。')
P(doc, '第五步是核对平台应承担与实际结算，逐条处理争议。')
P(doc, '第六步是按门店乘月份重算费用归属，剔除加盟店，编制台账到应计的差异桥并验证闭合。')
P(doc, '第七步是判定48项活动的审批与执行有效性，汇总门店执行偏差与责任。')
P(doc, '第八步是核对门店销售汇总，按实收重述收入并测算调整额。')
P(doc, '第九步是完成全部勾稽，最后给出一句话核心结论。')

H(doc, '十一、标准答案')
P(doc, '标准答案共8个包，逐包覆盖制度适用与活动有效性、逐单复核、异常券识别、平台对账与争议、'
       '费用归属重算、执行偏差、收入调整与勾稽自检。', indent=False)
P(doc, concl['核心结论'], bold=True)
for k in ['支撑一（制度适用）', '支撑二（逐单复核）', '支撑三（台账质量）',
          '支撑四（平台与活动）', '支撑五（收入确认）', '风险提示']:
    if k in concl:
        P(doc, concl[k])
P(doc, f'关键结果：24家门店（直营21、加盟3）核销{TX_N}单，公司应计营销费用{TOT_C}元（不含税），'
       f'台账原始计提{TOT_A}元，台账较应计高{DIFF}元；重复计提{DUP_N}条、过期未核销{EXP_N}条、漏计提{MISS_N}条；'
       f'平台差额{P_DIFF}元；争议{DIS_N}条；收入确认调整{INC_N}个门店月；{N_CHK}项勾稽全部闭合，不一致项为0。')

doc.core_properties.author = '审计部'
NPARA = len(doc.paragraphs)
NCHAR = sum(len(p.text) for p in doc.paragraphs)
doc.save(DOCX)
scrub(DOCX, ('审计部-王梓', '合伙人-林志远', '14', '2026-07-13T02:41:27Z', '2026-07-15T01:18:36Z'),
      (2026, 7, 15, 10, 0, 0), kind='docx', app_ver='16.0000', doc_stats=(NCHAR, NPARA))

# 母版工作簿一并去除工具签名
scrub(SRC_Q, ('营销部-林晓', '审计部-王梓', '58', '2026-06-30T03:18:44Z', '2026-07-13T06:22:15Z'),
      (2026, 7, 13, 14, 30, 0), kind='xlsx', app_ver='16.0300')
scrub(SRC_A, ('审计部-王梓', '合伙人-林志远', '21', '2026-07-16T01:09:38Z', '2026-07-22T02:44:07Z'),
      (2026, 7, 22, 11, 0, 0), kind='xlsx', app_ver='16.0000')

print('done')
print('C', TOT_C, 'A', TOT_A, 'diff', DIFF, 'afra', AFRA)
print('dup', DUP_N, DUP_AMT, 'exp', EXP_N, EXP_AMT, 'miss', MISS_N, MISS_AMT)
print('acts', N_VALID, APPROVE_BAD, N_AFTER, OVER_N, N_CROSSACT)
print('abn', N_ABN, 'overcp', OVER_CP, 'cross', CROSS_N, 'mem', MEM_BAD, 'dev', N_DEV)
print('inc', INC_N, INC_AMT, 'checks', N_CHK, 'paras', NPARA, 'chars', NCHAR)
