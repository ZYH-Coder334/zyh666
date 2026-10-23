#!/usr/bin/env python3
"""涉密资质甲级 v6 附件拆分：8输入+8答案，含新增表"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os, shutil, zipfile, io

THIN = Side(style="thin", color="CCCCCC")
BORDER = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
HDR_FILL = PatternFill("solid", fgColor="2F5597")
HDR_FONT = Font(name="宋体", bold=True, color="FFFFFF", size=10)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
BODY = Font(name="宋体", size=9)

SPLIT = '/tmp/miji_v6_split'
shutil.rmtree(SPLIT, ignore_errors=True)
os.makedirs(f'{SPLIT}/input', exist_ok=True)
os.makedirs(f'{SPLIT}/ans', exist_ok=True)

APP_XML = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
           '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
           'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
           '<Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity>'
           '<ScaleCrop>false</ScaleCrop><Company>华信安通科技有限公司</Company>'
           '<LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc>'
           '<HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0300</AppVersion></Properties>')


def core_xml(author, created, modified):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            f'<dc:creator>{author}</dc:creator><cp:lastModifiedBy>{author}</cp:lastModifiedBy>'
            f'<dcterms:created xsi:type="dcterms:W3CDTF">{created}</dcterms:created>'
            f'<dcterms:modified xsi:type="dcterms:W3CDTF">{modified}</dcterms:modified></cp:coreProperties>')


def scrub(path, author, created, modified):
    with zipfile.ZipFile(path) as zf:
        items = [(i, zf.read(i.filename)) for i in zf.infolist()]
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for info, data in items:
            if info.filename == 'docProps/app.xml':
                data = APP_XML.encode('utf-8')
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
        src = ws.column_dimensions[get_column_letter(i)].width
        nws.column_dimensions[get_column_letter(i)].width = src or 14
    nws.freeze_panes = "A2"
    wb.save(path)
    return nws.max_row - 1


def make_rar(files, rar_path):
    with zipfile.ZipFile(rar_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for f in files:
            zf.write(f, arcname=os.path.basename(f))


src_q = openpyxl.load_workbook('/workspace/涉密资质甲级_v6_题目.xlsx', data_only=True)
src_a = openpyxl.load_workbook('/workspace/涉密资质甲级_v6_标准答案.xlsx', data_only=True)

IN, AN = f'{SPLIT}/input', f'{SPLIT}/ans'
rows = {}
for sn in src_q.sheetnames:
    p = f'{IN}/{sn}.xlsx'
    rows[sn] = dump(src_q[sn], p)
    scrub(p, '综合管理部', '2025-08-18T02:15:00Z', '2025-08-20T06:40:00Z')
for sn in src_a.sheetnames:
    p = f'{AN}/{sn}.xlsx'
    rows[sn] = dump(src_a[sn], p)
    scrub(p, '合规咨询组', '2025-08-28T01:20:00Z', '2025-08-29T09:05:00Z')

OUT = '/workspace/涉密资质甲级_附件'
shutil.rmtree(f'{OUT}/输入附件', ignore_errors=True)
shutil.rmtree(f'{OUT}/标准答案', ignore_errors=True)
os.makedirs(f'{OUT}/输入附件', exist_ok=True)
os.makedirs(f'{OUT}/标准答案', exist_ok=True)

INPUT_PACKS = [
    ('01企业证照与股权结构.rar', ['企业基本信息', '股权结构穿透', '经营范围明细']),
    ('02人员名册与社保明细.rar', ['人员花名册', '社保缴纳明细', '社保基数对照表']),
    ('03高级职称与专业目录.rar', ['高级职称人员', '职称专业目录', '发证机关权限']),
    ('04涉密人员台账.rar', ['涉密人员台账']),
    ('05项目清单与发票批文.rar', ['项目清单', '发票清单', '涉密项目批文', '非密项目认定', '关联方名单']),
    ('06涉密场所与设备.rar', ['场所设备清单', '设备清单', '场所三铁检查']),
    ('07保密制度清单.rar', ['保密制度清单']),
    ('08管理层要求与资质标准.rar', ['管理层要求', '甲级资质条件', '集成收入口径', '收入确认规则', '保密审查规则']),
]

ANSWER_PACKS = [
    ('01差距分析报告.rar', ['差距分析结论']),
    ('02社保核验答案.rar', ['社保核验答案']),
    ('03高级职称审查答案.rar', ['高级职称审查答案']),
    ('04涉密人员审查答案.rar', ['涉密人员审查答案']),
    ('05集成收入核算答案.rar', ['集成收入核算答案']),
    ('06场所与制度审查答案.rar', ['场所审查答案', '设备权属审查答案', '保密制度审查答案']),
    ('07错误要求驳回.rar', ['错误要求驳回']),
    ('08勾稽与一致性自检.rar', ['数据一致性核验', '勾稽自检']),
]

for name, sheets in INPUT_PACKS:
    make_rar([f'{IN}/{s}.xlsx' for s in sheets], f'{OUT}/输入附件/{name}')
for name, sheets in ANSWER_PACKS:
    make_rar([f'{AN}/{s}.xlsx' for s in sheets], f'{OUT}/标准答案/{name}')

print('=== 输入附件 ===')
for name, sheets in INPUT_PACKS:
    print(f'  {name}: ' + '、'.join(f'{s}({rows[s]}行)' for s in sheets))
print('=== 标准答案 ===')
for name, sheets in ANSWER_PACKS:
    print(f'  {name}: ' + '、'.join(f'{s}({rows[s]}行)' for s in sheets))
print(f'\n输入 {len(INPUT_PACKS)}包/{sum(len(s) for _,s in INPUT_PACKS)}表；答案 {len(ANSWER_PACKS)}包/{sum(len(s) for _,s in ANSWER_PACKS)}表')
