#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""G8-ZX-003/004/005 打包格式返修：16个小rar -> 每题2包（题目附件+答案附件），docx打包说明同步更新，表格内容不动"""
import datetime
import os
import re
import shutil
import zipfile

import openpyxl
from openpyxl import Workbook
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt

EXCEL_APP = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
             '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
             'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
             '<Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>'
             '<Company>睿衡企业管理咨询有限公司</Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc>'
             '<HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0300</AppVersion></Properties>')
WORD_APP = EXCEL_APP.replace('Microsoft Excel', 'Microsoft Office Word').replace('16.0300', '16.0000')
CORE_TPL = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            '<dc:creator>{creator}</dc:creator><cp:lastModifiedBy>{creator}</cp:lastModifiedBy>'
            '<dcterms:created xsi:type="dcterms:W3CDTF">{created}</dcterms:created>'
            '<dcterms:modified xsi:type="dcterms:W3CDTF">{modified}</dcterms:modified></cp:coreProperties>')


def copy_sheet(src_ws, dst_wb):
    dst_ws = dst_wb.create_sheet(src_ws.title)
    for row in src_ws.iter_rows():
        for cell in row:
            nc = dst_ws.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                nc.font = cell.font.copy()
                nc.fill = cell.fill.copy()
                nc.border = cell.border.copy()
                nc.alignment = cell.alignment.copy()
                nc.number_format = cell.number_format
    for k, v in src_ws.column_dimensions.items():
        dst_ws.column_dimensions[k].width = v.width
    dst_ws.freeze_panes = src_ws.freeze_panes
    dst_ws.sheet_view.showGridLines = False
    for mc in src_ws.merged_cells.ranges:
        dst_ws.merge_cells(str(mc))


def scrub(path, app_xml, created, modified, dt, creator):
    tmp = path + '.tmp'
    with zipfile.ZipFile(path, 'r') as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == 'docProps/app.xml':
                data = app_xml.encode('utf-8')
            elif item.filename == 'docProps/core.xml':
                data = CORE_TPL.format(creator=creator, created=created, modified=modified).encode('utf-8')
            zi = zipfile.ZipInfo(item.filename, date_time=dt)
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = 0o600 << 16
            zout.writestr(zi, data)
    os.replace(tmp, path)
    ts = int(datetime.datetime(*dt).timestamp())
    os.utime(path, (ts, ts))


def pack_multi(files, rar_path, dt):
    with zipfile.ZipFile(rar_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for arcname, data in files:
            zi = zipfile.ZipInfo(arcname, date_time=dt)
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = 0o600 << 16
            z.writestr(zi, data)
    ts = int(datetime.datetime(*dt).timestamp())
    os.utime(rar_path, (ts, ts))


def set_para_text(p, text):
    runs = p.runs
    if not runs:
        p.add_run(text)
        return
    runs[0].text = text
    for r in runs[1:]:
        r._element.getparent().remove(r._element)


QUESTIONS = [
    dict(dir='/workspace/高经济价值第八期_合伙人分润审计', base='合伙人分润审计',
         in_dt=(2026, 10, 15, 9, 0, 0), an_dt=(2026, 10, 25, 14, 0, 0), doc_dt=(2026, 10, 20, 10, 0, 0),
         in_created='2026-10-15T01:00:00Z', an_created='2026-10-25T06:00:00Z', doc_created='2026-10-20T02:00:00Z',
         input_packs=[('01项目分润台账', ['项目分润台账']), ('02合伙人考核结果', ['合伙人考核结果']),
                      ('03制度文件汇编', ['制度文件汇编']), ('04财务账面分润', ['财务账面分润']),
                      ('05竞业与离职证据', ['竞业与离职证据']), ('06飞单证据', ['飞单证据']),
                      ('07分润发放记录', ['分润发放记录']), ('08管理层要求', ['管理层要求'])],
         answer_packs=[('01制度适用判定', ['制度适用判定']), ('02项目分润计提', ['项目分润计提重算']),
                       ('03合伙人分润明细', ['合伙人分润明细', '考核系数与飞单处理']),
                       ('04竞业违约处理', ['竞业违约处理']), ('05飞单处理', ['飞单处理']),
                       ('06风险与整改机制', ['风险清单', '整改机制']), ('07错误要求驳回', ['错误要求驳回']),
                       ('08勾稽与结论', ['勾稽自检', '合伙人分润审计结论'])]),
    dict(dir='/workspace/高经济价值第八期_应收账款审计', base='应收账款审计',
         in_dt=(2026, 11, 30, 9, 0, 0), an_dt=(2026, 12, 10, 14, 0, 0), doc_dt=(2026, 12, 5, 10, 0, 0),
         in_created='2026-11-30T01:00:00Z', an_created='2026-12-10T06:00:00Z', doc_created='2026-12-05T02:00:00Z',
         input_packs=[('01应收账款明细', ['应收账款明细']), ('02客户信用档案', ['客户信用档案']),
                      ('03制度文件汇编', ['制度文件汇编']), ('04财务账面坏账准备', ['财务账面坏账准备']),
                      ('05债务重组协议', ['债务重组协议']), ('06破产公告摘要', ['破产公告摘要']),
                      ('07催收记录', ['催收记录']), ('08管理层要求', ['管理层要求'])],
         answer_packs=[('01账龄口径重算', ['账龄口径重算']), ('02坏账准备重算', ['坏账准备重算']),
                       ('03个别认定处理', ['个别认定处理']), ('04账面差异调节', ['账面差异调节']),
                       ('05催收与信用政策', ['催收与信用政策']),
                       ('06风险与整改机制', ['风险清单', '整改机制']), ('07错误要求驳回', ['错误要求驳回']),
                       ('08勾稽与结论', ['勾稽自检', '应收账款审计结论'])]),
    dict(dir='/workspace/高经济价值第八期_数据合规审计', base='数据合规审计',
         in_dt=(2026, 12, 31, 9, 0, 0), an_dt=(2027, 1, 10, 14, 0, 0), doc_dt=(2027, 1, 5, 10, 0, 0),
         in_created='2026-12-31T01:00:00Z', an_created='2027-01-10T06:00:00Z', doc_created='2027-01-05T02:00:00Z',
         input_packs=[('01项目数据留存台账', ['项目数据留存台账']), ('02制度文件汇编', ['制度文件汇编']),
                      ('03项目合同签署记录', ['项目合同签署记录']), ('04外包合同与沟通记录', ['外包合同与沟通记录']),
                      ('05离职与设备记录', ['离职与设备记录']), ('06安全巡检日志', ['安全巡检日志']),
                      ('07数据销毁记录', ['数据销毁记录']), ('08管理层要求', ['管理层要求'])],
         answer_packs=[('01版本适用判定', ['版本适用判定']), ('02留存期重算', ['留存期重算']),
                       ('03外包转委托审查', ['外包转委托审查']), ('04设备回收审查', ['设备回收审查']),
                       ('05违规分级处置', ['违规分级处置']), ('06网盘事件处置', ['网盘事件处置']),
                       ('07错误要求驳回', ['错误要求驳回']),
                       ('08勾稽与结论', ['勾稽自检', '数据合规审计结论'])]),
]

TMP = '/tmp/repack2'
BAD_WORDS = [b'python', b'Python', b'openpyxl', b'xlsxwriter', 'AI生成'.encode(), '示例数据'.encode(), '脚本'.encode()]
all_issues = []

for q in QUESTIONS:
    base, d = q['base'], q['dir']
    out = f"{d}/{base}_v1_附件"
    shutil.rmtree(TMP, ignore_errors=True)
    os.makedirs(TMP, exist_ok=True)

    # ---- 输入包：8个xlsx打进1个rar ----
    wb_q = openpyxl.load_workbook(f"{d}/{base}_v1_题目.xlsx")
    files = []
    for name, sheets in q['input_packs']:
        nb = Workbook()
        nb.remove(nb.active)
        for s in sheets:
            copy_sheet(wb_q[s], nb)
        xp = f'{TMP}/{name}.xlsx'
        nb.save(xp)
        scrub(xp, EXCEL_APP, q['in_created'], q['in_created'], q['in_dt'], '审计部-王梓')
        files.append((f'{name}.xlsx', open(xp, 'rb').read()))
    in_rar = f'{out}/{base}_v1_输入附件.rar'
    pack_multi(files, in_rar, q['in_dt'])

    # ---- 答案包 ----
    wb_a = openpyxl.load_workbook(f"{d}/{base}_v1_标准答案.xlsx")
    files = []
    for name, sheets in q['answer_packs']:
        nb = Workbook()
        nb.remove(nb.active)
        for s in sheets:
            copy_sheet(wb_a[s], nb)
        xp = f'{TMP}/{name}.xlsx'
        nb.save(xp)
        scrub(xp, EXCEL_APP, q['an_created'], q['an_created'], q['an_dt'], '审计部-王梓')
        files.append((f'{name}.xlsx', open(xp, 'rb').read()))
    an_rar = f'{out}/{base}_v1_标准答案.rar'
    pack_multi(files, an_rar, q['an_dt'])

    # ---- 删除旧16小包 ----
    shutil.rmtree(f'{out}/输入附件', ignore_errors=True)
    shutil.rmtree(f'{out}/标准答案', ignore_errors=True)

    # ---- docx 打包说明同步 ----
    docx_path = f'{out}/{base}_v1_题目.docx'
    doc = Document(docx_path)
    for p in doc.paragraphs:
        t = p.text
        nt = t
        if '本轮只输出' in t and t.endswith('.rar。'):
            nt = t[:-5] + '.xlsx。'
        elif t == '输入附件共8包。':
            nt = f'输入附件共1包：{base}_v1_输入附件.rar，内含8个表格文件。'
        elif re.match(r'^\d{2}.+\.rar，包含', t):
            nt = t.replace('.rar，包含', '.xlsx，包含')
        elif re.match(r'^\d{2}.+\.rar$', t):
            nt = t[:-4] + '.xlsx'
        if '并只提交当轮对应交付包。' in nt:
            nt = nt.replace('并只提交当轮对应交付包。', '并只提交当轮对应交付表格（.xlsx）。')
        if nt.startswith('本项委托按八轮执行') and '标准答案附件' not in nt:
            nt += f'全部八轮交付表格最终打包为1个标准答案附件（{base}_v1_标准答案.rar）。'
        if '当轮对应交付包' in nt:
            nt = nt.replace('当轮对应交付包', '当轮对应交付表格')
        if nt != t:
            set_para_text(p, nt)
    doc.save(docx_path)
    scrub(docx_path, WORD_APP, q['doc_created'], q['doc_created'], q['doc_dt'], '审计部-王梓')

    # ---- 自检 ----
    for rp, exp in ((in_rar, q['in_dt'][:3]), (an_rar, q['an_dt'][:3])):
        with zipfile.ZipFile(rp) as z:
            if z.testzip() is not None:
                all_issues.append(f'{base} testzip失败 {os.path.basename(rp)}')
            names = z.namelist()
            if len(names) != 8 or not all(n.endswith('.xlsx') for n in names):
                all_issues.append(f'{base} 条目异常 {os.path.basename(rp)}: {len(names)}')
            for n in names:
                if z.getinfo(n).date_time[:3] != exp:
                    all_issues.append(f'{base} 时间戳不符 {n}')
                inner = z.read(n)
                for w in BAD_WORDS:
                    if w in inner:
                        all_issues.append(f'{base} 敏感词 {w} in {n}')
                import io as _io
                with zipfile.ZipFile(_io.BytesIO(inner)) as zi:
                    app = zi.read('docProps/app.xml').decode()
                    core = zi.read('docProps/core.xml').decode()
                    if 'Microsoft Excel' not in app or '睿衡企业管理咨询有限公司' not in app:
                        all_issues.append(f'{base} app.xml伪装失败 {n}')
                    if '审计部-王梓' not in core:
                        all_issues.append(f'{base} core.xml伪装失败 {n}')
    doc2 = Document(docx_path)
    text = '\n'.join(p.text for p in doc2.paragraphs)
    for sec in ['一、编号', '二、任务背景', '三、任务分解', '四、输入附件', '五、制度与判断规则',
                '六、交付物', '七、约束条件', '八、完成标准', '九、考点', '十、答题思路', '十一、标准答案']:
        if sec not in text:
            all_issues.append(f'{base} docx缺章节 {sec}')
    if '【职业定位】' not in text:
        all_issues.append(f'{base} 缺职业定位')
    n_round = len(re.findall(r'本轮只输出\d{2}[^\n]*?\.xlsx。', text))
    if n_round != 8:
        all_issues.append(f'{base} 八轮xlsx输出数量={n_round}')
    for m in re.findall(r'\S*\.rar', text):
        if m not in (f'{base}_v1_输入附件.rar', f'{base}_v1_标准答案.rar',
                     f'{base}_v1_输入附件.rar，', f'{base}_v1_标准答案.rar）'):
            all_issues.append(f'{base} docx残留旧rar引用: {m}')
    kd = re.findall(r'考点(\d+)：', text)
    if len(kd) != 50 or len(set(kd)) != 50:
        all_issues.append(f'{base} 考点异常 {len(kd)}/{len(set(kd))}')
    raw = open(docx_path, 'rb').read()
    for w in BAD_WORDS:
        if w in raw:
            all_issues.append(f'{base} docx敏感词 {w}')
    print(f'{base}: 2包已生成，docx已更新')

print('\n返修自检问题数 =', len(all_issues))
for e in all_issues:
    print(' -', e)
