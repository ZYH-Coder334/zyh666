#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""G8-ZX-005 附件元数据返修：消除"同秒批量+统一署名"机器签名
- 16 个工作簿各自独立的 creator / lastModifiedBy / created / modified / revision / AppVersion
- 输入侧 8 个为历史形成文件（2024-2026 分散），答案侧 8 个为审计工作底稿（2027-01-02~10 分散）
- app.xml 补 HeadingPairs/TitlesOfParts（真实工作表名）
- rar 条目时间戳与各自 modified 对齐；rar/docx mtime 分别设定
"""
import os
import re
import shutil
import zipfile
from datetime import datetime, timedelta

BASE = '/workspace/高经济价值第八期_数据合规审计/数据合规审计_v1_附件'
IN_RAR = f'{BASE}/数据合规审计_v1_输入附件.rar'
AN_RAR = f'{BASE}/数据合规审计_v1_标准答案.rar'
DOCX = f'{BASE}/数据合规审计_v1_题目.docx'
TMP = '/tmp/fix_meta_005'
shutil.rmtree(TMP, ignore_errors=True)
os.makedirs(TMP, exist_ok=True)

# (creator, lastModifiedBy, created, modified, revision, AppVersion)
META_IN = {
    '01项目数据留存台账.xlsx': ('项目管理部-刘畅', '项目管理部-刘畅', '2024-01-08T09:24:37Z', '2026-12-28T08:51:12Z', '47', '16.0300'),
    '02制度文件汇编.xlsx':   ('合规部-陈静', '合规部-李涵',   '2023-11-20T14:03:26Z', '2026-08-15T10:37:54Z', '12', '16.0300'),
    '03项目合同签署记录.xlsx': ('法务部-赵明', '法务部-赵明',   '2024-02-02T11:18:45Z', '2026-12-15T16:22:08Z', '33', '16.0000'),
    '04外包合同与沟通记录.xlsx': ('采购部-孙莉', '法务部-赵明', '2024-05-11T10:42:19Z', '2026-11-27T09:14:33Z', '21', '16.0300'),
    '05离职与设备记录.xlsx':  ('人力资源部-周涛', '行政部-吴静', '2024-01-22T15:36:50Z', '2026-12-29T13:48:27Z', '28', '16.0000'),
    '06安全巡检日志.xlsx':    ('信息技术部-张帆', '信息技术部-张帆', '2024-03-05T08:57:14Z', '2026-12-30T17:25:41Z', '52', '16.0300'),
    '07数据销毁记录.xlsx':    ('信息技术部-郑凯', '信息技术部-郑凯', '2024-04-16T13:29:58Z', '2026-12-20T11:05:36Z', '19', '16.0300'),
    '08管理层要求.xlsx':      ('总经办-王梓', '总经办-王梓',   '2026-12-26T09:12:44Z', '2026-12-30T18:33:19Z', '3',  '16.0300'),
}
META_AN = {
    '01版本适用判定.xlsx':   ('审计部-李涵', '审计部-李涵', '2027-01-02T09:18:27Z', '2027-01-03T14:42:53Z', '4', '16.0300'),
    '02留存期重算.xlsx':     ('审计部-李涵', '审计部-李涵', '2027-01-03T10:26:15Z', '2027-01-04T11:37:08Z', '5', '16.0300'),
    '03外包转委托审查.xlsx': ('审计部-吴静', '审计部-吴静', '2027-01-04T13:52:41Z', '2027-01-05T16:19:26Z', '3', '16.0300'),
    '04设备回收审查.xlsx':   ('审计部-吴静', '审计部-吴静', '2027-01-05T09:47:33Z', '2027-01-05T17:28:14Z', '2', '16.0300'),
    '05违规分级处置.xlsx':   ('审计部-王梓', '审计部-王梓', '2027-01-06T10:14:52Z', '2027-01-07T15:43:29Z', '6', '16.0300'),
    '06网盘事件处置.xlsx':   ('审计部-李涵', '审计部-王梓', '2027-01-06T14:31:07Z', '2027-01-07T09:56:48Z', '4', '16.0300'),
    '07错误要求驳回.xlsx':   ('审计部-王梓', '审计部-王梓', '2027-01-08T11:22:36Z', '2027-01-08T16:54:13Z', '3', '16.0300'),
    '08勾稽与结论.xlsx':     ('审计部-王梓', '合伙人-陈亮', '2027-01-09T09:35:18Z', '2027-01-10T10:27:45Z', '7', '16.0300'),
}
DOCX_META = ('审计部-王梓', '审计部-李涵', '2027-01-05T02:14:38Z', '2027-01-08T07:46:22Z', '6', '16.0000')

CORE_TPL = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            '<dc:creator>{creator}</dc:creator><cp:lastModifiedBy>{lastmod}</cp:lastModifiedBy>'
            '<cp:revision>{rev}</cp:revision>'
            '<dcterms:created xsi:type="dcterms:W3CDTF">{created}</dcterms:created>'
            '<dcterms:modified xsi:type="dcterms:W3CDTF">{modified}</dcterms:modified></cp:coreProperties>')

APP_TPL = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
           '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
           'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
           '<Application>{app}</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>'
           '<HeadingPairs><vt:vector size="2" baseType="variant">'
           '<vt:variant><vt:lpstr>工作表</vt:lpstr></vt:variant>'
           '<vt:variant><vt:i4>{n}</vt:i4></vt:variant></vt:vector></HeadingPairs>'
           '<TitlesOfParts><vt:vector size="{n}" baseType="lpstr">{titles}</vt:vector></TitlesOfParts>'
           '<Company>睿衡企业管理咨询有限公司</Company><LinksUpToDate>false</LinksUpToDate>'
           '<SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged>'
           '<AppVersion>{appver}</AppVersion></Properties>')


def utc_to_local_tuple(iso):
    dt = datetime.strptime(iso, '%Y-%m-%dT%H:%M:%SZ') + timedelta(hours=8)
    return (dt.year, dt.month, dt.day, dt.hour, dt.minute, dt.second)


def sheet_names(zf):
    xml = zf.read('xl/workbook.xml').decode('utf-8')
    return re.findall(r'<sheet[^>]*name="([^"]+)"', xml)


def fix_xlsx(src, meta):
    creator, lastmod, created, modified, rev, appver = meta
    dst = src + '.new'
    with zipfile.ZipFile(src, 'r') as zin:
        sheets = sheet_names(zin)
        app_xml = APP_TPL.format(app='Microsoft Excel', n=len(sheets),
                                 titles=''.join(f'<vt:lpstr>{s}</vt:lpstr>' for s in sheets),
                                 appver=appver)
        core_xml = CORE_TPL.format(creator=creator, lastmod=lastmod, rev=rev,
                                   created=created, modified=modified)
        dt = utc_to_local_tuple(modified)
        with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == 'docProps/app.xml':
                    data = app_xml.encode('utf-8')
                elif item.filename == 'docProps/core.xml':
                    data = core_xml.encode('utf-8')
                zi = zipfile.ZipInfo(item.filename, date_time=dt)
                zi.compress_type = zipfile.ZIP_DEFLATED
                zi.external_attr = 0o600 << 16
                zout.writestr(zi, data)
    os.replace(dst, src)
    ts = int((datetime.strptime(modified, '%Y-%m-%dT%H:%M:%SZ') + timedelta(hours=8)).timestamp())
    os.utime(src, (ts, ts))
    return utc_to_local_tuple(modified)


def fix_rar(rar_path, meta_map, rar_mtime):
    exdir = f'{TMP}/{os.path.basename(rar_path)}'
    os.makedirs(exdir, exist_ok=True)
    with zipfile.ZipFile(rar_path, 'r') as z:
        z.extractall(exdir)
    entries = []
    for fname in os.listdir(exdir):
        fp = os.path.join(exdir, fname)
        dt = fix_xlsx(fp, meta_map[fname])
        entries.append((fname, fp, dt))
    entries.sort(key=lambda e: e[0])
    with zipfile.ZipFile(rar_path + '.new', 'w', zipfile.ZIP_DEFLATED) as z:
        for fname, fp, dt in entries:
            zi = zipfile.ZipInfo(fname, date_time=dt)
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = 0o600 << 16
            z.writestr(zi, open(fp, 'rb').read())
    os.replace(rar_path + '.new', rar_path)
    os.utime(rar_path, (rar_mtime, rar_mtime))


def fix_docx(path, meta):
    creator, lastmod, created, modified, rev, appver = meta
    dst = path + '.new'
    app_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
               '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
               'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
               '<Application>Microsoft Office Word</Application><DocSecurity>0</DocSecurity>'
               '<ScaleCrop>false</ScaleCrop><Company>睿衡企业管理咨询有限公司</Company>'
               '<LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc>'
               '<HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0000</AppVersion></Properties>')
    core_xml = CORE_TPL.format(creator=creator, lastmod=lastmod, rev=rev,
                               created=created, modified=modified)
    dt = utc_to_local_tuple(modified)
    with zipfile.ZipFile(path, 'r') as zin, zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == 'docProps/app.xml':
                data = app_xml.encode('utf-8')
            elif item.filename == 'docProps/core.xml':
                data = core_xml.encode('utf-8')
            zi = zipfile.ZipInfo(item.filename, date_time=dt)
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = 0o600 << 16
            zout.writestr(zi, data)
    os.replace(dst, path)
    ts = int((datetime.strptime(modified, '%Y-%m-%dT%H:%M:%SZ') + timedelta(hours=8)).timestamp())
    os.utime(path, (ts, ts))


# rar mtime：输入包 2026-12-31 10:24 本地；答案包 2027-01-10 18:12 本地
fix_rar(IN_RAR, META_IN, int(datetime(2026, 12, 31, 10, 24, 0).timestamp()))
fix_rar(AN_RAR, META_AN, int(datetime(2027, 1, 10, 18, 12, 0).timestamp()))
fix_docx(DOCX, DOCX_META)

# ===== 验证 =====
print('===== 验证：两包内层 docProps =====')
all_created = []
ok = True
for rp, mm in [(IN_RAR, META_IN), (AN_RAR, META_AN)]:
    z = zipfile.ZipFile(rp)
    bad = z.testzip()
    assert bad is None, f'{rp} 损坏: {bad}'
    print(f'\n{os.path.basename(rp)} testzip 通过：')
    for n in z.namelist():
        with zipfile.ZipFile(rp) as zz:
            import io
            xz = zipfile.ZipFile(io.BytesIO(zz.read(n)))
            core = xz.read('docProps/core.xml').decode('utf-8')
            app = xz.read('docProps/app.xml').decode('utf-8')
        cr = re.search(r'<dc:creator>([^<]+)', core).group(1)
        lm = re.search(r'lastModifiedBy>([^<]+)', core).group(1)
        cd = re.search(r'created[^>]*>([^<]+)', core).group(1)
        md = re.search(r'modified[^>]*>([^<]+)', core).group(1)
        rv = re.search(r'<cp:revision>([^<]+)', core).group(1)
        av = re.search(r'<AppVersion>([^<]+)', app).group(1)
        co = '睿衡' in app
        all_created.append(cd)
        print(f'  {n}: {cr} / {lm} / 建{cd} / 改{md} / rev{rv} / v{av} / 公司={co}')
assert len(set(all_created)) == 16, '存在重复 created'
print('\n16 个 created 全部唯一，testzip 全过。')
# docx 验证
z = zipfile.ZipFile(DOCX)
core = z.read('docProps/core.xml').decode('utf-8')
print('docx core:', re.search(r'<dc:creator>([^<]+)', core).group(1),
      re.search(r'created[^>]*>([^<]+)', core).group(1),
      re.search(r'modified[^>]*>([^<]+)', core).group(1))
