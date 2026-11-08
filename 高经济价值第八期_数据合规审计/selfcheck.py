#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""G8-ZX-005 三项自检：无AI痕迹 / 任务说明流程 / 完备性"""
import os, zipfile, re

BASE = '/workspace/高经济价值第八期_数据合规审计/数据合规审计_v1_附件'
BAD_WORDS = [b'python', b'Python', b'openpyxl', b'docx', 'AI生成'.encode(), '示例数据'.encode(),
             '脚本'.encode(), b'xlsxwriter', '人工智能生成'.encode()]
errors = []

# 1) 16包 testzip + 顶层无残留 + 单条目
rars = []
for sub in ('输入附件', '标准答案'):
    d = f'{BASE}/{sub}'
    fs = sorted(f for f in os.listdir(d) if f.endswith('.rar'))
    if len(fs) != 8:
        errors.append(f'{sub} 数量={len(fs)}')
    for f in fs:
        rars.append(f'{d}/{f}')
for rp in rars:
    with zipfile.ZipFile(rp) as z:
        if z.testzip() is not None:
            errors.append(f'testzip失败 {rp}')
        names = z.namelist()
        if len(names) != 1 or not names[0].endswith('.xlsx'):
            errors.append(f'rar条目异常 {rp}: {names}')
        if any('docProps' in n for n in names):
            errors.append(f'rar顶层残留docProps {rp}')
        inner = z.read(names[0])
        for w in BAD_WORDS:
            if w in inner:
                errors.append(f'内层敏感词 {w} in {rp}')
        # 内层 docProps 伪装
        with zipfile.ZipFile(__import__('io').BytesIO(inner)) as zi:
            app = zi.read('docProps/app.xml').decode()
            core = zi.read('docProps/core.xml').decode()
            if 'Microsoft Excel' not in app or '睿衡企业管理咨询有限公司' not in app:
                errors.append(f'app.xml伪装失败 {rp}')
            if '审计部-王梓' not in core:
                errors.append(f'core.xml伪装失败 {rp}')
            # 时间戳统一
            dts = {i.date_time for i in zi.infolist()}

# docx 自检
docx_path = f'{BASE}/数据合规审计_v1_题目.docx'
raw = open(docx_path, 'rb').read()
for w in BAD_WORDS:
    if w in raw and w != b'docx':
        errors.append(f'docx敏感词 {w}')
with zipfile.ZipFile(docx_path) as z:
    app = z.read('docProps/app.xml').decode()
    if 'Microsoft Office Word' not in app or '睿衡企业管理咨询有限公司' not in app:
        errors.append('docx app.xml伪装失败')

from docx import Document
doc = Document(docx_path)
text = '\n'.join(p.text for p in doc.paragraphs)
for sec in ['一、编号', '二、任务背景', '三、任务分解', '四、输入附件', '五、制度与判断规则',
            '六、交付物', '七、约束条件', '八、完成标准', '九、考点', '十、答题思路', '十一、标准答案']:
    if sec not in text:
        errors.append(f'docx缺章节 {sec}')
if '【职业定位】' not in text:
    errors.append('缺职业定位')
n_round = len(re.findall(r'本轮只输出\d{2}[^\n]*?\.rar。', text))
if n_round != 8:
    errors.append(f'八轮"本轮只输出"数量={n_round}')
kd = re.findall(r'考点(\d+)：', text)
if len(kd) != 50 or len(set(kd)) != 50:
    errors.append(f'考点数量/重复异常: {len(kd)}/{len(set(kd))}')
for phrase in ['AI', '人工智能', '示例数据', '脚本']:
    if phrase in text:
        errors.append(f'docx正文可疑措辞: {phrase}')

# 输入包无答案表名泄漏
leak_names = ['勾稽', '结论', '判定', '重算', '审查', '处置', '驳回']
for rp in rars:
    if '输入附件' in rp:
        with zipfile.ZipFile(rp) as z:
            inner = z.read(z.namelist()[0])
            with zipfile.ZipFile(__import__('io').BytesIO(inner)) as zi:
                wbxml = zi.read('xl/workbook.xml').decode()
                for ln in leak_names:
                    if ln in wbxml:
                        errors.append(f'输入包泄漏sheet名 {ln} in {os.path.basename(rp)}')

# 时间戳统一（rar文件本身）
import datetime
for rp in rars:
    with zipfile.ZipFile(rp) as z:
        dt = z.infolist()[0].date_time
        exp = (2026, 12, 31) if '输入附件' in rp else (2027, 1, 10)
        if dt[:3] != exp:
            errors.append(f'时间戳不符 {os.path.basename(rp)}: {dt}')

print('自检完成，问题数 =', len(errors))
for e in errors:
    print(' -', e)
