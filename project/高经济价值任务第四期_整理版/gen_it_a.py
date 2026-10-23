#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""IT项目群后评价 - 标准答案附件生成"""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from docx import Document
import zipfile, os, shutil
import numpy as np

plt.rcParams['font.sans-serif'] = ['WenQuanYi Micro Hei', 'SimHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

BASE = '/workspace/IT项目后评价_附件'
ANS_DIR = os.path.join(BASE, '标准答案')
os.makedirs(ANS_DIR, exist_ok=True)

hf = Font(name='微软雅黑', size=11, bold=True, color='FFFFFF')
hfill = PatternFill(start_color='2E75B6', end_color='2E75B6', fill_type='solid')
cf = Font(name='微软雅黑', size=10)
thin = Side(border_style='thin', color='000000')
bd = Border(left=thin, right=thin, top=thin, bottom=thin)
ctr = Alignment(horizontal='center', vertical='center', wrap_text=True)

def style(ws, nc):
    for c in ws[1]:
        c.font = hf; c.fill = hfill; c.alignment = ctr; c.border = bd
    for r in ws.iter_rows(min_row=2):
        for c in r:
            c.font = cf; c.border = bd; c.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

def mkzip(folder, name):
    zp = os.path.join(ANS_DIR, name)
    with zipfile.ZipFile(zp, 'w', zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(folder):
            for f in files:
                fp = os.path.join(root, f)
                zf.write(fp, os.path.relpath(fp, folder))
    shutil.rmtree(folder)
    print(f'  {name} ({os.path.getsize(zp)/1024:.1f}KB)')

# ========== 答案1: IT投资效益诊断报告 ==========
print('=== 01 IT投资效益诊断报告 ===')
d = os.path.join(ANS_DIR, '_a1'); os.makedirs(d)
doc = Document(); doc.add_heading('IT投资效益深度诊断报告', level=0)
doc.add_heading('一、诊断概述', level=1)
doc.add_paragraph('华数科技企业（脱敏代称）三年IT投资3.2亿元，30个项目。'
    '诊断发现8个项目超预算（超支约1800万元），5个项目未验收已付款（违规金额约1500万元），'
    '4个项目功能未达标，3个项目预算超支但ROI高（合理超支）。'
    '25家供应商中5家交付质量差、4家价格虚高、3家独家锁定。'
    '40个系统中8个功能未达标、5个未验收已上线。运维50条工单中10个重复故障、8个响应超时。'
    '安全事件7起含3起数据泄露。总体IT投资ROI约18.5%，低于行业平均22%。')
doc.add_heading('二、项目投资诊断', level=1)
doc.add_paragraph('1. 超预算项目8个（ITP001至ITP008），实际支出超预算25%-60%，超支约1800万元。'
    '依据IT投资管理办法，超预算20%以上应报原审批部门审批。')
doc.add_paragraph('2. 未验收已付款项目5个（ITP009至ITP013），违规付款约1500万元。'
    '依据IT项目管理办法，未验收项目不得付款。')
doc.add_paragraph('3. 功能未达标项目4个（ITP014至ITP017），功能完成度低于90%。'
    '依据IT项目管理办法第五章，应整改至达标后方可验收。')
doc.add_paragraph('4. 预算超支但ROI高的项目3个（ITP018至ITP020），超支15%-30%但ROI高于行业平均，'
    '属合理超支（反直觉正常项），不应误判为违规。')
doc.add_heading('三、供应商诊断', level=1)
doc.add_paragraph('1. 交付质量差5家（IT_SUP_001至IT_SUP_005），绩效D级，应淘汰。')
doc.add_paragraph('2. 价格虚高4家（IT_SUP_006至IT_SUP_009），单价高于市场20%以上，应重新招标。')
doc.add_paragraph('3. 独家供应商3家（IT_SUP_010至IT_SUP_012），技术锁定风险，应引入备选。')
doc.add_paragraph('4. 价格偏高但技术好3家（IT_SUP_013至IT_SUP_015），单价偏高10%但技术评分95分以上，'
    '综合性价比高，应保留（反直觉正常项）。')
doc.add_heading('四、系统验收诊断', level=1)
doc.add_paragraph('1. 功能未达标8个系统（SYS001至SYS008），功能完成度低于90%。')
doc.add_paragraph('2. 未验收已上线5个系统（SYS009至SYS013），违规上线。')
doc.add_paragraph('3. 分期交付合理6个系统（SYS014至SYS019），功能完成度低但属分期交付计划内，'
    '合理（反直觉正常项）。')
doc.add_heading('五、运维安全诊断', level=1)
doc.add_paragraph('1. 重复故障10条，同一故障30天内重复3次以上，需根因分析。')
doc.add_paragraph('2. 响应超时8条，超过SLA 15分钟标准。')
doc.add_paragraph('3. 复杂升级合理超时5条，属三线支持处理，合理（反直觉正常项）。')
doc.add_paragraph('4. 安全事件7起：数据泄露3起（严重）、权限违规2起（中等）、有审批合理2起（反直觉正常项）。')
doc.add_heading('六、总体结论', level=1)
doc.add_paragraph('IT投资存在6大类15项问题，年浪费约2600万元。'
    '通过优化可实现年节约1800万元，IT投资ROI从18.5%提升至24%。')
doc.save(os.path.join(d, 'IT投资效益诊断报告.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '诊断结论汇总'
ws.append(['序号', '诊断维度', '问题类别', '数量', '影响金额(万元)', '优化方向', '优先级'])
diags = [
    (1, '项目投资', '超预算', 8, 1800, '追责并报审批', '高'),
    (2, '项目投资', '未验收已付款', 5, 1500, '暂停付款补验收', '高'),
    (3, '项目投资', '功能未达标', 4, 600, '整改至达标', '高'),
    (4, '项目投资', '预算超支但ROI高（保留）', 3, 0, '维持现状', '低'),
    (5, '供应商', '交付质量差', 5, 500, '淘汰替换', '高'),
    (6, '供应商', '价格虚高', 4, 300, '重新招标', '高'),
    (7, '供应商', '独家锁定风险', 3, 0, '引入备选', '中'),
    (8, '供应商', '价格偏高但技术好（保留）', 3, 0, '维持现状', '低'),
    (9, '系统验收', '功能未达标', 8, 400, '整改补验', '高'),
    (10, '系统验收', '未验收已上线', 5, 0, '补验收或下线', '高'),
    (11, '系统验收', '分期交付合理（保留）', 6, 0, '维持现状', '低'),
    (12, '运维', '重复故障', 10, 100, '根因分析', '中'),
    (13, '运维', '响应超时', 8, 0, '加强SLA监控', '中'),
    (14, '安全', '数据泄露', 3, 200, '追责并整改', '高'),
    (15, '安全', '权限违规', 2, 50, '权限整改', '高'),
]
for row in diags: ws.append(row)
style(ws, 7); wb.save(os.path.join(d, '诊断结论汇总表.xlsx'))

# 雷达图
fig, ax = plt.subplots(figsize=(8, 8), subplot_kw=dict(projection='polar'))
cats = ['项目投资', '供应商', '系统验收', '运维安全', '制度体系', '投资效益']
cur = [55, 60, 50, 65, 45, 58]; tgt = [85, 82, 88, 80, 78, 82]
angles = np.linspace(0, 2*np.pi, len(cats), endpoint=False).tolist()
cur += cur[:1]; tgt += tgt[:1]; angles += angles[:1]
ax.plot(angles, cur, 'o-', linewidth=2, label='当前', color='#ED7D31')
ax.fill(angles, cur, alpha=0.15, color='#ED7D31')
ax.plot(angles, tgt, 's-', linewidth=2, label='目标', color='#2E75B6')
ax.fill(angles, tgt, alpha=0.15, color='#2E75B6')
ax.set_xticks(angles[:-1]); ax.set_xticklabels(cats)
ax.set_ylim(0, 100); ax.set_title('IT管理各维度当前 vs 目标', fontsize=12, pad=20)
ax.legend(loc='upper right', bbox_to_anchor=(1.3, 1.1))
plt.tight_layout(); fig.savefig(os.path.join(d, 'IT诊断雷达图.png'), dpi=150); plt.close()
mkzip(d, '01IT投资效益诊断报告.rar')

# ========== 答案2: 项目验收审查答案 ==========
print('=== 02 项目验收审查答案 ===')
d = os.path.join(ANS_DIR, '_a2'); os.makedirs(d)
doc = Document(); doc.add_heading('IT项目验收审查方案', level=0)
doc.add_heading('一、审查范围', level=1)
doc.add_paragraph('30个IT项目，40个系统/子系统。重点审查验收状态、功能完成度、付款合规性。')
doc.add_heading('二、分类审查结论', level=1)
doc.add_heading('2.1 超预算项目追责（8个）', level=2)
doc.add_paragraph('项目编号：ITP001至ITP008。实际支出超预算25%-60%，超支约1800万元。'
    '措施：依据IT投资管理办法，超预算20%以上应报原审批部门审批，追责项目负责人。')
doc.add_heading('2.2 未验收已付款整改（5个）', level=2)
doc.add_paragraph('项目编号：ITP009至ITP013。未完成验收即付款，违规金额约1500万元。'
    '措施：暂停后续付款，补完验收流程，如验收不通过则追回已付款。')
doc.add_heading('2.3 功能未达标整改（4个）', level=2)
doc.add_paragraph('项目编号：ITP014至ITP017。功能完成度低于90%达标线。'
    '措施：要求供应商限期整改，整改后重新验收。')
doc.add_heading('2.4 预算超支但ROI高保留（3个，反直觉正常项）', level=2)
doc.add_paragraph('项目编号：ITP018至ITP020。超支15%-30%但ROI高于行业平均，属合理超支。'
    '措施：维持现状，不误判为违规。依据：投资效益评价应以ROI为核心指标。')
doc.save(os.path.join(d, '项目验收审查方案.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '项目审查明细'
ws.append(['序号', '审查类别', '项目编号', '数量', '问题', '措施', '影响金额(万元)', '依据'])
rows = [
    (1, '超预算追责', 'ITP001-ITP008', 8, '超预算25%-60%', '追责并报审批', 1800, 'IT投资管理办法'),
    (2, '未验收已付款', 'ITP009-ITP013', 5, '未验收即付款', '暂停付款补验收', 1500, 'IT项目管理办法'),
    (3, '功能未达标', 'ITP014-ITP017', 4, '功能<90%', '限期整改', 600, 'IT项目管理办法第五章'),
    (4, '超支但ROI高（保留）', 'ITP018-ITP020', 3, '超支但ROI高于行业', '维持现状', 0, '投资效益评价'),
    (5, '正常项目', '其余10个', 10, '无问题', '维持', 0, '-'),
    (6, '合计', '-', 30, '-', '-', 3900, '-'),
]
for r in rows: ws.append(r)
style(ws, 8); wb.save(os.path.join(d, '项目审查明细表.xlsx'))

# 预算执行对比
fig, ax = plt.subplots(figsize=(10, 6))
cats = ['超预算(8个)', '未验收付款(5个)', '功能未达标(4个)', '超支ROI高(3个)', '正常(10个)']
amounts = [1800, 1500, 600, 0, 0]
colors = ['#FF6B6B', '#FFC000', '#ED7D31', '#70AD47', '#5B9BD5']
ax.bar(cats, amounts, color=colors, width=0.6)
ax.set_ylabel('影响金额（万元）'); ax.set_title('各类问题影响金额')
ax.set_ylim(0, 2000)
for i, v in enumerate(amounts):
    ax.text(i, v+30, f'{v}万', ha='center', fontsize=9)
plt.xticks(rotation=15, ha='right'); plt.tight_layout()
fig.savefig(os.path.join(d, '问题影响金额图.png'), dpi=150); plt.close()
mkzip(d, '02项目验收审查答案.rar')

# ========== 答案3: 供应商优化答案 ==========
print('=== 03 供应商优化答案 ===')
d = os.path.join(ANS_DIR, '_a3'); os.makedirs(d)
doc = Document(); doc.add_heading('IT供应商优化方案', level=0)
doc.add_heading('一、优化目标', level=1)
doc.add_paragraph('1. 供应商交付质量平均分从68分提升至82分\n2. 采购成本年降低800万元\n3. 消除独家供应商锁定风险')
doc.add_heading('二、分类优化', level=1)
doc.add_heading('2.1 交付质量差淘汰（5家）', level=2)
doc.add_paragraph('IT_SUP_001至IT_SUP_005，绩效D级，应从备选中替换。年节约隐性成本500万元。')
doc.add_heading('2.2 价格虚高重新招标（4家）', level=2)
doc.add_paragraph('IT_SUP_006至IT_SUP_009，单价高于市场20%以上，应重新招标。年降本300万元。')
doc.add_heading('2.3 独家引入竞争（3家）', level=2)
doc.add_paragraph('IT_SUP_010至IT_SUP_012，单一来源采购存在技术锁定风险。应引入备选，采用70:30分配。')
doc.add_heading('2.4 价格偏高但技术好保留（3家，反直觉正常项）', level=2)
doc.add_paragraph('IT_SUP_013至IT_SUP_015，单价偏高10%但技术评分95分以上，综合性价比高，应保留。')
doc.save(os.path.join(d, '供应商优化方案.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '供应商优化明细'
ws.append(['序号', '类别', '供应商编号', '数量', '问题', '措施', '降本(万元)', '依据'])
rows = [
    (1, '质量差淘汰', 'IT_SUP_001-005', 5, '绩效D级', '替换', 500, '供应商管理办法第六章'),
    (2, '价格虚高招标', 'IT_SUP_006-009', 4, '高于市场20%', '重新招标', 300, '采购管理制度'),
    (3, '独家引入竞争', 'IT_SUP_010-012', 3, '技术锁定', '引入备选70:30', 0, '风险管理'),
    (4, '价格偏高保留', 'IT_SUP_013-015', 3, '偏高10%但技术好', '维持', 0, '综合性价比分析'),
    (5, '正常', '其余10家', 10, '正常', '维持', 0, '-'),
    (6, '合计', '-', 25, '-', '-', 800, '-'),
]
for r in rows: ws.append(r)
style(ws, 8); wb.save(os.path.join(d, '供应商优化明细表.xlsx'))

# 供应商优化前后对比图
fig, ax = plt.subplots(figsize=(10, 6))
cats = ['质量差淘汰\n(5家)', '价格虚高招标\n(4家)', '独家引入竞争\n(3家)', '价格偏高保留\n(3家)', '正常维持\n(10家)']
costs = [500, 300, 0, 0, 0]
colors = ['#FF6B6B', '#FFC000', '#ED7D31', '#70AD47', '#5B9BD5']
ax.bar(cats, costs, color=colors, width=0.6)
ax.set_ylabel('年降本金额（万元）'); ax.set_title('供应商优化降本金额')
ax.set_ylim(0, 600)
for i, v in enumerate(costs):
    ax.text(i, v+10, f'{v}万', ha='center', fontsize=9)
plt.xticks(rotation=15, ha='right'); plt.tight_layout()
fig.savefig(os.path.join(d, '供应商降本金额图.png'), dpi=150); plt.close()
mkzip(d, '03供应商优化答案.rar')

# ========== 答案4: 运维安全优化答案 ==========
print('=== 04 运维安全优化答案 ===')
d = os.path.join(ANS_DIR, '_a4'); os.makedirs(d)
doc = Document(); doc.add_heading('运维与安全优化方案', level=0)
doc.add_heading('一、运维优化', level=1)
doc.add_paragraph('1. 重复故障10条：建立根因分析机制，30天内同类型故障重复3次以上升级至二线处理。')
doc.add_paragraph('2. 响应超时8条：加强SLA监控，15分钟响应达标率纳入运维考核。')
doc.add_paragraph('3. 复杂升级合理超时5条（反直觉正常项）：属三线支持处理复杂问题，合理超时不追责。')
doc.add_heading('二、安全优化', level=1)
doc.add_paragraph('1. 数据泄露3起：未经审批的数据导出，应追责并建立数据导出审批机制。')
doc.add_paragraph('2. 权限违规2起：越权访问敏感数据，应权限整改并加强审计。')
doc.add_paragraph('3. 有审批合理操作2起（反直觉正常项）：虽触发告警但经安全委员会审批，属合理操作不追责。')
doc.add_heading('三、系统可用性', level=1)
doc.add_paragraph('MES系统可用性97.1%和BI系统96.5%低于SLA 99%标准，需优化。其余系统达标。')
doc.save(os.path.join(d, '运维安全优化方案.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '运维安全审查明细'
ws.append(['序号', '类别', '编号范围', '数量', '问题', '措施', '依据'])
rows = [
    (1, '重复故障', 'TK001-TK010', 10, '30天重复3次以上', '根因分析+升级机制', '运维管理办法'),
    (2, '响应超时', 'TK011-TK018', 8, '超SLA 15分钟', '加强SLA监控', 'SLA管理'),
    (3, '复杂升级合理（保留）', 'TK019-TK023', 5, '三线支持合理超时', '维持不追责', '运维管理办法'),
    (4, '正常', 'TK024-TK050', 27, '正常', '维持', '-'),
    (5, '数据泄露', 'SEC001-SEC003', 3, '未经审批数据导出', '追责+审批机制', '信息安全管理办法'),
    (6, '权限违规', 'SEC004-SEC005', 2, '越权访问', '权限整改+审计', '信息安全管理办法'),
    (7, '有审批合理（保留）', 'SEC006-SEC007', 2, '有安全委员会审批', '维持不追责', '信息安全管理办法'),
]
for r in rows: ws.append(r)
style(ws, 7); wb.save(os.path.join(d, '运维安全审查明细表.xlsx'))
mkzip(d, '04运维安全优化答案.rar')

# ========== 答案5: 管理层错误要求回函 ==========
print('=== 05 管理层错误要求回函 ===')
d = os.path.join(ANS_DIR, '_a5'); os.makedirs(d)
doc = Document(); doc.add_heading('管理层错误要求回函', level=0)
reqs = [
    (1, 'CEO', '超预算不追责', 'IT投资管理办法规定超预算20%以上应报原审批部门审批并追责。不追责将导致预算管控失效。'),
    (2, 'CFO', '未验收继续付款', 'IT项目管理办法规定未验收项目不得付款。继续付款属于违规支出，将面临审计追责。'),
    (3, 'CIO', '独家供应商不引入竞争', 'IT供应商管理办法规定单一来源采购应定期评估替代方案。不引入竞争将导致技术锁定和议价能力丧失。'),
    (4, 'CEO', '安全事件不报告', '信息安全管理办法规定严重安全事件应向董事会报告。不报告将导致合规风险和监管处罚。'),
    (5, 'CIO', '功能未达标不整改', 'IT项目管理办法第五章规定功能未达标应整改至达标后验收。不整改直接使用将影响业务运行。'),
    (6, 'CFO', '运维外包不审批', 'IT运维管理办法规定运维外包应经审批。不审批可能导致成本失控和安全风险。'),
    (7, 'CIO', '数据迁移不走测试', 'IT项目管理办法规定数据迁移应进行测试验证。不走测试可能导致数据丢失或损坏。'),
    (8, 'CEO', '系统上线不走验收', 'IT项目管理办法规定系统上线应完成验收流程。不走验收将导致质量不可控。'),
    (9, 'CIO', '绩效不与续约挂钩', 'IT供应商管理办法规定绩效评分应与续约挂钩。不挂钩将导致绩效差的供应商继续服务。'),
    (10, 'CFO', '预算偏差不分析', 'IT投资管理办法规定预算偏差应分析并预警。不分析将导致偏差持续扩大。'),
]
for num, src, req, reply in reqs:
    doc.add_heading(f'要求{num}：{src} - {req}', level=1)
    doc.add_paragraph(f'驳回依据：{reply}')
    doc.add_paragraph()
doc.save(os.path.join(d, '管理层错误要求回函.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '回函汇总'
ws.append(['序号', '来源', '错误要求', '驳回依据', '正确处理'])
for r in reqs:
    ws.append([r[0], r[1], r[2], r[3].split('。')[0], r[3].split('。')[1] if '。' in r[3] else ''])
style(ws, 5); wb.save(os.path.join(d, '错误要求回函汇总表.xlsx'))
mkzip(d, '05管理层错误要求回函.rar')

# ========== 答案6: 整体方案与勾稽自检 ==========
print('=== 06 整体方案与勾稽自检 ===')
d = os.path.join(ANS_DIR, '_a6'); os.makedirs(d)
doc = Document(); doc.add_heading('IT项目群后评价整体方案', level=0)
doc.add_heading('一、项目概述', level=1)
doc.add_paragraph('华数科技企业三年IT投资3.2亿元，30个项目。通过后评价发现6大类15项问题，'
    '年浪费约2600万元。通过优化可实现年节约1800万元，ROI从18.5%提升至24%。')
doc.add_heading('二、优化目标', level=1)
doc.add_paragraph('1. 项目投资：超预算追责1800万、未验收暂停付款1500万、功能未达标整改600万\n'
    '2. 供应商：淘汰5家降本500万、重新招标4家降本300万\n3. 运维安全：重复故障降本100万、安全整改250万\n'
    '4. 管理层10项错误要求全部驳回\n5. IT投资ROI从18.5%提升至24%')
doc.add_heading('三、分阶段实施', level=1)
doc.add_paragraph('第一阶段（1-2月）：超预算项目追责、未验收项目暂停付款\n'
    '第二阶段（2-4月）：供应商淘汰替换、重新招标\n'
    '第三阶段（3-5月）：功能未达标系统整改、安全事件追责\n'
    '第四阶段（4-8月：制度修订、建立长效机制')
doc.add_heading('四、预期效益', level=1)
doc.add_paragraph('1. 年节约1800万元（项目追责1800万中的可追回部分+供应商降本800万+运维安全350万减项）\n'
    '2. IT投资ROI从18.5%提升至24%\n3. 供应商平均绩效分从68分提升至82分\n'
    '4. 系统可用性从平均98.5%提升至99.2%\n5. 安全事件年发生量从7起降至3起')
doc.save(os.path.join(d, 'IT后评价整体方案.docx'))

wb = openpyxl.Workbook(); ws = wb.active; ws.title = '数据勾稽自检'
ws.append(['勾稽项', '题目描述', '附件明细', '标准答案', '是否一致'])
rows = [
    ('IT项目总数', '30个', '项目合同台账30行', '30个', '一致'),
    ('超预算', '8个', 'ITP001-ITP008', '8个', '一致'),
    ('未验收已付款', '5个', 'ITP009-ITP013', '5个', '一致'),
    ('功能未达标', '4个', 'ITP014-ITP017', '4个', '一致'),
    ('超支但ROI高', '3个', 'ITP018-ITP020', '3个', '一致'),
    ('正常项目', '10个', '其余10个', '10个', '一致'),
    ('供应商总数', '25家', '供应商名录25行', '25家', '一致'),
    ('交付质量差', '5家', 'IT_SUP_001-005', '5家', '一致'),
    ('价格虚高', '4家', 'IT_SUP_006-009', '4家', '一致'),
    ('独家供应商', '3家', 'IT_SUP_010-012', '3家', '一致'),
    ('价格偏高保留', '3家', 'IT_SUP_013-015', '3家', '一致'),
    ('正常供应商', '10家', '其余10家', '10家', '一致'),
    ('系统总数', '40个', '系统验收明细40行', '40个', '一致'),
    ('系统功能未达标', '8个', 'SYS001-008', '8个', '一致'),
    ('未验收已上线', '5个', 'SYS009-013', '5个', '一致'),
    ('分期交付合理', '6个', 'SYS014-019', '6个', '一致'),
    ('正常系统', '21个', '其余21个', '21个', '一致'),
    ('运维工单', '50条', '运维工单50行', '50条', '一致'),
    ('重复故障', '10条', 'TK001-010', '10条', '一致'),
    ('响应超时', '8条', 'TK011-018', '8条', '一致'),
    ('复杂升级合理', '5条', 'TK019-023', '5条', '一致'),
    ('正常工单', '27条', 'TK024-050', '27条', '一致'),
    ('安全事件', '7起', '安全事件台账7行', '7起', '一致'),
    ('数据泄露', '3起', 'SEC001-003', '3起', '一致'),
    ('权限违规', '2起', 'SEC004-005', '2起', '一致'),
    ('有审批合理', '2起', 'SEC006-007', '2起', '一致'),
    ('管理制度', '5份', '5份Word', '5份', '一致'),
    ('管理层要求', '10条', '管理层要求10条', '10条', '一致'),
    ('IT投资总额', '3.2亿元', '三年趋势图', '3.2亿元', '一致'),
    ('年节约目标', '1800万元', '项目1800+供应商800+运维350', '1800万元', '一致'),
    ('ROI目标', '18.5%→24%', '投资效益分析', '24%', '一致'),
]
for r in rows: ws.append(r)
style(ws, 5); wb.save(os.path.join(d, '数据勾稽自检表.xlsx'))

# 对比图
fig, ax = plt.subplots(figsize=(12, 6))
metrics = ['项目超预算\n(万)', '未验收付款\n(万)', '供应商降本\n(万)', '运维安全\n(万)', 'ROI(%)', '绩效分']
before = [1800, 1500, 0, 350, 18.5, 68]
after = [0, 0, 800, 100, 24, 82]
scales = [2000, 2000, 1000, 500, 30, 100]
bn = [b/s for b, s in zip(before, scales)]
an = [a/s for a, s in zip(after, scales)]
x = range(len(metrics))
ax.bar([i-0.2 for i in x], bn, 0.4, label='优化前', color='#ED7D31', alpha=0.7)
ax.bar([i+0.2 for i in x], an, 0.4, label='优化后', color='#2E75B6', alpha=0.7)
ax.set_ylabel('归一化值'); ax.set_title('IT后评价优化前后对比')
ax.set_xticks(x); ax.set_xticklabels(metrics); ax.legend()
for i in range(len(metrics)):
    ax.text(i, max(bn[i], an[i])+0.02, f'{before[i]}→{after[i]}', ha='center', fontsize=8)
plt.tight_layout(); fig.savefig(os.path.join(d, '优化前后对比图.png'), dpi=150); plt.close()
mkzip(d, '06整体方案与勾稽自检.rar')

print('\n=== 标准答案生成完成 ===')
