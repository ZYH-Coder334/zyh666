#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""政府投资建设项目全过程跟踪审计 - 题目附件生成"""
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import random

random.seed(42)

wb = openpyxl.Workbook()
wb.remove(wb.active)

header_font = Font(name='微软雅黑', size=11, bold=True)
header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
header_font_white = Font(name='微软雅黑', size=11, bold=True, color='FFFFFF')
cell_font = Font(name='微软雅黑', size=10)
thin = Side(border_style='thin', color='000000')
border = Border(left=thin, right=thin, top=thin, bottom=thin)
center = Alignment(horizontal='center', vertical='center', wrap_text=True)
left = Alignment(horizontal='left', vertical='center', wrap_text=True)

def style_ws(ws, ncols):
    for cell in ws[1]:
        cell.font = header_font_white
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.font = cell_font
            cell.alignment = left
            cell.border = border
    for col in range(1, ncols+1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 16

# ============ 1. 项目基本信息 ============
ws = wb.create_sheet('项目基本信息')
ws.append(['字段', '内容'])
info = [
    ('项目名称', '某市城市轨道交通3号线一期工程'),
    ('项目类型', '政府投资项目'),
    ('概算总投资', '45亿元'),
    ('预算总额', '48亿元'),
    ('预计结算', '52亿元'),
    ('超概算金额', '7亿元'),
    ('超概算比例', '15.6%'),
    ('建设内容', '车站18座、区间16个、车辆段1处'),
    ('建设期', '5年'),
    ('当前进度', '第4年'),
    ('资金来源', '财政资金70% + 银行贷款30%'),
    ('招标方式', '公开招标为主'),
    ('审计依据', '《政府投资条例》《招标投标法》《建筑法》《建设工程价款结算暂行办法》'),
    ('审计范围', '招投标合规性、合同履行、资金使用、工程变更管理、投资控制、绩效评价'),
    ('委托方', '某市审计局'),
    ('被审计单位', '某市轨道交通建设项目公司'),
]
for k, v in info:
    ws.append([k, v])
style_ws(ws, 2)
ws.column_dimensions['A'].width = 18
ws.column_dimensions['B'].width = 60

# ============ 2. 招投标文件汇总（35套） ============
ws = wb.create_sheet('招投标文件汇总')
ws.append(['招标编号', '项目名称', '采购方式', '投标人数量', '中标价(万元)', '评标委员会人数', '是否经财政审批', '备注'])

bids = []
# 正常20套
for i in range(1, 21):
    bids.append((f'ZB{i:03d}', f'轨道交通3号线土建工程第{i}标段', '公开招标', random.randint(5, 12),
                 round(random.uniform(8000, 35000), 2), random.choice([5, 7, 9]), '是', '正常公开招标'))
# 围标串标3套 ZB021-ZB023
for i in range(21, 24):
    bids.append((f'ZB{i:03d}', f'轨道交通3号线机电安装工程第{i-20}标段', '公开招标', 3,
                 round(random.uniform(12000, 18000), 2), 5, '是', '投标人报价偏差<0.1%，疑似围标串标'))
# 规避招标2套 ZB024-ZB025
for i in range(24, 26):
    bids.append((f'ZB{i:03d}', f'轨道交通3号线装修工程第{i-23}标段', '单一来源采购', 1,
                 round(random.uniform(400, 490), 2), 0, '否', '原项目500万拆分为5个99万项目，规避公开招标'))
# 评标违规5套 ZB026-ZB030
for i in range(26, 31):
    bids.append((f'ZB{i:03d}', f'轨道交通3号线信号系统工程第{i-25}标段', '公开招标', random.randint(4, 8),
                 round(random.uniform(15000, 25000), 2), random.choice([3, 4]), '是', '评标委员会人数不足5人或专业不达标'))
# 合法采购5套 ZB031-ZB035（反直觉正常项）
for i in range(31, 36):
    bids.append((f'ZB{i:03d}', f'轨道交通3号线专用设备采购第{i-30}标段',
                 random.choice(['单一来源采购', '竞争性谈判']), 1,
                 round(random.uniform(3000, 8000), 2), 5, '是', '经财政部门审批，单一来源/竞争性谈判合法'))

for b in bids:
    ws.append(b)
style_ws(ws, 8)
ws.column_dimensions['B'].width = 30
ws.column_dimensions['H'].width = 40

# ============ 3. 合同文件汇总（60份） ============
ws = wb.create_sheet('合同文件汇总')
ws.append(['合同编号', '项目名称', '承包人', '合同金额(万元)', '分包情况', '实际施工人', '是否有补充协议', '补充协议是否审批'])

contracts = []
# 正常30份
for i in range(1, 31):
    contracts.append((f'HT{i:03d}', f'土建工程第{i}标段合同', f'中铁第{i}工程局',
                      round(random.uniform(8000, 35000), 2), '无分包', f'中铁第{i}工程局',
                      random.choice(['有', '无']), '是'))
# 违法分包8份 HT031-HT038
for i in range(31, 39):
    contracts.append((f'HT{i:03d}', f'车站结构工程第{i-30}合同', f'中建第{i}局',
                      round(random.uniform(12000, 20000), 2), '主体工程分包给无资质单位', '某劳务队（无资质）',
                      '有', '是'))
# 转包挂靠3份 HT039-HT041
for i in range(39, 42):
    contracts.append((f'HT{i:03d}', f'区间隧道工程第{i-38}合同', f'中铁第{i}局',
                      round(random.uniform(15000, 25000), 2), '整体转包', f'某私人施工队',
                      '有', '是'))
# 合同变更未审批12份 HT042-HT053
for i in range(42, 54):
    contracts.append((f'HT{i:03d}', f'轨道铺设工程第{i-41}合同', f'中铁建第{i}局',
                      round(random.uniform(10000, 18000), 2), '无分包', f'中铁建第{i}局',
                      '有', '否'))
# 合法分包7份 HT054-HT060（反直觉正常项）
for i in range(54, 61):
    contracts.append((f'HT{i:03d}', f'机电安装工程第{i-53}合同', f'中铁电气化局',
                      round(random.uniform(10000, 20000), 2), '非主体劳务分包给有资质单位', f'中铁电气化局',
                      random.choice(['有', '无']), '是'))

for c in contracts:
    ws.append(c)
style_ws(ws, 8)
ws.column_dimensions['B'].width = 28
ws.column_dimensions['E'].width = 30
ws.column_dimensions['F'].width = 22

# ============ 4. 工程款支付记录（300条） ============
ws = wb.create_sheet('工程款支付记录')
ws.append(['支付编号', '合同编号', '支付金额(万元)', '累计支付比例(%)', '验收单编号', '收款账户', '是否合同约定账户', '支付日期'])

payments = []
pid = 1
# 正常240条
for _ in range(240):
    ht = f'HT{random.randint(1, 30):03d}'
    payments.append((f'ZF{pid:04d}', ht, round(random.uniform(500, 3000), 2),
                     round(random.uniform(30, 85), 1), f'YS{random.randint(1, 200):04d}',
                     f'合同约定账户{random.randint(1, 10)}', '是',
                     f'202{random.randint(1,3)}-{random.randint(1,12):02d}-{random.randint(1,28):02d}'))
    pid += 1
# 超付15条
for _ in range(15):
    ht = f'HT{random.randint(1, 30):03d}'
    payments.append((f'ZF{pid:04d}', ht, round(random.uniform(3000, 6000), 2),
                     round(random.uniform(100, 120), 1), f'YS{random.randint(1, 200):04d}',
                     f'合同约定账户{random.randint(1, 10)}', '是',
                     f'202{random.randint(1,3)}-{random.randint(1,12):02d}-{random.randint(1,28):02d}'))
    pid += 1
# 挪用资金5条
for _ in range(5):
    ht = f'HT{random.randint(1, 30):03d}'
    payments.append((f'ZF{pid:04d}', ht, round(random.uniform(2000, 5000), 2),
                     round(random.uniform(40, 70), 1), f'YS{random.randint(1, 200):04d}',
                     f'某私人账户{random.randint(1, 5)}', '否',
                     f'202{random.randint(1,3)}-{random.randint(1,12):02d}-{random.randint(1,28):02d}'))
    pid += 1
# 无依据支付8条
for _ in range(8):
    ht = f'HT{random.randint(1, 30):03d}'
    payments.append((f'ZF{pid:04d}', ht, round(random.uniform(1500, 4000), 2),
                     round(random.uniform(40, 75), 1), '无',
                     f'合同约定账户{random.randint(1, 10)}', '是',
                     f'202{random.randint(1,3)}-{random.randint(1,12):02d}-{random.randint(1,28):02d}'))
    pid += 1
# 合法预付款32条（反直觉正常项）
for _ in range(32):
    ht = f'HT{random.randint(1, 30):03d}'
    payments.append((f'ZF{pid:04d}', ht, round(random.uniform(800, 2500), 2),
                     round(random.uniform(10, 30), 1), '预付款（无验收单）',
                     f'合同约定账户{random.randint(1, 10)}', '是',
                     f'202{random.randint(1,3)}-{random.randint(1,12):02d}-{random.randint(1,28):02d}'))
    pid += 1

for p in payments:
    ws.append(p)
style_ws(ws, 8)
ws.column_dimensions['F'].width = 22
ws.column_dimensions['G'].width = 18

# ============ 5. 工程变更审批单（120份） ============
ws = wb.create_sheet('工程变更审批单')
ws.append(['变更编号', '变更内容', '变更金额(万元)', '发包方审批签字', '监理审批签字', '变更单价与合同单价对比', '变更原因', '紧急程度'])

changes = []
cid = 1
# 正常70份
for _ in range(70):
    ratio = round(random.uniform(0.8, 1.1), 2)
    changes.append((f'BG{cid:04d}', f'工程变更{cid}', round(random.uniform(50, 500), 2),
                    '已签', '已签', f'{ratio}:1（合理）',
                    random.choice(['设计优化', '地质条件变化', '业主要求']), '一般'))
    cid += 1
# 未经审批15份
for _ in range(15):
    changes.append((f'BG{cid:04d}', f'工程变更{cid}', round(random.uniform(100, 800), 2),
                    '未签', '已签', f'{round(random.uniform(0.8,1.1),2)}:1',
                    random.choice(['施工方自行变更', '现场签证']), '一般'))
    cid += 1
# 计价不合理20份
for _ in range(20):
    ratio = round(random.uniform(1.35, 1.8), 2)
    changes.append((f'BG{cid:04d}', f'工程变更{cid}', round(random.uniform(200, 1000), 2),
                    '已签', '已签', f'{ratio}:1（偏高30%以上）',
                    random.choice(['材料涨价', '工艺调整']), '一般'))
    cid += 1
# 非必要变更8份
for _ in range(8):
    changes.append((f'BG{cid:04d}', f'工程变更{cid}', round(random.uniform(80, 300), 2),
                    '已签', '已签', f'{round(random.uniform(0.9,1.1),2)}:1',
                    '施工方自行优化（属施工方义务）', '一般'))
    cid += 1
# 紧急抢险变更7份（反直觉正常项）
for _ in range(7):
    changes.append((f'BG{cid:04d}', f'工程变更{cid}', round(random.uniform(150, 600), 2),
                    '后补', '已签', f'{round(random.uniform(0.9,1.1),2)}:1',
                    '紧急抢险（先施工后补审批）', '紧急'))
    cid += 1

for c in changes:
    ws.append(c)
style_ws(ws, 8)
ws.column_dimensions['B'].width = 16
ws.column_dimensions['F'].width = 22
ws.column_dimensions['G'].width = 30

# ============ 6. 概算预算对比表 ============
ws = wb.create_sheet('概算预算对比表')
ws.append(['分项工程', '概算金额(万元)', '预算金额(万元)', '差额(万元)', '超支比例(%)', '超支原因'])

items = [
    ('车站工程', 180000, 195000, 15000, 8.3, '地质条件变化导致围护结构增加'),
    ('区间隧道工程', 120000, 128000, 8000, 6.7, '盾构掘进难度超预期'),
    ('轨道工程', 40000, 42000, 2000, 5.0, '钢轨价格上涨'),
    ('机电安装工程', 50000, 55000, 5000, 10.0, '设备选型升级'),
    ('装修工程', 25000, 28000, 3000, 12.0, '装修标准提高'),
    ('车辆段工程', 20000, 18000, -2000, -10.0, '优化设计节约'),
    ('预备费', 15000, 14000, -1000, -6.7, '预备费动用较少'),
    ('合计', 450000, 480000, 30000, 6.7, '主要因车站和机电工程超支'),
]
for item in items:
    ws.append(item)
style_ws(ws, 6)
ws.column_dimensions['A'].width = 20
ws.column_dimensions['F'].width = 35

# ============ 7. 资金拨付台账（200条） ============
ws = wb.create_sheet('资金拨付台账')
ws.append(['拨付编号', '申请单位', '申请金额(万元)', '拨付金额(万元)', '拨付进度(%)', '是否按进度', '超额比例(%)', '拨付日期'])

funds = []
fid = 1
units = ['项目公司', '土建项目部', '机电项目部', '轨道项目部', '装修项目部']
# 正常170条
for _ in range(170):
    amt = round(random.uniform(500, 5000), 2)
    funds.append((f'BF{fid:04d}', random.choice(units), amt, amt,
                  round(random.uniform(30, 90), 1), '是', 0,
                  f'202{random.randint(1,3)}-{random.randint(1,12):02d}-{random.randint(1,28):02d}'))
    fid += 1
# 未按进度10条
for _ in range(10):
    req = round(random.uniform(1000, 5000), 2)
    pay = round(req * random.uniform(0.5, 0.8), 2)
    funds.append((f'BF{fid:04d}', random.choice(units), req, pay,
                  round(random.uniform(30, 70), 1), '否', 0,
                  f'202{random.randint(1,3)}-{random.randint(1,12):02d}-{random.randint(1,28):02d}'))
    fid += 1
# 拨付超额15条
for _ in range(15):
    req = round(random.uniform(1000, 5000), 2)
    pay = round(req * random.uniform(1.1, 1.3), 2)
    funds.append((f'BF{fid:04d}', random.choice(units), req, pay,
                  round(random.uniform(40, 80), 1), '否',
                  round((pay/req - 1)*100, 1),
                  f'202{random.randint(1,3)}-{random.randint(1,12):02d}-{random.randint(1,28):02d}'))
    fid += 1
# 合法预付5条（反直觉正常项）
for _ in range(5):
    req = round(random.uniform(2000, 6000), 2)
    pay = round(req * random.uniform(1.05, 1.15), 2)
    funds.append((f'BF{fid:04d}', random.choice(units), req, pay,
                  round(random.uniform(10, 25), 1), '是',
                  round((pay/req - 1)*100, 1),
                  f'202{random.randint(1,3)}-{random.randint(1,12):02d}-{random.randint(1,28):02d}'))
    fid += 1

for f in funds:
    ws.append(f)
style_ws(ws, 8)
ws.column_dimensions['B'].width = 16

# ============ 8. 会议纪要与管理层要求（10条） ============
ws = wb.create_sheet('会议纪要与管理层要求')
ws.append(['序号', '来源', '内容'])

reqs = [
    (1, '分管副市长', '围标串标项目不追究，继续履行合同'),
    (2, '建委主任', '违法分包合同不处理，维持现状'),
    (3, '财政局副局长', '超付工程款不追回，内部消化'),
    (4, '项目公司总经理', '未经审批的变更予以认可，不调减'),
    (5, '发改委副主任', '概算超支不向市政府报告，自行解决'),
    (6, '审计局某处长', '绩效评价只写优点，不写问题'),
    (7, '项目公司财务总监', '资金挪用问题不审计，不追究'),
    (8, '建委招标办主任', '评标违规问题不披露，内部处理'),
    (9, '项目公司总工', '非必要变更全部认可，不剔除'),
    (10, '项目公司总经理', '合同变更未审批也予以认可，不要求补充'),
]
for r in reqs:
    ws.append(r)
style_ws(ws, 3)
ws.column_dimensions['B'].width = 20
ws.column_dimensions['C'].width = 55

wb.save('/workspace/政府投资项目审计_题目附件.xlsx')
print('题目附件生成完成')
