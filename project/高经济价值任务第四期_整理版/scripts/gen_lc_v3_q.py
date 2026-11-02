#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""劳动合同合规审查v3清洁版题目附件：
删除任务说明/强制分批/法规摘要/评分矩阵/判定列，
新增劳动合同竞业条款表(8份违法)修复完备性。
全部为纯业务原始材料，无出题标记。"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

THIN = Side(style="thin", color="CCCCCC")
BORDER = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
HDR_FILL = PatternFill("solid", fgColor="2B5B84")
HDR_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=10)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

def ws_setup(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

def title_row(ws, row, text, span):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name="Arial", bold=True, size=12, color="2B5B84")

def hdr_row(ws, row, headers):
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=i, value=h)
        c.font = HDR_FONT; c.fill = HDR_FILL; c.alignment = CENTER; c.border = BORDER

def body_row(ws, r, values, total=False):
    for i, v in enumerate(values, 1):
        c = ws.cell(row=r, column=i, value=v)
        c.font = Font(name="Arial", size=9, bold=total)
        c.alignment = WRAP; c.border = BORDER
        if total:
            c.fill = PatternFill("solid", fgColor="E8EEF4")

wb = openpyxl.Workbook()

# === Sheet 1: 公司基本信息 ===
ws = wb.active; ws.title = "公司基本信息"
ws_setup(ws, [22, 50])
title_row(ws, 1, "华汇新零售基本信息", 2)
hdr_row(ws, 2, ["项目", "内容"])
info = [
    ["公司名称", "华汇新零售股份有限公司（脱敏代称）"],
    ["成立时间", "2015年3月"],
    ["主营业务", "连锁便利店"],
    ["门店数量", "68家"],
    ["员工总数", "680人"],
    ["裁员人数", "200人"],
    ["裁员占比", "29.4%"],
    ["当地最低工资标准", "2500元/月"],
    ["当地上年度职工月均工资", "8000元/月"],
    ["工时制度", "标准工时制、综合计算工时制、不定时工时制"],
    ["裁员原因", "经营战略调整，关闭华东区域3家门店"],
    ["补偿方案", "N+1，N按基本工资计算"],
    ["通知时间", "提前15日通知员工"],
    ["工会说明", "未向工会说明裁员方案"],
    ["劳动行政报告", "未向劳动行政部门报告"],
    ["优先留用名单", "未制定"],
]
r = 3
for row in info:
    body_row(ws, r, row); r += 1

# === Sheet 2: 劳动合同条款（80份） ===
# 删除"合规状态""备注"判定列，只保留客观字段
ws = wb.create_sheet("劳动合同条款")
ws_setup(ws, [8, 8, 12, 12, 10, 10, 14])
title_row(ws, 1, "劳动合同条款样本（80份）", 7)
hdr_row(ws, 2, ["合同编号", "员工编号", "合同期限", "试用期", "竞业期限", "竞业补偿金(元/月)", "是否签订单独竞业协议"])

contracts = []
cid = 1
emp = 101
# 12份试用期违法（试用期超法定上限）
illegal_probation = [
    ("6个月", "3个月"), ("6个月", "2个月"), ("1年", "3个月"), ("1年", "3个月"),
    ("2年", "4个月"), ("6个月", "3个月"), ("1年", "3个月"), ("2年", "4个月"),
    ("1年", "3个月"), ("6个月", "2个月"), ("1年", "3个月"), ("2年", "4个月"),
]
for dur, prob in illegal_probation:
    contracts.append([f"HT{cid:03d}", f"E{emp:03d}", dur, prob, "无", 0, "否"])
    cid += 1; emp += 1

# 8份劳动合同内含竞业条款违法（5份补偿金低于最低工资+3份未约定）
# HT013-HT017: 补偿金1200元（低于最低工资2500），未签订单独协议
for i in range(5):
    contracts.append([f"HT{cid:03d}", f"E{emp:03d}", "3年", "6个月", "2年", 1200, "否"])
    cid += 1; emp += 1
# HT018-HT020: 未约定补偿金，未签订单独协议
for i in range(3):
    contracts.append([f"HT{cid:03d}", f"E{emp:03d}", "3年", "6个月", "2年", 0, "否"])
    cid += 1; emp += 1

# 10份反直觉正常项（试用期=法定上限，合法边界）
normal_probation = [
    ("3年", "6个月"), ("1年", "2个月"), ("6个月", "1个月"), ("3年", "6个月"), ("1年", "2个月"),
    ("6个月", "1个月"), ("3年", "6个月"), ("1年", "2个月"), ("6个月", "1个月"), ("3年", "6个月"),
]
for dur, prob in normal_probation:
    contracts.append([f"HT{cid:03d}", f"E{emp:03d}", dur, prob, "无", 0, "否"])
    cid += 1; emp += 1

# 50份正常合同
for i in range(50):
    dur = "3年" if i % 3 == 0 else ("1年" if i % 3 == 1 else "2年")
    prob = "6个月" if dur == "3年" else ("2个月" if dur == "1年" else "1个月")
    contracts.append([f"HT{cid:03d}", f"E{emp:03d}", dur, prob, "无", 0, "否"])
    cid += 1; emp += 1

r = 3
for row in contracts:
    body_row(ws, r, row); r += 1
body_row(ws, r, ["合计", f"{len(contracts)}份", "", "", "", "", ""], total=True)

# === Sheet 3: 员工名册与社保（150人） ===
# 删除"裁员类别""备注判定"列，特殊人员信息放"备注"列用自然语言描述
ws = wb.create_sheet("员工名册与社保")
ws_setup(ws, [8, 8, 10, 14, 12, 12, 10, 10, 20])
title_row(ws, 1, "员工名册与社保缴纳基数（150人）", 9)
hdr_row(ws, 2, ["序号", "工号", "部门", "岗位", "入职日期", "月工资(元)", "社保基数(元)", "工时制度", "备注"])

employees = []
eid = 1
# 50人社保基数低于实际工资
for i in range(50):
    wage = [8000, 9000, 7500, 8500, 10000, 7000, 9500, 8200, 7800, 8800][i % 10]
    base = int(wage * 0.6)  # 按60%缴纳
    dept = ["门店", "配送", "采购", "财务", "行政"][i % 5]
    pos = ["店员", "配送员", "采购员", "会计", "行政专员"][i % 5]
    yr = 2016 + (i % 8)
    employees.append([eid, f"E{100+eid:03d}", dept, pos, f"{yr}-0{(i%9)+1}-15", wage, base, "标准工时", ""])
    eid += 1

# 6名禁止裁减对象（工伤2+孕期3+哺乳期1），信息放备注列自然语言
special = [
    [eid, f"E{100+eid:03d}", "门店", "店员", "2018-05-20", 6500, 6500, "标准工时", "2023年8月工伤，工伤认定书编号HW2023-0815"],
    [eid+1, f"E{100+eid+1:03d}", "门店", "店长", "2017-03-10", 9000, 9000, "标准工时", "怀孕3个月，预产期2024年4月"],
    [eid+2, f"E{100+eid+2:03d}", "配送", "配送员", "2019-06-15", 7500, 7500, "综合工时", "怀孕5个月，预产期2024年2月"],
    [eid+3, f"E{100+eid+3:03d}", "门店", "店员", "2018-11-08", 6500, 6500, "标准工时", "怀孕7个月，预产期2023年12月"],
    [eid+4, f"E{100+eid+4:03d}", "采购", "采购员", "2016-09-20", 8000, 8000, "标准工时", "哺乳期，孩子2023年6月出生"],
    [eid+5, f"E{100+eid+5:03d}", "门店", "店员", "2020-02-14", 6000, 6000, "标准工时", "2023年10月工伤，工伤认定书编号HW2023-1020"],
]
for row in special:
    employees.append(row)
eid += 6

# 4名应签未签无固定期限（连续工作满10年）
for i in range(4):
    yr = 2013 + i
    employees.append([eid, f"E{100+eid:03d}", "门店", "店长", f"{yr}-03-01", 9500, 9500, "标准工时", "连续工作满10年，公司未签订无固定期限合同"])
    eid += 1

# 2名试用期员工（在裁员名单中）
for i in range(2):
    employees.append([eid, f"E{100+eid:03d}", "门店", "店员", "2023-08-01", 5500, 5500, "标准工时", "试用期6个月，2024年1月到期"])
    eid += 1

# 88人正常员工补齐150人
for i in range(88):
    wage = [8000, 7500, 8500, 9000, 7000, 9500, 6500, 8800, 7800, 8200][i % 10]
    dept = ["门店", "配送", "采购", "财务", "行政", "人事", "IT", "运营"][i % 8]
    pos = ["店员", "配送员", "采购员", "会计", "行政专员", "HR专员", "运维", "运营专员"][i % 8]
    yr = 2015 + (i % 9)
    employees.append([eid, f"E{100+eid:03d}", dept, pos, f"{yr}-0{(i%9)+1}-10", wage, wage, "标准工时", ""])
    eid += 1

r = 3
for row in employees:
    body_row(ws, r, row); r += 1
body_row(ws, r, ["合计", f"{len(employees)}人", "", "", "", "", "", "", ""], total=True)

# === Sheet 4: 裁员方案与程序 ===
# 删除"法律要求""是否合规"判定列
ws = wb.create_sheet("裁员方案与程序")
ws_setup(ws, [8, 30, 30])
title_row(ws, 1, "公司拟定的裁员方案", 3)
hdr_row(ws, 2, ["序号", "方案内容", "说明"])
plan = [
    [1, "提前15日通知被裁减员工", "公司拟定通知时间为15日"],
    [2, "未向工会或全体职工说明裁员方案", "公司未进行工会说明程序"],
    [3, "未向劳动行政部门报告裁员方案", "公司未履行报告程序"],
    [4, "裁员名单含工伤员工2人、孕期员工3人、哺乳期员工1人", "共6名特殊人员"],
    [5, "补偿方案为N+1，N按基本工资计算", "公司拟定的补偿口径"],
    [6, "未制定优先留用人员名单", "公司未考虑优先留用"],
    [7, "试用期员工2人在裁员名单中", "公司未特殊考虑试用期员工"],
    [8, "4名连续工作满10年员工在裁员名单中", "公司未签订无固定期限合同"],
    [9, "门店关闭后员工统一裁减", "公司未考虑转岗安置"],
    [10, "补偿金一次性支付", "公司拟定支付方式"],
    [11, "未公布裁员标准和依据", "公司未公开裁员规则"],
]
r = 3
for row in plan:
    body_row(ws, r, row); r += 1

# === Sheet 5: 加班工资明细（80人） ===
# 删除"法定应发""差额"判定列，只保留公司已付金额，让模型自己算
ws = wb.create_sheet("加班工资明细")
ws_setup(ws, [8, 8, 14, 12, 10, 10, 10, 10, 10, 10])
title_row(ws, 1, "加班工资明细表（80人，单位：元）", 10)
hdr_row(ws, 2, ["序号", "工号", "工时制度", "日工资", "工作日加班(h)", "休息日加班(h)", "法定节假日加班(h)", "公司已付加班工资", "月份", "备注"])

overtime = []
oid = 1
# 30人标准工时制（公司按1倍计算，应1.5/2/3倍）
for i in range(30):
    dw = 250 + (i % 5) * 50  # 日工资250-450
    wd_ot = 20 + (i % 10)    # 工作日加班20-29h
    we_ot = 8 + (i % 8)      # 休息日加班8-15h
    ho_ot = i % 4            # 法定节假日加班0-3h
    # 公司按1倍计算（只付了加班小时数×日工资/8）
    paid = round((wd_ot + we_ot + ho_ot) * dw / 8, 2)
    overtime.append([oid, f"E{200+oid:03d}", "标准工时", dw, wd_ot, we_ot, ho_ot, paid, "2023-10", ""])
    oid += 1

# 25人综合计算工时制（公司未支付加班工资）
for i in range(25):
    dw = 280 + (i % 4) * 40
    wd_ot = 30 + (i % 10)  # 周期内超过法定标准
    overtime.append([oid, f"E{230+oid-30:03d}", "综合计算工时", dw, wd_ot, 0, 0, 0, "2023-10", ""])
    oid += 1

# 10人不定时工时制（公司未支付法定节假日加班工资）
for i in range(10):
    dw = 350 + (i % 3) * 50
    overtime.append([oid, f"E{255+oid-55:03d}", "不定时工时", dw, 0, 0, 8 + (i % 4), 0, "2023-10", ""])
    oid += 1

# 15人已正确计算加班工资（反直觉正常项）
for i in range(15):
    dw = 300 + (i % 4) * 50
    wd_ot = 10 + (i % 6)
    we_ot = 4 + (i % 3)
    ho_ot = i % 3
    # 正确计算：工作日1.5倍+休息日2倍+法定3倍
    paid = round((wd_ot * dw / 8 * 1.5) + (we_ot * dw / 8 * 2) + (ho_ot * dw / 8 * 3), 2)
    overtime.append([oid, f"E{265+oid-65:03d}", "标准工时", dw, wd_ot, we_ot, ho_ot, paid, "2023-10", ""])
    oid += 1

r = 3
for row in overtime:
    body_row(ws, r, row); r += 1
body_row(ws, r, ["合计", f"{len(overtime)}人", "", "", "", "", "", "", "", ""], total=True)

# === Sheet 6: 竞业限制协议清单（30份单独协议） ===
# 删除"补偿是否合规"判定列
ws = wb.create_sheet("竞业限制协议清单")
ws_setup(ws, [8, 8, 10, 14, 14, 14, 10])
title_row(ws, 1, "竞业限制协议清单（30份单独协议）", 7)
hdr_row(ws, 2, ["序号", "协议编号", "员工编号", "岗位", "竞业期限", "补偿金(元/月)", "违约金(元)"])

noncompete = []
ncid = 1
# 12份补偿金低于最低工资（1200元/月，违法）
for i in range(12):
    noncompete.append([ncid, f"JY{ncid:03d}", f"E{300+ncid:03d}", "店长", "2年", 1200, 50000])
    ncid += 1

# 8份未约定补偿金（违法）
for i in range(8):
    noncompete.append([ncid, f"JY{ncid:03d}", f"E{300+ncid:03d}", "采购员", "2年", 0, 30000])
    ncid += 1

# 5份补偿金等于最低工资（合法边界，反直觉正常项）
for i in range(5):
    noncompete.append([ncid, f"JY{ncid:03d}", f"E{300+ncid:03d}", "运营专员", "1年", 2500, 40000])
    ncid += 1

# 5份补偿金高于最低工资（合规）
for i in range(5):
    noncompete.append([ncid, f"JY{ncid:03d}", f"E{300+ncid:03d}", "区域经理", "2年", 3500, 80000])
    ncid += 1

r = 3
for row in noncompete:
    body_row(ws, r, row); r += 1
body_row(ws, r, ["合计", f"{len(noncompete)}份", "", "", "", "", ""], total=True)

# === Sheet 7: 会议纪要与邮件（7项管理层要求） ===
# 删除"涉及法律""正确处理方式"判定列
ws = wb.create_sheet("会议纪要与邮件")
ws_setup(ws, [18, 50])
title_row(ws, 1, "管理层会议纪要、邮件与电话备忘录", 2)
hdr_row(ws, 2, ["来源", "内容"])
minutes = [
    ["会议纪要-CEO", "经济补偿N按基本工资计算就行，不用算绩效和加班费，能省不少。"],
    ["邮件-HR总监", "综合计算工时制的员工不支付加班工资，他们周期内总工时没超就行。"],
    ["电话备忘录-CEO", "试用期员工不用给经济补偿，他们还没转正呢。"],
    ["邮件-CEO", "裁员方案不用向劳动行政部门报告了，200人而已，走个内部流程就行。"],
    ["会议纪要-财务VP", "竞业限制补偿金按最低工资的50%支付，也就是1250元，省成本。"],
    ["邮件-CEO", "把那几个孕期员工也裁了，反正门店要关了，她们也没法继续工作。"],
    ["电话备忘录-HR总监", "社保缴纳基数不用调整，按现在缴的就行，调整了反而引起注意。"],
    ["会议纪要-CEO", "不优先留用无固定期限合同的员工，统一裁减更简单。"],
    ["邮件-财务VP", "已经签了无固定期限合同的员工，按固定期限处理就行，不用额外补偿。"],
]
r = 3
for row in minutes:
    body_row(ws, r, row); r += 1

# === Sheet 8: 银行流水（裁员补偿支付记录） ===
ws = wb.create_sheet("银行流水")
ws_setup(ws, [10, 14, 14, 30])
title_row(ws, 1, "裁员补偿支付银行流水（单位：万元）", 4)
hdr_row(ws, 2, ["月份", "支付金额", "支付对象", "备注"])
flow = [
    ["10月", 0, "无", "尚未支付"],
    ["11月", 0, "无", "尚未支付"],
    ["12月", 0, "无", "尚未支付"],
    ["合计", 0, "", "公司尚未支付任何裁员补偿金"],
]
r = 3
for row in flow:
    body_row(ws, r, row); r += 1

# 保存
qpath = "/workspace/劳动合同合规审查v3_题目附件.xlsx"
wb.save(qpath)

# 统计
ht_total = len(contracts)
ht_probation_illegal = 12
ht_noncompete_illegal = 8
emp_total = len(employees)
emp_social_mismatch = 50
emp_special = 6
emp_indefinite = 4
emp_probation = 2
ot_total = len(overtime)
nc_total = len(noncompete)
nc_illegal = 12 + 8  # 12份低于最低工资 + 8份未约定
nc_legal_boundary = 5
nc_compliant = 5
total_noncompete_illegal = ht_noncompete_illegal + nc_illegal  # 8 + 20 = 28

print(f"=== v3题目附件统计 ===")
print(f"劳动合同: {ht_total}份")
print(f"  试用期违法: {ht_probation_illegal}份")
print(f"  劳动合同内竞业违法: {ht_noncompete_illegal}份")
print(f"员工名册: {emp_total}人")
print(f"  社保基数不一致: {emp_social_mismatch}人")
print(f"  禁止裁减对象: {emp_special}人")
print(f"  应签未签无固定期限: {emp_indefinite}人")
print(f"  试用期员工在裁员名单: {emp_probation}人")
print(f"加班工资: {ot_total}人")
print(f"竞业限制单独协议: {nc_total}份")
print(f"  违法(低于最低工资): 12份")
print(f"  违法(未约定): 8份")
print(f"  合法边界(=最低工资): {nc_legal_boundary}份")
print(f"  合规: {nc_compliant}份")
print(f"竞业违法合计: {total_noncompete_illegal}份 (劳动合同内{ht_noncompete_illegal} + 单独协议{nc_illegal})")
print(f"共 {len(wb.sheetnames)} 个sheet:")
for s in wb.sheetnames:
    print(f"  - {s}")
print(f"\n题目附件已保存: {qpath}")
