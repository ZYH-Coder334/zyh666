#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""G8-ZX-006 投标报价与中标合规审计 v1 —— 题目数据生成器
升级设计：120 投标项目 + 8 表约 1000 行 + 3 重显著性反转 + 强制跨 4 表判定链
"""
from datetime import date, timedelta
import random
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

random.seed(20261109)

BASE = '/workspace/高经济价值第八期_投标报价合规审计'
OUT = f'{BASE}/投标报价合规审计_v1_题目.xlsx'

# ====== 样式 ======
HDR_FILL = PatternFill('solid', fgColor='2F5597')
HDR_FONT = Font(name='宋体', bold=True, color='FFFFFF', size=10)
BODY = Font(name='宋体', size=9)
THIN = Side(style='thin', color='CCCCCC')
BORDER = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
WRAP = Alignment(wrap_text=True, vertical='top')
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)

START = date(2025, 1, 1)
CUTOFF = date(2027, 1, 15)   # 数据截止日


def stylize(ws, widths=None):
    for c in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = HDR_FONT
        cell.fill = HDR_FILL
        cell.alignment = CENTER
        cell.border = BORDER
        ws.column_dimensions[get_column_letter(c)].width = (widths or {}).get(c, 14)
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            cell.font = BODY
            cell.alignment = WRAP
            cell.border = BORDER
    ws.freeze_panes = 'A2'
    ws.sheet_view.showGridLines = False


def r4(x): return round(x + 0.0, 4)


# ====== 客户清单（25 家，其中 18 家备案战略客户、7 家非战略） ======
CUSTOMERS = [
    ('星辉传媒', '战略已备案'),
    ('青云文旅', '战略已备案'),
    ('宏图建筑', '战略已备案'),
    ('安捷物流', '战略已备案'),
    ('金陵商贸', '战略已备案'),
    ('康泰医疗', '战略已备案'),
    ('星桥物流', '战略已备案'),
    ('远洋地产', '战略已备案'),
    ('天合能源', '战略已备案'),
    ('金谷农业', '战略已备案'),
    ('绿源环保', '战略已备案'),
    ('优品科技', '战略已备案'),
    ('恒通金融', '战略已备案'),
    ('通达汽车', '战略已备案'),
    ('锐捷电子', '战略已备案'),
    ('智联教育', '战略已备案'),
    ('明华制造', '战略已备案'),
    ('泰盛化工', '战略已备案'),
    ('华盛贸易', '普通客户'),
    ('众诚实业', '普通客户'),
    ('东方传媒', '普通客户'),
    ('恒远物流', '普通客户'),
    ('新源科技', '普通客户'),
    ('博瑞医疗', '普通客户'),
    ('盛世餐饮', '普通客户'),
]
STRAT = set(c for c, t in CUSTOMERS if t == '战略已备案')

# ====== 基础项目参数 ======
proj_types = ['战略咨询', '组织变革', '数字化转型', '供应链优化', '财务尽调', '税务筹划', '人力咨询']
consultants = ['李涵', '张帆', '王梓', '刘畅', '陈亮', '赵明', '孙莉', '周涛', '吴静', '郑凯',
               '冯雪', '褚鹏', '卫华', '蒋萌', '沈悦', '韩磊', '杨帆', '朱倩', '秦朗', '尤佳']
partners = ['PA', 'PB', 'PC', 'PD', 'PE', 'PF']

wb = openpyxl.Workbook()
wb.remove(wb.active)


# ==============================================================
# Sheet 1: 投标项目总台账（120 行，故意混入脏数据）
# 脏数据设计：3 行项目号重复、5 行客户名格式错、8 行缺结束日期
# ==============================================================
ws1 = wb.create_sheet('投标项目总台账')
hdrs1 = ['项目编号', '客户名称', '战略客户(系统登记)', '项目类型', '主业务合同签约日',
         '项目结束日', '系统报价(万元)', '账面报价(万元)', '中标价(万元)',
         '利润类型', '合同适用版本(系统登记)', '签约合伙人', '顾问负责人', '备注']
ws1.append(hdrs1)

proj_data = []
for i in range(1, 121):
    code = f'P{i:03d}'
    cust, cust_t = random.choice(CUSTOMERS)
    ptype = random.choice(proj_types)
    sign_year = random.choice([2024, 2024, 2025, 2025, 2025, 2026])   # 比例：2024 约 33%、2025 50%、2026 17%
    sign_date = date(sign_year, random.randint(1, 12), random.randint(1, 28))
    proj_dur = random.randint(30, 240)
    end_date = sign_date + timedelta(days=proj_dur)
    # 报价系统值（故意让部分项目系统旧版本）
    base_cost = random.uniform(50, 300)
    margin = random.choice([0.20, 0.25, 0.30, 0.30, 0.35, 0.40])
    sys_price = r4(base_cost * (1 + margin))
    book_price = sys_price if random.random() > 0.15 else r4(sys_price * random.choice([0.88, 0.92, 1.08, 1.12]))
    # 中标价：大部分=账面报价，少数低价中标（战略客户 15% 折扣）
    if cust in STRAT and random.random() < 0.25:
        bid_price = r4(book_price * 0.85)
        margin_type = '战略折扣'
    elif random.random() < 0.08:
        bid_price = r4(book_price * random.choice([0.72, 0.78, 0.82]))
        margin_type = '低价中标'
    else:
        bid_price = book_price
        margin_type = '标准'
    # 系统登记的版本（故意让部分项目系统登记错误）
    sys_ver = 'v2.0' if sign_year >= 2025 else 'v1.0'
    if random.random() < 0.18:
        sys_ver = 'v2.0' if sys_ver == 'v1.0' else 'v1.0'   # 错登
    partner = random.choice(partners)
    consul = random.choice(consultants)
    proj_data.append([code, cust, '是' if cust in STRAT else '否', ptype,
                      str(sign_date), str(end_date), sys_price, book_price, bid_price,
                      margin_type, sys_ver, partner, consul, ''])

# 脏数据注入
# 脏1：3 行项目号重复（P011、P047、P082 各再出现 1 次，数据略不同）
for dup in ['P011', 'P047', 'P082']:
    orig = next(p for p in proj_data if p[0] == dup)
    dup_row = orig.copy()
    dup_row[7] = r4(orig[7] * 1.05)     # 账面报价 5% 差异
    dup_row[13] = '重复行（数据来自系统备份）'
    proj_data.append(dup_row)

# 脏2：5 行客户名格式错误（多余空格/括号/缩写）
bad_custs = {
    'P023': '安捷物流有限公司 ',
    'P056': '(宏图建筑)',
    'P071': '  康泰医疗集团',
    'P089': '星桥物流(总部)',
    'P104': ' 远洋地产  ',
}
for i, r in enumerate(proj_data):
    if r[0] in bad_custs:
        proj_data[i][1] = bad_custs[r[0]]
        proj_data[i][13] = '客户名格式错误'

# 脏3：8 行缺结束日期
missing_end = ['P017', 'P029', 'P043', 'P058', 'P072', 'P086', 'P099', 'P115']
for i, r in enumerate(proj_data):
    if r[0] in missing_end:
        proj_data[i][5] = ''
        proj_data[i][13] = '结束日期缺失'

for r in proj_data:
    ws1.append(r)

stylize(ws1, widths={1: 10, 2: 18, 3: 14, 4: 12, 5: 14, 6: 14, 7: 12, 8: 12, 9: 12,
                     10: 12, 11: 14, 12: 10, 13: 10, 14: 32})
print(f'投标项目总台账: {len(proj_data)} 行')


# ==============================================================
# Sheet 2: 制度适用规则汇编
# 设计：最显眼的置顶规则红字标"战略客户全部适用v2"（假），
# 真正生效的规则藏在下方（只有 18 家备案客户适用，且须管委会书面审批）
# ==============================================================
ws2 = wb.create_sheet('制度适用规则汇编')
hdrs2 = ['制度编号', '制度名称', '生效日期', '适用范围', '核心条款', '当前状态', '备注']
ws2.append(hdrs2)

rules = [
    # —— 置顶：最显眼的旧制度（已废止但未下架，红字！显著性反转 ①）——
    ['PRC-BID-v1.1', '投标报价管理制度（已废止）', '2024-06-01', '全部项目',
     '战略客户统一适用v2.0折扣率（最低15%），普通客户适用v1.0（最低5%），无需单独备案',
     '已废止（2025-01-01起）', '⚠️ 此制度已废止，仅作历史参考——最显眼但无效'],
    # —— 现行制度（真正生效的规则）——
    ['PRC-BID-v2.0', '投标报价管理制度', '2025-01-01', '2025-01-01及以后签约的项目',
     '战略客户折扣适用条件：①客户已在《战略客户备案清单》内；②折扣≤25%且≥15%；③须经管委会书面审批通过，三个条件同时满足方可适用v2.0',
     '现行有效', '三个条件缺一不可，仅"是战略客户"不够'],
    ['PRC-BID-v2.0', '投标报价管理制度', '2025-01-01', '2025-01-01及以后签约的项目',
     '普通客户报价须按成本加成≥30%，不得低于；低于30%的须经合伙人委员会书面审批',
     '现行有效', '普通客户底价 30%'],
    ['PRC-BID-v1.0', '投标报价管理制度', '2023-04-01', '2024-12-31及以前签约的项目',
     '全部客户按v1.0执行：成本加成≥20%，战略客户折扣可放宽至10%，无需单独备案',
     '现行有效（仅适用旧项目）', '签约年份 2024 年及以前'],
    ['PRC-APP-v1.3', '报价审批生效要件规定', '2025-06-15', '全部项目',
     '报价审批必须同时满足：①ERP系统审批状态显示"已通过"；②审批人签字齐全并上传系统；③审批时间在报价有效期内（报价提交后30天）。三者缺一审批无效',
     '现行有效', '—— 关键规则 ——'],
    ['PRC-REL-v1.0', '关联方与陪标识别规则', '2024-02-01', '全部项目',
     '下列情形视为关联方：①投标保证金从同一银行账户转出；②投标联系人IP地址相同；③法定代表人或主要负责人存在近亲属关系；④投标文件结构高度相似。发现任一即视为陪标',
     '现行有效', '—— 关键规则 ——'],
    ['PRC-PEN-v1.0', '违规处置分级', '2023-01-01', '全部项目',
     '违规分为：①整改（补流程或退款）；②解除合同（陪标/造假）；③追究刑事责任（挪用保证金）；④报告行业协会（陪标严重）',
     '现行有效', '—— 处置分级 ——'],
]
for r in rules:
    ws2.append(r)

# 把第 2 行（旧制度）整行标红，让它"最显眼"
for cell in ws2[2]:
    cell.font = Font(name='宋体', size=9, color='CC0000', bold=True)
    cell.fill = PatternFill('solid', fgColor='FFF2CC')

stylize(ws2, widths={1: 18, 2: 28, 3: 14, 4: 22, 5: 60, 6: 16, 7: 36})
# 覆盖第2行的样式（stylize 会把所有行刷成同一样式）
for cell in ws2[2]:
    cell.font = Font(name='宋体', size=9, color='CC0000', bold=True)
    cell.fill = PatternFill('solid', fgColor='FFF2CC')
print(f'制度适用规则汇编: {len(rules)} 行')


# ==============================================================
# Sheet 3: 合同签署与备案审批表（180 行，关联项目+备案+审批）
# 设计：部分低价中标项目无备案记录、部分备案记录审批时间在报价有效期外
# ==============================================================
ws3 = wb.create_sheet('合同签署与备案审批表')
hdrs3 = ['备案编号', '项目编号', '客户名称', '备案类型', '备案申请日', '审批完成日',
         '审批人', '折扣率(%)', '是否管委会审批', 'ERP系统状态', '备注']
ws3.append(hdrs3)

# 为所有 120 个项目生成备案记录（部分低价中标项目缺备案）
bid_price_map = {r[0]: r[8] for r in proj_data if r[13] != '重复行（数据来自系统备份）'}    # 只取不重复的
book_price_map = {r[0]: r[7] for r in proj_data if r[13] != '重复行（数据来自系统备份）'}
# 去掉重复行，以项目号唯一
unique_proj = {}
for r in proj_data:
    if r[0] not in unique_proj or r[13] != '重复行（数据来自系统备份）':
        unique_proj[r[0]] = r

bid_price_map = {k: v[8] for k, v in unique_proj.items()}
book_price_map = {k: v[7] for k, v in unique_proj.items()}
cust_map = {k: v[1] for k, v in unique_proj.items()}
sign_year_map = {k: int(v[4][:4]) for k, v in unique_proj.items()}
margin_type_map = {k: v[9] for k, v in unique_proj.items()}

rec_count = 0
reg_no = 1
# 真实备案战略客户集合（从题目生成时同步）
STRAT_SET = set(c for c, t in CUSTOMERS if t == '战略已备案')

for code, row in unique_proj.items():
    cust = cust_map[code]
    sign_year = sign_year_map[code]
    book_p = book_price_map[code]
    bid_p = bid_price_map[code]
    marg = margin_type_map[code]
    
    # 正常项目：有备案
    if marg == '标准':
        apply_day = date(2025, 1, 15) + timedelta(days=random.randint(0, 600))
        approve_day = apply_day + timedelta(days=random.randint(3, 15))   # 3-15 天内完成审批
        ws3.append([f'BG{reg_no:04d}', code, cust, '标准报价备案',
                    str(apply_day), str(approve_day),
                    random.choice(['李涵', '张帆', '王梓']),
                    '0.0', '否', '已通过', '正常标准报价备案'])
        reg_no += 1
        rec_count += 1
        continue
    
    # 战略折扣项目（需备案 v2 三个条件）
    if marg == '战略折扣':
        is_strat_real = cust in STRAT_SET
        discount_pct = f'{r4((book_p - bid_p) / book_p * 100):.1f}'
        
        # 分支 A：真实备案战略客户 + 2025+ 签约 → 随机决定是否满足 v2.0 三条件
        if is_strat_real and sign_year >= 2025:
            roll = random.random()
            if roll < 0.55:   # 55% 概率满足 v2.0 三条件 → 合规对照组
                apply_day = date(2025, 3, 1) + timedelta(days=random.randint(100, 400))
                approve_day = apply_day + timedelta(days=random.randint(3, 20))   # 3-20 天
                ws3.append([f'BG{reg_no:04d}', code, cust, '战略折扣备案',
                            str(apply_day), str(approve_day),
                            random.choice(['PA', 'PB']),
                            discount_pct,
                            '是', '已通过', f'合规：{approve_day - apply_day}天内审批，管委会已签字'])
                reg_no += 1
                rec_count += 1
                continue
            elif roll < 0.75:  # 20% 概率缺 1 条件：管委会未签字 → v1.0 降级
                apply_day = date(2025, 3, 1) + timedelta(days=random.randint(100, 400))
                approve_day = apply_day + timedelta(days=random.randint(3, 20))
                ws3.append([f'BG{reg_no:04d}', code, cust, '战略折扣备案',
                            str(apply_day), str(approve_day),
                            random.choice(['李涵', '张帆']),
                            discount_pct,
                            '否', '已通过', f'降级原因：无管委会审批签字（仅有部门负责人签字）'])
                reg_no += 1
                rec_count += 1
                continue
            else:             # 25% 概率无备案 → 违规
                ws3.append([f'BG{reg_no:04d}', code, cust, '战略折扣备案',
                            '', '', '', discount_pct,
                            '否', '未提交', '违规：战略折扣无备案、无审批'])
                reg_no += 1
                rec_count += 1
                continue
        
        # 分支 B：非真实备案客户 or 2024 年签约 → 一律 v1.0 降级/违规
        else:
            apply_day = date(2025, 3, 1) + timedelta(days=random.randint(100, 400))
            approve_day = apply_day + timedelta(days=random.randint(35, 60))
            ws3.append([f'BG{reg_no:04d}', code, cust, '战略折扣备案',
                        str(apply_day), str(approve_day),
                        random.choice(['李涵', '张帆']),
                        discount_pct,
                        '否', '已通过',
                        f"降级原因：{'非真实备案战略客户' if not is_strat_real else '签约年份 2024 年适用 v1.0'} + 审批" +
                        f"{approve_day - apply_day}天超 30 天有效期"])
            reg_no += 1
            rec_count += 1
            continue
    
    # 低价中标（违规！）：全部无备案
    if marg == '低价中标':
        ws3.append([f'BG{reg_no:04d}', code, cust, '低价中标备案',
                    '', '', '', f'{r4((book_p - bid_p) / book_p * 100):.1f}',
                    '否', '未提交', '—— 低价中标，无备案，无审批记录，属严重违规 ——'])
        reg_no += 1
        rec_count += 1

stylize(ws3, widths={1: 12, 2: 10, 3: 16, 4: 16, 5: 14, 6: 14,
                     7: 10, 8: 10, 9: 16, 10: 12, 11: 40})
print(f'合同签署与备案审批表: {rec_count} 行')


# ==============================================================
# Sheet 4: 报价审批流记录表（580 行，每步审批留痕，部分被驳回的低价申请）
# 设计：P078 陪标项目的审批流藏有"陪标保证金同一账户转出"的证据
# ==============================================================
ws4 = wb.create_sheet('报价审批流记录表')
hdrs4 = ['审批记录号', '项目编号', '客户名称', '审批阶段', '审批人', '审批结果',
         '审批日期', '申请报价(万元)', '批准报价(万元)', '驳回原因', '备注']
ws4.append(hdrs4)

ap_rec = 1
for code, row in unique_proj.items():
    cust = cust_map[code]
    book_p = book_price_map[code]
    bid_p = bid_price_map[code]
    marg = margin_type_map[code]
    
    # 每项目生成 3-5 条审批流
    n_steps = random.randint(3, 5)
    prev_p = book_p
    for step in range(n_steps):
        if step == n_steps - 1:
            # 最后一步：合伙人审批
            if marg == '低价中标':
                ws4.append([f'AP{ap_rec:05d}', code, cust, f'第{step+1}步-合伙人终审',
                            random.choice(partners), '驳回',
                            str(date(2025, 6, 1) + timedelta(days=random.randint(0, 500))),
                            prev_p, '', '低于普通客户成本加成30%底价，无管委会审批',
                            '' if random.random() > 0.05 else '【关键：此项目低价中标，合伙人审批全部被驳回但报价实际执行——审批流程形同虚设】'])
            elif marg == '战略折扣' and random.random() < 0.15:
                ws4.append([f'AP{ap_rec:05d}', code, cust, f'第{step+1}步-合伙人终审',
                            random.choice(partners), '驳回',
                            str(date(2025, 5, 1) + timedelta(days=random.randint(0, 500))),
                            prev_p, '', '折扣率超25%上限或无管委会书面审批',
                            ''])
            else:
                ws4.append([f'AP{ap_rec:05d}', code, cust, f'第{step+1}步-合伙人终审',
                            random.choice(partners), '通过',
                            str(date(2025, 5, 1) + timedelta(days=random.randint(0, 500))),
                            prev_p, bid_p if step == n_steps - 1 else prev_p,
                            '', ''])
        else:
            ws4.append([f'AP{ap_rec:05d}', code, cust, f'第{step+1}步-{random.choice(["项目经理", "财务经理", "合规审核"])}',
                        random.choice(consultants), random.choice(['通过', '通过', '通过', '驳回']),
                        str(date(2025, 3, 1) + timedelta(days=random.randint(0, 500))),
                        prev_p, prev_p, '', ''])
        prev_p = bid_p
        ap_rec += 1

# P078 陪标项目的特殊证据：在审批流备注里藏 JSON
# 找出 P078 的最后一条审批记录并改写
for i, r in enumerate(ws4.iter_rows(min_row=2, values_only=True)):
    if r[1] == 'P078' and r[3].startswith('第') and '终审' in r[3]:
        row_idx = i + 2
        ws4.cell(row=row_idx, column=11).value = ('{"陪标线索": true, "保证金同一账户转出": true, '
                                                   '"投标IP相同": true, "关联方": "华盛贸易"}')
        break

stylize(ws4, widths={1: 12, 2: 10, 3: 16, 4: 18, 5: 10, 6: 10, 7: 14,
                     8: 12, 9: 12, 10: 28, 11: 40})
print(f'报价审批流记录表: {ap_rec - 1} 行')


# ==============================================================
# Sheet 5: 战略客户备案清单（45 行，置顶写"全部适用v2"但实际只 18 家备案）
# 显著性反转 ①：最显眼的置顶红字误导 AI
# ==============================================================
ws5 = wb.create_sheet('战略客户备案清单')
hdrs5 = ['备案编号', '客户名称', '备案类型', '备案生效日', '折扣权限',
         '是否经管委会审批', '管委会审批文号', '备注']
ws5.append(hdrs5)

# 置顶误导性"通知"——用一行加粗大字写"以下客户全部适用 v2.0"
ws5.append(['⚠️【特别通知】', '以下 45 家客户均为战略客户，自 2025-01-01 起全部适用 PRC-BID-v2.0 折扣规则',
            '', '', '', '', '', '此行为置顶醒目提示（但不构成规则本身）'])

# 只有前 18 家（即 STRAT 集合里的）真正备案了 v2.0 折扣
for i, (cust, cust_t) in enumerate(CUSTOMERS):
    bg = f'SC{i+1:03d}'
    if cust in STRAT and cust_t == '战略已备案':
        ws5.append([bg, cust, 'v2.0 折扣备案（真实生效）',
                    str(date(2024, 12, 28) + timedelta(days=random.randint(0, 30))),
                    '15%-25%', '是', f'GB-{(2025)*100 + i:04d}', '—— 有效备案 ——'])
    elif cust == '华盛贸易':
        ws5.append([bg, cust, 'v1.0 普通客户（无 v2.0 备案）',
                    str(date(2024, 3, 15)),
                    '无', '否', '', '普通客户，未申请 v2.0 备案'])
    else:
        ws5.append([bg, cust, '普通客户（无 v2.0 备案）',
                    str(date(2023, 6, 10)),
                    '无', '否', '', '未申请战略客户备案'])

stylize(ws5, widths={1: 12, 2: 16, 3: 20, 4: 14, 5: 12, 6: 16, 7: 14, 8: 36})
# 覆盖置顶行样式
for cell in ws5[2]:
    cell.font = Font(name='宋体', size=10, color='CC0000', bold=True)
    cell.fill = PatternFill('solid', fgColor='FFE699')
print(f'战略客户备案清单: {len(CUSTOMERS) + 1} 行（含置顶误导）')


# ==============================================================
# Sheet 6: 月度行政通报（12 行正文 + 隐藏的页脚规则 —— 显著性反转 ③）
# 设计：关键的"报价审批生效要件"藏在 sheet 级 comment 里（AI 默认跳过）
# ==============================================================
ws6 = wb.create_sheet('月度行政通报')
hdrs6 = ['通报编号', '月份', '主题', '正文摘要', '附件位置', '发布人', '备注']
ws6.append(hdrs6)

bulletins = [
    ['TB-2026-01', '2026-01', '年度投标工作总结', '全年投标 124 个，中标率 68.7%，同比+2.3%', '', '行政管理部', ''],
    ['TB-2026-02', '2026-02', '2026年报价审批流程调整通知', '即日起报价审批须由 ERP 系统线上完成，纸质审批单不再作为有效凭证', '', '合规部', ''],
    ['TB-2026-03', '2026-03', '第一季度投标质量分析', '低价中标项目 12 个，其中 4 个后续出现客户投诉', '', '质量部', ''],
    ['TB-2026-04', '2026-04', '关联方交易自查要求', '全员自查 2025 年以来参与的投标项目，填报关联方信息表', '', '内控部', ''],
    ['TB-2026-05', '2026-05', '战略客户折扣使用情况通报', '已备案 18 家战略客户，2025 年共 23 个项目使用折扣', '', '客户管理部', ''],
    ['TB-2026-06', '2026-06', '——（以下为空，凑数据用）', '', '', '', ''],
]
for r in bulletins:
    ws6.append(r)

# 写入 sheet 级 comment（openpyxl 支持，xlsx 有这个字段）
try:
    from openpyxl.comments import Comment
    sheet_comment = Comment(
        '⚠️【关键规则，藏于页脚，AI 读 xlsx 内容时默认跳过此注释】\n'
        'PRC-APP-v1.3《报价审批生效要件规定》的补充说明（2025-06-15 生效）：\n'
        '1. 报价审批必须同时满足：①ERP系统状态=已通过；②审批人签字齐全并上传系统；③审批时间在报价提交后30天内。三者缺一审批无效。\n'
        '2. 仅有纸质审批单而无 ERP 线上审批记录的，视为未完成审批，对应的折扣/低价报价不得执行。\n'
        '3. 低价中标项目无论何种理由，均须经管委会书面审批并在 ERP 系统上完成全部流程。',
        '合规部-李涵'
    )
    ws6.comment = sheet_comment
    print('页脚关键规则注释已写入')
except Exception as e:
    print(f'写入 sheet comment 失败：{e}——改为第 7 行备注列')
    ws6.append(['', '', '', '', '', '',
                '【关键规则·藏于此处·AI 易跳过】PRC-APP-v1.3 补充：报价审批生效要件须 ERP 已通过+签字上传+30天内，三者缺一无效；仅有纸质单而无 ERP 线上审批视为未完成；低价中标须管委会书面审批并走完 ERP 全流程'])

stylize(ws6, widths={1: 14, 2: 10, 3: 22, 4: 56, 5: 14, 6: 10, 7: 60})
print(f'月度行政通报: {len(bulletins)} 行')


# ==============================================================
# Sheet 7: 关联方与陪标线索表（30 行，P078 陪标证据藏备注 JSON）
# 显著性反转 ②：备注列是 JSON，AI 大概率跳过非结构化内容
# ==============================================================
ws7 = wb.create_sheet('关联方与陪标线索表')
hdrs7 = ['线索编号', '涉及项目', '关联方A', '关联方B', '关系类型', '发现渠道', '置信度', '详细证据(JSON)']
ws7.append(hdrs7)

rel_rows = []
for i in range(1, 26):
    code = f'P{random.randint(1, 120):03d}'
    rel_rows.append([f'RL{i:03d}', code, random.choice(CUSTOMERS)[0], random.choice(CUSTOMERS)[0],
                     random.choice(['法定代表人近亲属', '投标保证金同账户', '投标联系人IP相同', '投标文件高度相似', '无关联']),
                     random.choice(['内部自查', '外部举报', '系统监控']),
                     random.choice(['低', '低', '中', '中', '高']),
                     '{"details": "......"}'])
# P078 的陪标证据（置信度高，JSON 完整）
rel_rows.append(['RL999', 'P078', '华盛贸易', '睿衡企业管理咨询有限公司（投标关联方）',
                 '投标保证金同账户转出+法定代表人近亲属', '外部举报',
                 '高',
                 '{"保证金转出银行账户": "工行 6222***8801", "投标联系人IP": "192.168.10.*", "法定代表人关系": "父子", "投标文件相似度": "92%", "举报时间": "2026-11-28"}'])

for r in rel_rows:
    ws7.append(r)

stylize(ws7, widths={1: 12, 2: 10, 3: 16, 4: 20, 5: 22, 6: 14, 7: 8, 8: 60})
print(f'关联方与陪标线索表: {len(rel_rows)} 行')


# ==============================================================
# Sheet 8: 管理层投标报价指令（12 条错误 + 1 条正确，混在一起）
# 设计：正确指令藏在中间（第 7 条），周围全是错误的，AI 大概率跳过
# ==============================================================
ws8 = wb.create_sheet('管理层投标报价指令')
hdrs8 = ['指令编号', '发布时间', '发布人', '指令内容', '指令级别', '是否合规（审计部标注）', '备注']
ws8.append(hdrs8)

mgmt_cmds = [
    ['CMD-01', '2026-08-12', '销售总监-陈总', '全部 2025 年以后签约的项目，无论是非战略客户，统一按 v2.0 执行最低 15% 折扣，不必单独备案', '口头指令', '错误', ''],
    ['CMD-02', '2026-09-03', '运营副总-李总', 'P078 项目陪标一事压下去，对审计说是"正常业务操作"', '微信文字', '错误', ''],
    ['CMD-03', '2026-09-15', '合伙人-PA', '审批流超 30 天有效期但业务上没问题的，后续补纸质单就行，ERP 状态可以改', '邮件', '错误', ''],
    ['CMD-04', '2026-10-02', '销售总监-陈总', '低价中标项目后续报价可以直接跳过合规审核，我来拍板', '口头指令', '错误', ''],
    ['CMD-05', '2026-10-18', '运营副总-李总', '关联方自查让全员写"无"就行，不用真查', '邮件', '错误', ''],
    ['CMD-06', '2026-11-01', '合伙人-PC', '战略客户折扣备案编号 BG0032 之后的，可以不填管委会审批文号', '口头指令', '错误', ''],
    # —— 正确指令，藏中间 ——
    ['CMD-07', '2026-11-10', '合规部负责人-王总', 'P078 陪标项目必须中止，立即解除与华盛贸易的合作关系并上报管委会；同时全面自查 2025 年以来所有低价中标项目，发现违规立即整改', '正式邮件', '正确', '—— 正确指令，被错误指令包围 ——'],
    ['CMD-08', '2026-11-15', '销售总监-陈总', 'P078 陪标上报会影响拿新项目，暂缓执行', '口头指令', '错误', ''],
    ['CMD-09', '2026-11-20', '运营副总-李总', '备案表里把低价中标项目补成"战略折扣"就行，别留痕迹', '微信文字', '错误', ''],
    ['CMD-10', '2026-11-25', '合伙人-PD', '客户说低价就行，没签协议就算了，我们按 v1.0 旧制度走', '邮件', '错误', ''],
    ['CMD-11', '2026-12-01', '销售总监-陈总', '关联方识别只查投标 IP，保证金同账户不算（客户共用会计正常）', '口头指令', '错误', ''],
    ['CMD-12', '2026-12-10', '运营副总-李总', '陪标不违法，就是行业惯例，不用上纲上线', '口头指令', '错误', ''],
]
for r in mgmt_cmds:
    ws8.append(r)

stylize(ws8, widths={1: 10, 2: 14, 3: 16, 4: 64, 5: 12, 6: 22, 7: 40})
print(f'管理层投标报价指令: {len(mgmt_cmds)} 行')

wb.save(OUT)
total_rows = sum(ws.max_row - 1 for ws in wb.worksheets)
print(f'\n===== 题目已生成：{OUT} =====')
print(f'8 张表，合计 {total_rows} 行数据')
for ws in wb.worksheets:
    print(f'  {ws.title}: {ws.max_row - 1} 行')
