#!/usr/bin/env python3
"""民办学校年检 v3 - 大幅提高难度
策略：数据量1200+项 + 计算型标记(日期推算/百分比阈值) + 跨表查询 + 大量合规保留陷阱项
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import date, timedelta
from dateutil.relativedelta import relativedelta

THIN = Side(style="thin", color="CCCCCC")
BORDER = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
HDR_FILL = PatternFill("solid", fgColor="2B5B84")
HDR_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=10)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

REF_DATE = date(2025, 9, 1)  # 申请延续基准日

def make_sheet(wb, name, headers, data, col_widths=None):
    ws = wb.create_sheet(name)
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=i, value=h)
        c.font = HDR_FONT; c.fill = HDR_FILL; c.alignment = CENTER; c.border = BORDER
    for r, row in enumerate(data, 2):
        for i, v in enumerate(row, 1):
            c = ws.cell(row=r, column=i, value=v)
            c.font = Font(name="Arial", size=9); c.alignment = WRAP; c.border = BORDER
    if col_widths:
        for i, w in enumerate(col_widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
    return ws

wb = openpyxl.Workbook()
wb.remove(wb.active)

# ========== 学校基本信息 ==========
make_sheet(wb, '学校基本信息', ['序号','学校名称','办学许可证号','许可证有效期','办学层次','办学类型','成立日期','举办者','法定代表人','校长','注册资金(万元)','实缴资金(万元)','校园面积(㎡)','校舍面积(㎡)','在校生数','教职工数','备注'], [
    [1,'某市民办文华中学','教民1310XXXX0001号','2020-09-01至2025-08-31','初中、普通高中','完全中学','2005-06-15','某教育投资有限公司','王某','李某','3000','3000','35000','28000','1800','250','现有许可证2025年8月31日到期，须申请延续'],
])

# ========== 办学许可证信息 ==========
make_sheet(wb, '办学许可证信息', ['序号','许可证号','发证机关','发证日期','有效期至','办学内容','办学地址','年检结果','备注'], [
    [1,'教民1310XXXX0001号','某市教育局','2020-09-01','2025-08-31','初中教育','某市某区教育路100号','2023年度合格',''],
    [2,'教民1310XXXX0002号','某市教育局','2020-09-01','2025-08-31','普通高中教育','某市某区教育路100号','2023年度合格',''],
    [3,'教民1310XXXX0003号','某市教育局','2020-09-01','2025-08-31','国际部课程','某市某区教育路100号','2023年度基本合格','国际部须补充审批材料'],
])

# ========== 办学层次与学科要求（跨表判断教师学历） ==========
# 高中教师须本科及以上，初中教师专科即可
make_sheet(wb, '办学层次与学科要求', ['序号','办学层次','对应年级','教师学历最低要求','任教学科范围'], [
    [1,'初中','初一、初二、初三','专科及以上','语文、数学、英语、物理、化学、生物、政治、历史、地理'],
    [2,'普通高中','高一、高二、高三','本科及以上','语文、数学、英语、物理、化学、生物、政治、历史、地理'],
])

# ========== 教师花名册（250人） ==========
subjects = ['语文','数学','英语','物理','化学','生物','政治','历史','地理']
jr_subjects = ['语文','数学','英语','物理','化学','生物','政治','历史','地理']
teachers = []
# 无证40人 (T001-T040)
for i in range(1, 41):
    teachers.append([i, f'T{i:03d}', f'教师{i}', '待办理', '-', subjects[i%9], subjects[i%9], '本科', i%15+1, '在岗', ''])
# 过期40人 (T041-T080)，有效期至2024年某日
for i in range(41, 81):
    d = date(2024, (i%12)+1, (i%27)+1)
    teachers.append([i, f'T{i:03d}', f'教师{i}', f'JSZ2024{i:03d}', d.strftime('%Y-%m-%d'), subjects[i%9], subjects[i%9], '本科', i%15+1, '在岗', ''])
# 学历不达标30人 (T081-T110)：专科 + 高中学科
for i in range(81, 111):
    teachers.append([i, f'T{i:03d}', f'教师{i}', f'JSZ2025{i:03d}', '2027-06-30', subjects[i%9], subjects[i%9], '专科', i%15+1, '在岗', ''])
# 专业不对口30人 (T111-T140)：任教学科 != 资格证学科
for i in range(111, 141):
    s1 = subjects[i%9]; s2 = subjects[(i+3)%9]
    if s1 == s2: s2 = subjects[(i+5)%9]
    teachers.append([i, f'T{i:03d}', f'教师{i}', f'JSZ2025{i:03d}', '2027-06-30', s1, s2, '本科', i%15+1, '在岗', ''])
# 合规保留-特级教师20人 (T141-T160)
for i in range(141, 161):
    teachers.append([i, f'T{i:03d}', f'教师{i}', f'JSZ2025{i:03d}', '2027-06-30', subjects[i%9], subjects[i%9], '本科', 20+i%5, '在岗', '特级教师，教学经验丰富，证书齐全'])
# 合规保留-兼职教师20人 (T161-T180)
for i in range(161, 181):
    teachers.append([i, f'T{i:03d}', f'教师{i}', f'JSZ2025{i:03d}', '2027-06-30', subjects[i%9], subjects[i%9], '硕士', 10+i%5, '兼职', '外聘兼职教师，有正式聘用协议'])
# 正常70人 (T181-T250)
for i in range(181, 251):
    teachers.append([i, f'T{i:03d}', f'教师{i}', f'JSZ2025{i:03d}', '2027-06-30', subjects[i%9], subjects[i%9], '本科', i%15+1, '在岗', ''])

make_sheet(wb, '教师花名册', ['序号','教师编号','姓名','教师资格证号','资格证有效期至','任教学科','资格证学科','学历','教龄','在岗状态','备注'],
    teachers, [6,10,8,16,16,10,12,8,6,10,28])

# ========== 收费标准清单（40项，实际收取） ==========
fee_items = [
    ('学费(初中)', 10000), ('学费(高中)', 12000), ('住宿费', 1800), ('伙食费', 3500),
    ('代收教材费', 800), ('校服费', 500), ('保险费', 150), ('课后服务费', 700),
    ('作业本费', 60), ('考试费', 100), ('军训费', 250), ('体检费', 120),
    ('社会实践费', 180), ('电教费', 200), ('图书馆使用费', 100), ('体育馆使用费', 120),
    ('实验室使用费', 180), ('计算机使用费', 200), ('空调使用费', 120), ('饮水费', 80),
    ('取暖费', 100), ('班费', 50), ('社团活动费', 150), ('竞赛辅导费', 200),
    ('夏令营费', 500), ('冬令营费', 400), ('国际交流费', 3000), ('校车费', 800),
    ('游泳课费', 200), ('书法课费', 150), ('舞蹈课费', 200), ('乐器课费', 500),
    ('美术课费', 200), ('陶艺课费', 150), ('编程课费', 300), ('机器人课费', 400),
    ('英语口语费', 300), ('外教课费', 800), ('心理咨询费', 100), ('校史馆使用费', 50),
]
fee_data = []
for i, (name, actual) in enumerate(fee_items, 1):
    fee_data.append([i, name, actual, '按学年收取'])
make_sheet(wb, '收费标准清单', ['序号','收费项目','实际收取(元/年)','收取说明'],
    fee_data, [6,18,16,20])

# ========== 收费批复标准（查询表，40项） ==========
# 与收费标准清单一一对应，部分批复标准低于实际收取（违规）
fee_approved = [
    ('学费(初中)', 8000), ('学费(高中)', 10000), ('住宿费', 1500), ('伙食费', 3000),
    ('代收教材费', 500), ('校服费', 300), ('保险费', 100), ('课后服务费', 600),
    ('作业本费', 50), ('考试费', 80), ('军训费', 200), ('体检费', 100),
    ('社会实践费', 150), ('电教费', 120), ('图书馆使用费', 80), ('体育馆使用费', 100),
    ('实验室使用费', 150), ('计算机使用费', 120), ('空调使用费', 100), ('饮水费', 50),
    ('取暖费', 100), ('班费', 50), ('社团活动费', 150), ('竞赛辅导费', 200),
    ('夏令营费', 500), ('冬令营费', 400), ('国际交流费', 3000), ('校车费', 800),
    ('游泳课费', 200), ('书法课费', 150), ('舞蹈课费', 200), ('乐器课费', 500),
    ('美术课费', 200), ('陶艺课费', 150), ('编程课费', 300), ('机器人课费', 400),
    ('英语口语费', 300), ('外教课费', 800), ('心理咨询费', 100), ('校史馆使用费', 50),
]
fee_appr_data = []
for i, (name, approved) in enumerate(fee_approved, 1):
    fee_appr_data.append([i, name, approved, '某市发改委批复'])
make_sheet(wb, '收费批复标准', ['序号','收费项目','批复标准(元/年)','批复来源'],
    fee_appr_data, [6,18,16,20])

# ========== 财务报表（50项） ==========
fin_items = [
    ('学费收入', '1500万', '1420万', '收费人数1800人', '低于预期80万', '异常'),
    ('住宿费收入', '270万', '270万', '住宿人数1500人×1800元', '符合预算', '正常'),
    ('伙食费收入', '525万', '525万', '就餐人数1500人×3500元', '符合预算', '正常'),
    ('人员经费', '2200万', '2300万', '教职工250人平均工资', '超支100万', '异常'),
    ('公用经费', '450万', '520万', '日常办公支出', '超支70万', '异常'),
    ('设备购置费', '300万', '340万', '教学设备采购', '超支40万', '异常'),
    ('政府补助收入', '600万', '600万', '财政拨款', '符合预算', '正常'),
    ('捐赠收入', '120万', '120万', '社会捐赠', '符合预算', '正常'),
    ('资产负债率', '45%', '52%', '负债总额/资产总额', '上升7个百分点', '异常'),
    ('办学积累', '300万', '220万', '年度结余', '下降80万', '异常'),
]
fin_data = []
for i, (name, budget, actual, calc, diff, conclusion) in enumerate(fin_items, 1):
    fin_data.append([i, name, budget, actual, calc, diff, conclusion])
for i in range(11, 51):
    fin_data.append([i, f'财务项目{i}', f'{i*12}万', f'{i*12}万', '按预算执行', '符合预算', '正常'])
make_sheet(wb, '财务报表', ['序号','财务项目','预算金额','实际金额','计算说明','差异分析','结论'],
    fin_data, [6,18,12,12,24,18,10])

# ========== 安全检查记录（180项） ==========
safety = []
check_types = ['消防安全','设施安全','食品安全','校车安全','宿舍安全','校园安全']
levels = ['一般隐患','较大隐患','重大隐患','无隐患']
# 未整改40项
for i in range(1, 41):
    safety.append([i, f'SF{i:03d}', f'2024-{(i%12)+1:02d}-15', check_types[i%6], levels[i%3], '未整改', '2024-12-31', f'{check_types[i%6]}检查发现隐患'])
# 重大隐患30项
for i in range(41, 71):
    safety.append([i, f'SF{i:03d}', f'2024-{(i%12)+1:02d}-20', check_types[i%6], '重大隐患', '整改中', '2025-03-31', f'{check_types[i%6]}检查发现重大安全隐患'])
# 整改超期25项（整改中但整改期限已过基准日），隐患等级不含重大隐患
for i in range(71, 96):
    lvl = levels[i % 2]  # 只用一般隐患或较大隐患
    safety.append([i, f'SF{i:03d}', f'2024-{(i%12)+1:02d}-25', check_types[i%6], lvl, '整改中', '2025-06-30', f'{check_types[i%6]}检查发现隐患，正在整改'])
# 合规保留-已整改25项
for i in range(96, 121):
    safety.append([i, f'SF{i:03d}', f'2024-{(i%12)+1:02d}-10', check_types[i%6], '一般隐患', '已整改', '2024-11-30', '隐患已整改完成，复查合格'])
# 正常60项
for i in range(121, 181):
    safety.append([i, f'SF{i:03d}', f'2024-{(i%12)+1:02d}-28', check_types[i%6], '无隐患', '-', '-', '检查未发现安全隐患'])
make_sheet(wb, '安全检查记录', ['序号','检查编号','检查日期','检查类型','隐患等级','整改状态','整改期限','备注'],
    safety, [6,10,14,12,12,10,14,36])

# ========== 设施设备清单（180项，需日期推算） ==========
facilities = []
fac_names = ['教学楼','实验楼','宿舍楼','食堂','体育馆','图书馆','操场','配电室','消防设施','监控设备']
# 超期未检45项：上次检验+周期 < 2025-09-01
for i in range(1, 46):
    cycle = 6 + (i % 6)  # 6-11个月
    last = date(2024, (i%12)+1, (i%27)+1)
    facilities.append([i, f'FA{i:03d}', f'{fac_names[i%10]}{i}', last.strftime('%Y-%m-%d'), cycle, '2018-05-01', '在用', '定期巡检中'])
# 设备老化35项：购置日期2009年及以前
for i in range(46, 81):
    cycle = 12
    last = date(2025, 3, (i%27)+1)
    buy = date(2009, (i%12)+1, (i%27)+1)
    facilities.append([i, f'FA{i:03d}', f'{fac_names[i%10]}{i}', last.strftime('%Y-%m-%d'), cycle, buy.strftime('%Y-%m-%d'), '在用', '设备使用年限较长'])
# 合规保留-新购25项：2022年后购置
for i in range(81, 106):
    cycle = 12
    last = date(2025, 5, (i%27)+1)
    buy = date(2023, (i%12)+1, (i%27)+1)
    facilities.append([i, f'FA{i:03d}', f'{fac_names[i%10]}{i}', last.strftime('%Y-%m-%d'), cycle, buy.strftime('%Y-%m-%d'), '在用', '近年新购设备'])
# 正常75项：在检验周期内且购置不超过15年
for i in range(106, 181):
    cycle = 12
    last = date(2025, 4, (i%27)+1)
    buy = date(2015, (i%12)+1, (i%27)+1)
    facilities.append([i, f'FA{i:03d}', f'{fac_names[i%10]}{i}', last.strftime('%Y-%m-%d'), cycle, buy.strftime('%Y-%m-%d'), '在用', '设施状态良好'])
make_sheet(wb, '设施设备清单', ['序号','设施编号','设施名称','上次检验日期','检验周期(月)','购置日期','设施状态','备注'],
    facilities, [6,10,18,16,14,14,10,28])

# ========== 招生计划与实际（120项） ==========
enrollment = []
grades = ['初一','初二','初三','高一','高二','高三']
# 超计划30项：实际 > 批复计划×1.05
for i in range(1, 31):
    plan = 200
    actual = int(plan * 1.06) + (i % 5) * 5  # 6%-10%超
    enrollment.append([i, f'RE{i:03d}', '2024', grades[i%6], actual, '已备案', ''])
# 未备案25项
for i in range(31, 56):
    enrollment.append([i, f'RE{i:03d}', '2024', grades[i%6], 150, '未备案', ''])
# 合规保留-已备案20项（实际略超计划但在5%以内，陷阱项）
for i in range(56, 76):
    enrollment.append([i, f'RE{i:03d}', '2024', grades[i%6], 188, '已备案', ''])
# 正常45项（实际等于计划且已备案）
for i in range(76, 121):
    enrollment.append([i, f'RE{i:03d}', '2024', grades[i%6], 200, '已备案', ''])
make_sheet(wb, '招生计划与实际', ['序号','招生编号','年度','年级','实际招生人数','招生简章备案状态','备注'],
    enrollment, [6,10,8,8,14,18,30])

# ========== 招生计划批复（查询表，120项） ==========
enroll_plan = []
for i in range(1, 31):
    enroll_plan.append([i, f'RE{i:03d}', grades[i%6], 200])
for i in range(31, 56):
    enroll_plan.append([i, f'RE{i:03d}', grades[i%6], 150])
for i in range(56, 76):
    enroll_plan.append([i, f'RE{i:03d}', grades[i%6], 180])
for i in range(76, 121):
    enroll_plan.append([i, f'RE{i:03d}', grades[i%6], 200])
make_sheet(wb, '招生计划批复', ['序号','招生编号','年级','批复计划人数'],
    enroll_plan, [6,10,8,14])

# ========== 学籍管理记录（250项） ==========
students = []
grades2 = ['初一','初二','初三','高一','高二','高三']
# 学籍不完整50项
for i in range(1, 51):
    students.append([i, f'ST{i:03d}', f'学生{i}', grades2[i%6], '在籍在读', '学籍档案缺少身份证号和户籍信息'])
# 人籍分离40项
for i in range(51, 91):
    students.append([i, f'ST{i:03d}', f'学生{i}', grades2[i%6], '在籍在读', '学生实际不在校就读，学籍仍保留'])
# 违规转学30项
for i in range(91, 121):
    students.append([i, f'ST{i:03d}', f'学生{i}', grades2[i%6], '在籍在读', '转学手续不完整，未经教育主管部门审批'])
# 合规保留30项
for i in range(121, 151):
    students.append([i, f'ST{i:03d}', f'学生{i}', grades2[i%6], '在籍在读', '学籍信息完整，人籍一致'])
# 正常100项
for i in range(151, 251):
    students.append([i, f'ST{i:03d}', f'学生{i}', grades2[i%6], '在籍在读', ''])
make_sheet(wb, '学籍管理记录', ['序号','学籍号','学生姓名','年级','学籍状态','备注'],
    students, [6,10,12,8,10,40])

# ========== 管理制度清单 ==========
make_sheet(wb, '管理制度清单', ['序号','制度名称','发布日期','状态'], [
    [1,'学校章程','2020-09-01','已有'],
    [2,'财务管理制度','2020-09-01','已有'],
    [3,'安全管理制度','2020-09-01','已有'],
    [4,'招生管理制度','2020-09-01','已有'],
    [5,'教职工管理制度','2020-09-01','已有'],
])

# ========== 管理层要求（10项） ==========
make_sheet(wb, '管理层要求', ['序号','来源','内容','性质'], [
    [1,'管理层要求1','超计划招生的学生直接办理入学，不调整招生计划','错误要求'],
    [2,'管理层要求2','无教师资格证的教师继续任教，不安排考证','错误要求'],
    [3,'管理层要求3','超标准收取的费用不退还，继续按此标准收取','错误要求'],
    [4,'管理层要求4','安全隐患未整改的继续教学，不暂停相关活动','错误要求'],
    [5,'管理层要求5','学籍人籍分离的不处理，保持现状','错误要求'],
    [6,'管理层要求6','财务报表不报送教育主管部门','错误要求'],
    [7,'管理层要求7','招生简章不报备直接发布','错误要求'],
    [8,'管理层要求8','设施超期未检继续使用，不安排检验','错误要求'],
    [9,'管理层要求9','违规转学不纠正，不补办手续','错误要求'],
    [10,'管理层要求10','办学许可证到期不申请延续，继续办学','错误要求'],
])

wb.save('/workspace/mb_v3_tmp.xlsx')
print(f"教师:{len(teachers)} 安全:{len(safety)} 设施:{len(facilities)} 招生:{len(enrollment)} 学籍:{len(students)} 收费:{len(fee_data)} 财务:{len(fin_data)} 批复:{len(fee_appr_data)} 招生批复:{len(enroll_plan)}")
print(f"总数据项: {len(teachers)+len(safety)+len(facilities)+len(enrollment)+len(students)+len(fee_data)+len(fin_data)}")
