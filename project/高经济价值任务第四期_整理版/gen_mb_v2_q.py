#!/usr/bin/env python3
"""民办学校年检 v2 - 大幅提高难度+延长运行时间
策略：数据量2.3倍+隐晦标记（多字段交叉判断）+反直觉项+30轮
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

THIN = Side(style="thin", color="CCCCCC")
BORDER = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
HDR_FILL = PatternFill("solid", fgColor="2B5B84")
HDR_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=10)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

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
    [1,'某市民办文华中学','教民1310XXXX0001号','2020-09-01至2025-08-31','初中','普通高中','2005-06-15','某教育投资有限公司','王某','李某','3000','3000','35000','28000','1200','95','现有许可证即将到期，须申请延续'],
])

# ========== 办学许可证信息 ==========
make_sheet(wb, '办学许可证信息', ['序号','许可证号','发证机关','发证日期','有效期至','办学内容','办学地址','年检结果','备注'], [
    [1,'教民1310XXXX0001号','某市教育局','2020-09-01','2025-08-31','初中教育','某市某区教育路100号','2023年度合格',''],
    [2,'教民1310XXXX0002号','某市教育局','2020-09-01','2025-08-31','普通高中教育','某市某区教育路100号','2023年度合格',''],
    [3,'教民1310XXXX0003号','某市教育局','2020-09-01','2025-08-31','国际部课程','某市某区教育路100号','2023年度基本合格','国际部须补充审批材料'],
])

# ========== 教师花名册（150人，隐晦标记） ==========
# columns: 序号, 教师编号, 姓名, 教师资格证号, 资格证有效期至, 任教学科, 资格证学科, 学历, 教龄, 在岗状态, 备注
teachers = []
subjects = ['语文','数学','英语','物理','化学','生物','政治','历史','地理']
# 无证20人
for i in range(1, 21):
    teachers.append([i, f'T{i:03d}', f'教师{i}', '无', '-', subjects[i%9], '-', '本科', i%10+1, '在岗', ''])
# 过期20人（有效期至2024年，已过期）
for i in range(21, 41):
    teachers.append([i, f'T{i:03d}', f'教师{i}', f'JSZ2024{i:03d}', '2024-06-30', subjects[i%9], subjects[i%9], '本科', i%10+1, '在岗', ''])
# 学历不达标15人
for i in range(41, 56):
    teachers.append([i, f'T{i:03d}', f'教师{i}', f'JSZ2025{i:03d}', '2027-06-30', subjects[i%9], subjects[i%9], '专科', i%10+1, '在岗', ''])
# 专业不对口15人
for i in range(56, 71):
    teachers.append([i, f'T{i:03d}', f'教师{i}', f'JSZ2025{i:03d}', '2027-06-30', '物理', '历史', '本科', i%10+1, '在岗', ''])
# 合规保留-特级教师8人
for i in range(71, 79):
    teachers.append([i, f'T{i:03d}', f'教师{i}', f'JSZ2025{i:03d}', '2027-06-30', subjects[i%9], subjects[i%9], '本科', 20, '在岗', '特级教师，教学经验丰富'])
# 合规保留-兼职教师8人
for i in range(79, 87):
    teachers.append([i, f'T{i:03d}', f'教师{i}', f'JSZ2025{i:03d}', '2027-06-30', subjects[i%9], subjects[i%9], '硕士', 10, '兼职', '外聘兼职教师，有聘用协议'])
# 正常教师63人
for i in range(87, 151):
    teachers.append([i, f'T{i:03d}', f'教师{i}', f'JSZ2025{i:03d}', '2027-06-30', subjects[i%9], subjects[i%9], '本科', i%10+1, '在岗', ''])

make_sheet(wb, '教师花名册', ['序号','教师编号','姓名','教师资格证号','资格证有效期至','任教学科','资格证学科','学历','教龄','在岗状态','备注'],
    teachers, [6,10,8,16,16,10,12,8,6,10,28])

# ========== 收费标准清单（20项） ==========
fee_items = [
    ('学费(初中)', 8000, 10000, '超标准2000', '某市发改委批复8000'),
    ('学费(高中)', 10000, 12000, '超标准2000', '某市发改委批复10000'),
    ('住宿费', 1500, 1500, '无差异', '某市发改委批复1500'),
    ('伙食费', 3000, 3000, '无差异', '成本核算'),
    ('代收教材费', 500, 800, '超标准300', '某市发改委批复500'),
    ('校服费', 300, 500, '超标准200', '某市发改委批复300'),
    ('保险费', 100, 150, '超标准50', '自愿原则'),
    ('课后服务费', 600, 600, '无差异', '某市发改委批复600'),
    ('作业本费', 50, 50, '无差异', '成本核算'),
    ('考试费', 80, 80, '无差异', '成本核算'),
    ('军训费', 200, 200, '无差异', '成本核算'),
    ('体检费', 100, 100, '无差异', '成本核算'),
    ('社会实践费', 150, 150, '无差异', '成本核算'),
    ('电教费', 120, 200, '超标准80', '某市发改委批复120'),
    ('图书馆使用费', 80, 80, '无差异', '成本核算'),
    ('体育馆使用费', 100, 100, '无差异', '成本核算'),
    ('实验室使用费', 150, 150, '无差异', '成本核算'),
    ('计算机使用费', 120, 200, '超标准80', '某市发改委批复120'),
    ('空调使用费', 100, 100, '无差异', '成本核算'),
    ('饮水费', 50, 80, '超标准30', '成本核算应为50'),
]
fee_data = []
for i, (name, approved, actual, diff, basis) in enumerate(fee_items, 1):
    fee_data.append([i, name, approved, actual, diff, basis, '合规' if actual<=approved else '违规', '无' if actual<=approved else '退还超收费用'])
make_sheet(wb, '收费标准清单', ['序号','收费项目','批准标准(元/年)','实际收取(元/年)','差异说明','依据','审查结论','处理建议'],
    fee_data, [6,18,16,16,18,24,10,18])

# ========== 财务报表（30项） ==========
fin_data = []
fin_items = [
    ('学费收入', '1200万', '1150万', '收费人数1200人×10000元', '低于预期50万', '异常'),
    ('住宿费收入', '180万', '180万', '住宿人数1200人×1500元', '符合预算', '正常'),
    ('伙食费收入', '360万', '360万', '就餐人数1200人×3000元', '符合预算', '正常'),
    ('人员经费', '1500万', '1550万', '教职工95人平均工资', '超支50万', '异常'),
    ('公用经费', '300万', '350万', '日常办公支出', '超支50万', '异常'),
    ('设备购置费', '200万', '220万', '教学设备采购', '超支20万', '异常'),
    ('政府补助收入', '500万', '500万', '财政拨款', '符合预算', '正常'),
    ('捐赠收入', '100万', '100万', '社会捐赠', '符合预算', '正常'),
    ('资产负债率', '45%', '48%', '负债总额/资产总额', '上升3个百分点', '异常'),
    ('办学积累', '200万', '150万', '年度结余', '下降50万', '异常'),
]
for i, (name, budget, actual, calc, diff, conclusion) in enumerate(fin_items, 1):
    fin_data.append([i, name, budget, actual, calc, diff, conclusion])
# 补充20项正常财务项目
for i in range(11, 31):
    fin_data.append([i, f'财务项目{i}', f'{i*10}万', f'{i*10}万', '按预算执行', '符合预算', '正常'])
make_sheet(wb, '财务报表', ['序号','财务项目','预算金额','实际金额','计算说明','差异分析','结论'],
    fin_data, [6,18,12,12,24,18,10])

# ========== 安全检查记录（80项） ==========
safety = []
check_types = ['消防安全','设施安全','食品安全','校车安全','宿舍安全','校园安全']
levels = ['一般隐患','较大隐患','重大隐患','无隐患']
# 隐患未整改20项
for i in range(1, 21):
    safety.append([i, f'SF{i:03d}', f'2024-{(i%12)+1:02d}-15', check_types[i%6], levels[i%3], '未整改', f'{check_types[i%6]}检查发现隐患，尚未完成整改'])
# 检查不合格15项（重大隐患）
for i in range(21, 36):
    safety.append([i, f'SF{i:03d}', f'2024-{(i%12)+1:02d}-20', check_types[i%6], '重大隐患', '整改中', f'{check_types[i%6]}检查发现重大安全隐患'])
# 已整改复查合格10项（合规保留）
for i in range(36, 46):
    safety.append([i, f'SF{i:03d}', f'2024-{(i%12)+1:02d}-25', check_types[i%6], '一般隐患', '已整改', '隐患已整改完成，复查合格'])
# 正常35项
for i in range(46, 81):
    safety.append([i, f'SF{i:03d}', f'2024-{(i%12)+1:02d}-28', check_types[i%6], '无隐患', '-', '检查未发现安全隐患'])
make_sheet(wb, '安全检查记录', ['序号','检查编号','检查日期','检查类型','隐患等级','整改状态','备注'],
    safety, [6,10,14,12,12,10,36])

# ========== 设施设备清单（100项） ==========
facilities = []
fac_names = ['教学楼','实验楼','宿舍楼','食堂','体育馆','图书馆','操场','配电室','消防设施','监控设备']
# 超期未检25项
for i in range(1, 26):
    facilities.append([i, f'FA{i:03d}', f'{fac_names[i%10]}{i}', '超期未检', '上次检验日期为2022年，已超过检验周期'])
# 设备老化20项
for i in range(26, 46):
    facilities.append([i, f'FA{i:03d}', f'{fac_names[i%10]}{i}', '设备老化', '设备使用年限超过10年，存在安全隐患'])
# 新建验收合格10项（合规保留）
for i in range(46, 56):
    facilities.append([i, f'FA{i:03d}', f'{fac_names[i%10]}{i}', '新建验收合格', '2024年新建并通过验收，状态良好'])
# 正常45项
for i in range(56, 101):
    facilities.append([i, f'FA{i:03d}', f'{fac_names[i%10]}{i}', '正常', '设施状态良好，定期检验'])
make_sheet(wb, '设施设备清单', ['序号','设施编号','设施名称','设施状态','备注'],
    facilities, [6,10,18,14,40])

# ========== 招生计划与实际（60项） ==========
enrollment = []
grades = ['初一','初二','初三','高一','高二','高三']
# 超计划招生15项
for i in range(1, 16):
    plan = 200; actual = 200 + (i % 5 + 1) * 10
    enrollment.append([i, f'RE{i:03d}', '2024', grades[i%6], plan, actual, f'超计划招生{actual-plan}人'])
# 简章未备案12项
for i in range(16, 28):
    enrollment.append([i, f'RE{i:03d}', '2024', grades[i%6], 150, 150, '招生简章未报教育主管部门备案'])
# 按计划已备案8项（合规保留）
for i in range(28, 36):
    enrollment.append([i, f'RE{i:03d}', '2024', grades[i%6], 180, 180, '按计划招生，招生简章已备案'])
# 正常25项
for i in range(36, 61):
    enrollment.append([i, f'RE{i:03d}', '2024', grades[i%6], 200, 200, ''])
make_sheet(wb, '招生计划与实际', ['序号','招生编号','年度','年级','计划人数','实际人数','备注'],
    enrollment, [6,10,8,8,10,10,36])

# ========== 学籍管理记录（120项） ==========
students = []
# 学籍不完整25项
for i in range(1, 26):
    students.append([i, f'ST{i:03d}', f'学生{i}', '初一', '在籍在读', '学籍档案缺少身份证号和户籍信息'])
# 人籍分离20项
for i in range(26, 46):
    students.append([i, f'ST{i:03d}', f'学生{i}', '初二', '在籍在读', '学生实际不在校就读，学籍仍保留'])
# 违规转学15项
for i in range(46, 61):
    students.append([i, f'ST{i:03d}', f'学生{i}', '初三', '在籍在读', '转学手续不完整，未经教育主管部门审批'])
# 完整准确8项（合规保留）
for i in range(61, 69):
    students.append([i, f'ST{i:03d}', f'学生{i}', '高一', '在籍在读', '学籍信息完整，人籍一致'])
# 正常52项
for i in range(69, 121):
    students.append([i, f'ST{i:03d}', f'学生{i}', grades2:=['初一','初二','初三','高一','高二','高三'][i%6], '在籍在读', ''])
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

wb.save('/workspace/mb_v2_tmp.xlsx')
print(f"教师:{len(teachers)} 安全:{len(safety)} 设施:{len(facilities)} 招生:{len(enrollment)} 学籍:{len(students)} 收费:{len(fee_data)} 财务:{len(fin_data)}")
print(f"总数据项: {len(teachers)+len(safety)+len(facilities)+len(enrollment)+len(students)+len(fee_data)+len(fin_data)}")
