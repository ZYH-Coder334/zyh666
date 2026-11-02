#!/usr/bin/env python3
"""民办学校年检 v3 标准答案生成 - 含日期推算与百分比阈值计算"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import date
from dateutil.relativedelta import relativedelta

THIN = Side(style="thin", color="CCCCCC")
BORDER = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
HDR_FILL = PatternFill("solid", fgColor="2B5B84")
HDR_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=10)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

REF_DATE = date(2025, 9, 1)
HIGH_SCHOOL_SUBJECTS = {'语文','数学','英语','物理','化学','生物','政治','历史','地理'}

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

q = openpyxl.load_workbook('/workspace/mb_v3_tmp.xlsx', data_only=True)
wb = openpyxl.Workbook(); wb.remove(wb.active)

# ========== 合规诊断结论 ==========
diagnosis = [
    [1,'学校2005年成立，至申请日已超过20年，满足办学年限要求','符合','无需整改','《民办教育促进法》第十一条'],
    [2,'3张办学许可证有效期至2025年8月31日，须提前6个月申请延续；0003号年检基本合格须补充审批','不符合','立即申请延续，补充国际部审批材料','《民办教育促进法》第十四条'],
    [3,'注册资金3000万元已实缴到位，满足办学资金要求','符合','无需整改','《民办教育促进法》第十一条'],
    [4,'校园面积35000㎡，校舍面积28000㎡，满足办学条件','符合','无需整改','《民办教育促进法》第十一条'],
    [5,'在校生1800人，教职工250人，师生比满足要求','符合','无需整改','《民办教育促进法》第十一条'],
    [6,'教师花名册250人：无证40人、资格证过期40人、学历不达标30人、专业不对口30人','不符合','停止无证教师教学并安排考证；过期证30日内续期；提升学历；调整任教学科','《教师法》第十条'],
    [7,'20名特级教师和20名兼职教师证件齐全，予以保留','合规保留','无需调离','《教师法》第十一条'],
    [8,'收费标准40项，其中20项超批复标准收费，须退还超收费用','不符合','退还超收费用','《民办教育促进法》第三十六条'],
    [9,'财务报表50项，其中6项异常，须核实差异原因','需关注','核实差异原因','《民办教育促进法》第三十五条'],
    [10,'安全检查180项：未整改40项、重大隐患30项、整改超期25项','不符合','立即整改并复查验收','《中小学幼儿园安全管理办法》第六条'],
    [11,'设施设备180项：超期未检39项、设备老化35项','不符合','立即检验或更换设备','《中小学幼儿园安全管理办法》第十八条'],
    [12,'招生120项：超计划招生30项、简章未备案25项','不符合','调整招生计划或拒绝超招学生入学；先备案再发布','《民办教育促进法》第二十八条'],
    [13,'学籍250项：学籍不完整50项、人籍分离40项、违规转学30项','不符合','补齐信息、纠正人籍一致、补办转学手续','《教育部中小学生学籍管理办法》'],
    [14,'管理层提出10项错误要求，全部须依法驳回','不符合','逐项驳回','《民办教育促进法》《教师法》等'],
    [15,'5份管理制度均存在缺陷条款，须修订完善','需改进','修订制度','各管理制度'],
]
make_sheet(wb, '合规诊断结论', ['序号','诊断结论','审查结论','处理建议','引用条款'],
    diagnosis, [6,60,10,40,24])

# ========== 师资队伍审查答案（250人） ==========
teacher_ans = []
ws = q['教师花名册']
for r in ws.iter_rows(min_row=2, values_only=True):
    idx, tid, name, cert, valid, subject, cert_subj, edu, years, status, remark = r
    if cert == '待办理':
        teacher_ans.append([idx, tid, '无教师资格证', '不符合', '立即停止教学，安排考证', '《教师法》第十条'])
    elif valid != '-' and date.fromisoformat(str(valid)) < REF_DATE:
        teacher_ans.append([idx, tid, '教师资格证已过期', '不符合', '30日内续期，未续期前不得教学', '《教师法》第十条'])
    elif edu == '专科' and subject in HIGH_SCHOOL_SUBJECTS:
        teacher_ans.append([idx, tid, '学历不达标（高中教师须本科及以上）', '不符合', '提升学历至本科及以上', '《教师法》第十条'])
    elif subject != cert_subj:
        teacher_ans.append([idx, tid, '专业不对口', '不符合', '调整任教学科或变更资格证学科', '《教师法》第十条'])
    elif '特级' in str(remark):
        teacher_ans.append([idx, tid, '特级教师证件齐全', '合规保留', '合规保留', '《教师法》第十一条'])
    elif '兼职' in str(status):
        teacher_ans.append([idx, tid, '兼职教师有协议证件齐全', '合规保留', '合规保留', '《教师法》第十一条'])
    else:
        teacher_ans.append([idx, tid, '正常', '符合', '', ''])
make_sheet(wb, '师资队伍审查答案', ['序号','教师编号','问题类型','审查结论','处理建议','引用条款'],
    teacher_ans, [6,10,28,10,36,16])

# ========== 收费审查答案（跨表查询收费批复标准） ==========
fee_ans = []
# 构建批复标准查询字典
appr_dict = {}
ws_appr = q['收费批复标准']
for r in ws_appr.iter_rows(min_row=2, values_only=True):
    appr_dict[str(r[1])] = int(r[2])
ws = q['收费标准清单']
for r in ws.iter_rows(min_row=2, values_only=True):
    idx, name, actual, basis = r
    approved = appr_dict.get(str(name), 0)
    if int(actual) > approved:
        fee_ans.append([idx, name, f'超批复标准{int(actual)-approved}元/年', '不符合', '退还超收费用', '《民办教育促进法》第三十六条'])
    else:
        fee_ans.append([idx, name, '符合批复标准', '合规保留', '合规保留', ''])
make_sheet(wb, '收费审查答案', ['序号','收费项目','问题类型','审查结论','处理建议','引用条款'],
    fee_ans, [6,18,24,10,18,24])

# ========== 财务审查答案（50项） ==========
fin_ans = []
ws = q['财务报表']
for r in ws.iter_rows(min_row=2, values_only=True):
    idx, name, budget, actual, calc, diff, conclusion = r
    if conclusion == '异常':
        fin_ans.append([idx, name, f'{diff}', '需关注', '核实差异原因', '《民办教育促进法》第三十五条'])
    else:
        fin_ans.append([idx, name, '符合预算', '合规保留', '合规保留', ''])
make_sheet(wb, '财务审查答案', ['序号','财务项目','问题类型','审查结论','处理建议','引用条款'],
    fin_ans, [6,18,18,10,18,24])

# ========== 安全检查审查答案（180项） ==========
safety_ans = []
ws = q['安全检查记录']
for r in ws.iter_rows(min_row=2, values_only=True):
    idx, sid, sdate, ctype, level, rect, deadline, remark = r
    if rect == '未整改':
        safety_ans.append([idx, sid, '安全隐患未整改', '不符合', '立即整改并复查验收', '《中小学幼儿园安全管理办法》第六条'])
    elif level == '重大隐患':
        safety_ans.append([idx, sid, '检查不合格（重大隐患）', '不符合', '限期整改并复查', '《中小学幼儿园安全管理办法》第六条'])
    elif rect == '整改中' and deadline != '-' and date.fromisoformat(str(deadline)) < REF_DATE:
        safety_ans.append([idx, sid, '整改超期未完成', '不符合', '限期完成整改', '《中小学幼儿园安全管理办法》第六条'])
    elif rect == '已整改':
        safety_ans.append([idx, sid, '已整改复查合格', '合规保留', '合规保留', ''])
    else:
        safety_ans.append([idx, sid, '正常', '符合', '', ''])
make_sheet(wb, '安全检查审查答案', ['序号','检查编号','问题类型','审查结论','处理建议','引用条款'],
    safety_ans, [6,10,22,10,24,28])

# ========== 设施设备审查答案（180项，含日期推算） ==========
fac_ans = []
ws = q['设施设备清单']
for r in ws.iter_rows(min_row=2, values_only=True):
    idx, fid, fname, last_date, cycle, buy_date, status, remark = r
    last = date.fromisoformat(str(last_date))
    next_due = last + relativedelta(months=int(cycle))
    buy = date.fromisoformat(str(buy_date))
    if next_due < REF_DATE:
        fac_ans.append([idx, fid, f'设施超期未检验（下次应检{next_due}）', '不符合', '立即检验', '《中小学幼儿园安全管理办法》第十八条'])
    elif buy < date(2010, 1, 1):
        fac_ans.append([idx, fid, '设备老化（购置超过15年）', '不符合', '立即更换设备', '《中小学幼儿园安全管理办法》第十八条'])
    elif buy >= date(2022, 1, 1):
        fac_ans.append([idx, fid, '近年新购设备', '合规保留', '合规保留', ''])
    else:
        fac_ans.append([idx, fid, '正常', '符合', '', ''])
make_sheet(wb, '设施设备审查答案', ['序号','设施编号','问题类型','审查结论','处理建议','引用条款'],
    fac_ans, [6,10,30,10,18,28])

# ========== 招生审查答案（120项，跨表查询招生计划批复+百分比阈值） ==========
enroll_ans = []
# 构建批复计划查询字典
plan_dict = {}
ws_plan = q['招生计划批复']
for r in ws_plan.iter_rows(min_row=2, values_only=True):
    plan_dict[str(r[1])] = int(r[3])
ws = q['招生计划与实际']
for r in ws.iter_rows(min_row=2, values_only=True):
    idx, eid, year, grade, actual, filed, remark = r
    plan = plan_dict.get(str(eid), 0)
    threshold = plan * 1.05
    if int(actual) > threshold:
        enroll_ans.append([idx, eid, f'超计划招生{int(actual)-plan}人（超{round((int(actual)-plan)/plan*100,1)}%）', '不符合', '调整招生计划或拒绝超招学生入学', '《民办教育促进法》第二十八条'])
    elif filed == '未备案':
        enroll_ans.append([idx, eid, '招生简章未备案', '不符合', '先备案再发布', '《教育部关于做好普通中小学招生入学工作的通知》'])
    elif '已备案' in str(filed) and int(actual) > plan:
        enroll_ans.append([idx, eid, f'实际略超计划{int(actual)-plan}人但未超5%阈值', '合规保留', '合规保留', '《民办教育促进法》第二十八条'])
    elif '已备案' in str(filed) and int(actual) == plan:
        enroll_ans.append([idx, eid, '正常', '符合', '', ''])
    else:
        enroll_ans.append([idx, eid, '正常', '符合', '', ''])
make_sheet(wb, '招生审查答案', ['序号','招生编号','问题类型','审查结论','处理建议','引用条款'],
    enroll_ans, [6,10,32,10,30,32])

# ========== 学籍审查答案（250项） ==========
stu_ans = []
ws = q['学籍管理记录']
for r in ws.iter_rows(min_row=2, values_only=True):
    idx, sid, sname, grade, status, remark = r
    if '缺少身份证号和户籍信息' in str(remark):
        stu_ans.append([idx, sid, '学籍信息不完整', '不符合', '补齐身份证号和户籍信息', '《教育部中小学生学籍管理办法》'])
    elif '实际不在校就读' in str(remark):
        stu_ans.append([idx, sid, '人籍分离', '不符合', '纠正人籍一致', '《教育部中小学生学籍管理办法》'])
    elif '转学手续不完整' in str(remark):
        stu_ans.append([idx, sid, '违规转学', '不符合', '补办转学手续', '《教育部中小学生学籍管理办法》'])
    elif '学籍信息完整' in str(remark):
        stu_ans.append([idx, sid, '学籍完整信息准确', '合规保留', '合规保留', ''])
    else:
        stu_ans.append([idx, sid, '正常', '符合', '', ''])
make_sheet(wb, '学籍审查答案', ['序号','学籍号','问题类型','审查结论','处理建议','引用条款'],
    stu_ans, [6,10,18,10,28,28])

# ========== 错误要求驳回 ==========
reject = [
    [1,'超计划招生学生直接入学不调整计划','《民办教育促进法》第二十八条','调整招生计划或拒绝超招学生入学'],
    [2,'无教师资格证教师继续任教','《教师法》第十条','立即停止教学并安排考证'],
    [3,'超标准收费不退还','《民办教育促进法》第三十六条','退还超收费用'],
    [4,'安全隐患未整改继续教学','《中小学幼儿园安全管理办法》第六条','立即整改并复查验收'],
    [5,'学籍人籍分离不处理','《教育部中小学生学籍管理办法》','纠正人籍一致'],
    [6,'财务报表不报送教育部门','《民办教育促进法》第三十五条','定期报送教育部门'],
    [7,'招生简章不报备直接发布','《教育部关于做好普通中小学招生入学工作的通知》','先备案再发布'],
    [8,'设施超期未检继续使用','《中小学幼儿园安全管理办法》第十八条','立即检验'],
    [9,'违规转学不纠正','《教育部中小学生学籍管理办法》','补办转学手续'],
    [10,'许可证到期不申请延续','《民办教育促进法》第十四条','提前6个月申请延续'],
]
make_sheet(wb, '错误要求驳回', ['序号','错误要求','违反法规','处理'],
    reject, [6,36,36,30])

# ========== 勾稽自检 ==========
selfcheck = [
    [1,'教师花名册总数','250人（40无证+40过期+30学历+30对口+20特级+20兼职+70正常）','一致'],
    [2,'安全检查记录总数','180条（40未整改+30重大隐患+25整改超期+25合规保留+60正常）','一致'],
    [3,'设施设备总数','180个（39超期+35老化+25合规保留+81正常）','一致'],
    [4,'招生记录总数','120条（30超计划+25未备案+20合规保留+45正常）','一致'],
    [5,'学籍记录总数','250条（50不完整+40人籍分离+30违规转学+30合规保留+100正常）','一致'],
    [6,'收费标准总数','40项（20违规+20合规保留）','一致'],
    [7,'财务报表总数','50项（6异常+44合规保留）','一致'],
    [8,'教师违规数','140项（40+40+30+30）','一致'],
    [9,'安全违规数','95项（40+30+25）','一致'],
    [10,'设施违规数','74项（39+35）','一致'],
    [11,'招生违规数','55项（30+25）','一致'],
    [12,'学籍违规数','120项（50+40+30）','一致'],
    [13,'收费违规数','20项','一致'],
    [14,'合规保留总数','204项（20特级+20兼职+25安全+25设施+20招生+30学籍+20收费+44财务）','一致'],
    [15,'违规总数','522项（基本条件2+师资140+收费20+安全95+招生55+学籍120+设施74+管理层10+财务6）','一致'],
]
make_sheet(wb, '勾稽自检', ['序号','勾稽项','验证数据','一致性'],
    selfcheck, [6,20,60,10])

wb.save('/workspace/mb_v3_ans_tmp.xlsx')

# 统计验证
from collections import Counter
print("=== 答案统计 ===")
for sn in ['师资队伍审查答案','安全检查审查答案','设施设备审查答案','招生审查答案','学籍审查答案','收费审查答案','财务审查答案']:
    ws = wb[sn]
    c = Counter(r[3] for r in ws.iter_rows(min_row=2, values_only=True))
    print(f"{sn}: {dict(c)}")
print(f"\n违规合计: 基本条件2 + 师资140 + 收费20 + 财务6 + 安全95 + 设施74 + 招生55 + 学籍120 + 管理层10 = {2+140+20+6+95+74+55+120+10}")
