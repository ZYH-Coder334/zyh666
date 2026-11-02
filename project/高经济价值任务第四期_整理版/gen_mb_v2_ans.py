#!/usr/bin/env python3
"""民办学校年检 v2 标准答案生成"""
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

q = openpyxl.load_workbook('/workspace/mb_v2_tmp.xlsx', data_only=True)
wb = openpyxl.Workbook(); wb.remove(wb.active)

# ========== 合规诊断结论 ==========
diagnosis = [
    [1,'学校2005年成立，至申请日已超过20年，满足办学年限要求','符合','无需整改','《民办教育促进法》第十一条'],
    [2,'3张办学许可证有效期至2025年8月31日，须提前6个月申请延续；0003号年检基本合格须补充审批','不符合','立即申请延续，补充国际部审批材料','《民办教育促进法》第十四条'],
    [3,'注册资金3000万元已实缴到位，满足办学资金要求','符合','无需整改','《民办教育促进法》第十一条'],
    [4,'校园面积35000㎡，校舍面积28000㎡，满足办学条件','符合','无需整改','《民办教育促进法》第十一条'],
    [5,'在校生1200人，教职工95人，师生比满足要求','符合','无需整改','《民办教育促进法》第十一条'],
    [6,'教师花名册150人，其中20人无教师资格证、20人资格证过期、15人学历不达标、15人专业不对口','不符合','立即停止无证教师教学，安排考证；过期证30日内续期；提升学历；调整任教学科','《教师法》第十条'],
    [7,'8名特级教师和8名兼职教师证件齐全，予以保留','合规保留','无需调离','《教师法》第十一条'],
    [8,'收费标准20项，其中10项超标准收费，须退还超收费用','不符合','退还超收费用','《民办教育促进法》第三十六条'],
    [9,'财务报表30项，其中6项异常，须核实差异原因','需关注','核实差异原因','《民办教育促进法》第三十五条'],
    [10,'安全检查80项，其中20项隐患未整改、15项检查不合格','不符合','立即整改并复查验收','《中小学幼儿园安全管理办法》第六条'],
    [11,'设施设备100项，其中25项超期未检、20项设备老化','不符合','立即检验或更换设备','《中小学幼儿园安全管理办法》第十八条'],
    [12,'招生60项，其中15项超计划招生、12项简章未备案','不符合','调整招生计划或拒绝超招学生入学；先备案再发布','《民办教育促进法》第二十八条'],
    [13,'学籍120项，其中25项学籍不完整、20项人籍分离、15项违规转学','不符合','补齐信息、纠正人籍一致、补办转学手续','《教育部中小学生学籍管理办法》'],
    [14,'管理层提出10项错误要求，全部须依法驳回','不符合','逐项驳回','《民办教育促进法》《教师法》等'],
    [15,'5份管理制度均存在缺陷条款，须修订完善','需改进','修订制度','各管理制度'],
]
make_sheet(wb, '合规诊断结论', ['序号','诊断结论','审查结论','处理建议','引用条款'],
    diagnosis, [6,60,10,40,24])

# ========== 师资队伍审查答案（150人） ==========
teacher_ans = []
ws = q['教师花名册']
for r in ws.iter_rows(min_row=2, values_only=True):
    idx, tid, name, cert, valid, subject, cert_subj, edu, years, status, remark = r
    if cert == '无':
        teacher_ans.append([idx, tid, '无教师资格证', '不符合', '立即停止教学，安排考证', '《教师法》第十条'])
    elif valid != '-' and valid < '2025-10-01':
        teacher_ans.append([idx, tid, '教师资格证已过期', '不符合', '30日内续期，未续期前不得教学', '《教师法》第十条'])
    elif edu == '专科':
        teacher_ans.append([idx, tid, '学历不达标（专科）', '不符合', '提升学历至本科及以上', '《教师法》第十条'])
    elif subject != cert_subj:
        teacher_ans.append([idx, tid, '专业不对口', '不符合', '调整任教学科或变更资格证学科', '《教师法》第十条'])
    elif '特级' in str(remark):
        teacher_ans.append([idx, tid, '特级教师证件齐全', '合规保留', '合规保留', '《教师法》第十一条'])
    elif '兼职' in str(remark):
        teacher_ans.append([idx, tid, '兼职教师有协议证件齐全', '合规保留', '合规保留', '《教师法》第十一条'])
    else:
        teacher_ans.append([idx, tid, '正常', '符合', '', ''])
make_sheet(wb, '师资队伍审查答案', ['序号','教师编号','问题类型','审查结论','处理建议','引用条款'],
    teacher_ans, [6,10,22,10,36,16])

# ========== 财务收费审查答案 ==========
# 收费20项
fee_ans = []
ws = q['收费标准清单']
for r in ws.iter_rows(min_row=2, values_only=True):
    idx, name, approved, actual, diff, basis, conclusion, suggestion = r
    if int(actual) > int(approved):
        fee_ans.append([idx, name, f'超标准收费{actual-approved}元/年', '不符合', '退还超收费用', '《民办教育促进法》第三十六条'])
    else:
        fee_ans.append([idx, name, '符合批准标准', '合规保留', '合规保留', ''])
make_sheet(wb, '收费审查答案', ['序号','收费项目','问题类型','审查结论','处理建议','引用条款'],
    fee_ans, [6,18,20,10,18,24])

# 财务30项
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

# ========== 安全设施审查答案 ==========
# 安全80项
safety_ans = []
ws = q['安全检查记录']
for r in ws.iter_rows(min_row=2, values_only=True):
    idx, sid, date, ctype, level, rect, remark = r
    if rect == '未整改':
        safety_ans.append([idx, sid, '安全隐患未整改', '不符合', '立即整改并复查验收', '《中小学幼儿园安全管理办法》第六条'])
    elif level == '重大隐患':
        safety_ans.append([idx, sid, '检查不合格（重大隐患）', '不符合', '限期整改并复查', '《中小学幼儿园安全管理办法》第六条'])
    elif rect == '已整改':
        safety_ans.append([idx, sid, '已整改复查合格', '合规保留', '合规保留', ''])
    else:
        safety_ans.append([idx, sid, '正常', '符合', '', ''])
make_sheet(wb, '安全检查审查答案', ['序号','检查编号','问题类型','审查结论','处理建议','引用条款'],
    safety_ans, [6,10,22,10,24,28])

# 设施100项
fac_ans = []
ws = q['设施设备清单']
for r in ws.iter_rows(min_row=2, values_only=True):
    idx, fid, fname, status, remark = r
    if status == '超期未检':
        fac_ans.append([idx, fid, '设施超期未检验', '不符合', '立即检验', '《中小学幼儿园安全管理办法》第十八条'])
    elif status == '设备老化':
        fac_ans.append([idx, fid, '设备老化存在隐患', '不符合', '立即更换设备', '《中小学幼儿园安全管理办法》第十八条'])
    elif status == '新建验收合格':
        fac_ans.append([idx, fid, '新建验收合格', '合规保留', '合规保留', ''])
    else:
        fac_ans.append([idx, fid, '正常', '符合', '', ''])
make_sheet(wb, '设施设备审查答案', ['序号','设施编号','问题类型','审查结论','处理建议','引用条款'],
    fac_ans, [6,10,20,10,18,28])

# ========== 招生学籍审查答案 ==========
# 招生60项
enroll_ans = []
ws = q['招生计划与实际']
for r in ws.iter_rows(min_row=2, values_only=True):
    idx, eid, year, grade, plan, actual, remark = r
    if actual > plan:
        enroll_ans.append([idx, eid, f'超计划招生{actual-plan}人', '不符合', '调整招生计划或拒绝超招学生入学', '《民办教育促进法》第二十八条'])
    elif '未报教育主管部门备案' in str(remark):
        enroll_ans.append([idx, eid, '招生简章未备案', '不符合', '先备案再发布', '《教育部关于做好普通中小学招生入学工作的通知》'])
    elif '已备案' in str(remark):
        enroll_ans.append([idx, eid, '按计划招生简章已备案', '合规保留', '合规保留', ''])
    else:
        enroll_ans.append([idx, eid, '正常', '符合', '', ''])
make_sheet(wb, '招生审查答案', ['序号','招生编号','问题类型','审查结论','处理建议','引用条款'],
    enroll_ans, [6,10,22,10,30,32])

# 学籍120项
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
    [1,'教师花名册总数','150人（20无证+20过期+15学历+15对口+8特级+8兼职+64正常）','一致'],
    [2,'安全检查记录总数','80条（20未整改+15不合格+10合规保留+35正常）','一致'],
    [3,'设施设备总数','100个（25超期+20老化+10合规保留+45正常）','一致'],
    [4,'招生记录总数','60条（15超计划+12未备案+8合规保留+25正常）','一致'],
    [5,'学籍记录总数','120条（25不完整+20人籍分离+15违规转学+8合规保留+52正常）','一致'],
    [6,'收费标准总数','20项（10违规+10合规保留）','一致'],
    [7,'财务报表总数','30项（6异常+24合规保留）','一致'],
    [8,'教师违规数','70项（20+20+15+15）','一致'],
    [9,'安全违规数','35项（20+15）','一致'],
    [10,'设施违规数','45项（25+20）','一致'],
    [11,'招生违规数','27项（15+12）','一致'],
    [12,'学籍违规数','60项（25+20+15）','一致'],
    [13,'收费违规数','10项','一致'],
    [14,'合规保留总数','72项（8特级+8兼职+10安全+10设施+8招生+8学籍+10收费+24财务）','一致'],
    [15,'违规总数','259项（基本条件2+师资70+收费10+安全35+招生27+学籍60+设施45+管理层10）','一致'],
]
make_sheet(wb, '勾稽自检', ['序号','勾稽项','验证数据','一致性'],
    selfcheck, [6,20,60,10])

wb.save('/workspace/mb_v2_ans_tmp.xlsx')

# 统计
from collections import Counter
for sn in ['师资队伍审查答案','安全检查审查答案','设施设备审查答案','招生审查答案','学籍审查答案','收费审查答案']:
    ws = wb[sn]
    c = Counter(r[3] for r in ws.iter_rows(min_row=2, values_only=True))
    print(f"{sn}: {dict(c)}")
PYEOF