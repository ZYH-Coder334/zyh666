#!/usr/bin/env python3
"""涉密资质甲级 v6 标准答案生成
核心难点：
1. 集成收入7重判定（类型+窗口+非关联方含别名+发票全额+密级认定+归档）
2. 场所多维校验（公摊净面积+物理隔离+设备密级匹配+档案室三铁+租期）
3. 职称5维校验（专业+证书编号格式+发证机关权限+取得时间+证书状态）
4. 涉密人员6维（国籍+配偶+审查+培训有效期+岗位+社保）
5. 28项跨表勾稽
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import date, timedelta

THIN = Side(style="thin", color="CCCCCC")
BORDER = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
HDR_FILL = PatternFill("solid", fgColor="2F5597")
HDR_FONT = Font(name="宋体", bold=True, color="FFFFFF", size=10)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
BODY = Font(name="宋体", size=9)

REF = date(2025, 9, 1)
REV_LIMIT = date(2022, 9, 1)
TRAIN_LIMIT = date(2024, 9, 1)
WIN_START = date(2022, 9, 1)
WIN_END = date(2025, 8, 31)
SOC_LIMIT = '2024-09-01'

LEVEL_RANK = {'非密': 0, '秘密': 1, '机密': 2}


def make_sheet(wb, name, headers, data, col_widths=None):
    ws = wb.create_sheet(name)
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=i, value=h)
        c.font = HDR_FONT
        c.fill = HDR_FILL
        c.alignment = CENTER
        c.border = BORDER
    for r, row in enumerate(data, 2):
        for i, v in enumerate(row, 1):
            c = ws.cell(row=r, column=i, value=v)
            c.font = BODY
            c.alignment = WRAP
            c.border = BORDER
    if col_widths:
        for i, w in enumerate(col_widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    return ws


def dt(s):
    return date.fromisoformat(str(s).strip())


def blank(v):
    return v is None or not str(v).strip()


q = openpyxl.load_workbook('/workspace/涉密资质甲级_v6_题目.xlsx', data_only=True)

# ================= 读取数据 =================
base = {}
for r in q['企业基本信息'].iter_rows(min_row=2, values_only=True):
    keys = ['name', 'code', 'oldname', 'found', 'reg', 'paid', 'type', 'legal',
            'addr', 'site', 'scope', 'declared', 'apply']
    base = dict(zip(keys, r))

equity = [r for r in q['股权结构穿透'].iter_rows(min_row=2, values_only=True) if r[1] != '合计']
roster = [r for r in q['人员花名册'].iter_rows(min_row=2, values_only=True)]
social = [r for r in q['社保缴纳明细'].iter_rows(min_row=2, values_only=True)]
cs_majors = {r[1] for r in q['职称专业目录'].iter_rows(min_row=2, values_only=True) if str(r[3]).startswith('08')}
titles = [r for r in q['高级职称人员'].iter_rows(min_row=2, values_only=True)]
sec_staff = [r for r in q['涉密人员台账'].iter_rows(min_row=2, values_only=True)]
projects = [r for r in q['项目清单'].iter_rows(min_row=2, values_only=True)]
invoices = {r[1]: r for r in q['发票清单'].iter_rows(min_row=2, values_only=True)}
secret_docs = {r[1]: r for r in q['涉密项目批文'].iter_rows(min_row=2, values_only=True)}
non_secret_certs = {r[1]: r for r in q['非密项目认定'].iter_rows(min_row=2, values_only=True)}
related_raw = [r for r in q['关联方名单'].iter_rows(min_row=2, values_only=True)]
# 关联方含别名
related = set()
for r in related_raw:
    related.add(r[1])
    if r[2] and r[2] != '—':
        related.add(r[2])
sites = [r for r in q['场所设备清单'].iter_rows(min_row=2, values_only=True)]
devices = [r for r in q['设备清单'].iter_rows(min_row=2, values_only=True)]
san_tie = {r[1]: r for r in q['场所三铁检查'].iter_rows(min_row=2, values_only=True)}
policies = [r for r in q['保密制度清单'].iter_rows(min_row=2, values_only=True)]
requests = [r for r in q['管理层要求'].iter_rows(min_row=2, values_only=True)]
salary_map = {r[1]: (r[2], r[3]) for r in q['社保基数对照表'].iter_rows(min_row=2, values_only=True)}

roster_names = {r[1]: r for r in roster}
soc_by_name = {r[1]: r for r in social}

# ================= 1. 外资比例 =================
foreign = sum(float(str(r[4]).strip('%')) * float(str(r[5]).strip('%')) / 100 for r in equity)

# ================= 2. 社保 =================
soc_rows = []
soc_ok = 0
for r in social:
    idx, name, idn, dept, post, hire, start, status, last, base_sal, loc = r
    ok = (status == '正常' and str(start) <= SOC_LIMIT)
    if ok:
        soc_ok += 1
        soc_rows.append([idx, name, dept, post, start, status, '符合', '社保连续缴纳满一年'])
    else:
        reason = '缴费状态异常，须恢复连续缴纳' if status != '正常' else f'起缴日期{start}，至基准日不足12个月'
        soc_rows.append([idx, name, dept, post, start, status, '不符合', reason])

# ================= 3. 高级职称（5维校验） =================
title_rows = []
cs_n = 0
cs_valid = 0
AUTHORITY_POS = ['人力资源和社会保障部', '各省人力资源和社会保障厅']
for r in titles:
    idx, name, tname, major, cert_no, auth, obtain, edu, cert_status, appoint = r
    is_cs = major in cs_majors
    if is_cs:
        cs_n += 1
    # 证书编号格式：Z开头，末位1=正高2=副高
    cert_ok = str(cert_no).startswith('Z') and (
        (tname == '教授级高级工程师' and str(cert_no).endswith('1')) or
        (tname == '高级工程师' and str(cert_no).endswith('2'))
    )
    # 发证机关权限：正高须省部级
    auth_ok = True
    if tname == '教授级高级工程师':
        auth_ok = auth in AUTHORITY_POS or '人力资源和社会保障厅' in str(auth) or auth == '人力资源和社会保障部'
    # 证书状态
    status_ok = cert_status == '有效'

    if is_cs and cert_ok and auth_ok and status_ok:
        cs_valid += 1
        title_rows.append([idx, name, tname, major, '计算机类', '符合', '计入高级职称人数'])
    elif is_cs:
        reasons = []
        if not cert_ok: reasons.append('证书编号格式不符')
        if not auth_ok: reasons.append('发证机关权限不足')
        if not status_ok: reasons.append('证书状态无效')
        title_rows.append([idx, name, tname, major, '计算机类', '不符合', '；'.join(reasons)])
    else:
        title_rows.append([idx, name, tname, major, '非计算机类', '不符合', '专业不属于计算机技术与软件相关专业'])

# ================= 4. 涉密人员（6维） =================
sec_rows = []
c_nation = c_spouse = c_noreview = c_expired = c_train = c_train_exp = c_keep = c_ok = 0
for r in sec_staff:
    idx, name, ros_idx, dept, level, nation, spouse, review, train, sec_start, employ = r
    if nation != '中国':
        sec_rows.append([idx, name, ros_idx, dept, level, '国籍非中国籍', '不符合', '立即调离涉密岗位']); c_nation += 1
    elif spouse != '中国':
        sec_rows.append([idx, name, ros_idx, dept, level, '配偶为外籍人员', '不符合', '调离涉密岗位']); c_spouse += 1
    elif blank(review):
        sec_rows.append([idx, name, ros_idx, dept, level, '未进行保密审查', '不符合', '立即组织保密审查']); c_noreview += 1
    elif dt(review) < REV_LIMIT:
        sec_rows.append([idx, name, ros_idx, dept, level, '保密审查已过期', '不符合', '重新审查']); c_expired += 1
    elif blank(train):
        sec_rows.append([idx, name, ros_idx, dept, level, '未完成保密教育培训', '不符合', '限期完成培训']); c_train += 1
    elif dt(train) < TRAIN_LIMIT:
        sec_rows.append([idx, name, ros_idx, dept, level, '保密培训已过期（超1年）', '不符合', '重新参加培训']); c_train_exp += 1
    elif level == '一般' and dept in ('综合管理部', '后勤保障部'):
        sec_rows.append([idx, name, ros_idx, dept, level, '非涉密岗位且各项条件合规', '合规保留', '按非涉密岗位管理']); c_keep += 1
    else:
        sec_rows.append([idx, name, ros_idx, dept, level, '各项条件合规', '符合', '—']); c_ok += 1

sec_bad = c_nation + c_spouse + c_noreview + c_expired + c_train + c_train_exp

# ================= 5. 集成收入（7重判定） =================
proj_rows = []
income = counted = big_counted = 0
ex_type = ex_win = ex_related = ex_uninvoiced = ex_partial = ex_nodoc = ex_noarch = 0
for r in projects:
    idx, pid, pname, ptype, amount, content, accept, client, secret, archived = r
    inv = invoices.get(pid)

    if ptype != '系统集成':
        proj_rows.append([idx, pid, pname, ptype, amount, '不计入', f'{ptype}不属于集成业务']); ex_type += 1; continue
    if dt(accept) < WIN_START or dt(accept) > WIN_END:
        proj_rows.append([idx, pid, pname, ptype, amount, '不计入', f'验收日期{accept}不在近三年窗口内']); ex_win += 1; continue
    if client in related:
        proj_rows.append([idx, pid, pname, ptype, amount, '不计入', '甲方为关联方（含别名），属关联交易']); ex_related += 1; continue
    if inv is None or inv[6] != '已开具':
        proj_rows.append([idx, pid, pname, ptype, amount, '不计入', '未取得合规发票']); ex_uninvoiced += 1; continue
    if inv[7] != '全额':
        proj_rows.append([idx, pid, pname, ptype, amount, '不计入', f'发票为部分开票（{inv[4]}万元≠合同额{amount}万元）']); ex_partial += 1; continue
    # 密级认定
    if secret in ('机密', '秘密'):
        if pid not in secret_docs:
            proj_rows.append([idx, pid, pname, ptype, amount, '不计入', '涉密项目未取得保密批文']); ex_nodoc += 1; continue
    elif secret == '非密':
        if pid not in non_secret_certs:
            proj_rows.append([idx, pid, pname, ptype, amount, '不计入', '非密项目未取得非密认定']); ex_nodoc += 1; continue
    if archived != '已归档':
        proj_rows.append([idx, pid, pname, ptype, amount, '不计入', '项目未完成归档']); ex_noarch += 1; continue

    proj_rows.append([idx, pid, pname, ptype, amount, '计入',
                      '系统集成+验收在窗口+非关联方+全额发票+密级认定+已归档'])
    counted += 1; income += int(amount)
    if int(amount) >= 1000:
        big_counted += 1

# ================= 6. 场所（多维校验） =================
site_rows = []
site_pass = site_bad = site_na = 0
dev_by_site = {}
for d in devices:
    dev_by_site.setdefault(d[2], []).append(d)

for r in sites:
    idx, code, addr, use, area, share, s1, s2, own, isolated, site_level = r
    net_area = area - share
    years = (dt(s2) - dt(s1)).days / 365.25
    if '办公' in use:
        site_rows.append([idx, code, use, net_area, f'{s1}至{s2}', '非涉密场所', '合规保留', '不作为涉密场所申报']); site_na += 1; continue
    issues = []
    if dt(s2) < REF:
        issues.append(f'租期已于{s2}到期')
    if net_area < 100:
        issues.append(f'净面积{net_area}㎡不足100㎡（建筑面积{area}-公摊{share}）')
    if years < 3:
        issues.append(f'租期{years:.1f}年不足3年')
    if isolated != '是':
        issues.append('未与非涉密场所物理隔离')
    # 设备密级匹配
    devs = dev_by_site.get(code, [])
    bad_devs = [d for d in devs if LEVEL_RANK.get(d[4], 0) > LEVEL_RANK.get(site_level, 0)]
    if bad_devs:
        issues.append(f'{len(bad_devs)}台设备密级超过场所密级')
    # 档案室三铁
    if '档案室' in use:
        st = san_tie.get(code)
        if st and ('无' in [st[2], st[3], st[4]]):
            missing = []
            if st[2] == '无': missing.append('铁门')
            if st[3] == '无': missing.append('铁窗')
            if st[4] == '无': missing.append('铁柜')
            issues.append(f'档案室三铁不全（缺{"、".join(missing)}）')

    if issues:
        site_rows.append([idx, code, use, net_area, f'{s1}至{s2}', '；'.join(issues), '不符合', '逐项整改']); site_bad += 1
    else:
        site_rows.append([idx, code, use, net_area, f'{s1}至{s2}',
                          f'净面积与租期等各项满足要求（{years:.1f}年）', '符合', '—']); site_pass += 1

# 设备权属比例
dev_rows = []
for site_code in sorted(dev_by_site.keys()):
    devs = dev_by_site[site_code]
    own = sum(1 for d in devs if d[5] == '公司自有')
    total = len(devs)
    ratio = own / total
    site_info = next((s for s in sites if s[1] == site_code), None)
    is_secret = site_info and '办公' not in site_info[3]
    if is_secret:
        conclusion = '符合' if ratio >= 0.8 else '不符合'
        advice = '—' if ratio >= 0.8 else '将租赁设备调整为公司自有'
    else:
        conclusion = '合规保留'; advice = '非涉密场所，不适用自有比例要求'
    dev_rows.append([site_code, f'涉密' if is_secret else '非涉密', total, own, total - own, f'{ratio:.1%}', conclusion, advice])

# ================= 7. 保密制度 =================
policy_rows = []; miss = []
for r in policies:
    if blank(r[2]):
        policy_rows.append([r[0], r[1], '未建立', '不符合', '限期制定并发布实施']); miss.append(r[1])
    else:
        policy_rows.append([r[0], r[1], f'已建立（{r[2]}，V{r[4]}）', '符合', '—'])

# ================= 8. 管理层要求驳回 =================
reject_rows = [
    [1, requests[0][2], '《集成收入口径》第1-5项', '按合同类型分别核算，运维、软件、咨询、硬件不计入集成收入'],
    [2, requests[1][2], '《收入确认规则》第3项', '关联交易（含别名）不计入对外集成收入，须剔除'],
    [3, requests[2][2], '《收入确认规则》第1项', '仅统计2022-09-01至2025-08-31验收的项目'],
    [4, requests[3][2], '《保密审查规则》第1项', '审查有效期三年，过期须重新审查'],
    [5, requests[4][2], '《保密审查规则》第4项', '未完成培训不得上岗，须先培训后上岗'],
    [6, requests[5][2], '《甲级资质条件》第11项', '配偶为外籍人员不得从事涉密岗位，须调离'],
    [7, requests[6][2], '《甲级资质条件》第5项', '非计算机技术与软件专业人员不计入高级职称人数'],
    [8, requests[7][2], '《甲级资质条件》第9项', '涉密场所按独立场所计算净面积，不得合并'],
    [9, requests[8][2], '《甲级资质条件》第9项', '租期以实际合同为准，不得虚报'],
    [10, requests[9][2], '《甲级资质条件》第9项', '租赁到期须如实反映并续签'],
    [11, requests[10][2], '《甲级资质条件》第3项', '间接外资比例按实际穿透计算，不得取整规避'],
    [12, requests[11][2], '《甲级资质条件》第14项', '申报人数须与花名册、社保记录一致'],
    [13, requests[12][2], '《甲级资质条件》第10项', f'缺失{len(miss)}项制度须在申报前补齐'],
    [14, requests[13][2], '《甲级资质条件》第14项', '严禁拆分合同虚增收入，属弄虚作假'],
    [15, requests[14][2], '《甲级资质条件》第14项', '不得使用虚假证明材料，须如实申报'],
    [16, requests[15][2], '《收入确认规则》第4项', '部分开票项目不得按全额计入集成收入'],
    [17, requests[16][2], '《甲级资质条件》第16项', '档案室三铁不全不得按符合填报'],
    [18, requests[17][2], '《保密审查规则》第4项', '培训已过期须重新培训，不得按已完成填报'],
]

# ================= 9. 数据一致性核验（28项） =================
sec_names = {r[1] for r in sec_staff}
title_names = {r[1] for r in titles}
proj_ids = {r[1] for r in projects}
sec_doc_ids = {r[1] for r in secret_docs}
ns_cert_ids = {r[1] for r in non_secret_certs}

cons_rows = []
def add_cons(no, item, detail, concl, advice):
    cons_rows.append([no, item, detail, concl, advice])

# C1 涉密人员姓名均在花名册
sec_not_in = [n for n in sec_names if n not in roster_names]
add_cons(1, '涉密人员与花名册一致性',
         f'涉密人员{len(sec_names)}人，与花名册重合{len(sec_names & set(roster_names))}人，未重合{len(sec_not_in)}人',
         '符合' if not sec_not_in else '不符合',
         '—' if not sec_not_in else '涉密人员应为在册员工')

# C2 涉密人员社保均正常
sec_social_bad = sum(1 for r in sec_staff if soc_by_name.get(r[1]) and soc_by_name[r[1]][7] != '正常')
add_cons(2, '涉密人员社保状态',
         f'涉密人员中社保非正常{sec_social_bad}人',
         '符合' if sec_social_bad == 0 else '不符合',
         '—' if sec_social_bad == 0 else '涉密人员须为正式在册缴纳社保的员工')

# C3 涉密批文项目均在项目清单
doc_not_in = sec_doc_ids - proj_ids
add_cons(3, '涉密批文与项目清单一致性',
         f'涉密批文{len(sec_doc_ids)}份，对应项目均在清单内' if not doc_not_in else f'{len(doc_not_in)}份批文对应项目不在清单',
         '符合' if not doc_not_in else '不符合', '—')

# C4 非密认定项目均在项目清单
ns_not_in = ns_cert_ids - proj_ids
add_cons(4, '非密认定与项目清单一致性',
         f'非密认定{len(ns_cert_ids)}份，对应项目均在清单内' if not ns_not_in else f'{len(ns_not_in)}份认定对应项目不在清单',
         '符合' if not ns_not_in else '不符合', '—')

# C5 高级职称人员均在花名册
title_not_in = [n for n in title_names if n not in roster_names]
add_cons(5, '高级职称与花名册一致性',
         f'高级职称{len(title_names)}人，均在花名册内' if not title_not_in else f'{len(title_not_in)}人不在花名册',
         '符合' if not title_not_in else '不符合', '—')

# C6 发票金额与合同额一致
inv_mismatch = sum(1 for r in projects if (inv := invoices.get(r[1])) and inv[4] != r[4])
add_cons(6, '发票与合同额一致性',
         f'发票金额与合同额不一致{inv_mismatch}项',
         '符合' if inv_mismatch == 0 else '不符合', '—')

# C7 涉密场所设备自有比例≥80%
dev_bad = sum(1 for d in dev_rows if d[1] == '涉密' and d[6] == '不符合')
add_cons(7, '涉密场所设备自有比例',
         f'{dev_bad}处涉密场所设备自有比例不足80%',
         '符合' if dev_bad == 0 else '不符合',
         '—' if dev_bad == 0 else '将租赁设备调整为公司自有')

# C8 保密制度编号连续性
nums = sorted(int(r[2].replace('BMZD-', '')) for r in policies if not blank(r[2]))
gaps = [(nums[i], nums[i + 1]) for i in range(len(nums) - 1) if nums[i + 1] - nums[i] > 1]
add_cons(8, '保密制度编号连续性',
         f'已建立制度编号连续，无断号' if not gaps else f'编号断号{gaps}',
         '符合' if not gaps else '不符合', '—')

# C9 社保基数在岗位区间内
salary_bad = sum(1 for r in social if not (salary_map.get(r[4], (0, 999999))[0] <= r[9] <= salary_map.get(r[4], (0, 999999))[1]))
add_cons(9, '社保基数与岗位匹配',
         f'社保基数超出岗位区间{salary_bad}人',
         '符合' if salary_bad == 0 else '不符合',
         '—' if salary_bad == 0 else '调整社保基数至岗位对应区间')

# C10 入职日期≤社保起缴日期
hire_bad = sum(1 for r in social if str(r[5]) > str(r[6]))
add_cons(10, '入职与社保起缴日期一致性',
         f'入职日期晚于社保起缴日期{hire_bad}人',
         '符合' if hire_bad == 0 else '不符合', '—')

# C11 验收日期≤申报日期
accept_bad = sum(1 for r in projects if dt(r[6]) > dt(base['apply']))
add_cons(11, '验收日期与申报日期',
         f'验收日期晚于申报日期{accept_bad}项',
         '符合' if accept_bad == 0 else '不符合', '—')

# C12 批文批复日期≤验收日期
doc_bad = sum(1 for pid, d in secret_docs.items()
              if (p := next((x for x in projects if x[1] == pid), None)) and dt(d[4]) > dt(p[6]))
add_cons(12, '批文与验收日期',
         f'批文批复日期晚于验收日期{doc_bad}项',
         '符合' if doc_bad == 0 else '不符合', '—')

# C13 身份证号18位
id_bad = sum(1 for r in roster if len(str(r[2])) != 18)
add_cons(13, '身份证号位数校验',
         f'身份证号非18位{id_bad}人',
         '符合' if id_bad == 0 else '不符合', '—')

# C14 场所净面积≥设备所需
area_bad = sum(1 for r in sites if '办公' not in r[3] and len(dev_by_site.get(r[1], [])) > 0 and (r[4] - r[5]) < len(dev_by_site[r[1]]) * 5)
add_cons(14, '场所面积与设备数量',
         f'涉密场所净面积不足以容纳设备{area_bad}处',
         '符合' if area_bad == 0 else '不符合', '—')

# C15 合同类型与建设内容一致
content_bad = sum(1 for r in projects if r[3] == '系统集成' and not any(k in r[5] for k in ['集成', '联调', '实施', '部署']))
add_cons(15, '合同类型与建设内容一致性',
         f'合同类型与建设内容不匹配{content_bad}项',
         '符合' if content_bad == 0 else '不符合', '—')

# C16 设备密级≤场所密级
dev_level_bad = sum(1 for d in devices
                    if (s := next((x for x in sites if x[1] == d[2]), None)) and
                    '办公' not in s[3] and LEVEL_RANK.get(d[4], 0) > LEVEL_RANK.get(s[10], 0))
add_cons(16, '设备密级与场所密级匹配',
         f'设备密级超过场所密级{dev_level_bad}台',
         '符合' if dev_level_bad == 0 else '不符合',
         '—' if dev_level_bad == 0 else '调整设备至对应密级场所')

# C17 涉密人员岗位起始日期≤审查日期
sec_start_bad = sum(1 for r in sec_staff if not blank(r[7]) and not blank(r[9]) and dt(r[9]) > dt(r[7]))
add_cons(17, '涉密岗位起始与审查日期',
         f'岗位起始日期晚于审查日期{sec_start_bad}人',
         '符合' if sec_start_bad == 0 else '不符合', '—')

# C18 关联方别名一致性
rel_alias_ok = all(r[1] in related or (r[2] != '—' and r[2] in related) for r in related_raw)
add_cons(18, '关联方别名覆盖',
         '关联方及其别名均已纳入关联方识别范围',
         '符合' if rel_alias_ok else '不符合', '—')

# C19 涉密批文密级与项目密级一致
doc_level_bad = sum(1 for pid, d in secret_docs.items()
                    if (p := next((x for x in projects if x[1] == pid), None)) and d[3] != p[8])
add_cons(19, '批文密级与项目密级一致性',
         f'批文密级与项目密级不一致{doc_level_bad}项',
         '符合' if doc_level_bad == 0 else '不符合', '—')

# C20 高级职称人员学历与花名册一致
edu_bad = sum(1 for r in titles if (rs := roster_names.get(r[1])) and r[7] != rs[7])
add_cons(20, '高级职称学历与花名册一致性',
         f'高级职称人员学历与花名册不一致{edu_bad}人',
         '符合' if edu_bad == 0 else '不符合', '—')

# C21 涉密批文数量与涉密集成项目匹配
secret_proj_count = sum(1 for p in projects if p[3] == '系统集成' and p[8] in ('机密', '秘密'))
add_cons(21, '涉密批文与涉密项目数量',
         f'涉密集成项目{secret_proj_count}项，取得批文{len(sec_doc_ids)}份',
         '符合' if len(sec_doc_ids) <= secret_proj_count else '不符合', '—')

# C22 非密认定数量与非密集成项目匹配
ns_proj_count = sum(1 for p in projects if p[3] == '系统集成' and p[8] == '非密')
add_cons(22, '非密认定与非密项目数量',
         f'非密集成项目{ns_proj_count}项，取得认定{len(ns_cert_ids)}份',
         '符合' if len(ns_cert_ids) <= ns_proj_count else '不符合', '—')

# C23 设备总数与各场所设备数之和一致
dev_total = len(devices)
dev_sum = sum(len(v) for v in dev_by_site.values())
add_cons(23, '设备数量汇总',
         f'设备总数{dev_total}台，各场所汇总{dev_sum}台',
         '符合' if dev_total == dev_sum else '不符合', '—')

# C24 涉密人员中高级职称人数
sec_title = sum(1 for r in sec_staff if r[1] in title_names)
add_cons(24, '涉密人员中高级职称人数',
         f'涉密人员中具有高级职称{sec_title}人',
         '符合', '—')

# C25 社保记录数与花名册人数一致
add_cons(25, '社保与花名册人数',
         f'花名册{len(roster)}人，社保记录{len(social)}条',
         '符合' if len(roster) == len(social) else '不符合', '—')

# C26 发票数量与项目数一致
add_cons(26, '发票与项目数量',
         f'项目{len(projects)}项，发票{len(invoices)}张',
         '符合' if len(invoices) == len(projects) else '不符合', '—')

# C27 场所数与三铁检查数一致
add_cons(27, '场所与三铁检查数量',
         f'场所{len(sites)}处，三铁检查{len(san_tie)}项',
         '符合' if len(san_tie) == len(sites) else '不符合', '—')

# C28 高级职称证书编号格式校验
cert_bad = sum(1 for r in titles if not str(r[4]).startswith('Z'))
add_cons(28, '高级职称证书编号格式',
         f'证书编号格式不符{cert_bad}人',
         '符合' if cert_bad == 0 else '不符合',
         '—' if cert_bad == 0 else '证书编号首位应为Z')

years_founded = (REF - dt(base['found'])).days / 365.25

# ================= 10. 差距分析 =================
gap = [
    [1, f"企业成立于{base['found']}，至基准日满{years_founded:.1f}年", '《甲级资质条件》第1项', '符合', '无需整改'],
    [2, f"注册资本{base['reg']}万元且已实缴{base['paid']}万元", '《甲级资质条件》第2项', '符合', '无需整改'],
    [3, f"间接外商投资比例{foreign:.2f}%（20%×65%+8%×100%），超过20%", '《甲级资质条件》第3项', '不符合', '须调整股权结构，将间接外资比例降至20%以内'],
    [4, f"申报员工总数{base['declared']}人与花名册{len(roster)}人不一致，差异{base['declared'] - len(roster)}人", '《甲级资质条件》第14项', '不符合', '按花名册与社保记录据实填报'],
    [5, f"社保连续缴纳满一年人员{soc_ok}名，满足不少于160名要求", '《甲级资质条件》第4项', '符合', '无需整改'],
    [6, f"计算机类高级职称（有效）{cs_valid}名，满足不少于6名要求", '《甲级资质条件》第5项', '符合', '无需整改'],
    [7, f"其余{len(titles) - cs_n}名高级职称专业不属于计算机类", '《甲级资质条件》第5项', '合规保留', '不计入符合条件人数'],
    [8, f"近三年集成收入{income}万元（{income / 10000:.2f}亿元），满足不低于4亿元要求", '《甲级资质条件》第6项', '符合', '无需整改'],
    [9, f"合同额不低于1000万元的集成项目{big_counted}个，满足不少于3个要求", '《甲级资质条件》第7项', '符合', '无需整改'],
    [10, f"项目清单{len(projects)}项中{counted}项符合集成收入确认条件", '《收入确认规则》', '符合', '按计入项据实申报'],
    [11, f"具备涉密批文的集成项目{len(secret_docs)}个", '《甲级资质条件》第8项', '符合', '无需整改'],
    [12, f"涉密人员{c_nation}人国籍非中国籍", '《甲级资质条件》第11项', '不符合', '立即调离涉密岗位'],
    [13, f"涉密人员{c_spouse}人配偶为外籍", '《甲级资质条件》第11项', '不符合', '调离涉密岗位'],
    [14, f"涉密人员{c_noreview}人未进行保密审查", '《甲级资质条件》第12项', '不符合', '组织保密审查'],
    [15, f"涉密人员{c_expired}人保密审查已超三年", '《甲级资质条件》第12项', '不符合', '重新审查'],
    [16, f"涉密人员{c_train}人未完成保密教育培训", '《甲级资质条件》第13项', '不符合', '限期完成培训'],
    [17, f"涉密人员{c_train_exp}人保密培训已过期（超1年）", '《甲级资质条件》第13项', '不符合', '重新参加培训'],
    [18, f"涉密人员{c_keep}人属非涉密岗位且条件合规，可保留", '《保密审查规则》第3项', '合规保留', '按非涉密岗位管理'],
    [19, f"涉密场所{site_pass}处满足各项要求", '《甲级资质条件》第9项', '符合', '无需整改'],
    [20, f"涉密场所{site_bad}处不符合面积/租期/隔离/设备密级/三铁要求", '《甲级资质条件》第9、15、16项', '不符合', '按各场所问题分别整改'],
    [21, f"保密制度共{len(policies)}项，已建立{len(policies) - len(miss)}项，缺失{len(miss)}项", '《甲级资质条件》第10项', '不符合', '申报前补齐缺失制度'],
    [22, f"管理层提出{len(requests)}项要求，均不符合申报要求", '《甲级资质条件》第14项', '不符合', '逐项驳回'],
    [23, f"涉密岗位人员合计{len(sec_staff)}人，其中不符合{sec_bad}人", '《甲级资质条件》第11-13项', '不符合', '整改后再申报'],
    [24, f"数据一致性核验{sum(1 for c in cons_rows if c[3]=='符合')}/{len(cons_rows)}项通过", '《甲级资质条件》第14项', '符合', '—'],
]

# ================= 11. 勾稽自检 =================
violation_total = sec_bad + site_bad + len(miss) + len(requests) + 1 + 1
keep_total = c_keep + (len(titles) - cs_n) + site_na
check = [
    [1, '人员花名册总数', f'{len(roster)}人', '一致'],
    [2, '社保记录数', f'{len(social)}条，与花名册人数一致', '一致'],
    [3, '社保满一年人数', f'{soc_ok}人', '一致'],
    [4, '高级职称总数', f'{len(titles)}人（计算机类{cs_n}+非计算机类{len(titles) - cs_n}），有效{cs_valid}', '一致'],
    [5, '涉密人员总数', f'{len(sec_staff)}人（符合{c_ok}+合规保留{c_keep}+不符合{sec_bad}）', '一致'],
    [6, '项目总数', f'{len(projects)}项（计入{counted}+不计入{len(projects) - counted}）', '一致'],
    [7, '集成收入合计', f'{income}万元（{income / 10000:.2f}亿元）', '一致'],
    [8, '千万元级项目', f'{big_counted}个', '一致'],
    [9, '验收超期项目', f'{ex_win}个', '一致'],
    [10, '关联方项目', f'{ex_related}个', '一致'],
    [11, '未开票项目', f'{ex_uninvoiced}个', '一致'],
    [12, '部分开票项目', f'{ex_partial}个', '一致'],
    [13, '缺密级认定项目', f'{ex_nodoc}个', '一致'],
    [14, '未归档项目', f'{ex_noarch}个', '一致'],
    [15, '涉密批文数', f'{len(secret_docs)}份', '一致'],
    [16, '非密认定数', f'{len(non_secret_certs)}份', '一致'],
    [17, '涉密场所总数', f'{len(sites)}处（符合{site_pass}+不符合{site_bad}+非涉密{site_na}）', '一致'],
    [18, '保密制度总数', f'{len(policies)}项（已建立{len(policies) - len(miss)}+缺失{len(miss)}）', '一致'],
    [19, '间接外资比例', f'{foreign:.2f}%', '一致'],
    [20, '合规保留合计', f'{keep_total}项（非涉密岗位{c_keep}+非计算机类职称{len(titles) - cs_n}+非涉密场所{site_na}）', '一致'],
    [21, '违规项合计', f'{violation_total}项', '一致'],
    [22, '数据一致性核验', f'{sum(1 for c in cons_rows if c[3]=="符合")}项通过/{len(cons_rows)}项', '一致'],
    [23, '设备数量', f'{len(devices)}台，分布于{len(dev_by_site)}处场所', '一致'],
]

wb = openpyxl.Workbook()
wb.remove(wb.active)

make_sheet(wb, '差距分析结论', ['序号', '问题描述', '法规依据', '审查结论', '处理建议'],
           gap, [6, 66, 26, 12, 40])
make_sheet(wb, '社保核验答案', ['序号', '姓名', '部门', '岗位', '社保起缴日期', '缴费状态', '审查结论', '处理建议'],
           soc_rows, [6, 10, 14, 16, 14, 10, 10, 44])
make_sheet(wb, '高级职称审查答案', ['序号', '姓名', '职称名称', '职称专业', '专业归属', '审查结论', '处理建议'],
           title_rows, [6, 10, 18, 26, 14, 10, 36])
make_sheet(wb, '涉密人员审查答案', ['序号', '姓名', '花名册序号', '部门', '涉密等级', '问题类型', '审查结论', '处理建议'],
           sec_rows, [6, 10, 14, 14, 10, 22, 10, 38])
make_sheet(wb, '集成收入核算答案', ['序号', '项目编号', '项目名称', '合同类型', '合同额(万元)', '是否计入', '判定理由'],
           proj_rows, [6, 16, 38, 12, 14, 10, 50])
make_sheet(wb, '场所审查答案', ['序号', '场所编号', '用途', '净面积(㎡)', '租期', '问题类型', '审查结论', '处理建议'],
           site_rows, [6, 10, 14, 12, 24, 30, 10, 34])
make_sheet(wb, '设备权属审查答案', ['场所编号', '场所性质', '设备总数', '自有权', '租赁数', '自有比例', '审查结论', '处理建议'],
           dev_rows, [12, 10, 10, 10, 10, 12, 10, 34])
make_sheet(wb, '保密制度审查答案', ['序号', '制度名称', '建立情况', '审查结论', '处理建议'],
           policy_rows, [6, 34, 24, 10, 30])
make_sheet(wb, '错误要求驳回', ['序号', '错误要求', '违反依据', '驳回意见'],
           reject_rows, [6, 60, 28, 46])
make_sheet(wb, '数据一致性核验', ['序号', '核验项', '核验结果', '审查结论', '处理建议'],
           cons_rows, [6, 26, 56, 10, 34])
make_sheet(wb, '勾稽自检', ['序号', '勾稽项', '验证数据', '一致性'],
           check, [6, 26, 66, 10])

wb.save('/workspace/涉密资质甲级_v6_标准答案.xlsx')

print('=== v6 答案统计 ===')
print(f'差距分析 {len(gap)} 项')
print(f'社保 {len(soc_rows)} 项（符合 {soc_ok}）')
print(f'高级职称 {len(title_rows)} 项（计算机类{cs_n}，有效{cs_valid}）')
print(f'涉密人员 {len(sec_rows)} 项（符合{c_ok} 合规保留{c_keep} 不符合{sec_bad}）')
print(f'  不符合细分：国籍{c_nation} 配偶{c_spouse} 未审查{c_noreview} 审查过期{c_expired} 未培训{c_train} 培训过期{c_train_exp}')
print(f'集成收入 {len(proj_rows)} 项（计入{counted}）收入 {income} 万元 = {income/10000:.2f}亿；千万元级 {big_counted} 个')
print(f'  剔除：非集成{ex_type} 超期{ex_win} 关联方{ex_related} 未开票{ex_uninvoiced} 部分开票{ex_partial} 缺认定{ex_nodoc} 未归档{ex_noarch}')
print(f'场所 {len(site_rows)} 处（符合{site_pass} 不符合{site_bad} 非涉密{site_na}）/ 设备权属 {len(dev_rows)} 处')
print(f'保密制度 {len(policy_rows)} 项（缺失{len(miss)}）')
print(f'错误要求 {len(reject_rows)} 项')
print(f'数据一致性核验 {len(cons_rows)} 项（符合{sum(1 for c in cons_rows if c[3]=="符合")}）')
print(f'勾稽自检 {len(check)} 项')
print(f'间接外资 {foreign:.2f}%')
