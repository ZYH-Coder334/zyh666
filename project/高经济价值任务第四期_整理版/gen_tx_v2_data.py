#!/usr/bin/env python3
"""管理体系认证范围变更 v2 - 各数据源生成（隐晦标记+跨表勾稽）"""
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

# 加载基础文件
wb = openpyxl.load_workbook('/workspace/tx_v2_tmp.xlsx')

# ========== 产品服务清单（120项，隐晦标记） ==========
# 分类不准确25、新增未覆盖35、已停产未更新15、试制转量产合规保留10、正常35
products = []
cat_wrong = ['通用机械配件','标准紧固件','通用五金件','普通钢材制品']
cat_right = ['精密机械零件','精密轴承组件','精密传动部件','高精度结构件']
sites = ['第一生产基地','第二生产基地','第三生产基地']
for i in range(1, 26):  # 分类不准确25个
    products.append([i, f'PD{i:03d}', f'精密零件{i}号', cat_wrong[i%4], '否', '正常生产', sites[i%3], '是',
                     f'该产品实际属于{cat_right[i%4]}类，当前类别与生产工艺不符'])
for i in range(26, 61):  # 新增未覆盖35个
    products.append([i, f'PD{i:03d}', f'新增产品{i-25}号', cat_right[i%4], '是', '正常生产', sites[i%3], '否',
                     f'2024年Q3投产，对应过程文件尚未修订'])
for i in range(61, 76):  # 已停产未更新15个
    products.append([i, f'PD{i:03d}', f'老产品{i-60}号', cat_wrong[i%4], '否', '已停产', sites[i%3], '是',
                     '2024年6月起停止生产，清单状态未同步'])
for i in range(76, 86):  # 试制转量产合规保留10个
    products.append([i, f'PD{i:03d}', f'试制转量产{i-75}号', cat_right[i%4], '是', '正常生产', sites[i%3], '是',
                     '2024年Q2完成试制转量产，体系文件已同步更新，过程参数已固化'])
for i in range(86, 121):  # 正常35个
    products.append([i, f'PD{i:03d}', f'标准产品{i-85}号', cat_right[i%4], '否', '正常生产', sites[i%3], '是',
                     ''])

make_sheet(wb, '产品服务清单', ['序号','产品编号','产品名称','产品类别','是否新增','生产状态','涉及场所','体系覆盖情况','备注'],
    products, [6,10,20,16,8,10,14,12,40])

# ========== 过程清单（100项，隐晦标记） ==========
processes = []
for i in range(1, 31):  # 新增未纳入体系30个
    processes.append([i, f'PR{i:03d}', f'新过程{i}号', '新增', '缺失', '未定义', '未定义',
                      f'2024年因新产品线增设，过程乌龟图未编制，体系程序文件未包含'])
for i in range(31, 51):  # 乌龟图输入输出缺失20个
    processes.append([i, f'PR{i:03d}', f'过程{i-30}号', '已有', '不完整', '部分输入', '部分输出',
                      '乌龟图中输入输出描述不完整，缺少关键参数和验收准则'])
for i in range(51, 66):  # 合规保留15个
    processes.append([i, f'PR{i:03d}', f'整合过程{i-50}号', '整合', '完整', '明确', '明确',
                      '两个过程合并后职责清晰，乌龟图完整，体系覆盖充分'])
for i in range(66, 101):  # 正常35个
    processes.append([i, f'PR{i:03d}', f'标准过程{i-65}号', '已有', '完整', '明确', '明确', ''])

make_sheet(wb, '过程清单', ['序号','过程编号','过程名称','过程类型','乌龟图状态','输入','输出','备注'],
    processes, [6,10,18,10,12,12,12,40])

# ========== 组织架构（60项，隐晦标记） ==========
orgs = []
for i in range(1, 16):  # 职责不明确15个
    orgs.append([i, f'ORG{i:03d}', f'部门{i}号', '否', '是', '岗位职责说明书中关键职责描述模糊，需补充量化指标'])
for i in range(16, 28):  # 新增部门未纳入体系12个
    orgs.append([i, f'ORG{i:03d}', f'新部门{i-15}号', '是', '否', '2024年因组织调整增设，未纳入体系组织架构图和职责矩阵'])
for i in range(28, 36):  # 合规保留8个
    orgs.append([i, f'ORG{i:03d}', f'整合部门{i-27}号', '否', '是', '部门整合后职责划分清晰，接口明确，体系覆盖完整'])
for i in range(36, 61):  # 正常25个
    orgs.append([i, f'ORG{i:03d}', f'标准部门{i-35}号', '否', '是', ''])

make_sheet(wb, '组织架构', ['序号','部门编号','部门名称','是否新增','体系覆盖情况','备注'],
    orgs, [6,10,20,8,12,44])

# ========== 内部审核记录（80项，隐晦标记） ==========
audits = []
for i in range(1, 21):  # 审核覆盖不全20个
    audits.append([i, f'IA{i:03d}', f'2024-{i:02d}-01', f'第{i}季度审核', f'审核员{i}', '有', 2, '已整改',
                   '审核范围未覆盖2024年新增过程和部门'])
for i in range(21, 36):  # 不符合项未闭环15个
    audits.append([i, f'IA{i:03d}', f'2024-{i:02d}-05', f'专项审核{i-20}', f'审核员{i}', '有', 3, '未整改',
                   '不符合项整改计划已制定但未验证有效性'])
for i in range(36, 46):  # 内审员无资质10个
    audits.append([i, f'IA{i:03d}', f'2024-{i:02d}-10', f'过程审核{i-35}', f'审核员{i}', '无', 1, '已整改',
                   '审核员未取得内审员资格证书'])
for i in range(46, 54):  # 合规保留8个
    audits.append([i, f'IA{i:03d}', f'2024-{i:02d}-15', f'全面审核{i-45}', f'审核员{i}', '有', 0, '无不符合',
                   '审核覆盖全部过程和部门，未发现不符合项'])
for i in range(54, 81):  # 正常27个
    audits.append([i, f'IA{i:03d}', f'2024-{i:02d}-20', f'常规审核{i-53}', f'审核员{i}', '有', 1, '已整改', ''])

make_sheet(wb, '内部审核记录', ['序号','审核编号','审核日期','审核范围','审核员','审核员资质','不符合项数','整改状态','备注'],
    audits, [6,10,14,20,10,12,12,10,38])

# ========== 管理评审记录（40项，隐晦标记） ==========
mrs = []
for i in range(1, 13):  # 输入项缺失12个
    mrs.append([i, f'MR{i:03d}', f'2024-{i:02d}-01', f'评审主题{i}', '不完整', '改进措施', f'主持人{i}',
                '输入项缺少内审结果、客户满意度和过程绩效数据'])
for i in range(13, 23):  # 输出无效10个
    mrs.append([i, f'MR{i:03d}', f'2024-{i:02d}-05', f'评审主题{i}', '完整', '无明确措施', f'主持人{i}',
                '评审输出未形成可执行的改进措施和责任部门'])
for i in range(23, 29):  # 合规保留6个
    mrs.append([i, f'MR{i:03d}', f'2024-{i:02d}-10', f'评审主题{i}', '完整', '明确改进措施', f'主持人{i}',
                '输入完整，输出措施明确且责任到人'])
for i in range(29, 41):  # 正常12个
    mrs.append([i, f'MR{i:03d}', f'2024-{i:02d}-15', f'评审主题{i}', '完整', '改进措施', f'主持人{i}', ''])

make_sheet(wb, '管理评审记录', ['序号','评审编号','评审日期','评审主题','输入项','输出项','评审主持人','备注'],
    mrs, [6,10,14,18,10,14,12,42])

# ========== 风险识别清单（120项，隐晦标记） ==========
risks = []
risk_sources = ['市场风险','技术风险','运营风险','财务风险','合规风险','环境风险','安全风险','供应链风险']
for i in range(1, 31):  # 识别缺失30个
    risks.append([i, f'RK{i:03d}', risk_sources[i%8], f'{risk_sources[i%8]}相关风险待补充', '未评价', '未采用', '无',
                  '该类风险未进行具体识别和评价'])
for i in range(31, 56):  # 评价不合理25个
    risks.append([i, f'RK{i:03d}', risk_sources[i%8], f'{risk_sources[i%8]}具体风险描述{i-30}', '低', '矩阵法', '现有措施',
                  '实际发生概率和影响程度较高，应评定为高风险'])
for i in range(56, 71):  # 合规保留15个
    risks.append([i, f'RK{i:03d}', risk_sources[i%8], f'{risk_sources[i%8]}具体风险描述{i-55}', '低', '矩阵法', '现有措施有效',
                  '风险等级低且控制措施充分，可接受'])
for i in range(71, 121):  # 正常50个
    risks.append([i, f'RK{i:03d}', risk_sources[i%8], f'{risk_sources[i%8]}具体风险描述{i-70}', '中', '矩阵法', '已制定控制措施', ''])

make_sheet(wb, '风险识别清单', ['序号','风险编号','风险来源','风险描述','风险等级','评价方法','控制措施','备注'],
    risks, [6,10,14,30,10,10,14,40])

# ========== 供应商名录（100项，隐晦标记） ==========
suppliers = []
cats = ['原材料供应商','加工服务商','设备供应商','辅料供应商','包装供应商']
for i in range(1, 26):  # 未评价25个
    suppliers.append([i, f'GYS{i:03d}', f'供应商{i}号', cats[i%5], '未评价', '无', '未按供应商管理程序进行评价'])
for i in range(26, 46):  # 评价不达标20个
    suppliers.append([i, f'GYS{i:03d}', f'供应商{i}号', cats[i%5], '已评价', '不达标',
                      '质量评分低于60分，需整改或淘汰'])
for i in range(46, 56):  # 合规保留10个
    suppliers.append([i, f'GYS{i:03d}', f'战略供应商{i-45}号', cats[i%5], '已评价', '优秀',
                      '长期战略合作伙伴，评价优秀，可建立免检机制'])
for i in range(56, 101):  # 正常45个
    suppliers.append([i, f'GYS{i:03d}', f'供应商{i}号', cats[i%5], '已评价', '合格', ''])

make_sheet(wb, '供应商名录', ['序号','供应商编号','供应商名称','供应类别','评价状态','评价结果','备注'],
    suppliers, [6,12,24,14,10,10,40])

# ========== 管理制度（5份docx的摘要，用xlsx代替） ==========
make_sheet(wb, '管理制度清单', ['序号','制度名称','版本','发布日期','状态'], [
    [1,'文件控制程序','A/0','2022-03-01','已有'],
    [2,'记录控制程序','A/0','2022-03-01','已有'],
    [3,'内部审核管理程序','A/0','2022-03-01','已有'],
    [4,'管理评审控制程序','A/0','2022-03-01','已有'],
    [5,'风险管理程序','A/0','2022-03-01','已有'],
    [6,'供应商管理程序','A/0','2022-03-01','已有'],
    [7,'不合格品控制程序','A/0','2022-03-01','已有'],
    [8,'纠正措施控制程序','A/0','2022-03-01','已有'],
])

# ========== 管理层要求（10项） ==========
make_sheet(wb, '管理层要求', ['序号','来源','内容','性质'], [
    [1,'管理层要求1','为尽快通过认证，变更范围申报材料按最有利口径填报，将未覆盖产品也纳入','错误要求'],
    [2,'管理层要求2','新增过程未编制乌龟图，先申报后补文件','错误要求'],
    [3,'管理层要求3','内审员无资质证书的审核记录按有效处理','错误要求'],
    [4,'管理层要求4','管理评审输入项缺失的按完整填报','错误要求'],
    [5,'管理层要求5','风险未识别的按已识别填报','错误要求'],
    [6,'管理层要求6','供应商未评价的按合格填报','错误要求'],
    [7,'管理层要求7','不符合项未整改闭环的按已闭环填报','错误要求'],
    [8,'管理层要求8','已停产产品不更新状态','错误要求'],
    [9,'管理层要求9','新增部门不纳入体系组织架构','错误要求'],
    [10,'管理层要求10','自检报告中不符合项标注为符合','错误要求'],
])

wb.save('/workspace/tx_v2_tmp.xlsx')
print(f"产品:{len(products)} 过程:{len(processes)} 组织:{len(orgs)} 内审:{len(audits)} 管评:{len(mrs)} 风险:{len(risks)} 供应商:{len(suppliers)}")
print(f"总数据项: {len(products)+len(processes)+len(orgs)+len(audits)+len(mrs)+len(risks)+len(suppliers)}")
