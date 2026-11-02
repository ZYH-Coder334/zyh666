#!/usr/bin/env python3
"""民办学校年检 v4 - 拆分打包为6输入+6答案rar(zip格式)"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os, shutil, zipfile

THIN = Side(style="thin", color="CCCCCC")
BORDER = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
HDR_FILL = PatternFill("solid", fgColor="2B5B84")
HDR_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=10)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

def copy_sheet_to_new(src_ws, path, col_widths=None):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet(src_ws.title)
    for r_idx, row in enumerate(src_ws.iter_rows(values_only=True), 1):
        for c_idx, val in enumerate(row, 1):
            c = ws.cell(row=r_idx, column=c_idx, value=val)
            if r_idx == 1:
                c.font = HDR_FONT; c.fill = HDR_FILL; c.alignment = CENTER; c.border = BORDER
            else:
                c.font = Font(name="Arial", size=9); c.alignment = WRAP; c.border = BORDER
    if col_widths:
        for i, w in enumerate(col_widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
    wb.save(path)
    return ws.max_row - 1

src_q = openpyxl.load_workbook('/workspace/mb_v4_tmp.xlsx', data_only=True)
src_a = openpyxl.load_workbook('/workspace/mb_v4_ans_tmp.xlsx', data_only=True)

SPLIT = '/tmp/mb_v4_split'
shutil.rmtree(SPLIT, ignore_errors=True)
os.makedirs(f'{SPLIT}/input', exist_ok=True)
os.makedirs(f'{SPLIT}/ans', exist_ok=True)

# ========== 输入附件 ==========
print("=== 拆分输入附件 ===")
# 01 学校基本信息与许可证（含3个跨表查询表）
copy_sheet_to_new(src_q['学校基本信息'], f'{SPLIT}/input/学校基本信息.xlsx', [6,18,18,22,14,12,12,18,10,10,12,12,12,12,10,10,30])
copy_sheet_to_new(src_q['办学许可证信息'], f'{SPLIT}/input/办学许可证信息.xlsx', [6,18,14,12,12,14,24,14,24])
copy_sheet_to_new(src_q['办学层次与学科要求'], f'{SPLIT}/input/办学层次与学科要求.xlsx', [6,12,18,18,40])
copy_sheet_to_new(src_q['教师资格证管理规定'], f'{SPLIT}/input/教师资格证管理规定.xlsx', [6,18,16,16,24])
copy_sheet_to_new(src_q['设施设备使用年限标准'], f'{SPLIT}/input/设施设备使用年限标准.xlsx', [6,12,14,20,18])
copy_sheet_to_new(src_q['学籍状态分类标准'], f'{SPLIT}/input/学籍状态分类标准.xlsx', [6,14,24,28])
# 02 师资队伍
copy_sheet_to_new(src_q['教师花名册'], f'{SPLIT}/input/教师花名册.xlsx', [6,10,8,16,16,10,12,8,12,6,10,28])
# 03 财务收费
copy_sheet_to_new(src_q['收费标准清单'], f'{SPLIT}/input/收费标准清单.xlsx', [6,18,16,20])
copy_sheet_to_new(src_q['收费批复标准'], f'{SPLIT}/input/收费批复标准.xlsx', [6,18,16,20])
copy_sheet_to_new(src_q['财务报表'], f'{SPLIT}/input/财务报表.xlsx', [6,18,12,12,24,18,10])
# 04 安全设施
copy_sheet_to_new(src_q['安全检查记录'], f'{SPLIT}/input/安全检查记录.xlsx', [6,10,14,12,12,10,14,36])
copy_sheet_to_new(src_q['设施设备清单'], f'{SPLIT}/input/设施设备清单.xlsx', [6,10,18,12,16,14,14,10,28])
# 05 招生学籍
copy_sheet_to_new(src_q['招生计划与实际'], f'{SPLIT}/input/招生计划与实际.xlsx', [6,10,8,8,14,18,30])
copy_sheet_to_new(src_q['招生计划批复'], f'{SPLIT}/input/招生计划批复.xlsx', [6,10,8,14])
copy_sheet_to_new(src_q['学籍管理记录'], f'{SPLIT}/input/学籍管理记录.xlsx', [6,10,12,8,10,40])
# 06 管理制度与管理层
copy_sheet_to_new(src_q['管理制度清单'], f'{SPLIT}/input/管理制度清单.xlsx', [6,20,14,10])
copy_sheet_to_new(src_q['管理层要求'], f'{SPLIT}/input/管理层要求.xlsx', [6,14,50,12])
print("输入附件拆分完成")

# ========== 标准答案 ==========
print("=== 拆分标准答案 ===")
copy_sheet_to_new(src_a['合规诊断结论'], f'{SPLIT}/ans/合规诊断结论.xlsx', [6,60,10,40,24])
copy_sheet_to_new(src_a['师资队伍审查答案'], f'{SPLIT}/ans/师资队伍审查答案.xlsx', [6,10,28,10,36,16])
copy_sheet_to_new(src_a['收费审查答案'], f'{SPLIT}/ans/收费审查答案.xlsx', [6,18,24,10,18,24])
copy_sheet_to_new(src_a['财务审查答案'], f'{SPLIT}/ans/财务审查答案.xlsx', [6,18,18,10,18,24])
copy_sheet_to_new(src_a['安全检查审查答案'], f'{SPLIT}/ans/安全检查审查答案.xlsx', [6,10,22,10,24,28])
copy_sheet_to_new(src_a['设施设备审查答案'], f'{SPLIT}/ans/设施设备审查答案.xlsx', [6,10,30,10,18,28])
copy_sheet_to_new(src_a['招生审查答案'], f'{SPLIT}/ans/招生审查答案.xlsx', [6,10,32,10,30,32])
copy_sheet_to_new(src_a['学籍审查答案'], f'{SPLIT}/ans/学籍审查答案.xlsx', [6,10,18,10,28,28])
copy_sheet_to_new(src_a['错误要求驳回'], f'{SPLIT}/ans/错误要求驳回.xlsx', [6,36,36,30])
copy_sheet_to_new(src_a['勾稽自检'], f'{SPLIT}/ans/勾稽自检.xlsx', [6,20,60,10])
print("标准答案拆分完成")

# ========== 打包 ==========
OUT = '/workspace/民办学校年检_附件'
shutil.rmtree(f'{OUT}/输入附件', ignore_errors=True)
shutil.rmtree(f'{OUT}/标准答案', ignore_errors=True)
os.makedirs(f'{OUT}/输入附件', exist_ok=True)
os.makedirs(f'{OUT}/标准答案', exist_ok=True)

def make_rar(files, rar_path):
    with zipfile.ZipFile(rar_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for f in files:
            zf.write(f, arcname=os.path.basename(f))
    return os.path.getsize(rar_path)

print("\n=== 打包输入附件 ===")
make_rar([f'{SPLIT}/input/学校基本信息.xlsx', f'{SPLIT}/input/办学许可证信息.xlsx', f'{SPLIT}/input/办学层次与学科要求.xlsx', f'{SPLIT}/input/教师资格证管理规定.xlsx', f'{SPLIT}/input/设施设备使用年限标准.xlsx', f'{SPLIT}/input/学籍状态分类标准.xlsx'], f'{OUT}/输入附件/01学校基本信息与许可证.rar')
make_rar([f'{SPLIT}/input/教师花名册.xlsx'], f'{OUT}/输入附件/02师资队伍与教师资格.rar')
make_rar([f'{SPLIT}/input/收费标准清单.xlsx', f'{SPLIT}/input/收费批复标准.xlsx', f'{SPLIT}/input/财务报表.xlsx'], f'{OUT}/输入附件/03财务管理与收费标准.rar')
make_rar([f'{SPLIT}/input/安全检查记录.xlsx', f'{SPLIT}/input/设施设备清单.xlsx'], f'{OUT}/输入附件/04安全卫生与设施设备.rar')
make_rar([f'{SPLIT}/input/招生计划与实际.xlsx', f'{SPLIT}/input/招生计划批复.xlsx', f'{SPLIT}/input/学籍管理记录.xlsx'], f'{OUT}/输入附件/05招生与学籍管理.rar')
make_rar([f'{SPLIT}/input/管理制度清单.xlsx', f'{SPLIT}/input/管理层要求.xlsx'], f'{OUT}/输入附件/06管理制度与管理层要求.rar')
print("输入附件打包完成")

print("\n=== 打包标准答案 ===")
make_rar([f'{SPLIT}/ans/合规诊断结论.xlsx'], f'{OUT}/标准答案/01合规诊断报告.rar')
make_rar([f'{SPLIT}/ans/师资队伍审查答案.xlsx'], f'{OUT}/标准答案/02师资队伍审查答案.rar')
make_rar([f'{SPLIT}/ans/收费审查答案.xlsx', f'{SPLIT}/ans/财务审查答案.xlsx'], f'{OUT}/标准答案/03财务收费审查答案.rar')
make_rar([f'{SPLIT}/ans/安全检查审查答案.xlsx', f'{SPLIT}/ans/设施设备审查答案.xlsx'], f'{OUT}/标准答案/04安全设施审查答案.rar')
make_rar([f'{SPLIT}/ans/招生审查答案.xlsx', f'{SPLIT}/ans/学籍审查答案.xlsx'], f'{OUT}/标准答案/05招生学籍审查答案.rar')
make_rar([f'{SPLIT}/ans/错误要求驳回.xlsx', f'{SPLIT}/ans/勾稽自检.xlsx'], f'{OUT}/标准答案/06整改方案与勾稽自检.rar')
print("标准答案打包完成")

print(f"\n输入文件: {len(os.listdir(f'{SPLIT}/input'))}个")
print(f"答案文件: {len(os.listdir(f'{SPLIT}/ans'))}个")
