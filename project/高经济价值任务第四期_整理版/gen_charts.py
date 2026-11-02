#!/usr/bin/env python3
"""生成图表"""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib import rcParams

rcParams['font.sans-serif'] = ['DejaVu Sans', 'SimHei', 'WenQuanYi Micro Hei']
rcParams['axes.unicode_minus'] = False

IN = "/workspace/数字化项目止损_附件/输入附件"
OUT = "/workspace/数字化项目止损_附件/标准答案"

# 输入附件图表
# 1. 预算执行情况
fig, ax = plt.subplots(figsize=(10,5))
items = ['ERP','MES','E-commerce','Integration','Hardware','Training','O&M','Reserve']
budget = [1200,1000,500,300,400,200,300,100]
actual = [1500,1200,450,280,500,180,400,200]
x = range(len(items))
w = 0.35
ax.bar([i-w/2 for i in x], budget, w, label='Budget', color='#4472C4')
ax.bar([i+w/2 for i in x], actual, w, label='Actual', color='#ED7D31')
ax.set_xticks(list(x)); ax.set_xticklabels(items, fontsize=8)
ax.set_ylabel('Amount (10k)'); ax.set_title('Budget vs Actual Execution')
ax.legend(); plt.tight_layout()
plt.savefig(f"{IN}/预算执行对比.png", dpi=100)
plt.close()
print("  预算执行对比.png")

# 2. 供应商资质分布
fig, ax = plt.subplots(figsize=(6,6))
ax.pie([5,10], labels=['Unqualified(5)','Qualified(10)'], autopct='%1.0f%%',
       colors=['#FF6B6B','#4ECDC4'], startangle=90)
ax.set_title('Vendor Qualification Distribution')
plt.tight_layout()
plt.savefig(f"{IN}/供应商资质分布.png", dpi=100)
plt.close()
print("  供应商资质分布.png")

# 3. 需求状态分布
fig, ax = plt.subplots(figsize=(8,5))
cats = ['Unconfirmed','Missing','Scope Creep','Normal']
counts = [12,10,8,30]
ax.bar(cats, counts, color=['#FF6B6B','#FFA500','#FFD93D','#6BCB77'])
ax.set_ylabel('Count'); ax.set_title('Requirement Status Distribution')
for i,v in enumerate(counts): ax.text(i, v+0.5, str(v), ha='center')
plt.tight_layout()
plt.savefig(f"{IN}/需求状态分布.png", dpi=100)
plt.close()
print("  需求状态分布.png")

# 4. 交付物状态分布
fig, ax = plt.subplots(figsize=(6,6))
ax.pie([10,8,5,17], labels=['Unqualified(10)','Undelivered(8)','Fake Acceptance(5)','Normal(17)'],
       autopct='%1.0f%%', colors=['#FF6B6B','#FFA500','#FFD93D','#6BCB77'], startangle=90)
ax.set_title('Deliverable Status Distribution')
plt.tight_layout()
plt.savefig(f"{IN}/交付物状态分布.png", dpi=100)
plt.close()
print("  交付物状态分布.png")

# 5. 合同采购方式分布
fig, ax = plt.subplots(figsize=(6,6))
ax.pie([8,4,3,5], labels=['No-Bid(8)','Missing Terms(4)','Duplicate(3)','Compliant(5)'],
       autopct='%1.0f%%', colors=['#FF6B6B','#FFA500','#FFD93D','#6BCB77'], startangle=90)
ax.set_title('Contract Procurement Distribution')
plt.tight_layout()
plt.savefig(f"{IN}/合同采购分布.png", dpi=100)
plt.close()
print("  合同采购分布.png")

# 标准答案图表
# 6. 违规分布
fig, ax = plt.subplots(figsize=(10,5))
cats = ['Budget','Fund','Vendor','Contract','Requirement','Change','Deliverable','Acceptance']
counts = [5,15,5,15,30,13,23,14]
ax.barh(cats, counts, color='#FF6B6B')
ax.set_xlabel('Violation Count'); ax.set_title('Violation Distribution by Category')
for i,v in enumerate(counts): ax.text(v+0.3, i, str(v), va='center')
plt.tight_layout()
plt.savefig(f"{OUT}/违规分布.png", dpi=100)
plt.close()
print("  违规分布.png")

# 7. 整改前后对比
fig, ax = plt.subplots(figsize=(8,5))
stages = ['Phase 1\n1-15d','Phase 2\n16-30d','Phase 3\n31-60d','Phase 4\n61-90d']
before = [0,0,0,120]
after = [0,35,60,120]
ax.bar(stages, before, label='Before(violations remain)', color='#FF6B6B', alpha=0.7)
ax.bar(stages, after, label='After(cumulative rectified)', color='#6BCB77', alpha=0.7)
ax.set_ylabel('Violation Count'); ax.set_title('Violation Rectification Progress')
ax.legend(); plt.tight_layout()
plt.savefig(f"{OUT}/整改前后对比.png", dpi=100)
plt.close()
print("  整改前后对比.png")

print("\n=== 图表生成完成 ===")
