"""
跑一遍逻辑校验，看看企业填的36项指标有没有互相打架的情况。

读取 企业验证数据模板.xlsx 的"指标填写表"K列，逐条检查跨指标约束。
通过后就可以跑 03 做正式评价了。

用法：
  python 02_数据逻辑校验.py
"""

import pandas as pd
import os
from tkinter import Tk, messagebox

DIR = os.path.dirname(os.path.abspath(__file__))
XLSX = os.path.join(DIR, "企业验证数据模板.xlsx")

if not os.path.exists(XLSX):
    root = Tk()
    root.withdraw()
    messagebox.showwarning(
        "缺少模板文件",
        f"未找到企业验证数据模板.xlsx！\n\n"
        f"请先运行 01_生成空白填报表单.py 生成模板文件。\n\n"
        f"查找位置：\n{XLSX}"
    )
    root.destroy()
    exit(1)

df = pd.read_excel(XLSX, sheet_name="指标填写表")
values = df["填写值"].tolist()

# 先看看填完了没
filled_count = sum(1 for v in values if v != "" and not pd.isna(v))
if filled_count < 36:
    root = Tk()
    root.withdraw()
    messagebox.showwarning(
        "数据未填写完整",
        f"只填写了 {filled_count}/36 项指标数据。\n\n"
        f"请打开 企业验证数据模板.xlsx\n"
        f"在【指标填写表】的 K列（填写值）\n"
        f"逐项填入贵企业的36项指标数据后，再运行本脚本。"
    )
    root.destroy()
    exit(1)

values = [float(v) for v in values]
names = df["指标名称"].tolist()
d = dict(zip(names, values))

print("=== 企业数据逻辑约束检查 ===\n")
print(f"  数据来源: 企业验证数据模板.xlsx → 指标填写表(K列)")
print(f"  已填写指标: {filled_count}/36\n")

violations = []
warnings = []

# 约束1: 设备联网率不应该比数据自动采集率的70%还低
v1 = d["设备联网率(%)"]
v7 = d["数据自动采集覆盖率(%)"]
threshold_c1 = v7 * 0.7
if v1 < threshold_c1:
    violations.append(
        f"C1  设备联网率({v1}%) < 数据自动采集率({v7}%) × 0.7 = {threshold_c1:.1f}%\n"
        f"    → 联网率一般不低于采集率的七成，核实一下是否低估了设备联网情况"
    )

# 约束2: MES部署率不该超过ERP覆盖率
v4 = d["ERP系统覆盖率(%)"]
v5 = d["MES系统部署率(%)"]
if v4 < v5:
    violations.append(
        f"C2  ERP覆盖率({v4}%) < MES部署率({v5}%)\n"
        f"    → MES部署范围一般不会超过ERP，核实两套系统的实际部署情况"
    )

# 约束3: 数据分析应用率不能超过数据采集覆盖
v8 = d["数据分析应用率(%)"]
if v8 > v7:
    violations.append(
        f"C3  数据分析应用率({v8}%) > 数据自动采集覆盖率({v7}%)\n"
        f"    → 分析覆盖面不可能跑在采集前头，检查数据来源"
    )

# 约束4: 智能检测覆盖率一般不会超过自动化设备占比
v10 = d["自动化设备占比(%)"]
v11 = d["智能检测与质控覆盖率(%)"]
if v10 < v11:
    violations.append(
        f"C4  自动化设备占比({v10}%) < 智能检测覆盖率({v11}%)\n"
        f"    → 智能检测通常部署在自动化设备上，覆盖率不该高过自动化占比"
    )

# 约束5: 库存数字化一般领先物流可视化（提示级别）
v14 = d["库存管理数字化率(%)"]
v15 = d["物流可视化追踪率(%)"]
if v14 < v15 * 0.8:
    warnings.append(
        f"C5  库存管理数字化率({v14}%) < 物流追踪率({v15}%) × 0.8\n"
        f"    → 库存数字化通常走在物流可视化前头，数值可能不太对"
    )

# 约束7: 效率提升幅度不应该远超自动化水平
v28 = d["生产效率提升率(%)"]
threshold_c7 = v10 * 1.2
if v28 > threshold_c7:
    violations.append(
        f"C7  生产效率提升率({v28}%) > 自动化占比({v10}%) × 1.2 = {threshold_c7:.1f}%\n"
        f"    → 效率提升不太应该大幅超过自动化水平，查查统计口径"
    )

# 约束8: 数字化营销收入占比和CRM成熟度应该大致匹配（提示级别）
v16 = d["CRM系统应用率(%)"]
v17 = d["数字化营销收入占比(%)"]
if v17 > v16 * 1.5:
    warnings.append(
        f"C8  数字化营销收入占比({v17}%) > CRM应用率({v16}%) × 1.5\n"
        f"    → 数字化营销收入和CRM成熟度大致应该对得上"
    )

# 约束9: 订单按时交付率范围检查
v30 = d["订单按时交付率(%)"]
if v30 < 70 or v30 > 100:
    violations.append(
        f"C9  订单按时交付率({v30}%) 不在 [70, 100] 范围内\n"
        f"    → 模型要求交付率在70%~100%之间，超出的值会被截断"
    )

# ── 输出违规和提示 ──

if violations:
    print(f"⚠ 发现 {len(violations)} 处逻辑约束违规:\n")
    for i, v in enumerate(violations, 1):
        print(f"  [{i}] {v}\n")
else:
    print("✅ 全部逻辑约束通过！")

if warnings:
    print(f"💡 {len(warnings)} 条提示信息:\n")
    for i, w in enumerate(warnings, 1):
        print(f"  [{i}] {w}\n")

# ── 门槛指标速览 ──

print(f"{'─' * 60}")
print("5项关键门槛指标值:")
threshold_items = {
    "设备联网率(%)": d["设备联网率(%)"],
    "ERP系统覆盖率(%)": d["ERP系统覆盖率(%)"],
    "MES系统部署率(%)": d["MES系统部署率(%)"],
    "数据分析应用率(%)": d["数据分析应用率(%)"],
    "数字化专业人才占比(%)": d["数字化专业人才占比(%)"],
}
for name, val in threshold_items.items():
    print(f"  {name:<30} = {val:>6.1f}")

print(f"\n共检查 9 条逻辑约束（6条硬约束 + 3条补充提示）。")
print(f"有违规就核实数据来源后修正，再重新跑一遍。")
print(f"全部通过后就可以跑: python 03_运行企业验证.py --from-template")

# 弹窗汇总校验结果
root = Tk()
root.withdraw()
if violations:
    msg = f"发现 {len(violations)} 处逻辑违规：\n\n"
    for v in violations[:6]:
        msg += f"⚠ {v.split(chr(10))[0]}\n"
    if len(violations) > 6:
        msg += f"\n... 共 {len(violations)} 处，详见终端输出。\n"
    if warnings:
        msg += f"\n另有 {len(warnings)} 条提示信息，详见终端输出。\n"
    msg += f"\n请核实数据来源后修正，再重新校验。"
    messagebox.showwarning("校验结果：存在违规", msg)
else:
    wmsg = ""
    if warnings:
        wmsg = f"\n\n另有 {len(warnings)} 条提示信息，详见终端输出。"
    messagebox.showinfo(
        "校验结果：全部通过",
        f"全部逻辑约束通过！{wmsg}\n\n"
        f"可以继续运行 03_运行企业验证.py 进行评价。"
    )
root.destroy()
