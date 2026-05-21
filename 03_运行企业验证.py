"""
跑完整的企业数据验证流程，输出评价结果到 Excel。

把企业填好的数据喂给论文评价模型，跑一遍完整的评价，
输出转型阶段、综合得分、维度得分、方案对比、敏感性分析和优化建议。

用法：
  双击运行本脚本，按弹窗提示操作即可。
  数据来源：同目录下的【企业验证数据模板.xlsx】的【指标填写表】工作表中的【填写值】列。
  
  操作流程：
    1. 先运行 01_生成空白填报表单.py 生成模板
    2. 在模板中填写企业数据
    3. 双击运行本脚本，按弹窗输入企业名称和行业
    4. 自动输出 验证结果.xlsx
"""

import numpy as np
import pandas as pd
import json
import os
import sys
import time
import warnings
from tkinter import Tk, messagebox, simpledialog
warnings.filterwarnings('ignore')

OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))

# 模型数据文件在上级的 论文模型代码 目录下
MODEL_DATA_DIR = os.path.join(OUTPUT_DIR, "..", "01_论文模型代码")
MODEL_DATA_FILE = os.path.join(MODEL_DATA_DIR, "评价数据模板.xlsx")

STAGE_NAMES = {
    1: "L1 规划级", 2: "L2 规范级", 3: "L3 整合级",
    4: "L4 优化级", 5: "L5 引领级"
}


# ────── 加载模型数据 ──────

def load_indicators(xlsx_path):
    """从"指标体系"工作表读指标定义和上下界"""
    df = pd.read_excel(xlsx_path, sheet_name="指标体系")
    indicators = df["指标名称"].tolist()
    dimensions = ["技术基础", "业务应用", "组织管理", "转型成效"]
    dim_indices = {}
    for dim in dimensions:
        dim_indices[dim] = df[df["一级维度"] == dim].index.tolist()
    bounds = {"min": df["下界"].tolist(), "max": df["上界"].tolist()}
    indicator_dir = [1] * len(indicators)
    threshold_indicators = df[df["是否门槛指标"] == "是"]["指标名称"].tolist()
    return indicators, dimensions, dim_indices, bounds, indicator_dir, threshold_indicators


def load_reference_enterprises(xlsx_path):
    """从"参考企业数据"工作表加载行业参考企业矩阵"""
    df = pd.read_excel(xlsx_path, sheet_name="参考企业数据")
    result = {}
    metadata = {}
    for _, row in df.iterrows():
        name = f"{row['企业编号']}_{row['企业名称']}"
        result[name] = [row[str(i)] for i in range(1, 37)]
        metadata[name] = {
            "编号": row["企业编号"], "名称": row["企业名称"],
            "阶段": row["对应阶段"], "行业": row["对应行业"],
        }
    return result, metadata


def load_case_enterprises(xlsx_path):
    """从"案例企业数据"工作表加载案例企业"""
    df = pd.read_excel(xlsx_path, sheet_name="案例企业数据")
    result = {}
    metadata = {}
    for _, row in df.iterrows():
        name = row["企业名称"]
        result[name] = [row[str(i)] for i in range(1, 37)]
        metadata[name] = {
            "编号": row["企业编号"], "名称": row["企业名称"],
            "阶段": row["对应阶段"], "行业": row["对应行业"], "短板": row["短板维度"],
        }
    return result, metadata


def load_stage_thresholds(xlsx_path):
    """从"阶段判定阈值"工作表加载各阶段门槛值"""
    df = pd.read_excel(xlsx_path, sheet_name="阶段判定阈值")
    thresholds = {}
    for _, row in df.iterrows():
        stage = int(row["阶段"])
        thresholds[stage] = {
            "设备联网率(%)":         float(row["设备联网率(%)"]),
            "ERP系统覆盖率(%)":      float(row["ERP系统覆盖率(%)"]),
            "MES系统部署率(%)":      float(row["MES系统部署率(%)"]),
            "数据分析应用率(%)":     float(row["数据分析应用率(%)"]),
            "数字化专业人才占比(%)": float(row["数字化专业人才占比(%)"]),
        }
    thresholds[1] = {}
    return thresholds


def load_weight_adjustments(xlsx_path):
    """从"权重调节系数"工作表加载各阶段调节系数"""
    df = pd.read_excel(xlsx_path, sheet_name="权重调节系数")
    dimensions = ["技术基础", "业务应用", "组织管理", "转型成效"]
    adjustments = {}
    for _, row in df.iterrows():
        stage = int(row["阶段"])
        adjustments[stage] = {dim: float(row[dim]) for dim in dimensions}
    return adjustments


def load_all_data(xlsx_path=None):
    """一口气加载 Excel 里的全部模型数据"""
    if xlsx_path is None:
        xlsx_path = MODEL_DATA_FILE
    if not os.path.exists(xlsx_path):
        raise FileNotFoundError(
            f"数据文件不存在: {xlsx_path}\n"
            f"请确保 评价数据模板.xlsx 在 论文模型代码/ 目录下"
        )
    indicators, dimensions, dim_indices, bounds, indicator_dir, threshold_indicators = \
        load_indicators(xlsx_path)
    ref_enterprises, ref_meta = load_reference_enterprises(xlsx_path)
    case_enterprises, case_meta = load_case_enterprises(xlsx_path)
    stage_thresholds = load_stage_thresholds(xlsx_path)
    weight_adjustments = load_weight_adjustments(xlsx_path)
    assert len(indicators) == 36, f"指标数量应该是36项，当前是{len(indicators)}项"
    return {
        "indicators": indicators, "dimensions": dimensions,
        "dim_indices": dim_indices, "bounds": bounds,
        "indicator_dir": indicator_dir, "threshold_indicators": threshold_indicators,
        "ref_enterprises": ref_enterprises, "ref_meta": ref_meta,
        "case_enterprises": case_enterprises, "case_meta": case_meta,
        "stage_thresholds": stage_thresholds, "weight_adjustments": weight_adjustments,
    }


# ────── 核心算法 ──────

def normalize_data(data_array, bounds=None):
    """极差标准化，把各指标值映射到 [0,1]"""
    n, m = data_array.shape
    normalized = np.zeros((n, m))
    for j in range(m):
        if bounds is not None:
            lo, hi = bounds["min"][j], bounds["max"][j]
        else:
            lo, hi = np.min(data_array[:, j]), np.max(data_array[:, j])
        if abs(hi - lo) < 1e-10:
            normalized[:, j] = 0.5
        else:
            normalized[:, j] = np.clip((data_array[:, j] - lo) / (hi - lo), 0, 1)
    return normalized


def calc_entropy_weights(normalized_matrix):
    """熵权法，算各指标的客观权重"""
    n, m = normalized_matrix.shape
    mat = np.where(normalized_matrix < 1e-10, 1e-10, normalized_matrix)
    col_sums = mat.sum(axis=0)
    p_mat = mat / col_sums
    k = 1.0 / np.log(n)
    entropy = -k * np.sum(p_mat * np.log(p_mat), axis=0)
    entropy = np.clip(entropy, 0, 1)
    redundancy = 1 - entropy
    total = redundancy.sum()
    if total < 1e-10:
        return np.ones(m) / m
    return redundancy / total


def determine_stage(raw_data_dict, stage_thresholds):
    """门槛指标法：看5项关键指标能跨过哪个阶段的线"""
    key_indicators = {
        "设备联网率(%)":         raw_data_dict.get("设备联网率(%)", 0),
        "ERP系统覆盖率(%)":      raw_data_dict.get("ERP系统覆盖率(%)", 0),
        "MES系统部署率(%)":      raw_data_dict.get("MES系统部署率(%)", 0),
        "数据分析应用率(%)":     raw_data_dict.get("数据分析应用率(%)", 0),
        "数字化专业人才占比(%)": raw_data_dict.get("数字化专业人才占比(%)", 0),
    }
    for stage in [5, 4, 3, 2]:
        thresholds = stage_thresholds[stage]
        if all(key_indicators[k] >= v for k, v in thresholds.items()):
            return stage, STAGE_NAMES[stage]
    return 1, STAGE_NAMES[1]


def get_adjusted_dim_weights(base_dim_weights, stage, weight_adjustments):
    """根据所处的转型阶段调整维度权重"""
    adjustments = weight_adjustments[stage]
    adjusted = {}
    for dim in base_dim_weights:
        adjusted[dim] = max(base_dim_weights[dim] + adjustments.get(dim, 0), 0.05)
    total = sum(adjusted.values())
    return {dim: w / total for dim, w in adjusted.items()}


def topsis(normalized_matrix, weights):
    """TOPSIS，算每个企业的贴近度"""
    weighted = normalized_matrix * weights
    z_pos = weighted.max(axis=0)
    z_neg = weighted.min(axis=0)
    d_pos = np.sqrt(((weighted - z_pos) ** 2).sum(axis=1))
    d_neg = np.sqrt(((weighted - z_neg) ** 2).sum(axis=1))
    closeness = d_neg / (d_pos + d_neg + 1e-12)
    return closeness, d_pos, d_neg


def calc_dimension_closeness(norm_row, weights, full_norm_matrix, dim_indices):
    """算单个企业在四个维度上各自的贴近度"""
    dim_scores = {}
    for dim, idx_list in dim_indices.items():
        dim_w = weights[idx_list]
        dim_full = full_norm_matrix[:, idx_list] * dim_w
        z_pos = dim_full.max(axis=0)
        z_neg = dim_full.min(axis=0)
        row_w = norm_row[idx_list] * dim_w
        d_p = np.sqrt(((row_w - z_pos) ** 2).sum())
        d_n = np.sqrt(((row_w - z_neg) ** 2).sum())
        dim_scores[dim] = round(float(d_n / (d_p + d_n + 1e-12) * 100), 2)
    return dim_scores


# ────── 主评价流程 ──────

def run_evaluation(data=None, verbose=True):
    """跑一遍完整的评价"""
    if data is None:
        data = load_all_data()

    indicators = data["indicators"]
    dimensions = data["dimensions"]
    dim_indices = data["dim_indices"]
    bounds = data["bounds"]
    ref_enterprises = data["ref_enterprises"]
    case_enterprises = data["case_enterprises"]
    stage_thresholds = data["stage_thresholds"]
    weight_adjustments = data["weight_adjustments"]

    def log(msg=""):
        if verbose:
            print(msg)

    log("=" * 72)
    log("  制造业企业数字化转型动态评价模型 V5")
    log("  基于改进熵权-TOPSIS方法")
    log("=" * 72)

    ref_names = list(ref_enterprises.keys())
    case_names = list(case_enterprises.keys())
    all_names = ref_names + case_names
    ref_data = [ref_enterprises[k] for k in ref_names]
    case_data = [case_enterprises[k] for k in case_names]
    all_data = np.array(ref_data + case_data, dtype=float)
    n_ref = len(ref_names)
    n_case = len(case_names)
    n_total = n_ref + n_case

    log(f"\n【评价矩阵规模】")
    log(f"  行业参考企业: {n_ref} 个（覆盖L1~L5五个转型阶段典型水平）")
    log(f"  待评案例企业: {n_case} 个")
    log(f"  评价指标数量: {len(indicators)} 项（4维度×3二级×3三级）")

    norm_matrix = normalize_data(all_data, bounds=bounds)
    log(f"\n【数据标准化】完成")

    weights = calc_entropy_weights(norm_matrix)
    base_dim_weights = {dim: float(weights[dim_indices[dim]].sum()) for dim in dimensions}

    log(f"\n【熵权法基础维度权重（{n_total}个样本）】")
    for dim in dimensions:
        log(f"  {dim}: {base_dim_weights[dim]:.4f}  ({base_dim_weights[dim]*100:.1f}%)")

    weight_series = pd.Series(weights, index=indicators).sort_values(ascending=False)
    log(f"\n【指标权重 Top10】")
    for i, (name, w) in enumerate(weight_series.head(10).items(), 1):
        log(f"  {i:2d}. {name:<32} {w:.4f}  ({w*100:.2f}%)")

    log(f"\n{'=' * 72}")
    log(f"  待评企业评价结果")
    log(f"{'=' * 72}")

    results = []
    for ci, case_name in enumerate(case_names):
        raw = case_enterprises[case_name]
        raw_dict = dict(zip(indicators, raw))

        log(f"\n{'─' * 72}")
        log(f"  ◆ {case_name}")
        log(f"{'─' * 72}")

        stage_num, stage_name = determine_stage(raw_dict, stage_thresholds)
        log(f"\n  ▶ 转型阶段判定: {stage_name}")
        key_vals = {k: raw_dict[k] for k in [
            "设备联网率(%)", "ERP系统覆盖率(%)", "MES系统部署率(%)",
            "数据分析应用率(%)", "数字化专业人才占比(%)"
        ]}
        for kn, kv in key_vals.items():
            threshold = stage_thresholds.get(stage_num + 1, {}).get(kn, "─")
            log(f"     {kn.split('(')[0]:<18} = {kv:5.1f}  (下级阈值: {threshold})")

        adj_dim_weights = get_adjusted_dim_weights(base_dim_weights, stage_num, weight_adjustments)
        log(f"\n  ▶ 动态权重调整（{stage_name}）:")
        for dim in dimensions:
            bw = base_dim_weights[dim]
            aw = adj_dim_weights[dim]
            arrow = "↑" if aw > bw + 0.005 else ("↓" if aw < bw - 0.005 else "─")
            log(f"     {dim:<8}: {bw:.4f} → {aw:.4f}  {arrow}  ({aw - bw:+.4f})")

        final_weights = np.zeros(len(indicators))
        for dim, idx_list in dim_indices.items():
            db = base_dim_weights[dim]
            da = adj_dim_weights[dim]
            ratio = da / db if db > 1e-10 else 0
            for idx in idx_list:
                final_weights[idx] = weights[idx] * ratio
        final_weights /= final_weights.sum()

        closeness_all, _, _ = topsis(norm_matrix, final_weights)
        final_score = float(closeness_all[n_ref + ci]) * 100

        dim_scores = calc_dimension_closeness(
            norm_matrix[n_ref + ci], final_weights, norm_matrix, dim_indices
        )

        log(f"\n  ▶ TOPSIS综合评价结果: {final_score:.2f} 分")
        log(f"\n  ▶ 各维度得分:")
        sorted_dims = sorted(dim_scores.items(), key=lambda x: x[1], reverse=True)
        for dim, score in sorted_dims:
            bar = "█" * int(score / 5) + "░" * (20 - int(score / 5))
            log(f"     {dim:<8}: {score:5.1f}  [{bar}]")

        best_dim = sorted_dims[0]
        worst_dim = sorted_dims[-1]
        log(f"\n  ▶ 优势维度: {best_dim[0]}  ({best_dim[1]:.1f}分)")
        log(f"  ▶ 短板维度: {worst_dim[0]}  ({worst_dim[1]:.1f}分)")

        ref_scores = {rn: float(closeness_all[ri]) * 100 for ri, rn in enumerate(ref_names)}
        log(f"\n  ▶ 行业参考对标:")
        for rn, rs in ref_scores.items():
            marker = "  ◀" if abs(rs - final_score) < 8 else ""
            log(f"     {rn:<32}: {rs:.1f} 分{marker}")

        results.append({
            "企业名称": case_name, "转型阶段": stage_name,
            "综合得分": round(final_score, 2), "各维度得分": dim_scores,
            "动态维度权重": {dim: round(w, 4) for dim, w in adj_dim_weights.items()},
            "优势维度": best_dim[0], "短板维度": worst_dim[0],
        })

    log(f"\n{'=' * 72}")
    log(f"  多企业横向对比")
    log(f"{'=' * 72}\n")
    header = f"  {'企业':<24} {'阶段':<12} {'综合分':>6}  " + \
             "  ".join(f"{d[:4]:>6}" for d in dimensions)
    log(header)
    log("  " + "─" * (len(header) - 2))
    for r in results:
        dv = "  ".join(f"{r['各维度得分'][d]:>6.1f}" for d in dimensions)
        log(f"  {r['企业名称'][:22]:<24} {r['转型阶段']:<12} {r['综合得分']:>6.1f}  {dv}")

    suggestions = {
        "技术基础": [
            "优先推进关键工序设备数字化改造，把设备联网率拉到行业平均线以上（60%+）",
            "加快MES系统选型上线，打通生产过程的数据采集和实时监控",
            "统一数据采集标准，建好企业级数据治理平台",
        ],
        "业务应用": [
            "推一把现有系统集成，打破ERP/MES/SCM之间的数据孤岛",
            "深化供应链数字化协同，把供应商数据互通和库存可视化做起来",
            "部署CRM，把客户全生命周期管起来",
        ],
        "组织管理": [
            "定个中长期的数字化人才培养计划，加大专业人才引进和内部培训投入",
            "完善转型战略规划，设专职数字化领导岗位（CIO/CDO）",
            "把数字化KPI纳进考核体系，转型目标挂钩高管绩效，强化执行",
        ],
        "转型成效": [
            "建一套数字化投入产出的量化跟踪机制，识别高价值应用场景",
            "以生产效率和质量提升为突破口，形成可复制的数字化赋能模板",
            "把数字化和新产品/新服务创新结合起来，让技术投入转化为增长动力",
        ],
    }
    log(f"\n{'=' * 72}")
    log(f"  各企业优化建议")
    log(f"{'=' * 72}")
    for r in results:
        log(f"\n  ◆ {r['企业名称']} [{r['转型阶段']}]")
        log(f"    当前短板: {r['短板维度']}（{r['各维度得分'][r['短板维度']]:.1f}分）")
        log(f"    建议措施:")
        for s in suggestions[r["短板维度"]]:
            log(f"      · {s}")

    log(f"\n\n【完成】评价流程结束")
    return results, weights, norm_matrix, n_ref


# ────── 等权重对比 + 三方案对比 + 敏感性分析 ──────

def compare_with_equal_weights(norm_matrix, weights, n_ref, case_names, results_v5,
                                indicators, dim_indices):
    """跟等权重TOPSIS比一比，看动态权重拉动了多少"""
    print(f"\n{'=' * 72}")
    print(f"  等权重TOPSIS对比（验证动态权重有效性）")
    print(f"{'=' * 72}\n")
    equal_w = np.ones(len(indicators)) / len(indicators)
    closeness_eq, _, _ = topsis(norm_matrix, equal_w)
    print(f"  {'企业':<28}  {'V5动态权重得分':>14}  {'等权重得分':>10}  {'差值':>8}")
    print(f"  {'─' * 68}")
    for ci, r in enumerate(results_v5):
        eq_score = float(closeness_eq[n_ref + ci]) * 100
        delta = r["综合得分"] - eq_score
        print(f"  {r['企业名称'][:26]:<28}  {r['综合得分']:>14.2f}  {eq_score:>10.2f}  {delta:>+8.2f}")


def compare_three_methods(norm_matrix, weights, n_ref, case_names, results_v5,
                          indicators, dim_indices, base_dim_weights,
                          weight_adjustments, case_enterprises, stage_thresholds):
    """三套方案放一起比——动态权重 vs 传统熵权 vs 等权重"""
    dimensions = list(base_dim_weights.keys())

    print(f"\n{'=' * 72}")
    print(f"  三方案对比：综合得分与分维度得分")
    print(f"{'=' * 72}")

    print(f"\n  {'─' * 68}")
    print(f"  【方案B】传统熵权TOPSIS（不调权重）")
    print(f"  {'─' * 68}")
    closeness_b, _, _ = topsis(norm_matrix, weights)
    print(f"  {'企业':<24} {'综合得分':>8}  {'技术基础':>8}  {'业务应用':>8}  {'组织管理':>8}  {'转型成效':>8}  {'短板':>8}")
    print(f"  {'─' * 86}")
    method_b_results = []
    for ci, cn in enumerate(case_names):
        score_b = float(closeness_b[n_ref + ci]) * 100
        dim_b = calc_dimension_closeness(norm_matrix[n_ref + ci], weights, norm_matrix, dim_indices)
        worst_dim = min(dim_b, key=dim_b.get)
        row = f"  {cn:<24} {score_b:>8.2f}"
        for d in dimensions:
            row += f"  {dim_b[d]:>8.2f}"
        row += f"  {worst_dim:>8}"
        print(row)
        method_b_results.append({"企业": cn, "综合得分": round(score_b, 2), "分维度得分": dim_b, "短板": worst_dim})

    print(f"\n  {'─' * 68}")
    print(f"  【方案C】等权重TOPSIS（36项指标全部等权）")
    print(f"  {'─' * 68}")
    equal_w = np.ones(len(indicators)) / len(indicators)
    closeness_c, _, _ = topsis(norm_matrix, equal_w)
    print(f"  {'企业':<24} {'综合得分':>8}  {'技术基础':>8}  {'业务应用':>8}  {'组织管理':>8}  {'转型成效':>8}  {'短板':>8}")
    print(f"  {'─' * 86}")
    method_c_results = []
    for ci, cn in enumerate(case_names):
        score_c = float(closeness_c[n_ref + ci]) * 100
        dim_c = calc_dimension_closeness(norm_matrix[n_ref + ci], equal_w, norm_matrix, dim_indices)
        worst_dim = min(dim_c, key=dim_c.get)
        row = f"  {cn:<24} {score_c:>8.2f}"
        for d in dimensions:
            row += f"  {dim_c[d]:>8.2f}"
        row += f"  {worst_dim:>8}"
        print(row)
        method_c_results.append({"企业": cn, "综合得分": round(score_c, 2), "分维度得分": dim_c, "短板": worst_dim})

    print(f"\n  {'─' * 68}")
    print(f"  三方案短板识别对比")
    print(f"  {'─' * 68}")
    print(f"  {'企业':<24} {'方案A短板':>12} {'方案B短板':>12} {'方案C短板':>12}")
    print(f"  {'─' * 64}")
    for ci, cn in enumerate(case_names):
        print(f"  {cn:<24} {results_v5[ci]['短板维度']:>12} "
              f"{method_b_results[ci]['短板']:>12} {method_c_results[ci]['短板']:>12}")

    return method_b_results, method_c_results


def sensitivity_analysis(norm_matrix, weights, n_ref, case_enterprises, indicators,
                          case_idx=0, perturbation=0.1, n_random=500, dynamic_weights=None):
    """敏感性分析：单指标扰动 + 蒙特卡洛"""

    case_names = list(case_enterprises.keys())
    base_w = dynamic_weights if dynamic_weights is not None else weights
    weight_label = "动态权重" if dynamic_weights is not None else "基础权重"

    print(f"\n{'=' * 72}")
    print(f"  敏感性分析（{case_names[case_idx]}，基准: {weight_label}）")
    print(f"{'=' * 72}")

    baseline_c, _, _ = topsis(norm_matrix, base_w)
    baseline = float(baseline_c[n_ref + case_idx]) * 100

    w_series = pd.Series(base_w, index=indicators).sort_values(ascending=False)
    top5_names = list(w_series.head(5).index)
    top5_idx = [indicators.index(n) for n in top5_names]

    print(f"\n  基准得分: {baseline:.2f}\n")
    print(f"  [单指标扰动 ±{int(perturbation*100)}%]")
    print(f"  {'指标':<32} {'扰动':>6} {'得分':>8} {'变化':>8}")
    print(f"  {'─' * 58}")

    for idx in top5_idx:
        for direction, label in [(1 + perturbation, f"+{int(perturbation*100)}%"),
                                  (1 - perturbation, f"-{int(perturbation*100)}%")]:
            w_p = base_w.copy()
            w_p[idx] *= direction
            w_p /= w_p.sum()
            c_p, _, _ = topsis(norm_matrix, w_p)
            s_p = float(c_p[n_ref + case_idx]) * 100
            print(f"  {indicators[idx]:<32} {label:>6} {s_p:>8.2f} {s_p - baseline:>+8.2f}")

    print(f"\n  [蒙特卡洛组合扰动  n={n_random}，每个权重独立±{int(perturbation*100)}%]")
    rng = np.random.default_rng(42)
    perturbed_scores = []
    for _ in range(n_random):
        noise = 1 + rng.uniform(-perturbation, perturbation, size=len(base_w))
        w_p = base_w * noise
        w_p /= w_p.sum()
        c_p, _, _ = topsis(norm_matrix, w_p)
        perturbed_scores.append(float(c_p[n_ref + case_idx]) * 100)

    arr = np.array(perturbed_scores)
    print(f"  均值:  {arr.mean():.2f}")
    print(f"  标准差: {arr.std():.2f}")
    print(f"  95%置信区间: [{np.percentile(arr, 2.5):.2f}, {np.percentile(arr, 97.5):.2f}]")
    width = np.percentile(arr, 97.5) - np.percentile(arr, 2.5)
    print(f"  区间宽度: {width:.2f}")
    print(f"  结论: {'结果稳健（区间宽度<1分）' if width < 1 else '结果稳健（区间宽度<10分）' if width < 10 else '结果受权重影响较大'}")
    return {
        "基准权重类型": weight_label, "基准得分": round(baseline, 2),
        "均值": round(float(arr.mean()), 2), "标准差": round(float(arr.std()), 2),
        "95%置信区间": [round(float(np.percentile(arr, 2.5)), 2),
                       round(float(np.percentile(arr, 97.5)), 2)],
        "区间宽度": round(float(width), 2),
    }


# ────── 企业数据验证工具 ──────

def load_simulated_data_from_template():
    """从模板 Excel 的指标填写表读企业填入的数据"""
    template_path = os.path.join(OUTPUT_DIR, "企业验证数据模板.xlsx")
    if not os.path.exists(template_path):
        raise FileNotFoundError(
            f"没找到模板文件: {template_path}\n"
            f"先生成: python 01_生成空白填报表单.py"
        )
    df = pd.read_excel(template_path, sheet_name="指标填写表")
    values = df["填写值"].tolist()
    if all(v == "" or pd.isna(v) for v in values):
        raise ValueError("模板的【填写值】列还是空的，先填数据。")
    values = [float(v) if v != "" and not pd.isna(v) else 0.0 for v in values]
    return {
        "企业编号": "D", "企业名称": "外部验证企业",
        "对应行业": "制造业", "企业规模": "未知", "指标值": values,
    }


def inject_enterprise_into_model(enterprise_info):
    """把新企业的数据塞进模型的数据结构里"""
    original_data = load_all_data()
    new_enterprise_name = enterprise_info.get("企业名称", "外部验证企业")
    new_enterprise_values = enterprise_info["指标值"]

    # 先粗略估计一下阶段
    indicators = original_data["indicators"]
    stage_thresholds = original_data["stage_thresholds"]
    raw_dict = dict(zip(indicators, new_enterprise_values))
    estimated_stage_num, estimated_stage_name = determine_stage(raw_dict, stage_thresholds)

    # 推断短板
    bounds = original_data["bounds"]
    norm_vals = []
    for i, v in enumerate(new_enterprise_values):
        lo, hi = bounds["min"][i], bounds["max"][i]
        if abs(hi - lo) < 1e-10:
            norm_vals.append(0.5)
        else:
            norm_vals.append(np.clip((v - lo) / (hi - lo), 0, 1))
    dim_scores = {}
    for dim, idx_list in original_data["dim_indices"].items():
        dim_scores[dim] = np.mean([norm_vals[i] for i in idx_list])
    weak_dim = min(dim_scores, key=dim_scores.get)

    # 加进案例企业列表
    original_data["case_enterprises"][new_enterprise_name] = new_enterprise_values
    original_data["case_meta"][new_enterprise_name] = {
        "编号": enterprise_info.get("企业编号", "D"),
        "名称": new_enterprise_name,
        "阶段": estimated_stage_name,
        "行业": enterprise_info.get("对应行业", "制造业"),
        "短板": weak_dim,
    }
    return original_data, estimated_stage_name, dim_scores


def run_full_validation(original_data, enterprise_name):
    """跑完整验证，收集所有结果"""
    print("\n" + "=" * 72)
    print(f"  企业数据验证 · 模拟运行")
    print(f"  验证对象: {enterprise_name}")
    print("=" * 72)

    indicators = original_data["indicators"]
    raw_values = original_data["case_enterprises"][enterprise_name]
    raw_dict = dict(zip(indicators, raw_values))
    stage_num, stage_name = determine_stage(raw_dict, original_data["stage_thresholds"])
    print(f"\n  ▶ 阶段判定: {stage_name}")

    print(f"\n  ◆ 5项关键门槛指标:")
    key_indicators = {
        "设备联网率(%)": raw_dict["设备联网率(%)"],
        "ERP系统覆盖率(%)": raw_dict["ERP系统覆盖率(%)"],
        "MES系统部署率(%)": raw_dict["MES系统部署率(%)"],
        "数据分析应用率(%)": raw_dict["数据分析应用率(%)"],
        "数字化专业人才占比(%)": raw_dict["数字化专业人才占比(%)"],
    }
    for name, val in key_indicators.items():
        cur_t = original_data["stage_thresholds"].get(stage_num, {}).get(name, "─")
        next_t = original_data["stage_thresholds"].get(stage_num + 1, {}).get(name, "─") if stage_num < 5 else "─(最高级)"
        print(f"     {name:<28} = {val:>6.1f}  (本阶段阈值: {cur_t}, 下级阈值: {next_t})")

    # 完整评价
    results, weights, norm_matrix, n_ref = run_evaluation(data=original_data, verbose=True)

    # 等权重对比
    compare_with_equal_weights(
        norm_matrix, weights, n_ref,
        list(original_data["case_enterprises"].keys()), results,
        original_data["indicators"], original_data["dim_indices"]
    )

    base_dim_weights = {}
    for dim in original_data["dimensions"]:
        idx_list = original_data["dim_indices"][dim]
        base_dim_weights[dim] = sum(weights[idx] for idx in idx_list)

    case_names = list(original_data["case_enterprises"].keys())

    method_b_results, method_c_results = compare_three_methods(
        norm_matrix, weights, n_ref, case_names, results,
        original_data["indicators"], original_data["dim_indices"], base_dim_weights,
        original_data["weight_adjustments"], original_data["case_enterprises"],
        original_data["stage_thresholds"]
    )

    dynamic_weights_per_case = {}
    for ci, cn in enumerate(case_names):
        raw = original_data["case_enterprises"][cn]
        raw_dict = dict(zip(original_data["indicators"], raw))
        s_num, _ = determine_stage(raw_dict, original_data["stage_thresholds"])
        adj_dim_weights = get_adjusted_dim_weights(base_dim_weights, s_num, original_data["weight_adjustments"])
        final_w = np.zeros(len(original_data["indicators"]))
        for dim, idx_list in original_data["dim_indices"].items():
            db = base_dim_weights[dim]
            da = adj_dim_weights[dim]
            ratio = da / db if db > 1e-10 else 0
            for idx in idx_list:
                final_w[idx] = weights[idx] * ratio
        final_w /= final_w.sum()
        dynamic_weights_per_case[cn] = final_w

    sa_results = []
    for ci in range(len(case_names)):
        sa_r = sensitivity_analysis(
            norm_matrix, weights, n_ref,
            original_data["case_enterprises"], original_data["indicators"],
            case_idx=ci, dynamic_weights=dynamic_weights_per_case[case_names[ci]]
        )
        sa_results.append(sa_r)

    print(f"\n{'=' * 72}")
    print(f"  敏感性分析汇总（动态权重基准）")
    print(f"{'=' * 72}")
    print(f"  {'企业':<36} {'基准得分':>8} {'95%CI':>22} {'区间宽度':>8}")
    print(f"  {'─' * 78}")
    for ci, cn in enumerate(case_names):
        r = sa_results[ci]
        ci_str = f"[{r['95%置信区间'][0]:.2f}, {r['95%置信区间'][1]:.2f}]"
        marker = " ◀ 新验证企业" if cn == enterprise_name else ""
        print(f"  {cn:<36} {r['基准得分']:>8.2f} {ci_str:>22} {r['区间宽度']:>8.2f}{marker}")

    out_path = os.path.join(OUTPUT_DIR, "验证结果.xlsx")
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df_summary = pd.DataFrame([{
            "企业": r["企业名称"], "转型阶段": r["转型阶段"], "综合得分": r["综合得分"],
            "技术基础": r["各维度得分"].get("技术基础", ""),
            "业务应用": r["各维度得分"].get("业务应用", ""),
            "组织管理": r["各维度得分"].get("组织管理", ""),
            "转型成效": r["各维度得分"].get("转型成效", ""),
            "优势维度": r["优势维度"], "短板维度": r["短板维度"],
        } for r in results])
        df_summary.to_excel(writer, sheet_name="评价结果汇总", index=False)

        weight_rows = []
        for r in results:
            for dim, w in r["动态维度权重"].items():
                weight_rows.append({"企业": r["企业名称"], "维度": dim, "动态权重": w})
        df_weights = pd.DataFrame(weight_rows)
        df_weights.to_excel(writer, sheet_name="维度权重", index=False)

        comparison_rows = []
        for r in results:
            comparison_rows.append({
                "企业": r["企业名称"], "方案": "方案A_动态权重", "综合得分": r["综合得分"],
                "技术基础": r["各维度得分"].get("技术基础", ""),
                "业务应用": r["各维度得分"].get("业务应用", ""),
                "组织管理": r["各维度得分"].get("组织管理", ""),
                "转型成效": r["各维度得分"].get("转型成效", ""), "短板": r["短板维度"],
            })
        for mb in method_b_results:
            comparison_rows.append({
                "企业": mb["企业"], "方案": "方案B_传统熵权TOPSIS", "综合得分": mb["综合得分"],
                "技术基础": mb["分维度得分"].get("技术基础", ""),
                "业务应用": mb["分维度得分"].get("业务应用", ""),
                "组织管理": mb["分维度得分"].get("组织管理", ""),
                "转型成效": mb["分维度得分"].get("转型成效", ""), "短板": mb["短板"],
            })
        for mc in method_c_results:
            comparison_rows.append({
                "企业": mc["企业"], "方案": "方案C_等权重TOPSIS", "综合得分": mc["综合得分"],
                "技术基础": mc["分维度得分"].get("技术基础", ""),
                "业务应用": mc["分维度得分"].get("业务应用", ""),
                "组织管理": mc["分维度得分"].get("组织管理", ""),
                "转型成效": mc["分维度得分"].get("转型成效", ""), "短板": mc["短板"],
            })
        df_comparison = pd.DataFrame(comparison_rows)
        df_comparison.to_excel(writer, sheet_name="三方案对比", index=False)

        df_sensitivity = pd.DataFrame([
            {"企业": cn, **r} for cn, r in zip(case_names, sa_results)
        ])
        df_sensitivity.to_excel(writer, sheet_name="敏感性分析", index=False)

    print(f"\n{'=' * 72}")
    print(f"  ✅ 验证完成！结果已保存至: {out_path}")
    print(f"{'=' * 72}")
    return results, out_path


def generate_validation_summary(results, enterprise_name):
    """打出验证摘要"""
    target_result = None
    for r in results:
        if r["企业名称"] == enterprise_name:
            target_result = r
            break
    if target_result is None:
        print(f"⚠ 没找到企业 {enterprise_name} 的评价结果")
        return

    r = target_result
    print(f"\n{'#' * 72}")
    print(f"#  验证摘要报告")
    print(f"{'#' * 72}")
    print(f"")
    print(f"  验证企业: {r['企业名称']}")
    print(f"  转型阶段: {r['转型阶段']}")
    print(f"  综合得分: {r['综合得分']} 分")
    print(f"")
    print(f"  维度得分:")
    for dim, score in sorted(r["各维度得分"].items(), key=lambda x: x[1], reverse=True):
        bar = "█" * int(score / 5) + "░" * (20 - int(score / 5))
        print(f"    {dim:<8}: {score:>5.1f}  [{bar}]")
    print(f"")
    print(f"  优势维度: {r['优势维度']}")
    print(f"  短板维度: {r['短板维度']}")
    print(f"")
    print(f"  动态维度权重:")
    for dim, w in r["动态维度权重"].items():
        print(f"    {dim:<8}: {w:.4f} ({w*100:.1f}%)")
    print(f"")
    print(f"  模型稳定性:")
    print(f"    评价方法: 改进熵权-TOPSIS（含阶段前置判定与动态权重调整）")
    print(f"    参考企业: 6个行业级标杆（覆盖L1-L5）")
    print(f"    方案对比: 动态权重 vs 传统熵权 vs 等权重")
    print(f"    敏感性分析: 蒙特卡洛500次权重扰动")
    print(f"")
    print(f"  输出文件: 验证结果.xlsx（含4个工作表）")
    print(f"")


# ────── 入口 ──────

if __name__ == "__main__":
    root = Tk()
    root.withdraw()

    template_path = os.path.join(OUTPUT_DIR, "企业验证数据模板.xlsx")
    if not os.path.exists(template_path):
        messagebox.showerror(
            "找不到模板",
            f"找不到企业验证数据模板！\n\n"
            f"请确保以下文件存在：\n{template_path}\n\n"
            f"如未生成模板，请先运行：\n  01_生成空白填报表单.py"
        )
        sys.exit(1)

    # 提前检查是否已填写数据
    df_check = pd.read_excel(template_path, sheet_name="指标填写表")
    check_values = df_check["填写值"].tolist()
    filled_count = sum(1 for v in check_values if v != "" and not pd.isna(v))
    if filled_count == 0:
        messagebox.showwarning(
            "数据未填写",
            f"企业验证数据模板中尚未填写任何数据。\n\n"
            f"请先打开 企业验证数据模板.xlsx\n"
            f"在【指标填写表】的 K列（填写值）\n"
            f"逐项填入贵企业的36项指标数据后，再运行本脚本。"
        )
        sys.exit(1)

    messagebox.showinfo(
        "企业数字化评价验证工具",
        "欢迎使用制造业企业数字化转型评价模型\n\n"
        "本工具将根据您在【企业验证数据模板.xlsx】中\n"
        "填写的企业数据，自动完成：\n"
        "  · 转型阶段判定\n"
        "  · 综合评分\n"
        "  · 方案对比\n"
        "  · 敏感性分析\n\n"
        "请确保已在模板的【指标填写表】中填好数据。\n"
        "点击确定继续..."
    )

    enterprise_name = simpledialog.askstring(
        "企业名称",
        "请输入企业名称：",
        initialvalue="外部验证企业"
    )
    if not enterprise_name or enterprise_name.strip() == "":
        enterprise_name = "外部验证企业"
    enterprise_name = enterprise_name.strip()

    enterprise_industry = simpledialog.askstring(
        "所属行业",
        "请输入企业所属行业：",
        initialvalue="制造业"
    )
    if not enterprise_industry or enterprise_industry.strip() == "":
        enterprise_industry = "制造业"
    enterprise_industry = enterprise_industry.strip()

    print("╔══════════════════════════════════════════════════════════════╗")
    print("║    制造业企业数字化转型评价模型 · 数据验证工具               ║")
    print("║    论文《制造业企业数字化转型评价方法研究》                  ║")
    print("╚══════════════════════════════════════════════════════════════╝")

    print(f"\n从企业验证数据模板.xlsx加载...")
    try:
        enterprise_info = load_simulated_data_from_template()
    except ValueError as e:
        messagebox.showerror("数据错误", str(e))
        sys.exit(1)

    enterprise_info["企业名称"] = enterprise_name
    enterprise_info["对应行业"] = enterprise_industry

    print(f"  ✓ 已读取 {len(enterprise_info['指标值'])} 项指标数据")
    print(f"  ✓ 企业: {enterprise_info['企业名称']}")
    print(f"  ✓ 行业: {enterprise_info['对应行业']}")

    print(f"\n正在初始化评价模型并注入企业数据...")

    model_data, estimated_stage, dim_preview = inject_enterprise_into_model(enterprise_info)

    print(f"  ✓ 初步阶段估计: {estimated_stage}")
    print(f"  ✓ 维度标准化均值预览:")
    for dim, score in sorted(dim_preview.items(), key=lambda x: x[1], reverse=True):
        bar = "█" * int(score * 20) + "░" * (20 - int(score * 20))
        print(f"      {dim:<8}: {score:.3f} [{bar}]")

    print(f"\n开始运行完整验证流程...\n")
    time.sleep(0.5)

    results, out_path = run_full_validation(model_data, enterprise_info["企业名称"])
    generate_validation_summary(results, enterprise_info["企业名称"])

    print(f"验证流程全部完成。详细结果请查看: {out_path}")

    messagebox.showinfo(
        "验证完成",
        f"评价流程全部完成！\n\n"
        f"结果文件：验证结果.xlsx\n"
        f"所在位置：\n{OUTPUT_DIR}\n\n"
        f"包含4个工作表：\n"
        f"  1. 评价结果汇总\n"
        f"  2. 维度权重\n"
        f"  3. 三方案对比\n"
        f"  4. 敏感性分析"
    )
