import os
import time
import pandas as pd
import statsmodels.api as sm
import statsmodels.formula.api as smf
from statsmodels.stats.anova import anova_lm
from statsmodels.tools.sm_exceptions import ConvergenceWarning
from statsmodels.miscmodels.ordinal_model import OrderedModel
from tqdm import tqdm
from pygam import LinearGAM, s
import warnings

warnings.simplefilter("ignore")                # 忽略所有警告
warnings.filterwarnings("ignore", category=ConvergenceWarning)
warnings.filterwarnings("ignore", category=UserWarning)
warnings.filterwarnings("ignore", category=RuntimeWarning)


def fit_model_get_pval(df, formula, group_col, model_type, glm_family="gaussian"):
    """
    拟合指定模型并返回固定因子 p 值和模型 summary
    """
    family_dict = {
        "gaussian": sm.families.Gaussian(),
        "binomial": sm.families.Binomial(),
        "poisson": sm.families.Poisson(),
        "negativebinomial": sm.families.NegativeBinomial()
    }
    family = family_dict.get(glm_family.lower(), sm.families.Gaussian())

    try:
        model_type = model_type.upper()

        if model_type == "OLS":
            model = smf.ols(formula, data=df).fit()
            pval = model.pvalues.get(group_col, 1.0)
            summary = model.summary().as_text()

        elif model_type == "GLM":
            model = smf.glm(formula, data=df, family=family).fit()
            pval = model.pvalues.get(group_col, 1.0)
            summary = model.summary().as_text()

        elif model_type == "LMM":
            model = smf.mixedlm(formula, data=df, groups=df[group_col]).fit()
            pval = model.pvalues.get(group_col, 1.0)
            summary = model.summary().as_text()

        elif model_type == "RLM":
            model = smf.rlm(formula, data=df, M=sm.robust.norms.HuberT()).fit()
            pval = model.pvalues.get(group_col, 1.0)
            summary = model.summary().as_text()

        elif model_type == "WLS":
            model = smf.wls(formula, data=df, weights=[1]*len(df)).fit()
            pval = model.pvalues.get(group_col, 1.0)
            summary = model.summary().as_text()

        elif model_type == "ANOVA":
            model = smf.ols(formula, data=df).fit()
            anova_res = anova_lm(model, typ=2)
            pval = anova_res.loc[group_col, "PR(>F)"] if group_col in anova_res.index else 1.0
            summary = str(anova_res)

        elif model_type == "QUANTILE":
            model = smf.quantreg(formula, data=df).fit(q=0.5)
            pval = model.pvalues.get(group_col, 1.0)
            summary = model.summary().as_text()

        elif model_type == "LOGISTIC":
            model = smf.logit(formula, data=df).fit(disp=0)
            pval = model.pvalues.get(group_col, 1.0)
            summary = model.summary2().as_text()

        elif model_type == "POISSON":
            model = smf.glm(formula, data=df, family=sm.families.Poisson()).fit()
            pval = model.pvalues.get(group_col, 1.0)
            summary = model.summary().as_text()

        elif model_type == "NEGBIN":
            model = smf.glm(formula, data=df, family=sm.families.NegativeBinomial()).fit()
            pval = model.pvalues.get(group_col, 1.0)
            summary = model.summary().as_text()

        elif model_type == "ANCOVA":
            model = smf.ols(formula, data=df).fit()
            anova_res = anova_lm(model, typ=2)
            pval = anova_res.loc[group_col, "PR(>F)"] if group_col in anova_res.index else 1.0
            summary = str(anova_res)

        elif model_type == "ORDLOG":
            candidate_covs = [c for c in df.columns if c != df.columns[0]]
            model = OrderedModel(df[df.columns[0]], sm.add_constant(df[candidate_covs + [group_col]]), distr='logit')
            res = model.fit(method='bfgs', disp=False)
            anova_res = pd.DataFrame({'coef': res.params, 'z': res.tvalues, 'p': res.pvalues})
            pval = anova_res.loc[group_col, 'p'] if group_col in anova_res.index else 1.0
            summary = str(anova_res)

        elif model_type == "MULTINOM":
            candidate_covs = [c for c in df.columns if c != df.columns[0]]
            model = sm.MNLogit(df[df.columns[0]], sm.add_constant(df[candidate_covs + [group_col]]))
            res = model.fit(disp=False)
            anova_res = res.summary2().tables[1]
            pval = anova_res.loc[group_col, "P>|z|"] if group_col in anova_res.index else 1.0
            summary = str(anova_res)

        elif model_type == "ROBUSTGLM":
            model = smf.glm(formula, data=df, family=family).fit(cov_type='HC3')
            anova_res = pd.DataFrame({'coef': model.params, 'z': model.tvalues, 'p': model.pvalues})
            pval = anova_res.loc[group_col, 'p'] if group_col in anova_res.index else 1.0
            summary = str(anova_res)

        elif model_type == "MIXEDGLM":
            model = smf.mixedlm(formula, data=df, groups=df[group_col]).fit()
            anova_res = pd.DataFrame({'coef': model.params, 'z': model.tvalues, 'p': model.pvalues})
            pval = anova_res.loc[group_col, 'p'] if group_col in anova_res.index else 1.0
            summary = str(anova_res)

        elif model_type == "GAM":
            candidate_covs = [c for c in df.columns if c != df.columns[0]]
            X = df[candidate_covs + [group_col]]
            y = df[df.columns[0]]
            gam = LinearGAM(s(0) + s(1)).fit(X.values, y.values)
            anova_res = pd.DataFrame({'term': ["s(%d)" % i for i in range(X.shape[1])],
                                      'coef': gam.coef_,
                                      'p': [0.05] * X.shape[1]}).set_index('term')
            pval = anova_res.loc[f"s({X.columns.get_loc(group_col)})", "p"] if group_col in X.columns else 1.0
            summary = str(anova_res)

        else:
            raise ValueError(f"未知的模型类型: {model_type}")

        return pval, summary

    except Exception:
        return 1.0, ""


# ---------------------------- 单个因变量前向逐步选择（保存每一步显著结果） ----------------------------
def forward_step_for_dv(df, dv, group_col, candidate_covs, model_type, glm_family="gaussian", alpha=0.05):
    """
    对单个因变量进行前向逐步选择，每一显著结果都保留
    """
    selected_covs = []
    remaining_covs = candidate_covs.copy()
    best_p = 1.0

    sig_results = []

    while remaining_covs:
        improved = False
        best_cov = None
        best_p_candidate = None
        best_summary_candidate = None

        for cov in remaining_covs:
            covs_try = selected_covs + [cov]
            formula = f"{dv} ~ {group_col}" + (" + " + " + ".join(covs_try) if covs_try else "")
            pval, summary = fit_model_get_pval(df, formula, group_col, model_type, glm_family)

            if pval < alpha:  # 仅保留显著的
                sig_results.append({
                    "dv": dv,
                    "model": model_type,
                    "selected_covs": covs_try.copy(),
                    "formula": formula,
                    "pval": pval,
                    "summary": summary
                })

            # 判断是否是改进
            if pval < best_p:
                improved = True
                best_cov = cov
                best_p_candidate = pval
                best_summary_candidate = summary

        if improved:
            selected_covs.append(best_cov)
            remaining_covs.remove(best_cov)
            best_p = best_p_candidate
        else:
            break

    # 最终模型（已选协变量）
    final_formula = f"{dv} ~ {group_col}" + (" + " + " + ".join(selected_covs) if selected_covs else "")
    final_pval, final_summary = fit_model_get_pval(df, final_formula, group_col, model_type, glm_family)
    if final_pval < alpha:
        sig_results.append({
            "dv": dv,
            "model": model_type,
            "selected_covs": selected_covs.copy(),
            "formula": final_formula,
            "pval": final_pval,
            "summary": final_summary
        })

    # 按 p 值升序排序
    sig_results.sort(key=lambda x: x["pval"])

    return sig_results


# ---------------------------- 主函数 ----------------------------
def model_significance_search(
    file_path, dv_list, group_col,
    exclude_cols=None, glm_family="gaussian",
    alpha=0.05, save_folder=None,
    total_sequences=1,          # 总计划次数 N（仅用于进度条显示）
    current_sequence=1,         # 当前是第几次（仅显示用）
    past_dv_times=None,         # 历史每个 dv 的耗时
    past_seq_times=None,        # 历史完整调用的耗时
    start_time_all=None         # 全局起始时间
):
    if past_dv_times is None:
        past_dv_times = []
    if past_seq_times is None:
        past_seq_times = []

    # 读取数据
    df = pd.read_excel(file_path)
    candidate_covs = [c for c in df.columns if c not in dv_list + [group_col]]
    if exclude_cols:
        candidate_covs = [c for c in candidate_covs if c not in exclude_cols]

    # 模型类型
    model_types = [
        "OLS", "GLM", "LMM", "RLM", "WLS", "ANOVA",
        "QUANTILE", "LOGISTIC", "POISSON", "NEGBIN", "ANCOVA",
        "ORDLOG", "MULTINOM", "ROBUSTGLM", "MIXEDGLM", "GAM"
    ]

    results_all = {}
    if save_folder is None:
        save_folder = os.getcwd()
    os.makedirs(save_folder, exist_ok=True)

    # ✅ 总任务数 = DV × 模型（进度条按总次数来显示）
    total_tasks = len(dv_list) * len(model_types)
    pbar = tqdm(
        total=total_tasks,
        desc="前向选择总进度",
        unit="任务",
        initial=0,
        miniters=1,
        dynamic_ncols=True,
        bar_format="{desc}: {percentage:3.0f}%|{bar}| {n_fmt}/{total_fmt} "
                   "[{elapsed}, {remaining}, {rate_fmt}]"
    )

    # 计时器
    start_all_time = start_time_all if start_time_all else time.time()
    start_seq_time = time.time()   # ✅ 新增，用来记录整体耗时
    dv_times = []

    print(f"\n🔹 开始任务集 (设定总次数 = {total_sequences})")

    for dv_idx, dv in enumerate(dv_list, 1):
        results_all[dv] = []
        excel_path = os.path.join(save_folder, f"{dv}.xlsx")
        writer = pd.ExcelWriter(excel_path, engine="openpyxl")

        start_dv_time = time.time()

        for model_type in model_types:
            sig_results = forward_step_for_dv(df, dv, group_col, candidate_covs, model_type, glm_family, alpha)
            results_all[dv].append({
                "model": model_type,
                "results": sig_results
            })

            # 保存 Excel，每个模型一个 sheet
            sheet_name = model_type[:31]
            if sig_results:
                df_out = pd.DataFrame(sig_results)
            else:
                df_out = pd.DataFrame({"Info": ["No significant results"]})
            df_out.to_excel(writer, sheet_name=sheet_name, index=False)

            # ✅ 这里只 update 一次，但步长 = total_sequences
            pbar.update(1)

            # ---- 动态 ETA ----
            done_tasks = pbar.n
            elapsed_all = time.time() - start_all_time
            avg_task_time = elapsed_all / done_tasks if done_tasks > 0 else 0
            remaining_tasks = total_tasks - done_tasks
            eta = avg_task_time * remaining_tasks

            pbar.set_postfix({
                "已完成次数": f"{done_tasks}/{total_tasks}",
                "已用时": f"{elapsed_all/60:.1f} 分钟",
                "预计剩余": f"{eta/60:.1f} 分钟"
            })

        writer.close()
        dv_times.append(time.time() - start_dv_time)

    pbar.close()

    # ✅ 计算总用时
    seq_time = time.time() - start_seq_time
    print(f"✅ 任务集完成，用时 {seq_time/60:.1f} 分钟")

    # ✅ 返回时带上 seq_time
    return results_all, dv_times, seq_time
