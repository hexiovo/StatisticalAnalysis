import os
import pandas as pd
import numpy as np
import statsmodels.formula.api as smf

def mediation_search(file_path, x_var, y_var, exclude_cols=None, output_dir=None):
    """
    自动测试Excel中每个变量作为中介变量（M）的显著性，输出a、b、c'路径及p值结果。
    参数：
        file_path : str, 输入xlsx文件路径
        x_var : str, 自变量列名
        y_var : str, 因变量列名
        exclude_cols : list, 要排除的列名
        output_dir : str, 输出目录（默认为输入文件所在目录）
    """

    # ---------- 1. 读取数据 ----------
    df = pd.read_excel(file_path)
    print(f"读取数据，共 {df.shape[0]} 行，{df.shape[1]} 列。")

    # ---------- 2. 排除指定列 ----------
    if exclude_cols:
        df = df.drop(columns=exclude_cols, errors='ignore')
    print(f"排除后剩余 {df.shape[1]} 列。")

    # ---------- 3. 虚拟变量化 ----------
    df = pd.get_dummies(df, drop_first=True)
    print("已自动虚拟变量化。")

    if x_var not in df.columns or y_var not in df.columns:
        raise ValueError(f"未找到自变量 {x_var} 或因变量 {y_var}")

    # ---------- 4. 确定候选变量 ----------
    candidates = [c for c in df.columns if c not in [x_var, y_var]]
    results = []

    # ---------- 5. 循环进行中介分析 ----------
    for m in candidates:
        try:
            # a 路径：M ~ X
            model_a = smf.ols(f"{m} ~ {x_var}", data=df).fit()
            p_a = model_a.pvalues.get(x_var, np.nan)
            beta_a = model_a.params.get(x_var, np.nan)

            # b、c' 路径：Y ~ X + M
            model_b = smf.ols(f"{y_var} ~ {x_var} + {m}", data=df).fit()
            p_b = model_b.pvalues.get(m, np.nan)
            beta_b = model_b.params.get(m, np.nan)

            # c' 路径（X→Y 的直接效应）
            p_c_prime = model_b.pvalues.get(x_var, np.nan)
            beta_c_prime = model_b.params.get(x_var, np.nan)

            # c 路径（总效应：Y ~ X）
            model_c = smf.ols(f"{y_var} ~ {x_var}", data=df).fit()
            p_c = model_c.pvalues.get(x_var, np.nan)
            beta_c = model_c.params.get(x_var, np.nan)

        except Exception as e:
            print(f"变量 {m} 出错：{e}")
            p_a, p_b, p_c, p_c_prime = np.nan, np.nan, np.nan, np.nan
            beta_a, beta_b, beta_c, beta_c_prime = np.nan, np.nan, np.nan, np.nan

        results.append({
            "中介变量": m,
            "β(X→M)": beta_a, "p(X→M)": p_a,
            "β(M→Y)": beta_b, "p(M→Y)": p_b,
            "β(X→Y)总效应(c)": beta_c, "p(X→Y)总效应(c)": p_c,
            "β(X→Y)直接效应(c')": beta_c_prime, "p(X→Y)直接效应(c')": p_c_prime
        })

    # ---------- 6. 输出结果 ----------
    df_result = pd.DataFrame(results)
    if output_dir is None:
        output_dir = os.path.dirname(file_path)
    os.makedirs(output_dir, exist_ok=True)

    base_name = os.path.splitext(os.path.basename(file_path))[0]
    save_path = os.path.join(output_dir, f"{base_name}_mediation.xlsx")

    df_result.to_excel(save_path, index=False)
    print(f"中介分析结果已保存至：{save_path}")

    return df_result


def extract_mediation(output_dir, p_threshold=0.05, summary_name="summary_significant_mediation.xlsx"):
    """
    汇总 output_dir 下各因变量子文件夹中的中介分析结果文件。
    文件名包含 '_mediation' 的 Excel 文件。
    检查两个间接效应和总效应的 p 值是否都 <= 阈值。
    满足条件的整行数据（包括 β 和 p 列）保存到汇总文件，添加 y_var 列。

    参数：
        output_dir : str
            主输出文件夹路径
        p_threshold : float
            显著性阈值，默认 0.05
        summary_name : str
            汇总文件名，默认 summary_significant_mediation.xlsx
    """
    summary_records = []

    # 遍历每个因变量子文件夹
    for y_var in os.listdir(output_dir):
        subdir = os.path.join(output_dir, y_var)
        if not os.path.isdir(subdir):
            continue

        # 筛选文件名包含 "_mediation"
        xlsx_files = [f for f in os.listdir(subdir) if f.endswith('.xlsx') and '_mediation' in f.lower()]
        print(f"\n📂 正在处理因变量 {y_var} 下的中介文件，共 {len(xlsx_files)} 个")

        for fname in xlsx_files:
            fpath = os.path.join(subdir, fname)
            try:
                df = pd.read_excel(fpath)
                if df.empty:
                    continue

                # 检查是否包含必要列
                required_p_cols = ['p(X→M)', 'p(M→Y)', 'p(X→Y)总效应(c)']
                missing_cols = [c for c in required_p_cols if c not in df.columns]
                if missing_cols:
                    print(f"⚠️ 文件 {fname} 缺少列 {missing_cols}，跳过")
                    continue

                # 判断是否显著
                sig_mask = (df['p(X→M)'] <= p_threshold) & \
                           (df['p(M→Y)'] <= p_threshold) & \
                           (df['p(X→Y)总效应(c)'] <= p_threshold)

                sig_rows = df[sig_mask]
                if not sig_rows.empty:
                    sig_rows = sig_rows.copy()
                    sig_rows.insert(0, "file_name", fname)
                    sig_rows.insert(0, "y_var", y_var)
                    summary_records.append(sig_rows)

            except Exception as e:
                print(f"❌ 文件 {fname} 读取失败，错误：{e}")

    # 保存汇总
    if summary_records:
        summary_df = pd.concat(summary_records, ignore_index=True)
        summary_path = os.path.join(output_dir, summary_name)
        summary_df.to_excel(summary_path, index=False)
        print(f"\n✅ 汇总完成！共提取 {len(summary_df)} 条显著结果，已保存至 {summary_path}")
    else:
        print("\n⚠️ 未找到满足条件（两个间接效应和总效应 p ≤ 阈值）的结果。")