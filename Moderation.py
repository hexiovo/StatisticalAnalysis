import pandas as pd
import statsmodels.formula.api as smf
import os
import numpy as np

def moderation_search(file_path, x_var, y_var, exclude_cols=None, output_dir=None):
    """
    自动测试Excel中每个变量作为调节变量（Z）的显著性（交互项p值）及三条路径结果。
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

    # ---------- 4. 候选调节变量 ----------
    candidates = [c for c in df.columns if c not in [x_var, y_var]]
    results = []

    # ---------- 5. 循环进行调节分析 ----------
    for z in candidates:
        try:
            # 完整模型：Y ~ X + Z + X*Z
            model = smf.ols(f"{y_var} ~ {x_var} * {z}", data=df).fit()

            # 构造交互项名（statsmodels自动命名为 X:Z）
            interaction = f"{x_var}:{z}"

            # 提取三条路径的系数与p值
            beta_x = model.params.get(x_var, np.nan)
            p_x = model.pvalues.get(x_var, np.nan)

            beta_z = model.params.get(z, np.nan)
            p_z = model.pvalues.get(z, np.nan)

            beta_inter = model.params.get(interaction, np.nan)
            p_inter = model.pvalues.get(interaction, np.nan)

        except Exception as e:
            print(f"变量 {z} 出错：{e}")
            beta_x = beta_z = beta_inter = np.nan
            p_x = p_z = p_inter = np.nan

        results.append({
            "调节变量": z,
            "β(X→Y)": beta_x, "p(X→Y)": p_x,
            "β(Z→Y)": beta_z, "p(Z→Y)": p_z,
            "β(X×Z→Y)": beta_inter, "p(X×Z→Y)": p_inter
        })

    # ---------- 6. 输出结果 ----------
    df_result = pd.DataFrame(results)
    if output_dir is None:
        output_dir = os.path.dirname(file_path)
    os.makedirs(output_dir, exist_ok=True)

    base_name = os.path.splitext(os.path.basename(file_path))[0]
    save_path = os.path.join(output_dir, f"{base_name}_moderation.xlsx")

    df_result.to_excel(save_path, index=False)
    print(f"调节分析结果已保存至：{save_path}")

    return df_result


def extract_moderation(output_dir, p_threshold=0.05, summary_name="summary_significant_moderation.xlsx"):
    """
    汇总 output_dir 下各因变量子文件夹中的调节分析结果文件。
    文件名包含 '_moderate' 或 '_moderation' 的 Excel 文件。
    检查交互项 p(X×Z→Y) 是否 ≤ 阈值。
    满足条件的整行数据（保留所有列）保存到汇总文件，添加 y_var 和 file_name 列。

    参数：
        output_dir : str
            主输出文件夹路径
        p_threshold : float
            显著性阈值，默认 0.05
        summary_name : str
            汇总文件名，默认 summary_significant_moderation.xlsx
    """
    summary_records = []

    # 遍历每个因变量子文件夹
    for y_var in os.listdir(output_dir):
        subdir = os.path.join(output_dir, y_var)
        if not os.path.isdir(subdir):
            continue

        # 筛选文件名包含 "_moderate" 或 "_moderation"
        xlsx_files = [f for f in os.listdir(subdir) if f.endswith('.xlsx') and ('_moderate' in f.lower() or '_moderation' in f.lower())]
        print(f"\n📂 正在处理因变量 {y_var} 下的调节文件，共 {len(xlsx_files)} 个")

        for fname in xlsx_files:
            fpath = os.path.join(subdir, fname)
            try:
                df = pd.read_excel(fpath)
                if df.empty:
                    continue

                # 检查交互项 p 列是否存在
                inter_p_col = 'p(X×Z→Y)'
                if inter_p_col not in df.columns:
                    print(f"⚠️ 文件 {fname} 缺少列 {inter_p_col}，跳过")
                    continue

                # 判断交互项显著
                sig_rows = df[df[inter_p_col] <= p_threshold]
                if not sig_rows.empty:
                    sig_rows = sig_rows.copy()
                    # 在最前面添加 y_var 和 file_name
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
        print("\n⚠️ 未找到满足条件（交互项显著）的结果。")