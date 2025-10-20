import pandas as pd
import statsmodels.formula.api as smf
import os
import numpy as np

def mediation_search(file_path, x_var, y_var, exclude_cols=None, output_dir=None):
    """
    自动测试Excel中每个变量作为中介变量（M）的显著性，输出p值结果。
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
            # ---------- a 路径 ----------
            model_a = smf.ols(f"{m} ~ {x_var}", data=df).fit()
            p_a = model_a.pvalues.get(x_var, np.nan)  # 若不存在则为 NaN

            # ---------- b 路径 ----------
            model_b = smf.ols(f"{y_var} ~ {x_var} + {m}", data=df).fit()
            p_b = model_b.pvalues.get(m, np.nan)

        except Exception as e:
            print(f"变量 {m} 出错：{e}")
            p_a, p_b = np.nan, np.nan  # 如果任意步骤失败，赋值 NaN

        results.append({
            "中介变量": m,
            "p(X→M)": p_a,
            "p(M→Y)": p_b
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
