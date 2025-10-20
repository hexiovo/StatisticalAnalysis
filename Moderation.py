import pandas as pd
import statsmodels.formula.api as smf
import os
import numpy as np

def moderation_search(file_path, x_var, y_var, exclude_cols=None, output_dir=None):
    """
    自动测试Excel中每个变量作为调节变量（Z）的显著性（交互项p值）。
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

    # ---------- 4. 候选变量 ----------
    candidates = [c for c in df.columns if c not in [x_var, y_var]]
    results = []

    # ---------- 5. 调节分析 ----------
    for z in candidates:
        try:
            # 回归模型：Y ~ X + Z + X*Z
            model = smf.ols(f"{y_var} ~ {x_var} * {z}", data=df).fit()

            # 交互项名称在 pvalues 字典中一般为 “X:Z”
            interaction = f"{x_var}:{z}"
            p_inter = model.pvalues.get(interaction, np.nan)

        except Exception as e:
            print(f"变量 {z} 出错：{e}")
            p_inter = np.nan  # 若回归出错则设为NaN

        results.append({
            "调节变量": z,
            "交互项p值": p_inter
        })

    # ---------- 6. 输出 ----------
    df_result = pd.DataFrame(results)
    if output_dir is None:
        output_dir = os.path.dirname(file_path)
    os.makedirs(output_dir, exist_ok=True)

    base_name = os.path.splitext(os.path.basename(file_path))[0]
    save_path = os.path.join(output_dir, f"{base_name}_moderation.xlsx")

    df_result.to_excel(save_path, index=False)
    print(f"调节分析结果已保存至：{save_path}")
