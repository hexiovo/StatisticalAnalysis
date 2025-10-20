from Statistic import model_significance_search
import os
import time

def model_search_pipeline(
        input_dir: str,
        task_names: object = None,
        dv_list: object = None,
        group_col: object = "组别",
        exclude_cols: object = None,
        glm_family: object = "gaussian",
        alpha: object = 0.05
) -> dict:
    """
    批量运行多个任务文件的模型显著性搜索。

    Parameters
    ----------
    input_dir : str, optional

    task_names : list of str, optional
        要处理的任务名列表（每个任务应对应一个 Excel 文件）。默认 ["CT", "AUT", "PIT"]。

    dv_list : list of str, optional
        因变量（Dependent Variables）名称列表。默认 ["同伴观点采择倾向", "同伴影响"]。

    group_col : str, optional
        分组变量列名，默认 "组别"。

    exclude_cols : list or None, optional
        要排除的列名列表。默认 None（不排除任何列）。

    glm_family : str, optional
        GLM 模型的分布族，可选 "gaussian"、"binomial" 等。默认 "gaussian"。

    alpha : float, optional
        显著性检验阈值，默认 0.05。

    model_func : callable, required
        模型函数，用于实际执行模型搜索。例如 `model_significance_search`。
        函数应接受以下参数：
        (file_path, dv_list, group_col, exclude_cols, glm_family, alpha,
         save_folder, total_sequences, current_sequence, past_dv_times,
         past_seq_times, start_time_all)

    Returns
    -------
    results_summary : dict
        返回各任务的结果、耗时统计等汇总信息。
    """

    # ---------- 计时初始化 ----------
    total_sequences = len(task_names)
    past_dv_times = []
    past_seq_times = []
    start_time_all = time.time()

    results_summary = {}

    print(f"📊 开始批量任务，共 {total_sequences} 个任务。\n")

    # ---------- 循环运行 ----------
    for idx, task_name in enumerate(task_names, start=1):
        input_path = os.path.join(input_dir, f"{task_name}.xlsx")
        save_folder = os.path.join(input_dir, f"{task_name}_model")

        if not os.path.exists(input_path):
            print(f"⚠️ 跳过任务 {task_name}：输入文件 {input_path} 不存在")
            continue

        print(f"\n🚀 开始处理任务 {task_name} ({idx}/{total_sequences}) ...")



        # ---------- 调用核心模型 ----------
        results, dv_times, seq_time = model_significance_search(
            file_path=input_path,
            dv_list=dv_list,
            group_col=group_col,
            exclude_cols=exclude_cols,
            glm_family=glm_family,
            alpha=alpha,
            save_folder=save_folder,
            total_sequences=total_sequences,
            current_sequence=idx,
            past_dv_times=past_dv_times,
            past_seq_times=past_seq_times,
            start_time_all=start_time_all
        )

        # ---------- 更新统计 ----------
        past_dv_times.extend(dv_times)
        past_seq_times.append(seq_time)

        results_summary[task_name] = {
            "results": results,
            "dv_times": dv_times,
            "seq_time": seq_time
        }

    # ---------- 总耗时 ----------
    total_time = time.time() - start_time_all
    print(f"\n🎉 所有任务完成，总耗时 {total_time / 60:.1f} 分钟")

    results_summary["total_time_min"] = round(total_time / 60, 2)
    results_summary["past_dv_times"] = past_dv_times
    results_summary["past_seq_times"] = past_seq_times

    return results_summary
