from Pipeline import *
from misc import *
import os
import time

# ========================= 主程序入口 =========================
if __name__ == "__main__":


    # ---------- ModelSearch ----------
    # dv_list = ["同伴观点采择倾向", "同伴影响"]   # 因变量列表
    # group_col = "组别"                           # 分组变量
    #
    # # ---------- 输入输出路径 ----------
    # input_dir = r"F:\Project\AI-Group\data\Pre&1A\Model_Search"
    # task_names = ["CT", "AUT", "PIT"]  # 三个任务名
    #
    # model_search_pipeline(
    #     task_names = task_names,
    #     dv_list = dv_list,
    #     group_col = group_col,
    #     input_dir=input_dir
    # )


    # ------- convert sav to xlsx ---------
    # input_dir = r"F:\Project\AI-Group\data\Pre&1A\all\SPSS"  # 输入文件夹路径
    # output_dir = r"F:\Project\AI-Group\data\Pre&1A\all\ALL"  # 输出文件夹路径
    # convert_sav_to_xlsx(input_dir, output_dir)


    # --- Mediation & Moderation Search ---
    x_vals = "组别"
    y_vals = ["新颖性变化","同伴观点采择倾向","适用性变化"]
    exclude_cols = ["AI拟人化","序号","姓名"]
    inputdir = r"F:\Project\AI-Group\data\Pre&1A\all\ALL"
    outputdir = r"F:\Project\AI-Group\data\Pre&1A\all\统计\中介调节"

    mediation_moderation_pipeline(
        input_dir=inputdir,
        x_var=x_vals,
        y_var=y_vals,
        exclude_cols=exclude_cols,
        output_dir=outputdir
    )