import os
import pandas as pd
import pyreadstat

def convert_sav_to_xlsx(input_dir, output_dir):
    """
    批量将SPSS格式（.sav）的数据文件转换为Excel（.xlsx），第一行为变量名。
    参数:
        input_dir: str 输入文件夹路径
        output_dir: str 输出文件夹路径
    """
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # 遍历目录下所有文件
    for root, _, files in os.walk(input_dir):
        for file in files:
            if file.endswith('.sav'):
                sav_path = os.path.join(root, file)
                rel_path = os.path.relpath(root, input_dir)
                out_dir = os.path.join(output_dir, rel_path)
                os.makedirs(out_dir, exist_ok=True)

                out_path = os.path.join(out_dir, os.path.splitext(file)[0] + '.xlsx')

                print(f"正在转换: {sav_path}")

                try:
                    # 读取SPSS文件
                    df, meta = pyreadstat.read_sav(sav_path)

                    # 若有变量标签，优先使用标签名
                    var_labels = meta.column_labels
                    if any(var_labels):  # 存在变量标签时
                        df.columns = [label if label else name for label, name in zip(var_labels, df.columns)]

                    # 写入Excel文件
                    df.to_excel(out_path, index=False)
                    print(f"✅ 已保存: {out_path}")

                except Exception as e:
                    print(f"❌ 转换失败: {sav_path}")
                    print(f"   错误原因: {e}")

    print("全部文件转换完成。")
