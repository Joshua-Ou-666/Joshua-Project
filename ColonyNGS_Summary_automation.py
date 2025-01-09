# Created by OuJiaYu. Version 0.1 2025/01/09

import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog

# 读取原始Excel文件并进行基本的修改
def sheet_modification():
    # 使用tkinter选择文件
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls;*.xlsx")])
    
    if not file_path:
        print("未选择文件")
        return

    # 读取原始Excel文件
    df = pd.read_excel(file_path, engine='openpyxl')

    # 检查df的'Gene_Mfg_ID'列是否存在';'号
    if df['Gene_Mfg_ID'].str.contains(';').any():
        print("原始表格尚未正确整理, Gene_Mfg_ID列存在异常数值, 请检查和校正, 程序结束")
        exit()
    else:
        print("原始表格录入正确, 正在处理\n")

    # 删除列'Definition', 'selected', 'True_clone_number'
    df = df.drop(columns=['Definition', 'selected', 'True_clone_number'])

    # 增加新的表头
    new_columns = [
        'Dynamic_Depth', 'Dynamic_depth_Judgement', 'Lowest_MAP_Judgement', 
        'Large_Indel_Judgement', 'GATK_Judgement', 'Clone_Definition', 
        'Clone_Definition_modified (此列用于人工校正)', 'Manual_Selected','ANY_Clone_is_TRUE/CONCERN', 'Note'
    ]
    # 删去这些列，留存：'TRUE_Number', 'CONCERN_Number', 'Picking_Clone_Number', 'ErrFree_rate', 'CONCERN_rate'

    # 为新列赋予默认值
    for col in new_columns:
        df[col] = ''

    # 修改后的列顺序
    fixed_columns = ['Gene_Mfg_ID', 'PRJ', 'Inquiry_ID', 'GeneName', 'VectorID', 'NC_Length', 'Mfg_ID_Abbr', 'Clone#', 'Clone_Plate', 
                     'Clone_Position', 'Reformat_Plate', 'Reformat_Position', 'I7', 'I5', 'PLBC', 'Sample_Name', 'Ref_id', 'Ref_full_len', 
                     'Ref_analysis_len', 'highest_%MAP', 'lowest_%MAP', 'q10_%MAP', 'nbases>=99%MAP', 'nbases>=95%MAP', 'nbases>=90%MAP', 
                     'nbases_failed(0)', 'pos_failed(0)', 'nbases_failed(50)', 'pos_failed(50)', 'nbases_failed(90)', 'pos_failed(90)', 
                     'fq_name', '#rname', 'startpos', 'endpos', 'numreads', 'meandepth', 'meanbaseq', 'meanmapq', 'I7_reads', 'mapped_rate', 
                     'ratio', 'site_3', 'site_4', 'site_5', 'site_6', 'site_7', 'site_8', 'site_9', 'site_10', 'site_11', 'site_12', 'site_13', 
                     'site_14', 'site_15', 'site_16', 'site_17', 'fwd_ratio1', 'site_22', 'site_23', 'site_24', 'site_25', 'site_26', 'site_27', 
                     'site_28', 'site_29', 'site_30', 'site_31', 'site_32', 'site_33', 'site_34', 'site_35', 'site_36', 'rev_ratio1', 'file_reads', 
                     'flagstat', 'minimum_depth', 'median_depth', 'Dynamic_Depth', 'lowest_MAP', 'site_1', 'site_2', 'fwd_ratio2', 'site_20', 'site_21', 
                     'rev_ratio2', 'GATK_INDEL', 'GATK_SNP', 'GATK_VCF_AD', 'GATK_VCF_DP', 'GATK_AD/DP', 'Dynamic_depth_Judgement', 'Lowest_MAP_Judgement', 
                     'Large_Indel_Judgement', 'GATK_Judgement', 'Clone_Definition', 'Clone_Definition_modified (此列用于人工校正)', 
                     'Manual_Selected', 'ANY_Clone_is_TRUE/CONCERN', 'Note'
    ]
    # 删去这些列，留存：'TRUE_Number', 'CONCERN_Number', 'Picking_Clone_Number', 'ErrFree_rate', 'CONCERN_rate'
    
    # 按照固定的列顺序重新排列
    df = df[fixed_columns]

    # 删除不需要的列
    columns_to_remove = ['Ref_id', 'fq_name', '#rname', 'I7_reads', 'site_11', 'site_12', 'site_13', 
                         'site_14', 'site_15', 'site_16', 'site_17', 'fwd_ratio1', 'site_30', 'site_31', 
                         'site_32', 'site_33', 'site_34', 'site_35', 'site_36', 'rev_ratio1'
    ] 
    df = df.drop(columns=columns_to_remove)

    # 获取原始文件的目录
    output_dir = os.path.dirname(file_path)

    #获取原始文件的文件名
    input_file_name = os.path.basename(file_path)

    return df, output_dir, input_file_name

# 进行judgement的添加    
def judgement_addition():
    
    print( "请输入soft clip count的数值, 默认请输入50: ")
    soft_clip_count = float(input())
    print( "请输入large indel ratio的值, 默认请输入5.26: ")
    large_indel_ratio = float(input())

    print("开始添加judgement,请耐心等候\n")

    ## 进行Dynamic_Depth和Dynamic_depth_Judgement的判断
    # 为Dynamic_Depth列赋值，公式为Dynamic_Depth=0.1 * Median_Depth
    df['Dynamic_Depth'] = df['median_depth'] * 0.1
    # 当minimum_depth 大于 Dynamic_Depth时，在Dynamic_depth_Judgement列输出“TRUE”，反之则为“FALSE"
    df['Dynamic_depth_Judgement'] = df.apply(lambda row: 'TRUE' if row['minimum_depth'] > row['Dynamic_Depth'] else 'FALSE', axis=1)

    ## 进行Lowest_MAP_Judgement的判断
    # 当lowest_MAP 介于80-90时，在Lowest_MAP_Judgement (90)列输出“Lowest_MAP_CONCERN”, 大于等于90时输出“TRUE”，反之则为“FALSE"
    df['Lowest_MAP_Judgement'] = df.apply(lambda row: 'Lowest_MAP_CONCERN' if row['lowest_MAP'] >= 80 else 'FALSE', axis=1)
    df['Lowest_MAP_Judgement'] = df.apply(lambda row: 'TRUE' if row['lowest_MAP'] >= 90 else row['Lowest_MAP_Judgement'], axis=1)

    ## 进行Large Indel Judgement的判断
    # 当site_2 小于等于 50 且 for_ratio2 小于5.26 且 site21 小于等于 50 且 rev_ratio2 小于 5.26时，在Large Indel Judgement (50&5.26)列输出“TRUE”，反之则为“FALSE"
    df['Large_Indel_Judgement'] = df.apply(lambda row: 'TRUE' if row['site_2'] <= soft_clip_count and row['fwd_ratio2'] < large_indel_ratio and row['site_21'] <= soft_clip_count and row['rev_ratio2'] < large_indel_ratio else 'FALSE', axis=1) 
    # 特别地, 当 for_ratio2 大于 5.26 但 rev_ratio2 小于 5.26时，在Large Indel Judgement (50&5.26)列输出“LARGE_INDEL_FOR”，反之则输出原值 
    df['Large_Indel_Judgement'] = df.apply(lambda row: 'LARGE_INDEL_FOR' if row['fwd_ratio2'] > large_indel_ratio and row['rev_ratio2'] < large_indel_ratio else row['Large_Indel_Judgement'], axis=1)
    # 特别地, 当 for_ratio2 小于 5.26 但 rev_ratio2 大于 5.26时，在Large Indel Judgement (50&5.26)列输出“LARGE_INDEL_REV”，反之则输出原值 
    df['Large_Indel_Judgement'] = df.apply(lambda row: 'LARGE_INDEL_REV' if row['fwd_ratio2'] < large_indel_ratio and row['rev_ratio2'] > large_indel_ratio else row['Large_Indel_Judgement'], axis=1)
    # 特别地, 当site_2 且 site_21 且 fwd_ratio2 且 rev_ratio2 都为NaN时，输出“TRUE”, 反之则输出原值
    df['Large_Indel_Judgement'] = df.apply(lambda row: 'TRUE' if pd.isnull(row['site_2']) and pd.isnull(row['site_21']) and pd.isnull(row['fwd_ratio2']) and pd.isnull(row['rev_ratio2']) else row['Large_Indel_Judgement'], axis=1)

    ## 进行GATK_Judgement的判断
    # 当GATK_INDEL 且 GATK_SNP 且 GATK_AD/DP 全为NaN时，在GATK_Judgement列输出“GATK_TRUE", 反之则输出“GATK_FALSE”
    df['GATK_Judgement'] = df.apply(lambda row: 'GATK_TRUE' if pd.isnull(row['GATK_INDEL']) and pd.isnull(row['GATK_SNP']) and pd.isnull(row['GATK_AD/DP']) else 'GATK_FALSE', axis=1)
    
    # 处理GATK分析的复合数据
    # 当GATK_INDEL不为空, 且GATK_SNP为空时，对GATK_INDEL和GATK_AD/DP进行处理
    def process_composite_data_INDEL_only(row):
        if pd.notnull(row['GATK_INDEL']) and pd.isnull(row['GATK_SNP']) and pd.notnull(row['GATK_AD/DP']):
            indel_values = str(row['GATK_INDEL']).split(';') if isinstance(row['GATK_INDEL'], str) else [row['GATK_INDEL']]
            ad_dp_values = str(row['GATK_AD/DP']).split(';') if isinstance(row['GATK_AD/DP'], str) else [row['GATK_AD/DP']]

            # 确保两个列表的长度相同
            if len(indel_values) == len(ad_dp_values):
                combined_values = list(zip(indel_values, ad_dp_values))
            else:
                combined_values = []
                
            # print(combined_values)
            indel_false = any(float(ad_dp) < 0.65 for indel, ad_dp in combined_values)
            indel_concern = all(float(ad_dp) >= 0.65 for indel, ad_dp in combined_values)

            # Check INDEL values using combined_values
            if indel_false:
                return 'GATK_INDEL_FALSE'
            elif indel_concern:
                return 'GATK_INDEL_CONCERN'
            else:
                return row['GATK_Judgement']
        else:
            return row['GATK_Judgement']
            #print (row['GATK_Judgement'])
    #运行INDEL_only函数并赋值给GATK_Judgement列        
    df['GATK_Judgement'] = df.apply(process_composite_data_INDEL_only, axis=1)

    # 当GATK_INDEL为空, 且GATK_SNP不为空时，对GATK_SNP和GATK_AD/DP进行处理
    def process_composite_data_SNP_only(row):
        if pd.isnull(row['GATK_INDEL']) and pd.notnull(row['GATK_SNP']) and pd.notnull(row['GATK_AD/DP']):
            snp_values = str(row['GATK_SNP']).split(';') if isinstance(row['GATK_SNP'], str) else [row['GATK_SNP']]
            ad_dp_values = str(row['GATK_AD/DP']).split(';') if isinstance(row['GATK_AD/DP'], str) else [row['GATK_AD/DP']]
            
            # 确保两个列表的长度相同
            if len(snp_values) == len(ad_dp_values):
                combined_values = list(zip(snp_values, ad_dp_values))
            else:
                combined_values = []
            
            # print(combined_values)

            snp_false = any(float(ad_dp) < 0.8 for snp, ad_dp in combined_values)
            snp_concern = all(0.8 <= float(ad_dp) < 0.9 for snp, ad_dp in combined_values)
            snp_true = all(float(ad_dp) >= 0.9 for snp, ad_dp in combined_values)
            
            if snp_false:
                return 'GATK_SNP_FALSE'
            elif snp_concern:
                return 'GATK_SNP_CONCERN'
            elif snp_true:
                return 'GATK_SNP_TRUE'
            else:
                return row['GATK_Judgement']
        else:
            return row['GATK_Judgement']
    #运行SNP_only函数并赋值给GATK_Judgement列
    df['GATK_Judgement'] = df.apply(process_composite_data_SNP_only, axis=1)

    # 当GATK_INDEL不为空, 且GATK_SNP不为空时，对GATK_INDEL, GATK_SNP和GATK_AD/DP进行处理
    def process_composite_data_BOTH(row):
        if pd.notnull(row['GATK_INDEL']) and pd.notnull(row['GATK_SNP']) and pd.notnull(row['GATK_AD/DP']):
            indel_values = str(row['GATK_INDEL']).split(';') if isinstance(row['GATK_INDEL'], str) else [row['GATK_INDEL']]
            snp_values = str(row['GATK_SNP']).split(';') if isinstance(row['GATK_SNP'], str) else [row['GATK_SNP']]
            ad_dp_values = str(row['GATK_AD/DP']).split(';') if isinstance(row['GATK_AD/DP'], str) else [row['GATK_AD/DP']]
            
            # Split ad_dp_values into two parts
            ad_dp_values_indel = ad_dp_values[:len(indel_values)]
            ad_dp_values_snp = ad_dp_values[len(indel_values):]

            # Pair indel_values with ad_dp_values_indel
            combined_values_indel = list(zip(indel_values, ad_dp_values_indel))

            # Pair snp_values with ad_dp_values_snp
            combined_values_snp = list(zip(snp_values, ad_dp_values_snp))

            # Check INDEL values
            indel_all_above_0_65 = all(float(ad_dp) >= 0.65 for indel, ad_dp in combined_values_indel)
            indel_any_below_0_65 = any(float(ad_dp) < 0.65 for indel, ad_dp in combined_values_indel)
            
            # Check SNP values
            snp_all_above_0_9 = all(float(ad_dp) > 0.9 for snp, ad_dp in combined_values_snp)
            snp_all_above_0_8 = all(float(ad_dp) >= 0.8 for snp, ad_dp in combined_values_snp)
            snp_any_below_0_8 = any(float(ad_dp) < 0.8 for snp, ad_dp in combined_values_snp)
            
            if indel_all_above_0_65 and snp_all_above_0_9:
                return 'GATK_INDEL_CONCERN'
            elif indel_all_above_0_65 and snp_all_above_0_8 and not snp_all_above_0_9:
                return 'GATK_INDEL&SNP_CONCERN'
            elif indel_all_above_0_65 and snp_any_below_0_8:
                return 'GATK_SNP_FALSE'
            elif indel_any_below_0_65 and snp_all_above_0_8:
                return 'GATK_INDEL_FALSE'
            elif indel_any_below_0_65 and snp_any_below_0_8:
                return 'GATK_FALSE'
            else:
                return row['GATK_Judgement']
        else: 
            return row['GATK_Judgement']
    #运行BOTH函数并赋值给GATK_Judgement列
    df['GATK_Judgement'] = df.apply(process_composite_data_BOTH, axis=1)

    # 保存增加judment的表格为new_judgment.xls
    # new_output_file_path = os.path.join(output_dir, 'judgment.xls')
    # df.to_excel(new_output_file_path, index=False)
    # print(f"增加judgment判断表格已保存为: {new_output_file_path}")
    # return df, output_dir
    return df, output_dir, input_file_name

# 进行克隆判断
def clone_judgement():
    print ("开始克隆判断,请耐心等候")

    # 初始化'Clone_Definition'列
    df['Clone_Definition'] = 'To_be_defined'

    # 当Dynamic_depth_Judgement 且 Lowest_MAP_Judgement 且 Large_Indel_Judgement 且 GATK_Judgement 全为“TRUE”时，在Clone_Definition列输出“TRUE”，反之则为“DISCARD”
    df['Clone_Definition'] = df.apply(lambda row: 'TRUE' if row['Dynamic_depth_Judgement'] == 'TRUE' 
                                               and row['Lowest_MAP_Judgement'] == 'TRUE' 
                                               and row['Large_Indel_Judgement'] == 'TRUE' 
                                               and row['GATK_Judgement'] == 'GATK_TRUE' 
                                               else 'DISCARD', axis=1)

    # 当Dynamic_depth_Judgement 且 Large_Indel_Judgement 且 GATK_Judgement 全为“TRUE”, 且Lowest_MAP为 "Lowest_MAP_CONCERN",在Clone_Definition列输出“Lowest_MAP_CONCERN”，反之则输出原值
    df['Clone_Definition'] = df.apply(lambda row: 'Lowest_MAP_CONCERN' if row['Dynamic_depth_Judgement'] == 'TRUE' 
                                               and row['Lowest_MAP_Judgement'] == 'Lowest_MAP_CONCERN' 
                                               and row['Large_Indel_Judgement'] == 'TRUE' 
                                               and row['GATK_Judgement'] == 'GATK_TRUE' 
                                               else row['Clone_Definition'], axis=1)
    
    # 当Dynamic_depth_Judgement 且 Lowest_MAP 且 GATK_Judgement 全为“TRUE”, 且Large Indel Judgement (50&5.26) 为 "LARGE_INDEL_FOR"或"LARGE_INDEL_REV",在Clone_Definition列输出“LARGE_INDEL_CONCERN”，反之则输出原值
    df['Clone_Definition'] = df.apply(lambda row: 'LARGE_INDEL_CONCERN' if row['Dynamic_depth_Judgement'] == 'TRUE'
                                               and row['Lowest_MAP_Judgement'] == 'TRUE'
                                               and row['GATK_Judgement'] == 'GATK_TRUE'
                                               and (row['Large_Indel_Judgement'] == 'LARGE_INDEL_FOR' or row['Large_Indel_Judgement'] == 'LARGE_INDEL_REV')
                                               else row['Clone_Definition'], axis=1)
    
    # 当Dynamic_depth_Judgement 且 Lowest_MAP_Judgement 且 Large_Indel_Judgement 且 GATK_Judgement 为“GATK_INDEL_CONCERN"时，在Clone_Definition列输出“GATK_INDEL_CONCERN”，反之则输出原值
    df['Clone_Definition'] = df.apply(lambda row: 'GATK_INDEL_CONCERN' if row['Dynamic_depth_Judgement'] == 'TRUE' 
                                               and row['Lowest_MAP_Judgement'] == 'TRUE' 
                                               and row['Large_Indel_Judgement' ] == 'TRUE' 
                                               and row['GATK_Judgement'] == 'GATK_INDEL_CONCERN' 
                                               else row['Clone_Definition'], axis=1)
    
    # 当Dynamic_depth_Judgement 且 Lowest_MAP_Judgement 且 Large_Indel_Judgement 且 GATK_Judgement 为“GATK_SNP_CONCERN"时，在Clone_Definition列输出“GATK_SNP_CONCERN”，反之则输出原值
    df['Clone_Definition'] = df.apply(lambda row: 'GATK_SNP_CONCERN' if row['Dynamic_depth_Judgement'] == 'TRUE' 
                                               and row['Lowest_MAP_Judgement'] == 'TRUE' 
                                               and row['Large_Indel_Judgement' ] == 'TRUE' 
                                               and row['GATK_Judgement'] == 'GATK_SNP_CONCERN' 
                                               else row['Clone_Definition'], axis=1)
    
    # 当Dynamic_depth_Judgement 且 Large Indel Judgement为"TRUE", 且Lowest_MAP_Judgement为 "Lowest_MAP_CONCERN" 且 GATK_Judgement 为“GATK_SNP_CONCERN"时，在Clone_Definition列输出“Lowest_MAP&SNP_CONCERN”，反之则输出原值
    df['Clone_Definition'] = df.apply(lambda row: 'Lowest_MAP&SNP_CONCERN' if row['Dynamic_depth_Judgement'] == 'TRUE' 
                                               and row['Lowest_MAP_Judgement'] == 'Lowest_MAP_CONCERN' 
                                               and row['Large_Indel_Judgement' ] == 'TRUE' 
                                               and row['GATK_Judgement'] == 'GATK_SNP_CONCERN' 
                                               else row['Clone_Definition'], axis=1)

    # 将Clone_Definition列的值赋给Clone_Definition_modified (此列用于人工校正)列
    df['Clone_Definition_modified (此列用于人工校正)'] = df['Clone_Definition']

    # 获取新的文件名
    new_file_name = os.path.splitext(input_file_name)[0] + '_Clone_Judgement.xlsx'
    
    # 保存增加克隆判断的表格为新的文件名
    new_output_file_path = os.path.join(output_dir, new_file_name)
    df.to_excel(new_output_file_path, index=False)
    print(f"增加clone judgement判断表格已保存为: {new_output_file_path}")

# 进行统计
def statistics():
    def load_files():
        # 使用tkinter选择模板文件
        root = tk.Tk()
        root.withdraw()  # 隐藏主窗口
        print("请选择统计模板文件")
        template_file_path = filedialog.askopenfilename(title="选择统计模板文件", filetypes=[("Excel files", "*.xls;*.xlsx")])
    
        if not template_file_path:
            print("未选择统计模板文件")
            return None, None

        # 使用tkinter选择克隆判断文件
        print("请选择克隆判断文件")
        data_file_path = filedialog.askopenfilename(title="请选择克隆判断文件, 文件名应带有'Clone_Judgement'", filetypes=[("Excel files", "*.xls;*.xlsx")])
    
        if not data_file_path:
            print("未选择克隆判断文件")
            return None, None

        data_file_name = os.path.basename(data_file_path).replace('_Clone_Judgement', '')

        data_file_dir = os.path.dirname(data_file_path)

        return template_file_path, data_file_path, data_file_name, data_file_dir

    # 调用load_files函数获取文件路径
    template_file_path, data_file_path, data_file_name, data_file_dir = load_files()
    print("读取统计模板文件和克隆判断文件成功, 开始统计数据, 请耐心等候")

    # 读取模板文件
    template_df = pd.read_excel(template_file_path, engine='openpyxl')
    # 读取克隆判断文件
    data_df = pd.read_excel(data_file_path, engine='openpyxl')
    
    # 统计模板文件中每一个Gene_Mfg_ID在完整数据文件中出现的次数，并输出到Clone_Number列
    # 如果在完整数据文件中找不到某一个值，则输出0
    template_df['Clone_Number'] = template_df['Gene_Mfg_ID'].apply(lambda x: data_df['Gene_Mfg_ID'].tolist().count(x) if x in data_df['Gene_Mfg_ID'].tolist() else 0)

    # 统计模板文件中每一个Gene_Mfg_ID在完整数据文件的'Clone_Definition_modified (此列用于人工校正)'中出现TRUE的次数，并输出到TRUE_number列。如果在完整数据文件中找不到某一个值，则输出0
    template_df['TRUE_Number'] = template_df['Gene_Mfg_ID'].apply(lambda x: 
        data_df[(data_df['Gene_Mfg_ID'] == x) & 
                (data_df['Clone_Definition_modified (此列用于人工校正)'] == 'TRUE')].shape[0]
        if x in data_df['Gene_Mfg_ID'].tolist() else 0)
    
    # 统计模板文件中每一个Gene_Mfg_ID在完整数据文件的'Clone_Definition_modified (此列用于人工校正)'中出现Lowest_MAP_CONCERN的次数，以及出现GATK_SNP_CONCERN的次数，将两项相加，并输出到Lowest_MAP/GATK_SNP_CONCERN_number列。如果在完整数据文件中找不到某一个值，则输出0
    template_df['Lowest_MAP/GATK_SNP_CONCERN_Number'] = template_df['Gene_Mfg_ID'].apply(lambda x: 
        data_df[(data_df['Gene_Mfg_ID'] == x) & 
                (data_df['Clone_Definition_modified (此列用于人工校正)'] == 'Lowest_MAP_CONCERN')].shape[0] + 
        data_df[(data_df['Gene_Mfg_ID'] == x) & 
                (data_df['Clone_Definition_modified (此列用于人工校正)'] == 'GATK_SNP_CONCERN')].shape[0] + 
        data_df[(data_df['Gene_Mfg_ID'] == x) & 
                (data_df['Clone_Definition_modified (此列用于人工校正)'] == 'Lowest_MAP&SNP_CONCERN')].shape[0]
        if x in data_df['Gene_Mfg_ID'].tolist() else 0
    )

    # 统计模板文件中每一个Gene_Mfg_ID在完整数据文件的'Clone_Definition_modified (此列用于人工校正)'中出现LARGE_INDEL_CONCERN的次数，并输出到LARGE_INDEL_CONCERN_number列。如果在完整数据文件中找不到某一个值，则输出0
    template_df['LARGE_INDEL_CONCERN_Number'] = template_df['Gene_Mfg_ID'].apply(lambda x:
        data_df[(data_df['Gene_Mfg_ID'] == x) & 
                (data_df['Clone_Definition_modified (此列用于人工校正)'] == 'LARGE_INDEL_CONCERN')].shape[0]
        if x in data_df['Gene_Mfg_ID'].tolist() else 0
    )
    
    # 统计模板文件中每一个Gene_Mfg_ID在完整数据文件的'Clone_Definition_modified (此列用于人工校正)'中出现GATK_INDEL_CONCERN的次数，并输出到GATK_INDEL_CONCERN_number列。如果在完整数据文件中找不到某一个值，则输出0
    template_df['GATK_INDEL_CONCERN_Number'] = template_df['Gene_Mfg_ID'].apply(lambda x: 
        data_df[(data_df['Gene_Mfg_ID'] == x) & 
                (data_df['Clone_Definition_modified (此列用于人工校正)'] == 'GATK_INDEL_CONCERN')].shape[0]
        if x in data_df['Gene_Mfg_ID'].tolist() else 0
    )

    ## 基于file_reads校正clone number,获得picking clone number
    # 建立函数找出file_reads的最优cutoff
    def find_optimal_cutoff(cutoff_df):
        # 提取需要的列
        file_read_cutoff_df = cutoff_df[['Gene_Mfg_ID', 'file_reads', 'flagstat']]

        # 计算cutoff=500时flagstat < 90的数量
        max_flagstat_below_90 = file_read_cutoff_df[file_read_cutoff_df['file_reads'] <= 500]['flagstat'].lt(90).sum()
        print(f"Max_flagstat_below_90: {max_flagstat_below_90}")

        # 初始化变量
        optimal_cutoff = None

        # 设置file_reads的cutoff范围，从100到500，并以20为单位增加
        for cutoff in range(100, 501, 20):
            below_90_count = file_read_cutoff_df[file_read_cutoff_df['file_reads'] <= cutoff]['flagstat'].lt(90).sum()

            # 检查是否达到Max_flagstat_below_90的95%
            if below_90_count >= 0.95 * max_flagstat_below_90:
                optimal_cutoff = cutoff
                break

        print(f"Optimal cutoff: {optimal_cutoff}")
        return optimal_cutoff
    
    # 找到最优cutoff,统计模板文件中每一个Gene_Mfg_ID在完整数据文件的file_reads列中小于等于最优cutoff的次数，并输出到Clone_Correction列。如果在完整数据文件中找不到某一个值，则输出0
    file_read_cutoff = find_optimal_cutoff(data_df)
    template_df['Clone_Correction'] = template_df['Gene_Mfg_ID'].apply(lambda x: 
        data_df[(data_df['Gene_Mfg_ID'] == x) & 
                (data_df['file_reads'] <= file_read_cutoff)].shape[0]
        if x in data_df['Gene_Mfg_ID'].tolist() else 0
    )
    
    # 计算校正的clone number,输出到Picking_Clone_Number列
    template_df['Picking_Clone_Number'] = template_df['Clone_Number'] - template_df['Clone_Correction']

    # 计算ErrFree_Rate,输出到ErrFree_rate列。如果Picking_Clone_Number为0，则输出0
    template_df['ErrFree_Rate'] = template_df.apply(lambda row: row['TRUE_Number'] / row['Picking_Clone_Number'] if row['Picking_Clone_Number'] != 0 else 0, axis=1)

    # 计算CONCERN_rate,输出到CONCERN_rate列。如果Picking_Clone_Number为0，则输出0
    template_df['CONCERN_Rate'] = template_df.apply(lambda row: (row['Lowest_MAP/GATK_SNP_CONCERN_Number'] + row['LARGE_INDEL_CONCERN_Number'] + row['GATK_INDEL_CONCERN_Number']) / row['Picking_Clone_Number'] if row['Picking_Clone_Number'] != 0 else 0, axis=1)

    # 计算ErrFree+CONCERN_Rate,输出到ErrFree+CONCERN_Rate列
    template_df['ErrFree+CONCERN_Rate'] = template_df['ErrFree_Rate'] + template_df['CONCERN_Rate']

    # 给ANY_Clone_is_TRUE/CONCERN列赋值
    template_df['ANY_Clone_is_TRUE/CONCERN'] = template_df.apply(lambda row: 'YES' if row['ErrFree+CONCERN_Rate'] > 0 else 'NO', axis=1)
    
    # 将Clone_Correction的表头增加cutoff值
    template_df.rename(columns={'Clone_Correction': f'Clone_Correction(cutoff={file_read_cutoff})'}, inplace=True)
    
    # 获取新的文件名
    new_file_name = os.path.splitext(data_file_name)[0] + '_Statistics.xlsx'

    # 保存为新的Excel文件
    merged_output_file_path = os.path.join(data_file_dir, new_file_name)
    template_df.to_excel(merged_output_file_path, index=False)
    print(f"统计数据已保存为: {merged_output_file_path}")

    # 将统计数据输出到完整数据文件的同名列
    columns_to_update = ['ANY_Clone_is_TRUE/CONCERN']
    for col in columns_to_update:
        data_df[col] = data_df['Gene_Mfg_ID'].map(template_df.set_index('Gene_Mfg_ID')[col])

    # 获取新的文件名，保存更新后的完整数据文件
    new_data_file_name = os.path.splitext(data_file_name)[0] + '_Judgement_updated.xlsx'
    updated_data_file_path = os.path.join(data_file_dir, new_data_file_name)
    data_df.to_excel(updated_data_file_path, index=False)
    print(f"增加统计注释的克隆判断文件已保存为: {updated_data_file_path}\n")

# 进行克隆选择
def clone_selection():
    def load_files():
        # 使用tkinter选择模板文件
        root = tk.Tk()
        root.withdraw()  # 隐藏主窗口
        print("请选择完成统计注释的克隆判断文件, 文件名应带有'Clone_Judgement_updated'")
        data_file_path = filedialog.askopenfilename(title="选择完成统计注释的克隆判断文件", filetypes=[("Excel files", "*.xls;*.xlsx")])
    
        if not data_file_path:
            print("未选择完整数据文件")
            return None, None

        data_file_name = os.path.basename(data_file_path).replace('_Clone_Judgement_updated', '')

        data_file_dir = os.path.dirname(data_file_path)

        return data_file_path, data_file_name, data_file_dir

    data_file_path, data_file_name, data_file_dir = load_files()
    

    # 读取完整数据文件
    df = pd.read_excel(data_file_path, engine='openpyxl')

    #检查导入的文件是否已有克隆判断
    if 'To_be_defined' in df['Clone_Definition'].values:
        print("导入的文件中存在未定义的克隆，请检查文件内容或代码逻辑是否遍历所有克隆, 程序自动退出")
        exit()
    else:
        print("读取文件成功, 开始进行克隆选择, 请耐心等候")
    
    # 初始化Manual_Selected列
    df['Manual_Selected'] = 'FALSE'

    # 对每个Gene_Mfg_ID的克隆进行分组
    grouped = df.groupby('Gene_Mfg_ID')

    # 增加一列'Handled', 用于标记处理过的Gene_Mfg_ID组, 默认为To_Be_Done
    df['Handled'] = 'To_Be_Done'

    # 根据克隆判断结果进一步分组
    for name, group in grouped:
        # 统计每个Gene_Mfg_ID组内的克隆数量
        clone_number = len(group)
        true_clones = group[group['Clone_Definition_modified (此列用于人工校正)'] == 'TRUE']
        lowest_map_concern_clones = group[group['Clone_Definition_modified (此列用于人工校正)'] == 'Lowest_MAP_CONCERN']
        large_indel_concern_clones = group[group['Clone_Definition_modified (此列用于人工校正)'] == 'LARGE_INDEL_CONCERN']
        gatk_indel_concern_clones = group[group['Clone_Definition_modified (此列用于人工校正)'] == 'GATK_INDEL_CONCERN']
        gatk_snp_concern_clones = group[group['Clone_Definition_modified (此列用于人工校正)'] == 'GATK_SNP_CONCERN']
        lowest_map_snp_concern_clones = group[group['Clone_Definition_modified (此列用于人工校正)'] == 'Lowest_MAP&SNP_CONCERN']
        discard_clones = group[group['Clone_Definition_modified (此列用于人工校正)'] == 'DISCARD']

    # 如果存在TRUE克隆，则选择最优克隆
        if not true_clones.empty:
            # 检查是否存在median_depth > 50的TRUE克隆数量 > 1
            true_clones_above_50 = true_clones[true_clones['median_depth'] > 50]
            if len(true_clones_above_50) > 1:
                # 在median_depth > 50的TRUE克隆内部进行flagstat降序排序
                sorted_clones = true_clones_above_50.sort_values(by=['flagstat', 'lowest_MAP'], ascending=[False, False])
            else:
                # 在所有TRUE克隆内部进行flagstat降序排序
                sorted_clones = true_clones.sort_values(by=['flagstat', 'lowest_MAP'], ascending=[False, False])
            # 选择排序后的第一个克隆作为最优克隆
            best_clone_index = sorted_clones.index[0]
            df.at[best_clone_index, 'Manual_Selected'] = 'TRUE'
            # 当某一克隆的'Manual_Selected'为TRUE时, 将该克隆所在的Gene_Mfg_ID组的'Handled'列设置为Done
            df.loc[df['Gene_Mfg_ID'] == df.at[best_clone_index, 'Gene_Mfg_ID'], 'Handled'] = 'Done'
            # print(f"Gene_Mfg_ID: {df.at[best_clone_index, 'Gene_Mfg_ID']}, Selected: TRUE")
    
        # 如果Gene_Mfg_ID组内仅存在lowest_map_concern克隆，则选择最优克隆
        if true_clones.empty and not lowest_map_concern_clones.empty and large_indel_concern_clones.empty and gatk_indel_concern_clones.empty and gatk_snp_concern_clones.empty and lowest_map_snp_concern_clones.empty:
            # 检查是否存在median_depth > 50且flagstat > 98.5的克隆数量 > 1
            lowest_map_concern_clones_above_50 = lowest_map_concern_clones[(lowest_map_concern_clones['median_depth'] > 50) & (lowest_map_concern_clones['flagstat'] > 98.5)]
            if len(lowest_map_concern_clones_above_50) > 1:
                # 在median_depth > 50的Lowest_MAP_CONCERN克隆内部进行flagstat降序排序
                sorted_clones = lowest_map_concern_clones_above_50.sort_values(by=['lowest_MAP'], ascending=(False))
            else:
                # 在所有Lowest_MAP_CONCERN克隆内部进行flagstat降序排序
                sorted_clones = lowest_map_concern_clones.sort_values(by=['lowest_MAP'], ascending=False)
            best_clone_index = sorted_clones.index[0]
            df.at[best_clone_index, 'Manual_Selected'] = 'Lowest_MAP_CONCERN_Selected'
            # 当某一克隆的'Manual_Selected'为Lowest_MAP_CONCERN_Selected时, 将该克隆所在的Gene_Mfg_ID组的'Handled'列设置为Done，在Note列中注释'需要人工检查'
            df.loc[df['Gene_Mfg_ID'] == df.at[best_clone_index, 'Gene_Mfg_ID'], 'Handled'] = 'Done'
            df.loc[df['Gene_Mfg_ID'] == df.at[best_clone_index, 'Gene_Mfg_ID'], 'Note'] = '需要人工检查'
            # print(f"Gene_Mfg_ID: {df.at[best_clone_index, 'Gene_Mfg_ID']}, Selected: Lowest_MAP_CONCERN_Selected")

        # 如果Gene_Mfg_ID组内仅存在GATK_INDEL_CONCERN克隆，则选择最优克隆
        if true_clones.empty and lowest_map_concern_clones.empty and large_indel_concern_clones.empty and not gatk_indel_concern_clones.empty and gatk_snp_concern_clones.empty and lowest_map_snp_concern_clones.empty:
            # 检查是否存在median_depth > 50且flagstat > 98.5的克隆数量 > 1
            gatk_indel_concern_clones_above_50 = gatk_indel_concern_clones[(gatk_indel_concern_clones['median_depth'] > 50) & (gatk_indel_concern_clones['flagstat'] > 98.5)]
            if len(gatk_indel_concern_clones_above_50) > 1:
                # 在median_depth > 50的GATK_INDEL_CONCERN克隆内部进行flagstat降序排序
                sorted_clones = gatk_indel_concern_clones_above_50.sort_values(by=['GATK_AD/DP'], ascending=[False])
            else:
                # 在所有GATK_INDEL_CONCERN克隆内部进行flagstat降序排序
                sorted_clones = gatk_indel_concern_clones.sort_values(by=['GATK_AD/DP'], ascending=[False])   
            best_clone_index = sorted_clones.index[0]
            df.at[best_clone_index, 'Manual_Selected'] = 'GATK_INDEL_CONCERN_Selected'
            # 当某一克隆的'Manual_Selected'为GATK_INDEL_CONCERN_Selected时, 将该克隆所在的Gene_Mfg_ID组的'Handled'列设置为Done，在Note列中注释'需要人工检查'
            df.loc[df['Gene_Mfg_ID'] == df.at[best_clone_index, 'Gene_Mfg_ID'], 'Handled'] = 'Done'
            df.loc[df['Gene_Mfg_ID'] == df.at[best_clone_index, 'Gene_Mfg_ID'], 'Note'] = '需要人工检查'
            # print(f"Gene_Mfg_ID: {df.at[best_clone_index, 'Gene_Mfg_ID']}, Selected: GATK_INDEL_CONCERN_Selected")

        # 如果Gene_Mfg_ID组内仅存在GATK_SNP_CONCERN克隆，则选择最优克隆
        if true_clones.empty and lowest_map_concern_clones.empty and large_indel_concern_clones.empty and gatk_indel_concern_clones.empty and not gatk_snp_concern_clones.empty and lowest_map_snp_concern_clones.empty:
            # 检查是否存在median_depth > 50且flagstat > 98.5的克隆数量 > 1
            gatk_snp_concern_clones_above_50 = gatk_snp_concern_clones[(gatk_snp_concern_clones['median_depth'] > 50 )& (gatk_snp_concern_clones['flagstat'] > 98.5)]
            if len(gatk_snp_concern_clones_above_50) > 1:
                # 在median_depth > 50的GATK_SNP_CONCERN克隆内部进行flagstat降序排序
                sorted_clones = gatk_snp_concern_clones_above_50.sort_values(by=['GATK_AD/DP'], ascending=[False])
            else:
                # 在所有GATK_SNP_CONCERN克隆内部进行flagstat降序排序
                sorted_clones = gatk_snp_concern_clones.sort_values(by=['GATK_AD/DP'], ascending=[False])
            best_clone_index = sorted_clones.index[0]
            df.at[best_clone_index, 'Manual_Selected'] = 'GATK_SNP_CONCERN_Selected'
            # 当某一克隆的'Manual_Selected'为GATK_SNP_CONCERN_Selected时, 将该克隆所在的Gene_Mfg_ID组的'Handled'列设置为Done，在Note列中注释'需要人工检查'
            df.loc[df['Gene_Mfg_ID'] == df.at[best_clone_index, 'Gene_Mfg_ID'], 'Handled'] = 'Done'
            df.loc[df['Gene_Mfg_ID'] == df.at[best_clone_index, 'Gene_Mfg_ID'], 'Note'] = '需要人工检查'
            #　print(f"Gene_Mfg_ID: {df.at[best_clone_index, 'Gene_Mfg_ID']}, Selected: GATK_SNP_CONCERN_Selected")
    
        # 如果Gene_Mfg_ID组内仅存在Lowest_MAP&SNP_CONCERN克隆，则选择最优克隆
        if true_clones.empty and lowest_map_concern_clones.empty and large_indel_concern_clones.empty and gatk_indel_concern_clones.empty and gatk_snp_concern_clones.empty and not lowest_map_snp_concern_clones.empty:
            # 检查是否存在median_depth > 50且flagstat > 98.5的克隆数量 > 1
            lowest_map_snp_concern_clones_above_50 = lowest_map_snp_concern_clones[(lowest_map_snp_concern_clones['median_depth'] > 50 ) & (lowest_map_snp_concern_clones['flagstat'] > 98.5)]
            if len(lowest_map_snp_concern_clones_above_50) > 1:
                # 在median_depth > 50的Lowest_MAP&SNP_CONCERN克隆内部进行flagstat降序排序
                sorted_clones = lowest_map_snp_concern_clones_above_50.sort_values(by=['lowest_MAP'], ascending=[False])
            else:
                # 在所有Lowest_MAP&SNP_CONCERN克隆内部进行flagstat降序排序
                sorted_clones = lowest_map_snp_concern_clones.sort_values(by=['lowest_MAP'], ascending=[False])
            best_clone_index = sorted_clones.index[0]
            df.at[best_clone_index, 'Manual_Selected'] = 'Lowest_MAP&SNP_CONCERN_Selected'
            # 当某一克隆的'Manual_Selected'为Lowest_MAP&SNP_CONCERN_Selected时, 将该克隆所在的Gene_Mfg_ID组的'Handled'列设置为Done，在Note列中注释'需要人工检查'
            df.loc[df['Gene_Mfg_ID'] == df.at[best_clone_index, 'Gene_Mfg_ID'], 'Handled'] = 'Done'
            df.loc[df['Gene_Mfg_ID'] == df.at[best_clone_index, 'Gene_Mfg_ID'], 'Note'] = '需要人工检查'
            # print(f"Gene_Mfg_ID: {df.at[best_clone_index, 'Gene_Mfg_ID']}, Selected: Lowest_MAP&SNP_CONCERN_Selected")

        # 如果Gene_Mfg_ID组内仅存在LARGE_INDEL_CONCERN克隆，则选择最优克隆
        if true_clones.empty and lowest_map_concern_clones.empty and not large_indel_concern_clones.empty and gatk_indel_concern_clones.empty and gatk_snp_concern_clones.empty and lowest_map_snp_concern_clones.empty:
            # 检查是否存在median_depth > 50且flagstat > 98.5的克隆数量 > 1
            large_indel_concern_clones_above_50 = large_indel_concern_clones[(large_indel_concern_clones['median_depth'] > 50) & (large_indel_concern_clones['flagstat'] > 98.5)]
            if len(large_indel_concern_clones_above_50) > 1:
                # 在median_depth > 50的LARGE_INDEL_CONCERN克隆内部进行flagstat降序排序
                sorted_clones = large_indel_concern_clones_above_50.sort_values(by=['flagstat'], ascending=[False])
            else:
                # 在所有LARGE_INDEL_CONCERN克隆内部进行flagstat降序排序
                sorted_clones = large_indel_concern_clones.sort_values(by=['flagstat'], ascending=[False])
            best_clone_index = sorted_clones.index[0]
            df.at[best_clone_index, 'Manual_Selected'] = 'LARGE_INDEL_CONCERN_Selected'
            # 当某一克隆的'Manual_Selected'为LARGE_INDEL_CONCERN_Selected时, 将该克隆所在的Gene_Mfg_ID组的'Handled'列设置为Done，在Note列中注释'需要人工检查'
            df.loc[df['Gene_Mfg_ID'] == df.at[best_clone_index, 'Gene_Mfg_ID'], 'Handled'] = 'Done'
            df.loc[df['Gene_Mfg_ID'] == df.at[best_clone_index, 'Gene_Mfg_ID'], 'Note'] = '需要人工检查'
            # print(f"Gene_Mfg_ID: {df.at[best_clone_index, 'Gene_Mfg_ID']}, Selected: LARGE_INDEL_CONCERN_Selected")
        
        # 以上步骤未能完成克隆选择的Gene_Mfg_ID组
        current_gene = group['Gene_Mfg_ID'].iloc[0]
        if not any(df.loc[df['Gene_Mfg_ID'] == current_gene, 'Handled'] == 'Done'):
            
            # 如果Gene_Mfg_ID组内仅存在DISCARD克隆，则在note列中注释'全部为DISCARD克隆', 在handled列中注释'Done'
            if len(discard_clones) == clone_number:
                df.loc[df['Gene_Mfg_ID'] == group['Gene_Mfg_ID'].iloc[0], 'Note'] = '全部为DISCARD克隆'
                df.loc[df['Gene_Mfg_ID'] == group['Gene_Mfg_ID'].iloc[0], 'Handled'] = 'Done'
            
            # 如果Gene_Mfg_ID组内存在多种CONCERN克隆
            else:
                # 将Lowest_MAP_CONCERN, GATK_SNP_CONCERN, Lowest_MAP&SNP_CONCERN 克隆合并
                Mapping_rate_CONCERN_clones = pd.concat([lowest_map_concern_clones, gatk_snp_concern_clones, lowest_map_snp_concern_clones])
                
                # 如果mapping_rate_CONCERN_clones不为空，那么在这些克隆内部进行lowest_MAP降序排序
                if not Mapping_rate_CONCERN_clones.empty:
                    # 检查是否存在median_depth > 50且flagstat > 98.5的克隆数量 > 1
                    Mapping_rate_CONCERN_clones_above_50 = Mapping_rate_CONCERN_clones[(Mapping_rate_CONCERN_clones['median_depth'] > 50) & (Mapping_rate_CONCERN_clones['flagstat'] > 98.5)]
                    if len(Mapping_rate_CONCERN_clones_above_50) > 1:
                        sorted_clones = Mapping_rate_CONCERN_clones_above_50.sort_values(by=['lowest_MAP'], ascending=[False])
                    else:
                        sorted_clones = Mapping_rate_CONCERN_clones.sort_values(by=['lowest_MAP'], ascending=[False])
                    best_clone_index = sorted_clones.index[0]
                    df.at[best_clone_index, 'Manual_Selected'] = 'Lowest_MAP_CONCERN_Selected'
                    # 当某一克隆的'Manual_Selected'为Lowest_MAP_CONCERN_Selected时, 将该克隆所在的Gene_Mfg_ID组的'Handled'列设置为Done
                    df.loc[df['Gene_Mfg_ID'] == df.at[best_clone_index, 'Gene_Mfg_ID'], 'Handled'] = 'Done'
                    df.loc[df['Gene_Mfg_ID'] == df.at[best_clone_index, 'Gene_Mfg_ID'], 'Note'] = '需要人工检查'

                # 如果mapping_rate_CONCERN_clones为空，则在note列中注释'存在CONCERN克隆，需要人工检查和选择'，在handled列中注释'Done'    
                else:
                    df.loc[df['Gene_Mfg_ID'] == group['Gene_Mfg_ID'].iloc[0], 'Note'] = '存在CONCERN克隆,需要人工检查和选择'
                    df.loc[df['Gene_Mfg_ID'] == group['Gene_Mfg_ID'].iloc[0], 'Handled'] = 'Done'

    # 检查是否有未处理的Gene_Mfg_ID组, 如果有则保存为Mfg_ID_TBD.xls, 如果没有则保存为Clone_Definition&selection.xls
    if 'To_Be_Done' in df['Handled'].tolist():
        print("存在未处理的Gene_Mfg_ID组, 请人工检查分析未处理的Gene_Mfg_ID组, 或检查代码逻辑是否遍历所有Gene_Mfg_ID组")
        # df_to_be_done = df[df['Handled'] == 'To_Be_Done']
        # df_TBD_file_name = os.path.splitext(data_file_name)[0] + 'Mfg_ID_TBD.xls'
        # new_output_file_path = os.path.join(data_file_dir, df_TBD_file_name)
        # df_to_be_done.to_excel(new_output_file_path, index=False)
        # print(f"未处理的Gene_Mfg_ID组已保存为表格: {new_output_file_path}")
    
    # 保存增加克隆选择的表格为Clone_selection.xls     
    new_data_file_name = os.path.splitext(data_file_name)[0] + '_Clone_Definition&Selection.xlsx'
    new_output_file_path = os.path.join(data_file_dir, new_data_file_name)
    df.to_excel(new_output_file_path, index=False)
    print(f"增加克隆选择的完整数据表格已保存为: {new_output_file_path}")
    # return df, output_dir

if __name__ == "__main__":
    print("欢迎使用ColonyNGS数据自动化处理程序\n")
    
    print("输入 '1' 进行克隆判断 -> 统计 -> 克隆选择;\n输入 '2' 进行统计 -> 克隆选择; \n输入 '3' 进行克隆选择; \n其它字符则退出。\n")
    user_input = input("请确认操作种类? \n")
    if user_input == '3':
        print('开始进行克隆选择，请准备完成统计注释的克隆判断文件\n')
        user_input = input("请确认是否继续操作？输入 'Y' 继续，输入 任意字符 退出: \n")
        if user_input != 'Y':
            print("操作已取消")
            exit()
        elif user_input != 'Y':
            clone_selection()
            print("克隆选择已完成\n")
    elif user_input == '2':
        print("开始进行统计，请准备统计模板文件和克隆判断文件, 检查统计模板文件的Gene_Mfg_ID列已录入完整基因列表, 强烈建议将其它信息列录入完整\n")
        user_input = input("请确认是否继续操作？输入 'Y' 继续，输入 任意字符 退出: \n")
        if user_input != 'Y':
            print("操作已取消")
            exit()
        elif user_input == 'Y':
            statistics()
            print("统计已完成\n")
            print('开始进行克隆选择，请准备完成统计注释的克隆判断文件\n')
            user_input = input("请确认是否继续操作？输入 'Y' 继续，输入 任意字符 退出: \n")
            if user_input != 'Y':
                print("操作已取消")
                exit()
            elif user_input != 'Y':
                clone_selection()
                print("克隆选择已完成\n")    
            clone_selection()
            print("克隆选择已完成\n")
    elif user_input == '1':
        print("请准备好原始表格, 检查原始表格每一行的第一个单元格都是Gene_Mfg_ID, 否则需要手动修改\n")
        print('请使用excel将原始表格另存为.xls格式, 该步骤不可省略\n')
        user_input = input("请确认是否完成上述操作？输入 'Y' 确认，输入 任意字符 退出: \n")
        if user_input != 'Y':
            print("操作已取消")
            exit()
        elif user_input == 'Y':
            print('请选择原始表格: \n')
            df, output_dir, input_file_name = sheet_modification()
            judgement_addition()
            clone_judgement()
            print("克隆判断已完成\n")
            print("开始进行统计，请准备统计模板文件和克隆判断文件, 检查统计模板文件的Gene_Mfg_ID列已录入完整基因列表, 强烈建议将其它信息列录入完整\n")
            user_input = input("请确认是否继续操作？输入 'Y' 继续，输入 任意字符 退出: \n")
            if user_input != 'Y':
                print("操作已取消")
                exit()
            elif user_input == 'Y':
                statistics()
                print("统计已完成\n")
                print('开始进行克隆选择，请准备完成统计注释的克隆判断文件\n')
                user_input = input("请确认是否继续操作？输入 'Y' 继续，输入 任意字符 退出: \n")
                if user_input != 'Y':
                    print("操作已取消")
                    exit()
                elif user_input == 'Y':
                    clone_selection()
                    print("克隆选择已完成\n")
    elif user_input != '1' or '2' or '3':
        print("操作已取消")
        exit()
    print("程序已完成所有操作，感谢使用！")
    


    
    



