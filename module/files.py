#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文件处理模块程序
支持CSV、Excel的读取和文档处理功能
"""
import pandas as pd
import os

def classify_filename(filename,classify_name1,classify_name2):

    """
    通过文件名分类
    支持两种类型识别方式：
    1. 按文件名后缀识别（如 _point 或 _alarm_rule）
    2. 按文件名包含的关键词识别
    """
    # # 方式1：按最后的下划线分割识别类型（推荐）
    # if '_' in filename:
    #     suffix = filename.split('_')[-1].lower().replace('.csv', '')
    #     if suffix in ['point', 'alarmrule']:  # 处理可能的格式差异
    #         return suffix
    #     if suffix.startswith('alarm'):  # 处理 alarm_rule 可能被分割为 alarm
    #         return 'alarm_rule'
    
    # 方式2：通过关键词匹配识别
    filename_lower = filename.lower()
    if classify_name1 in filename_lower:
        return classify_name1
    elif classify_name2 in filename_lower:
        return classify_name2
    return None

def _read_csv_with_fallback(file_path: str):
    """尝试多种常见编码读取 CSV，避免 'utf-8' 解码失败。
    优先顺序：utf-8-sig -> gbk/cp936 -> gb18030 -> big5 -> latin1
    """
    tried = []
    # 常见中文/通用编码回退序列
    encodings = [
        'utf-8-sig',
        'gbk', 'cp936',
        'gb18030',
        'big5',
        'latin1',  # 最后兜底：字节到字符一一映射，保证不报错
    ]
    for enc in encodings:
        try:
            df = pd.read_csv(file_path, encoding=enc)
            print(f"已使用编码 {enc} 成功读取: {file_path}")
            return df
        except UnicodeDecodeError:
            tried.append(enc)
            continue
        except Exception:
            # 对于分隔符或引擎问题，尝试使用 python 引擎再试一次
            try:
                df = pd.read_csv(file_path, encoding=enc, engine='python')
                print(f"已使用编码 {enc} + python 引擎 成功读取: {file_path}")
                return df
            except Exception:
                tried.append(enc)
                continue
    raise UnicodeDecodeError("csv", b"", 0, 1, f"所有尝试的编码均失败: {tried}")

def read_file(file_path):
    """读取 CSV 或 Excel 文件并返回 DataFrame"""
    try:
        _, file_extension = os.path.splitext(file_path)
        if file_extension.lower() == '.csv':
            print(f"正在读取 CSV 文件: {file_path}")
            return _read_csv_with_fallback(file_path)
        elif file_extension.lower() == '.xlsx':
            print(f"正在读取 Excel 文件: {file_path}")
            return pd.read_excel(file_path)
        else:
            print(f"文件格式不支持: {file_extension}")
            return None
    except Exception as e:
        print(f"读取文件 {file_path} 时出错: {e}")
        return None
    
def save_to_excel(df, output_path):
    """保存 DataFrame 到 Excel 文件"""
    try:
        df.to_excel(output_path, index=False)
        print(f"拼接结果已保存到: {output_path}")
    except Exception as e:
        print(f"保存 Excel 文件时出错: {e}")

def save_to_csv(df,output_path):
    """保存 DataFrame 到 Excel 文件"""
    try:
        df.to_csv(output_path, index=False, encoding='utf-8-sig')
        print(f"文件已保存: {output_path}，行数: {len(df)}")
    except Exception as e:
        print(f"保存 CSV 文件时出错: {e}")

def data_processing(df, processing_config):
    """ 
    根据配置处理 DataFrame，包括删除、添加、重命名列以及替换列内容等操作
    :param df: 输入的 DataFrame 
    :param processing_config: 处理配置字典，包含以下键：
        - 'delete': 要删除的列名列表
        - 'add': 要添加的新列及其默认值字典
        - 'rename': 重命名列的字典
        - 'replace': 替换列内容的字典，格式为 {列名: {旧值: 新值}}
    """ 
    # 1. 删除表头（列）
    if 'delete' in processing_config:
        df = df.drop(columns=processing_config['delete'], errors='ignore')
    
    # 2. 添加新表头（列）
    if 'add' in processing_config:
        for col, default_value in processing_config['add'].items():
            df[col] = default_value

    # 3. 重命名表头
    if 'rename' in processing_config:
        df = df.rename(columns=processing_config['rename'])
    
    
    # 4. 替换列内容
    if 'replace' in processing_config:
        for col, replacements in processing_config['replace'].items():
            if col in df.columns:
                df[col] = df[col].replace(replacements)

    # # 5. 重新排序列
    # if 'reorder' in processing_config:
    #     # 只保留配置中存在的列
    #     existing_columns = [col for col in processing_config['reorder'] if col in df.columns]
    #     df = df[existing_columns]
    
    # 6. 筛选表头'source'中内容为'IGS','EMS','Math'的行
    df = df[df['source'].isin(['IGS','EMS','Meter','ECU'])]

    # 7. 按'addr'列排序
    df = df.sort_values(by='description')



    # 添加返回处理后的 DataFrame
    return df


if __name__ == "__main__":
    # 示例配置
    processing_configs1= {
        # 处理point文件
        'point':{
            'delete': ['分类', '页面名称', '页面内容', 'Unnamed: 3', '数据来源', '页面点位', '备注', '备注.1'],
            # 'add': {'point_type': 2,
            #         },
            'rename': {'点位描述': 'description',
                    '点位名称': 'name',
                    },
            # 'replace': {'system__calculate_type': {'EMS': 1, 'Math': '2'},
            #             'is_status': {'1': 'TRUE', '0': 'FALSE'},
            #             'is_active': {'1': 'TRUE', '0': 'FALSE'},
            #             'disconnect_op': {'置0': 2, '置空': 3, '保持': 1, '默认值': 4},
            #             'storage': {'秒':1, '分钟':2, '小时': 3, '天': 4, '月': 5, '年': 6},
            #             }
        },
        # 处理alarm文件
        'alarm':{
            'alarm_rule': {
            'rename': {'rule_code': 'alarm_code', 'cond': 'condition'},
            'add': {'acknowledged': False},
            'replace': {'level': {'1': 'low', '2': 'medium', '3': 'high'}},
            'delete': ['legacy_field'],
            'reorder': ['alarm_code', 'condition', 'level']
            },
        }
    }

    input_dir = r'./data'  # 替换为输入目录
    output_dir = './data'  # 替换为输出目录

    # 处理目录下所有文件
    for filename in os.listdir(input_dir):
        if filename.endswith(".csv") or filename.endswith(".xlsx"):

            # 通过文件名分类
            file_type = classify_filename(filename,'point','alarm')

            if not file_type or file_type not in processing_configs:
                print(f"⚠️ 未识别的文件类型: {filename}，已跳过")
                continue

            # 创建类型专属输出目录
            type_output_dir = os.path.join(output_dir, file_type)
            os.makedirs(type_output_dir, exist_ok=True)

            # 输入输出文件路径
            input_path = os.path.join(input_dir, filename)
            output_path = os.path.join(type_output_dir, filename)

            # 读取文件
            df = read_file(input_path)

            if df is not None:
                try:
                    # 处理数据
                    df = data_processing(df, processing_configs[file_type])
                    print(f"✅ 已处理 {filename} → {file_type} 类型")
                except Exception as e:
                    print(f"❌ 处理失败 {filename}: {str(e)}")

                # 保存结果到 csv
                save_to_csv(df, output_path)

    print("\n处理完成！输出目录结构：")
    print(f"{output_dir}/")
    for root, dirs, files in os.walk(output_dir):
        level = root.replace(output_dir, '').count(os.sep)
        indent = ' ' * 4 * level
        print(f'{indent}{os.path.basename(root)}/')
        sub_indent = ' ' * 4 * (level + 1)
        for f in files[:3]:  # 显示前3个文件示例
            print(f'{sub_indent}{f}')
        if len(files) > 3:
            print(f'{sub_indent}...（共 {len(files)} 个文件）')

    print("All files processed!")            
                