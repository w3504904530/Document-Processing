import pandas as pd
import os


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

    # 添加返回处理后的 DataFrame
    return df


if __name__ == "__main__":
    # 示例配置
    processing_config = {
        'delete': ['new_name','upload_name','description.1'],
        'add': {'point_type': 2,
                },
        'rename': {'source': 'system__calculate_type',
                   'category': 'source',
                   'calculate_interval': 'system__calculate_interval',
                   'script': 'system__script',
                   'value': 'system__value',
                   '断联操作': 'disconnect_op'
                   },
        'replace': {'system__calculate_type': {'EMS': 1, 'Math': '2'},
                    'is_status': {'1': 'TRUE', '0': 'FALSE'},
                    'is_active': {'1': 'TRUE', '0': 'FALSE'},
                    'disconnect_op': {'置0': 2, '置空': 3, '保持': 1, '默认值': 4},
                    'storage': {'秒':1, '分钟':2, '小时': 3, '天': 4, '月': 5, '年': 6},
                    }
    }
    
    # 读取文件
    input_file = '数据处理\data\system_point.csv'  # 替换为实际文件路径
    df = read_file(input_file)
    
    if df is not None:
        # 处理数据
        df = data_processing(df, processing_config)
        
        # 保存结果到 Excel
        output_file = '数据处理\data\output_system_point.csv'  # 替换为实际输出路径
        save_to_csv(df, output_file)