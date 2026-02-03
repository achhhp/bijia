#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
供应商比价软件 - Web版本
功能：
1. 上传至少3份报价单（Excel或CSV格式）
2. 自动找出每个物料的最低价和对应供应商
3. 统计每个供应商中了哪些低价物料
4. 展示结果并导出分析报告
"""

import os
import tempfile
import pandas as pd
from flask import Flask, request, render_template, redirect, url_for, send_file
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB
app.config['ALLOWED_EXTENSIONS'] = {'xlsx', 'xls', 'csv'}

# 全局变量，用于存储分析结果
analysis_result = None
vendor_stats = None


def allowed_file(filename):
    """检查文件是否允许上传"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']


def extract_vendor_name_from_content(df):
    """从文件内容中提取供应商名称"""
    # 尝试从第一行或前几行提取供应商名称
    for i in range(min(10, len(df))):
        row = df.iloc[i]
        for j, cell in enumerate(row):
            cell_str = str(cell)
            print(f"检查单元格 [{i},{j}]: {cell_str}")
            if '供应商' in cell_str or '供货' in cell_str:
                # 提取供应商名称
                # 例如："供货供应商：瓦房店市君创百货贸易商店" -> "瓦房店市君创百货贸易商店"
                if '：' in cell_str:
                    parts = cell_str.split('：')
                    if len(parts) > 1:
                        vendor_name = parts[1].strip()
                        # 去除采购时间等无关信息
                        if '采购时间' in vendor_name:
                            vendor_name = vendor_name.split('采购时间')[0].strip()
                        print(f"提取到供应商名称：{vendor_name}")
                        return vendor_name
                elif ':' in cell_str:
                    parts = cell_str.split(':')
                    if len(parts) > 1:
                        vendor_name = parts[1].strip()
                        # 去除采购时间等无关信息
                        if '采购时间' in vendor_name:
                            vendor_name = vendor_name.split('采购时间')[0].strip()
                        print(f"提取到供应商名称：{vendor_name}")
                        return vendor_name
    
    # 尝试另一种方式：直接检查第二行（根据用户提供的格式）
    if len(df) >= 2:
        row = df.iloc[1]
        for j, cell in enumerate(row):
            cell_str = str(cell)
            print(f"检查第二行单元格 [{1},{j}]: {cell_str}")
            # 第二行可能直接包含供应商名称
            if cell_str and not cell_str.startswith('Unnamed'):
                # 尝试从单元格中提取供应商名称
                if '：' in cell_str:
                    parts = cell_str.split('：')
                    if len(parts) > 1:
                        vendor_name = parts[1].strip()
                        # 去除采购时间等无关信息
                        if '采购时间' in vendor_name:
                            vendor_name = vendor_name.split('采购时间')[0].strip()
                        print(f"从第二行提取到供应商名称：{vendor_name}")
                        return vendor_name
                elif ':' in cell_str:
                    parts = cell_str.split(':')
                    if len(parts) > 1:
                        vendor_name = parts[1].strip()
                        # 去除采购时间等无关信息
                        if '采购时间' in vendor_name:
                            vendor_name = vendor_name.split('采购时间')[0].strip()
                        print(f"从第二行提取到供应商名称：{vendor_name}")
                        return vendor_name
    
    print("无法从文件内容中提取供应商名称")
    return None


def process_dataframe(df, filename, vendor_name):
    """处理数据框，查找必要的列"""
    # 获取所有列名，用于调试
    all_columns = list(df.columns)
    
    # 尝试从文件内容中提取供应商名称
    content_vendor_name = extract_vendor_name_from_content(df)
    if content_vendor_name:
        vendor_name = content_vendor_name
        # 进一步清理供应商名称，去除采购时间等无关信息
        if '采购时间' in vendor_name:
            vendor_name = vendor_name.split('采购时间')[0].strip()
        # 去除任何非中文字符和数字以外的内容
        vendor_name = ''.join([c for c in vendor_name if c.isalnum() or c in ' .-']).strip()
        print(f"从文件内容中提取的供应商名称：{vendor_name}")
    else:
        # 如果无法从内容中提取，使用文件名作为供应商名称
        vendor_name = filename.split('.')[0]
        print(f"从文件名中提取的供应商名称：{vendor_name}")
    
    # 尝试处理没有明确列名的情况
    # 情况1：如果列名是Unnamed，尝试跳过标题行，找到实际数据
    if all('Unnamed' in str(col) for col in df.columns) or len(df.columns) == 1:
        print(f"文件 {filename} 可能包含标题行，尝试找到实际数据...")
        
        # 跳过标题行，找到实际数据的起始位置
        data_start_row = 0
        for i in range(len(df)):
            row = df.iloc[i]
            # 检查是否包含数字（可能是价格或数量）
            has_number = False
            for cell in row:
                try:
                    float(cell)
                    has_number = True
                    break
                except:
                    pass
            if has_number:
                data_start_row = i
                break
        
        print(f"实际数据起始行：{data_start_row}")
        
        # 如果找到了数据起始行，使用该行作为列名
        if data_start_row > 0:
            new_columns = df.iloc[data_start_row].tolist()
            df = df.iloc[data_start_row + 1:].reset_index(drop=True)
            df.columns = new_columns
            # 更新列名列表
            all_columns = list(df.columns)
            print(f"新的列名：{all_columns}")
    
    # 标准化列名（去除空格、特殊字符，转换为小写）
    normalized_columns = {}
    for col in df.columns:
        # 标准化列名：去除空格、特殊字符，转换为小写
        normalized_col = ''.join(e for e in str(col) if e.isalnum() or e == '_').lower()
        normalized_columns[normalized_col] = col
    
    # 查找序号列
    serial_col = None
    serial_col_options = ['序号', '编号', 'id', 'no', 'number']
    for option in serial_col_options:
        normalized_option = ''.join(e for e in option if e.isalnum() or e == '_').lower()
        if normalized_option in normalized_columns:
            serial_col = normalized_columns[normalized_option]
            break
    
    # 如果通过名称找不到序号列，尝试通过位置查找（假设第一列是序号列）
    if not serial_col and len(df.columns) >= 1:
        serial_col = df.columns[0]
        print(f"通过位置选择序号列：{serial_col}")
    
    # 查找物料列（支持更多可能的列名）
    item_col = None
    item_col_options = ['品名', '名称', '物料', '货品', '商品', '品目', '项目', '货物', '物料名称', '商品名称', 'item', 'name', 'product']
    for option in item_col_options:
        normalized_option = ''.join(e for e in option if e.isalnum() or e == '_').lower()
        if normalized_option in normalized_columns:
            item_col = normalized_columns[normalized_option]
            break
    
    # 如果通过名称找不到物料列，尝试通过位置查找（假设第二列是物料列）
    if not item_col and len(df.columns) >= 2:
        item_col = df.columns[1]
        print(f"通过位置选择物料列：{item_col}")
    
    # 查找价格列（支持更多可能的列名）
    price_col = None
    price_col_options = ['单项报价', '报价', '价格', '单价', '单位价格', '单价(元)', '价格(元)', '报价(元)', '单位报价', 'price', 'cost', 'unitprice']
    for option in price_col_options:
        normalized_option = ''.join(e for e in option if e.isalnum() or e == '_').lower()
        if normalized_option in normalized_columns:
            price_col = normalized_columns[normalized_option]
            break
    
    # 如果通过名称找不到价格列，尝试通过位置查找（假设第三列是价格列）
    if not price_col and len(df.columns) >= 3:
        price_col = df.columns[2]
        print(f"通过位置选择价格列：{price_col}")
    
    # 查找分项小计列
    subtotal_col = None
    subtotal_col_options = ['分项小计', '小计', '金额', 'total', 'amount']
    for option in subtotal_col_options:
        normalized_option = ''.join(e for e in option if e.isalnum() or e == '_').lower()
        if normalized_option in normalized_columns:
            subtotal_col = normalized_columns[normalized_option]
            break
    
    # 检查是否找到必要的列
    if not item_col or not price_col:
        # 提供更详细的错误信息，包含文件的所有列名
        return None, f"文件 {filename} 缺少必要的列：品名/名称 或 单项报价/价格。文件中的列名：{', '.join(all_columns)}"
    
    # 清理数据
    try:
        df = df.dropna(subset=[item_col, price_col])
        # 标准化物料名称，去除特殊字符和空格，使不同格式的相同物料名称标准化
        def normalize_item_name(name):
            name = str(name).strip()
            # 去除特殊字符
            name = ''.join(e for e in name if e.isalnum() or e in ' -')
            # 去除多余空格
            name = ' '.join(name.split())
            # 转换为小写
            name = name.lower()
            return name
        
        df['物料'] = df[item_col].apply(normalize_item_name)  # 标准化物料名称
        df['原始物料名称'] = df[item_col].astype(str).str.strip()  # 保存原始物料名称
        df['价格'] = pd.to_numeric(df[price_col], errors='coerce')  # 重命名为'价格'，与后续代码保持一致
        
        # 过滤掉价格为0或空的行（视为弃权）
        df = df[(df['价格'] > 0) & (df['价格'].notna())]
        print(f"过滤后剩余物料数量：{len(df)}")
        
        # 处理序号
        if serial_col:
            df['序号'] = df[serial_col].astype(str).str.strip()
        else:
            # 如果没有序号列，生成默认序号
            df['序号'] = range(1, len(df) + 1)
        
        # 处理分项小计
        print(f"开始处理分项小计，当前列名：{list(df.columns)}")
        print(f"标准化列名：{normalized_columns}")
        print(f"当前数据前5行：{df.head().to_dict('records')}")
        
        if subtotal_col:
            print(f"找到分项小计列：{subtotal_col}")
            df['分项小计'] = pd.to_numeric(df[subtotal_col], errors='coerce')
            print(f"分项小计列数据预览：{df['分项小计'].head().tolist()}")
        else:
            # 如果没有分项小计列，尝试计算
            print("未找到分项小计列，尝试计算")
            
            # 尝试识别数量列
            quantity_col = None
            
            # 方法1：尝试通过列名识别
            quantity_col_options = ['需求量', '数量', 'qty', 'num', 'amount', 'count', '采购数量', '订购数量', '数量单位', '需 求量']
            print(f"尝试通过列名识别数量列，选项：{quantity_col_options}")
            
            for option in quantity_col_options:
                normalized_option = ''.join(e for e in option if e.isalnum() or e == '_').lower()
                if normalized_option in normalized_columns:
                    quantity_col = normalized_columns[normalized_option]
                    print(f"通过列名找到数量列：{quantity_col}")
                    break
            
            # 方法1.1：直接检查所有列名，寻找包含数量相关关键词的列
            if not quantity_col:
                print("通过标准选项未找到数量列，尝试检查所有列名")
                for col in df.columns:
                    # 标准化列名，去除空格和特殊字符
                    normalized_col = ''.join(e for e in str(col) if e.isalnum() or e == '_').lower()
                    # 检查是否包含数量相关的关键词
                    if any(keyword in normalized_col for keyword in ['qty', 'num', 'amount', 'count', 'quantity', 'demand']):
                        quantity_col = col
                        print(f"通过关键词找到数量列：{col}")
                        break
                    # 检查是否包含中文数量相关词汇
                    if any(keyword in str(col) for keyword in ['需求', '数量', '用量', '订购']):
                        quantity_col = col
                        print(f"通过中文关键词找到数量列：{col}")
                        break
            
            # 方法2：如果列名都是Unnamed，尝试通过位置和内容识别
            if not quantity_col and all('Unnamed' in str(col) for col in df.columns):
                print("列名都是Unnamed，尝试通过位置和内容识别数量列")
                # 对于典型的报价单格式，数量列通常在价格列附近
                # 尝试检查第3-6列（假设序号、物料、价格在前三列）
                for i in range(2, min(6, len(df.columns))):
                    col = df.columns[i]
                    print(f"检查第 {i+1} 列：{col}")
                    
                    # 检查该列是否包含数字
                    sample_values = df[col].head()
                    print(f"该列数据预览：{sample_values.tolist()}")
                    
                    # 计算该列中可以转换为数字的值的比例
                    numeric_count = 0
                    total_count = 0
                    for val in sample_values:
                        total_count += 1
                        try:
                            float(val)
                            numeric_count += 1
                        except:
                            pass
                    
                    # 如果超过50%的值是数字，且不是价格列和序号列，则认为是数量列
                    if numeric_count / total_count > 0.5 and col != price_col and col != serial_col:
                        quantity_col = col
                        print(f"通过内容识别找到数量列：{col}，数字比例：{numeric_count/total_count:.2f}")
                        break
            
            # 方法3：检查所有列，寻找可能的数量列
            if not quantity_col:
                print("通过名称和位置未找到数量列，尝试检查所有列")
                
                # 特殊处理：直接检查所有Unnamed列，寻找包含数量数据的列
                print("尝试检查所有Unnamed列，寻找数量数据")
                # 存储所有可能的数量列
                possible_quantity_cols = []
                
                for col in df.columns:
                    if 'Unnamed' in str(col):
                        try:
                            sample_values = df[col].head()
                            print(f"检查Unnamed列：{col}，数据预览：{sample_values.tolist()}")
                            
                            # 跳过价格列（通常包含"报价"、"价格"等关键词，或者列名中包含价格相关的内容）
                            col_str = str(col).lower()
                            header_str = ""
                            if len(sample_values) > 0:
                                try:
                                    header_str = str(sample_values[0]).lower()
                                except:
                                    pass
                            if any(keyword in col_str or keyword in header_str for keyword in ['price', '报价', '价格', '单价']):
                                print(f"跳过列：{col}（可能是价格列）")
                                continue
                            
                            # 检查该列是否包含数字
                            numeric_count = 0
                            total_count = 0
                            valid_values = []
                            for val in sample_values:
                                total_count += 1
                                try:
                                    val_float = float(val)
                                    numeric_count += 1
                                    valid_values.append(val_float)
                                except:
                                    pass
                            
                            # 如果超过50%的值是数字，则认为是可能的数量列
                            if total_count > 0 and numeric_count / total_count > 0.5:
                                # 检查该列的值是否合理（通常数量大于1）
                                has_large_value = any(val > 1 for val in valid_values)
                                if has_large_value:
                                    possible_quantity_cols.append((col, numeric_count / total_count, valid_values))
                                    print(f"找到可能的数量列：{col}，数字比例：{numeric_count/total_count:.2f}")
                        except Exception as e:
                            print(f"检查Unnamed列时出错：{str(e)}")
                            import traceback
                            traceback.print_exc()
                            continue
                
                # 从可能的数量列中选择最合适的
                if possible_quantity_cols:
                    # 按数字比例排序，选择最高的
                    possible_quantity_cols.sort(key=lambda x: x[1], reverse=True)
                    quantity_col = possible_quantity_cols[0][0]
                    print(f"选择最佳数量列：{quantity_col}")
                
                # 特殊处理：直接检查第5列（根据用户的报价格式，第5列通常是数量列）
                # 强制使用第5列作为数量列，因为根据用户的报价格式，第5列确实是数量列
                if len(df.columns) >= 5:
                    try:
                        potential_quantity_col = df.columns[4]  # 第5列（索引为4）
                        print(f"强制使用第5列 {potential_quantity_col} 作为数量列")
                        sample_values = df[potential_quantity_col].head()
                        print(f"该列数据预览：{sample_values.tolist()}")
                        
                        # 检查该列是否包含数字
                        numeric_count = 0
                        total_count = 0
                        valid_values = []
                        for val in sample_values:
                            total_count += 1
                            try:
                                val_float = float(val)
                                numeric_count += 1
                                valid_values.append(val_float)
                            except:
                                pass
                        
                        # 检查该列的值是否合理（通常数量大于1）
                        has_large_value = any(val > 1 for val in valid_values) if valid_values else False
                        
                        # 无论如何都使用第5列作为数量列
                        quantity_col = potential_quantity_col
                        print(f"强制使用第5列作为数量列：{quantity_col}")
                    except Exception as e:
                        print(f"检查第5列时出错：{str(e)}")
                        import traceback
                        traceback.print_exc()
                        # 如果出错，尝试使用默认数量1
                        quantity_col = None
                
                # 如果仍然没有找到，尝试检查价格列附近的列
                if not quantity_col:
                    # 优先检查价格列附近的列（通常数量列在价格列附近）
                    price_col_index = -1
                    for i, col in enumerate(df.columns):
                        if col == price_col:
                            price_col_index = i
                            break
                    
                    # 首先检查价格列附近的列
                    if price_col_index != -1:
                        try:
                            print(f"价格列在位置：{price_col_index}")
                            # 检查价格列前后的列
                            found_quantity = False
                            for i in range(max(0, price_col_index - 2), min(len(df.columns), price_col_index + 3)):
                                if i == price_col_index:
                                    continue
                                col = df.columns[i]
                                # 跳过序号列
                                if col == serial_col:
                                    print(f"跳过列：{col}（序号列）")
                                    continue
                                
                                print(f"检查价格列附近的列：{col}（位置：{i}")
                                sample_values = df[col].head()
                                print(f"该列数据预览：{sample_values.tolist()}")
                                
                                # 检查该列是否包含数字
                                numeric_count = 0
                                total_count = 0
                                for val in sample_values:
                                    total_count += 1
                                    try:
                                        float(val)
                                        numeric_count += 1
                                    except:
                                        pass
                                
                                # 如果超过50%的值是数字，则认为是数量列
                                if total_count > 0 and numeric_count / total_count > 0.5:
                                    quantity_col = col
                                    print(f"通过价格列附近检查找到数量列：{col}，数字比例：{numeric_count/total_count:.2f}")
                                    found_quantity = True
                                    break
                        except Exception as e:
                            print(f"检查价格列附近的列时出错：{str(e)}")
                            import traceback
                            traceback.print_exc()
                            
            # 如果找到数量列，计算分项小计
            if quantity_col:
                try:
                    print(f"使用 {quantity_col} 列计算分项小计")
                    print(f"价格列数据预览：{df['价格'].head().tolist()}")
                    print(f"数量列数据预览：{df[quantity_col].head().tolist()}")
                    
                    # 转换数量列为数字
                    numeric_quantity = pd.to_numeric(df[quantity_col], errors='coerce')
                    print(f"转换后的数量数据预览：{numeric_quantity.head().tolist()}")
                    
                    # 存储数量列，方便后续分析使用
                    df['数量'] = numeric_quantity
                    print(f"存储的数量列数据预览：{df['数量'].head().tolist()}")
                    
                    # 计算分项小计
                    df['分项小计'] = df['价格'] * numeric_quantity
                    print(f"计算后的分项小计预览：{df['分项小计'].head().tolist()}")
                except Exception as e:
                    print(f"计算分项小计时出错：{str(e)}")
                    import traceback
                    traceback.print_exc()
                    df['分项小计'] = 0
                    df['数量'] = 1  # 出错时默认数量为1
            else:
                # 尝试使用固定值1作为数量，确保分项小计不为0
                print("未找到数量列，使用默认数量1计算分项小计")
                try:
                    df['数量'] = 1  # 存储默认数量
                    df['分项小计'] = df['价格'] * 1
                    print(f"使用默认数量1计算后的分项小计预览：{df['分项小计'].head().tolist()}")
                except Exception as e:
                    print(f"使用默认数量计算分项小计时出错：{str(e)}")
                    import traceback
                    traceback.print_exc()
                    df['分项小计'] = df['价格']
                    df['数量'] = 1  # 出错时默认数量为1
        
        df = df.dropna(subset=['价格'])
        
        # 调试信息：打印处理后的数据
        print(f"文件 {filename} 处理后的数据（前5行）：")
        print(df[['序号', '物料', '价格', '分项小计']].head())
        print(f"供应商名称：{vendor_name}")
        
        return (vendor_name, df), None
    except Exception as e:
        return None, f"处理文件 {filename} 数据时出错：{str(e)}"


def analyze_prices(files):
    """分析价格，找出最低价"""
    data = {}
    errors = []
    
    # 解析所有文件
    for file_path in files:
        result, error = parse_file(file_path)
        if error:
            errors.append(error)
        else:
            vendor_name, df = result
            data[vendor_name] = df
    
    if not data:
        return None, None, errors
    
    # 收集所有物料
    all_items = set()
    for vendor, df in data.items():
        all_items.update(df['物料'].tolist())
    
    # 分析每个物料的最低价
    analysis_data = []
    for item in all_items:
        min_price = float('inf')
        min_vendor = ""
        min_serial = ""
        min_subtotal = 0
        
        min_quantity = 1  # 默认数量为1
        min_item_name = item  # 默认使用标准化的物料名称
        
        for vendor, df in data.items():
            item_data = df[df['物料'] == item]
            if not item_data.empty:
                price = item_data['价格'].iloc[0]
                if price < min_price:
                    min_price = price
                    min_vendor = vendor
                    min_serial = item_data['序号'].iloc[0]
                    min_subtotal = item_data['分项小计'].iloc[0]
                    # 尝试获取原始物料名称
                    min_item_name = item_data.get('原始物料名称', item_data['物料']).iloc[0]
                    # 尝试获取数量信息
                    # 首先检查是否有'数量'列
                    if '数量' in item_data.columns:
                        try:
                            min_quantity = float(item_data['数量'].iloc[0])
                            print(f"找到数量列：数量，值：{min_quantity}")
                        except Exception as e:
                            print(f"解析数量时出错：{str(e)}")
                            pass
                    # 如果没有找到，尝试检查所有列，寻找可能的数量列
                    if min_quantity == 1:
                        for col in item_data.columns:
                            # 跳过序号列、价格列和分项小计列
                            if col in ['序号', '价格', '分项小计']:
                                continue
                            # 检查列名是否包含数量相关的关键词
                            if any(keyword in str(col).lower() for keyword in ['qty', 'num', 'amount', 'count', 'quantity', '需求量', '数量', '需 求量', '需求']):
                                try:
                                    min_quantity = float(item_data[col].iloc[0])
                                    print(f"找到数量列：{col}，值：{min_quantity}")
                                    break
                                except Exception as e:
                                    print(f"解析数量时出错：{str(e)}")
                                    pass
                    # 如果找不到数量列，尝试通过分项小计和价格计算
                    if min_quantity == 1 and min_subtotal > 0 and min_price > 0:
                        try:
                            min_quantity = min_subtotal / min_price
                            print(f"通过计算得到数量：{min_quantity}")
                        except Exception as e:
                            print(f"计算数量时出错：{str(e)}")
                            pass
                    # 确保数量不为0或负数
                    if min_quantity <= 0:
                        min_quantity = 1
                        print("数量值无效，使用默认值1")
        
        if min_vendor:
            analysis_data.append({
                '序号': min_serial,
                '物料名称': min_item_name,
                '最低价': min_price,
                '供应商': min_vendor,
                '数量': min_quantity,
                '分项小计': min_subtotal
            })
    
    # 生成分析结果并按序号排序
    analysis_result = pd.DataFrame(analysis_data)
    # 确保序号列存在
    if '序号' in analysis_result.columns:
        print(f"排序前的序号：{analysis_result['序号'].head().tolist()}")
        # 使用更直接的排序方法
        try:
            # 方法1：尝试将序号转换为数字
            analysis_result['序号_numeric'] = pd.to_numeric(analysis_result['序号'], errors='coerce')
            # 先按数字序号排序，再按字符串序号排序
            analysis_result = analysis_result.sort_values(['序号_numeric', '序号'], na_position='last').reset_index(drop=True)
            # 删除临时列
            analysis_result = analysis_result.drop('序号_numeric', axis=1)
            print("分析结果按序号排序成功（方法1）")
            print(f"排序后的序号：{analysis_result['序号'].head().tolist()}")
        except Exception as e:
            print(f"方法1排序失败：{str(e)}")
            try:
                # 方法2：使用lambda函数直接排序
                analysis_result = analysis_result.iloc[sorted(range(len(analysis_result)), key=lambda i: (
                    isinstance(analysis_result.iloc[i]['序号'], str),
                    float(analysis_result.iloc[i]['序号']) if str(analysis_result.iloc[i]['序号']).replace('.', '', 1).isdigit() else analysis_result.iloc[i]['序号']
                ))].reset_index(drop=True)
                print("分析结果按序号排序成功（方法2）")
                print(f"排序后的序号：{analysis_result['序号'].head().tolist()}")
            except Exception as e2:
                print(f"方法2排序失败：{str(e2)}")
                try:
                    # 方法3：简单按字符串排序
                    analysis_result = analysis_result.sort_values('序号').reset_index(drop=True)
                    print("分析结果按序号排序成功（方法3）")
                    print(f"排序后的序号：{analysis_result['序号'].head().tolist()}")
                except Exception as e3:
                    print(f"方法3排序失败：{str(e3)}")
    else:
        print("分析结果中没有序号列，无法排序")
    
    # 统计每个供应商的中标情况
    vendor_stats = {}
    for _, row in analysis_result.iterrows():
        vendor = row['供应商']
        item = row['物料名称']
        serial = row['序号']
        price = row['最低价']
        subtotal = row['分项小计']
        
        if vendor not in vendor_stats:
            vendor_stats[vendor] = []
        # 获取数量信息
        quantity = 1  # 默认数量为1
        # 尝试从分析结果中获取数量
        if '数量' in row:
            quantity = row['数量']
        # 如果找不到数量，尝试通过分项小计和价格计算
        elif subtotal > 0 and price > 0:
            try:
                quantity = subtotal / price
            except:
                pass
        
        vendor_stats[vendor].append({
            '序号': serial,
            '物料名称': item,
            '单项报价': price,
            '数量': quantity,
            '分项小计': subtotal
        })
    
    return analysis_result, vendor_stats, errors


def parse_file(file_path):
    """解析报价单文件"""
    try:
        filename = os.path.basename(file_path)
        
        # 提取供应商名称（默认用文件名作为供应商名称）
        vendor_name = filename.split('.')[0]
        
        # 尝试读取Excel文件的所有工作表
        if file_path.endswith('.csv'):
            # CSV文件直接读取
            df = pd.read_csv(file_path)
            return process_dataframe(df, filename, vendor_name)
        else:
            # Excel文件，尝试读取所有工作表
            try:
                # 首先尝试读取第一个工作表
                df = pd.read_excel(file_path)
                result, error = process_dataframe(df, filename, vendor_name)
                if not error:
                    return result, None
                
                # 如果第一个工作表失败，尝试读取所有工作表
                xl = pd.ExcelFile(file_path)
                for sheet_name in xl.sheet_names:
                    df = pd.read_excel(xl, sheet_name)
                    result, error = process_dataframe(df, filename, vendor_name)
                    if not error:
                        return result, None
                
                # 所有工作表都失败
                return None, f"文件 {filename} 的所有工作表都缺少必要的列：品名/名称 或 单项报价/价格"
            except Exception as e:
                return None, f"读取Excel文件 {file_path} 时出错：{str(e)}"
    except Exception as e:
        return None, f"解析文件 {file_path} 时出错：{str(e)}"


def analyze_prices_from_uploads(uploaded_files):
    """直接从上传的文件流分析价格，避免保存到临时文件"""
    data = {}
    errors = []
    
    # 处理每个上传的文件
    for file in uploaded_files:
        if not file or not allowed_file(file.filename):
            errors.append(f'文件 {file.filename} 格式不支持，请上传Excel或CSV文件')
            continue
        
        try:
            filename = secure_filename(file.filename)
            vendor_name = filename.split('.')[0]  # 用文件名作为供应商名称
            
            # 直接从文件流读取数据，避免保存到临时文件
            if file.filename.endswith('.csv'):
                # 对于CSV文件，直接读取
                try:
                    df = pd.read_csv(file.stream)
                    # 调试信息：打印前几行数据
                    print(f"文件 {filename} 的前5行数据：")
                    print(df.head())
                    print(f"文件 {filename} 的列名：")
                    print(list(df.columns))
                    result, error = process_dataframe(df, filename, vendor_name)
                except Exception as csv_error:
                    errors.append(f"读取CSV文件 {filename} 时出错：{str(csv_error)}")
                    continue
            else:  # Excel文件
                # 对于Excel文件，使用BytesIO
                import io
                try:
                    file_content = file.stream.read()
                    excel_file = io.BytesIO(file_content)
                    
                    # 尝试读取第一个工作表
                    try:
                        df = pd.read_excel(excel_file)
                        # 调试信息：打印前几行数据
                        print(f"文件 {filename} 的前5行数据：")
                        print(df.head())
                        print(f"文件 {filename} 的列名：")
                        print(list(df.columns))
                        result, error = process_dataframe(df, filename, vendor_name)
                        
                        # 如果第一个工作表失败，尝试读取所有工作表
                        if error:
                            excel_file.seek(0)  # 重置文件指针
                            xl = pd.ExcelFile(excel_file)
                            print(f"文件 {filename} 的工作表：")
                            print(xl.sheet_names)
                            for sheet_name in xl.sheet_names:
                                print(f"尝试读取工作表：{sheet_name}")
                                df = pd.read_excel(xl, sheet_name)
                                # 调试信息：打印前几行数据
                                print(f"工作表 {sheet_name} 的前5行数据：")
                                print(df.head())
                                print(f"工作表 {sheet_name} 的列名：")
                                print(list(df.columns))
                                result, error = process_dataframe(df, filename, vendor_name)
                                if not error:
                                    print(f"成功读取工作表：{sheet_name}")
                                    break
                    finally:
                        excel_file.close()
                except Exception as excel_error:
                    errors.append(f"读取Excel文件 {filename} 时出错：{str(excel_error)}")
                    continue
            
            if error:
                errors.append(error)
            else:
                vendor_name, df = result
                data[vendor_name] = df
                print(f"成功解析文件：{filename}，供应商：{vendor_name}，物料数量：{len(df)}")
        except Exception as e:
            errors.append(f"处理文件 {file.filename} 时出错：{str(e)}")
            import traceback
            traceback.print_exc()
    
    print(f"解析完成，成功解析的文件数量：{len(data)}")
    if not data:
        return None, None, errors
    
    # 识别最全报价单的供应商（包含最多物料的供应商）
    max_items = 0
    most_complete_vendor = None
    for vendor, df in data.items():
        item_count = len(df['物料'].unique())
        if item_count > max_items:
            max_items = item_count
            most_complete_vendor = vendor
        elif item_count == max_items:
            # 如果数量相同，使用最后一个供应商
            most_complete_vendor = vendor
    print(f"最全报价单供应商：{most_complete_vendor}，包含物料数量：{max_items}")
    
    # 收集所有物料
    all_items = set()
    for vendor, df in data.items():
        all_items.update(df['物料'].tolist())
    
    # 分析每个物料的最低价
    analysis_data = []
    # 收集所有供应商名称
    all_vendors = list(data.keys())
    
    for item in all_items:
        min_price = float('inf')
        min_vendor = ""
        min_serial = ""
        min_subtotal = 0
        min_quantity = 1  # 默认数量为1
        min_item_name = item
        
        # 收集各供应商的报价
        vendor_prices = {}
        for vendor in all_vendors:
            vendor_prices[vendor] = None
        
        for vendor, df in data.items():
            item_data = df[df['物料'] == item]
            if not item_data.empty:
                price = item_data['价格'].iloc[0]
                vendor_prices[vendor] = price
                
                if price < min_price:
                    min_price = price
                    min_vendor = vendor
                    min_serial = item_data['序号'].iloc[0]
                    min_subtotal = item_data['分项小计'].iloc[0]
                    # 尝试获取原始物料名称
                    min_item_name = item_data.get('原始物料名称', item_data['物料']).iloc[0]
                    # 尝试获取数量信息
                    # 首先检查是否有'数量'列
                    if '数量' in item_data.columns:
                        try:
                            min_quantity = float(item_data['数量'].iloc[0])
                            print(f"找到数量列：数量，值：{min_quantity}")
                        except Exception as e:
                            print(f"解析数量时出错：{str(e)}")
                            pass
                    # 如果没有找到，尝试检查所有列，寻找可能的数量列
                    if min_quantity == 1:
                        for col in item_data.columns:
                            # 跳过序号列、价格列和分项小计列
                            if col in ['序号', '价格', '分项小计']:
                                continue
                            # 检查列名是否包含数量相关的关键词
                            if any(keyword in str(col).lower() for keyword in ['qty', 'num', 'amount', 'count', 'quantity', '需求量', '数量', '需 求量', '需求']):
                                try:
                                    min_quantity = float(item_data[col].iloc[0])
                                    print(f"找到数量列：{col}，值：{min_quantity}")
                                    break
                                except Exception as e:
                                    print(f"解析数量时出错：{str(e)}")
                                    pass
                    # 如果找不到数量列，尝试通过分项小计和价格计算
                    if min_quantity == 1 and min_subtotal > 0 and min_price > 0:
                        try:
                            min_quantity = min_subtotal / min_price
                            print(f"通过计算得到数量：{min_quantity}")
                        except Exception as e:
                            print(f"计算数量时出错：{str(e)}")
                            pass
                    # 确保数量不为0或负数
                    if min_quantity <= 0:
                        min_quantity = 1
                        print("数量值无效，使用默认值1")
        
        if min_vendor:
            # 创建分析数据项
            item_data = {
                '序号': min_serial,
                '物料名称': min_item_name,
                '最低价': min_price,
                '供应商': min_vendor,
                '数量': min_quantity,
                '分项小计': min_subtotal
            }
            
            # 添加各供应商的报价列
            for vendor in all_vendors:
                item_data[f'报价_{vendor}'] = vendor_prices[vendor]
            
            analysis_data.append(item_data)
    
    # 生成分析结果
    analysis_result = pd.DataFrame(analysis_data)
    
    # 按对应序号排序（直接使用分析结果中的序号列）
    def get_sortable_value(serial):
        if isinstance(serial, (int, float)):
            return (0, serial)
        try:
            return (0, int(serial))
        except ValueError:
            try:
                return (0, float(serial))
            except ValueError:
                return (1, str(serial))
    
    # 创建排序键列，直接使用分析结果中的序号
    analysis_result['排序键'] = analysis_result['序号'].apply(get_sortable_value)
    # 按排序键排序
    analysis_result = analysis_result.sort_values(by='排序键').reset_index(drop=True)
    # 删除临时列
    analysis_result = analysis_result.drop('排序键', axis=1)
    print("按对应序号排序完成")
    
    # 统计每个供应商的中标情况
    vendor_stats = {}
    for _, row in analysis_result.iterrows():
        vendor = row['供应商']
        # 获取数量信息
        quantity = 1  # 默认数量为1
        if '数量' in row:
            quantity = row['数量']
        # 如果找不到数量，尝试通过分项小计和价格计算
        elif '分项小计' in row and '最低价' in row and row['分项小计'] > 0 and row['最低价'] > 0:
            try:
                quantity = row['分项小计'] / row['最低价']
            except:
                pass
        
        item_info = {
            '序号': row['序号'],
            '物料名称': row['物料名称'],
            '单项报价': row['最低价'],
            '数量': quantity,
            '分项小计': row['分项小计']
        }
        
        if vendor not in vendor_stats:
            vendor_stats[vendor] = []
        vendor_stats[vendor].append(item_info)
    
    return analysis_result, vendor_stats, errors


@app.route('/', methods=['GET', 'POST'])
def index():
    global analysis_result, vendor_stats
    
    if request.method == 'POST':
        # 检查是否有文件上传
        if 'files' not in request.files:
            return render_template('index.html', error='请选择文件上传')
        
        files = request.files.getlist('files')
        
        # 检查文件数量
        if len(files) < 3:
            return render_template('index.html', error='请至少上传3份报价单')
        
        # 分析价格（直接处理上传的文件流，避免保存到临时文件）
        analysis_result, vendor_stats, errors = analyze_prices_from_uploads(files)
        
        if errors:
            return render_template('index.html', error='\n'.join(errors))
        
        if analysis_result is None:
            return render_template('index.html', error='没有成功解析的报价单')
        
        # 转换分析结果为列表，用于前端展示
        price_results = analysis_result.to_dict('records')
        
        # 提取所有供应商名称，用于前端显示
        all_vendors = []
        if price_results:
            # 从第一个结果中提取供应商报价列
            first_item = price_results[0]
            for key in first_item:
                if key.startswith('报价_'):
                    vendor = key.replace('报价_', '')
                    all_vendors.append(vendor)
        print(f"提取的供应商列表：{all_vendors}")
        
        # 转换供应商中标统计为列表，用于前端展示，并按序号排序
        vendor_results = []
        # 定义排序键函数
        def sort_key(item):
            serial = item['序号']
            try:
                return (0, int(serial))
            except ValueError:
                try:
                    return (0, float(serial))
                except ValueError:
                    return (1, str(serial))
        
        for vendor, items in vendor_stats.items():
            # 按序号排序
            sorted_items = sorted(items, key=sort_key)
            for item in sorted_items:
                vendor_results.append({
                    '供应商': vendor,
                    '中标品名': item['物料名称'],
                    '对应序号': item['序号'],
                    '单项报价': item['单项报价'],
                    '数量': item.get('数量', 1),
                    '分项小计': item['分项小计']
                })
        
        return render_template('index.html', price_results=price_results, vendor_results=vendor_results, all_vendors=all_vendors)
    
    return render_template('index.html')


@app.route('/export')
def export():
    global analysis_result, vendor_stats
    
    if analysis_result is None:
        return redirect(url_for('index'))
    
    # 创建临时Excel文件
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        tmp_path = tmp.name
    
    try:
        # 创建Excel写入器
        with pd.ExcelWriter(tmp_path, engine='openpyxl') as writer:
            # 写入最低价分析结果
            analysis_result.to_excel(writer, sheet_name='最低价分析', index=False)
            
            # 为每个供应商创建单独的工作表
            for vendor, items in vendor_stats.items():
                # 按序号排序
                def sort_key(item):
                    serial = item['序号']
                    try:
                        return (0, int(serial))
                    except ValueError:
                        try:
                            return (0, float(serial))
                        except ValueError:
                            return (1, str(serial))
                
                sorted_items = sorted(items, key=sort_key)
                
                # 创建供应商工作表数据
                vendor_data = []
                for item in sorted_items:
                    vendor_data.append({
                        '对应序号': item['序号'],
                        '中标品名': item['物料名称'],
                        '单项报价': item['单项报价'],
                        '数量': item.get('数量', 1),
                        '分项小计': item['分项小计']
                    })
                
                # 创建供应商工作表
                vendor_df = pd.DataFrame(vendor_data)
                # 工作表名称不能超过31个字符
                sheet_name = vendor[:31]
                vendor_df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        # 发送文件
        return send_file(tmp_path, as_attachment=True, download_name='供应商比价分析报告.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    finally:
        # 清理临时文件
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except:
                pass


if __name__ == '__main__':
    # 创建templates目录（如果不存在）
    if not os.path.exists('templates'):
        os.makedirs('templates')
    app.run(debug=True, host='0.0.0.0', port=5000)
