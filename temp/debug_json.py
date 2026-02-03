# 用于调试JSON文件的脚本
import json
import sys
import os

def validate_json(file_path):
    """验证JSON文件格式并返回解析后的数据"""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            # 首先尝试完整读取和解析
            try:
                data = json.load(f)
                print("✓ JSON文件格式正确！")
                return True, data
            except json.JSONDecodeError as e:
                # 如果解析失败，尝试定位具体问题
                print(f"✗ JSON格式错误: {e}")
                print(f"错误位置: 第{e.lineno}行, 第{e.colno}列")
                
                # 重新打开文件并显示错误行附近的内容
                f.seek(0)
                lines = f.readlines()
                error_line = e.lineno - 1  # 转换为0-based索引
                
                # 显示错误行及其前后几行
                start = max(0, error_line - 2)
                end = min(len(lines), error_line + 3)
                
                print("\n错误行附近内容:")
                for i in range(start, end):
                    prefix = "-> " if i == error_line else "   "
                    print(f"{prefix}{i+1}: {lines[i].rstrip()}")
                
                if error_line < len(lines):
                    # 显示错误位置的字符
                    if e.colno <= len(lines[error_line]):
                        print(f"\n错误字符: '{lines[error_line][e.colno-1]}'")
                
                return False, None
    except FileNotFoundError:
        print(f"✗ 文件不存在: {file_path}")
        return False, None
    except UnicodeDecodeError:
        print(f"✗ 文件编码错误，请确保文件使用UTF-8编码")
        return False, None
    except Exception as e:
        print(f"✗ 读取文件时出错: {e}")
        return False, None

def check_structure(data, max_depth=3, current_depth=0):
    """检查并打印JSON的结构"""
    indent = "  " * current_depth
    
    if current_depth >= max_depth:
        if isinstance(data, dict):
            print(f"{indent}... 包含 {len(data)} 个键 (已截断)")
        elif isinstance(data, list):
            print(f"{indent}... 包含 {len(data)} 个元素 (已截断)")
        return
    
    if isinstance(data, dict):
        print(f"{indent}对象 (包含 {len(data)} 个键):")
        for key, value in data.items():
            print(f"{indent}  '{key}':", end=" ")
            if isinstance(value, dict):
                print(f"对象 (包含 {len(value)} 个键)")
                check_structure(value, max_depth, current_depth + 1)
            elif isinstance(value, list):
                print(f"数组 (包含 {len(value)} 个元素)")
                if value:
                    print(f"{indent}  首个元素类型: {type(value[0]).__name__}")
                    if len(value) > 1:
                        print(f"{indent}  最后元素类型: {type(value[-1]).__name__}")
                    # 只检查数组的第一个元素来避免输出过多
                    if isinstance(value[0], (dict, list)):
                        check_structure(value[0], max_depth, current_depth + 1)
            else:
                print(f"{type(value).__name__}")
                if isinstance(value, str) and len(value) > 50:
                    print(f"{indent}    值预览: '{value[:50]}...'")
    elif isinstance(data, list):
        print(f"{indent}数组 (包含 {len(data)} 个元素):")
        if data:
            print(f"{indent}  首个元素类型: {type(data[0]).__name__}")
            if len(data) > 1:
                print(f"{indent}  最后元素类型: {type(data[-1]).__name__}")
            # 只检查数组的第一个元素来避免输出过多
            if isinstance(data[0], (dict, list)):
                check_structure(data[0], max_depth, current_depth + 1)
    else:
        print(f"{indent}{type(data).__name__}")

def create_sample_json():
    """创建一个示例JSON文件用于演示"""
    sample_data = {
        "id": "example-id",
        "nodes": [
            {
                "id": 1,
                "type": "example_node",
                "pos": [100, 200],
                "inputs": [],
                "outputs": []
            }
        ],
        "links": []
    }
    
    with open("sample.json", "w", encoding="utf-8") as f:
        json.dump(sample_data, f, ensure_ascii=False, indent=2)
    
    print("已创建示例JSON文件: sample.json")

def find_common_json_errors(data_str):
    """查找常见的JSON错误"""
    common_errors = []
    
    # 检查是否缺少逗号
    lines = data_str.splitlines()
    for i, line in enumerate(lines):
        stripped = line.strip()
        # 检查是否有对象属性后缺少逗号
        if ':' in stripped and not stripped.endswith(',') and not stripped.endswith('}') and not stripped.endswith(']'):
            # 检查下一行是否是新的属性或结束括号
            if i + 1 < len(lines):
                next_stripped = lines[i+1].strip()
                if next_stripped and next_stripped[0] in '{["' and not next_stripped.startswith('//'):
                    common_errors.append(f"第{i+1}行: 可能缺少逗号")
    
    # 检查是否有多余的逗号
    if ',}' in data_str or ',]' in data_str:
        common_errors.append("发现多余的逗号在对象或数组末尾")
    
    # 检查字符串是否正确闭合
    quote_count = data_str.count('"')
    if quote_count % 2 != 0:
        common_errors.append("引号不匹配，可能有未闭合的字符串")
    
    # 检查括号是否匹配
    open_braces = data_str.count('{')
    close_braces = data_str.count('}')
    if open_braces != close_braces:
        common_errors.append(f"大括号不匹配: 打开{open_braces}个，关闭{close_braces}个")
    
    open_brackets = data_str.count('[')
    close_brackets = data_str.count(']')
    if open_brackets != close_brackets:
        common_errors.append(f"方括号不匹配: 打开{open_brackets}个，关闭{close_brackets}个")
    
    return common_errors

def main():
    print("==== JSON文件调试工具 ====\n")
    
    # 显示当前目录的JSON文件
    print("当前目录中的JSON文件:")
    json_files = [f for f in os.listdir('.') if f.endswith('.json')]
    if json_files:
        for i, file in enumerate(json_files, 1):
            print(f"{i}. {file}")
    else:
        print("- 当前目录没有JSON文件")
    
    print("\n使用说明:")
    print("1. 将需要调试的JSON文件复制到当前目录")
    print("2. 运行: python debug_json.py <filename.json>")
    print("3. 或者运行: python debug_json.py --sample 创建示例文件")
    print("\n注意: 由于权限限制，无法直接访问工作目录外的文件")
    
    # 处理命令行参数
    if len(sys.argv) > 1:
        if sys.argv[1] == '--sample':
            create_sample_json()
        else:
            file_path = sys.argv[1]
            print(f"\n正在验证文件: {file_path}")
            is_valid, data = validate_json(file_path)
            
            if is_valid and data:
                print("\nJSON结构分析:")
                check_structure(data)
                
                # 尝试查找常见错误
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                common_errors = find_common_json_errors(content)
                if common_errors:
                    print("\n可能的问题:")
                    for error in common_errors:
                        print(f"- {error}")
                else:
                    print("\n未发现常见的JSON格式问题")

if __name__ == "__main__":
    main()