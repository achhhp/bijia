#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import json
import sys
from collections import defaultdict

# 节点类型与功能的映射字典
NODE_FUNCTIONS = {
    # 核心节点
    "UpscaleModelLoader": "加载图像放大模型，用于提升图像分辨率和质量",
    "KSampler": "采样器节点，负责根据条件生成或处理潜在空间表示(latent)",
    "VAEEncode": "将图像编码为潜在空间表示(latent)",
    "VAEDecode": "将潜在空间表示(latent)解码为图像",
    "CLIPLoader": "加载CLIP模型，用于文本条件处理",
    "PreviewImage": "预览生成或处理后的图像",
    "Image Save": "保存图像到指定路径",
    
    # 图像处理节点
    "easy imageBatchToImageList": "将图像批次转换为图像列表",
    "ImageListToImageBatch": "将图像列表转换为图像批次",
    
    # TTP工具集节点
    "TTP_Tile_image_size": "计算图像分块大小，用于分块处理大图像",
    "TTP_Image_Tile_Batch": "将图像分割成多个块，便于并行处理",
    "TTP_Image_Assy": "将处理后的图像块重新组装成完整图像",
    
    # 内存管理节点
    "easy cleanGpuUsed": "清理GPU内存，优化资源使用",
    
    # 模型优化节点
    "TeaCache": "模型缓存优化，提高模型加载和推理效率",
    "PathchSageAttentionKJ": "路径感知注意力优化，提升模型性能",
    "ModelSamplingSD3": "SD3模型采样参数设置",
    
    # 条件处理节点
    "FluxGuidance": "流量引导控制，调整条件引导强度",
    "ConditioningZeroOut": "条件归零处理",
    "ReferenceLatent": "参考潜在空间表示，用于参考图像引导生成",
    
    # 其他功能节点
    "Image Comparer (rgthree)": "图像比较器，用于直观比较两张图像的差异"
}

# 节点组信息分析
def analyze_node_groups(workflow_data):
    groups = workflow_data.get("groups", [])
    print(f"\n=== 节点组分析 ({len(groups)}个组) ===")
    for group in groups:
        group_id = group.get("id", "未知")
        title = group.get("title", "未命名组")
        print(f"  组 {group_id}: {title}")

# 分析节点之间的连接关系
def analyze_node_connections(workflow_data):
    links = workflow_data.get("links", [])
    nodes = {node["id"]: node for node in workflow_data.get("nodes", [])}
    
    # 构建输入输出连接映射
    input_connections = defaultdict(list)
    output_connections = defaultdict(list)
    
    # 简化版连接分析（因为完整格式较复杂）
    print(f"\n=== 连接分析 ({len(links)}个连接) ===")
    print("  注意：完整连接关系较复杂，这里只显示概览")
    
    # 统计每个节点的输入输出数量
    for node_id, node in nodes.items():
        input_count = len(node.get("inputs", []))
        output_count = len(node.get("outputs", []))
        print(f"  节点 {node_id} ({node.get('type', '未知')}): {input_count}个输入, {output_count}个输出")

# 主分析函数
def analyze_workflow(filename):
    try:
        with open(filename, 'r', encoding='utf-8') as f:
            workflow_data = json.load(f)
        
        print(f"\n===== ComfyUI工作流分析: {filename} =====")
        
        # 工作流基本信息
        print("\n=== 基本信息 ===")
        print(f"  ID: {workflow_data.get('id', '未知')}")
        print(f"  版本: {workflow_data.get('version', '未知')}")
        print(f"  最后节点ID: {workflow_data.get('last_node_id', '未知')}")
        print(f"  最后链接ID: {workflow_data.get('last_link_id', '未知')}")
        
        # 节点概览
        nodes = workflow_data.get("nodes", [])
        print(f"\n=== 节点概览 ({len(nodes)}个节点) ===")
        
        # 按节点类型分组
        nodes_by_type = defaultdict(list)
        for node in nodes:
            node_type = node.get("type", "未知类型")
            nodes_by_type[node_type].append(node)
        
        # 打印每种类型的节点数量
        print("  节点类型统计:")
        for node_type, node_list in sorted(nodes_by_type.items(), key=lambda x: len(x[1]), reverse=True):
            print(f"    - {node_type}: {len(node_list)}个")
        
        # 详细节点分析
        print("\n=== 详细节点分析 ===")
        for node in nodes:
            node_id = node.get("id", "未知")
            node_type = node.get("type", "未知类型")
            node_function = NODE_FUNCTIONS.get(node_type, "未知功能")
            
            # 获取节点配置值
            widgets_values = node.get("widgets_values", [])
            
            print(f"\n  节点 {node_id}: {node_type}")
            print(f"    功能: {node_function}")
            
            # 打印关键配置值
            if widgets_values:
                print(f"    主要配置: {widgets_values[:3]} {'...' if len(widgets_values) > 3 else ''}")
            
            # 分析输入输出
            inputs = node.get("inputs", [])
            outputs = node.get("outputs", [])
            print(f"    输入端口: {len(inputs)}个")
            print(f"    输出端口: {len(outputs)}个")
        
        # 分析节点组
        analyze_node_groups(workflow_data)
        
        # 分析连接关系
        analyze_node_connections(workflow_data)
        
        # 工作流程分析
        print("\n=== 工作流程分析 ===")
        print("  这是一个图像放大工作流，主要包含以下流程:")
        print("  1. 图像预处理：加载图像并进行分块")
        print("  2. 模型加载：加载CLIP、VAE和放大模型")
        print("  3. 图像编码：将图像转换为潜在空间表示")
        print("  4. 潜在空间处理：使用KSampler进行优化采样")
        print("  5. 图像解码：将处理后的潜在表示转回图像")
        print("  6. 图像重组：将分块处理的图像重新组装")
        print("  7. 结果输出：预览和保存处理后的图像")
        
        # 特殊功能节点说明
        print("\n=== 关键节点详细说明 ===")
        if "TTP_Image_Tile_Batch" in nodes_by_type:
            print("  TTP_Image_Tile_Batch (图像分块):")
            print("    - 功能：将大图像分割成多个小块，便于GPU处理")
            print("    - 优势：可以处理超过GPU内存限制的大图像")
        
        if "UpscaleModelLoader" in nodes_by_type:
            print("  UpscaleModelLoader (放大模型加载):")
            print("    - 功能：加载专用的图像放大模型")
            print("    - 配置：模型名称等参数通过widgets_values设置")
        
        if "KSampler" in nodes_by_type:
            print("  KSampler (采样器):")
            print("    - 功能：在潜在空间中进行采样和优化")
            print("    - 参数：步数、CFG、采样器类型等影响生成质量")
        
        print("\n工作流分析完成！")
        
    except Exception as e:
        print(f"分析过程中出现错误: {e}")

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("用法: python analyze_comfyui_workflow.py <workflow_file.json>")
        sys.exit(1)
    
    filename = sys.argv[1]
    analyze_workflow(filename)