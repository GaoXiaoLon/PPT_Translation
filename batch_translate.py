#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import argparse
import glob
import time
from tqdm import tqdm
from dotenv import load_dotenv
from translator import DeepSeekTranslator, PPTTranslator

# 加载环境变量
load_dotenv()

def get_all_ppt_files(directory, recursive=False):
    """
    获取目录中所有PPT文件
    
    Args:
        directory (str): 目录路径
        recursive (bool): 是否递归查找子目录
        
    Returns:
        list: PPT文件路径列表
    """
    if recursive:
        pattern = os.path.join(directory, "**", "*.ppt*")
        files = glob.glob(pattern, recursive=True)
    else:
        pattern = os.path.join(directory, "*.ppt*")
        files = glob.glob(pattern)
    
    return [f for f in files if f.endswith(('.ppt', '.pptx'))]

def batch_translate(input_files, output_dir, source_lang, target_lang, domain, api_key, model):
    """
    批量翻译PPT文件
    
    Args:
        input_files (list): 输入PPT文件路径列表
        output_dir (str): 输出目录
        source_lang (str): 源语言
        target_lang (str): 目标语言
        domain (str): 专业领域
        api_key (str): API密钥
        model (str): 模型名称
    """
    # 确保输出目录存在
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # 初始化翻译器
    translator = DeepSeekTranslator(api_key=api_key, model=model)
    ppt_translator = PPTTranslator(
        translator, 
        source_lang=source_lang, 
        target_lang=target_lang,
        domain=domain
    )
    
    print(f"\n=== 批量翻译开始 ===")
    print(f"源语言: {source_lang}")
    print(f"目标语言: {target_lang}")
    print(f"专业领域: {domain}")
    print(f"总文件数: {len(input_files)}")
    print(f"输出目录: {output_dir}")
    print("=" * 50)
    
    # 记录开始时间
    start_time = time.time()
    
    # 处理每个文件
    for i, input_file in enumerate(tqdm(input_files, desc="翻译进度")):
        try:
            # 生成输出文件路径
            file_name = os.path.basename(input_file)
            name, ext = os.path.splitext(file_name)
            output_file = os.path.join(output_dir, f"{name}_translated{ext}")
            
            print(f"\n[{i+1}/{len(input_files)}] 翻译文件: {file_name}")
            
            # 翻译PPT
            ppt_translator.translate_ppt(input_file, output_file)
            
            print(f"  ✓ 已保存至: {output_file}")
            
        except Exception as e:
            print(f"  ✗ 翻译失败: {str(e)}")
    
    # 计算总耗时
    elapsed_time = time.time() - start_time
    minutes, seconds = divmod(elapsed_time, 60)
    hours, minutes = divmod(minutes, 60)
    
    print("\n=== 批量翻译完成 ===")
    print(f"总耗时: {int(hours)}小时 {int(minutes)}分钟 {seconds:.2f}秒")
    print(f"已处理文件数: {len(input_files)}")
    print(f"输出目录: {output_dir}")
    print("=" * 50)

def main():
    """主函数"""
    parser = argparse.ArgumentParser(description="PPT批量翻译工具 - 使用DeepSeek API进行专业翻译")
    parser.add_argument("--input-dir", required=True, help="输入PPT文件目录")
    parser.add_argument("--output-dir", default="translated_ppts", help="输出PPT文件目录")
    parser.add_argument("--source", default="en", help="源语言代码，默认为'en'")
    parser.add_argument("--target", default="zh", help="目标语言代码，默认为'zh'")
    parser.add_argument("--api-key", help="DeepSeek API密钥（可选，也可在.env文件中设置）")
    parser.add_argument("--model", default="deepseek-chat", help="DeepSeek模型名称")
    parser.add_argument("--domain", choices=["computer", "os", "general"], default="general",
                        help="专业领域，可选值: computer (计算机), os (操作系统), general (通用)")
    parser.add_argument("--recursive", action="store_true", help="递归查找子目录中的文件")
    parser.add_argument("--filter", default="*.ppt*", help="文件过滤模式，默认为'*.ppt*'")
    
    args = parser.parse_args()
    
    # 检查输入目录
    if not os.path.exists(args.input_dir) or not os.path.isdir(args.input_dir):
        print(f"错误: 输入目录不存在: {args.input_dir}")
        return
    
    # 获取所有PPT文件
    input_files = get_all_ppt_files(args.input_dir, args.recursive)
    
    if not input_files:
        print(f"错误: 在目录 {args.input_dir} 中未找到PPT文件")
        return
    
    # 获取API密钥
    api_key = args.api_key or os.getenv("DEEPSEEK_API_KEY")
    if not api_key:
        print("错误: DeepSeek API密钥未提供。请在.env文件中设置DEEPSEEK_API_KEY或使用--api-key参数")
        return
    
    try:
        # 批量翻译
        batch_translate(
            input_files=input_files,
            output_dir=args.output_dir,
            source_lang=args.source,
            target_lang=args.target,
            domain=args.domain,
            api_key=api_key,
            model=args.model
        )
        
    except Exception as e:
        print(f"批量翻译过程中出错: {e}")

if __name__ == "__main__":
    main() 