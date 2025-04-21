#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import argparse
import requests
import json
from tqdm import tqdm
from dotenv import load_dotenv
try:
    from pptx import Presentation
except ImportError:
    print("错误: 找不到python-pptx库。请使用'pip install python-pptx'安装。")
    Presentation = None
from concurrent.futures import ThreadPoolExecutor
from terminology import terminology_manager

# 加载环境变量
load_dotenv()

class DeepSeekTranslator:
    """使用DeepSeek API进行翻译的类"""
    
    def __init__(self, api_key=None, model="deepseek-chat"):
        self.api_key = api_key or os.getenv("DEEPSEEK_API_KEY")
        if not self.api_key:
            raise ValueError("DeepSeek API密钥未提供。请在.env文件中设置DEEPSEEK_API_KEY或作为参数传递。")
        
        self.model = model
        self.api_url = "https://api.deepseek.com/v1/chat/completions"
        self.headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {self.api_key}"
        }

    def translate(self, text, source_lang="en", target_lang="zh", domain=None):
        """
        使用DeepSeek API翻译文本
        
        Args:
            text (str): 要翻译的文本
            source_lang (str): 源语言代码
            target_lang (str): 目标语言代码 
            domain (str, optional): 专业领域
        
        Returns:
            str: 翻译后的文本
        """
        if not text or text.strip() == "":
            return text
            
        # 加载领域术语库
        if domain:
            terminology_manager.load_terminology(domain)
            
        # 构建专业提示
        domain_prompt = ""
        if domain:
            domain_prompts = {
                "computer": "你是一位精通计算机科学专业术语的翻译专家，",
                "os": "你是一位精通操作系统专业术语的翻译专家，",
                "general": ""
            }
            domain_prompt = domain_prompts.get(domain, f"你是一位精通{domain}领域专业术语的翻译专家，")
        
        system_prompt = f"{domain_prompt}请将以下{source_lang}文本翻译成{target_lang}，保持专业准确，同时保留原文的格式和标点符号。翻译时，专业术语应使用目标语言中对应的标准术语。"
        
        # 添加术语库术语到提示中
        terminology_dict = terminology_manager.terminology_dict
        if terminology_dict and len(terminology_dict) > 0:
            terms = []
            for k, v in terminology_dict.items():
                if k in text:  # 只添加文本中出现的术语
                    terms.append(f"{k} = {v}")
            
            if terms:
                terminology_prompt = "请特别注意以下专业术语的翻译：\n" + "\n".join(terms)
                system_prompt += "\n\n" + terminology_prompt
        
        data = {
            "model": self.model,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": text}
            ],
            "temperature": 0.3,  # 低温度以确保翻译的确定性
            "max_tokens": 4096
        }
        
        try:
            response = requests.post(
                self.api_url,
                headers=self.headers,
                data=json.dumps(data)
            )
            response.raise_for_status()
            result = response.json()
            translated_text = result["choices"][0]["message"]["content"]
            
            # 使用术语库增强翻译结果
            enhanced_translation = terminology_manager.enhance_translation(text, translated_text)
            
            return enhanced_translation
        except Exception as e:
            print(f"翻译时出错: {e}")
            return text  # 出错时返回原文

class PPTTranslator:
    """PPT文件翻译器类"""
    
    def __init__(self, translator, source_lang="en", target_lang="zh", domain=None):
        if Presentation is None:
            raise ImportError("请安装python-pptx库以使用此功能")
        
        self.translator = translator
        self.source_lang = source_lang
        self.target_lang = target_lang
        self.domain = domain
        
    def translate_ppt(self, input_file, output_file, progress_callback=None):
        """
        翻译PPT文件
        
        Args:
            input_file (str): 输入PPT文件路径
            output_file (str): 输出PPT文件路径
            progress_callback (function, optional): 进度回调函数
        """
        # 加载PPT
        prs = Presentation(input_file)
        total_slides = len(prs.slides)
        
        print(f"开始翻译PPT: {input_file}")
        print(f"总计 {total_slides} 张幻灯片")
        
        # 遍历所有幻灯片和形状进行翻译
        for i, slide in enumerate(tqdm(prs.slides, desc="翻译幻灯片")):
            # 更新进度
            if progress_callback:
                progress_callback(i+1, total_slides)
                
            for shape in slide.shapes:
                if hasattr(shape, "text_frame") and shape.text_frame:
                    if shape.text_frame.text.strip():
                        # 收集所有段落文本
                        paragraphs = []
                        for paragraph in shape.text_frame.paragraphs:
                            # 保存原始段落格式信息
                            para_text = paragraph.text
                            if para_text.strip():
                                paragraphs.append(para_text)
                        
                        # 批量翻译所有段落文本
                        if paragraphs:
                            text_to_translate = "\n".join(paragraphs)
                            translated_text = self.translator.translate(
                                text_to_translate, 
                                self.source_lang, 
                                self.target_lang,
                                self.domain
                            )
                            
                            # 分割翻译结果并更新段落
                            translated_paragraphs = translated_text.split('\n')
                            for j, paragraph in enumerate(shape.text_frame.paragraphs):
                                if j < len(translated_paragraphs) and paragraph.text.strip():
                                    # 保留原始格式
                                    for run_idx, run in enumerate(paragraph.runs):
                                        if run_idx == 0 and translated_paragraphs[j].strip():
                                            run.text = translated_paragraphs[j]
                                        elif run_idx > 0:
                                            run.text = ""
                
                # 处理表格文本
                if hasattr(shape, "table"):
                    for row in shape.table.rows:
                        for cell in row.cells:
                            if cell.text.strip():
                                cell.text = self.translator.translate(
                                    cell.text, 
                                    self.source_lang, 
                                    self.target_lang,
                                    self.domain
                                )
        
        # 保存翻译后的PPT
        prs.save(output_file)
        print(f"翻译完成，已保存至: {output_file}")
        
def main():
    """主函数"""
    parser = argparse.ArgumentParser(description="PPT翻译工具 - 使用DeepSeek API进行专业翻译")
    parser.add_argument("--input", required=True, help="输入PPT文件路径")
    parser.add_argument("--output", required=True, help="输出PPT文件路径")
    parser.add_argument("--source", default="en", help="源语言代码，默认为'en'")
    parser.add_argument("--target", default="zh", help="目标语言代码，默认为'zh'")
    parser.add_argument("--api-key", help="DeepSeek API密钥（可选，也可在.env文件中设置）")
    parser.add_argument("--model", default="deepseek-chat", help="DeepSeek模型名称")
    parser.add_argument("--domain", choices=["computer", "os", "general"], 
                        help="专业领域，可选值: computer (计算机), os (操作系统), general (通用)")
    
    args = parser.parse_args()
    
    # 检查输入文件
    if not os.path.exists(args.input):
        print(f"错误: 输入文件不存在: {args.input}")
        return
    
    try:
        # 初始化翻译器
        translator = DeepSeekTranslator(api_key=args.api_key, model=args.model)
        ppt_translator = PPTTranslator(
            translator, 
            source_lang=args.source, 
            target_lang=args.target,
            domain=args.domain
        )
        
        # 翻译PPT
        ppt_translator.translate_ppt(args.input, args.output)
        
    except ImportError as e:
        print(f"导入错误: {e}")
        print("请安装必要的依赖: pip install python-pptx")
    except Exception as e:
        print(f"翻译过程中出错: {e}")

if __name__ == "__main__":
    main() 