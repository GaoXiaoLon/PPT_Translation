#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import re

class TerminologyManager:
    """专业术语管理类"""
    
    def __init__(self):
        self.terminology_dirs = ["专业词库"]
        self.terminology_dict = {}
        self.initialized = False
    
    def load_terminology(self, domain=None):
        """
        加载专业术语库
        
        Args:
            domain (str, optional): 专业领域
        """
        if self.initialized and domain is None:
            return self.terminology_dict
            
        result = {}
        
        for dir_path in self.terminology_dirs:
            if not os.path.exists(dir_path):
                continue
                
            files = []
            if domain:
                # 加载特定领域的术语库
                domain_file = f"{domain}_terms.txt"
                domain_path = os.path.join(dir_path, domain_file)
                if os.path.exists(domain_path):
                    files.append(domain_path)
            else:
                # 加载所有术语库
                for file in os.listdir(dir_path):
                    if file.endswith("_terms.txt"):
                        files.append(os.path.join(dir_path, file))
        
            # 解析术语文件
            for file_path in files:
                with open(file_path, 'r', encoding='utf-8') as f:
                    for line in f:
                        line = line.strip()
                        if not line or line.startswith('#'):
                            continue
                            
                        parts = line.split('=', 1)
                        if len(parts) == 2:
                            term = parts[0].strip()
                            translation = parts[1].strip()
                            result[term] = translation
        
        self.terminology_dict = result
        self.initialized = True
        return result
    
    def get_translation(self, term, default=None):
        """
        获取术语的翻译
        
        Args:
            term (str): 术语
            default: 默认返回值
            
        Returns:
            str: 翻译结果
        """
        if not self.initialized:
            self.load_terminology()
            
        return self.terminology_dict.get(term, default)
    
    def enhance_translation(self, text, translation):
        """
        使用术语库增强翻译结果
        
        Args:
            text (str): 原文
            translation (str): 翻译结果
            
        Returns:
            str: 增强后的翻译
        """
        if not self.initialized:
            self.load_terminology()
            
        # 如果术语库为空，直接返回原翻译
        if not self.terminology_dict:
            return translation
            
        # 在原文中查找术语并替换翻译中对应部分
        enhanced = translation
        for term, term_translation in self.terminology_dict.items():
            # 使用正则表达式查找完整的词
            pattern = r'\b' + re.escape(term) + r'\b'
            if re.search(pattern, text, re.IGNORECASE):
                # 在翻译中寻找可能的错误翻译并替换为标准术语
                # 这里采用简单的替换策略，实际应用中可能需要更复杂的算法
                enhanced = re.sub(pattern, term_translation, enhanced, flags=re.IGNORECASE)
                
        return enhanced

# 创建全局术语管理器实例
terminology_manager = TerminologyManager() 