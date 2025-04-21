#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
import os
import subprocess
import pkg_resources

def check_python_version():
    """检查Python版本"""
    major, minor, micro = sys.version_info[:3]
    print(f"当前Python版本: {major}.{minor}.{micro}")
    if major < 3 or (major == 3 and minor < 6):
        print("错误: 需要Python 3.6或更高版本")
        return False
    return True

def check_dependencies():
    """检查依赖包是否已安装"""
    required_packages = [
        "python-pptx",
        "requests",
        "python-dotenv",
        "tqdm"
    ]
    
    missing_packages = []
    
    for package in required_packages:
        try:
            pkg_resources.get_distribution(package)
            print(f"✓ {package} 已安装")
        except pkg_resources.DistributionNotFound:
            missing_packages.append(package)
            print(f"✗ {package} 未安装")
    
    if missing_packages:
        print("\n缺少以下依赖包:")
        for package in missing_packages:
            print(f"  - {package}")
        print("\n请使用以下命令安装:")
        print("  pip install -r requirements.txt")
        return False
    
    return True

def check_env_file():
    """检查.env文件是否存在并包含API密钥"""
    if not os.path.exists(".env"):
        if os.path.exists(".env.example"):
            print("警告: 找不到.env文件，但已找到.env.example")
            print("请复制.env.example为.env并设置您的DeepSeek API密钥")
        else:
            print("错误: 找不到.env文件，请创建并设置您的DeepSeek API密钥")
        return False
    
    with open(".env", "r") as f:
        content = f.read()
        if "DEEPSEEK_API_KEY=" not in content or "your_api_key_here" in content:
            print("警告: .env文件中未设置有效的DeepSeek API密钥")
            return False
    
    print("✓ .env文件已配置")
    return True

def check_directories():
    """检查必要的目录是否存在"""
    required_dirs = [
        "专业词库"
    ]
    
    for directory in required_dirs:
        if not os.path.exists(directory):
            print(f"警告: 找不到 {directory} 目录")
            print(f"创建 {directory} 目录...")
            try:
                os.makedirs(directory)
                print(f"✓ 已创建 {directory} 目录")
            except Exception as e:
                print(f"错误: 无法创建 {directory} 目录: {str(e)}")
                return False
        else:
            print(f"✓ {directory} 目录已存在")
    
    return True

def main():
    """主函数"""
    print("========================================")
    print("         PPT翻译工具环境检查")
    print("========================================\n")
    
    all_checks_passed = True
    
    # 检查Python版本
    print("[1] 检查Python版本")
    python_check = check_python_version()
    all_checks_passed = all_checks_passed and python_check
    print()
    
    # 检查依赖包
    print("[2] 检查依赖包")
    dependencies_check = check_dependencies()
    all_checks_passed = all_checks_passed and dependencies_check
    print()
    
    # 检查环境文件
    print("[3] 检查环境配置")
    env_check = check_env_file()
    all_checks_passed = all_checks_passed and env_check
    print()
    
    # 检查目录
    print("[4] 检查目录结构")
    dir_check = check_directories()
    all_checks_passed = all_checks_passed and dir_check
    print()
    
    # 结果总结
    print("========================================")
    if all_checks_passed:
        print("✓ 所有检查通过！可以启动PPT翻译工具。")
        print("  运行 python gui.py 或双击 start_translator.bat 开始使用")
    else:
        print("✗ 检查未全部通过，请解决上述问题后再启动翻译工具")
    print("========================================")
    
    input("\n按回车键继续...")
    
    return 0 if all_checks_passed else 1

if __name__ == "__main__":
    sys.exit(main()) 