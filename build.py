#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PPT2ImageConverter 打包脚本
生成独立可执行文件
"""

import subprocess
import sys
import os
import shutil


def check_dependencies():
    """检查依赖是否安装"""
    print("[1/4] 检查依赖...")
    
    required = ["pyinstaller", "comtypes", "img2pdf", "Pillow"]
    missing = []
    
    for pkg in required:
        try:
            if pkg == "Pillow":
                __import__("PIL")
            else:
                __import__(pkg.lower())
        except ImportError:
            missing.append(pkg)
    
    if missing:
        print(f"  缺少依赖: {', '.join(missing)}")
        print("  正在安装...")
        subprocess.run([sys.executable, "-m", "pip", "install", "-q"] + missing)
        print("  依赖安装完成")
    else:
        print("  所有依赖已安装")
    print()


def build_exe():
    """构建 exe"""
    print("[2/4] 构建可执行文件...")
    print("  这可能需要几分钟，请耐心等待...")
    print()
    
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--name", "PPT2ImageConverter",
        "--onefile",
        "--windowed",
        "--clean",
        "--noconfirm",
        # 隐藏导入
        "--hidden-import", "comtypes",
        "--hidden-import", "comtypes.client",
        "--hidden-import", "img2pdf",
        "--hidden-import", "PIL",
        "--hidden-import", "PIL.Image",
        "--hidden-import", "tkinter",
        "--hidden-import", "tkinter.filedialog",
        "--hidden-import", "tkinter.scrolledtext",
        "PPT2ImageConverter.py"
    ]
    
    result = subprocess.run(cmd, capture_output=True, text=True)
    
    if result.returncode == 0:
        print("  构建成功!")
        return True
    else:
        print("  构建失败!")
        print(result.stderr[-2000:] if len(result.stderr) > 2000 else result.stderr)
        return False


def create_package():
    """创建分发包"""
    print("[3/4] 创建分发包...")
    
    # 清理旧文件
    package_dir = "PPT2ImageConverter_分发包"
    if os.path.exists(package_dir):
        shutil.rmtree(package_dir)
    os.makedirs(package_dir)
    
    # 复制 exe
    exe_source = os.path.join("dist", "PPT2ImageConverter.exe")
    exe_target = os.path.join(package_dir, "PPT2ImageConverter.exe")
    
    if not os.path.exists(exe_source):
        print(f"  错误: 未找到 {exe_source}")
        return False
    
    shutil.copy2(exe_source, exe_target)
    print(f"  复制: PPT2ImageConverter.exe")
    
    # 创建使用说明
    readme_content = """PPT 转高清图片/PDF 工具
========================

使用方法
--------
方法1 - 拖拽使用（推荐）:
  选中 PPT 文件，拖拽到 PPT2ImageConverter.exe 图标上

方法2 - 先打开程序:
  1. 双击运行 PPT2ImageConverter.exe
  2. 点击"选择文件"按钮，选择 PPT
  3. 点击"开始导出"

分辨率选项
--------
- 1K (1280x720)  - 快速预览
- 2K (2560x1440) - 日常使用（默认）
- 4K (3840x2160) - 高质量展示

输出文件
--------
转换完成后，在原 PPT 同目录下生成：
- 文件名_2K/     - PNG 图片文件夹
- 文件名.pdf     - 合并后的 PDF

系统要求
--------
- Windows 7/10/11
- Microsoft PowerPoint 2010+
- 无需安装 Python 或其他环境

注意事项
--------
- 导出过程中请勿关闭 PowerPoint
- 大文件导出可能需要较长时间
- 导出的图片保持原始质量，无压缩
"""
    
    readme_path = os.path.join(package_dir, "使用说明.txt")
    with open(readme_path, "w", encoding="utf-8") as f:
        f.write(readme_content)
    print(f"  创建: 使用说明.txt")
    
    # 打包为 zip
    zip_name = "PPT2ImageConverter_分发包"
    zip_path = zip_name + ".zip"
    
    if os.path.exists(zip_path):
        os.remove(zip_path)
    
    shutil.make_archive(zip_name, "zip", package_dir)
    print(f"  打包: {zip_path}")
    
    return True


def clean_build_files():
    """清理构建临时文件"""
    print("[4/4] 清理临时文件...")
    
    dirs_to_remove = ["build", "dist"]
    files_to_remove = ["PPT2ImageConverter.spec"]
    
    for d in dirs_to_remove:
        if os.path.exists(d):
            shutil.rmtree(d)
            print(f"  删除: {d}/")
    
    for f in files_to_remove:
        if os.path.exists(f):
            os.remove(f)
            print(f"  删除: {f}")


def main():
    print("=" * 50)
    print("PPT2ImageConverter 打包工具")
    print("=" * 50)
    print()
    
    # 执行步骤
    check_dependencies()
    
    if build_exe():
        if create_package():
            clean_build_files()
            
            print()
            print("=" * 50)
            print("打包完成!")
            print()
            print("分发包文件:")
            print("  PPT2ImageConverter_分发包/")
            print("  ├── PPT2ImageConverter.exe  (主程序)")
            print("  └── 使用说明.txt")
            print()
            print("  PPT2ImageConverter_分发包.zip  (可直接发送)")
            print()
            print("使用方式:")
            print("  1. 解压分发包")
            print("  2. 拖拽 PPT 文件到 exe 图标上")
            print("  3. 或双击 exe 后选择文件")
            print("=" * 50)
        else:
            print()
            print("创建分发包失败")
    else:
        print()
        print("构建失败，请检查错误信息")
    
    print()
    input("按回车键退出...")


if __name__ == "__main__":
    main()
