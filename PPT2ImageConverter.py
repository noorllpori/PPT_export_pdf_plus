#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PPT 2K/4K 导出工具 - GUI 版本
支持拖拽到 exe 图标，可选择分辨率
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import threading
import os
import sys
import comtypes.client
import img2pdf


def check_powerpoint_installed():
    """
    检测是否安装了 Microsoft PowerPoint
    返回: (是否安装, 错误信息)
    """
    try:
        # 尝试创建 PowerPoint COM 对象
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        powerpoint.Quit()
        return True, None
    except Exception as e:
        error_msg = str(e)
        if "-2147221008" in error_msg or "80040154" in error_msg:
            return False, "未检测到 Microsoft PowerPoint\n\n本工具需要 PowerPoint 才能导出图片。\n请安装 Microsoft Office PowerPoint 后再试。"
        elif "-2147467262" in error_msg:
            return False, "PowerPoint COM 组件注册失败\n\n请尝试修复 Microsoft Office 安装。"
        else:
            return False, f"PowerPoint 检测失败:\n{error_msg}\n\n请确保已安装 Microsoft PowerPoint。"


class PPTConverterApp:
    # 分辨率配置
    RESOLUTIONS = {
        "1K (1280x720)": (1280, 720),
        "2K (2560x1440) - 默认": (2560, 1440),
        "4K (3840x2160)": (3840, 2160),
    }
    
    def __init__(self, root, initial_file=None):
        self.root = root
        self.root.title("PPT 转高清图片/PDF 工具")
        self.root.geometry("600x520")
        self.root.minsize(500, 420)
        
        self.initial_file = initial_file
        self.current_file = None
        
        self.setup_ui()
        
        # 如果有初始文件，自动填充
        if initial_file and os.path.exists(initial_file):
            self.root.after(100, lambda: self.load_file(initial_file))
    
    def setup_ui(self):
        """设置界面"""
        # 主容器
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(4, weight=1)
        
        # ===== 标题 =====
        title_label = ttk.Label(
            main_frame, 
            text="PPT 转高清图片/PDF", 
            font=("Microsoft YaHei", 16, "bold")
        )
        title_label.grid(row=0, column=0, pady=(0, 10))
        
        # ===== 文件选择区域 =====
        self.drop_frame = tk.Frame(
            main_frame, 
            bg="#e3f2fd", 
            highlightbackground="#2196f3", 
            highlightthickness=2,
            height=100
        )
        self.drop_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=10)
        self.drop_frame.grid_propagate(False)
        
        self.drop_label = tk.Label(
            self.drop_frame,
            text="点击选择 PPT 文件\n或拖拽文件到窗口",
            bg="#e3f2fd",
            font=("Microsoft YaHei", 11),
            fg="#1976d2"
        )
        self.drop_label.place(relx=0.5, rely=0.5, anchor="center")
        
        # 绑定点击事件
        self.drop_frame.bind("<Button-1>", self.select_file)
        self.drop_label.bind("<Button-1>", self.select_file)
        
        # ===== 文件信息显示 =====
        self.file_frame = ttk.LabelFrame(main_frame, text="文件信息", padding="10")
        self.file_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=10)
        self.file_frame.columnconfigure(0, weight=1)
        
        self.file_label = ttk.Label(
            self.file_frame, 
            text="未选择文件",
            wraplength=500,
            foreground="#999"
        )
        self.file_label.grid(row=0, column=0, sticky=tk.W)
        
        # ===== 设置区域 =====
        settings_frame = ttk.LabelFrame(main_frame, text="导出设置", padding="10")
        settings_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=10)
        settings_frame.columnconfigure(1, weight=1)
        
        # 分辨率选择
        ttk.Label(settings_frame, text="分辨率:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        
        self.resolution_var = tk.StringVar(value="2K (2560x1440) - 默认")
        self.resolution_combo = ttk.Combobox(
            settings_frame,
            textvariable=self.resolution_var,
            values=list(self.RESOLUTIONS.keys()),
            state="readonly",
            width=30
        )
        self.resolution_combo.grid(row=0, column=1, sticky=(tk.W, tk.E))
        self.resolution_combo.current(1)
        
        # DPI 设置
        ttk.Label(settings_frame, text="PDF DPI:").grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(10, 0))
        
        self.dpi_var = tk.IntVar(value=300)
        dpi_frame = ttk.Frame(settings_frame)
        dpi_frame.grid(row=1, column=1, sticky=tk.W, pady=(10, 0))
        
        ttk.Radiobutton(dpi_frame, text="150 (较小)", variable=self.dpi_var, value=150).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Radiobutton(dpi_frame, text="300 (推荐)", variable=self.dpi_var, value=300).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Radiobutton(dpi_frame, text="600 (高质量)", variable=self.dpi_var, value=600).pack(side=tk.LEFT)
        
        # 输出选项
        self.export_png_var = tk.BooleanVar(value=True)
        self.export_pdf_var = tk.BooleanVar(value=True)
        
        ttk.Checkbutton(
            settings_frame, 
            text="导出 PNG 图片", 
            variable=self.export_png_var
        ).grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=(10, 0))
        
        ttk.Checkbutton(
            settings_frame, 
            text="合并为 PDF", 
            variable=self.export_pdf_var
        ).grid(row=3, column=0, columnspan=2, sticky=tk.W)
        
        # ===== 日志区域 =====
        log_frame = ttk.LabelFrame(main_frame, text="进度日志", padding="5")
        log_frame.grid(row=4, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(
            log_frame, 
            height=8, 
            wrap=tk.WORD,
            font=("Consolas", 9)
        )
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.log_text.config(state=tk.DISABLED)
        
        # ===== 按钮区域 =====
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=5, column=0, sticky=(tk.W, tk.E), pady=(10, 0))
        btn_frame.columnconfigure(0, weight=1)
        
        self.start_btn = ttk.Button(
            btn_frame, 
            text="开始导出", 
            command=self.start_export,
            state=tk.DISABLED
        )
        self.start_btn.grid(row=0, column=1, padx=(0, 10))
        
        ttk.Button(
            btn_frame, 
            text="退出", 
            command=self.root.quit
        ).grid(row=0, column=2)
        
        # 版本信息
        version_label = ttk.Label(
            main_frame,
            text="v1.0 | 支持 .pptx .ppt 格式 | 拖拽文件到窗口即可",
            font=("Microsoft YaHei", 8),
            foreground="#999"
        )
        version_label.grid(row=6, column=0, pady=(10, 0))
    
    def select_file(self, event=None):
        """选择文件"""
        file_path = filedialog.askopenfilename(
            title="选择 PowerPoint 文件",
            filetypes=[
                ("PowerPoint 文件", "*.pptx *.ppt"),
                ("所有文件", "*.*")
            ]
        )
        if file_path:
            self.load_file(file_path)
    
    def load_file(self, file_path):
        """加载文件"""
        if not file_path or not os.path.exists(file_path):
            return
            
        self.current_file = file_path
        file_name = os.path.basename(file_path)
        
        try:
            file_size = os.path.getsize(file_path) / (1024 * 1024)
            size_str = f"{file_size:.2f} MB"
        except:
            size_str = "未知大小"
        
        self.file_label.config(
            text=f"文件名: {file_name}\n路径: {file_path}\n大小: {size_str}",
            foreground="#000"
        )
        
        # 更新选择区域样式
        self.drop_frame.config(bg="#c8e6c9", highlightbackground="#4caf50")
        self.drop_label.config(bg="#c8e6c9", text=f"已选择: {file_name}", fg="#2e7d32")
        
        self.start_btn.config(state=tk.NORMAL)
        self.log(f"已加载文件: {file_name}")
    
    def log(self, message):
        """添加日志"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.root.update_idletasks()
    
    def start_export(self):
        """开始导出"""
        if not self.current_file:
            messagebox.showwarning("提示", "请先选择 PPT 文件")
            return
        
        # 禁用按钮
        self.start_btn.config(state=tk.DISABLED, text="导出中...")
        self.log("")
        self.log("=" * 40)
        
        # 在新线程中执行导出
        thread = threading.Thread(target=self.export_worker, daemon=True)
        thread.start()
    
    def export_worker(self):
        """导出工作线程"""
        try:
            # 获取设置
            res_name = self.resolution_var.get()
            width, height = self.RESOLUTIONS[res_name]
            dpi = self.dpi_var.get()
            export_png = self.export_png_var.get()
            export_pdf = self.export_pdf_var.get()
            
            if not export_png and not export_pdf:
                self.root.after(0, lambda: messagebox.showwarning("提示", "请至少选择一种输出格式"))
                return
            
            pptx_path = self.current_file
            base_name = os.path.splitext(os.path.basename(pptx_path))[0]
            work_dir = os.path.dirname(pptx_path)
            res_short = res_name.split()[0]
            png_folder = os.path.join(work_dir, f"{base_name}_{res_short}")
            output_pdf = os.path.join(work_dir, f"{base_name}.pdf")
            
            self.root.after(0, lambda: self.log(f"开始导出: {base_name}"))
            self.root.after(0, lambda: self.log(f"分辨率: {width}x{height}"))
            self.root.after(0, lambda: self.log(f"PDF DPI: {dpi}"))
            self.root.after(0, lambda: self.log("=" * 40))
            
            slide_count = 0
            
            # 步骤 1: 导出 PNG
            if export_png:
                self.root.after(0, lambda: self.log("\n[步骤 1/2] 启动 PowerPoint..."))
                
                os.makedirs(png_folder, exist_ok=True)
                
                powerpoint = None
                presentation = None
                
                try:
                    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
                    powerpoint.Visible = 1  # 显示窗口（某些版本不支持隐藏）
                    
                    self.root.after(0, lambda: self.log("正在打开 PPT 文件..."))
                    presentation = powerpoint.Presentations.Open(pptx_path)
                    slide_count = presentation.Slides.Count
                    
                    self.root.after(0, lambda: self.log(f"共 {slide_count} 页，开始导出 PNG...\n"))
                    
                    for i in range(1, slide_count + 1):
                        slide = presentation.Slides(i)
                        output_file = os.path.join(png_folder, f"Slide_{i:02d}.png")
                        slide.Export(output_file, "PNG", width, height)
                        self.root.after(0, lambda i=i: self.log(f"  完成: Slide_{i:02d}.png"))
                    
                    presentation.Close()
                    powerpoint.Quit()
                    
                    self.root.after(0, lambda: self.log(f"\nPNG 导出完成"))
                    
                except Exception as e:
                    if presentation:
                        try:
                            presentation.Close()
                        except:
                            pass
                    if powerpoint:
                        try:
                            powerpoint.Quit()
                        except:
                            pass
                    raise e
            
            # 步骤 2: 合并 PDF
            if export_pdf and export_png:
                self.root.after(0, lambda: self.log("\n[步骤 2/2] 合并为 PDF..."))
                
                png_files = sorted([f for f in os.listdir(png_folder) 
                                   if f.lower().endswith('.png')])
                
                if not png_files:
                    self.root.after(0, lambda: self.log("错误: 没有找到 PNG 文件"))
                    return
                
                image_paths = [os.path.join(png_folder, f) for f in png_files]
                
                with open(output_pdf, "wb") as f:
                    f.write(img2pdf.convert(image_paths, dpi=dpi))
                
                file_size = os.path.getsize(output_pdf) / (1024 * 1024)
                self.root.after(0, lambda: self.log(f"PDF 生成完成"))
                self.root.after(0, lambda: self.log(f"  大小: {file_size:.2f} MB"))
            
            # 完成
            self.root.after(0, lambda: self.log("\n" + "=" * 40))
            self.root.after(0, lambda: self.log("导出完成！"))
            self.root.after(0, lambda: self.log("=" * 40))
            
            result_msg = f"导出完成！\n\n"
            if export_png:
                result_msg += f"PNG 文件夹: {png_folder}\n"
            if export_pdf:
                result_msg += f"PDF 文件: {output_pdf}\n"
            result_msg += f"页数: {slide_count}"
            
            self.root.after(0, lambda: messagebox.showinfo("完成", result_msg))
            
        except Exception as e:
            error_msg = str(e)
            self.root.after(0, lambda: self.log(f"\n错误: {error_msg}"))
            self.root.after(0, lambda: messagebox.showerror("错误", f"导出失败:\n{error_msg}"))
        
        finally:
            self.root.after(0, lambda: self.start_btn.config(state=tk.NORMAL, text="开始导出"))


def main():
    # 首先检测 PowerPoint 是否安装
    is_installed, error_msg = check_powerpoint_installed()
    if not is_installed:
        # 创建一个临时窗口显示错误
        root = tk.Tk()
        root.withdraw()  # 隐藏主窗口
        messagebox.showerror("错误 - 缺少 PowerPoint", error_msg)
        root.destroy()
        sys.exit(1)
    
    # 检查命令行参数（拖拽传入的文件）
    initial_file = None
    if len(sys.argv) > 1:
        # Windows 拖拽到 exe 上时，路径可能带引号
        initial_file = sys.argv[1].strip('"')
    
    # 创建窗口
    root = tk.Tk()
    app = PPTConverterApp(root, initial_file)
    root.mainloop()


if __name__ == "__main__":
    main()
