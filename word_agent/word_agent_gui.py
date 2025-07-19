"""
Word Agent GUI - 现代化的Word编辑助手图形界面
使用tkinter构建，提供用户友好的界面
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import os
import json
from typing import Optional
import sys

# 导入agent相关模块
from llm_agent import LLMWordAgent, LLMConfig


class WordAgentGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Word Agent - AI驱动的Word编辑助手")
        self.root.geometry("900x700")
        self.root.minsize(800, 600)
        
        # 设置现代化主题
        self.setup_theme()
        
        # 初始化变量
        self.agent: Optional[LLMWordAgent] = None
        self.input_file_path = tk.StringVar()
        self.output_file_path = tk.StringVar()
        self.api_key = tk.StringVar()
        self.base_url = tk.StringVar(value="https://openrouter.ai/api/v1")
        self.model_name = tk.StringVar(value="moonshotai/kimi-k2")
        self.is_processing = False
        
        # 创建界面
        self.create_widgets()
        
        # 绑定事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def setup_theme(self):
        """设置现代化主题"""
        style = ttk.Style()
        
        # 配置颜色主题
        self.colors = {
            'primary': '#2563eb',      # 蓝色
            'primary_light': '#3b82f6',
            'success': '#10b981',      # 绿色
            'error': '#ef4444',        # 红色
            'warning': '#f59e0b',      # 黄色
            'background': '#f8fafc',   # 浅灰背景
            'surface': '#ffffff',      # 白色表面
            'text': '#1f2937',         # 深灰文字
            'text_light': '#6b7280',   # 浅灰文字
        }
        
        # 配置样式
        style.configure('Title.TLabel', font=('Segoe UI', 16, 'bold'), foreground=self.colors['text'])
        style.configure('Heading.TLabel', font=('Segoe UI', 12, 'bold'), foreground=self.colors['text'])
        style.configure('Info.TLabel', font=('Segoe UI', 10), foreground=self.colors['text_light'])
        style.configure('Primary.TButton', font=('Segoe UI', 10, 'bold'))
        style.configure('Success.TButton', font=('Segoe UI', 10))
        style.configure('Error.TButton', font=('Segoe UI', 10))
        
        # 设置根窗口背景
        self.root.configure(bg=self.colors['background'])
    
    def create_widgets(self):
        """创建界面组件"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # 标题
        title_label = ttk.Label(main_frame, text="Word Agent - AI驱动的Word编辑助手", style='Title.TLabel')
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # API配置区域
        self.create_api_config_section(main_frame, 1)
        
        # 文件选择区域
        self.create_file_section(main_frame, 2)
        
        # 指令输入区域
        self.create_instruction_section(main_frame, 3)
        
        # 控制按钮区域
        self.create_control_section(main_frame, 4)
        
        # 进度显示区域
        self.create_progress_section(main_frame, 5)
        
        # 状态栏
        self.create_status_bar(main_frame, 6)
    
    def create_api_config_section(self, parent, row):
        """创建API配置区域"""
        # 配置框架
        config_frame = ttk.LabelFrame(parent, text="API配置", padding="10")
        config_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        config_frame.columnconfigure(1, weight=1)
        
        # API Key
        ttk.Label(config_frame, text="API Key:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        api_key_entry = ttk.Entry(config_frame, textvariable=self.api_key, show="*", width=50)
        api_key_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        
        # 显示/隐藏API Key按钮
        self.show_api_key = tk.BooleanVar()
        show_btn = ttk.Checkbutton(config_frame, text="显示", variable=self.show_api_key, 
                                 command=lambda: api_key_entry.config(show="" if self.show_api_key.get() else "*"))
        show_btn.grid(row=0, column=2, sticky=tk.W)
        
        # Base URL
        ttk.Label(config_frame, text="Base URL:").grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(10, 0))
        ttk.Entry(config_frame, textvariable=self.base_url, width=50).grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(0, 10), pady=(10, 0))
        
        # Model
        ttk.Label(config_frame, text="模型:").grid(row=2, column=0, sticky=tk.W, padx=(0, 10), pady=(10, 0))
        ttk.Entry(config_frame, textvariable=self.model_name, width=50).grid(row=2, column=1, sticky=(tk.W, tk.E), padx=(0, 10), pady=(10, 0))
        
        # 测试连接按钮
        test_btn = ttk.Button(config_frame, text="测试连接", command=self.test_connection)
        test_btn.grid(row=2, column=2, sticky=tk.W, pady=(10, 0))
    
    def create_file_section(self, parent, row):
        """创建文件选择区域"""
        file_frame = ttk.LabelFrame(parent, text="文件配置", padding="10")
        file_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)
        
        # 输入文件
        ttk.Label(file_frame, text="输入文件:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        ttk.Entry(file_frame, textvariable=self.input_file_path, width=60).grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        ttk.Button(file_frame, text="选择文件", command=self.select_input_file).grid(row=0, column=2, sticky=tk.W)
        
        # 提示信息
        info_label = ttk.Label(file_frame, text="留空将创建新文档", style='Info.TLabel')
        info_label.grid(row=1, column=1, sticky=tk.W, pady=(5, 0))
        
        # 输出文件
        ttk.Label(file_frame, text="输出文件:").grid(row=2, column=0, sticky=tk.W, padx=(0, 10), pady=(10, 0))
        ttk.Entry(file_frame, textvariable=self.output_file_path, width=60).grid(row=2, column=1, sticky=(tk.W, tk.E), padx=(0, 10), pady=(10, 0))
        ttk.Button(file_frame, text="选择位置", command=self.select_output_file).grid(row=2, column=2, sticky=tk.W, pady=(10, 0))
    
    def create_instruction_section(self, parent, row):
        """创建指令输入区域"""
        instruction_frame = ttk.LabelFrame(parent, text="编辑指令", padding="10")
        instruction_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        instruction_frame.columnconfigure(0, weight=1)
        instruction_frame.rowconfigure(1, weight=1)
        parent.rowconfigure(row, weight=1)
        
        # 指令标签
        ttk.Label(instruction_frame, text="请输入您的编辑指令:", style='Heading.TLabel').grid(row=0, column=0, sticky=tk.W, pady=(0, 10))
        
        # 指令文本框
        self.instruction_text = scrolledtext.ScrolledText(instruction_frame, height=8, width=80, wrap=tk.WORD)
        self.instruction_text.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 示例指令
        example_text = """示例指令：
1. 创建一个标题为"会议纪要"的文档
2. 添加三个段落介绍会议背景
3. 插入一个3x3的表格记录会议内容
4. 将所有字体改为宋体
5. 在文档末尾添加签名栏"""
        
        self.instruction_text.insert(tk.END, example_text)
    
    def create_control_section(self, parent, row):
        """创建控制按钮区域"""
        control_frame = ttk.Frame(parent)
        control_frame.grid(row=row, column=0, columnspan=3, pady=(0, 10))
        
        # 开始处理按钮
        self.start_btn = ttk.Button(control_frame, text="开始处理", command=self.start_processing, style='Primary.TButton')
        self.start_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 停止处理按钮
        self.stop_btn = ttk.Button(control_frame, text="停止处理", command=self.stop_processing, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 清空日志按钮
        clear_btn = ttk.Button(control_frame, text="清空日志", command=self.clear_log)
        clear_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # 打开输出文件按钮
        self.open_btn = ttk.Button(control_frame, text="打开输出文件", command=self.open_output_file, state=tk.DISABLED)
        self.open_btn.pack(side=tk.LEFT)
    
    def create_progress_section(self, parent, row):
        """创建进度显示区域"""
        progress_frame = ttk.LabelFrame(parent, text="处理进度", padding="10")
        progress_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        progress_frame.columnconfigure(0, weight=1)
        progress_frame.rowconfigure(1, weight=1)
        
        # 进度条
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, mode='indeterminate')
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # 日志显示
        self.log_text = scrolledtext.ScrolledText(progress_frame, height=15, width=80, wrap=tk.WORD)
        self.log_text.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置日志文本标签
        self.log_text.tag_config("info", foreground=self.colors['text'])
        self.log_text.tag_config("success", foreground=self.colors['success'])
        self.log_text.tag_config("error", foreground=self.colors['error'])
        self.log_text.tag_config("warning", foreground=self.colors['warning'])
    
    def create_status_bar(self, parent, row):
        """创建状态栏"""
        status_frame = ttk.Frame(parent)
        status_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        status_frame.columnconfigure(1, weight=1)
        
        # 状态标签
        self.status_label = ttk.Label(status_frame, text="就绪", style='Info.TLabel')
        self.status_label.grid(row=0, column=0, sticky=tk.W)
        
        # Token使用统计
        self.token_label = ttk.Label(status_frame, text="Token使用: 0", style='Info.TLabel')
        self.token_label.grid(row=0, column=2, sticky=tk.E)
    
    def select_input_file(self):
        """选择输入文件"""
        file_path = filedialog.askopenfilename(
            title="选择Word文档",
            filetypes=[("Word文档", "*.docx"), ("所有文件", "*.*")]
        )
        if file_path:
            self.input_file_path.set(file_path)
    
    def select_output_file(self):
        """选择输出文件"""
        file_path = filedialog.asksaveasfilename(
            title="保存Word文档",
            defaultextension=".docx",
            filetypes=[("Word文档", "*.docx"), ("所有文件", "*.*")]
        )
        if file_path:
            self.output_file_path.set(file_path)
    
    def test_connection(self):
        """测试API连接"""
        if not self.api_key.get():
            messagebox.showerror("错误", "请先输入API Key")
            return
        
        try:
            config = LLMConfig(
                api_key=self.api_key.get(),
                base_url=self.base_url.get(),
                model=self.model_name.get(),
                temperature=0.2,
                max_tokens=100
            )
            
            # 创建临时agent进行测试
            test_agent = LLMWordAgent(config)
            
            # 简单的测试调用
            test_commands, _ = test_agent.call_llm("请输出一个简单的JSON数组: [{\"action\": \"add_paragraph\", \"text\": \"test\"}]")
            
            if test_commands:
                messagebox.showinfo("成功", "API连接测试成功！")
                self.log_message("API连接测试成功", "success")
            else:
                messagebox.showerror("错误", "API连接测试失败")
                
        except Exception as e:
            messagebox.showerror("错误", f"API连接测试失败: {str(e)}")
            self.log_message(f"API连接测试失败: {str(e)}", "error")
    
    def start_processing(self):
        """开始处理"""
        # 验证输入
        if not self.api_key.get():
            messagebox.showerror("错误", "请先输入API Key")
            return
        
        if not self.output_file_path.get():
            messagebox.showerror("错误", "请选择输出文件路径")
            return
        
        instruction = self.instruction_text.get(1.0, tk.END).strip()
        if not instruction:
            messagebox.showerror("错误", "请输入编辑指令")
            return
        
        # 设置UI状态
        self.set_processing_state(True)
        
        # 在后台线程中处理
        thread = threading.Thread(target=self.process_document, args=(instruction,))
        thread.daemon = True
        thread.start()
    
    def process_document(self, instruction: str):
        """处理文档（在后台线程中运行）"""
        try:
            # 创建LLM配置
            config = LLMConfig(
                api_key=self.api_key.get(),
                base_url=self.base_url.get(),
                model=self.model_name.get(),
                temperature=0.2,
                max_tokens=2000
            )
            
            # 创建agent
            self.agent = LLMWordAgent(config)
            self.agent.set_progress_callback(self.log_message)
            
            # 处理文档
            result = self.agent.edit_document(
                docx_path=self.input_file_path.get(),
                user_inst=instruction,
                out_path=self.output_file_path.get()
            )
            
            # 更新UI
            self.root.after(0, self.on_processing_complete, result)
            
        except Exception as e:
            error_msg = f"处理失败: {str(e)}"
            self.root.after(0, self.log_message, error_msg, "error")
            self.root.after(0, self.set_processing_state, False)
    
    def on_processing_complete(self, result):
        """处理完成回调"""
        self.set_processing_state(False)
        
        if result["status"] == "success":
            self.log_message("文档处理完成！", "success")
            self.open_btn.config(state=tk.NORMAL)
            
            # 显示结果摘要
            if "best_attempt" in result:
                best = result["best_attempt"]
                summary = f"成功完成 {best['success_count']} 个操作"
                if best['error_count'] > 0:
                    summary += f"，{best['error_count']} 个操作失败"
                else:
                    summary += "，0 个错误"
                self.log_message(summary, "success")
                
                # 显示是否找到完美执行
                if result.get("perfect_attempt_found", False):
                    self.log_message("✅ 发现完美执行（无错误）！", "success")
            
            messagebox.showinfo("成功", "文档处理完成！")
            
        elif result["status"] == "partial_success":
            self.log_message("文档部分处理完成", "warning")
            self.open_btn.config(state=tk.NORMAL)
            
            if "best_attempt" in result:
                best = result["best_attempt"]
                summary = f"部分成功：完成 {best['success_count']} 个操作，{best['error_count']} 个操作失败"
                self.log_message(summary, "warning")
            
            messagebox.showwarning("部分成功", "文档处理部分完成，请查看日志了解详情")
            
        else:
            self.log_message("文档处理失败", "error")
            error_msg = result.get("error", "未知错误")
            messagebox.showerror("失败", f"文档处理失败: {error_msg}")
        
        # 更新token使用统计
        if "token_usage" in result:
            token_usage = result["token_usage"]
            total_tokens = token_usage.get("total_tokens", 0)
            prompt_tokens = token_usage.get("prompt_tokens", 0)
            completion_tokens = token_usage.get("completion_tokens", 0)
            
            # 显示详细的token统计
            token_summary = f"Token使用: {total_tokens} (输入: {prompt_tokens}, 输出: {completion_tokens})"
            self.token_label.config(text=token_summary)
            
            # 在日志中也显示token统计
            self.log_message(f"📊 {token_summary}", "info")
        
        # 显示总尝试次数
        if "total_attempts" in result:
            attempts_msg = f"总共尝试 {result['total_attempts']} 次"
            self.log_message(f"🔄 {attempts_msg}", "info")
    
    def stop_processing(self):
        """停止处理"""
        self.is_processing = False
        self.log_message("用户取消处理", "warning")
        self.set_processing_state(False)
    
    def set_processing_state(self, processing: bool):
        """设置处理状态"""
        self.is_processing = processing
        
        if processing:
            self.start_btn.config(state=tk.DISABLED)
            self.stop_btn.config(state=tk.NORMAL)
            self.progress_bar.start()
            self.status_label.config(text="处理中...")
        else:
            self.start_btn.config(state=tk.NORMAL)
            self.stop_btn.config(state=tk.DISABLED)
            self.progress_bar.stop()
            self.status_label.config(text="就绪")
    
    def log_message(self, message: str, level: str = "info"):
        """记录日志消息"""
        import datetime
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {message}\n"
        
        # 在主线程中更新UI
        def update_log():
            self.log_text.insert(tk.END, formatted_message, level)
            self.log_text.see(tk.END)
        
        if threading.current_thread() == threading.main_thread():
            update_log()
        else:
            self.root.after(0, update_log)
    
    def clear_log(self):
        """清空日志"""
        self.log_text.delete(1.0, tk.END)
    
    def open_output_file(self):
        """打开输出文件"""
        output_path = self.output_file_path.get()
        if output_path and os.path.exists(output_path):
            try:
                os.startfile(output_path)  # Windows
            except AttributeError:
                try:
                    os.system(f"open '{output_path}'")  # macOS
                except:
                    os.system(f"xdg-open '{output_path}'")  # Linux
        else:
            messagebox.showerror("错误", "输出文件不存在")
    
    def on_closing(self):
        """关闭程序"""
        if self.is_processing:
            if messagebox.askokcancel("退出", "正在处理中，确定要退出吗？"):
                self.stop_processing()
                self.root.destroy()
        else:
            self.root.destroy()


def main():
    """主函数"""
    root = tk.Tk()
    app = WordAgentGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main() 