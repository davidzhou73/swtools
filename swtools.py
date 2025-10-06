import pathlib
import pandas as pd
import paramiko
import time
import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading

class SWToolsGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("批量生成/执行设备脚本工具")
        # 设置窗口大小和居中
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        win_width = min(1024, screen_width)
        win_height = min(768, screen_height)
        x = (screen_width - win_width) // 2
        y = (screen_height - win_height) // 2
        self.root.geometry(f"{win_width}x{win_height}+{x}+{y}")
        # 文件路径变量
        self.device_file_path = tk.StringVar()
        self.save_path = tk.StringVar()
        self.create_widgets()

    def create_widgets(self):
        """创建主界面控件，包括文件选择、按钮、信息显示区域"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky="nsew")
        # 配置网格权重，保证窗口自适应
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        # 文件选择区域
        file_frame = ttk.LabelFrame(main_frame, text="文件选择", padding="5")
        file_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        file_frame.columnconfigure(1, weight=1)
        # 设备配置和脚本文件选择
        ttk.Label(file_frame, text="设备脚本配置文件:").grid(row=0, column=0, sticky="w", padx=(0, 5))
        ttk.Entry(file_frame, textvariable=self.device_file_path, width=50).grid(row=0, column=1, sticky="ew", padx=(0, 5))
        ttk.Button(file_frame, text="浏览", command=self.browse_device_file).grid(row=0, column=2)
        # 保存路径
        ttk.Label(file_frame, text="脚本/执行结果保存路径:").grid(row=1, column=0, sticky="w", padx=(0, 5))
        ttk.Entry(file_frame, textvariable=self.save_path, width=50).grid(row=1, column=1, sticky="ew", padx=(0, 5))
        # 控制按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=1, column=0, columnspan=2, pady=(0, 10))
        self.save_button = ttk.Button(button_frame, text="生成脚本并保存", command=self.save_commands, state=tk.DISABLED)
        self.save_button.pack(side=tk.LEFT, padx=(0, 5))
        self.execute_button = ttk.Button(button_frame, text="生成脚本并执行", command=self.execute_commands, state=tk.DISABLED)
        self.execute_button.pack(side=tk.LEFT, padx=(0, 5))
        self.clear_button = ttk.Button(button_frame, text="清空执行结果", command=self.clear_log)
        self.clear_button.pack(side=tk.LEFT)
        # 使用 PanedWindow 分隔设备信息和执行结果区域，可拖动分隔条
        paned = ttk.PanedWindow(main_frame, orient=tk.VERTICAL)
        paned.grid(row=2, column=0, columnspan=2, sticky="nsew", pady=(0, 10))
        main_frame.rowconfigure(2, weight=1)
        # 设备信息显示区域
        device_frame = ttk.LabelFrame(paned, text="设备信息", padding="5")
        device_frame.columnconfigure(0, weight=1)
        self.device_text = scrolledtext.ScrolledText(device_frame, height=8, width=80)
        self.device_text.grid(row=0, column=0, sticky="nsew")
        device_frame.rowconfigure(0, weight=1)
        paned.add(device_frame, weight=1)
        # 执行结果显示区域
        result_frame = ttk.LabelFrame(paned, text="执行结果", padding="5")
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(0, weight=1)
        self.result_text = scrolledtext.ScrolledText(result_frame, height=15, width=80)
        self.result_text.grid(row=0, column=0, sticky="nsew")
        paned.add(result_frame, weight=2)

    def browse_device_file(self):
        """弹出文件选择对话框，选择设备脚本配置文件"""
        filename = filedialog.askopenfilename(
            title="选择设备脚本配置文件",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if filename:
            self.device_file_path.set(filename)
            self.load_device_info()

    def load_device_info(self):
        """读取设备信息并显示到界面"""
        device_file = self.device_file_path.get()
        try:
            self.log_message("正在读取设备信息...")
            df = pd.read_excel(device_file)
            if "设备名称" not in df.columns:
                messagebox.showerror("错误", "Excel文件中未找到'设备名称'列")
                self.log_message("错误: Excel文件中未找到'设备名称'列")
                return
            device_names = df.loc[:, "设备名称"].unique()
            self.device_text.delete(1.0, tk.END)
            if device_names.size > 0:
                self.device_text.insert(tk.END, f"找到{device_names.size}个设备:\n\n")
                i = 1
                for name in device_names:
                    name = name.strip()
                    self.device_text.insert(tk.END, f"{i:>3d}. 设备名称: {name}\n")
                    i += 1
                self.save_button.config(state=tk.NORMAL)
                if "IP地址" in df.columns and "账号" in df.columns and "密码" in df.columns:
                    self.execute_button.config(state=tk.NORMAL)
                else:
                    self.log_message("提示: Excel文件中未找到设备登录相关信息列（'IP地址'、'账号'、'密码'），不能直接执行脚本")
                self.log_message("设备信息读取完成")
            else:
                self.save_button.config(state=tk.DISABLED)
                self.execute_button.config(state=tk.DISABLED)
                messagebox.showinfo("提示", "未找到任何设备信息")
                self.log_message("未找到任何设备信息")
        except Exception as e:
            messagebox.showerror("错误", f"读取设备信息失败: {str(e)}")
            self.log_message(f"错误: {str(e)}")

    def save_commands(self):
        """生成设备命令脚本并保存到指定路径"""
        device_file = self.device_file_path.get()
        if not device_file:
            messagebox.showerror("错误", "请先选择设备脚本配置文件")
            return
        save_folder = filedialog.askdirectory(title="选择保存脚本生成结果的路径")
        if not save_folder:
            messagebox.showinfo("提示", "未选择保存路径，操作取消")
            return
        self.save_path.set(save_folder)
        try:
            self.log_message("开始保存脚本生成结果...")
            df = pd.read_excel(device_file)
            # 判断是否有登录信息，决定脚本命令起始列
            if "IP地址" in df.columns and "账号" in df.columns and "密码" in df.columns:
                cmd_start = 4
            else:
                cmd_start = 1
            device_names = df.loc[:, "设备名称"].unique()
            for device in device_names:
                file_path = pathlib.Path(save_folder).joinpath(f"{device}_cmd.txt")
                with open(f"{file_path}", "w", encoding="utf-8") as f:
                    device_df = df[df.loc[:, "设备名称"] == device]
                    for _, row in device_df.iterrows():
                        row.columns = [str(col).strip() for col in row]
                        f.writelines("#\n")
                        for cmd_idx in range(cmd_start, len(row.columns)):
                            f.writelines(f"{row.iloc[cmd_idx]}\n")
                    f.writelines("#\n")
                self.log_message(f"设备 {device} 脚本写入完成，结果保存到 {file_path}")
        except Exception as e:
            self.log_message(f"错误: {str(e)}")

    def execute_commands(self):
        """批量远程执行设备命令，保存执行结果"""
        device_file = self.device_file_path.get()
        if not device_file:
            messagebox.showerror("错误", "请先选择设备脚本配置文件")
            return
        save_folder = filedialog.askdirectory(title="选择保存脚本执行结果的路径")
        if not save_folder:
            messagebox.showinfo("提示", "未选择保存路径，操作取消")
            return
        self.save_path.set(save_folder)
        # 在新线程中执行命令，防止界面卡死
        self.execute_button.config(state=tk.DISABLED)
        thread = threading.Thread(target=self._execute_commands_thread)
        thread.daemon = True
        thread.start()

    def _execute_commands_thread(self):
        """线程函数：读取设备信息并执行命令"""
        try:
            device_file = self.device_file_path.get()
            df = pd.read_excel(device_file)
            if "IP地址" not in df.columns or "账号" not in df.columns or "密码" not in df.columns:
                self.log_message("错误: Excel文件中未找到设备登录相关信息列（'IP地址'、'账号'、'密码'），不能直接执行脚本")
                return
            device_names = df.loc[:, "设备名称"].unique()
            for device in device_names:
                device_df = df[df.loc[:, "设备名称"] == device]
                for _, row in device_df.iterrows():
                    row.columns = [str(col).strip() for col in row]
                    cmds = []
                    for cmd_idx in range(4, len(row.columns)):
                        cmds.append(row.iloc[cmd_idx])
                self.ssh_device_with_log(device, row["IP地址"], row["账号"], row["密码"], cmds)
        except Exception as e:
            raise Exception(f"读取设备脚本配置文件失败: {str(e)}")
        finally:
            self.root.after(0, self._execution_finished)

    def _execution_finished(self):
        """命令执行完成后恢复按钮状态"""
        self.execute_button.config(state=tk.NORMAL)

    def clear_log(self):
        """清空执行结果显示区域"""
        self.result_text.delete(1.0, tk.END)

    def log_message(self, message):
        """在执行结果区域追加日志信息，带时间戳"""
        timestamp = datetime.datetime.now().strftime('%H:%M:%S')
        log_entry = f"[{timestamp}] {message}\n"
        def update_log():
            self.result_text.insert(tk.END, log_entry)
            self.result_text.see(tk.END)
        if threading.current_thread() is threading.main_thread():
            update_log()
        else:
            self.root.after(0, update_log)

    def ssh_device_with_log(self, device, ip, user, passwd, cmds):
        """通过SSH连接设备并执行命令，保存返回结果到文件"""
        try:
            ssh = paramiko.SSHClient()
            ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            self.log_message(f"正在连接 {ip} ...")
            ssh.connect(ip, 22, user, passwd, timeout=10)
            channel = ssh.invoke_shell()
            file_path = pathlib.Path(self.save_path.get()).joinpath(f"{device}_result.log")
            with open(file_path, 'w', encoding='utf-8') as f:
                for cmd in cmds:
                    cmd = cmd.strip()
                    if cmd:
                        channel.send(cmd + '\n')
                        time.sleep(5)
                        output = channel.recv(65535)
                        buff = output.decode('utf-8', errors='ignore')
                        # 写入文件
                        f.write("="*15 + f"命令: {cmd}" + "="*15 + "\n")
                        f.write(buff)
                        f.write("\n" + "="*50 + "\n")
            ssh.close()
            self.log_message(f"设备 {device} 脚本执行完成，结果保存到 {file_path}")
        except Exception as e:
            self.log_message(f"错误: {str(e)}")

def main():
    """程序入口"""
    root = tk.Tk()
    app = SWToolsGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()