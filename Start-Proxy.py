# proxy_launcher_optimized.py
import os
import subprocess
import tkinter as tk
from tkinterdnd2 import TkinterDnD, DND_FILES
import threading
import queue


class ProxyLauncher(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("代理启动器 v3")
        self.geometry("500x300")
        self.log_queue = queue.Queue()

        # 代理设置
        self.proxy_frame = tk.LabelFrame(self, text="代理配置")
        self.proxy_frame.pack(pady=10, fill="x", padx=10)

        # HTTP/HTTPS 代理
        tk.Label(self.proxy_frame, text="HTTP/HTTPS代理:").grid(row=0, column=0, sticky="w")
        self.proxy_entry = tk.Entry(self.proxy_frame)
        self.proxy_entry.insert(0, "http://127.0.0.1:7890")
        self.proxy_entry.grid(row=0, column=1, padx=5, sticky="ew")

        # 不代理的地址
        tk.Label(self.proxy_frame, text="不代理的地址:").grid(row=1, column=0, sticky="w")
        self.noproxy_entry = tk.Entry(self.proxy_frame)
        self.noproxy_entry.insert(0, "localhost,127.0.0.1,192.168.1.1")
        self.noproxy_entry.grid(row=1, column=1, padx=5, sticky="ew")

        # 置顶/取消置顶按钮
        self.topmost_button = tk.Button(self.proxy_frame, text="窗口置顶", command=self.toggle_topmost)
        self.topmost_button.grid(row=0, column=2, rowspan=2, padx=10, pady=5, sticky="ns")  # 放在右侧

        # 拖放区域
        self.drop_label = tk.Label(self, text="拖拽应用程序或快捷方式到这里",
                                   relief="groove", padx=20, pady=20)
        self.drop_label.pack(pady=20, fill="x", padx=50)

        # 日志输出
        self.log_text = tk.Text(self, height=10, state="disabled")
        self.log_text.pack(fill="both", expand=True, padx=10, pady=5)

        # 绑定事件
        self.drop_label.drop_target_register(DND_FILES)
        self.drop_label.dnd_bind('<<Drop>>', self.handle_drop)
        self.after(100, self.process_log_queue)

    def toggle_topmost(self):
        """切换窗口置顶状态"""
        current_state = self.attributes("-topmost")
        self.attributes("-topmost", not current_state)
        if not current_state:
            self.topmost_button.config(text="取消置顶")
        else:
            self.topmost_button.config(text="窗口置顶")

    def resolve_shortcut(self, lnk_path):
        """解析Windows快捷方式"""
        try:
            ps_script = f'''
            $sh = New-Object -ComObject WScript.Shell
            $sc = $sh.CreateShortcut('{lnk_path.replace("'", "''")}')
            Write-Output $sc.TargetPath
            '''
            result = subprocess.run(
                ['powershell', '-Command', ps_script],
                capture_output=True,
                text=True,
                check=True,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            return result.stdout.strip()
        except Exception as e:
            self.log(f"快捷方式解析失败: {str(e)}")
            return None

    def handle_drop(self, event):
        """异步处理拖放事件"""

        def async_launch():
            filepath = event.data.strip('{}')
            original_path = filepath

            # 处理快捷方式
            if filepath.lower().endswith('.lnk'):
                target = self.resolve_shortcut(filepath)
                if target and os.path.exists(target):
                    filepath = target
                else:
                    self.log_queue.put("错误：无效的快捷方式")
                    return

            if not os.path.isfile(filepath):
                self.log_queue.put(f"错误：路径无效 - {filepath}")
                return

            # 准备环境变量
            env = os.environ.copy()
            env.update({
                "HTTP_PROXY": self.proxy_entry.get(),
                "HTTPS_PROXY": self.proxy_entry.get(),
                "NO_PROXY": self.noproxy_entry.get()
            })

            # 异步启动程序
            try:
                subprocess.Popen(
                    [filepath],
                    env=env,
                    creationflags=subprocess.DETACHED_PROCESS,
                    shell=True  # 用系统的shell来启动一个新进程
                )
                self.log_queue.put(f"成功启动: {os.path.basename(filepath)}")
                if original_path != filepath:
                    self.log_queue.put(f"快捷方式来源: {original_path}")
            except Exception as e:
                self.log_queue.put(f"启动失败: {str(e)}")

        # 启动线程
        threading.Thread(target=async_launch, daemon=True).start()

    def process_log_queue(self):
        """处理日志队列（主线程中执行）"""
        while not self.log_queue.empty():
            msg = self.log_queue.get_nowait()
            self.log_text.config(state="normal")
            self.log_text.insert("end", msg + "\n")
            self.log_text.see("end")
            self.log_text.config(state="disabled")
        self.after(100, self.process_log_queue)

    def log(self, message):
        """兼容旧日志方法"""
        self.log_queue.put(message)


if __name__ == '__main__':
    app = ProxyLauncher()
    app.mainloop()
