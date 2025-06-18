import os
import json
import time
import win32gui
import tkinter as tk
from tkinter import filedialog, messagebox

version = "1.1.0"

class NeteaseMusicApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"NCM Catcher v{version}")
        self.root.geometry("500x500")
        self.setup_config()
        self.create_widgets()
        self.auto_running = False
        self.last_song = None

    def setup_config(self):
        """初始化配置文件"""
        if not os.path.exists("config.json"):
            default_config = {
                "save_to_file": True,
                "order_number": ".",
                "rate": 5,
                "save_path": os.path.join(os.getcwd(), "songs.txt")
            }
            with open("config.json", "w", encoding="utf-8") as f:
                json.dump(default_config, f, ensure_ascii=False, indent=4)
        
        with open("config.json", "r", encoding="utf-8") as f:
            self.config = json.load(f)

    def save_config(self):
        """保存配置到文件"""
        with open("config.json", "w", encoding="utf-8") as f:
            json.dump(self.config, f, ensure_ascii=False, indent=4)

    def create_widgets(self):
        """创建GUI界面组件"""
        # 顶部标题
        title_frame = tk.Frame(self.root)
        title_frame.pack(fill=tk.X, padx=10, pady=10)
        tk.Label(title_frame, text="网易云音乐歌单姬", font=("Arial", 14)).pack()
        
        # 当前歌曲显示区域
        self.song_frame = tk.LabelFrame(self.root, text="当前歌曲信息")
        self.song_frame.pack(fill=tk.X, padx=10, pady=5)
        self.song_label = tk.Label(self.song_frame, text="等待获取...", font=("Arial", 12))
        self.song_label.pack(padx=10, pady=10)
        
        # 控制按钮区域
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)
        
        tk.Button(btn_frame, text="获取当前歌曲", command=self.get_info, width=15).pack(side=tk.LEFT, padx=5)
        self.auto_btn = tk.Button(btn_frame, text="开始自动获取", command=self.toggle_auto, width=15)
        self.auto_btn.pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="设置保存路径", command=self.set_path, width=15).pack(side=tk.LEFT, padx=5)
        
        # 配置设置区域
        config_frame = tk.LabelFrame(self.root, text="配置设置")
        config_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # 序号设置
        tk.Label(config_frame, text="序号格式:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.order_entry = tk.Entry(config_frame, width=10)
        self.order_entry.grid(row=0, column=1, padx=5, pady=5)
        self.order_entry.insert(0, self.config["order_number"])
        tk.Button(config_frame, text="更新", command=self.update_order, width=8).grid(row=0, column=2, padx=5)
        
        # 保存选项
        self.save_var = tk.BooleanVar(value=self.config["save_to_file"])
        tk.Checkbutton(config_frame, text="保存到文件",
                        variable=self.save_var, 
                        command=self.toggle_save).grid(row=1, column=0, columnspan=3, padx=5, pady=5, sticky=tk.W
        )
        
        # 保存路径显示
        path_frame = tk.Frame(config_frame)
        path_frame.grid(row=2, column=0, columnspan=3, sticky=tk.W, padx=5, pady=5)
        tk.Label(path_frame, text="保存路径:").pack(side=tk.LEFT)
        self.path_label = tk.Label(path_frame, text=self.config["save_path"], fg="blue")
        self.path_label.pack(side=tk.LEFT, padx=5)
        
        # 日志区域
        log_frame = tk.LabelFrame(self.root, text="操作日志")
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.log_text = tk.Text(log_frame, height=8)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.log_text.config(state=tk.DISABLED)
        
        # 状态栏
        self.status_var = tk.StringVar(value="就绪")
        status_bar = tk.Label(self.root, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def log_message(self, message):
        """添加消息到日志区域"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"{time.strftime('%H:%M:%S')} - {message}\n")
        self.log_text.see(tk.END)  # 滚动到最新消息
        self.log_text.config(state=tk.DISABLED)
        self.status_var.set(message)
    
    def get_netease_music_info(self):
        """获取网易云音乐当前播放歌曲信息"""
        MAIN_WINDOW_CLASS = "OrpheusBrowserHost"
        
        try:
            hwnd = win32gui.FindWindow(MAIN_WINDOW_CLASS, None)
            if not hwnd:
                return {"error": "未找到网易云音乐窗口"}
            
            title = win32gui.GetWindowText(hwnd)
            
            if not title or title == "网易云音乐":
                return {"error": "窗口标题无播放信息"}
            
            if " - " in title:
                parts = title.split(" - ", 1)
                return {"song": parts[0].strip()}
            
            return {"song": title, "artist": "未知歌手"}
        
        except Exception as e:
            return {"error": f"系统错误: {str(e)}"}

    def write_file(self, song_info):
        """将歌曲信息写入文件"""
        if not os.path.exists(self.config["save_path"]):
            with open(self.config["save_path"], "w", encoding="utf-8") as f:
                pass  # 创建空文件
        
        with open(self.config["save_path"], "r", encoding="utf-8") as f:
            lines = f.readlines()
            num = len(lines) + 1
        
        with open(self.config["save_path"], "a", encoding="utf-8") as file:
            file.write(f'{num}{self.config["order_number"]}{song_info["song"]}\n')
        
        self.log_message(f"已保存: {song_info['song']}")

    def get_info(self):
        """获取并显示当前歌曲信息"""
        song_info = self.get_netease_music_info()
        
        if "error" in song_info:
            self.song_label.config(text="错误: " + song_info["error"], fg="red")
            self.log_message("错误: " + song_info["error"])
        else:
            song_name = song_info["song"]
            self.song_label.config(text=song_name, fg="black")
            self.log_message(f"获取成功: {song_name}")
            
            if self.config["save_to_file"]:
                self.write_file(song_info)
    
    def toggle_auto(self):
        """切换自动获取状态"""
        if not self.auto_running:
            self.auto_running = True
            self.auto_btn.config(text="停止自动获取")
            self.log_message("开始自动获取歌曲信息...")
            self.auto_get_info()
        else:
            self.auto_running = False
            self.auto_btn.config(text="开始自动获取")
            self.log_message("已停止自动获取")
    
    def auto_get_info(self):
        """自动获取歌曲信息的定时任务"""
        if not self.auto_running:
            return
            
        song_info = self.get_netease_music_info()
        
        if "error" not in song_info:
            song_name = song_info["song"]
            if song_name != self.last_song:
                self.last_song = song_name
                self.song_label.config(text=song_name, fg="black")
                self.log_message(f"检测到新歌曲: {song_name}")
                
                if self.config["save_to_file"]:
                    self.write_file(song_info)
        else:
            self.song_label.config(text="错误: " + song_info["error"], fg="red")
            self.log_message("错误: " + song_info["error"])
        
        # 设置下次检查
        self.root.after(int(self.config["rate"]) * 1000, self.auto_get_info)
    
    def set_path(self):
        """设置文件保存路径"""
        initial_file = self.config.get("save_path", os.path.join(os.getcwd(), "songs.txt"))
        file_path = filedialog.asksaveasfilename(
            title="选择保存文件",
            defaultextension=".txt",
            initialfile=os.path.basename(initial_file),
            filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")]
        )
        
        if file_path:
            self.config["save_path"] = file_path
            self.path_label.config(text=file_path)
            self.save_config()
            self.log_message(f"保存路径已更新: {file_path}")
    
    def update_order(self):
        """更新序号格式"""
        new_order = self.order_entry.get().strip()
        if new_order:
            self.config["order_number"] = new_order
            self.save_config()
            self.log_message(f"序号格式已更新: {new_order}")
    
    def toggle_save(self):
        """切换保存到文件选项"""
        self.config["save_to_file"] = self.save_var.get()
        self.save_config()
        status = "开启" if self.config["save_to_file"] else "关闭"
        self.log_message(f"保存到文件功能已{status}")

def main():
    try:
        root = tk.Tk()
        NeteaseMusicApp(root)
        root.mainloop()
    except Exception as e:
        messagebox.showerror("启动错误", f"错误: {str(e)}")

if __name__ == "__main__":
    main()