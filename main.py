import os
import re
import json
import time
import win32gui
import tkinter as tk
from tkinter import filedialog, messagebox

version = "1.2.0"

class NeteaseMusicApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"Song Catcher v{version}")
        self.root.geometry("500x600")
        self.setup_config()
        self.create_widgets()
        self.auto_ncm_running = False
        self.auto_wx_running = False
        self.last_ncm_song = None
        self.last_wx_song = None

    def setup_config(self):
        """初始化配置文件并确保兼容性"""
        config_path = "config.json"
        
        # 如果配置文件不存在，创建默认配置
        if not os.path.exists(config_path):
            default_config = {
                "save_to_file": True,
                "order_number": ".",
                "ncm_rate": 5,  # 网易云轮询间隔
                "wx_rate": 5,    # 全民K歌轮询间隔
                "save_path": os.path.join(os.getcwd(), "songs.txt"),
            }
            with open(config_path, "w", encoding="utf-8") as f:
                json.dump(default_config, f, ensure_ascii=False, indent=4)
        
        # 加载配置文件
        with open(config_path, "r", encoding="utf-8") as f:
            self.config = json.load(f)
            
        # 确保配置项存在
        if "ncm_rate" not in self.config:
            self.config["ncm_rate"] = 5
        if "wx_rate" not in self.config:
            self.config["wx_rate"] = 5
            
        # 保存更新后的配置文件
        self.save_config()

    def save_config(self):
        """保存配置到文件"""
        with open("config.json", "w", encoding="utf-8") as f:
            json.dump(self.config, f, ensure_ascii=False, indent=4)

    def create_widgets(self):
        """创建GUI界面组件"""
        # 顶部标题
        title_frame = tk.Frame(self.root)
        title_frame.pack(fill=tk.X, padx=10, pady=10)
        tk.Label(title_frame, text="歌单姬", font=("Arial", 14)).pack()
        
        # 当前歌曲显示区域
        self.song_frame = tk.LabelFrame(self.root, text="当前歌曲信息")
        self.song_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # 网易云歌曲显示
        ncm_frame = tk.Frame(self.song_frame)
        ncm_frame.pack(fill=tk.X, padx=5, pady=5)
        tk.Label(ncm_frame, text="网易云:", font=("Arial", 10)).pack(side=tk.LEFT)
        self.ncm_label = tk.Label(ncm_frame, text="未获取", font=("Arial", 10), wraplength=400, justify=tk.LEFT)
        self.ncm_label.pack(side=tk.LEFT, padx=5)
        
        # 全民K歌歌曲显示
        wx_frame = tk.Frame(self.song_frame)
        wx_frame.pack(fill=tk.X, padx=5, pady=5)
        tk.Label(wx_frame, text="全民K歌:", font=("Arial", 10)).pack(side=tk.LEFT)
        self.wx_label = tk.Label(wx_frame, text="未获取", font=("Arial", 10), wraplength=400, justify=tk.LEFT)
        self.wx_label.pack(side=tk.LEFT, padx=5)
        
        # 手动获取按钮区域
        manual_frame = tk.LabelFrame(self.root, text="手动获取")
        manual_frame.pack(fill=tk.X, padx=10, pady=5)
        
        btn_frame = tk.Frame(manual_frame)
        btn_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # 获取网易云歌曲按钮
        tk.Button(btn_frame, text="获取网易云歌曲", command=self.get_ncm_info, width=15).pack(side=tk.LEFT, padx=5)
        # 获取全民K歌歌曲按钮
        tk.Button(btn_frame, text="获取全民K歌歌曲", command=self.get_wx_info, width=15).pack(side=tk.LEFT, padx=5)
        # 设置保存路径按钮
        tk.Button(btn_frame, text="设置保存路径", command=self.set_path, width=15).pack(side=tk.LEFT, pady=5)
        
        # 自动获取区域
        auto_frame = tk.LabelFrame(self.root, text="自动获取")
        auto_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # 网易云自动获取
        ncm_auto_frame = tk.Frame(auto_frame)
        ncm_auto_frame.pack(fill=tk.X, padx=5, pady=5)
        tk.Label(ncm_auto_frame, text="网易云自动获取:").pack(side=tk.LEFT)
        self.ncm_auto_btn = tk.Button(ncm_auto_frame, text="开始", command=self.toggle_auto_ncm, width=8)
        self.ncm_auto_btn.pack(side=tk.LEFT, padx=5)
        tk.Label(ncm_auto_frame, text="间隔(秒):").pack(side=tk.LEFT)
        self.ncm_rate_var = tk.IntVar(value=self.config["ncm_rate"])
        self.ncm_rate_spin = tk.Spinbox(ncm_auto_frame, from_=1, to=60, width=5, textvariable=self.ncm_rate_var)
        self.ncm_rate_spin.pack(side=tk.LEFT, padx=5)
        tk.Button(ncm_auto_frame, text="更新", command=self.update_ncm_rate, width=8).pack(side=tk.LEFT, padx=5)
        
        # 全民K歌自动获取
        wx_auto_frame = tk.Frame(auto_frame)
        wx_auto_frame.pack(fill=tk.X, padx=5, pady=5)
        tk.Label(wx_auto_frame, text="全民K歌自动获取:").pack(side=tk.LEFT)
        self.wx_auto_btn = tk.Button(wx_auto_frame, text="开始", command=self.toggle_auto_wx, width=8)
        self.wx_auto_btn.pack(side=tk.LEFT, padx=5)
        tk.Label(wx_auto_frame, text="间隔(秒):").pack(side=tk.LEFT)
        self.wx_rate_var = tk.IntVar(value=self.config["wx_rate"])
        self.wx_rate_spin = tk.Spinbox(wx_auto_frame, from_=1, to=60, width=5, textvariable=self.wx_rate_var)
        self.wx_rate_spin.pack(side=tk.LEFT, padx=5)
        tk.Button(wx_auto_frame, text="更新", command=self.update_wx_rate, width=8).pack(side=tk.LEFT, padx=5)
        
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
                        command=self.toggle_save).grid(row=1, column=0, columnspan=3, padx=5, pady=5, sticky=tk.W)
        
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
    
    def get_ncm_info(self):
        """获取网易云音乐当前播放歌曲信息"""
        try:
            NCM_WINDOW_CLASS = "OrpheusBrowserHost"
            hwnd = win32gui.FindWindow(NCM_WINDOW_CLASS, None)
            if not hwnd:
                self.ncm_label.config(text="未找到网易云音乐窗口", fg="red")
                self.log_message("错误: 未找到网易云音乐窗口")
                return
            
            title = win32gui.GetWindowText(hwnd)
            if not title or title == "网易云音乐":
                self.ncm_label.config(text="未在播放歌曲", fg="red")
                self.log_message("错误: 网易云未在播放歌曲")
                return
            
            if " - " in title:
                parts = title.split(" - ", 1)
                song_name = parts[0].strip()
            else:
                song_name = title
                
            self.ncm_label.config(text=song_name, fg="black")
            self.log_message(f"获取网易云歌曲成功: {song_name}")
            
            if self.config["save_to_file"]:
                self.save_song({"song": song_name, "source": "网易云"})
            
            self.last_ncm_song = song_name

        except Exception as e:
            self.ncm_label.config(text=f"系统错误: {str(e)}", fg="red")
            self.log_message(f"网易云错误: {str(e)}")
    
    def get_wx_info(self):
        """获取全民K歌当前播放歌曲信息"""
        try:
            wx_song = None
            
            # 定义回调函数
            def enum_windows_callback(hwnd, _):
                nonlocal wx_song
                if win32gui.IsWindowVisible(hwnd):
                    text = win32gui.GetWindowText(hwnd)
                    # 匹配全民K歌窗口标题
                    if re.search(r"(全民K歌|QQK歌|WeSing)", text):
                        if " - " in text:
                            wx_song = text.split(" - ")[1]
                        else:
                            wx_song = text
                        return False  # 找到后停止枚举
                return True  # 继续枚举
            
            # 枚举窗口
            win32gui.EnumWindows(enum_windows_callback, None)
            
            if not wx_song:
                self.wx_label.config(text="未找到全民K歌窗口", fg="red")
                self.log_message("错误: 未找到全民K歌窗口")
                return
            
            self.wx_label.config(text=wx_song, fg="black")
            self.log_message(f"获取全民K歌歌曲成功: {wx_song}")
            
            if self.config["save_to_file"]:
                self.save_song({"song": wx_song, "source": "全民K歌"})
            
            self.last_wx_song = wx_song

        except Exception as e:
            self.wx_label.config(text=f"系统错误: {str(e)}", fg="red")
            self.log_message(f"全民K歌错误: {str(e)}")
    
    def save_song(self, song_info):
        """将歌曲信息写入文件"""
        # 确保目录存在
        save_dir = os.path.dirname(self.config["save_path"])
        if save_dir and not os.path.exists(save_dir):
            os.makedirs(save_dir, exist_ok=True)
        
        # 处理新文件或读取现有内容
        if not os.path.exists(self.config["save_path"]):
            with open(self.config["save_path"], "w", encoding="utf-8") as f:
                pass  # 创建空文件
        
        with open(self.config["save_path"], "r", encoding="utf-8") as f:
            lines = f.readlines()
            num = len(lines) + 1
        
        # 写入新歌曲
        with open(self.config["save_path"], "a", encoding="utf-8") as file:
            file.write(f'{num}{self.config["order_number"]}{song_info["song"]}\n')
        
        # 记录日志
        source = song_info.get('source', '未知来源')
        song_name = song_info["song"]
        self.log_message(f"已保存 ({source}): {song_name}")
    
    def toggle_auto_ncm(self):
        """切换网易云自动获取状态"""
        if not self.auto_ncm_running:
            self.auto_ncm_running = True
            self.ncm_auto_btn.config(text="停止")
            self.log_message("开始自动获取网易云歌曲信息...")
            self.auto_get_ncm_info()
        else:
            self.auto_ncm_running = False
            self.ncm_auto_btn.config(text="开始")
            self.log_message("已停止网易云自动获取")
    
    def toggle_auto_wx(self):
        """切换全民K歌自动获取状态"""
        if not self.auto_wx_running:
            self.auto_wx_running = True
            self.wx_auto_btn.config(text="停止")
            self.log_message("开始自动获取全民K歌歌曲信息...")
            self.auto_get_wx_info()
        else:
            self.auto_wx_running = False
            self.wx_auto_btn.config(text="开始")
            self.log_message("已停止全民K歌自动获取")
    
    def auto_get_ncm_info(self):
        """自动获取网易云歌曲信息"""
        if not self.auto_ncm_running:
            return
            
        try:
            NCM_WINDOW_CLASS = "OrpheusBrowserHost"
            hwnd = win32gui.FindWindow(NCM_WINDOW_CLASS, None)
            if not hwnd:
                self.ncm_label.config(text="未找到网易云音乐窗口", fg="red")
                self.log_message("错误: 未找到网易云音乐窗口")
                return
            
            title = win32gui.GetWindowText(hwnd)
            if not title or title == "网易云音乐":
                self.ncm_label.config(text="未在播放歌曲", fg="red")
                self.log_message("错误: 网易云未在播放歌曲")
                return
            
            if " - " in title:
                parts = title.split(" - ", 1)
                song_name = parts[0].strip()
            else:
                song_name = title
                
            # 检查歌曲是否变化
            if song_name != self.last_ncm_song:
                self.ncm_label.config(text=song_name, fg="black")
                self.log_message(f"检测到新网易云歌曲: {song_name}")
                
                if self.config["save_to_file"]:
                    self.save_song({"song": song_name, "source": "网易云"})
                
                self.last_ncm_song = song_name

        except Exception as e:
            self.ncm_label.config(text=f"系统错误: {str(e)}", fg="red")
            self.log_message(f"网易云自动获取错误: {str(e)}")
        
        # 设置下次检查
        if self.auto_ncm_running:
            try:
                check_rate = max(1, int(self.config["ncm_rate"])) * 1000
            except (ValueError, TypeError):
                check_rate = 5000  # 5秒的默认值
                
            self.root.after(check_rate, self.auto_get_ncm_info)
    
    def auto_get_wx_info(self):
        """自动获取全民K歌歌曲信息"""
        if not self.auto_wx_running:
            return
            
        try:
            wx_song = None
            
            # 定义回调函数
            def enum_windows_callback(hwnd, _):
                nonlocal wx_song
                if win32gui.IsWindowVisible(hwnd):
                    text = win32gui.GetWindowText(hwnd)
                    # 匹配全民K歌窗口标题
                    if re.search(r"(全民K歌|QQK歌|WeSing)", text):
                        if " - " in text:
                            wx_song = text.split(" - ")[1]
                        else:
                            wx_song = text
                        return False  # 找到后停止枚举
                return True  # 继续枚举
            
            # 枚举窗口
            win32gui.EnumWindows(enum_windows_callback, None)
            
            if not wx_song:
                self.wx_label.config(text="未找到全民K歌窗口", fg="red")
                self.log_message("错误: 未找到全民K歌窗口")
                return
            
            # 检查歌曲是否变化
            if wx_song != self.last_wx_song:
                self.wx_label.config(text=wx_song, fg="black")
                self.log_message(f"检测到新全民K歌歌曲: {wx_song}")
                
                if self.config["save_to_file"]:
                    self.save_song({"song": wx_song, "source": "全民K歌"})
                
                self.last_wx_song = wx_song

        except Exception as e:
            self.wx_label.config(text=f"系统错误: {str(e)}", fg="red")
            self.log_message(f"全民K歌自动获取错误: {str(e)}")
        
        # 设置下次检查
        if self.auto_wx_running:
            try:
                check_rate = max(1, int(self.config["wx_rate"])) * 1000
            except (ValueError, TypeError):
                check_rate = 5000  # 5秒的默认值
                
            self.root.after(check_rate, self.auto_get_wx_info)
    
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
        else:
            messagebox.showwarning("格式无效", "序号格式不能为空")
    
    def toggle_save(self):
        """切换保存到文件选项"""
        self.config["save_to_file"] = self.save_var.get()
        self.save_config()
        status = "开启" if self.config["save_to_file"] else "关闭"
        self.log_message(f"保存到文件功能已{status}")
    
    def update_ncm_rate(self):
        """更新网易云轮询间隔时间"""
        try:
            new_rate = int(self.ncm_rate_var.get())
            if new_rate < 1 or new_rate > 60:
                raise ValueError("无效值")
            
            self.config["ncm_rate"] = new_rate
            self.save_config()
            self.log_message(f"网易云轮询间隔已更新: {new_rate}秒")
        except ValueError:
            messagebox.showerror("无效输入", "请输入1-60之间的有效整数")
            self.ncm_rate_var.set(self.config.get("ncm_rate", 5))
    
    def update_wx_rate(self):
        """更新全民K歌轮询间隔时间"""
        try:
            new_rate = int(self.wx_rate_var.get())
            if new_rate < 1 or new_rate > 60:
                raise ValueError("无效值")
            
            self.config["wx_rate"] = new_rate
            self.save_config()
            self.log_message(f"全民K歌轮询间隔已更新: {new_rate}秒")
        except ValueError:
            messagebox.showerror("无效输入", "请输入1-60之间的有效整数")
            self.wx_rate_var.set(self.config.get("wx_rate", 5))

def main():
    try:
        root = tk.Tk()
        app = NeteaseMusicApp(root)
        root.mainloop()
    except Exception as e:
        error_msg = f"程序启动错误: {str(e)}\n请检查配置文件config.json是否有效。"
        messagebox.showerror("启动错误", error_msg)

if __name__ == "__main__":
    main()