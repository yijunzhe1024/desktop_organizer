import os
import shutil
import time
import threading
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import sys
from pystray import Icon, MenuItem as item, Menu
from PIL import Image, ImageDraw


class DesktopOrganizer:
    def __init__(self, root):
        self.root = root
        self.root.title("桌面整理系统")
        self.root.geometry("600x500")
        self.root.resizable(True, True)

        # 设置中文字体支持
        self.style = ttk.Style()
        self.style.configure("TLabel", font=("SimHei", 10))
        self.style.configure("TButton", font=("SimHei", 10))
        self.style.configure("TCombobox", font=("SimHei", 10))

        # 监控目录（默认为桌面）
        self.monitor_dir = os.path.join(os.path.expanduser("~"), "Desktop")

        # 文件类型与目标文件夹的映射
        self.file_mappings = {
            '图片': ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.svg'],
            '文档': ['.txt', '.doc', '.docx', '.pdf', '.xls', '.xlsx', '.ppt', '.pptx', '.md'],
            '视频': ['.mp4', '.avi', '.mov', '.mkv', '.flv'],
            '音频': ['.mp3', '.wav', '.flac', '.m4a'],
            '压缩文件': ['.zip', '.rar', '.7z', '.tar', '.gz'],
            '程序': ['.exe', '.msi', '.py', '.java', '.cpp', '.html', '.css', '.js']
        }

        # 自动整理开关
        self.auto_organize = False
        self.check_interval = 30  # 检查间隔（秒）
        self.monitoring_thread = None
        self.running = False

        self.create_widgets()
        self.load_settings()

    def create_widgets(self):
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 监控目录设置
        dir_frame = ttk.LabelFrame(main_frame, text="监控目录", padding="10")
        dir_frame.pack(fill=tk.X, pady=5)

        ttk.Label(dir_frame, text="当前监控目录:").pack(side=tk.LEFT, padx=5)
        self.dir_entry = tk.StringVar(value=self.monitor_dir)
        ttk.Entry(dir_frame, textvariable=self.dir_entry, width=50).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(dir_frame, text="更改", command=self.change_dir).pack(side=tk.LEFT, padx=5)

        # 文件类型映射设置
        mapping_frame = ttk.LabelFrame(main_frame, text="文件类型映射", padding="10")
        mapping_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        # 左侧列表：文件类型
        left_frame = ttk.Frame(mapping_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        ttk.Label(left_frame, text="文件类型类别").pack(pady=5)
        self.type_listbox = tk.Listbox(left_frame, selectmode=tk.SINGLE, font=("SimHei", 10))
        self.type_listbox.pack(fill=tk.BOTH, expand=True, pady=5)
        self.update_type_listbox()

        type_buttons = ttk.Frame(left_frame)
        type_buttons.pack(fill=tk.X, pady=5)
        ttk.Button(type_buttons, text="添加类型", command=self.add_type).pack(side=tk.LEFT, padx=2, fill=tk.X,
                                                                              expand=True)
        ttk.Button(type_buttons, text="删除类型", command=self.delete_type).pack(side=tk.LEFT, padx=2, fill=tk.X,
                                                                                 expand=True)

        # 右侧列表：扩展名
        right_frame = ttk.Frame(mapping_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        ttk.Label(right_frame, text="关联的扩展名").pack(pady=5)
        self.ext_listbox = tk.Listbox(right_frame, selectmode=tk.SINGLE, font=("SimHei", 10))
        self.ext_listbox.pack(fill=tk.BOTH, expand=True, pady=5)

        ext_buttons = ttk.Frame(right_frame)
        ext_buttons.pack(fill=tk.X, pady=5)
        ttk.Button(ext_buttons, text="添加扩展名", command=self.add_extension).pack(side=tk.LEFT, padx=2, fill=tk.X,
                                                                                    expand=True)
        ttk.Button(ext_buttons, text="删除扩展名", command=self.delete_extension).pack(side=tk.LEFT, padx=2, fill=tk.X,
                                                                                       expand=True)

        # 绑定列表选择事件
        self.type_listbox.bind('<<ListboxSelect>>', self.on_type_select)

        # 自动整理设置
        auto_frame = ttk.LabelFrame(main_frame, text="自动整理设置", padding="10")
        auto_frame.pack(fill=tk.X, pady=5)

        self.auto_var = tk.BooleanVar(value=self.auto_organize)
        auto_check = ttk.Checkbutton(auto_frame, text="启用自动整理", variable=self.auto_var,
                                     command=self.toggle_auto_organize)
        auto_check.pack(side=tk.LEFT, padx=5)

        ttk.Label(auto_frame, text="检查间隔(秒):").pack(side=tk.LEFT, padx=5)
        self.interval_var = tk.StringVar(value=str(self.check_interval))
        interval_entry = ttk.Entry(auto_frame, textvariable=self.interval_var, width=10)
        interval_entry.pack(side=tk.LEFT, padx=5)
        ttk.Button(auto_frame, text="应用间隔", command=self.apply_interval).pack(side=tk.LEFT, padx=5)

        # 操作按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)

        ttk.Button(button_frame, text="立即整理", command=self.organize_now).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="保存设置", command=self.save_settings).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="关于", command=self.show_about).pack(side=tk.RIGHT, padx=5)

    def update_type_listbox(self):
        self.type_listbox.delete(0, tk.END)
        for type_name in self.file_mappings.keys():
            self.type_listbox.insert(tk.END, type_name)

    def on_type_select(self, event):
        selection = self.type_listbox.curselection()
        if not selection:
            return

        selected_type = self.type_listbox.get(selection[0])
        self.ext_listbox.delete(0, tk.END)
        for ext in self.file_mappings[selected_type]:
            self.ext_listbox.insert(tk.END, ext)

    def change_dir(self):
        from tkinter import filedialog
        new_dir = filedialog.askdirectory(title="选择监控目录")
        if new_dir:
            self.monitor_dir = new_dir
            self.dir_entry.set(new_dir)

    def add_type(self):
        type_name = simpledialog.askstring("添加文件类型", "请输入文件类型名称:")
        if type_name and type_name not in self.file_mappings:
            self.file_mappings[type_name] = []
            self.update_type_listbox()

    def delete_type(self):
        selection = self.type_listbox.curselection()
        if not selection:
            messagebox.showwarning("警告", "请先选择要删除的文件类型")
            return

        selected_type = self.type_listbox.get(selection[0])
        if messagebox.askyesno("确认", f"确定要删除'{selected_type}'吗?"):
            del self.file_mappings[selected_type]
            self.update_type_listbox()
            self.ext_listbox.delete(0, tk.END)

    def add_extension(self):
        selection = self.type_listbox.curselection()
        if not selection:
            messagebox.showwarning("警告", "请先选择文件类型")
            return

        selected_type = self.type_listbox.get(selection[0])
        ext = simpledialog.askstring("添加扩展名", "请输入扩展名(例如 .txt):")
        if ext:
            if not ext.startswith('.'):
                ext = '.' + ext
            if ext not in self.file_mappings[selected_type]:
                self.file_mappings[selected_type].append(ext)
                self.on_type_select(None)

    def delete_extension(self):
        selection = self.type_listbox.curselection()
        if not selection:
            messagebox.showwarning("警告", "请先选择文件类型")
            return

        ext_selection = self.ext_listbox.curselection()
        if not ext_selection:
            messagebox.showwarning("警告", "请先选择要删除的扩展名")
            return

        selected_type = self.type_listbox.get(selection[0])
        selected_ext = self.ext_listbox.get(ext_selection[0])
        self.file_mappings[selected_type].remove(selected_ext)
        self.on_type_select(None)

    def toggle_auto_organize(self):
        self.auto_organize = self.auto_var.get()
        if self.auto_organize and (not self.monitoring_thread or not self.running):
            self.start_monitoring()
        elif not self.auto_organize:
            self.stop_monitoring()

    def apply_interval(self):
        try:
            interval = int(self.interval_var.get())
            if interval > 0:
                self.check_interval = interval
                messagebox.showinfo("成功", f"检查间隔已设置为 {interval} 秒")
            else:
                messagebox.showerror("错误", "间隔时间必须大于0")
        except ValueError:
            messagebox.showerror("错误", "请输入有效的数字")

    def organize_now(self):
        try:
            count = self.organize_files()
            messagebox.showinfo("完成", f"整理完成，共移动了 {count} 个文件")
        except Exception as e:
            messagebox.showerror("错误", f"整理过程中发生错误: {str(e)}")

    def organize_files(self):
        """整理文件到对应目录"""
        count = 0

        # 创建所有需要的目录
        for dir_name in self.file_mappings.keys():
            dir_path = os.path.join(self.monitor_dir, dir_name)
            if not os.path.exists(dir_path):
                os.makedirs(dir_path)

        # 处理未分类文件目录
        unclassified_dir = os.path.join(self.monitor_dir, "未分类文件")
        if not os.path.exists(unclassified_dir):
            os.makedirs(unclassified_dir)

        # 遍历监控目录中的所有文件
        for item in os.listdir(self.monitor_dir):
            item_path = os.path.join(self.monitor_dir, item)

            # 跳过目录和隐藏文件
            if os.path.isdir(item_path) or item.startswith('.'):
                # 跳过我们创建的分类目录
                if item in self.file_mappings or item == "未分类文件":
                    continue

            # 获取文件扩展名
            _, ext = os.path.splitext(item)
            ext = ext.lower()

            # 查找对应的目录
            target_dir = None
            for dir_name, extensions in self.file_mappings.items():
                if ext in extensions:
                    target_dir = os.path.join(self.monitor_dir, dir_name)
                    break

            # 如果没有找到对应目录，使用未分类目录
            if not target_dir:
                target_dir = unclassified_dir

            # 移动文件
            try:
                target_path = os.path.join(target_dir, item)
                # 如果文件已存在，添加编号
                counter = 1
                while os.path.exists(target_path):
                    name, ext = os.path.splitext(item)
                    target_path = os.path.join(target_dir, f"{name}_{counter}{ext}")
                    counter += 1

                shutil.move(item_path, target_path)
                count += 1
            except Exception as e:
                print(f"移动文件 {item} 时出错: {str(e)}")

        return count

    def start_monitoring(self):
        self.running = True
        self.monitoring_thread = threading.Thread(target=self.monitor_loop, daemon=True)
        self.monitoring_thread.start()

    def stop_monitoring(self):
        self.running = False
        if self.monitoring_thread:
            self.monitoring_thread.join()
            self.monitoring_thread = None

    def monitor_loop(self):
        while self.running and self.auto_organize:
            self.organize_files()
            time.sleep(self.check_interval)

    def save_settings(self):
        """保存设置到文件"""
        import json

        settings = {
            'monitor_dir': self.monitor_dir,
            'file_mappings': self.file_mappings,
            'auto_organize': self.auto_organize,
            'check_interval': self.check_interval
        }

        try:
            with open('desktop_organizer_settings.json', 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=4)
            messagebox.showinfo("成功", "设置已保存")
        except Exception as e:
            messagebox.showerror("错误", f"保存设置失败: {str(e)}")

    def load_settings(self):
        """从文件加载设置"""
        import json

        if os.path.exists('desktop_organizer_settings.json'):
            try:
                with open('desktop_organizer_settings.json', 'r', encoding='utf-8') as f:
                    settings = json.load(f)

                self.monitor_dir = settings.get('monitor_dir', self.monitor_dir)
                self.dir_entry.set(self.monitor_dir)

                self.file_mappings = settings.get('file_mappings', self.file_mappings)
                self.update_type_listbox()

                self.auto_organize = settings.get('auto_organize', self.auto_organize)
                self.auto_var.set(self.auto_organize)

                self.check_interval = settings.get('check_interval', self.check_interval)
                self.interval_var.set(str(self.check_interval))

                # 如果设置了自动整理，启动监控
                if self.auto_organize:
                    self.start_monitoring()
            except Exception as e:
                print(f"加载设置失败: {str(e)}")

    def show_about(self):
        messagebox.showinfo("关于",
                            "桌面整理系统 v1.0.0 Yijunzhe1024\n\n可以根据文件类型自动整理指定目录中的文件，并支持最小化到系统托盘进行后台监控。\n\nGithub:https://github.com/yijunzhe1024/desktop_organizer")

    def create_tray_icon(self):
        """创建系统托盘图标"""
        # 创建一个简单的图标
        image = Image.new('RGB', (64, 64), color='blue')
        draw = ImageDraw.Draw(image)
        draw.text((10, 10), "DO", fill='white')

        # 托盘菜单
        menu = Menu(
            item('显示窗口', self.show_window),
            item('立即整理', self.tray_organize),
            item('退出', self.exit_app)
        )

        # 创建托盘图标
        self.tray_icon = Icon("桌面整理系统", image, "桌面整理系统", menu)

        # 启动托盘图标线程
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def show_window(self, icon=None, item=None):
        """显示主窗口"""
        self.root.deiconify()
        self.root.lift()

    def hide_window(self):
        """隐藏主窗口"""
        self.root.withdraw()

    def tray_organize(self, icon=None, item=None):
        """从托盘执行整理"""
        try:
            count = self.organize_files()
            self.tray_icon.notify(f"整理完成，共移动了 {count} 个文件", "桌面整理系统")
        except Exception as e:
            self.tray_icon.notify(f"整理失败: {str(e)}", "桌面整理系统")

    def exit_app(self, icon=None, item=None):
        """退出应用程序"""
        self.stop_monitoring()
        if self.tray_icon:
            self.tray_icon.stop()
        self.root.destroy()
        sys.exit(0)


if __name__ == "__main__":
    root = tk.Tk()

    # 创建应用实例
    app = DesktopOrganizer(root)


    # 处理窗口关闭事件 - 最小化到托盘
    def on_close():
        app.hide_window()


    root.protocol("WM_DELETE_WINDOW", on_close)

    # 创建托盘图标
    app.create_tray_icon()

    # 启动主循环
    root.mainloop()
