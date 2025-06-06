import os
import platform
import subprocess
import sys
import threading
import time
import tkinter as tk
from concurrent.futures import ThreadPoolExecutor, as_completed
from tkinter import filedialog, messagebox, ttk
import psutil

# 支持的视频文件扩展名
VIDEO_EXTENSIONS = ['.mp4', '.mkv', '.flv', '.avi', '.mov', '.wmv', '.m4s']


def find_video_files(root_dir):
    """递归查找所有视频文件"""
    video_files = []
    for root, dirs, files in os.walk(root_dir):
        for file in files:
            ext = os.path.splitext(file)[1].lower()
            if ext in VIDEO_EXTENSIONS:
                video_files.append(os.path.join(root, file))
        # 限制目录深度，避免搜索过深
        dirs[:] = [d for d in dirs if not d.startswith('.')]  # 跳过隐藏目录
    return video_files


def get_video_name(video_path):
    """从视频文件路径中提取文件名（不含扩展名）"""
    filename = os.path.basename(video_path)
    name, _ = os.path.splitext(filename)

    # 清理文件名中的非法字符
    invalid_chars = {'\\', '/', ':', '*', '?', '"', '<', '>', '|'}
    clean_name = ''.join(c for c in name if c not in invalid_chars)

    # 限制文件名长度
    if len(clean_name) > 150:
        clean_name = clean_name[:150]

    return clean_name


def process_single_file(video_path, output_dir, progress_callback=None):
    """处理单个视频文件（带进度回调）"""
    try:
        print(f"Processing: {video_path}")
        video_name = get_video_name(video_path)
        if not video_name:
            print(f"Invalid video name for: {video_path}")
            if progress_callback:
                progress_callback(False)
            return False

        os.makedirs(output_dir, exist_ok=True)
        final_mp3 = os.path.join(output_dir, f'{video_name}.mp3')

        if os.path.exists(final_mp3) and os.path.getsize(final_mp3) > 0:
            print(f"Skipping existing file: {final_mp3}")
            if progress_callback:
                progress_callback(True)
            return True

        temp_aac = os.path.splitext(final_mp3)[0] + "_temp.aac"

        try:
            print(f"Extracting audio for: {video_path}")
            extract_result = subprocess.run([
                'ffmpeg',
                '-hide_banner',
                '-loglevel', 'error',
                '-i', video_path,
                '-c:a', 'copy',
                '-vn',
                '-y',
                temp_aac
            ], capture_output=True, text=True)

            if extract_result.returncode != 0:
                raise RuntimeError(f"AAC extraction failed: {extract_result.stderr}")

            print(f"Converting to MP3: {video_path}")
            convert_result = subprocess.run([
                'ffmpeg',
                '-hide_banner',
                '-loglevel', 'error',
                '-i', temp_aac,
                '-c:a', 'libmp3lame',
                '-q:a', '0',
                '-y',
                final_mp3
            ], capture_output=True, text=True)

            if convert_result.returncode != 0:
                raise RuntimeError(f"MP3 conversion failed: {convert_result.stderr}")

            print(f"Successfully converted: {video_path} -> {final_mp3}")
            if progress_callback:
                progress_callback(True)
            return True

        except Exception as e:
            print(f"Error processing {video_path}: {str(e)}")
            if progress_callback:
                progress_callback(False)
            return False

        finally:
            if os.path.exists(temp_aac):
                try:
                    os.remove(temp_aac)
                except Exception as e:
                    print(f"Failed to remove temp file {temp_aac}: {str(e)}")

    except Exception as e:
        print(f"Unexpected error processing {video_path}: {str(e)}")
        if progress_callback:
            progress_callback(False)
        return False


def process_folders_parallel(input_dirs, output_dir, progress_callback=None, max_workers=None):
    """并行处理多个输入文件夹"""
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # 收集所有视频文件
    all_video_files = []
    for input_dir in input_dirs:
        all_video_files.extend(find_video_files(input_dir))

    total = len(all_video_files)
    if total == 0:
        print("没有找到任何视频文件")
        return 0, 0

    # 如果没有指定最大工作线程数，则自动设置
    if max_workers is None:
        logical_cores = psutil.cpu_count(logical=True)
        max_workers = max(1, logical_cores - 1)  # 确保至少1个线程

    # 使用线程池并行处理
    success = 0
    processed = 0
    lock = threading.Lock()  # 用于线程安全的计数

    def task_wrapper(video_path):
        nonlocal success, processed
        result = process_single_file(video_path, output_dir,
                                    lambda success_flag: progress_callback(success_flag) if progress_callback else None)
        with lock:
            processed += 1
            if result:
                success += 1
            # 更新进度
            if progress_callback:
                progress_callback(None)  # 触发进度更新
        return result

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = [executor.submit(task_wrapper, video_path) for video_path in all_video_files]

        # 等待所有任务完成
        for future in as_completed(futures):
            try:
                future.result()  # 捕获任何异常
            except Exception as e:
                print(f"处理过程中发生异常: {str(e)}")

    return total, success


class ProgressWindow:
    """进度显示窗口"""

    def __init__(self, total_tasks):
        self.root = tk.Tk()
        self.root.title("处理进度")
        self.root.geometry("400x200")
        self.root.resizable(False, False)

        # 防止重复创建Tk实例
        if hasattr(ProgressWindow, 'instance_created'):
            self.root.destroy()
            return
        ProgressWindow.instance_created = True

        self.total = total_tasks
        self.completed = 0
        self.success = 0
        self.start_time = time.time()
        self.running = True

        # 创建框架容器
        frame = ttk.Frame(self.root, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)

        self.label = ttk.Label(frame, text=f"准备处理 {total_tasks} 个文件...", font=("Arial", 10))
        self.label.pack(pady=5)

        self.progress = ttk.Progressbar(frame, orient="horizontal", length=380, mode="determinate")
        self.progress.pack(pady=5)

        # 状态信息框架
        status_frame = ttk.Frame(frame)
        status_frame.pack(fill=tk.X, pady=5)

        self.status = ttk.Label(status_frame, text="状态: 就绪")
        self.status.pack(side=tk.LEFT, padx=5)

        self.time_label = ttk.Label(status_frame, text="已用时间: 0秒")
        self.time_label.pack(side=tk.RIGHT, padx=5)

        # 统计信息框架
        stats_frame = ttk.Frame(frame)
        stats_frame.pack(fill=tk.X, pady=5)

        self.success_label = ttk.Label(stats_frame, text="成功: 0")
        self.success_label.pack(side=tk.LEFT, padx=5)

        self.failed_label = ttk.Label(stats_frame, text="失败: 0")
        self.failed_label.pack(side=tk.LEFT, padx=5)

        self.remaining_label = ttk.Label(stats_frame, text="剩余: 0")
        self.remaining_label.pack(side=tk.RIGHT, padx=5)

        # 速度信息
        self.speed_label = ttk.Label(frame, text="速度: 0 文件/秒")
        self.speed_label.pack(pady=5)

        # 按钮框架
        button_frame = ttk.Frame(frame)
        button_frame.pack(fill=tk.X, pady=5)

        self.cancel_button = ttk.Button(button_frame, text="取消", command=self.cancel)
        self.cancel_button.pack(side=tk.RIGHT)

        self.root.protocol("WM_DELETE_WINDOW", self.cancel)
        self.root.after(100, self.update_time)

    def update(self, success_flag=None):
        if not self.running:
            return

        if success_flag is not None:
            if success_flag:
                self.success += 1
            self.completed += 1

        progress_value = min(100, int((self.completed / self.total) * 100))
        self.progress["value"] = progress_value

        elapsed = time.time() - self.start_time
        speed = self.completed / elapsed if elapsed > 0 else 0

        self.label["text"] = f"处理中: {self.completed}/{self.total} ({progress_value}%)"
        self.status["text"] = "状态: 处理中..." if self.completed < self.total else "状态: 完成!"
        self.success_label["text"] = f"成功: {self.success}"
        self.failed_label["text"] = f"失败: {self.completed - self.success}"
        self.remaining_label["text"] = f"剩余: {self.total - self.completed}"
        self.speed_label["text"] = f"速度: {speed:.2f} 文件/秒"

        if self.completed >= self.total:
            self.cancel_button["text"] = "关闭"
            self.status["text"] = "状态: 处理完成!"
            self.running = False

    def update_time(self):
        if self.running:
            elapsed = time.time() - self.start_time
            self.time_label["text"] = f"已用时间: {elapsed:.1f}秒"
            self.root.after(1000, self.update_time)

    def cancel(self):
        self.running = False
        self.root.destroy()

    def close(self):
        self.running = False
        self.root.destroy()


def select_folders():
    """使用GUI选择多个文件夹（优化版）"""
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    print("请选择要处理的文件夹...")
    # 允许选择多个文件夹
    folders = filedialog.askdirectory(title="选择要处理的文件夹", mustexist=True)

    if not folders:
        print("没有选择文件夹")
        return []

    # 返回列表
    return [folders] if isinstance(folders, str) else list(folders)


def select_output_dir():
    """使用GUI选择输出目录"""
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    print("请选择输出目录...")
    output_dir = filedialog.askdirectory(title="选择输出目录", mustexist=False)

    if not output_dir:
        print("没有选择输出目录")
        return None

    return output_dir


def main():
    print("=== 一键.mp4To.mp3脚本 ===")

    # 检查ffmpeg是否可用
    try:
        subprocess.run(['ffmpeg', '-version'], check=True,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        print("FFmpeg 可用")
    except Exception as e:
        print("错误: 没有找到ffmpeg或无法执行。请确保ffmpeg已安装并添加到系统PATH中。")
        messagebox.showerror("错误", "没有找到FFmpeg! 请确保FFmpeg已安装并添加到系统PATH中。")
        return

    # 选择输入文件夹
    input_dirs = select_folders()
    if not input_dirs:
        return

    # 选择输出目录
    output_dir = select_output_dir()
    if not output_dir:
        return

    # 显示进度窗口
    total_files = sum(len(find_video_files(d)) for d in input_dirs)
    if total_files == 0:
        messagebox.showinfo("提示", "没有找到任何视频文件")
        return

    progress_window = ProgressWindow(total_files)

    # 启动处理线程
    def processing_thread():
        # 根据系统调整并行度
        max_workers = None
        if platform.system() == "Windows":
            max_workers = psutil.cpu_count(logical=True) - 1  # Windows上限制并发数

        total, success = process_folders_parallel(
            input_dirs,
            output_dir,
            lambda _: progress_window.update(),  # 使用lambda确保回调能触发
            max_workers
        )

        # 处理完成后关闭窗口
        progress_window.close()

        # 显示完成消息
        message = f"处理完成!\n共处理 {total} 个文件, 成功 {success} 个"
        print(message)
        messagebox.showinfo("完成", message)

        # 打开输出文件夹
        if os.name == 'nt':  # Windows
            os.startfile(output_dir)
        elif os.name == 'posix':  # macOS/Linux
            opener = 'open' if sys.platform == 'darwin' else 'xdg-open'
            subprocess.run([opener, output_dir], check=False)

    # 启动处理线程
    thread = threading.Thread(target=processing_thread, daemon=True)
    thread.start()

    # 启动主事件循环
    progress_window.root.mainloop()

    # 等待处理线程完成
    thread.join()


if __name__ == "__main__":
    main()