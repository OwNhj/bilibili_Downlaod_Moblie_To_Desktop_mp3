import json
import os
import platform
import subprocess
import sys
import threading
import time
import tkinter as tk
from concurrent.futures import ThreadPoolExecutor, as_completed
from functools import lru_cache
from tkinter import filedialog, messagebox, ttk

import psutil


def find_entry_json_files(root_dir):
    """递归查找所有entry.json文件（优化版）"""
    json_files = []
    for root, dirs, files in os.walk(root_dir):
        if 'entry.json' in files:
            json_files.append(os.path.join(root, 'entry.json'))
        # 限制目录深度，避免搜索过深
        dirs[:] = [d for d in dirs if not d.startswith('.')]  # 跳过隐藏目录
    return json_files


@lru_cache(maxsize=None)
def find_audio_file_cached(json_dir):
    """缓存音频文件查找结果"""
    return find_audio_file(json_dir)


def find_audio_file(json_dir):
    """在entry.json所在目录及其子目录中查找音频文件（优化版）"""
    # 支持的音频文件扩展名（按优先级排序）
    audio_extensions = ['.m4a', '.mp4', '.aac', '.flv', '.m4s']  # 更常见的音频格式优先

    # 先在entry.json同目录查找
    for ext in audio_extensions:
        audio_path = os.path.join(json_dir, f'audio{ext}')
        if os.path.exists(audio_path):
            return audio_path

    # 优化子目录搜索 - 只搜索常见音频目录
    common_audio_dirs = {'audio', 'sound', 'voice', 'music'}
    for root, dirs, files in os.walk(json_dir):
        # 限制搜索深度
        if root.count(os.sep) - json_dir.count(os.sep) > 2:
            del dirs[:]
            continue

        # 优先搜索常见音频目录
        dirs[:] = [d for d in dirs if d.lower() in common_audio_dirs or not d.startswith('.')]

        for file in files:
            file_lower = file.lower()
            if any(file_lower.endswith(ext) for ext in audio_extensions):
                # 优先匹配以"audio"开头的文件
                if file_lower.startswith('audio'):
                    return os.path.join(root, file)
                # 返回第一个找到的音频文件（如果不需要严格匹配audio开头）
                return os.path.join(root, file)

    return None


@lru_cache(maxsize=None)
def extract_title_name_cached(json_path):
    """缓存提取的title名称"""
    return extract_title_name(json_path)


def extract_title_name(json_path):
    """从entry.json中提取part字段（优化版）"""
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
            part_name = data.get('title', 'untitled')
            # 清理文件名中的非法字符（优化清理逻辑）
            invalid_chars = {'\\', '/', ':', '*', '?', '"', '<', '>', '|'}
            part_name = ''.join(c for c in part_name if c not in invalid_chars)
            # 限制文件名长度
            if len(part_name) > 150:
                part_name = part_name[:150]
            return part_name
    except Exception as e:
        print(f"Error reading {json_path}: {e}")
        return None


def process_single_file(json_path, output_dir, progress_callback=None):
    """处理单个entry.json对应的音频文件（带进度回调）"""
    try:
        # 使用缓存方法
        part_name = extract_title_name_cached(json_path)
        if not part_name:
            if progress_callback:
                progress_callback(False)
            return False

        # 使用缓存方法查找音频文件
        json_dir = os.path.dirname(json_path)
        audio_path = find_audio_file_cached(json_dir)
        if not audio_path:
            print(f"Audio file not found in {json_dir} or its subdirectories")
            if progress_callback:
                progress_callback(False)
            return False

        # 创建输出目录
        os.makedirs(output_dir, exist_ok=True)

        # 最终MP3文件路径
        final_mp3 = os.path.join(output_dir, f'{part_name}.mp3')

        # 检查文件是否已存在且不需要重新转换
        if os.path.exists(final_mp3) and os.path.getsize(final_mp3) > 0:
            print(f"Skipping existing file: {final_mp3}")
            if progress_callback:
                progress_callback(True)
            return True

        # 直接从音频文件转换为MP3（优化FFmpeg参数）
        result = subprocess.run([
            'ffmpeg',
            '-hide_banner',  # 隐藏不必要的输出
            '-loglevel', 'error',  # 只显示错误信息
            '-i', audio_path,
            '-c:a', 'libmp3lame',
            '-q:a', '0',
            '-y',
            final_mp3
        ], capture_output=True, text=True)

        if result.returncode == 0:
            print(f"Successfully converted: {audio_path} -> {final_mp3}")
            if progress_callback:
                progress_callback(True)
            return True
        else:
            print(f"FFmpeg error processing {json_path}:\n{result.stderr}")
            if progress_callback:
                progress_callback(False)
            return False

    except Exception as e:
        print(f"Error processing {json_path}: {str(e)}")
        if progress_callback:
            progress_callback(False)
        return False


def process_folders_parallel(input_dirs, output_dir, progress_callback=None, max_workers=None):
    """并行处理多个输入文件夹"""
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # 收集所有JSON文件
    all_json_files = []
    for input_dir in input_dirs:
        all_json_files.extend(find_entry_json_files(input_dir))

    total = len(all_json_files)
    if total == 0:
        print("没有找到任何entry.json文件")
        return 0, 0

    # 如果没有指定最大工作线程数，则自动设置
    if max_workers is None:
        logical_cores = psutil.cpu_count(logical=True)
        max_workers = logical_cores - 1

    # 使用线程池并行处理
    success = 0
    processed = 0
    lock = threading.Lock()  # 用于线程安全的计数

    def task_wrapper(json_path):
        nonlocal success, processed
        result = process_single_file(json_path, output_dir, progress_callback)
        with lock:
            processed += 1
            if result:
                success += 1
        return result

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = [executor.submit(task_wrapper, json_path) for json_path in all_json_files]

        # 等待所有任务完成
        for future in as_completed(futures):
            future.result()  # 捕获任何异常

    return total, success


class ProgressWindow:
    """进度显示窗口"""

    def __init__(self, total_tasks):
        self.root = tk.Tk()
        self.root.title("处理进度")
        self.root.geometry("400x200")
        self.root.resizable(False, False)

        # 防止重复创建Tk实例
        if not hasattr(ProgressWindow, 'instance_created'):
            ProgressWindow.instance_created = True
        else:
            self.root.destroy()
            return

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

    def update(self, success_flag):
        if not self.running:
            return

        self.completed += 1
        if success_flag:
            self.success += 1

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
    print("=== 音频提取转换工具 ===")

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
    total_files = sum(len(find_entry_json_files(d)) for d in input_dirs)
    if total_files == 0:
        messagebox.showinfo("提示", "没有找到任何entry.json文件")
        return

    progress_window = ProgressWindow(total_files)

    # 启动处理线程
    def processing_thread():
        # 根据系统调整并行度
        max_workers = None
        if platform.system() == "Windows":
            max_workers = min(8, (os.cpu_count() or 1))  # Windows上限制并发数

        total, success = process_folders_parallel(
            input_dirs,
            output_dir,
            progress_window.update,
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
    thread.join(timeout=1)


if __name__ == "__main__":
    main()