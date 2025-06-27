# Ensure the required packages are installed:
# pip install minimalmodbus pyserial pandas openpyxl
from datetime import datetime  # 用于记录时间
import minimalmodbus  # type: ignore
import time
import logging
import matplotlib.pyplot as plt  # 用于绘制压力-时间曲线
import threading  # 用于实现手动停止功能
import os  # 用于创建文件夹
from threading import Thread, Event  # 用于线程控制
import tkinter as tk  # 用于创建按钮界面
import matplotlib.animation as animation  # 用于实时动态绘图
import matplotlib  # 用于设置后端
from tkinter import filedialog  # 导入文件对话框模块
import pandas as pd  # 导入 pandas 用于生成 Excel 文件
matplotlib.use("TkAgg")  # 设置 TkAgg 后端以兼容 tkinter

# 配置日志记录
logging.basicConfig(filename="error.log", level=logging.ERROR, format="%(asctime)s - %(message)s")

# 配置 Modbus 仪器
instrument = minimalmodbus.Instrument(port="COM7", slaveaddress=1)
instrument.serial.baudrate = 9600
instrument.serial.bytesize = 8
instrument.serial.parity = minimalmodbus.serial.PARITY_NONE
instrument.serial.stopbits = 1
instrument.serial.timeout = 1  # 超时时间（秒）

# 初始化存储压力值和时间戳的列表
pressure_data = []
time_data = []
stop_flag = False  # 用于控制程序停止的标志

# 初始化停止事件
stop_event = Event()

# 创建桌面上的 "test" 文件夹
desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
image_folder = os.path.join(desktop_path, "test")
os.makedirs(image_folder, exist_ok=True)

def stop_program():
    global stop_flag
    input("按下回车键停止程序...")
    stop_flag = True

def check_serial_port():
    try:
        # 检查串口是否可用
        instrument.serial.open()
    except PermissionError as e:
        logging.error(f"串口访问被拒绝: {e}")
        print("串口访问被拒绝。请检查以下内容：")
        print("1. 确保没有其他程序占用串口。")
        print("2. 检查当前用户是否有权限访问串口。")
        print("   - 在 Windows 上，尝试以管理员身份运行程序。")
        print("   - 在 Linux 上，确保用户属于 `dialout` 或 `tty` 组。")
        exit(1)
    except Exception as e:
        logging.error(f"无法打开串口: {e}")
        print(f"无法打开串口: {e}")
        instrument.serial.close()

# 检查串口访问权限
# check_serial_port()

def send_command():
    try:
        # 发送指令：01 03 00 00 00 01 84 0A
        command = bytes.fromhex("01 03 00 00 00 01 84 0A")
        instrument.serial.write(command)
        time.sleep(0.1)  # 等待一段时间以确保指令被处理
    except Exception as e:
        logging.error(f"发送指令失败: {e}")

def read_pressure(retries=3):
    for attempt in range(retries):
        try:
            # 发送指令
            send_command()
            
            # 读取串口返回的原始字节数据
            raw_data = instrument.serial.read(7)  # 假设返回数据长度为7字节
            if len(raw_data) < 5:
                raise ValueError("返回的数据长度不足，无法解析")
            
            # 提取第四个和第五个字节
            byte4 = raw_data[3]
            byte5 = raw_data[4]
            
            # 将两个字节转换为浮点数（假设为大端字节序）
            value = (byte4 << 8) | byte5
            pressure = value / 100.0  # 假设需要除以100得到两位小数的浮点数
            
            return pressure
        except Exception as e:
            logging.error(f"读取失败 (尝试 {attempt + 1}/{retries}): {e}")
            time.sleep(1)  # 等待一段时间后重试
    return None  # 多次重试失败后返回 None

def start_realtime_plot():
    """启动实时动态绘图"""
    def plot():
        fig, ax = plt.subplots()
        ax.set_title("F-Time Curve (Real-time)")  # 修改括号内容
        ax.set_xlabel("Time (s)")
        ax.set_ylabel("F (kg)")
        ax.set_facecolor("white")  # 设置背景为白色
        line, = ax.plot([], [], label="F (kg)", color="blue")
        ax.legend()

        def init():
            """初始化绘图"""
            ax.set_xlim(0, 10)  # 初始 X 轴范围
            ax.set_ylim(-1, 1)  # 初始 Y 轴范围
            return line,

        def update(frame):
            """更新绘图"""
            if len(time_data) > 0:
                ax.set_xlim(0, max(time_data) + 1)  # 动态调整 X 轴范围
                ax.set_ylim(min(pressure_data) - 1, max(pressure_data) + 1)  # 动态调整 Y 轴范围
                line.set_data(time_data, pressure_data)
            return line,

        ani = animation.FuncAnimation(fig, update, init_func=init, blit=True, interval=50)
        plt.show()

    Thread(target=plot).start()  # 将绘图放入单独的线程中

def start_data_collection():
    """启动数据采集"""
    global zero_offset, start_time, pressure_data, time_data  # 声明全局变量
    zero_offset = None  # 初始化 zero_offset
    start_time = None
    pressure_data = []
    time_data = []
    stop_event.clear()
    print("数据采集已启动...")

    def collect_data():
        global zero_offset  # 在子线程中声明 zero_offset 为全局变量
        while not stop_event.is_set():
            pressure_value = read_pressure()
            if pressure_value is not None:
                current_time = datetime.now()
                if zero_offset is None:
                    zero_offset = pressure_value  # 调零时赋值
                    start_time = current_time
                    print(f"调零完成，偏移值: {zero_offset:.2f} kg")
                else:
                    adjusted_pressure = pressure_value - zero_offset
                    elapsed_time = (current_time - start_time).total_seconds()
                    print(f"实时压力值: {adjusted_pressure:.2f} kg, 时间: {elapsed_time:.2f} 秒")
                    pressure_data.append(adjusted_pressure)
                    time_data.append(elapsed_time)
            else:
                print("读取压力值失败，已记录日志。")
            time.sleep(0.05)

        print("数据采集线程已停止。")

    # 启动数据采集线程
    Thread(target=collect_data, daemon=True).start()

def manual_zero():
    """手动调零功能"""
    global zero_offset
    pressure_value = read_pressure()
    if pressure_value is not None:
        zero_offset = pressure_value
        print(f"手动调零完成，偏移值: {zero_offset:.2f} kg")
    else:
        print("调零失败，无法读取压力值。")

def select_save_path():
    """让用户选择图片保存路径"""
    global image_folder
    selected_folder = filedialog.askdirectory(initialdir=desktop_path, title="选择图片保存路径")
    if selected_folder:  # 如果用户选择了路径
        image_folder = selected_folder
        print(f"图片保存路径已更改为: {image_folder}")
    else:
        print("使用默认路径: 桌面的 test 文件夹")

def stop_data_collection():
    """停止数据采集"""
    stop_event.set()
    print("数据采集已停止。")
    if pressure_data and time_data:
        # 保存图片
        plt.figure(figsize=(10, 6))
        plt.plot(time_data, pressure_data, label="F (kg)", marker="o")
        plt.xlabel("Time (s)")
        plt.ylabel("F (kg)")
        plt.title("F-Time Curve")
        plt.legend()
        plt.gca().set_facecolor("white")  # 设置坐标轴背景为白色
        plt.grid(False)  # 去掉背景中的格子

        # 在曲线上显示每一次压力值的峰值
        for i in range(1, len(pressure_data) - 1):
            if (pressure_data[i] >= pressure_data[i - 1]) and (pressure_data[i] > pressure_data[i + 1]):
                plt.text(time_data[i], pressure_data[i], f"{pressure_data[i]:.2f}", fontsize=8, ha="center", va="bottom", color="red")

        # 保存图片到用户选择的路径或默认路径
        image_path = os.path.join(image_folder, f"采集结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png")
        plt.savefig(image_path)
        print(f"图片已保存到: {image_path}")
        plt.close()  # 关闭图形，避免与 Tkinter 冲突

        # 保存数据到 Excel 文件
        excel_path = os.path.join(image_folder, f"采集数据_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        try:
            data = {"Time (s)": time_data, "F (kg)": pressure_data}
            df = pd.DataFrame(data)
            df.to_excel(excel_path, index=False)
            print(f"Excel 数据已保存到: {excel_path}")
        except Exception as e:
            print(f"保存 Excel 文件时发生错误: {e}")

    # 不销毁窗口，仅停止采集
    print("采集已停止，但窗口仍然可用。")

def create_gui():
    """创建包含控制按钮和实时绘图的界面（美化布局并增大字号）"""
    global root, ax, line, ani  # 声明全局变量以便在其他函数中访问
    root = tk.Tk()
    root.title("LoaDC")
    root.configure(bg="#f4f8fb")

    # 主frame，分为左右两区
    main_frame = tk.Frame(root, bg="#f4f8fb")
    main_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)

    # 数据展示区（左侧，带边框）
    data_frame = tk.LabelFrame(
        main_frame, text="Data Display", bg="#f4f8fb", fg="#1a355e",
        bd=2, relief=tk.GROOVE, font=("Arial", 16, "bold")
    )
    data_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 12), pady=0)

    # 控制按钮区（右侧，带边框）
    control_frame = tk.LabelFrame(
        main_frame, text="Controls", bg="#f4f8fb", fg="#1a355e",
        bd=2, relief=tk.GROOVE, font=("Arial", 16, "bold")
    )
    control_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=0, pady=0)

    # 图表区
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    fig, ax = plt.subplots(figsize=(7, 4))
    fig.patch.set_facecolor("#f4f8fb")
    ax.set_title("F-Time Curve (Real-time)", color="#1a355e", fontsize=18)
    ax.set_xlabel("Time (s)", fontsize=14)
    ax.set_ylabel("F (kg)", fontsize=14)
    ax.set_facecolor("#ffffff")
    ax.tick_params(axis='both', labelsize=13)
    ax.grid(True, color="#d0d7e5", linestyle="--", linewidth=0.7, alpha=0.7)
    line, = ax.plot([], [], label="F (kg)", color="#2176ae", linewidth=2)
    ax.legend(fontsize=13)

    # 初始纵坐标范围为 0-10kg
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 10)

    def init():
        ax.set_xlim(0, 10)
        ax.set_ylim(0, 10)
        return line,

    def update(frame):
        if len(time_data) > 0:
            ax.set_xlim(0, max(time_data) + 1)
            line.set_data(time_data, pressure_data)
        return line,

    ani = animation.FuncAnimation(fig, update, init_func=init, blit=True, interval=50)

    canvas = FigureCanvasTkAgg(fig, master=data_frame)
    canvas_widget = canvas.get_tk_widget()
    canvas_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    # 控制按钮区美化
    def style_button(btn):
        btn.configure(
            bg="#e6eef8", fg="#1a355e", activebackground="#d0e2f2", activeforeground="#1a355e",
            relief=tk.FLAT, bd=0, width=18, height=2, cursor="hand2", font=("Arial", 14, "bold")
        )

    start_button = tk.Button(control_frame, text="Start Collection", command=start_data_collection)
    style_button(start_button)
    start_button.pack(fill=tk.X, padx=12, pady=(12, 4))

    stop_button = tk.Button(control_frame, text="Stop Collection", command=stop_data_collection)
    style_button(stop_button)
    stop_button.pack(fill=tk.X, padx=12, pady=4)

    zero_button = tk.Button(control_frame, text="Zero", command=manual_zero)
    style_button(zero_button)
    zero_button.pack(fill=tk.X, padx=12, pady=4)

    path_button = tk.Button(control_frame, text="Select Save Path", command=select_save_path)
    style_button(path_button)
    path_button.pack(fill=tk.X, padx=12, pady=4)

    def set_ylim_custom():
        top = tk.Toplevel(root)
        top.title("Set Y Axis Range")
        tk.Label(top, text="Y axis min (kg):", font=("Arial", 13)).grid(row=0, column=0, padx=5, pady=5)
        tk.Label(top, text="Y axis max (kg):", font=("Arial", 13)).grid(row=1, column=0, padx=5, pady=5)
        min_var = tk.StringVar(value="0")
        max_var = tk.StringVar(value="10")
        min_entry = tk.Entry(top, textvariable=min_var, font=("Arial", 13))
        max_entry = tk.Entry(top, textvariable=max_var, font=("Arial", 13))
        min_entry.grid(row=0, column=1, padx=5, pady=5)
        max_entry.grid(row=1, column=1, padx=5, pady=5)
        def apply():
            try:
                ymin = float(min_var.get())
                ymax = float(max_var.get())
                ax.set_ylim(ymin, ymax)
                canvas.draw_idle()
                top.destroy()
            except Exception:
                tk.messagebox.showerror("Error", "Please enter valid numbers.")
        apply_button = tk.Button(top, text="Apply", command=apply, font=("Arial", 13, "bold"))
        apply_button.grid(row=2, column=0, columnspan=2, pady=5)

    ylim10_button = tk.Button(control_frame, text="Y axis 0-10kg", command=lambda: (ax.set_ylim(0, 10), canvas.draw_idle()))
    style_button(ylim10_button)
    ylim10_button.pack(fill=tk.X, padx=12, pady=4)
    ylim50_button = tk.Button(control_frame, text="Y axis 0-50kg", command=lambda: (ax.set_ylim(0, 50), canvas.draw_idle()))
    style_button(ylim50_button)
    ylim50_button.pack(fill=tk.X, padx=12, pady=4)
    ylim_custom_button = tk.Button(control_frame, text="Custom Y axis", command=set_ylim_custom)
    style_button(ylim_custom_button)
    ylim_custom_button.pack(fill=tk.X, padx=12, pady=4)

    exit_button = tk.Button(control_frame, text="Exit", command=lambda: root.after(0, root.destroy))
    style_button(exit_button)
    exit_button.pack(fill=tk.X, padx=12, pady=(4, 12))

    # 悬停效果
    def on_enter(e): e.widget.config(bg="#d0e2f2")
    def on_leave(e): e.widget.config(bg="#e6eef8")
    for btn in [start_button, stop_button, zero_button, path_button, ylim10_button, ylim50_button, ylim_custom_button, exit_button]:
        btn.bind("<Enter>", on_enter)
        btn.bind("<Leave>", on_leave)

    root.mainloop()

# 主程序入口
if __name__ == "__main__":
    try:
        create_gui()
    except KeyboardInterrupt:
        print("程序终止")
    finally:
        if not stop_event.is_set():
            stop_data_collection()
        instrument.serial.close()
        if not stop_event.is_set():
            stop_data_collection()
        instrument.serial.close()
