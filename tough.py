# Ensure the required packages are installed:
# pip install minimalmodbus pyserial
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

    # 不销毁窗口，仅停止采集
    print("采集已停止，但窗口仍然可用。")

def create_gui():
    """创建包含控制按钮和实时绘图的界面"""
    global root, ax, line, ani  # 声明全局变量以便在其他函数中访问
    root = tk.Tk()
    root.title("LoaDC")  # 修改窗口标题为 LoaDC

    # 创建 matplotlib 图形嵌入到 tkinter 窗口中
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    fig, ax = plt.subplots()
    ax.set_title("F-Time Curve (Real-time)")
    ax.set_xlabel("Time (s)")
    ax.set_ylabel("F (kg)")
    ax.set_facecolor("white")
    line, = ax.plot([], [], label="F (kg)", color="blue")
    ax.legend()

    def init():
        """初始化绘图"""
        ax.set_xlim(0, 10)
        ax.set_ylim(-1, 1)  # 初始纵坐标范围
        return line,

    def update(frame):
        """更新绘图"""
        if len(time_data) > 0:
            ax.set_xlim(0, max(time_data) + 1)
            max_pressure = max(pressure_data) if pressure_data else 1  # 获取当前压力的最大值
            min_pressure = min(pressure_data) if pressure_data else 0  # 获取当前压力的最小值
            if max_pressure == min_pressure:  # 处理压力值相等的情况
                max_pressure += 1
            ax.set_ylim(min_pressure - 1, max_pressure + 1)  # 动态调整纵坐标范围
            line.set_data(time_data, pressure_data)
        return line,

    ani = animation.FuncAnimation(fig, update, init_func=init, blit=True, interval=50)

    # 将 matplotlib 图形嵌入到 tkinter 窗口中
    canvas = FigureCanvasTkAgg(fig, master=root)
    canvas_widget = canvas.get_tk_widget()
    canvas_widget.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

    # 创建控制按钮
    button_frame = tk.Frame(root)
    button_frame.pack(side=tk.BOTTOM, fill=tk.X)

    start_button = tk.Button(button_frame, text="启动采集", command=start_data_collection)
    start_button.pack(side=tk.LEFT, padx=5, pady=5)

    stop_button = tk.Button(button_frame, text="停止采集", command=stop_data_collection)
    stop_button.pack(side=tk.LEFT, padx=5, pady=5)

    zero_button = tk.Button(button_frame, text="调零", command=manual_zero)
    zero_button.pack(side=tk.LEFT, padx=5, pady=5)

    path_button = tk.Button(button_frame, text="选择保存路径", command=select_save_path)  # 添加选择路径按钮
    path_button.pack(side=tk.LEFT, padx=5, pady=5)

    exit_button = tk.Button(button_frame, text="退出程序", command=lambda: root.after(0, root.destroy))
    exit_button.pack(side=tk.LEFT, padx=5, pady=5)

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
