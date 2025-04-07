# Ensure the required packages are installed:
# pip install minimalmodbus pyserial
from datetime import datetime  # 用于记录时间
import minimalmodbus  # type: ignore
import time
import logging
import matplotlib.pyplot as plt  # 用于绘制压力-时间曲线
import threading  # 用于实现手动停止功能

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

# 调用读取压力值函数
try:
    zero_offset = None  # 用于存储调零偏移值
    stop_thread = threading.Thread(target=stop_program)  # 启动停止线程
    stop_thread.start()

    while not stop_flag:
        pressure_value = read_pressure()
        if pressure_value is not None:
            current_time = datetime.now()  # 获取当前时间
            if zero_offset is None:
                # 第一次读取时进行调零
                zero_offset = pressure_value
                print(f"调零完成，偏移值: {zero_offset:.2f} kg")
            else:
                # 后续读取减去调零偏移值
                adjusted_pressure = pressure_value - zero_offset
                print(f"实时压力值: {adjusted_pressure:.2f} kg")
                
                # 记录压力值和时间
                pressure_data.append(adjusted_pressure)
                time_data.append(current_time)
        else:
            print("读取压力值失败，已记录日志。")
        time.sleep(1)  # 采样间隔（根据需求调整）
except KeyboardInterrupt:
    print("程序终止")
finally:
    instrument.serial.close()
    
    # 绘制压力-时间曲线
    if pressure_data and time_data:
        plt.figure(figsize=(10, 6))
        plt.plot(time_data, pressure_data, label="压力值 (kg)")
        plt.xlabel("时间")
        plt.ylabel("压力 (kg)")
        plt.title("压力-时间曲线")
        plt.legend()
        plt.grid()
        plt.gcf().autofmt_xdate()  # 自动格式化时间轴
        plt.show()
