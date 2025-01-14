from tkinter import *
from tkinter import messagebox
from tkinter.scrolledtext import ScrolledText
import requests
import serial
import serial.tools.list_ports
import threading
import queue
import datetime
from openpyxl import Workbook

__version_ = '1.0.0'


def check_for_update():
    try:
        response = requests.get("https://github.com/906168212/SprayMonitorSystem/blob/master/version.json")
        if response.status_code == 200:
            data = response.text
            print(data)
            #latest_version = data['latest_version']
            #print(latest_version)
    except Exception as e:
        print(f"检查更新失败：{e}")


class Concert(Frame):
    '''初始化'''

    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.window_width = 900
        self.window_height = 600
        self.ser = None
        self.runing = False
        self.auto_read_flow_flag = True
        self.serial_lock = threading.Lock()  # 创建串口线程锁，保证收发数据的正常
        self.port_list = ['']
        self.baudrate_list = ['9600', '19200', '38400', '57600', '115200']
        self.send_queue = queue.Queue()  # 创建发送任务队列
        self.xlsx_name = ''
        check_for_update()
        self.version = 'V1.1.0'
        self.createWidget()
        self.excel_init()

    '''日志系统'''

    def excel_init(self):
        try:
            self.workbook = Workbook()  # 创建工作薄
            self.worksheet = self.workbook.active  # 选择默认工作表
            self.worksheet.merge_cells('B1:E1')  # 透传固定帧
            self.worksheet.merge_cells('F1:I1')  # CAN ID
            self.worksheet.merge_cells('J1:Q1')  # CAN 数据
            self.worksheet._current_row = 0
            self.worksheet.append(
                {1: '时间', 2: '透传固定帧', 6: 'CAN ID', 10: 'CAN数据'})  # ['时间','透传固定帧','CAN ID','CAN数据']
            self.xlsx_name = '试验CAN通信数据日志-' + str(
                datetime.datetime.now().strftime('%Y-%m-%d-%H-%M-%S')) + '.xlsx'
            self.workbook.save(self.xlsx_name)  # 保存工作薄
        except Exception as e:
            messagebox.showerror('Error', '初始化Excel时出错：' + str(e))

    def save_can_info(self, can_info):
        timestamp = datetime.datetime.now()  # 获取当前时间
        can_info = [timestamp] + can_info
        self.worksheet.append(can_info)
        self.workbook.save(self.xlsx_name)  # 保存工作薄

    '''接收线程'''

    def receive_thread_func(self):
        buffer = b''  # 数据缓存
        while self.runing:
            try:
                with self.serial_lock:  # 锁住串口线程，保证线程安全
                    if self.ser and self.ser.in_waiting:
                        buffer += self.ser.read(self.ser.in_waiting)
                while len(buffer) >= 16:
                    if buffer[0] == 0xAA:
                        message = buffer[:16]  # 提取完整信息
                        buffer = buffer[16:]  # 移除已处理部分
                        # 转为16进制字符串
                        hex_data = message.hex()
                        hex_array = [f"{int(x):02x}".upper() for x in message]
                        self.save_can_info(hex_array)
                        self.message_display('接收：' + ' '.join(f"{byte:02X}" for byte in message))
                        self.analysis_hex_data(hex_data)  # 解析数据
                    else:
                        buffer = buffer[1:]  # 丢弃无效数据头部
            except Exception as e:
                if self.runing:  # 仅在线程运行时显示错误
                    self.message_display('接收数据时出错：' + str(e), 'red')
                    self.message_display('================================')
                    messagebox.showerror('Error', '接收数据时出错：' + str(e))
                break

    '''发送线程'''

    def send_thread_func(self):
        while self.runing:
            try:
                if not self.send_queue.empty():
                    send_data = self.send_queue.get()
                    with self.serial_lock:  # 锁住串口线程，保证线程安全
                        self.ser.write(send_data)
                    # 转为16进制数组，CAN信息存储
                    hex_array = [f"{int(x):02x}".upper() for x in send_data]
                    self.save_can_info(hex_array)
                    send_data_str = ' '.join(i for i in hex_array)
                    self.message_display('发送：' + send_data_str)
                    self.send_queue.task_done()  # 任务完成，任务队列出队
            except Exception as e:
                if self.runing:  # 仅在线程运行时显示错误
                    self.message_display('发送数据时出错：' + str(e), 'red')
                    self.message_display('================================')
                    messagebox.showerror('Error', '发送数据时出错：' + str(e))
                break

    '''解析接收到的数据'''

    def analysis_hex_data(self, hex_data):
        if hex_data[:8] == 'aa010008':
            can_id = hex_data[8:16]
            data = hex_data[16:]
            '''根据CANID，解析数据'''
            match can_id:  # 流量数据，定义传过来的数据是计数值，需要进行转换为流量值
                case '00aa0401':
                    # hex_data[16:]中为8路流量数据，要先提取出来,成为8个元素的数组,拿到前五个元素
                    flow_data_array = [int(data[i:i + 2], 16) for i in range(0, len(data), 2)]
                    flow_values = [i / 25 for i in flow_data_array]
                    # 管路流量1L=596脉冲 则流量L/min=每秒脉冲数/596 * 1 * 60
                    flow_values[5] = flow_data_array[5] / 596 * 60
                    # 公式转换为流量 保留2位小数,放入self.flow_volumes进行显示
                    for i in range(len(flow_data_array) - 3):
                        self.flow_volumes[i].set(f'{flow_values[i]:.2f}')
                    self.channel_flow_var.set(f'{flow_values[5]:.2f}')
                case '00aa0501':  # 压力数据 获得值/100=实际值 预留位置
                    pressure_data = int(data[:2], 16) / 100
                    self.pressure_var.set(f'{pressure_data:.2f}')
                case _:  # 其他处理
                    pass

    '''打开/关闭串口'''

    def switch_serial_state(self):
        def button_permissions_open():
            self.port_choose['state'] = 'disabled'
            self.refresh_button['state'] = 'disabled'
            self.baudrate_choose['state'] = 'disabled'
            self.set_freq_button['state'] = 'normal'
            self.set_duty_phase_button['state'] = 'normal'

        def button_permissions_close():
            self.port_choose['state'] = 'normal'
            self.refresh_button['state'] = 'normal'
            self.baudrate_choose['state'] = 'normal'
            self.set_freq_button['state'] = 'disabled'
            self.set_duty_phase_button['state'] = 'disabled'

        if self.serial_button['text'] == '打开':
            try:
                self.ser = serial.Serial(self.port_var.get(), int(self.baudrate_var.get()), timeout=0.5)
                self.serial_button['text'] = '关闭'
                # 按钮权限变更
                button_permissions_open()
                self.message_display('串口打开！串口号：' + self.port_var.get() + ', 波特率：' + self.baudrate_var.get(),
                                     'green')
                # 接收/发送线程打开
                self.runing = True
                # 接收线程打开
                self.receive_thread = threading.Thread(target=self.receive_thread_func, daemon=True)
                self.receive_thread.start()
                self.message_display('接收线程打开！', 'green')
                # 发送线程打开
                self.send_thread = threading.Thread(target=self.send_thread_func, daemon=True)
                self.send_thread.start()
                self.message_display('发送线程打开！', 'green')
                # 启动1S定时器
                self.one_second_timer()
                self.message_display('流量监测定时打开-1S', 'green')
                self.message_display('================================')
            except Exception as e:
                if '拒绝访问' in e.args[0]:
                    self.message_display('端口占用中,请断开端口连接!', 'red')
                    self.message_display('================================')
                    messagebox.showerror('Error', self.port_var.get() + '端口占用中,请断开端口连接')
                else:
                    self.message_display('未知错误：' + str(e), 'red')
                    self.message_display('================================')
                    messagebox.showerror('Error', '未知错误：' + str(e))
                return
        else:
            self.runing = False
            # 等待线程结束
            if self.receive_thread.is_alive():
                self.receive_thread.join(timeout=1)
            self.message_display('接收线程关闭！', 'red')
            if self.send_thread.is_alive():
                self.send_thread.join()
            self.message_display('发送线程关闭！', 'red')
            # 确保串口不被占用时再关闭
            if self.ser.is_open:
                self.ser.close()
            self.message_display('串口关闭！', 'red')
            self.message_display('================================')
            self.serial_button['text'] = '打开'
            # 按钮权限切换
            button_permissions_close()

    '''获取串口列表并显示'''

    def get_port_list(self):
        self.port_list = list(serial.tools.list_ports.comports())  # 获取串口列表
        if len(self.port_list) == 0:  # 不存在串口时，使用默认值
            self.port_list = ['COM1', 'COM2', 'COM3', 'COM4', 'COM5', 'COM6', 'COM7', 'COM8', 'COM9', 'COM10']
            self.port_var.set(self.port_list[8])
            self.message_display('不存在可用串口，使用默认选择项！', 'yellow')
        else:
            self.port_var.set(self.port_list[0].device)
            self.port_list = [self.port_list[i].device for i in range(len(self.port_list))]
            self.message_display('串口加载完成！', 'green')
        menu = self.port_choose['menu']
        menu.delete(0, 'end')
        for port in self.port_list:
            menu.add_command(label=port, command=lambda value=port: self.port_var.set(value))
        self.message_display('================================')

    '''设置频率'''

    def set_frequency(self, freq):
        # 发送十六进制信息，格式为aa010008+4字节CANID+8字节数据，例如AA0100081F0000000000000000000000
        # 将freq str转换为int，然后转换为2字节的16进制，并拆分为高低字节位，用于放入send_data的数组中
        freq = 10 * int(freq)
        self.message_display('设置电磁阀频率为：' + str(freq / 10) + 'Hz', 'green')
        freq_high_low = [(freq >> 8) & 0xFF, freq & 0xFF]
        send_data = [0xAA, 0x01, 0x00, 0x08, 0x1F, 0x00, 0x00, 0x00, freq_high_low[0], freq_high_low[1], 0x00, 0x00,
                     0x00, 0x00, 0x00, 0x00]
        self.send_queue.put(send_data)  # 将发送任务放入队列

    '''设置占空比和相位'''

    def set_duty_phase(self, channel, duty, phase):
        channel = int(channel)
        duty = 10 * int(duty)
        phase = 10 * int(phase)
        self.message_display(
            '设置 ' + str(channel) + ' 号电磁阀：占空比：' + str(duty / 10) + '%，相位：' + str(phase / 10) + '°', 'green')
        duty_high_low = [(duty >> 8) & 0xFF, duty & 0xFF]
        phase_high_low = [(phase >> 8) & 0xFF, phase & 0xFF]
        send_data = [0xAA, 0x01, 0x00, 0x08, 0x1F, 0x00, 0x01, channel, duty_high_low[0], duty_high_low[1],
                     phase_high_low[0], phase_high_low[1], 0x00, 0x00, 0x00, 0x00]
        self.send_queue.put(send_data)  # 将发送任务放入队列

    '''读取脉冲计数值'''

    def read_flow_values(self):
        if not (self.runing and self.auto_read_flow_flag):
            return
        self.one_second_timer()
        send_data = [0xAA, 0x01, 0x00, 0x08, 0x00, 0xAA, 0x03, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00]
        self.send_queue.put(send_data)

    '''1S 定时器'''

    def one_second_timer(self):
        self.timer = threading.Timer(1, self.read_flow_values)
        self.timer.start()

    '''文本框显示'''

    def message_display(self, message, color=None):
        self.message_box.config(state='normal')  # 使文本框可编辑
        if color is None:
            color = 'black'
        self.message_box.tag_config(color, foreground=color)
        self.message_box.insert(END, '★' + str(datetime.datetime.now()) + ': ' + message + '\n', color)  # 插入消息
        self.message_box.see(END)  # 使文本框始终在最下面
        self.message_box.config(state='disabled')  # 禁止编辑

    '''切换自动发送流量flag'''

    def toggle_auto_read_flow(self):
        self.auto_read_flow_flag = not self.auto_read_flow_flag
        if self.auto_read_flow_flag:
            self.message_display('主动发送流量读取指令已开启！', 'green')
        else:
            self.message_display('主动发送流量读取指令已关闭！', 'red')
        self.read_flow_values()

    '''创建组件'''

    def createWidget(self):

        '''框架区'''
        # 标题框架
        title_f = Frame(root, width=self.window_width, borderwidth=1, relief='solid', padx=30, pady=10)
        title_f.pack()
        # 配置区框架
        config_f = Frame(root, width=self.window_width, padx=30)
        self.config_serial_f = Frame(config_f)
        config_f.pack()
        self.config_serial_f.pack(pady=5, padx=30, side='bottom', anchor='w')
        # 传感器区框架
        sensor_data_f = Frame(root, width=self.window_width, padx=30)
        data_labels_f = Frame(sensor_data_f)
        sensor_data_f.pack()
        data_labels_f.pack(pady=5, padx=30, side='bottom', anchor='w')
        # 指令区框架
        order_f = Frame(root, width=self.window_width, padx=30)
        order_freq_f = Frame(order_f)
        order_duty_phase_f = Frame(order_f)
        order_f.pack()
        order_duty_phase_f.pack(pady=5, padx=30, side='bottom', anchor='w')
        order_freq_f.pack(pady=5, padx=30, side='bottom', anchor='w')
        # 信息区框架
        message_f = Frame(root, width=self.window_width, padx=30)
        message_f.pack()
        # 状态区框架
        bottom_f = Frame(root, width=self.window_width, padx=30)
        bottom_f.pack(side='bottom', anchor='w')

        '''内容区'''
        ########### 标题 ###########
        self.title_label = Label(title_f, text='喷雾数据在线监控系统', font=('黑体', 25))
        self.title_label.pack()

        ########### 配置区，用于配置串口 ###########
        # 分割线
        self.config_label = Label(config_f,
                                  text='配置---------------------------------------------------------------------------',
                                  font=('黑体', 15))
        self.config_label.pack()
        # 创建串口选择框、波特率选择框、打开/关闭按钮以及刷新按钮
        # 串口选择框
        self.label_port = Label(self.config_serial_f, text='串口：', font=('黑体', 15))
        self.label_port.pack(side='left')
        self.port_var = StringVar()
        self.port_choose = OptionMenu(self.config_serial_f, self.port_var, *self.port_list)
        self.port_choose.pack(side='left')
        Label(self.config_serial_f, text='').pack(side='left', padx=10)
        # 波特率选择框
        self.label_baudrate = Label(self.config_serial_f, text='波特率：', font=('黑体', 15))
        self.label_baudrate.pack(side='left')
        self.baudrate_var = StringVar()
        self.baudrate_var.set(self.baudrate_list[0])
        self.baudrate_choose = OptionMenu(self.config_serial_f, self.baudrate_var, *self.baudrate_list)
        self.baudrate_choose.pack(side='left')
        # 串口开关按钮
        self.serial_button = Button(self.config_serial_f, text='打开', font=('黑体', 13),
                                    command=self.switch_serial_state)
        self.serial_button.pack(side='left', padx=30)
        # 刷新串口按钮
        self.refresh_button = Button(self.config_serial_f, text='刷新串口', font=('黑体', 13),
                                     command=self.get_port_list)
        self.refresh_button.pack(side='left')

        ########### 传感器数据区，观测流量、压力数据 ###########
        # 分割线
        self.flow_data_label = Label(sensor_data_f,
                                     text='传感器数据---------------------------------------------------------------------',
                                     font=('黑体', 15))
        self.flow_data_label.pack()
        # 创建2行3列 5个喷头流量，1个管路流量，1个管路压力
        # 流量5路
        self.flow_labels = [Label(data_labels_f, font=('黑体', 15), text='流量{}:'.format(i + 1)) for i in range(5)]
        self.flow_volumes = [StringVar() for _ in range(5)]
        self.flow_volume_labels = [Label(data_labels_f, font=('黑体', 15), width=5, bg='white', relief='solid',
                                         textvariable=self.flow_volumes[i]) for i in range(5)]
        self.flow_units = [Label(data_labels_f, font=('黑体', 15), text='L/min') for _ in range(5)]
        # 管路流量
        self.channel_flow_label = Label(data_labels_f, font=('黑体', 15), text='管路流量:')
        self.channel_flow_var = StringVar()
        self.channel_flow_data_label = Label(data_labels_f, font=('黑体', 15), textvariable=self.channel_flow_var,
                                             bg='white', relief='solid', width=5)
        self.channel_flow_unit = Label(data_labels_f, font=('黑体', 15), text='L/min')
        # 管路压力
        self.pressure_label = Label(data_labels_f, font=('黑体', 15), text='管路压力:')
        self.pressure_var = StringVar()
        self.pressure_data_label = Label(data_labels_f, font=('黑体', 15), textvariable=self.pressure_var, bg='white',
                                         relief='solid', width=5)
        self.pressure_unit = Label(data_labels_f, font=('黑体', 15), text='MPa  ')
        # 网格配置
        for i in range(5):
            self.flow_labels[i].grid(row=i // 3, column=i % 3 * 3)
            self.flow_volume_labels[i].grid(row=i // 3, column=i % 3 * 3 + 1)
            self.flow_units[i].grid(row=i // 3, column=i % 3 * 3 + 2, padx=3, pady=5)
            self.flow_volumes[i].set('0.00')
        self.channel_flow_label.grid(row=2, column=0)
        self.channel_flow_data_label.grid(row=2, column=1)
        self.channel_flow_var.set('0.00')
        self.channel_flow_unit.grid(row=2, column=2, pady=5)
        self.pressure_label.grid(row=2, column=3)
        self.pressure_data_label.grid(row=2, column=4)
        self.pressure_var.set('0.00')
        self.pressure_unit.grid(row=2, column=5, pady=5)

        ########### 指令发送区，差相驱动板指令发送 ###########
        # 分割线
        self.order_divider_label = Label(order_f,
                                         text='指令发送区---------------------------------------------------------------------',
                                         font=('黑体', 15))
        self.order_divider_label.pack()
        # 频率
        self.freq_label = Label(order_freq_f, text='频率:', font=('黑体', 15))
        self.freq_var = StringVar()
        self.freq_entry = Entry(order_freq_f, font=('黑体', 15), textvariable=self.freq_var, bg='white', relief='solid',
                                width=5, justify='center')
        self.freq_unit = Label(order_freq_f, text='Hz', font=('黑体', 15))
        self.freq_label.pack(side='left')
        self.freq_entry.pack(side='left')
        self.freq_unit.pack(side='left', padx=5)
        self.freq_var.set('1')
        # 设置频率按钮
        self.set_freq_button = Button(order_freq_f, state='disabled', text='频率设置', font=('黑体', 13),
                                      command=lambda: self.set_frequency(self.freq_entry.get()))
        self.set_freq_button.pack(side='left', padx=10)
        # 自动发送流量请求复选框
        self.auto_send_flow_var = IntVar()
        self.auto_send_flow_checkbutton = Checkbutton(order_freq_f, font=('黑体', 15), variable=self.auto_send_flow_var,
                                                      text='主动发送流量读取指令', onvalue=1, offvalue=0,
                                                      command=self.toggle_auto_read_flow)
        self.auto_send_flow_checkbutton.pack(side='left', padx=5)
        self.auto_send_flow_var.set(1)
        # 通道号
        self.duty_phase_channel_label = Label(order_duty_phase_f, text='通道:', font=('黑体', 15))
        self.duty_phase_channel_var = StringVar()
        self.duty_phase_channel_entry = Entry(order_duty_phase_f, font=('黑体', 15),
                                              textvariable=self.duty_phase_channel_var, bg='white', relief='solid',
                                              width=5, justify='center')
        self.duty_phase_channel_unit = Label(order_duty_phase_f, text='号', font=('黑体', 15))
        self.duty_phase_channel_label.pack(side='left')
        self.duty_phase_channel_entry.pack(side='left')
        self.duty_phase_channel_unit.pack(side='left', padx=5)
        self.duty_phase_channel_var.set('1')
        # 占空比
        self.duty_label = Label(order_duty_phase_f, text='占空比:', font=('黑体', 15))
        self.duty_var = StringVar()
        self.duty_entry = Entry(order_duty_phase_f, font=('黑体', 15), textvariable=self.duty_var, bg='white',
                                relief='solid', width=5, justify='center')
        self.duty_unit = Label(order_duty_phase_f, text='%', font=('黑体', 15))
        self.duty_label.pack(side='left')
        self.duty_entry.pack(side='left')
        self.duty_unit.pack(side='left', padx=5)
        self.duty_var.set('50')
        # 相位
        self.phase_label = Label(order_duty_phase_f, text='相位:', font=('黑体', 15))
        self.phase_var = StringVar()
        self.phase_entry = Entry(order_duty_phase_f, font=('黑体', 15), textvariable=self.phase_var, bg='white',
                                 relief='solid', width=5, justify='center')
        self.phase_unit = Label(order_duty_phase_f, text='°', font=('黑体', 15))
        self.phase_label.pack(side='left')
        self.phase_entry.pack(side='left')
        self.phase_unit.pack(side='left', padx=5)
        self.phase_var.set('0')
        # 设置占空比和相位
        self.set_duty_phase_button = Button(order_duty_phase_f, state='disabled', text='占空比和相位设置',
                                            font=('黑体', 13),
                                            command=lambda: self.set_duty_phase(self.duty_phase_channel_var.get(),
                                                                                self.duty_var.get(),
                                                                                self.phase_var.get()))
        self.set_duty_phase_button.pack(side='left', padx=10)

        ########### 系统状态信息区 ###########
        # 分割线
        self.message_divider_label = Label(message_f,
                                           text='系统状态信息-------------------------------------------------------------------',
                                           font=('黑体', 15))
        self.message_divider_label.pack()
        # 信息框
        self.message_box = ScrolledText(message_f, width=100, height=8, font=('宋体', 10))
        self.message_box.pack(side='left', padx=30, pady=5)
        self.message_box.config(state='disabled')  # 禁止编辑

        ########### 状态栏 ############
        self.version_label = Label(bottom_f, text='Version:', font=('黑体', 15))
        self.version_label.pack(side='left')
        self.version_currently = Label(bottom_f, text=self.version, font=('黑体', 15))
        self.version_currently.pack(side='left')

        ########### 显示信息、数据处理 ###########
        self.message_display('界面加载完成！', 'green')
        self.message_display('================================')
        self.get_port_list()


if __name__ == '__main__':
    root = Tk()
    root.geometry('900x600')  # 窗口大小
    root.title('果树生物量在线采集软件')
    app = Concert(root)
    root.mainloop()
