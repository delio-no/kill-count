import customtkinter as ctk
import threading
import time
import psutil
import win32process
import win32gui
import win32api
import win32con
import elevate
import keyboard
from CTkMessagebox import CTkMessagebox

from keyboard_code import VK_CODE

ctk.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"


class TimerApp:
    REMOTE_IP = "46.4.34.206"
    REMOTE_PORT = 7779

    def __init__(self, root):
        self.root = root
        self.root.title("Timer App")
        self.root.geometry("400x650")

        self.process_id = None
        self.hwnd = None
        self.connection = None
        self.filter_rule = None

        self.count_kill = 0  # default timer duration is 5 minutes
        self.kill = 0  # default timer duration is 5 minutes
        self.is_run = False
        self.end_time = None
        self.send_key_on_reset = ctk.BooleanVar(value=False)  # Checkbox variable

        self.key_thread = None  # Переменная для управления потоком
        self.key_thread_running = False  # Флаг для управления потоком

        self.create_widgets()
        self.setup_hotkeys()

    def create_widgets(self):
        self.status_label_frame = ctk.CTkFrame(self.root)
        self.status_label_frame.pack(pady=5)

        self.status_label = ctk.CTkLabel(self.status_label_frame, text="Статус окна EW: Не захвачено")
        self.status_label.pack(side="left", padx=(0, 10))

        self.status_indicator = ctk.CTkFrame(self.status_label_frame, width=20, height=20)
        self.status_indicator.pack(side="left")

        self.kill_label = ctk.CTkLabel(self.root, text="Убито: 0")
        self.kill_label.pack(pady=5)

        self.count_kill_label = ctk.CTkLabel(self.root, text=f"Количество убитых: {self.kill}")
        self.count_kill_label.pack(pady=5)

        self.start_button = ctk.CTkButton(self.root, text="Старт", command=self.start)
        self.start_button.pack(pady=5)

        self.stop_button = ctk.CTkButton(self.root, text="Стоп", command=self.stop)
        self.stop_button.pack(pady=5)

        self.set_kill_count_button = ctk.CTkButton(self.root, text="Количество мобов", command=self.set_count_kill)
        self.set_kill_count_button.pack(pady=5)

        self.capture_window_button = ctk.CTkButton(self.root, text="Захватить окно",
                                                   command=self.find_and_capture_window)
        self.capture_window_button.pack(pady=5)

        self.description_label = ctk.CTkLabel(self.root, text="Описание горячих клавиш:",
                                              font=("Arial", 12, "bold"))
        self.description_label.pack(pady=5)

        self.description_text = ctk.CTkTextbox(self.root, width=380, height=125)
        self.description_text.pack(pady=5)
        self.description_text.insert(ctk.END,
                                     "1. При нажатии 'ctrl+q' - захватывает HWID (PID) и HWND текущего активного окна.\n")
        self.description_text.insert(ctk.END,
                                     "2. При нажатии 'ctrl+w' - отправляет клавишу 'z' в текущий активный процесс и запускает таймер.\n")
        self.description_text.insert(ctk.END, "3. При нажатии 'ctrl+e' - сбрасывает таймер и отправляет клавишу 'z'.\n")
        self.description_text.configure(state="disabled")

        self.description_label = ctk.CTkLabel(self.root, text="Консоль:",
                                              font=("Arial", 12, "bold"))
        self.description_label.pack(pady=5)
        self.console = ctk.CTkTextbox(self.root, width=380, height=120)
        self.console.pack(pady=5)

    def setup_hotkeys(self):
        keyboard.add_hotkey('ctrl+q', self.capture_window)
        keyboard.add_hotkey('ctrl+w', self.start_hot_key)
        keyboard.add_hotkey('ctrl+e', self.stop)

    def get_active_window_pid(self):
        hwnd = win32gui.GetForegroundWindow()
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        process = psutil.Process(pid)
        if process.name() == "EndlessWar.exe":
            return pid, hwnd
        else:
            return None, None

    def timer_function(self):
        self.log('Таймер завершен')
        self.is_run = False

    def capture_window(self):
        self.process_id, self.hwnd = self.get_active_window_pid()
        self.connection = self.get_socket_connection()
        if self.hwnd and self.connection:
            self.set_local_ip()
            self.set_local_port()
            self.filter_rule = f"inbound and tcp.DstPort == {self.local_port} and ip.DstAddr == {self.local_ip}"

            self.status_label.configure(text="Статус окна EW: Захвачено")
            self.status_indicator.configure(fg_color="green")
            self.log('Окно захвачено')
        else:
            self.status_label.configure(text="Статус окна EW: Не захвачено")
            self.status_indicator.configure(fg_color="red")
            self.log('Окно не является EndlessWar.exe')

    def get_process_connections(self):
        """Получает соединения процесса по его PID."""
        if self.process_id:
            try:
                process = psutil.Process(self.process_id)
                return process.connections()
            except psutil.NoSuchProcess:
                self.log("Процесс не найден.")
                return None
        else:
            self.log("PID не установлен.")
            return None

    def find_connection(self):
        connections = self.get_process_connections()
        if connections:
            for conn in connections:
                if conn.raddr and conn.raddr.ip == self.REMOTE_IP and conn.raddr.port == self.REMOTE_PORT:
                    return conn
            self.log("Connect not found.")
            print("Connect not found.")
            return None
        else:
            return None

    def get_socket_connection(self):
        connection = self.find_connection()
        if connection:
            self.log(f"Connection: {connection}")
            self.log(f"Local IP: {connection.laddr.ip}")
            self.log(f"Local port: {connection.laddr.port}")
            print(f"Connection: {connection}")
            print(f"Local IP: {connection.laddr.ip}")
            print(f"Local port: {connection.laddr.port}")
            return connection
        else:
            self.log("Connect not found.")
            return None

    def find_and_capture_window(self):
        processes = [proc for proc in psutil.process_iter(['pid', 'name']) if proc.info['name'] == "EndlessWar.exe"]
        if len(processes) > 1:
            CTkMessagebox(title="Ошибка",
                          message="Найдено более одного процесса EndlessWar.exe. Используйте горячую клавишу 'ctrl+q' для захвата окна.")
            return
        elif len(processes) == 1:
            self.process_id = processes[0].info['pid']
            self.hwnd = self.get_hwnd_from_pid(self.process_id)
            if self.hwnd:
                self.status_label.configure(text="Статус окна EW: Захвачено")
                self.status_indicator.configure(fg_color="green")
                self.log('Окно EndlessWar.exe захвачено')
                print(self.hwnd)
                return
        self.status_label.configure(text="Статус окна EW: Не захвачено")
        self.status_indicator.configure(fg_color="red")
        self.log('Окно EndlessWar.exe не найдено')

    def get_hwnd_from_pid(self, pid):
        def callback(hwnd, hwnds):
            _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
            if found_pid == pid:
                hwnds.append(hwnd)
            return True

        hwnds = []
        win32gui.EnumWindows(callback, hwnds)
        return hwnds[0] if hwnds else None

    def set_local_ip(self):
        """Устанавливает локальный IP из соединения."""
        if self.connection:
            self.local_ip = self.connection.laddr.ip

    def set_local_port(self):
        """Устанавливает локальный порт из соединения."""
        if self.connection:
            self.local_port = self.connection.laddr.port

    def set_count_kill(self):
        try:
            self.count_kill = int(ctk.CTkInputDialog(title="Установить кол-во мобов",
                                                     text="Введите количество тмобов:").get_input())
            if self.count_kill is not None:
                self.count_kill_label.configure(text=f"Текущая количество мобов: {self.count_kill}")
                self.log(f"Таймер установлен на {self.count_kill} секунд")
            else:
                CTkMessagebox(title="Ошибка", message="Ошибка: введено не целое число.")
        except ValueError:
            CTkMessagebox(title="Ошибка", message="Ошибка: введено не целое число.")
        except TypeError:
            CTkMessagebox(title="Ошибка", message="Ошибка: введено не целое число.")

    def start(self):
        if not self.key_thread_running and self.hwnd:
            self.key_thread_running = True
            self.key_thread = threading.Thread(target=self.send_key_loop)
            self.key_thread.start()
            self.log("Нажатие клавиши '1' запущено")

    def start_hot_key(self):
        if self.hwnd:
            self.start()

    def stop(self):
        if self.key_thread_running:
            self.key_thread_running = False
            self.key_thread.join()  # Дождитесь завершения потока
            self.log("Нажатие клавиши '1' остановлено")

    def send_key_loop(self):
        while self.key_thread_running:
            self.send_key_1()
            time.sleep(0.1)  # Задержка между нажатиями

    def log(self, message):
        self.console.insert(ctk.END, message + "\n")
        self.console.see(ctk.END)

    def send_key_1(self):
        print('Use key - 1')
        try:
            win32api.PostMessage(self.hwnd, win32con.WM_KEYDOWN, int(VK_CODE['1']), 0)
            win32api.PostMessage(self.hwnd, win32con.WM_KEYUP, int(VK_CODE['1']), 0)
        except Exception as e:
            print('Ошибка передачи команды в окно')


elevate.elevate()

root = ctk.CTk()
app = TimerApp(root)

root.mainloop()