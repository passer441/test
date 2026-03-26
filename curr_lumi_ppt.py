import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import csv
import re
import time
from datetime import datetime
import serial  
import sys

try:
    import win32com.client
    import pythoncom
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

class RealDevice:
    def __init__(self):
        self.current_gray = 0
        self.keithley_inst = None 
        self.ca_app = None        
        self.ca_obj = None        
        self.ca_probe = None      
        self.ca_memory = None     

    def connect_keithley(self, port, baud, curr_limit, curr_range):
        try:
            self.keithley_inst = serial.Serial(port, int(baud), timeout=1)
            self.keithley_inst.write(b"*RST\n")
            time.sleep(0.1)
            
            self.keithley_inst.write(b":SOUR:FUNC VOLT\n")
            self.keithley_inst.write(b":SOUR:VOLT 0\n")
            self.keithley_inst.write(f":SENS:CURR:PROT {curr_limit}\n".encode('utf-8'))
            
            self.keithley_inst.write(b":SENS:FUNC 'CURR'\n")
            
            if curr_range.lower() == "auto":
                self.keithley_inst.write(b":SENS:CURR:RANG:AUTO ON\n")
            else:
                range_map = {"10uA": "10e-6", "100uA": "100e-6", "1mA": "1e-3", "10mA": "10e-3", "100mA": "100e-3", "1A": "1"}
                r_val = range_map.get(curr_range, "100e-3")
                self.keithley_inst.write(b":SENS:CURR:RANG:AUTO OFF\n")
                self.keithley_inst.write(f":SENS:CURR:RANG {r_val}\n".encode('utf-8'))
            
            self.keithley_inst.write(b":OUTP ON\n")
            
            return True, f"[{port}] 실제 Keithley 2400 연결 성공 (Range: {curr_range})"
        except Exception as e:
            return False, f"Keithley 연결 실패: {e}"

    def change_range(self, curr_range):
        if not self.keithley_inst:
            return False, "장비가 연결되어 있지 않습니다."
        try:
            if curr_range.lower() == "auto":
                self.keithley_inst.write(b":SENS:CURR:RANG:AUTO ON\n")
            else:
                range_map = {"10uA": "10e-6", "100uA": "100e-6", "1mA": "1e-3", "10mA": "10e-3", "100mA": "100e-3", "1A": "1"}
                r_val = range_map.get(curr_range, "100e-3")
                self.keithley_inst.write(b":SENS:CURR:RANG:AUTO OFF\n")
                self.keithley_inst.write(f":SENS:CURR:RANG {r_val}\n".encode('utf-8'))
            return True, f"측정 Range가 {curr_range}로 실시간 변경되었습니다."
        except Exception as e:
            return False, f"Range 변경 실패: {e}"

    def connect_ca310(self, sync_mode, channel, display_mode):
        try:
            if sys.maxsize > 2**32:
                return False, "64비트 파이썬 환경입니다. CA-SDK는 32비트 환경에서만 동작하므로 연결할 수 없습니다."
                
            if not WIN32_AVAILABLE:
                return False, "win32com 모듈을 사용할 수 없습니다."

            self.ca_app = win32com.client.Dispatch("CA200Srvr.Ca200")
            self.ca_app.AutoConnect()
            
            self.ca_obj = self.ca_app.SingleCa        
            self.ca_probe = self.ca_obj.SingleProbe   
            self.ca_memory = self.ca_obj.Memory       
            
            sync_dict = {"Universal": 3, "NTSC": 0, "PAL": 1, "Internal": 2}
            self.ca_obj.SyncMode = sync_dict.get(sync_mode, 3)
            self.ca_memory.ChannelNO = int(channel)
            self.ca_obj.DisplayMode = int(display_mode)
            
            return True, "[USB] 실제 CA-310 연결 성공"
        except Exception as e:
            return False, f"CA-310 연결 실패: {e}"

    def perform_zero_cal(self):
        try:
            if self.ca_obj:
                self.ca_obj.CalZero()
                return True, "실제 CA-310 Zero Calibration 완료"
            return False, "CA-310 장비 객체가 초기화되지 않았습니다."
        except Exception as e:
            return False, f"Zero Cal 실패: {e}"

    def get_keithley_data(self):
        try:
            if self.keithley_inst:
                self.keithley_inst.write(b":READ?\n")
                data = self.keithley_inst.readline().decode('utf-8').strip()
                if data:
                    values = data.split(',')
                    if len(values) >= 2:
                        try:
                            val = float(values[1])
                            if val > 1e30: 
                                return float('inf')
                            return val
                        except ValueError:
                            return 0.0
            return 0.0
        except Exception as e:
            print(f"Keithley 데이터 읽기 실패: {e}")
            return 0.0

    def get_ca310_data(self):
        try:
            if self.ca_obj and self.ca_probe:
                self.ca_obj.Measure()
                mode = self.ca_obj.DisplayMode
                
                if mode == 7: 
                    lv = self.ca_probe.Y
                    sx = self.ca_probe.X
                    sy = self.ca_probe.Z
                elif mode == 1: 
                    lv = self.ca_probe.Lv
                    sx = self.ca_probe.T
                    sy = self.ca_probe.duv
                elif mode == 5: 
                    lv = self.ca_probe.Lv
                    sx = self.ca_probe.ud
                    sy = self.ca_probe.vd
                else: 
                    lv = self.ca_probe.Lv
                    try:
                        sx = self.ca_probe.sx
                        sy = self.ca_probe.sy
                    except AttributeError:
                        sx = self.ca_probe.x
                        sy = self.ca_probe.y
                
                return lv, sx, sy
            return 0.0, 0.0, 0.0
        except Exception as e:
            print(f"CA-310 측정 실패: {e}")
            return 0.0, 0.0, 0.0
            
    def release_devices(self):
        try:
            if self.keithley_inst:
                self.keithley_inst.write(b":OUTP OFF\n")
                self.keithley_inst.close()
                self.keithley_inst = None
                
            self.ca_obj = None
            self.ca_probe = None
            self.ca_memory = None
            self.ca_app = None
        except Exception as e:
            print(f"장비 연결 해제 중 오류 발생: {e}")


class OLEDMeasurementApp:
    def __init__(self, root):
        self.root = root
        self.root.title("OLED Measurement System (Virtual Test Mode)")
        self.root.geometry("1150x700")
        
        self.device = RealDevice()
        
        self.k_port = tk.StringVar(value="COM3")
        self.k_baud = tk.StringVar(value="9600")
        self.k_volt = tk.StringVar(value="10V")
        self.k_curr_limit = tk.StringVar(value="1.05")
        
        self.k_curr_range = tk.StringVar(value="100mA") 
        self.k_range_desc = tk.StringVar(value="최대 측정 범위: 100mA") 
        self.k_curr_range.trace_add("write", self.update_range_desc)
        
        self.ca_sync = tk.StringVar(value="Universal")
        self.ca_mem = tk.IntVar(value=1)
        self.ca_disp_mode = tk.StringVar(value="Lvxy") 
        
        self.slides_dict = {}
        self.current_slide_idx = 1
        self.measure_results = []
        self.slideshow_started = False
        self.is_black_screen = False
        
        self.k_connected = False
        self.ca_connected = False
        self.ca_zero_calibrated = False
        self.ppt_app = None
        self.presentation = None
        
        self.loc_var = tk.StringVar(value="대기 중")
        self.tune_target_curr = tk.StringVar(value="N/A")
        self.tune_status = tk.StringVar(value="1번 슬라이더 대기 중")
        self.meas_status = tk.StringVar(value="1번 슬라이더 대기 중")  
        self.slide_num_var = tk.StringVar(value="- / -")
        self.current_gray_var = tk.StringVar(value="확인 대기")
        
        self.setup_ui()
        self.root.protocol("WM_DELETE_WINDOW", self.close_app)

    def wait(self, seconds):
        var = tk.IntVar()
        self.root.after(int(seconds * 1000), var.set, 1)
        self.root.wait_variable(var)

    def update_range_desc(self, *args):
        self.k_range_desc.set(f"최대 측정 범위: {self.k_curr_range.get()}")

    def setup_ui(self):
        main_pane = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_pane.pack(expand=True, fill="both", padx=10, pady=10)

        self.tabs = ttk.Notebook(main_pane)
        main_pane.add(self.tabs, weight=7)
        self.tabs.bind("<<NotebookTabChanged>>", self.on_tab_changed)

        self.tab1 = ttk.Frame(self.tabs); self.tabs.add(self.tab1, text="1. Keithley 설정")
        self.tab2 = ttk.Frame(self.tabs); self.tabs.add(self.tab2, text="2. CA-310 설정")
        self.tab3 = ttk.Frame(self.tabs); self.tabs.add(self.tab3, text="3. PPT 동기화")
        self.tab4 = ttk.Frame(self.tabs); self.tabs.add(self.tab4, text="4. Gray 튜닝")
        self.tab5 = ttk.Frame(self.tabs); self.tabs.add(self.tab5, text="5. 측정 실행")

        self.build_keithley_tab()
        self.build_ca310_tab()
        self.build_ppt_tab()
        self.build_tuning_tab()
        self.build_measure_tab()

        log_frame = ttk.LabelFrame(main_pane, text=" 시스템 로그 ", padding=5)
        main_pane.add(log_frame, weight=3)
        self.log_text = tk.Text(log_frame, width=35, state="normal", bg="#f0f0f0", font=("Consolas", 9))
        self.log_text.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        sb.pack(side="right", fill="y")
        self.log_text.config(yscrollcommand=sb.set)
        
        self.log_message("프로그램이 시작되었습니다.")

    def log_message(self, msg):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert("end", f"[{timestamp}] {msg}\n")
        self.log_text.see("end")

    def build_keithley_tab(self):
        frame = ttk.Frame(self.tab1, padding=20)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Keithley 2400 설정", font=("Malgun Gothic", 16, "bold")).pack(anchor="w", pady=10)
        
        guide = ttk.LabelFrame(frame, text=" 하드웨어 연결 안내 ", padding=10)
        guide.pack(fill="x", pady=5)
        ttk.Label(guide, text="1. 장비 [MENU] -> COMMUNICATION -> RS-232 선택\n2. Baud Rate를 9600으로 맞춤\n3. 장치 관리자에서 COM 포트 번호 확인 후 아래 입력").pack(anchor="w")

        form = ttk.LabelFrame(frame, text=" 통 세부 설정 ", padding=15)
        form.pack(fill="x", pady=10)
        ttk.Label(form, text="COM 포트:").grid(row=0, column=0, sticky="w", pady=5)
        ttk.Entry(form, textvariable=self.k_port).grid(row=0, column=1, padx=10, sticky="w")
        ttk.Label(form, text="Baud Rate:").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Entry(form, textvariable=self.k_baud).grid(row=1, column=1, padx=10, sticky="w")
        ttk.Label(form, text="전류 리미트(A):").grid(row=2, column=0, sticky="w", pady=5)
        ttk.Entry(form, textvariable=self.k_curr_limit).grid(row=2, column=1, padx=10, sticky="w")
        ttk.Label(form, text="측정 Range:").grid(row=3, column=0, sticky="w", pady=5)
        ttk.Combobox(form, textvariable=self.k_curr_range, values=["Auto", "10uA", "100uA", "1mA", "10mA", "100mA", "1A"]).grid(row=3, column=1, padx=10, sticky="w")
        ttk.Label(form, textvariable=self.k_range_desc, foreground="blue").grid(row=3, column=2, sticky="w", padx=10)
        ttk.Button(form, text="Range 실시간 적용", command=self.apply_range_action).grid(row=3, column=3, padx=10, sticky="w")
        
        ttk.Button(form, text="장비 연결 실행", command=self.connect_keithley_action).grid(row=4, column=0, columnspan=2, pady=15)

    def connect_keithley_action(self):
        self.k_connected, details = self.device.connect_keithley(self.k_port.get(), self.k_baud.get(), self.k_curr_limit.get(), self.k_curr_range.get())
        self.log_message(details)

    def apply_range_action(self):
        if not self.k_connected:
            messagebox.showwarning("경고", "먼저 장비를 연결해 주세요.")
            return
        success, msg = self.device.change_range(self.k_curr_range.get())
        if success:
            self.log_message(msg)
            messagebox.showinfo("성공", msg)
        else:
            self.log_message(msg)
            messagebox.showerror("오류", msg)

    def build_ca310_tab(self):
        frame = ttk.Frame(self.tab2, padding=20)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="CA-310 설정", font=("Malgun Gothic", 16, "bold")).pack(anchor="w", pady=10)
        form = ttk.LabelFrame(frame, text=" 측정 옵션 ", padding=15)
        form.pack(fill="x", pady=10)
        ttk.Label(form, text="Sync Mode:").grid(row=0, column=0, sticky="w", pady=5)
        ttk.Combobox(form, textvariable=self.ca_sync, values=["Universal", "NTSC", "PAL", "Internal"]).grid(row=0, column=1, padx=10)
        ttk.Label(form, text="Memory CH:").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Spinbox(form, textvariable=self.ca_mem, from_=0, to=99).grid(row=1, column=1, padx=10)
        ttk.Label(form, text="Display Mode:").grid(row=2, column=0, sticky="w", pady=5)
        ttk.Combobox(form, textvariable=self.ca_disp_mode, values=["Lvxy", "XYZ", "Tduv", "u'v'"]).grid(row=2, column=1, padx=10)
        
        btn_frame = ttk.Frame(form)
        btn_frame.grid(row=3, column=0, columnspan=2, pady=15)
        ttk.Button(btn_frame, text="장비 연결", command=self.connect_ca310_action).pack(side="left", padx=5)
   
    def connect_ca310_action(self):
        disp_dict = {"Lvxy": 0, "XYZ": 7, "Tduv": 1, "u'v'": 5}
        mode_val = disp_dict.get(self.ca_disp_mode.get(), 0)
        
        self.ca_connected, details = self.device.connect_ca310(self.ca_sync.get(), self.ca_mem.get(), mode_val)
        self.log_message(details)
        if self.ca_connected:
            mode = self.ca_disp_mode.get()
            if mode == "XYZ":
                self.measure_tree.heading("lv", text="Y")
                self.measure_tree.heading("cx", text="X")
                self.measure_tree.heading("cy", text="Z")
            elif mode == "Tduv":
                self.measure_tree.heading("lv", text="휘도(nit)")
                self.measure_tree.heading("cx", text="T")
                self.measure_tree.heading("cy", text="duv")
            elif mode == "u'v'":
                self.measure_tree.heading("lv", text="휘도(nit)")
                self.measure_tree.heading("cx", text="u'")
                self.measure_tree.heading("cy", text="v'")
            else:
                self.measure_tree.heading("lv", text="휘도(nit)")
                self.measure_tree.heading("cx", text="sx")
                self.measure_tree.heading("cy", text="sy")
                
            messagebox.showinfo("필수 작업", "CA-310 장비가 연결되었습니다. 반드시 Zero Cal을 실행합니다.")
            self.run_zero_cal_action()

    def run_zero_cal_action(self):
        if not self.ca_connected: return
        success, details = self.device.perform_zero_cal()
        if success:
            self.ca_zero_calibrated = True
            self.log_message(details)

    def build_ppt_tab(self):
        frame = ttk.Frame(self.tab3, padding=20)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="파워포인트 동기화", font=("Malgun Gothic", 16, "bold")).pack(pady=10)
        
        warning_frame = ttk.LabelFrame(frame, text=" 필수 주의사항 ", padding=10)
        warning_frame.pack(fill="x", pady=10)
        
        warn_text = (
            "1. 동기화 작업은 반드시 파워포인트 '일반 편집 화면' 상태에서 실행해야 합니다.\n"
            "2. 동기화 완료 후 파이썬으로 장비를 제어하려면 반드시 '슬라이드 쇼(F5)' 상태여야 합니다."
        )
        ttk.Label(warning_frame, text=warn_text, font=("Malgun Gothic", 12, "bold"), foreground="red", justify="center").pack(pady=10)

        sync_frame = ttk.Frame(frame)
        sync_frame.pack(pady=10)
        ttk.Button(sync_frame, text="PPT 동기화 실행", command=self.resync_ppt_action, width=20).pack()
        
        guide_text = (
            "첫 번째 슬라이드에는 offset 전류를 측정하도록 블랙 슬라이드를 추가해 주세요!\n"
            "슬라이드 노트에는 나중에 측정 데이터의 측정 위치를 식별할 수 있는 Label을 적어 주세요!\n"
            "그리고 각 슬라이드의 Target 전류는 슬라이드 노트에 ( 0.113 mA ) 와 같이 표기해 주세요!\n"
            "슬라이드 노트 작성 예) 양산 ( 0.231 mA )"
        )
        self.info_lbl = ttk.Label(frame, text=guide_text, justify="left", foreground="blue")
        self.info_lbl.pack(pady=20)

    def resync_ppt_action(self):
        try:
            self.ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
            
            was_in_slideshow = False
            saved_idx = 1
            try:
                if self.ppt_app.SlideShowWindows.Count > 0:
                    was_in_slideshow = True
                    saved_idx = self.ppt_app.SlideShowWindows(1).View.Slide.SlideIndex
                    self.ppt_app.SlideShowWindows(1).View.Exit()
                    time.sleep(0.5)
            except Exception as e:
                self.log_message(f"슬라이드 쇼 확인 중 무시된 예외: {e}")
                
            self.presentation = self.ppt_app.ActivePresentation
            
            slide_count = self.presentation.Slides.Count
            self.slides_dict = {}
            for i in range(1, slide_count + 1):
                slide = self.presentation.Slides(i)
                note = ""
                try:
                    for shape in slide.NotesPage.Shapes:
                        if shape.HasTextFrame and shape.TextFrame.HasText:
                            t = shape.TextFrame.TextRange.Text.strip()
                            if len(t) > len(note): note = t
                except Exception as e: 
                    self.log_message(f"슬라이드 노트 읽기 중 무시된 예외: {e}")
                self.slides_dict[i] = note if note else f"Slide {i}"
            
            if was_in_slideshow:
                self.presentation.SlideShowSettings.Run()
                time.sleep(0.5)
                try:
                    self.presentation.SlideShowWindow.View.GotoSlide(saved_idx)
                    self.current_slide_idx = self.presentation.SlideShowWindow.View.Slide.SlideIndex
                except:
                    self.current_slide_idx = 1
                self.slideshow_started = True
            else:
                self.current_slide_idx = 1
                self.slideshow_started = False
                
            self.update_tuning_info()
            self.tune_status.set(f"{self.current_slide_idx}번 슬라이더 대기 중")
            self.meas_status.set(f"{self.current_slide_idx}번 슬라이더 대기 중")
            self.log_message(f"파워포인트 동기화 완료 ({slide_count}장)")
            messagebox.showinfo("완료", "파워포인트와 동기화되었습니다.")
        except Exception as e:
            self.log_message(f"파워포인트 동기화 실패: {e}")
            messagebox.showerror("오류", f"파워포인트를 찾을 수 없습니다. 파워포인트 파일이 열려있는지 확인하세요.\n{e}")

    def build_tuning_tab(self):
        frame = ttk.Frame(self.tab4, padding=10)
        frame.pack(fill="both", expand=True)
        
        status_frame = ttk.LabelFrame(frame, text=" Gray 자동 튜닝 ", padding=10)
        status_frame.pack(fill="x", pady=5)
        
        ttk.Label(status_frame, text="현재 노트:").grid(row=0, column=0, sticky="w")
        ttk.Label(status_frame, textvariable=self.loc_var, foreground="blue").grid(row=0, column=1, sticky="w", padx=10)
        
        ttk.Label(status_frame, text="타겟 전류(mA):").grid(row=1, column=0, sticky="w")
        ttk.Label(status_frame, textvariable=self.tune_target_curr, foreground="red", font=("Arial", 10, "bold")).grid(row=1, column=1, sticky="w", padx=10)
        
        ttk.Label(status_frame, text="진행 상태:").grid(row=2, column=0, sticky="w")
        ttk.Label(status_frame, textvariable=self.tune_status, font=("Arial", 36, "bold")).grid(row=2, column=1, sticky="w", padx=10)
        
        ttk.Label(status_frame, text="슬라이드 번호:").grid(row=3, column=0, sticky="w")
        ttk.Label(status_frame, textvariable=self.slide_num_var, font=("Arial", 40, "bold")).grid(row=3, column=1, sticky="w", padx=10)

        cols = ("no", "slide_num", "label", "target", "gray", "offset", "meas", "comp")
        self.tune_tree = ttk.Treeview(frame, columns=cols, show="headings", height=8)
        for col, text in zip(cols, ["No", "슬라이드", "라벨", "목표(mA)", "설정 Gray", "Offset(mA)", "측정(mA)", "보정(mA)"]):
            self.tune_tree.heading(col, text=text)
            self.tune_tree.column(col, width=70, anchor="center")
        self.tune_tree.pack(fill="both", expand=True, pady=10)
        
        self.tune_tree.bind("<Control-c>", lambda e: self.copy_selected_to_clipboard(self.tune_tree))
        self.tune_tree.bind("<Control-C>", lambda e: self.copy_selected_to_clipboard(self.tune_tree))

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill="x", pady=5)
        
        nav = ttk.LabelFrame(btn_frame, text=" 제어 및 저장 ", padding=5)
        nav.pack(side="left", padx=5)
        ttk.Button(nav, text="▲ 이전 슬라이드", command=lambda: self.move_slide(-1)).pack(side="left", padx=5)
        ttk.Button(nav, text="▼ 다음 슬라이드", command=lambda: self.move_slide(1)).pack(side="left", padx=5)
        ttk.Button(nav, text="자동 튜닝 실행", command=self.run_auto_tune).pack(side="left", padx=5)
        ttk.Button(nav, text="수정된 PPT 저장", command=self.save_modified_ppt).pack(side="left", padx=5)

        self.toggle_btn_tune = tk.Button(btn_frame, text="블랙 화면", font=("Malgun Gothic", 12, "bold"), bg="black", fg="white", command=self.toggle_screen, width=15, height=2)
        self.toggle_btn_tune.pack(side="right", padx=10)

    def build_measure_tab(self):
        frame = ttk.Frame(self.tab5, padding=10)
        frame.pack(fill="both", expand=True)
        
        status_frame = ttk.LabelFrame(frame, text=" 최종 측정 모니터링 ", padding=10)
        status_frame.pack(fill="x", pady=5)
        
        ttk.Label(status_frame, text="현재 노트:").grid(row=0, column=0, sticky="w")
        ttk.Label(status_frame, textvariable=self.loc_var, foreground="blue").grid(row=0, column=1, sticky="w", padx=10)
        
        ttk.Label(status_frame, text="진행 상태:").grid(row=1, column=0, sticky="w")
        ttk.Label(status_frame, textvariable=self.meas_status, font=("Arial", 36, "bold")).grid(row=1, column=1, sticky="w", padx=10)
        
        ttk.Label(status_frame, text="슬라이드 번호:").grid(row=2, column=0, sticky="w")
        ttk.Label(status_frame, textvariable=self.slide_num_var, font=("Arial", 40, "bold")).grid(row=2, column=1, sticky="w", padx=10)
        
        ttk.Label(status_frame, text="현재 패턴 Gray:").grid(row=3, column=0, sticky="w")
        ttk.Label(status_frame, textvariable=self.current_gray_var, font=("Arial", 12, "bold"), foreground="purple").grid(row=3, column=1, sticky="w", padx=10)

        cols = ("no", "time", "slide_num", "label", "target", "gray", "offset", "meas", "comp", "lv", "cx", "cy")
        self.measure_tree = ttk.Treeview(frame, columns=cols, show="headings", height=8)
        for col, text in zip(cols, ["No", "측정 시간", "슬라이드", "라벨", "목표(mA)", "Gray", "Offset(mA)", "측정(mA)", "보정(mA)", "휘도(nit)", "cx", "cy"]):
            self.measure_tree.heading(col, text=text)
            self.measure_tree.column(col, width=65, anchor="center")
        self.measure_tree.pack(fill="both", expand=True, pady=10)
        
        self.measure_tree.bind("<Control-c>", lambda e: self.copy_selected_to_clipboard(self.measure_tree))
        self.measure_tree.bind("<Control-C>", lambda e: self.copy_selected_to_clipboard(self.measure_tree))

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill="x", pady=5)
        
        nav = ttk.LabelFrame(btn_frame, text=" 제어 및 저장 ", padding=5)
        nav.pack(side="left", padx=5)
        ttk.Button(nav, text="▲ 이전 슬라이드", command=lambda: self.move_slide(-1)).pack(side="left", padx=5)
        ttk.Button(nav, text="▼ 다음 슬라이드", command=lambda: self.move_slide(1)).pack(side="left", padx=5)
        ttk.Button(nav, text="측정 실행", command=self.run_measurement).pack(side="left", padx=5)
        ttk.Button(nav, text="데이터 CSV 저장", command=self.save_csv).pack(side="left", padx=5)

        self.toggle_btn = tk.Button(btn_frame, text="블랙 화면", font=("Malgun Gothic", 12, "bold"), bg="black", fg="white", command=self.toggle_screen, width=15, height=2)
        self.toggle_btn.pack(side="right", padx=10)

    def check_ppt_sync(self):
        if not self.presentation:
            return False
        try:
            if getattr(self.presentation, 'SlideShowWindow', None) is None:
                messagebox.showwarning("안내", "파워포인트 슬라이드 쇼가 실행 중이 아닙니다. 파워포인트 화면에서 슬라이드 쇼(F5)를 시작한 후 탭을 이동해 주세요.")
                self.slideshow_started = False
                return False
            
            real_idx = self.presentation.SlideShowWindow.View.Slide.SlideIndex
            if real_idx != self.current_slide_idx and real_idx in self.slides_dict:
                self.current_slide_idx = real_idx
                self.update_tuning_info()
                self.tune_status.set(f"{real_idx}번 슬라이더 대기 중")
                self.meas_status.set(f"{real_idx}번 슬라이더 대기 중")
                self.log_message(f"사용자 임의 조작 감지. 프로그램 슬라이드 동기화 (현재: {real_idx})")
            return True
        except Exception as e:
            messagebox.showwarning("안내", "파워포인트 제어 권한을 잃었습니다. 슬라이드 쇼가 강제로 종료되었을 수 있습니다. 슬라이드 쇼를 다시 실행하고 시도해 주세요.")
            self.slideshow_started = False
            return False

    def toggle_screen(self):
        if not self.check_ppt_sync(): return
        try:
            if self.is_black_screen:
                self.presentation.SlideShowWindow.View.State = 1
                self.presentation.SlideShowWindow.View.GotoSlide(self.current_slide_idx) 
                self.toggle_btn.config(text="블랙 화면", bg="black", fg="white")
                if hasattr(self, 'toggle_btn_tune'):
                    self.toggle_btn_tune.config(text="블랙 화면", bg="black", fg="white")
                self.is_black_screen = False
            else:
                self.presentation.SlideShowWindow.View.State = 3
                self.toggle_btn.config(text="슬라이드 화면", bg="white", fg="black")
                if hasattr(self, 'toggle_btn_tune'):
                    self.toggle_btn_tune.config(text="슬라이드 화면", bg="white", fg="black")
                self.is_black_screen = True
        except Exception as e:
            self.log_message(f"화면 전환 오류: {e}")
            messagebox.showerror("오류", f"화면 전환에 실패했습니다: {e}")

    def on_tab_changed(self, event):
        idx = self.tabs.index(self.tabs.select())
        if idx in [3, 4]:
            if not self.slides_dict:
                messagebox.showwarning("경고", "파워포인트를 먼저 동기화하세요.")
                self.tabs.select(2)
                return
            
            if not self.check_ppt_sync():
                self.tabs.select(2)
                return
            
            if idx == 4:
                if self.presentation:
                    try:
                        self.presentation.SlideShowWindow.View.GotoSlide(1)
                        self.current_slide_idx = 1
                        self.tune_status.set("1번 슬라이더 대기 중")
                        self.meas_status.set("1번 슬라이더 대기 중")
                    except Exception as e:
                        self.log_message(f"슬라이드 이동 오류: {e}")
                        
            self.update_tuning_info()

    def update_tuning_info(self):
        if not self.slides_dict: return
        note = self.slides_dict.get(self.current_slide_idx, "")
        self.loc_var.set(note)
        self.slide_num_var.set(f"{self.current_slide_idx} / {len(self.slides_dict)}")
        
        target_mA = None
        match = re.search(r'\(\s*([\d.]+)\s*([mu]?a)?\s*\)', note, re.IGNORECASE)
        if match:
            val = float(match.group(1))
            unit = match.group(2).lower() if match.group(2) else 'a'
            if unit == 'a': target_mA = val * 1000
            elif unit == 'ma': target_mA = val
            elif unit == 'ua': target_mA = val * 1e-3
            else: target_mA = val
            self.tune_target_curr.set(f"{target_mA:.4f}")
        else:
            self.tune_target_curr.set("파싱 불가")
            
        if self.presentation:
            gray_val = self.check_and_get_gray()
            if gray_val != -1:
                self.current_gray_var.set(str(gray_val))
            else:
                self.current_gray_var.set("에러 (패턴 없음/값 불일치/무채색 아님)")

    def check_and_get_gray(self):
        if not self.presentation: return -1
        try:
            slide = self.presentation.Slides(self.current_slide_idx)
            first_gray = None
            for shape in slide.Shapes:
                try:
                    if getattr(shape.Fill, 'Visible', 0) != 0: 
                        rgb = shape.Fill.ForeColor.RGB
                        r = rgb & 0xFF
                        g = (rgb >> 8) & 0xFF
                        b = (rgb >> 16) & 0xFF
                        
                        if not (r == g == b): 
                            return -1
                            
                        if first_gray is None: 
                            first_gray = r
                        elif first_gray != r: 
                            return -1
                except Exception as e:
                    self.log_message(f"도형 색상 확인 중 무시된 오류: {e}")
            return first_gray if first_gray is not None else -1
        except Exception as e:
            self.log_message(f"슬라이드 읽기 오류: {e}")
            return -1

    def move_slide(self, direction):
        if not self.slides_dict: return
        if not self.check_ppt_sync(): return
        new_idx = self.current_slide_idx + direction
        if 1 <= new_idx <= len(self.slides_dict):
            self.current_slide_idx = new_idx
            self.update_tuning_info()
            self.tune_status.set(f"{new_idx}번 슬라이더 대기 중")
            self.meas_status.set(f"{new_idx}번 슬라이더 대기 중")
            if self.presentation:
                try: self.presentation.SlideShowWindow.View.GotoSlide(new_idx)
                except Exception as e: messagebox.showerror("오류", f"슬라이드 이동 실패: {e}")

    def change_ppt_shape_color(self, gray_val):
        if not self.check_ppt_sync(): return
        try:
            color_val = gray_val | (gray_val << 8) | (gray_val << 16)
            slide = self.presentation.Slides(self.current_slide_idx)
            for shape in slide.Shapes:
                try:
                    if getattr(shape.Fill, 'Visible', 0) != 0:
                        shape.Fill.ForeColor.RGB = color_val
                    if getattr(shape.Line, 'Visible', 0) != 0:
                        shape.Line.ForeColor.RGB = color_val
                except Exception as e:
                    pass
            self.device.current_gray = gray_val
        except Exception as e:
            self.log_message(f"도형 색상 변경 실패: {e}")

    def run_auto_tune(self):
        if not self.slides_dict: return
        if not self.check_ppt_sync(): return
        
        target_str = self.tune_target_curr.get()
        if target_str == "파싱 불가":
            self.log_message("타겟 전류를 찾을 수 없어 튜닝을 취소합니다.")
            return
            
        target_mA = float(target_str)
        
        range_str = self.k_curr_range.get()
        if range_str.lower() != "auto":
            range_max_ma = {"10uA": 0.01, "100uA": 0.1, "1mA": 1.0, "10mA": 10.0, "100mA": 100.0, "1A": 1000.0}.get(range_str, 10.0)
            if target_mA > range_max_ma:
                messagebox.showwarning("경고", f"타겟 전류({target_mA:.4f}mA)가 현재 설정된 측정 Range({range_str})를 초과할 수 있습니다.\n(작업은 계속 진행됩니다)")
                self.log_message(f"경고: 타겟 전류({target_mA:.4f}mA)가 Range({range_str}) 초과 위험")
            
        target_A = target_mA / 1000.0
        low, high = 0, 255
        best_gray = 0
        min_diff = float('inf')
        
        self.tune_status.set("블랙 화면 전환/Offset 측정 중...")
        self.log_message("블랙 화면 전환 및 Offset 전류 측정 중 (2초 대기)...")
        try:
            self.presentation.SlideShowWindow.View.State = 1
            self.presentation.SlideShowWindow.View.GotoSlide(self.current_slide_idx) 
        except Exception as e:
            messagebox.showerror("오류", f"화면 상태 복원 실패: {e}")
            return
            
        self.root.update()
        self.wait(2.0)
        
        offset_A = self.device.get_keithley_data()
        if offset_A == float('inf'):
            messagebox.showerror("측정 에러", "전류 값이 범위를 초과했습니다. Range 설정을 높여주세요.")
            return
        offset_mA = offset_A * 1000.0
        
        try:
            self.presentation.SlideShowWindow.View.State = 1
        except Exception as e:
            messagebox.showerror("오류", f"화면 상태 복원 실패: {e}")
            return
            
        self.root.update()
        self.wait(0.5)
        
        self.log_message(f"목표 보정 전류 {target_mA:.4f}mA 튜닝 시작 (Offset: {offset_mA:.4f}mA)")
        
        while low <= high:
            mid = (low + high) // 2
            self.change_ppt_shape_color(mid)
    
            self.tune_status.set(f"튜닝 중... Gray: {mid}")
            self.root.update()
            self.wait(0.3)
    
            curr_A = self.device.get_keithley_data()
            if curr_A == float('inf'):
                messagebox.showerror("측정 에러", "전류 값이 범위를 초과했습니다. Range 설정을 높여주세요.")
                return
            comp_A = curr_A - offset_A
    
            diff = abs(comp_A - target_A)
    
            if diff < min_diff:
                min_diff = diff
                best_gray = mid
        
            if comp_A < target_A:
                low = mid + 1
            else:
                high = mid - 1

        self.change_ppt_shape_color(best_gray)
        self.root.update()  
        self.wait(2.0) 
                
        final_curr_A = self.device.get_keithley_data()
        if final_curr_A == float('inf'):
            messagebox.showerror("측정 에러", "전류 값이 범위를 초과했습니다. Range 설정을 높여주세요.")
            return
        final_comp_A = final_curr_A - offset_A
        final_curr_mA = final_curr_A * 1000.0
        final_comp_mA = final_comp_A * 1000.0
        
        self.tune_status.set(f"완료! Gray: {best_gray}")
        self.log_message(f"튜닝 완료 (Gray: {best_gray}, 측정: {final_curr_mA:.4f}mA, 보정: {final_comp_mA:.4f}mA)")
        
        res = [len(self.tune_tree.get_children())+1, self.current_slide_idx, self.loc_var.get(), f"{target_mA:.4f}", best_gray, f"{offset_mA:.4f}", f"{final_curr_mA:.4f}", f"{final_comp_mA:.4f}"]
        self.tune_tree.insert("", "end", values=res)
        self.tune_tree.yview_moveto(1)
        
        self.root.update()
        self.wait(1.0)
        self.move_slide(1)

    def run_measurement(self):
        if not self.slides_dict: return
        if not self.check_ppt_sync(): return
        
        gray_val = self.check_and_get_gray()
        if gray_val == -1:
            messagebox.showerror("에러", "패턴들의 gray 값이 서로 다르거나 무채색(white/gray)이 아닙니다.\n(측정은 계속 진행됩니다)")
            gray_val = "에러"
        
        target_str = self.tune_target_curr.get()
        target_current = target_str if target_str != "파싱 불가" else "N/A"
        
        if target_str != "파싱 불가":
            target_mA = float(target_str)
            range_str = self.k_curr_range.get()
            if range_str.lower() != "auto":
                range_max_ma = {"10uA": 0.01, "100uA": 0.1, "1mA": 1.0, "10mA": 10.0, "100mA": 100.0, "1A": 1000.0}.get(range_str, 10.0)
                if target_mA > range_max_ma:
                    messagebox.showwarning("경고", f"타겟 전류({target_mA:.4f}mA)가 현재 설정된 측정 Range({range_str})를 초과할 수 있습니다.\n(측정은 계속 진행됩니다)")
                    self.log_message(f"경고: 타겟 전류({target_mA:.4f}mA)가 Range({range_str}) 초과 위험")
                
        self.meas_status.set("블랙 화면 전환/Offset 측정 중...")
        self.log_message("블랙 화면 전환 및 Offset 전류 측정 중 (2초 대기)...")
        try:
            self.presentation.SlideShowWindow.View.State = 3
        except Exception as e:
            messagebox.showerror("오류", f"화면 상태 전환 실패: {e}")
            return
            
        self.root.update()
        self.wait(2.0)
        
        offset_A = self.device.get_keithley_data()
        if offset_A == float('inf'):
            messagebox.showerror("측정 에러", "전류 값이 범위를 초과했습니다. Range 설정을 높여주세요.")
            return
        offset_mA = offset_A * 1000.0
        
        self.meas_status.set("데이터 측정 중...")
        try:
            self.presentation.SlideShowWindow.View.State = 1
            self.presentation.SlideShowWindow.View.GotoSlide(self.current_slide_idx) 
        except Exception as e:
            messagebox.showerror("오류", f"화면 상태 복원 실패: {e}")
            return
            
        self.root.update()
        self.wait(0.5)
        
        curr_A = self.device.get_keithley_data()
        if curr_A == float('inf'):
            messagebox.showerror("측정 에러", "전류 값이 범위를 초과했습니다. Range 설정을 높여주세요.")
            return
        curr_mA = curr_A * 1000.0
        comp_mA = curr_mA - offset_mA
        
        lv, sx, sy = self.device.get_ca310_data()
        current_gray = gray_val 
        self.device.current_gray = current_gray
        self.current_gray_var.set(str(current_gray))
        
        self.meas_status.set("측정 완료!")
        self.log_message(f"최종 측정 기록 완료 (Gray: {current_gray})")
        
        current_time = datetime.now().strftime("%H:%M:%S")
        res = [len(self.measure_results)+1, current_time, self.current_slide_idx, self.loc_var.get(), target_current, current_gray, f"{offset_mA:.4f}", f"{curr_mA:.4f}", f"{comp_mA:.4f}", f"{lv:.2f}", f"{sx:.4f}", f"{sy:.4f}"]
        self.measure_results.append(res)
        self.measure_tree.insert("", "end", values=res)
        self.measure_tree.yview_moveto(1)
        
        self.root.update()
        self.wait(1.0)
        if self.current_slide_idx < len(self.slides_dict):
            self.move_slide(1)

    def copy_selected_to_clipboard(self, tree, event=None):
        selected_items = tree.selection()
        if not selected_items:
            return

        lines = []
        for item in selected_items:
            values = tree.item(item, 'values')
            lines.append("\t".join(map(str, values)))
            
        clipboard_text = "\n".join(lines)
        self.root.clipboard_clear()
        self.root.clipboard_append(clipboard_text)
        self.root.update() 
        self.log_message(f"선택된 {len(selected_items)}개의 행이 클립보드에 복사되었습니다.")

    def save_csv(self):
        if not self.measure_results: return
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")])
        if path:
            mode = self.ca_disp_mode.get()
            if mode == "XYZ":
                headers = ['No', 'Time', 'Slide', 'Label', 'Target(mA)', 'Gray', 'Offset(mA)', 'Measured(mA)', 'Compensated(mA)', 'Y', 'X', 'Z']
            elif mode == "Tduv":
                headers = ['No', 'Time', 'Slide', 'Label', 'Target(mA)', 'Gray', 'Offset(mA)', 'Measured(mA)', 'Compensated(mA)', 'Luminance', 'T', 'duv']
            elif mode == "u'v'":
                headers = ['No', 'Time', 'Slide', 'Label', 'Target(mA)', 'Gray', 'Offset(mA)', 'Measured(mA)', 'Compensated(mA)', 'Luminance', "u'", "v'"]
            else:
                headers = ['No', 'Time', 'Slide', 'Label', 'Target(mA)', 'Gray', 'Offset(mA)', 'Measured(mA)', 'Compensated(mA)', 'Luminance', 'cx', 'cy']
                
            with open(path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                writer.writerow(headers)
                writer.writerows(self.measure_results)
            self.log_message(f"CSV 저장 완료: {path}")

    def save_modified_ppt(self):
        if not self.presentation: return
        path = filedialog.asksaveasfilename(defaultextension=".pptx", filetypes=[("PowerPoint", "*.pptx;*.ppt")])
        if path:
            try:
                path = path.replace('/', '\\')
                self.presentation.SaveAs(path)
                self.log_message(f"수정된 PPT 저장 완료: {path}")
                messagebox.showinfo("저장 완료", "PPT 파일이 저장되었습니다.")
            except Exception as e:
                messagebox.showerror("저장 오류", str(e))

    def close_app(self):
        try:
            self.ppt_app = None
            self.presentation = None
        except Exception as e:
            print(f"COM Release error: {e}")
            
        self.device.release_devices()
        
        try:
            if WIN32_AVAILABLE:
                pythoncom.CoUninitialize()
        except Exception:
            pass
            
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = OLEDMeasurementApp(root)
    root.mainloop()