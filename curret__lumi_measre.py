import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import csv
import re
import time
from datetime import datetime
import serial  
import os       
import shutil   
import sys

try:
    import win32com.client
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False

class RealDevice:
    # [타 AI 및 EXAONE을 위한 상세 주석]
    # 이 클래스는 가상 장비 대신 실제 Keithley 2400과 Konica Minolta CA-310 장비를 제어합니다.
    
    current_gray = 0
    keithley_inst = None 
    ca_app = None        
    ca_obj = None        
    ca_probe = None      
    ca_memory = None     

    @classmethod
    def connect_keithley(cls, port, baud, curr_limit):
        """
        Keithley 2400 장비와 시리얼 통신(pyserial) 연결을 설정합니다.
        SCPI 명령어를 사용하여 장비를 초기화하고 전압을 인가하며 전류를 측정할 준비를 합니다.
        """
        try:
            # 1. 시리얼 포트 연결 (Timeout 1초)
            cls.keithley_inst = serial.Serial(port, int(baud), timeout=1)
            
            # 2. 장비 초기화 (*RST: Reset)
            cls.keithley_inst.write(b"*RST\n")
            time.sleep(0.1)
            
            # 3. 순수 전류계(Ammeter) 모드 설정: 전압 출력을 0V로 고정하여 회로에 영향을 주지 않음
            cls.keithley_inst.write(b":SOUR:FUNC VOLT\n")
            cls.keithley_inst.write(b":SOUR:VOLT 0\n")
            
            # [추가] UI에서 입력받은 전류 제한(Compliance) 값 설정
            cls.keithley_inst.write(f":SENS:CURR:PROT {curr_limit}\n".encode('utf-8'))
            
            # 4. 전류 측정 기능 활성화 (:SENS:FUNC 'CURR')
            cls.keithley_inst.write(b":SENS:FUNC 'CURR'\n")
            
            # 5. 장비 출력 켜기 (:OUTP ON)
            cls.keithley_inst.write(b":OUTP ON\n")
            
            return True, f"[{port}] 실제 Keithley 2400 연결 성공"
        except Exception as e:
            return False, f"Keithley 연결 실패: {e}"

    @classmethod
    def connect_ca310(cls, sync_mode, channel, display_mode):
        """
        CA-310 장비와 CA-SDK(win32com)를 통해 USB 통신 연결을 설정합니다.
        매뉴얼에 따라 단일 장비 및 단일 프로브를 대상으로 자동 설정(AutoConnect)을 수행합니다.
        """
        try:
            # [추가] 64비트 파이썬 강제 종료 방지
            if sys.maxsize > 2**32:
                return False, "64비트 파이썬 환경입니다. CA-SDK는 32비트 환경에서만 동작하므로 연결할 수 없습니다."
                
            if not WIN32_AVAILABLE:
                return False, "win32com 모듈을 사용할 수 없습니다."

            # 1. CA-SDK Application 루트 객체 생성 (타입 라이브러리에 정의된 CA200Srvr)
            cls.ca_app = win32com.client.Dispatch("CA200Srvr.Ca200")
            
            # 2. AutoConnect 실행 (USB에 연결된 1대의 장비와 1개의 프로브를 자동 인식하여 구성 설정)
            cls.ca_app.AutoConnect()
            
            # 3. 하위 제어 객체 생성
            cls.ca_obj = cls.ca_app.SingleCa        
            cls.ca_probe = cls.ca_obj.SingleProbe   
            cls.ca_memory = cls.ca_obj.Memory       
            
            # 4. Sync Mode 설정 (Universal: 3, NTSC: 0, PAL: 1)
            sync_dict = {"Universal": 3, "NTSC": 0, "PAL": 1}
            cls.ca_obj.SyncMode = sync_dict.get(sync_mode, 3)
            
            # 5. 메모리 채널 번호 매핑
            cls.ca_memory.ChannelNO = int(channel)
            
            # 6. 측정 디스플레이 모드 설정
            cls.ca_obj.DisplayMode = int(display_mode)
            
            return True, "[USB] 실제 CA-310 연결 성공"
        except Exception as e:
            return False, f"CA-310 연결 실패: {e}"

    @classmethod
    def perform_zero_cal(cls):
        """
        CA-310 장비의 영점 교정(Zero Calibration)을 실행합니다.
        실제 광학 측정을 진행하기 전에 노이즈 보정을 위해 반드시 선행되어야 합니다.
        """
        try:
            if cls.ca_obj:
                # CalZero 메서드 호출 시 영점 교정 절차 수행
                cls.ca_obj.CalZero()
                return True, "실제 CA-310 Zero Calibration 완료"
            return False, "CA-310 장비 객체가 초기화되지 않았습니다."
        except Exception as e:
            return False, f"Zero Cal 실패: {e}"

    @classmethod
    def get_keithley_data(cls):
        """
        Keithley 2400에 SCPI 읽기 명령을 전달하고 현재 흐르는 전류(A) 값을 반환합니다.
        """
        try:
            if cls.keithley_inst:
                # 1. 단일 측정 수행 및 버퍼에 결과 요청
                cls.keithley_inst.write(b":READ?\n")
                
                # 2. 결과 값 읽어오기 (결과 형식은 CSV 형태로 전압, 전류, 저항, 시간, 상태 값을 반환함)
                data = cls.keithley_inst.readline().decode('utf-8').strip()
                
                if data:
                    values = data.split(',')
                    # 3. 배열에서 1번째 인덱스인 전류(Current) 값을 추출해 실수 변환
                    # [추가] 타임아웃 및 데이터 파싱 에러 방지
                    if len(values) >= 2:
                        try:
                            return float(values[1])
                        except ValueError:
                            return 0.0
            return 0.0
        except Exception as e:
            print(f"Keithley 데이터 읽기 실패: {e}")
            return 0.0

    @classmethod
    def get_ca310_data(cls):
        """
        CA-310 장비에 광학 측정을 지시하고 설정된 모드에 맞는 데이터를 반환합니다.
        """
        try:
            if cls.ca_obj and cls.ca_probe:
                # 1. 측정 실행 트리거 (Measure 메서드 호출)
                cls.ca_obj.Measure()
                
                # 2. 현재 설정된 DisplayMode 확인하여 데이터 취득
                mode = cls.ca_obj.DisplayMode
                
                if mode == 7: # XYZ 모드
                    lv = cls.ca_probe.Y
                    sx = cls.ca_probe.X
                    sy = cls.ca_probe.Z
                elif mode == 1: # Tduv 모드
                    lv = cls.ca_probe.Lv
                    sx = cls.ca_probe.T
                    sy = cls.ca_probe.duv
                elif mode == 5: # u'v' 모드
                    lv = cls.ca_probe.Lv
                    sx = cls.ca_probe.ud
                    sy = cls.ca_probe.vd
                else: # 기본 Lvxy 모드 (0)
                    lv = cls.ca_probe.Lv
                    sx = cls.ca_probe.sx
                    sy = cls.ca_probe.sy
                
                return lv, sx, sy
            return 0.0, 0.0, 0.0
        except Exception as e:
            print(f"CA-310 측정 실패: {e}")
            return 0.0, 0.0, 0.0
            
    @classmethod
    def release_devices(cls):
        """
        애플리케이션 종료 시 장비와의 통신 포트 및 COM 객체를 안전하게 반환합니다.
        """
        try:
            if cls.keithley_inst:
                # 출력을 끄고 시리얼 포트를 해제함
                cls.keithley_inst.write(b":OUTP OFF\n")
                cls.keithley_inst.close()
                cls.keithley_inst = None
                
            # CA-SDK 관련 COM 객체들을 해제함
            cls.ca_obj = None
            cls.ca_probe = None
            cls.ca_memory = None
            cls.ca_app = None
        except Exception as e:
            print(f"장비 연결 해제 중 오류 발생: {e}")

class OLEDMeasurementApp:
    def __init__(self, root):
        self.root = root
        self.root.title("OLED Measurement System (Virtual Test Mode)")
        self.root.geometry("1150x700")
        
        self.k_port = tk.StringVar(value="COM3")
        self.k_baud = tk.StringVar(value="9600")
        self.k_volt = tk.StringVar(value="10V")
        self.k_curr_limit = tk.StringVar(value="1.05")
        
        self.ca_sync = tk.StringVar(value="Universal")
        self.ca_mem = tk.IntVar(value=1)
        self.ca_disp_mode = tk.StringVar(value="Lvxy") # 색좌표 모드 변수 추가
        
        self.ppt_path = tk.StringVar(value="선택된 파일 없음")
        self.temp_ppt_path = None  # 임시 파일 경로 저장 변수 추가
        
        self.slides = []
        self.current_slide_idx = 0
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
        self.meas_status = tk.StringVar(value="1번 슬라이더 대기 중")  # 5번 탭 진행 상태 변수 추가
        self.slide_num_var = tk.StringVar(value="- / -")
        
        self.setup_ui()
        # 닫기 버튼 이벤트 연결
        self.root.protocol("WM_DELETE_WINDOW", self.close_app)

    def setup_ui(self):
        main_pane = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_pane.pack(expand=True, fill="both", padx=10, pady=10)

        self.tabs = ttk.Notebook(main_pane)
        main_pane.add(self.tabs, weight=7)
        self.tabs.bind("<<NotebookTabChanged>>", self.on_tab_changed)

        self.tab1 = ttk.Frame(self.tabs); self.tabs.add(self.tab1, text="1. Keithley 설정")
        self.tab2 = ttk.Frame(self.tabs); self.tabs.add(self.tab2, text="2. CA-310 설정")
        self.tab3 = ttk.Frame(self.tabs); self.tabs.add(self.tab3, text="3. PPT 로드")
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

        form = ttk.LabelFrame(frame, text=" 통신 세부 설정 ", padding=15)
        form.pack(fill="x", pady=10)
        ttk.Label(form, text="COM 포트:").grid(row=0, column=0, sticky="w", pady=5)
        ttk.Entry(form, textvariable=self.k_port).grid(row=0, column=1, padx=10, sticky="w")
        ttk.Label(form, text="Baud Rate:").grid(row=1, column=0, sticky="w", pady=5)
        ttk.Entry(form, textvariable=self.k_baud).grid(row=1, column=1, padx=10, sticky="w")
        ttk.Label(form, text="전류 리미트(A):").grid(row=2, column=0, sticky="w", pady=5)
        ttk.Entry(form, textvariable=self.k_curr_limit).grid(row=2, column=1, padx=10, sticky="w")
        
        ttk.Button(form, text="장비 연결 실행", command=self.connect_keithley_action).grid(row=3, column=0, columnspan=2, pady=15)

    def connect_keithley_action(self):
        self.k_connected, details = RealDevice.connect_keithley(self.k_port.get(), self.k_baud.get(), self.k_curr_limit.get())
        self.log_message(details)

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
        ttk.Button(btn_frame, text="Zero Cal 실행", command=self.run_zero_cal_action).pack(side="left", padx=5)

    def connect_ca310_action(self):
        disp_dict = {"Lvxy": 0, "XYZ": 7, "Tduv": 1, "u'v'": 5}
        mode_val = disp_dict.get(self.ca_disp_mode.get(), 0)
        
        self.ca_connected, details = RealDevice.connect_ca310(self.ca_sync.get(), self.ca_mem.get(), mode_val)
        self.log_message(details)
        if self.ca_connected:
            # 선택된 모드에 따라 측정 결과 표의 헤더를 동적으로 변경합니다.
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
        success, details = RealDevice.perform_zero_cal()
        if success:
            self.ca_zero_calibrated = True
            self.log_message(details)

    def build_ppt_tab(self):
        frame = ttk.Frame(self.tab3, padding=20)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="실제 PPT 로드", font=("Malgun Gothic", 16, "bold")).pack(pady=20)
        file_frame = ttk.Frame(frame)
        file_frame.pack(pady=10)
        ttk.Entry(file_frame, textvariable=self.ppt_path, width=50).pack(side="left", padx=5)
        ttk.Button(file_frame, text="파일 선택", command=self.load_ppt_action).pack(side="left")
        
        guide_text = (
            "첫 번째 슬라이드에는 offset 전류를 측정하도록 블랙 슬라이드를 추가해 주세요!\n"
            "슬라이드 노트에는 나중에 측정 데이터의 측정 위치를 식별할 수 있는 Label을 적어 주세요!\n"
            "그리고 각 슬라이드의 Target 전류는 슬라이드 노트에 ( 0.113 mA ) 와 같이 표기해 주세요!\n"
            "슬라이드 노트 작성 예) 양산 ( 0.231 mA )"
        )
        self.info_lbl = ttk.Label(frame, text=guide_text, justify="left", foreground="blue")
        self.info_lbl.pack(pady=20)

    def load_ppt_action(self):
        if not WIN32_AVAILABLE: return
        original_path = filedialog.askopenfilename(filetypes=[("PowerPoint", "*.pptx;*.ppt")])
        if original_path:
            # ---------------------------------------------------------
            # 로딩 시작 안내 창 띄우기 (업데이트)
            # ---------------------------------------------------------
            loading_win = tk.Toplevel(self.root)
            loading_win.title("PPT 파일 로드 중")
            loading_win.geometry("300x100")
            # 모달 창으로 설정 (사용자가 다른 작업 못하게)
            loading_win.transient(self.root)
            loading_win.grab_set()
            
            tk.Label(loading_win, text="PPT 파일을 분석 중입니다...\n잠시만 기다려 주세요.", font=("Malgun Gothic", 11)).pack(expand=True)
            self.root.update()
            
            try:
                # 1. 파일 경로 정제
                original_path = original_path.replace('/', '\\')
                
                # 2. 임시 파일(_temp) 경로 생성
                base_name, ext = os.path.splitext(original_path)
                temp_path = f"{base_name}_temp{ext}"
                
                # 3. 원본 파일을 임시 파일로 복사
                shutil.copy2(original_path, temp_path)
                
                # 4. UI에는 임시 파일 경로를 표시 및 변수에 저장
                self.ppt_path.set(temp_path)
                self.temp_ppt_path = temp_path
                
                # 5. 파워포인트 앱 시작 및 임시 파일 열기
                self.ppt_app = win32com.client.Dispatch("PowerPoint.Application")
                self.presentation = self.ppt_app.Presentations.Open(temp_path, WithWindow=True)
                
                slide_count = self.presentation.Slides.Count
                self.slides = []
                for i in range(1, slide_count + 1):
                    slide = self.presentation.Slides(i)
                    note = ""
                    try:
                        for shape in slide.NotesPage.Shapes:
                            if shape.HasTextFrame and shape.TextFrame.HasText:
                                t = shape.TextFrame.TextRange.Text.strip()
                                if len(t) > len(note): note = t
                    except: pass
                    self.slides.append(note if note else f"Slide {i}")
                self.current_slide_idx = 0
                self.log_message(f"PPT 임시 복사본 로드 완료 ({slide_count}장)")
            except Exception as e:
                messagebox.showerror("오류", f"파일 로드 실패: {e}")
            finally:
                # 작업 완료 후 로딩 창 닫기
                loading_win.destroy()

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
        # 4번 탭 진행 현황 폰트 4배 (약 36pt) 확대
        ttk.Label(status_frame, textvariable=self.tune_status, font=("Arial", 36, "bold")).grid(row=2, column=1, sticky="w", padx=10)
        
        ttk.Label(status_frame, text="슬라이드 번호:").grid(row=3, column=0, sticky="w")
        ttk.Label(status_frame, textvariable=self.slide_num_var, font=("Arial", 40, "bold")).grid(row=3, column=1, sticky="w", padx=10)

        cols = ("no", "slide_num", "label", "target", "gray", "offset", "meas", "comp")
        self.tune_tree = ttk.Treeview(frame, columns=cols, show="headings", height=8)
        for col, text in zip(cols, ["No", "슬라이드", "라벨", "목표(mA)", "설정 Gray", "Offset(mA)", "측정(mA)", "보정(mA)"]):
            self.tune_tree.heading(col, text=text)
            self.tune_tree.column(col, width=70, anchor="center")
        self.tune_tree.pack(fill="both", expand=True, pady=10)

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
        
        # 5번 탭 진행 상태 추가
        ttk.Label(status_frame, text="진행 상태:").grid(row=1, column=0, sticky="w")
        ttk.Label(status_frame, textvariable=self.meas_status, font=("Arial", 36, "bold")).grid(row=1, column=1, sticky="w", padx=10)
        
        ttk.Label(status_frame, text="슬라이드 번호:").grid(row=2, column=0, sticky="w")
        ttk.Label(status_frame, textvariable=self.slide_num_var, font=("Arial", 40, "bold")).grid(row=2, column=1, sticky="w", padx=10)
        
        cols = ("no", "slide_num", "label", "target", "gray", "offset", "meas", "comp", "lv", "cx", "cy")
        self.measure_tree = ttk.Treeview(frame, columns=cols, show="headings", height=8)
        for col, text in zip(cols, ["No", "슬라이드", "라벨", "목표(mA)", "Gray", "Offset(mA)", "측정(mA)", "보정(mA)", "휘도(nit)", "cx", "cy"]):
            self.measure_tree.heading(col, text=text)
            self.measure_tree.column(col, width=65, anchor="center")
        self.measure_tree.pack(fill="both", expand=True, pady=10)

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
                messagebox.showerror("오류", "슬라이드 쇼 창을 찾을 수 없습니다. 슬라이드 쇼를 다시 시작해 주세요.")
                self.slideshow_started = False
                return False
            
            real_idx = self.presentation.SlideShowWindow.View.CurrentShowPosition - 1
            if real_idx != self.current_slide_idx and 0 <= real_idx < len(self.slides):
                self.current_slide_idx = real_idx
                self.update_tuning_info()
                self.log_message(f"사용자 임의 조작 감지. 프로그램 슬라이드 동기화 (현재: {real_idx + 1})")
            return True
        except Exception as e:
            messagebox.showerror("오류", f"파워포인트 제어 권한 상실.\n원인: {e}\n파일을 다시 로드하거나 프로그램을 재시작해 주세요.")
            self.slideshow_started = False
            return False

    def toggle_screen(self):
        if not self.check_ppt_sync(): return
        try:
            if self.is_black_screen:
                self.presentation.SlideShowWindow.View.State = 1
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
            if not self.slides:
                messagebox.showwarning("경고", "PPT 파일을 로드하세요.")
                self.tabs.select(2)
                return
            if not self.slideshow_started: self.start_slideshow()
            
            if self.slideshow_started:
                self.check_ppt_sync()
            
            if idx == 4:
                self.current_slide_idx = 0
                if self.presentation:
                    try:
                        self.presentation.SlideShowWindow.View.GotoSlide(1)
                    except Exception as e:
                        self.log_message(f"슬라이드 이동 오류: {e}")
                        
            self.update_tuning_info()

    def start_slideshow(self):
        try:
            self.presentation.SlideShowSettings.Run()
            self.slideshow_started = True
            self.current_slide_idx = 0
            self.log_message("PPT 슬라이드 쇼 시작")
        except Exception as e:
            messagebox.showerror("오류", f"슬라이드 쇼 시작 실패: {e}")

    def update_tuning_info(self):
        if not self.slides: return
        note = self.slides[self.current_slide_idx]
        self.loc_var.set(note)
        self.slide_num_var.set(f"{self.current_slide_idx + 1} / {len(self.slides)}")
        
        match = re.search(r'\(\s*([\d.]+)\s*([mu]?a)?\s*\)', note, re.IGNORECASE)
        if match:
            val = float(match.group(1))
            unit = match.group(2).lower() if match.group(2) else 'a'
            if unit == 'a': target_mA = val * 1000
            elif unit == 'ma': target_mA = val
            elif unit == 'ua': target_mA = val * 1e-3
            self.tune_target_curr.set(f"{target_mA:.4f}")
        else:
            self.tune_target_curr.set("파싱 불가")

    def move_slide(self, direction):
        if not self.slides: return
        if not self.check_ppt_sync(): return
        new_idx = self.current_slide_idx + direction
        if 0 <= new_idx < len(self.slides):
            self.current_slide_idx = new_idx
            self.update_tuning_info()
            self.tune_status.set(f"{new_idx + 1}번 슬라이더 대기 중")
            self.meas_status.set(f"{new_idx + 1}번 슬라이더 대기 중")
            if self.presentation:
                try: self.presentation.SlideShowWindow.View.GotoSlide(new_idx + 1)
                except Exception as e: messagebox.showerror("오류", f"슬라이드 이동 실패: {e}")

    def change_ppt_shape_color(self, gray_val):
        if not self.check_ppt_sync(): return
        try:
            color_val = gray_val | (gray_val << 8) | (gray_val << 16)
            slide = self.presentation.Slides(self.current_slide_idx + 1)
            for shape in slide.Shapes:
                try:
                    shape.Fill.ForeColor.RGB = color_val
                except:
                    pass
            RealDevice.current_gray = gray_val
        except Exception as e:
            self.log_message(f"도형 색상 변경 실패: {e}")

    def run_auto_tune(self):
        if not self.slides: return
        if not self.check_ppt_sync(): return
        
        target_str = self.tune_target_curr.get()
        if target_str == "파싱 불가":
            self.log_message("타겟 전류를 찾을 수 없어 튜닝을 취소합니다.")
            return
            
        target_mA = float(target_str)
        target_A = target_mA / 1000.0
        tolerance_A = target_A * 0.05
        low, high = 0, 255
        best_gray = 0
        
        self.tune_status.set("블랙 화면 전환/Offset 측정 중...")
        self.log_message("블랙 화면 전환 및 Offset 전류 측정 중 (2초 대기)...")
        try:
            self.presentation.SlideShowWindow.View.State = 3
        except Exception as e:
            messagebox.showerror("오류", f"화면 상태 전환 실패: {e}")
            return
            
        self.root.update()
        time.sleep(2.0)
        
        offset_A = RealDevice.get_keithley_data()
        offset_mA = offset_A * 1000.0
        
        try:
            self.presentation.SlideShowWindow.View.State = 1
        except Exception as e:
            messagebox.showerror("오류", f"화면 상태 복원 실패: {e}")
            return
            
        self.root.update()
        time.sleep(0.5)
        
        self.log_message(f"목표 보정 전류 {target_mA:.4f}mA 튜닝 시작 (Offset: {offset_mA:.4f}mA)")
        
        while low <= high:
            mid = (low + high) // 2
            self.change_ppt_shape_color(mid)
    
            self.tune_status.set(f"튜닝 중... Gray: {mid}")
            self.root.update()
            time.sleep(0.3)
    
            curr_A = RealDevice.get_keithley_data()
            comp_A = curr_A - offset_A
            comp_mA = comp_A * 1000.0
    
            if abs(comp_A - target_A) <= tolerance_A:
                best_gray = mid
                break
            elif comp_A < target_A:
                low = mid + 1
            else:
                high = mid - 1
                
        best_gray = mid
        self.change_ppt_shape_color(best_gray)
        final_curr_A = RealDevice.get_keithley_data()
        final_comp_A = final_curr_A - offset_A
        final_curr_mA = final_curr_A * 1000.0
        final_comp_mA = final_comp_A * 1000.0
        
        self.tune_status.set(f"완료! Gray: {best_gray}")
        self.log_message(f"튜닝 완료 (Gray: {best_gray}, 측정: {final_curr_mA:.4f}mA, 보정: {final_comp_mA:.4f}mA)")
        
        res = [len(self.tune_tree.get_children())+1, self.current_slide_idx + 1, self.loc_var.get(), f"{target_mA:.4f}", best_gray, f"{offset_mA:.4f}", f"{final_curr_mA:.4f}", f"{final_comp_mA:.4f}"]
        self.tune_tree.insert("", "end", values=res)
        self.tune_tree.yview_moveto(1)
        
        self.root.update()
        time.sleep(1.0)
        self.move_slide(1)

    def run_measurement(self):
        if not self.slides: return
        if not self.check_ppt_sync(): return
        
        target_str = self.tune_target_curr.get()
        target_current = target_str if target_str != "파싱 불가" else "N/A"
        
        self.meas_status.set("블랙 화면 전환/Offset 측정 중...")
        self.log_message("블랙 화면 전환 및 Offset 전류 측정 중 (2초 대기)...")
        try:
            self.presentation.SlideShowWindow.View.State = 3
        except Exception as e:
            messagebox.showerror("오류", f"화면 상태 전환 실패: {e}")
            return
            
        self.root.update()
        time.sleep(2.0)
        
        offset_A = RealDevice.get_keithley_data()
        offset_mA = offset_A * 1000.0
        
        self.meas_status.set("데이터 측정 중...")
        try:
            self.presentation.SlideShowWindow.View.State = 1
        except Exception as e:
            messagebox.showerror("오류", f"화면 상태 복원 실패: {e}")
            return
            
        self.root.update()
        time.sleep(0.5)
        
        curr_A = RealDevice.get_keithley_data()
        curr_mA = curr_A * 1000.0
        comp_mA = curr_mA - offset_mA
        
        lv, sx, sy = RealDevice.get_ca310_data()
        current_gray = RealDevice.current_gray
        
        self.meas_status.set("측정 완료!")
        self.log_message(f"최종 측정 기록 완료 (Gray: {current_gray})")
        
        res = [len(self.measure_results)+1, self.current_slide_idx + 1, self.loc_var.get(), target_current, current_gray, f"{offset_mA:.4f}", f"{curr_mA:.4f}", f"{comp_mA:.4f}", f"{lv:.2f}", f"{sx:.4f}", f"{sy:.4f}"]
        self.measure_results.append(res)
        self.measure_tree.insert("", "end", values=res)
        self.measure_tree.yview_moveto(1)
        
        self.root.update()
        time.sleep(1.0)
        if self.current_slide_idx < len(self.slides) - 1:
            self.move_slide(1)

    def save_csv(self):
        if not self.measure_results: return
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")])
        if path:
            mode = self.ca_disp_mode.get()
            if mode == "XYZ":
                headers = ['No', 'Slide', 'Label', 'Target(mA)', 'Gray', 'Offset(mA)', 'Measured(mA)', 'Compensated(mA)', 'Y', 'X', 'Z']
            elif mode == "Tduv":
                headers = ['No', 'Slide', 'Label', 'Target(mA)', 'Gray', 'Offset(mA)', 'Measured(mA)', 'Compensated(mA)', 'Luminance', 'T', 'duv']
            elif mode == "u'v'":
                headers = ['No', 'Slide', 'Label', 'Target(mA)', 'Gray', 'Offset(mA)', 'Measured(mA)', 'Compensated(mA)', 'Luminance', "u'", "v'"]
            else:
                headers = ['No', 'Slide', 'Label', 'Target(mA)', 'Gray', 'Offset(mA)', 'Measured(mA)', 'Compensated(mA)', 'Luminance', 'cx', 'cy']
                
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
        # 1. 파워포인트 앱 안전하게 종료
        if WIN32_AVAILABLE and self.ppt_app:
            try:
                if self.presentation:
                    try: self.presentation.SlideShowWindow.View.Exit()
                    except: pass
                    self.presentation.Close()
                self.ppt_app.Quit()
            except: pass
            
        # 2. 장비 연결 해제
        RealDevice.release_devices()
        
        # 3. 임시 파일 삭제 로직
        if self.temp_ppt_path and os.path.exists(self.temp_ppt_path):
            try:
                time.sleep(1)  # 파워포인트 프로세스가 완전히 반환되기를 잠시 대기
                os.remove(self.temp_ppt_path)
            except Exception as e:
                print(f"임시 파일 삭제 실패: {e}")
                
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = OLEDMeasurementApp(root)
    root.mainloop()