import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import csv
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
                            return float('inf')
            return float('inf')
        except Exception as e:
            print(f"Keithley 데이터 읽기 실패: {e}")
            return float('inf')

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
                
            if hasattr(self, 'ca_memory') and self.ca_memory: del self.ca_memory
            if hasattr(self, 'ca_probe') and self.ca_probe: del self.ca_probe
            if hasattr(self, 'ca_obj') and self.ca_obj: del self.ca_obj
            if hasattr(self, 'ca_app') and self.ca_app: del self.ca_app
            
            self.ca_memory = None
            self.ca_probe = None
            self.ca_obj = None
            self.ca_app = None
        except Exception as e:
            print(f"장비 연결 해제 중 오류 발생: {e}")


class OLEDMeasurementApp:
    def __init__(self, root):
        self.root = root
        self.root.title("OLED Measurement System")
        self.root.geometry("1000x700")
        
        self.device = RealDevice()
        
        self.k_port = tk.StringVar(value="COM3")
        self.k_baud = tk.StringVar(value="9600")
        self.k_curr_limit = tk.StringVar(value="1.05")
        
        self.k_curr_range = tk.StringVar(value="100mA") 
        self.k_range_desc = tk.StringVar(value="최대 측정 범위: 100mA") 
        self.k_curr_range.trace_add("write", self.update_range_desc)
        
        self.ca_sync = tk.StringVar(value="Universal")
        self.ca_mem = tk.IntVar(value=1)
        self.ca_disp_mode = tk.StringVar(value="Lvxy") 
        
        self.k_connected = False
        self.ca_connected = False

        self.measure_results = []
        self.meas_count_var = tk.IntVar(value=10)
        self.meas_interval_var = tk.DoubleVar(value=1.0)
        self.is_measuring = False
        self.stop_requested = False
        
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

        self.tab1 = ttk.Frame(self.tabs); self.tabs.add(self.tab1, text="1. Keithley 설정")
        self.tab2 = ttk.Frame(self.tabs); self.tabs.add(self.tab2, text="2. CA-310 설정")
        self.tab3 = ttk.Frame(self.tabs); self.tabs.add(self.tab3, text="3. 측정 실행")

        self.build_keithley_tab()
        self.build_ca310_tab()
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
                
            self.run_zero_cal_action()

    def run_zero_cal_action(self):
        if not self.ca_connected: return
        success, details = self.device.perform_zero_cal()
        if success:
            self.log_message(details)
        else:
            self.log_message(details)
            messagebox.showerror("Zero Cal 오류", f"Zero Calibration에 실패했습니다.\n{details}")

    def build_measure_tab(self):
        frame = ttk.Frame(self.tab3, padding=10)
        frame.pack(fill="both", expand=True)
        
        setting_frame = ttk.LabelFrame(frame, text=" 연속 측정 설정 ", padding=10)
        setting_frame.pack(fill="x", pady=5)
        
        ttk.Label(setting_frame, text="연속 측정 횟수:").grid(row=0, column=0, sticky="w", pady=5)
        ttk.Entry(setting_frame, textvariable=self.meas_count_var, width=10).grid(row=0, column=1, padx=10, sticky="w")
        
        ttk.Label(setting_frame, text="측정 간격 (sec):").grid(row=0, column=2, sticky="w", pady=5, padx=(20, 0))
        ttk.Entry(setting_frame, textvariable=self.meas_interval_var, width=10).grid(row=0, column=3, padx=10, sticky="w")

        cols = ("no", "time", "meas", "lv", "cx", "cy")
        self.measure_tree = ttk.Treeview(frame, columns=cols, show="headings", height=15)
        for col, text in zip(cols, ["No", "측정 시간", "측정 전류(mA)", "휘도(nit)", "cx", "cy"]):
            self.measure_tree.heading(col, text=text)
            self.measure_tree.column(col, width=100, anchor="center")
        self.measure_tree.pack(fill="both", expand=True, pady=10)
        
        self.measure_tree.bind("<Control-c>", self.copy_selected_to_clipboard)
        self.measure_tree.bind("<Control-C>", self.copy_selected_to_clipboard)

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill="x", pady=5)
        
        ttk.Button(btn_frame, text="한 번 측정", command=self.run_single_measurement).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="연속 측정", command=self.run_continuous_measurement).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="측정 중지", command=self.stop_measurement).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="데이터 CSV 저장", command=self.save_csv).pack(side="right", padx=5)
        ttk.Button(btn_frame, text="데이터 초기화", command=self.clear_data).pack(side="right", padx=20)

    def perform_measurement(self):
        if not self.k_connected or not self.ca_connected:
            messagebox.showwarning("경고", "Keithley와 CA-310 장비를 모두 연결한 후 측정해주세요.")
            return False

        curr_A = self.device.get_keithley_data()
        if curr_A == float('inf'):
            self.log_message("측정 에러: 전류 값이 범위를 초과했거나 통신에 실패했습니다.")
            messagebox.showerror("측정 에러", "전류 값이 범위를 초과했거나 통신에 실패했습니다.\nRange 설정을 확인해주세요.")
            return False

        curr_mA = curr_A * 1000.0
        lv, sx, sy = self.device.get_ca310_data()
        
        current_time = datetime.now().strftime("%H:%M:%S.%f")[:-3]
        no = len(self.measure_results) + 1
        
        curr_str = f"{curr_mA:.4f}"
        res = [no, current_time, curr_str, f"{lv:.2f}", f"{sx:.4f}", f"{sy:.4f}"]
        
        self.measure_results.append(res)
        self.measure_tree.insert("", "end", values=res)
        self.measure_tree.yview_moveto(1)
        self.log_message(f"측정 완료: 전류 {curr_str}mA, 휘도 {lv:.2f}")

        return True

    def run_single_measurement(self):
        if self.is_measuring: return
        self.is_measuring = True
        try:
            self.perform_measurement()
        finally:
            self.is_measuring = False

    def stop_measurement(self):
        if self.is_measuring:
            self.stop_requested = True
            self.log_message("측정 중지가 요청되었습니다. 현재 측정 후 종료됩니다.")

    def run_continuous_measurement(self):
        if self.is_measuring: return
        if not self.k_connected or not self.ca_connected:
            messagebox.showwarning("경고", "Keithley와 CA-310 장비를 모두 연결한 후 측정해주세요.")
            return
        
        count = self.meas_count_var.get()
        interval = self.meas_interval_var.get()
        
        if count <= 0:
            messagebox.showwarning("경고", "측정 횟수는 1 이상이어야 합니다.")
            return

        self.is_measuring = True
        self.stop_requested = False
        self.log_message(f"연속 측정 시작 (총 {count}회, 간격 {interval}초)")
        
        try:
            for i in range(count):
                if self.stop_requested:
                    self.log_message("사용자에 의해 연속 측정이 중단되었습니다.")
                    break

                success = self.perform_measurement()
                if not success:
                    self.log_message("측정 오류로 인해 연속 측정을 중단합니다.")
                    break

                self.root.update()
                
                if i < count - 1:
                    self.wait(interval)
        finally:
            self.log_message("연속 측정 종료")
            self.is_measuring = False

    def clear_data(self):
        if messagebox.askyesno("확인", "정말로 모든 측정 데이터를 초기화하시겠습니까?"):
            self.measure_results.clear()
            for item in self.measure_tree.get_children():
                self.measure_tree.delete(item)
            self.log_message("측정 데이터가 초기화되었습니다.")

    def get_current_headers(self):
        mode = self.ca_disp_mode.get()
        if mode == "XYZ":
            return ['No', 'Time', 'Measured(mA)', 'Y', 'X', 'Z']
        elif mode == "Tduv":
            return ['No', 'Time', 'Measured(mA)', 'Luminance', 'T', 'duv']
        elif mode == "u'v'":
            return ['No', 'Time', 'Measured(mA)', 'Luminance', "u'", "v'"]
        else:
            return ['No', 'Time', 'Measured(mA)', 'Luminance', 'cx', 'cy']

    def save_csv(self):
        if not self.measure_results:
            messagebox.showwarning("경고", "저장할 데이터가 없습니다.")
            return
            
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")])
        if path:
            try:
                headers = self.get_current_headers()
                with open(path, 'w', newline='', encoding='utf-8-sig') as f:
                    writer = csv.writer(f)
                    writer.writerow(headers)
                    writer.writerows(self.measure_results)
                self.log_message(f"CSV 저장 완료: {path}")
            except Exception as e:
                self.log_message(f"CSV 저장 실패: {e}")
                messagebox.showerror("저장 오류", f"파일 저장 중 오류가 발생했습니다.\n{e}")

    def copy_selected_to_clipboard(self, event=None):
        selected_items = self.measure_tree.selection()
        if not selected_items:
            return

        lines = []
        for item in selected_items:
            values = self.measure_tree.item(item, 'values')
            lines.append("\t".join(map(str, values)))
            
        clipboard_text = "\n".join(lines)
        self.root.clipboard_clear()
        self.root.clipboard_append(clipboard_text)
        self.root.update() 
        self.log_message(f"선택된 {len(selected_items)}개의 행이 클립보드에 복사되었습니다.")

    def close_app(self):
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