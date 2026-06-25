import os
import re
import sys
import queue
import tkinter as tk
from tkinter import filedialog, messagebox
import xlwings as xw
import pyperclip
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# --- 감시 이벤트 핸들러 (스레드 안전을 위해 Queue 사용) ---
class WatcherHandler(FileSystemEventHandler):
    def __init__(self, event_queue):
        super().__init__()
        self.q = event_queue

    def on_created(self, event):
        if not event.is_directory:
            self.q.put(os.path.basename(event.src_path))
            
    def on_moved(self, event):
        if not event.is_directory:
            self.q.put(os.path.basename(event.dest_path))

# --- 초경량 GUI 메인 클래스 ---
class MinimalExcelBot:
    def __init__(self, root):
        self.root = root
        self.root.title("엑셀 자동 봇 (초경량)")
        self.root.geometry("400x220")
        self.root.attributes('-topmost', True) # 화면 항상 위 고정
        self.root.configure(padx=10, pady=10)

        # 상태 변수
        self.target_folder = ""
        self.observer = None
        self.current_target_cell = None
        self.existing_files_cache = set()
        
        # 파일 감시 스레드와 GUI가 안전하게 소통하기 위한 큐
        self.event_queue = queue.Queue()

        self.init_ui()
        self.check_queue() # 주기적으로 큐 확인 시작

    def init_ui(self):
        # 1. 폴더 정보
        self.lbl_folder = tk.Label(self.root, text="📁 폴더: 선택되지 않음", fg="gray", font=("Arial", 10))
        self.lbl_folder.pack(pady=(0, 5), anchor="w")

        # 2. 메인 버튼
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(fill="x", pady=5)
        
        self.btn_start = tk.Button(btn_frame, text="▶️ 시작 / 재시작 (폴더 선택)", bg="#f0ad4e", font=("Arial", 10, "bold"), command=self.start_or_restart)
        self.btn_start.pack(fill="x", ipady=5)

        # 3. 실시간 정보창
        info_frame = tk.Frame(self.root, bd=1, relief="solid", bg="#f8f9fa", padx=5, pady=5)
        info_frame.pack(fill="x", pady=10)
        
        self.lbl_cell = tk.Label(info_frame, text="📍 타깃 셀: -", bg="#f8f9fa", font=("Arial", 10))
        self.lbl_cell.pack(anchor="w")
        
        self.lbl_clip = tk.Label(info_frame, text="📋 복사됨: -", bg="#f8f9fa", fg="#d63384", font=("Arial", 12, "bold"))
        self.lbl_clip.pack(anchor="w", pady=(5, 0))

        # 4. 상태 메시지
        self.lbl_status = tk.Label(self.root, text="대기 중...", fg="blue", font=("Arial", 10, "bold"))
        self.lbl_status.pack(pady=5)

    # --- 코어 로직 ---
    def is_exact_match(self, target_val, file_name):
        pattern = rf"(?:^|[-_.\s]){re.escape(target_val)}(?:[-_.\s]|$)"
        return bool(re.search(pattern, file_name, re.IGNORECASE))

    def clean_cell_value(self, raw_value):
        if raw_value is None: return ""
        if isinstance(raw_value, float) and raw_value.is_integer():
            return str(int(raw_value))
        return str(raw_value).strip()

    def build_file_cache(self):
        self.existing_files_cache.clear()
        for root, dirs, files in os.walk(self.target_folder):
            for f in files:
                self.existing_files_cache.add(f)

    def start_or_restart(self):
        if not self.target_folder:
            folder = filedialog.askdirectory(title="감시할 폴더 선택")
            if not folder: return
            
            self.target_folder = folder
            self.lbl_folder.config(text=f"👀 감시: {self.target_folder[:30]}...", fg="black")
            
            # 워치독 시작
            if self.observer:
                self.observer.stop()
                self.observer.join()
            
            event_handler = WatcherHandler(self.event_queue)
            self.observer = Observer()
            self.observer.schedule(event_handler, self.target_folder, recursive=True)
            self.observer.start()
            
            self.btn_start.config(text="🔄 재시작 (현재 선택된 셀부터 탐색)")

        try:
            wb = xw.books.active
            
            # 기존 타깃 셀이 있다면 검은색으로 원상복구
            if self.current_target_cell:
                try:
                    self.current_target_cell.font.color = (0, 0, 0)
                except Exception:
                    pass

            self.current_target_cell = wb.app.selection
            
            self.build_file_cache()
            self.lbl_status.config(text="🔄 탐색 중...")
            self.find_and_copy_next_target()
        except Exception as e:
            self.lbl_status.config(text="⚠️ 엑셀 연동 실패 (엑셀 창 확인)")

    def find_and_copy_next_target(self):
        loop_count = 0
        max_loops = 10000

        while loop_count < max_loops:
            loop_count += 1
            
            try: is_hidden = self.current_target_cell.api.EntireRow.Hidden
            except: is_hidden = False

            while is_hidden:
                self.current_target_cell = self.current_target_cell.offset(row_offset=1, column_offset=0)
                loop_count += 1
                if loop_count > max_loops: return
                try: is_hidden = self.current_target_cell.api.EntireRow.Hidden
                except: is_hidden = False

            raw_val = self.current_target_cell.value
            val = self.clean_cell_value(raw_val)

            if not val or val.lower() == 'none':
                self.lbl_status.config(text="✅ 남은 측정 항목 없음")
                self.lbl_cell.config(text="📍 타깃 셀: -")
                self.lbl_clip.config(text="📋 복사됨: -")
                self.current_target_cell = None
                return

            is_duplicate = any(self.is_exact_match(val, f) for f in self.existing_files_cache)

            if is_duplicate:
                self.current_target_cell.font.color = (0, 0, 0)
                self.current_target_cell.offset(row_offset=0, column_offset=1).value = "파일 있음"
                self.current_target_cell = self.current_target_cell.offset(row_offset=1, column_offset=0)
            else:
                self.current_target_cell.select()
                self.current_target_cell.font.color = (255, 0, 0) # 🔴 새로운 타깃만 빨간색
                pyperclip.copy(val)
                
                addr = self.current_target_cell.address.replace("$", "")
                self.lbl_cell.config(text=f"📍 타깃 셀: {addr}")
                self.lbl_clip.config(text=f"📋 복사됨: {val}")
                self.lbl_status.config(text="🟢 측정 대기 중...")
                return
                
        self.lbl_status.config(text="⚠️ 10,000행 초과 (안전 종료)")

    # --- 백그라운드 파일 감지 체크 (0.1초마다 실행) ---
    def check_queue(self):
        try:
            while True:
                file_name = self.event_queue.get_nowait()
                self.process_new_file(file_name)
        except queue.Empty:
            pass
        finally:
            self.root.after(100, self.check_queue)

    def process_new_file(self, file_name):
        self.existing_files_cache.add(file_name)
        if not self.current_target_cell: return

        try:
            raw_val = self.current_target_cell.value
            val = self.clean_cell_value(raw_val)
            if not val or val.lower() == 'none': return

            if self.is_exact_match(val, file_name):
                self.current_target_cell.font.color = (0, 0, 0)
                self.current_target_cell.offset(row_offset=0, column_offset=1).value = "파일 있음"
                self.current_target_cell = self.current_target_cell.offset(row_offset=1, column_offset=0)
                self.find_and_copy_next_target()
        except Exception as e:
            self.lbl_status.config(text="⚠️ COM/참조 오류 발생")

    # 💡 [버그 픽스] 이 함수가 클래스 밖으로 나가지 않도록 들여쓰기 완벽 적용!
    def on_closing(self):
        if self.current_target_cell:
            try:
                self.current_target_cell.font.color = (0, 0, 0)
            except Exception:
                pass # 사용자가 이미 엑셀 창을 끈 상태라면 무시

        if self.observer:
            self.observer.stop()
            self.observer.join()
            
        self.root.destroy()

# --- 실행부 ---
if __name__ == "__main__":
    root = tk.Tk()
    app = MinimalExcelBot(root)
    # X버튼(창 닫기)을 누를 때 on_closing 함수를 실행하도록 연결
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()