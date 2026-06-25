import sys
import os
import re
import traceback
import xlwings as xw
import pyperclip
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# --- [수정 3] 파일 생성(on_created) 및 이름 변경(on_moved) 완벽 감지 ---
class WatcherHandler(FileSystemEventHandler):
    def __init__(self, signal):
        super().__init__()
        self.signal = signal

    def on_created(self, event):
        if not event.is_directory:
            self.signal.emit(os.path.basename(event.src_path))
            
    def on_moved(self, event):
        if not event.is_directory:
            self.signal.emit(os.path.basename(event.dest_path))

# --- 메인 미니 프로그램 ---
class ExcelSimpleBot(QWidget):
    file_detected_signal = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.setWindowTitle("엑셀 자동 이동 봇 (무결점 아키텍처)")
        self.resize(450, 230)
        self.setWindowFlags(Qt.WindowStaysOnTopHint) 
        
        self.target_folder = ""
        self.observer = None
        
        # [수정 4] 엑셀 타깃 셀을 인스턴스 변수로 고정 보관
        self.current_target_cell = None 
        # [수정 5] 파일 목록 캐싱용 Set
        self.existing_files_cache = set()

        self.file_detected_signal.connect(self.process_new_file)
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()
        
        self.lbl_folder = QLabel("📁 감시할 폴더를 선택하세요")
        
        btn_folder = QPushButton("📂 폴더 변경")
        btn_folder.clicked.connect(self.change_folder)
        
        btn_start_restart = QPushButton("▶️ 시작 / 재시작 (선택된 셀부터 탐색)")
        btn_start_restart.clicked.connect(self.start_or_restart_program)
        btn_start_restart.setStyleSheet("background-color: #f0ad4e; color: black; font-weight: bold; padding: 8px;")
        
        info_frame = QFrame()
        info_frame.setStyleSheet("background-color: #f8f9fa; border: 1px solid #dee2e6; border-radius: 5px;")
        info_layout = QVBoxLayout()
        
        self.lbl_current_cell = QLabel("📍 현재 활성 셀: -")
        self.lbl_current_cell.setStyleSheet("font-size: 13px; font-weight: bold; color: #495057;")
        
        self.lbl_clipboard = QLabel("📋 클립보드: -")
        self.lbl_clipboard.setStyleSheet("font-size: 15px; font-weight: bold; color: #d63384;")
        
        info_layout.addWidget(self.lbl_current_cell)
        info_layout.addWidget(self.lbl_clipboard)
        info_frame.setLayout(info_layout)
        
        self.lbl_status = QLabel("대기 중...")
        self.lbl_status.setStyleSheet("font-weight: bold; color: blue;")
        
        layout.addWidget(self.lbl_folder)
        layout.addWidget(btn_start_restart)
        layout.addWidget(btn_folder)
        layout.addWidget(info_frame)
        layout.addWidget(self.lbl_status)
        self.setLayout(layout)

    def update_info_display(self, cell_address, clip_text):
        clean_address = cell_address.replace("$", "")
        self.lbl_current_cell.setText(f"📍 현재 타깃 셀: {clean_address}")
        self.lbl_clipboard.setText(f"📋 클립보드: {clip_text}")

    # --- [수정 5] 폴더 캐싱 로직 (시작할 때만 1회 실행) ---
    def build_file_cache(self):
        self.existing_files_cache.clear()
        for root, dirs, files in os.walk(self.target_folder):
            for f in files:
                self.existing_files_cache.add(f)

    # --- [수정 1] 정규식 기반 독립 문자열 매칭 함수 ---
    def is_exact_match(self, target_val, file_name):
        # 파일명 내에서 target_val이 독립적인 단어(구분자: _, -, ., 공백 등)로 존재하는지 검사
        pattern = rf"(?:^|[-_.\s]){re.escape(target_val)}(?:[-_.\s]|$)"
        return bool(re.search(pattern, file_name, re.IGNORECASE))

    # --- [수정 2] 숫자형 셀(1.0 -> 1) 정제 함수 ---
    def clean_cell_value(self, raw_value):
        if raw_value is None:
            return ""
        if isinstance(raw_value, float) and raw_value.is_integer():
            return str(int(raw_value))
        return str(raw_value).strip()

    def change_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "저장 폴더 선택")
        if folder:
            self.target_folder = folder
            self.lbl_folder.setText(f"👀 감시 중: {self.target_folder}")
            
            if self.observer:
                self.observer.stop()
                self.observer.join()
            event_handler = WatcherHandler(self.file_detected_signal)
            self.observer = Observer()
            self.observer.schedule(event_handler, self.target_folder, recursive=True)
            self.observer.start()
            return True
        return False

    def start_or_restart_program(self):
        if not self.target_folder:
            if not self.change_folder():
                self.lbl_status.setText("⚠️ 폴더를 선택해야 시작할 수 있습니다.")
                return
        
        try:
            # 시작 지점은 현재 사용자가 엑셀에서 선택한 셀로 지정
            wb = xw.books.active
            self.current_target_cell = wb.app.selection
            
            # 파일 캐시 빌드 (초기 1회)
            self.build_file_cache()
            
            self.lbl_status.setText("🔄 탐색 중...")
            self.find_and_copy_next_target()
        except Exception as e:
            self.lbl_status.setText(f"⚠️ 엑셀 연동 실패: {str(e)[:30]}")

    # --- 메인 탐색 로직 ---
    def find_and_copy_next_target(self):
        # [수정 7] 무한 루프 방지 상한선
        loop_count = 0
        max_loops = 10000 

        while loop_count < max_loops:
            loop_count += 1

            # [수정 8] Windows 전용 Hidden 속성 예외 처리 방어
            try:
                is_hidden = self.current_target_cell.api.EntireRow.Hidden
            except AttributeError:
                is_hidden = False # Mac 환경 등 미지원 시 False 처리

            # 1. 숨겨진 행 건너뛰기
            while is_hidden:
                self.current_target_cell = self.current_target_cell.offset(row_offset=1, column_offset=0)
                loop_count += 1
                if loop_count > max_loops:
                    self.lbl_status.setText("✅ 한계 도달 (측정 완료)")
                    return
                try:
                    is_hidden = self.current_target_cell.api.EntireRow.Hidden
                except:
                    is_hidden = False

            # 셀 값 정제
            raw_val = self.current_target_cell.value
            val = self.clean_cell_value(raw_val)
            
            # 빈 칸이면 종료
            if not val or val.lower() == 'none':
                self.lbl_status.setText("✅ 남은 측정 항목 없음")
                self.update_info_display("-", "없음 (종료)")
                self.current_target_cell = None
                return

            # 중복 검사 (정규식 기반)
            is_duplicate = any(self.is_exact_match(val, f) for f in self.existing_files_cache)

            if is_duplicate:
                # 글자색 초기화 및 "파일 있음" 기록
                self.current_target_cell.font.color = (0, 0, 0)
                self.current_target_cell.offset(row_offset=0, column_offset=1).value = "파일 있음"
                # 타깃 셀 한 칸 아래로 이동
                self.current_target_cell = self.current_target_cell.offset(row_offset=1, column_offset=0)
            else:
                # 🔴 타깃 확정 (색상 변경 및 클립보드 복사)
                self.current_target_cell.select() 
                self.current_target_cell.font.color = (255, 0, 0)
                
                pyperclip.copy(val)   
                self.update_info_display(self.current_target_cell.address, val)
                self.lbl_status.setText("🟢 측정 대기 중...")
                return 
                
        self.lbl_status.setText("⚠️ 10,000행 초과 (안전 종료)")

    # --- 새 파일 감지 로직 ---
    def process_new_file(self, file_name):
        # 감지된 파일을 캐시에 즉시 추가
        self.existing_files_cache.add(file_name)

        # 타깃 셀이 없으면(측정이 이미 끝났거나 시작 안 함) 무시
        if not self.current_target_cell:
            return

        try:
            # [수정 4] 사용자의 현재 클릭 위치가 아닌, 프로그램이 쥐고 있는 '타깃 셀' 확인
            raw_val = self.current_target_cell.value
            val = self.clean_cell_value(raw_val)

            if not val or val.lower() == 'none':
                return

            # [수정 1] 정확한 단어 매칭
            if self.is_exact_match(val, file_name):
                # 성공 처리
                self.current_target_cell.font.color = (0, 0, 0)
                self.current_target_cell.offset(row_offset=0, column_offset=1).value = "파일 있음"

                # 타깃을 다음 칸으로 업데이트하고 재탐색 가동
                self.current_target_cell = self.current_target_cell.offset(row_offset=1, column_offset=0)
                self.find_and_copy_next_target()

        except Exception as e:
            # [수정 6] 에러 원인 명확화
            err_msg = str(e).split('\n')[0][:40]
            self.lbl_status.setText(f"⚠️ COM/참조 오류: {err_msg}")

    def closeEvent(self, event):
        if self.observer:
            self.observer.stop()
            self.observer.join()
        event.accept()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ExcelSimpleBot()
    ex.show()
    sys.exit(app.exec_())
