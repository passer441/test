import tkinter as tk
from tkinter import ttk, messagebox
import win32com.client

class PPTSimpleControl:
    def __init__(self, root):
        self.root = root
        self.root.title("PPT Gray & Slide Control")
        self.root.geometry("400x480") 
        
        self.ppt_app = None
        self.presentation = None
        
        self.setup_ui()
        self.check_initial_connection() 

    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="10 10 10 10")
        main_frame.pack(expand=True, fill="both")

        frame_conn = ttk.LabelFrame(main_frame, text=" 상태 및 연결 ", padding=10)
        frame_conn.pack(fill="x", padx=5, pady=5)
        
        self.conn_status = tk.StringVar(value="연결 상태: 확인 중")
        self.slide_info = tk.StringVar(value="슬라이드: - / -")
        
        ttk.Label(frame_conn, textvariable=self.conn_status).pack(anchor="w")
        ttk.Label(frame_conn, textvariable=self.slide_info, font=("Arial", 9, "bold")).pack(anchor="w")
        ttk.Button(frame_conn, text="PPT 동기화", command=self.connect_ppt).pack(pady=5)

        frame_slide = ttk.LabelFrame(main_frame, text=" 화면 제어 ", padding=10)
        frame_slide.pack(fill="x", padx=5, pady=5)
        
        btn_move = ttk.Frame(frame_slide)
        btn_move.pack(fill="x", pady=2)
        ttk.Button(btn_move, text="◀ 이전", command=lambda: self.move_slide(-1)).pack(side="left", padx=5, expand=True)
        ttk.Button(btn_move, text="다음 ▶", command=lambda: self.move_slide(1)).pack(side="right", padx=5, expand=True)
        
        ttk.Button(frame_slide, text="블랙 화면 토글", command=self.toggle_black_screen).pack(fill="x", padx=5, pady=5)

        frame_gray = ttk.LabelFrame(main_frame, text=" Gray 조절 (0-255) ", padding=10)
        frame_gray.pack(fill="x", padx=5, pady=5)
        
        self.gray_var = tk.StringVar(value="현재 Gray: 확인 중")
        ttk.Label(frame_gray, textvariable=self.gray_var, font=("Arial", 10, "bold")).pack(pady=5)
        
        btn_grid5 = ttk.Frame(frame_gray)
        btn_grid5.pack(pady=2)
        ttk.Button(btn_grid5, text="▲▲ Gray +5", width=12, command=lambda: self.adjust_gray(5)).pack(side="left", padx=5)
        ttk.Button(btn_grid5, text="▼▼ Gray -5", width=12, command=lambda: self.adjust_gray(-5)).pack(side="left", padx=5)

        btn_grid1 = ttk.Frame(frame_gray)
        btn_grid1.pack(pady=2)
        ttk.Button(btn_grid1, text="▲ Gray +1", width=12, command=lambda: self.adjust_gray(1)).pack(side="left", padx=5)
        ttk.Button(btn_grid1, text="▼ Gray -1", width=12, command=lambda: self.adjust_gray(-1)).pack(side="left", padx=5)

    def check_initial_connection(self):
        try:
            self.ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
            self.presentation = self.ppt_app.ActivePresentation
            self.conn_status.set(f"연결: {self.presentation.Name}")
            self.update_info()
        except Exception:
            self.conn_status.set("연결 상태: 연결 실패")
            self.slide_info.set("슬라이드: - / -")

    def connect_ppt(self):
        try:
            self.ppt_app = win32com.client.GetActiveObject("PowerPoint.Application")
            self.presentation = self.ppt_app.ActivePresentation
            self.conn_status.set(f"연결: {self.presentation.Name}")
            self.update_info()
        except Exception:
            self.conn_status.set("연결 상태: 연결 실패")
            self.slide_info.set("슬라이드: - / -")
            messagebox.showerror("오류", "열려 있는 파워포인트를 찾을 수 없습니다.")

    def toggle_black_screen(self):
        try:
            if self.ppt_app and self.ppt_app.SlideShowWindows.Count > 0:
                view = self.ppt_app.SlideShowWindows(1).View
                if view.State == 3:
                    view.State = 1
                else:
                    view.State = 3
            else:
                messagebox.showwarning("알림", "슬라이드 쇼(F5)를 실행해 주세요.")
        except Exception as e:
            print(f"블랙 화면 전환 오류: {e}")

    def update_info(self):
        self.update_gray_display()
        self.update_slide_number()

    def update_slide_number(self):
        try:
            if self.ppt_app.SlideShowWindows.Count > 0:
                current = self.ppt_app.SlideShowWindows(1).View.Slide.SlideIndex
                total = self.presentation.Slides.Count
                self.slide_info.set(f"슬라이드: {current} / {total}")
            else:
                self.slide_info.set("슬라이드: 쇼 실행 필요")
        except:
            pass

    def move_slide(self, direction):
        if not self.presentation: self.connect_ppt()
        try:
            if self.ppt_app.SlideShowWindows.Count > 0:
                view = self.ppt_app.SlideShowWindows(1).View
                new_idx = view.Slide.SlideIndex + direction
                if 1 <= new_idx <= self.presentation.Slides.Count:
                    view.GotoSlide(new_idx)
                    self.root.after(500, self.update_info)
            else:
                messagebox.showwarning("알림", "슬라이드 쇼(F5)를 실행해 주세요.")
        except Exception as e:
            print(f"이동 오류: {e}")

    def get_current_gray(self):
        try:
            if self.ppt_app.SlideShowWindows.Count > 0:
                slide = self.ppt_app.SlideShowWindows(1).View.Slide
                for shape in slide.Shapes:
                    try:
                        rgb = shape.Fill.ForeColor.RGB
                        r, g, b = rgb & 0xFF, (rgb >> 8) & 0xFF, (rgb >> 16) & 0xFF
                        if r == g == b: return r
                    except: continue
            return -1
        except: return -1

    def update_gray_display(self):
        val = self.get_current_gray()
        self.gray_var.set(f"현재 Gray: {val}" if val != -1 else "현재 Gray: 인식 불가")

    def adjust_gray(self, delta):
        current_gray = self.get_current_gray()
        if current_gray == -1:
            messagebox.showwarning("경고", "조절할 수 있는 Gray 패턴이 없습니다.")
            return

        new_gray = max(0, min(255, current_gray + delta))
        color_val = new_gray | (new_gray << 8) | (new_gray << 16)

        try:
            slide = self.ppt_app.SlideShowWindows(1).View.Slide
            for shape in slide.Shapes:
                try:
                    shape.Fill.ForeColor.RGB = color_val
                    if shape.Line.Visible: shape.Line.ForeColor.RGB = color_val
                except: pass
            self.update_gray_display()
        except Exception as e:
            print(f"색상 변경 오류: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = PPTSimpleControl(root)
    root.mainloop()