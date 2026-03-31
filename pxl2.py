import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox, ttk
from PIL import Image, ImageTk
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.widgets import SpanSelector
import matplotlib
from matplotlib import font_manager
import json
import os
from datetime import datetime

matplotlib.use("TkAgg")

# ── 한글 폰트 자동 설정 ───────────────────────────────────────────
def _set_korean_font():
    """Windows / macOS / Linux 에서 한글 폰트를 자동으로 감지해 설정."""
    candidates = [
        "Malgun Gothic",   # Windows 기본 한글 폰트
        "AppleGothic",     # macOS
        "NanumGothic",     # Linux (나눔고딕 설치 시)
        "NanumBarunGothic",
        "Gulim",
        "Dotum",
        "Batang",
    ]
    available = {f.name for f in font_manager.fontManager.ttflist}
    for name in candidates:
        if name in available:
            matplotlib.rc("font", family=name)
            break
    # 마이너스 기호 깨짐 방지
    matplotlib.rcParams["axes.unicode_minus"] = False

_set_korean_font()

# ── 다크 테마 팔레트 ──────────────────────────────────────────────
BG_DARK    = "#1e1e2e"
BG_PANEL   = "#2a2a3e"
BG_CARD    = "#313145"
ACCENT     = "#7c6af7"
ACCENT2    = "#a78bfa"
TEXT_MAIN  = "#e2e0f0"
TEXT_SUB   = "#9492b0"
SUCCESS    = "#4ade80"
WARNING    = "#facc15"
BORDER     = "#44425a"

BTN_STYLE  = dict(bg=ACCENT, fg="white", activebackground=ACCENT2,
                  activeforeground="white", relief="flat",
                  font=("Segoe UI", 10, "bold"), padx=12, pady=6,
                  cursor="hand2", bd=0)
BTN2_STYLE = dict(bg=BG_CARD, fg=TEXT_MAIN, activebackground=BORDER,
                  activeforeground=TEXT_MAIN, relief="flat",
                  font=("Segoe UI", 10), padx=10, pady=5,
                  cursor="hand2", bd=0)


class PixelWidthAnalyzer:
    def __init__(self, root):
        self.root = root
        self.root.title("픽셀 발광폭 측정 프로그램  v2.0")
        self.root.configure(bg=BG_DARK)
        self.root.geometry("1100x750")
        self.root.minsize(900, 600)

        # 상태 변수
        self.image        = None
        self.gray_image   = None
        self.photo        = None
        self.pixel_scale  = 1.0
        self.unit         = "px"
        self.start_x = self.start_y = self.line_id = None
        self.mode         = "measure"
        self.history      = []          # 측정 이력
        self.img_path     = ""

        self._build_ui()
        self._show_welcome()

    # ── UI 빌드 ────────────────────────────────────────────────────
    def _build_ui(self):
        # 헤더
        header = tk.Frame(self.root, bg=BG_PANEL, pady=10)
        header.pack(fill=tk.X)
        tk.Label(header, text="⬛ 픽셀 발광폭 측정",
                 font=("Segoe UI", 15, "bold"),
                 bg=BG_PANEL, fg=TEXT_MAIN).pack(side=tk.LEFT, padx=18)
        tk.Label(header, text="이미지를 불러온 뒤 드래그하여 프로파일을 측정하세요",
                 font=("Segoe UI", 9), bg=BG_PANEL, fg=TEXT_SUB).pack(side=tk.LEFT)

        # 본문 (좌 사이드바 + 우 캔버스)
        body = tk.Frame(self.root, bg=BG_DARK)
        body.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)

        self._build_sidebar(body)
        self._build_canvas_area(body)

    def _build_sidebar(self, parent):
        sidebar = tk.Frame(parent, bg=BG_PANEL, width=220)
        sidebar.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 0))
        sidebar.pack_propagate(False)

        # ── 파일 섹션
        self._section(sidebar, "📁 파일")
        tk.Button(sidebar, text="이미지 불러오기", command=self.load_image,
                  **BTN_STYLE).pack(fill=tk.X, padx=14, pady=(4, 8))

        self.lbl_file = tk.Label(sidebar, text="파일 없음", wraplength=190,
                                 font=("Segoe UI", 8), bg=BG_PANEL, fg=TEXT_SUB)
        self.lbl_file.pack(padx=14, anchor="w")

        # ── 스케일 섹션
        self._section(sidebar, "📏 스케일 / 캘리브레이션")
        self.lbl_scale = tk.Label(sidebar, text="1 px = 1.0000 px",
                                  font=("Segoe UI", 9), bg=BG_PANEL, fg=SUCCESS)
        self.lbl_scale.pack(padx=14, anchor="w", pady=(2, 4))
        tk.Button(sidebar, text="스케일 설정 (드래그)", command=self.calibrate,
                  **BTN2_STYLE).pack(fill=tk.X, padx=14, pady=(0, 8))

        # ── 측정 이력 섹션
        self._section(sidebar, "📋 측정 이력")

        hist_frame = tk.Frame(sidebar, bg=BG_PANEL)
        hist_frame.pack(fill=tk.BOTH, expand=True, padx=14, pady=(4, 4))

        self.hist_list = tk.Listbox(hist_frame,
                                    bg=BG_CARD, fg=TEXT_MAIN,
                                    selectbackground=ACCENT,
                                    font=("Segoe UI", 8),
                                    relief="flat", bd=0,
                                    highlightthickness=1,
                                    highlightbackground=BORDER)
        self.hist_list.pack(fill=tk.BOTH, expand=True)

        btn_row = tk.Frame(sidebar, bg=BG_PANEL)
        btn_row.pack(fill=tk.X, padx=14, pady=(0, 6))
        tk.Button(btn_row, text="💾 저장", command=self.save_history,
                  **BTN2_STYLE).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 3))
        tk.Button(btn_row, text="🗑 초기화", command=self.clear_history,
                  **BTN2_STYLE).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(3, 0))

        # ── 도움말 버튼
        tk.Button(sidebar, text="❓ 사용법 보기", command=self._show_welcome,
                  **BTN2_STYLE).pack(fill=tk.X, padx=14, pady=(4, 14))

    def _build_canvas_area(self, parent):
        area = tk.Frame(parent, bg=BG_DARK)
        area.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # 상태바 (툴팁 느낌)
        self.status_var = tk.StringVar(value="👉 이미지를 불러오세요")
        status_bar = tk.Label(area, textvariable=self.status_var,
                               bg=BG_CARD, fg=TEXT_SUB,
                               font=("Segoe UI", 9), anchor="w", pady=5)
        status_bar.pack(fill=tk.X, padx=0)

        # 캔버스 + 스크롤바
        canvas_frame = tk.Frame(area, bg=BG_DARK)
        canvas_frame.pack(fill=tk.BOTH, expand=True)

        hbar = tk.Scrollbar(canvas_frame, orient=tk.HORIZONTAL,
                             bg=BG_CARD, troughcolor=BG_DARK)
        vbar = tk.Scrollbar(canvas_frame, orient=tk.VERTICAL,
                             bg=BG_CARD, troughcolor=BG_DARK)
        hbar.pack(side=tk.BOTTOM, fill=tk.X)
        vbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.canvas = tk.Canvas(canvas_frame, cursor="crosshair",
                                bg=BG_DARK, relief="flat",
                                xscrollcommand=hbar.set,
                                yscrollcommand=vbar.set,
                                highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)
        hbar.config(command=self.canvas.xview)
        vbar.config(command=self.canvas.yview)

        self.canvas.bind("<ButtonPress-1>",   self.on_press)
        self.canvas.bind("<B1-Motion>",        self.on_drag)
        self.canvas.bind("<ButtonRelease-1>",  self.on_release)
        self.canvas.bind("<Motion>",           self.on_mouse_move)

    def _section(self, parent, title):
        """섹션 헤더"""
        tk.Label(parent, text=title,
                 font=("Segoe UI", 9, "bold"),
                 bg=BG_PANEL, fg=ACCENT2).pack(anchor="w", padx=14, pady=(14, 2))
        tk.Frame(parent, bg=BORDER, height=1).pack(fill=tk.X, padx=14, pady=(0, 4))

    # ── 환영/도움말 팝업 ───────────────────────────────────────────
    def _show_welcome(self):
        win = tk.Toplevel(self.root)
        win.title("사용 방법")
        win.configure(bg=BG_DARK)
        win.geometry("460x380")
        win.resizable(False, False)
        win.grab_set()

        tk.Label(win, text="📖  픽셀 발광폭 측정 프로그램 사용법",
                 font=("Segoe UI", 12, "bold"),
                 bg=BG_DARK, fg=TEXT_MAIN).pack(pady=(20, 4))

        steps = [
            ("① 이미지 불러오기",
             "좌측 사이드바의 [이미지 불러오기] 버튼으로\n"
             "JPG, PNG, BMP, TIF 파일을 엽니다."),
            ("② 스케일 설정 (선택)",
             "[스케일 설정] 버튼 클릭 후 실제 길이를 아는\n"
             "구간을 드래그 → 실제 치수와 단위를 입력하면\n"
             "자동으로 um / mm 등 실측값으로 변환됩니다."),
            ("③ 프로파일 측정",
             "이미지 위에서 마우스를 드래그하면\n"
             "해당 선을 따라 밝기(휘도) 프로파일 그래프가 표시됩니다."),
            ("④ 발광폭 선택",
             "그래프에서 마우스로 드래그하면\n"
             "선택 구간의 너비가 자동으로 계산됩니다."),
        ]

        for title, desc in steps:
            f = tk.Frame(win, bg=BG_CARD, pady=8, padx=12)
            f.pack(fill=tk.X, padx=20, pady=3)
            tk.Label(f, text=title, font=("Segoe UI", 9, "bold"),
                     bg=BG_CARD, fg=ACCENT2).pack(anchor="w")
            tk.Label(f, text=desc, font=("Segoe UI", 8),
                     bg=BG_CARD, fg=TEXT_MAIN, justify="left").pack(anchor="w")

        tk.Button(win, text="시작하기", command=win.destroy,
                  **BTN_STYLE).pack(pady=14)

    # ── 이미지 로드 ────────────────────────────────────────────────
    def load_image(self):
        filepath = filedialog.askopenfilename(
            title="이미지 파일 선택",
            filetypes=[("이미지 파일", "*.jpg *.jpeg *.png *.bmp *.tif *.tiff"),
                       ("모든 파일", "*.*")])
        if not filepath:
            return

        img = Image.open(filepath)
        self.image = np.array(img)
        self.img_path = filepath

        if len(self.image.shape) >= 3:
            rgb_norm   = self.image[..., :3] / 255.0
            rgb_linear = np.power(rgb_norm, 2.2)
            lum_linear = np.dot(rgb_linear, [0.2126, 0.7152, 0.0722])
            self.gray_image = lum_linear * 255.0
        else:
            self.gray_image = self.image.astype(float)

        # 이미지 표시 (캔버스 크기에 맞게 축소 미리보기)
        canvas_w = self.canvas.winfo_width()  or 860
        canvas_h = self.canvas.winfo_height() or 680
        display_img = img.copy()
        display_img.thumbnail((canvas_w, canvas_h), Image.LANCZOS)
        self.scale_factor = img.width / display_img.width

        self.photo = ImageTk.PhotoImage(display_img)
        self.canvas.config(scrollregion=(0, 0,
                                         display_img.width,
                                         display_img.height))
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, anchor=tk.NW, image=self.photo)

        fname = os.path.basename(filepath)
        self.lbl_file.config(text=fname, fg=TEXT_MAIN)
        self._set_status(f"✅ 불러온 파일: {fname}  ({img.width} × {img.height} px)")

    # ── 캘리브레이션 ───────────────────────────────────────────────
    def calibrate(self):
        if self.image is None:
            messagebox.showwarning("경고", "이미지를 먼저 불러오세요.")
            return
        self.mode = "calibrate"
        self._set_status("📏 캘리브레이션 모드: 실제 길이를 아는 구간을 드래그하세요")

    # ── 캔버스 이벤트 ──────────────────────────────────────────────
    def on_mouse_move(self, event):
        if self.image is None:
            return
        cx = int(self.canvas.canvasx(event.x) * getattr(self, "scale_factor", 1))
        cy = int(self.canvas.canvasy(event.y) * getattr(self, "scale_factor", 1))
        h, w = self.gray_image.shape
        if 0 <= cx < w and 0 <= cy < h:
            val = self.gray_image[cy, cx]
            self._set_status(f"🖱  좌표: ({cx}, {cy})   밝기: {val:.1f}")

    def on_press(self, event):
        self.start_x = self.canvas.canvasx(event.x)
        self.start_y = self.canvas.canvasy(event.y)
        if self.line_id:
            self.canvas.delete(self.line_id)

    def on_drag(self, event):
        cur_x = self.canvas.canvasx(event.x)
        cur_y = self.canvas.canvasy(event.y)
        if self.line_id:
            self.canvas.delete(self.line_id)
        color = WARNING if self.mode == "calibrate" else "#f87171"
        self.line_id = self.canvas.create_line(
            self.start_x, self.start_y, cur_x, cur_y,
            fill=color, width=2, dash=(4, 2) if self.mode == "calibrate" else None)

        # 실시간 거리 표시
        sf = getattr(self, "scale_factor", 1)
        px = np.hypot((cur_x - self.start_x) * sf,
                      (cur_y - self.start_y) * sf)
        self._set_status(f"{'📏 캘리브레이션' if self.mode == 'calibrate' else '📊 측정'} 중 — "
                         f"{px:.1f} px  ({px * self.pixel_scale:.3f} {self.unit})")

    def on_release(self, event):
        if self.start_x is None or self.image is None:
            return

        end_x = self.canvas.canvasx(event.x)
        end_y = self.canvas.canvasy(event.y)
        sf    = getattr(self, "scale_factor", 1)

        length_px = np.hypot((end_x - self.start_x) * sf,
                              (end_y - self.start_y) * sf)
        if length_px < 3:
            self._set_status("⚠️  선이 너무 짧습니다. 더 길게 드래그하세요.")
            return

        if self.mode == "calibrate":
            self._do_calibrate(length_px)
        else:
            self._set_status("📊 프로파일 계산 중…")
            self.show_profile(
                self.start_x * sf, self.start_y * sf,
                end_x        * sf, end_y        * sf,
                length_px)

    # ── 캘리브레이션 처리 ──────────────────────────────────────────
    def _do_calibrate(self, length_px):
        win = tk.Toplevel(self.root)
        win.title("스케일 설정")
        win.configure(bg=BG_DARK)
        win.geometry("320x200")
        win.resizable(False, False)
        win.grab_set()

        tk.Label(win, text=f"선택 구간: {length_px:.2f} px",
                 font=("Segoe UI", 10), bg=BG_DARK, fg=TEXT_SUB).pack(pady=(18, 4))
        tk.Label(win, text="실제 길이 입력:",
                 font=("Segoe UI", 10), bg=BG_DARK, fg=TEXT_MAIN).pack()

        entry_frame = tk.Frame(win, bg=BG_DARK)
        entry_frame.pack(pady=6)
        length_var = tk.StringVar()
        unit_var   = tk.StringVar(value="um")
        tk.Entry(entry_frame, textvariable=length_var, width=10,
                 bg=BG_CARD, fg=TEXT_MAIN, insertbackground=TEXT_MAIN,
                 relief="flat", font=("Segoe UI", 11)).pack(side=tk.LEFT, padx=4)
        tk.Entry(entry_frame, textvariable=unit_var, width=5,
                 bg=BG_CARD, fg=TEXT_MAIN, insertbackground=TEXT_MAIN,
                 relief="flat", font=("Segoe UI", 11)).pack(side=tk.LEFT, padx=4)

        def apply():
            try:
                val = float(length_var.get())
                self.pixel_scale = val / length_px
                self.unit = unit_var.get() or "um"
                self.lbl_scale.config(
                    text=f"1 px = {self.pixel_scale:.4f} {self.unit}")
                self._set_status(
                    f"✅ 스케일 설정 완료: 1 px = {self.pixel_scale:.4f} {self.unit}")
                win.destroy()
            except ValueError:
                messagebox.showerror("오류", "숫자를 입력하세요.", parent=win)

        tk.Button(win, text="적용", command=apply, **BTN_STYLE).pack(pady=10)
        self.mode = "measure"

    # ── 프로파일 그래프 ────────────────────────────────────────────
    def show_profile(self, x0, y0, x1, y1, length_px):
        num_points = max(int(length_px), 2)
        x_idx = np.linspace(x0, x1, num_points)
        y_idx = np.linspace(y0, y1, num_points)

        h, w = self.gray_image.shape
        valid  = (x_idx >= 0) & (x_idx < w - 1) & (y_idx >= 0) & (y_idx < h - 1)
        x_idx  = x_idx[valid].astype(int)
        y_idx  = y_idx[valid].astype(int)

        if len(x_idx) == 0:
            messagebox.showwarning("오류", "유효한 측정 구간이 없습니다.")
            return

        profile   = self.gray_image[y_idx, x_idx]
        distances = np.linspace(0, len(profile) * self.pixel_scale, len(profile))

        # ── matplotlib 다크 테마
        plt.style.use("dark_background")
        fig, ax = plt.subplots(figsize=(9, 5))
        fig.patch.set_facecolor("#1e1e2e")
        ax.set_facecolor("#2a2a3e")

        ax.plot(distances, profile, color="#a78bfa", linewidth=1.8, label="Luminance")
        ax.fill_between(distances, profile, alpha=0.15, color="#a78bfa")
        ax.set_xlabel(f"Distance ({self.unit})", color="#9492b0")
        ax.set_ylabel("Relative Luminance", color="#9492b0")
        ax.set_title("밝기 프로파일  —  드래그하여 발광폭 선택", color="#e2e0f0", pad=12)
        ax.tick_params(colors="#9492b0")
        ax.grid(True, linestyle="--", alpha=0.25, color="#44425a")
        for spine in ax.spines.values():
            spine.set_edgecolor("#44425a")

        # FWHM 자동 계산 및 표시
        peak     = profile.max()
        half_max = peak / 2.0
        above    = np.where(profile >= half_max)[0]
        result_text = ""
        if len(above) >= 2:
            fwhm_px    = (above[-1] - above[0])
            fwhm_real  = fwhm_px * self.pixel_scale
            fwhm_left  = distances[above[0]]
            fwhm_right = distances[above[-1]]
            ax.axhline(half_max, color="#facc15", linestyle=":", linewidth=1,
                       label=f"Half Max ({half_max:.1f})")
            ax.axvspan(fwhm_left, fwhm_right, alpha=0.10, color="#facc15")
            ax.annotate("", xy=(fwhm_right, half_max),
                        xytext=(fwhm_left, half_max),
                        arrowprops=dict(arrowstyle="<->", color="#facc15", lw=1.5))
            ax.text((fwhm_left + fwhm_right) / 2, half_max * 1.02,
                    f"FWHM = {fwhm_real:.2f} {self.unit}",
                    color="#facc15", ha="center", va="bottom", fontsize=9)
            result_text = f"FWHM = {fwhm_real:.2f} {self.unit}"

        ax.legend(loc="upper right", fontsize=8,
                  facecolor="#313145", edgecolor="#44425a", labelcolor=TEXT_MAIN)

        # ── SpanSelector (수동 선택)
        markers = []

        def onselect(xmin, xmax):
            for m in markers:
                try: m.remove()
                except: pass
            markers.clear()

            width = xmax - xmin
            ax.set_title(f"선택 발광폭: {width:.3f} {self.unit}    "
                         f"[{xmin:.3f} → {xmax:.3f} {self.unit}]",
                         color="#e2e0f0", pad=12)

            l1 = ax.axvline(xmin, color="#f87171", linestyle="--", linewidth=1.5)
            l2 = ax.axvline(xmax, color="#f87171", linestyle="--", linewidth=1.5)
            span_patch = ax.axvspan(xmin, xmax, alpha=0.15, color="#f87171")

            ylim = ax.get_ylim()
            ty   = ylim[0] + (ylim[1] - ylim[0]) * 0.03
            t1 = ax.text(xmin, ty, f" {xmin:.2f}", color="#f87171",
                         va="bottom", ha="right", fontsize=8)
            t2 = ax.text(xmax, ty, f"{xmax:.2f} ", color="#f87171",
                         va="bottom", ha="left", fontsize=8)
            markers.extend([l1, l2, span_patch, t1, t2])
            fig.canvas.draw_idle()

            # 이력 추가
            ts = datetime.now().strftime("%H:%M:%S")
            entry = f"[{ts}]  {width:.3f} {self.unit}"
            self.history.append({
                "time": ts, "width": round(width, 4),
                "unit": self.unit, "xmin": round(xmin, 4), "xmax": round(xmax, 4),
                "fwhm": result_text, "file": os.path.basename(self.img_path)
            })
            self.hist_list.insert(tk.END, entry)
            self.hist_list.see(tk.END)

        self.span = SpanSelector(
            ax, onselect, "horizontal", useblit=False,
            props=dict(alpha=0.2, facecolor="#f87171"))

        # 최대값 마커
        peak_idx = np.argmax(profile)
        ax.plot(distances[peak_idx], profile[peak_idx], "o",
                color="#4ade80", markersize=7, label=f"Peak ({profile[peak_idx]:.1f})",
                zorder=5)

        plt.tight_layout()
        plt.show()
        self._set_status("✅ 측정 완료. 그래프를 드래그하여 발광폭을 선택하세요.")

    # ── 이력 저장 / 초기화 ─────────────────────────────────────────
    def save_history(self):
        if not self.history:
            messagebox.showinfo("알림", "저장할 측정 결과가 없습니다.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON 파일", "*.json"), ("모든 파일", "*.*")],
            title="측정 이력 저장")
        if path:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(self.history, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("저장 완료", f"측정 이력이 저장되었습니다:\n{path}")

    def clear_history(self):
        if messagebox.askyesno("초기화", "측정 이력을 모두 삭제할까요?"):
            self.history.clear()
            self.hist_list.delete(0, tk.END)
            self._set_status("🗑  이력이 초기화되었습니다.")

    # ── 유틸 ──────────────────────────────────────────────────────
    def _set_status(self, msg):
        self.status_var.set(msg)
        self.root.update_idletasks()


if __name__ == "__main__":
    root = tk.Tk()
    app  = PixelWidthAnalyzer(root)
    root.mainloop()