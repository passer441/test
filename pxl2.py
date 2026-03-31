import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import numpy as np
import matplotlib.pyplot as plt
import matplotlib
from matplotlib import font_manager
import json
import os
from datetime import datetime

matplotlib.use("TkAgg")

def _set_korean_font():
    candidates = [
        "Malgun Gothic", "AppleGothic", "NanumGothic",
        "NanumBarunGothic", "Gulim", "Dotum", "Batang"
    ]
    available = {f.name for f in font_manager.fontManager.ttflist}
    for name in candidates:
        if name in available:
            matplotlib.rc("font", family=name)
            break
    matplotlib.rcParams["axes.unicode_minus"] = False

_set_korean_font()

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
        self.root.title("픽셀 발광폭 측정 프로그램  v2.6")
        self.root.configure(bg=BG_DARK)
        self.root.geometry("1100x750")
        self.root.minsize(900, 600)

        self.base_image_pil = None
        self.image          = None
        self.gray_image     = None
        self.photo          = None
        self.pixel_scale    = 1.0
        self.zoom_factor    = 1.0
        self.scale_factor   = 1.0
        self.unit           = "px"
        self.start_x = self.start_y = self.line_id = None
        self.drawing_state  = 0
        self.mode           = "measure"
        self.history        = []
        self.img_path       = ""

        self._build_ui()

    def _build_ui(self):
        header = tk.Frame(self.root, bg=BG_PANEL, pady=10)
        header.pack(fill=tk.X)
        tk.Label(header, text="⬛ 픽셀 발광폭 측정", font=("Segoe UI", 15, "bold"), bg=BG_PANEL, fg=TEXT_MAIN).pack(side=tk.LEFT, padx=18)
        tk.Label(header, text="시작점을 클릭하고 마우스를 이동해 종점을 클릭하세요.", font=("Segoe UI", 9), bg=BG_PANEL, fg=TEXT_SUB).pack(side=tk.LEFT)

        body = tk.Frame(self.root, bg=BG_DARK)
        body.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)

        self._build_sidebar(body)
        self._build_canvas_area(body)

    def _build_sidebar(self, parent):
        sidebar = tk.Frame(parent, bg=BG_PANEL, width=220)
        sidebar.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 0))
        sidebar.pack_propagate(False)

        self._section(sidebar, "📁 파일")
        tk.Button(sidebar, text="이미지 불러오기", command=self.load_image, **BTN_STYLE).pack(fill=tk.X, padx=14, pady=(4, 8))

        self.lbl_file = tk.Label(sidebar, text="파일 없음", wraplength=190, font=("Segoe UI", 8), bg=BG_PANEL, fg=TEXT_SUB)
        self.lbl_file.pack(padx=14, anchor="w")

        self._section(sidebar, "📏 스케일 / 캘리브레이션")
        self.lbl_scale = tk.Label(sidebar, text="1 px = 1.0000 px", font=("Segoe UI", 9), bg=BG_PANEL, fg=SUCCESS)
        self.lbl_scale.pack(padx=14, anchor="w", pady=(2, 4))
        tk.Button(sidebar, text="스케일 설정", command=self.calibrate, **BTN2_STYLE).pack(fill=tk.X, padx=14, pady=(0, 8))

        self._section(sidebar, "📋 측정 이력 (다중 선택 가능)")

        hist_frame = tk.Frame(sidebar, bg=BG_PANEL)
        hist_frame.pack(fill=tk.BOTH, expand=True, padx=14, pady=(4, 4))

        self.hist_list = tk.Listbox(hist_frame, bg=BG_CARD, fg=TEXT_MAIN, selectbackground=ACCENT, font=("Segoe UI", 8), relief="flat", bd=0, highlightthickness=1, highlightbackground=BORDER, selectmode=tk.EXTENDED)
        self.hist_list.pack(fill=tk.BOTH, expand=True)
        self.hist_list.bind("<<ListboxSelect>>", self._update_selected_average)

        self.lbl_average = tk.Label(hist_frame, text="선택 항목 평균: 0.000", font=("Segoe UI", 9, "bold"), bg=BG_PANEL, fg=WARNING)
        self.lbl_average.pack(fill=tk.X, pady=(4, 0))

        btn_row = tk.Frame(sidebar, bg=BG_PANEL)
        btn_row.pack(fill=tk.X, padx=14, pady=(0, 6))
        tk.Button(btn_row, text="💾 저장", command=self.save_history, **BTN2_STYLE).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 3))
        tk.Button(btn_row, text="🗑 초기화", command=self.clear_history, **BTN2_STYLE).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(3, 0))

    def _build_canvas_area(self, parent):
        area = tk.Frame(parent, bg=BG_DARK)
        area.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.status_var = tk.StringVar(value="👉 이미지를 불러오세요")
        status_bar = tk.Label(area, textvariable=self.status_var, bg=BG_CARD, fg=TEXT_SUB, font=("Segoe UI", 9), anchor="w", pady=5)
        status_bar.pack(fill=tk.X, padx=0)

        canvas_frame = tk.Frame(area, bg=BG_DARK)
        canvas_frame.pack(fill=tk.BOTH, expand=True)

        hbar = tk.Scrollbar(canvas_frame, orient=tk.HORIZONTAL, bg=BG_CARD, troughcolor=BG_DARK)
        vbar = tk.Scrollbar(canvas_frame, orient=tk.VERTICAL, bg=BG_CARD, troughcolor=BG_DARK)
        hbar.pack(side=tk.BOTTOM, fill=tk.X)
        vbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.canvas = tk.Canvas(canvas_frame, cursor="crosshair", bg=BG_DARK, relief="flat", xscrollcommand=hbar.set, yscrollcommand=vbar.set, highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)
        hbar.config(command=self.canvas.xview)
        vbar.config(command=self.canvas.yview)

        self.canvas.bind("<ButtonPress-1>", self.on_left_click)
        self.canvas.bind("<Motion>", self.on_mouse_move)

        self.canvas.bind("<MouseWheel>", self.on_zoom)
        self.canvas.bind("<Button-4>", self.on_zoom)
        self.canvas.bind("<Button-5>", self.on_zoom)

        self.canvas.bind("<ButtonPress-2>", self.on_pan_start)
        self.canvas.bind("<B2-Motion>", self.on_pan_drag)
        self.canvas.bind("<ButtonPress-3>", self.on_pan_start)
        self.canvas.bind("<B3-Motion>", self.on_pan_drag)

        self.root.bind("<Escape>", self.cancel_drawing)

    def _section(self, parent, title):
        tk.Label(parent, text=title, font=("Segoe UI", 9, "bold"), bg=BG_PANEL, fg=ACCENT2).pack(anchor="w", padx=14, pady=(14, 2))
        tk.Frame(parent, bg=BORDER, height=1).pack(fill=tk.X, padx=14, pady=(0, 4))

    def load_image(self):
        filepath = filedialog.askopenfilename(title="이미지 파일 선택", filetypes=[("이미지 파일", "*.jpg *.jpeg *.png *.bmp *.tif *.tiff"), ("모든 파일", "*.*")])
        if not filepath:
            return

        self.base_image_pil = Image.open(filepath)
        self.image = np.array(self.base_image_pil)
        self.img_path = filepath

        if len(self.image.shape) >= 3:
            rgb_norm   = self.image[..., :3] / 255.0
            rgb_linear = np.power(rgb_norm, 2.2)
            lum_linear = np.dot(rgb_linear, [0.2126, 0.7152, 0.0722])
            self.gray_image = lum_linear * 255.0
        else:
            self.gray_image = self.image.astype(float)

        canvas_w = self.canvas.winfo_width() or 860
        canvas_h = self.canvas.winfo_height() or 680
        
        fit_zoom_w = canvas_w / self.base_image_pil.width
        fit_zoom_h = canvas_h / self.base_image_pil.height
        self.zoom_factor = min(fit_zoom_w, fit_zoom_h, 1.0)
        
        self.drawing_state = 0
        self._redraw_image()

        fname = os.path.basename(filepath)
        self.lbl_file.config(text=fname, fg=TEXT_MAIN)
        self._set_status(f"✅ 불러온 파일: {fname}  ({self.base_image_pil.width} × {self.base_image_pil.height} px)")

    def _redraw_image(self):
        if self.base_image_pil is None:
            return

        new_w = max(1, int(self.base_image_pil.width * self.zoom_factor))
        new_h = max(1, int(self.base_image_pil.height * self.zoom_factor))
        
        resample_filter = Image.NEAREST if self.zoom_factor > 1.0 else Image.LANCZOS
        display_img = self.base_image_pil.resize((new_w, new_h), resample_filter)

        self.photo = ImageTk.PhotoImage(display_img)
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, anchor=tk.NW, image=self.photo)
        self.canvas.config(scrollregion=(0, 0, new_w, new_h))
        
        self.scale_factor = 1.0 / self.zoom_factor
        if self.line_id:
            self.line_id = None
        self.drawing_state = 0

    def on_zoom(self, event):
        if self.base_image_pil is None:
            return

        if event.num == 4 or getattr(event, 'delta', 0) > 0:
            scale = 1.15
        elif event.num == 5 or getattr(event, 'delta', 0) < 0:
            scale = 0.85
        else:
            return

        new_zoom = self.zoom_factor * scale
        if 0.05 <= new_zoom <= 30.0:
            self.zoom_factor = new_zoom
            self._redraw_image()

    def on_pan_start(self, event):
        self.canvas.scan_mark(event.x, event.y)

    def on_pan_drag(self, event):
        self.canvas.scan_dragto(event.x, event.y, gain=1)

    def calibrate(self):
        if self.image is None:
            messagebox.showwarning("경고", "이미지를 먼저 불러오세요.")
            return
        self.mode = "calibrate"
        self._set_status("📏 캘리브레이션 모드: 시작점을 클릭하세요.")

    def on_left_click(self, event):
        if self.base_image_pil is None:
            return

        cx = self.canvas.canvasx(event.x)
        cy = self.canvas.canvasy(event.y)

        if self.drawing_state == 0:
            self.start_x = cx
            self.start_y = cy
            self.drawing_state = 1
            if self.line_id:
                self.canvas.delete(self.line_id)
            color = WARNING if self.mode == "calibrate" else "#f87171"
            self.line_id = self.canvas.create_line(self.start_x, self.start_y, cx, cy, fill=color, width=2, dash=(4, 2) if self.mode == "calibrate" else None)
            self._set_status("📍 시작점이 선택되었습니다. 마우스를 이동하여 종점을 클릭하세요. (취소: ESC 키)")
        else:
            end_x = cx
            end_y = cy
            self.drawing_state = 0
            
            self.canvas.coords(self.line_id, self.start_x, self.start_y, end_x, end_y)
            length_px = np.hypot((end_x - self.start_x) * self.scale_factor, (end_y - self.start_y) * self.scale_factor)

            if length_px < 3:
                self._set_status("⚠️  선이 너무 짧습니다. 다시 클릭하여 측정하세요.")
                self.canvas.delete(self.line_id)
                return

            if self.mode == "calibrate":
                self._do_calibrate(length_px)
            else:
                self._set_status("📊 프로파일 계산 중…")
                self.show_profile(self.start_x * self.scale_factor, self.start_y * self.scale_factor, 
                                  end_x * self.scale_factor, end_y * self.scale_factor, length_px)

    def cancel_drawing(self, event=None):
        if self.drawing_state == 1:
            self.drawing_state = 0
            if self.line_id:
                self.canvas.delete(self.line_id)
            self._set_status("🚫 측정이 취소되었습니다. 시작점을 다시 클릭하세요.")

    def on_mouse_move(self, event):
        if self.base_image_pil is None:
            return
            
        cx = self.canvas.canvasx(event.x)
        cy = self.canvas.canvasy(event.y)
        
        img_x = int(cx * self.scale_factor)
        img_y = int(cy * self.scale_factor)
        h, w = self.gray_image.shape
        
        status_text = ""
        if 0 <= img_x < w and 0 <= img_y < h:
            val = self.gray_image[img_y, img_x]
            status_text = f"🖱  좌표: ({img_x}, {img_y})   밝기: {val:.1f}"

        if self.drawing_state == 1 and self.start_x is not None:
            self.canvas.coords(self.line_id, self.start_x, self.start_y, cx, cy)
            length_px = np.hypot((cx - self.start_x) * self.scale_factor, (cy - self.start_y) * self.scale_factor)
            mode_str = '📏 캘리브레이션' if self.mode == 'calibrate' else '📊 측정'
            status_text += f"   |   {mode_str} 중 — {length_px:.1f} px ({length_px * self.pixel_scale:.3f} {self.unit})"

        if status_text:
            self._set_status(status_text)

    def _do_calibrate(self, length_px):
        win = tk.Toplevel(self.root)
        win.title("스케일 설정")
        win.configure(bg=BG_DARK)
        win.geometry("320x200")
        win.resizable(False, False)
        win.grab_set()

        tk.Label(win, text=f"선택 구간: {length_px:.2f} px", font=("Segoe UI", 10), bg=BG_DARK, fg=TEXT_SUB).pack(pady=(18, 4))
        tk.Label(win, text="실제 길이 입력:", font=("Segoe UI", 10), bg=BG_DARK, fg=TEXT_MAIN).pack()

        entry_frame = tk.Frame(win, bg=BG_DARK)
        entry_frame.pack(pady=6)
        length_var = tk.StringVar()
        unit_var   = tk.StringVar(value="um")
        tk.Entry(entry_frame, textvariable=length_var, width=10, bg=BG_CARD, fg=TEXT_MAIN, insertbackground=TEXT_MAIN, relief="flat", font=("Segoe UI", 11)).pack(side=tk.LEFT, padx=4)
        tk.Entry(entry_frame, textvariable=unit_var, width=5, bg=BG_CARD, fg=TEXT_MAIN, insertbackground=TEXT_MAIN, relief="flat", font=("Segoe UI", 11)).pack(side=tk.LEFT, padx=4)

        def apply():
            try:
                val = float(length_var.get())
                self.pixel_scale = val / length_px
                self.unit = unit_var.get() or "um"
                self.lbl_scale.config(text=f"1 px = {self.pixel_scale:.4f} {self.unit}")
                self._set_status(f"✅ 스케일 설정 완료: 1 px = {self.pixel_scale:.4f} {self.unit}")
                self.mode = "measure"
                win.destroy()
            except ValueError:
                messagebox.showerror("오류", "숫자를 입력하세요.", parent=win)

        def on_close():
            self.mode = "measure"
            win.destroy()

        win.protocol("WM_DELETE_WINDOW", on_close)
        tk.Button(win, text="적용", command=apply, **BTN_STYLE).pack(pady=10)

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

        plt.close('all')
        plt.style.use("dark_background")
        fig, ax = plt.subplots(figsize=(9, 5))
        fig.patch.set_facecolor("#1e1e2e")
        ax.set_facecolor("#2a2a3e")

        ax.plot(distances, profile, color="#a78bfa", linewidth=1.8, label="Luminance")
        ax.fill_between(distances, profile, alpha=0.15, color="#a78bfa")
        ax.set_xlabel(f"Distance ({self.unit})", color="#9492b0")
        ax.set_ylabel("Relative Luminance", color="#9492b0")
        ax.set_title("밝기 프로파일  —  시작점과 종점을 클릭하여 발광폭 지정", color="#e2e0f0", pad=12)
        ax.tick_params(colors="#9492b0")
        ax.grid(True, linestyle="--", alpha=0.25, color="#44425a")
        for spine in ax.spines.values():
            spine.set_edgecolor("#44425a")

        peak_idx = np.argmax(profile)
        peak = profile[peak_idx]
        half_max = peak / 2.0

        left_indices = np.where(profile[:peak_idx] <= half_max)[0]
        left_cross = left_indices[-1] if len(left_indices) > 0 else 0

        right_indices = np.where(profile[peak_idx:] <= half_max)[0]
        right_cross = peak_idx + right_indices[0] if len(right_indices) > 0 else len(profile) - 1

        result_text = ""
        if right_cross > left_cross:
            fwhm_real = (right_cross - left_cross) * self.pixel_scale
            fwhm_left = distances[left_cross]
            fwhm_right = distances[right_cross]

            ax.axhline(half_max, color="#facc15", linestyle=":", linewidth=1, label=f"Half Max ({half_max:.1f})")
            ax.axvspan(fwhm_left, fwhm_right, alpha=0.10, color="#facc15")
            ax.annotate("", xy=(fwhm_right, half_max), xytext=(fwhm_left, half_max), arrowprops=dict(arrowstyle="<->", color="#facc15", lw=1.5))
            ax.text((fwhm_left + fwhm_right) / 2, half_max * 1.02, f"FWHM = {fwhm_real:.2f} {self.unit}", color="#facc15", ha="center", va="bottom", fontsize=9)
            result_text = f"FWHM = {fwhm_real:.2f} {self.unit}"

        ax.legend(loc="upper right", fontsize=8, facecolor="#313145", edgecolor="#44425a", labelcolor=TEXT_MAIN)
        ax.plot(distances[peak_idx], profile[peak_idx], "o", color="#4ade80", markersize=7, label=f"Peak ({profile[peak_idx]:.1f})", zorder=5)

        self.graph_clicks = []
        self.graph_markers = []
        self.result_showing = False

        def on_mouse_motion(event):
            toolbar = fig.canvas.manager.toolbar
            if toolbar.mode != "":
                new_title = f"[{toolbar.mode}] 화면 제어 중입니다. 측정하려면 툴바 버튼을 해제하세요."
                color = "#facc15"
                self.result_showing = False
            else:
                if getattr(self, "result_showing", False):
                    return
                
                if len(self.graph_clicks) == 0:
                    new_title = "밝기 프로파일 — 시작점과 종점을 클릭하여 발광폭 지정"
                    color = "#e2e0f0"
                elif len(self.graph_clicks) == 1:
                    new_title = "종점을 클릭하세요."
                    color = "#e2e0f0"
                else:
                    return

            if ax.get_title() != new_title:
                ax.set_title(new_title, color=color, pad=12)
                fig.canvas.draw_idle()

        def on_plot_click(event):
            if fig.canvas.manager.toolbar.mode != "":
                return
            if event.inaxes != ax:
                return
            if event.button != 1:  # 좌클릭만 허용
                return
            if event.xdata is None:
                return

            self.graph_clicks.append(event.xdata)
            self.result_showing = False

            if len(self.graph_clicks) == 1:
                for m in self.graph_markers:
                    try: m.remove()
                    except: pass
                self.graph_markers.clear()
                
                ax.set_title("종점을 클릭하세요.", color="#e2e0f0", pad=12)
                l = ax.axvline(event.xdata, color="#f87171", linestyle="--", linewidth=1.5)
                self.graph_markers.append(l)
                fig.canvas.draw_idle()

            elif len(self.graph_clicks) == 2:
                for m in self.graph_markers:
                    try: m.remove()
                    except: pass
                self.graph_markers.clear()

                xmin, xmax = sorted(self.graph_clicks)
                width = xmax - xmin
                self.graph_clicks.clear()

                new_title = f"선택 발광폭: {width:.3f} {self.unit}    [{xmin:.3f} → {xmax:.3f} {self.unit}]"
                ax.set_title(new_title, color="#e2e0f0", pad=12)

                l1 = ax.axvline(xmin, color="#f87171", linestyle="--", linewidth=1.5)
                l2 = ax.axvline(xmax, color="#f87171", linestyle="--", linewidth=1.5)
                span_patch = ax.axvspan(xmin, xmax, alpha=0.15, color="#f87171")

                ylim = ax.get_ylim()
                ty   = ylim[0] + (ylim[1] - ylim[0]) * 0.03
                t1 = ax.text(xmin, ty, f" {xmin:.2f}", color="#f87171", va="bottom", ha="right", fontsize=8)
                t2 = ax.text(xmax, ty, f"{xmax:.2f} ", color="#f87171", va="bottom", ha="left", fontsize=8)
                self.graph_markers.extend([l1, l2, span_patch, t1, t2])
                fig.canvas.draw_idle()
                
                self.result_showing = True

                ts = datetime.now().strftime("%H:%M:%S")
                entry = f"[{ts}]  {width:.3f} {self.unit}"
                self.history.append({
                    "time": ts, "width": round(width, 4),
                    "unit": self.unit, "xmin": round(xmin, 4), "xmax": round(xmax, 4),
                    "fwhm": result_text, "file": os.path.basename(self.img_path)
                })
                self.hist_list.insert(tk.END, entry)
                self.hist_list.see(tk.END)
                self._update_selected_average()

        fig.canvas.mpl_connect('motion_notify_event', on_mouse_motion)
        fig.canvas.mpl_connect('button_press_event', on_plot_click)
        plt.tight_layout()
        plt.show()
        self._set_status("✅ 측정 완료. 툴바 버튼이 해제된 상태에서 그래프를 클릭하여 발광폭을 지정하세요.")

    def _update_selected_average(self, event=None):
        selected_indices = self.hist_list.curselection()
        if not selected_indices:
            self.lbl_average.config(text=f"선택 항목 평균: 0.000 {self.unit}")
            return
        
        avg_width = sum(self.history[i]["width"] for i in selected_indices) / len(selected_indices)
        self.lbl_average.config(text=f"선택 항목 평균: {avg_width:.3f} {self.unit}")

    def save_history(self):
        if not self.history:
            messagebox.showinfo("알림", "저장할 측정 결과가 없습니다.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON 파일", "*.json"), ("모든 파일", "*.*")], title="측정 이력 저장")
        if path:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(self.history, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("저장 완료", f"측정 이력이 저장되었습니다:\n{path}")

    def clear_history(self):
        if messagebox.askyesno("초기화", "측정 이력을 모두 삭제할까요?"):
            self.history.clear()
            self.hist_list.delete(0, tk.END)
            self._update_selected_average()
            self._set_status("🗑  이력이 초기화되었습니다.")

    def _set_status(self, msg):
        self.status_var.set(msg)
        self.root.update_idletasks()


if __name__ == "__main__":
    root = tk.Tk()
    app  = PixelWidthAnalyzer(root)
    root.mainloop()