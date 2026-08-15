# -*- coding: utf-8 -*-
"""
ap_theme.py — AN PHƯỚC GUI DESIGN SYSTEM (tự chứa, dùng cho Tkinter/ttk)
════════════════════════════════════════════════════════════════════════
Đồng bộ tone với web portal https://anphuoc-portal.onrender.com/
  • Navy chủ đạo : #1B2A4A
  • Crimson nhấn : #C41E3A
  • Nền sáng     : #F4F5F7
  • Chữ slate    : #1E293B
  • Card trắng   : #FFFFFF
  • Font         : Be Vietnam Pro (load runtime từ ./gui_assets/fonts)

Cách dùng tối thiểu:
    import ap_theme
    ap_theme.load_fonts()            # gọi TRƯỚC khi tạo Tk() nếu có thể
    root = tk.Tk()
    T = ap_theme.apply(root)         # trả về object chứa màu + font + style names
    ap_theme.Header(root, "Tiêu đề", "Phụ đề").pack(fill="x")
    ...
    ap_theme.Footer(root).pack(side="bottom", fill="x")

Mọi GUI của An Phước đều mang mark: "Phát triển bởi khoind@anphuoc.com.vn".
Tác giả: khoind@anphuoc.com.vn
"""
from __future__ import annotations
import os
import sys
import tkinter as tk
from tkinter import ttk
import tkinter.font as tkfont

# ─────────────────────────────────────────── BẢNG MÀU (brand An Phước) ──────
NAVY        = "#1B2A4A"   # thương hiệu / header
NAVY_DARK   = "#141F38"   # hover đậm
NAVY_LIGHT  = "#2A3D63"   # đường kẻ nhẹ trên nền navy
CRIMSON     = "#C41E3A"   # accent / primary action
CRIMSON_DK  = "#A5182F"   # hover đỏ
SURFACE     = "#F4F5F7"   # nền cửa sổ
CARD        = "#FFFFFF"   # nền card
TEXT        = "#1E293B"   # chữ chính
TEXT_MUTED  = "#6B7280"   # chữ phụ
TEXT_ONNAVY = "#FFFFFF"   # chữ trên nền navy
SUBTLE_ON_NAVY = "#AEB7CC"  # phụ đề trên navy
BORDER      = "#E2E5EA"   # viền card / ô nhập
OK_GREEN    = "#047857"
WARN_AMBER  = "#B45309"
ERR_RED     = "#BE123C"

MARK_TEXT = "Phát triển bởi khoind@anphuoc.com.vn"

# Family sẽ được xác định sau khi load font (fallback Segoe UI)
FONT_FAMILY = "Segoe UI"
_FONTS_LOADED = False


def _asset_dir() -> str:
    """Thư mục chứa font, cạnh file ap_theme.py này."""
    base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "gui_assets", "fonts")


def load_fonts(font_dir: str | None = None) -> str:
    """Đăng ký Be Vietnam Pro ở chế độ private (không cần cài vào Windows).
    Trả về tên family sẽ dùng. Nên gọi trước khi tạo Tk()."""
    global FONT_FAMILY, _FONTS_LOADED
    if _FONTS_LOADED:
        return FONT_FAMILY
    font_dir = font_dir or _asset_dir()
    if sys.platform.startswith("win") and os.path.isdir(font_dir):
        try:
            import ctypes
            FR_PRIVATE = 0x10
            gdi32 = ctypes.windll.gdi32
            loaded_any = False
            for fn in os.listdir(font_dir):
                if fn.lower().endswith(".ttf"):
                    path = os.path.join(font_dir, fn)
                    if gdi32.AddFontResourceExW(ctypes.c_wchar_p(path), FR_PRIVATE, 0):
                        loaded_any = True
            if loaded_any:
                FONT_FAMILY = "Be Vietnam Pro"
        except Exception:
            FONT_FAMILY = "Segoe UI"
    _FONTS_LOADED = True
    return FONT_FAMILY


class Theme:
    """Gói màu + font + tên style để code app dùng cho gọn."""
    def __init__(self, family: str):
        self.family = family
        # (family, size, weight) — Be Vietnam Pro có Bold nên "bold" map đúng
        self.f_display = (family, 20, "bold")
        self.f_h1      = (family, 15, "bold")
        self.f_h2      = (family, 12, "bold")
        self.f_body    = (family, 10)
        self.f_body_b  = (family, 10, "bold")
        self.f_small   = (family, 9)
        self.f_mono    = ("Consolas", 9)
        # màu (tiện truy cập qua T.navy ...)
        self.navy, self.crimson, self.surface = NAVY, CRIMSON, SURFACE
        self.card, self.text, self.muted, self.border = CARD, TEXT, TEXT_MUTED, BORDER
        self.ok, self.warn, self.err = OK_GREEN, WARN_AMBER, ERR_RED


def apply(root: tk.Misc) -> Theme:
    """Áp theme vào root + cấu hình ttk styles. Trả về Theme."""
    family = load_fonts()
    T = Theme(family)

    try:
        root.configure(bg=SURFACE)
    except Exception:
        pass

    style = ttk.Style(root)
    try:
        style.theme_use("clam")   # clam cho phép tuỳ biến màu button/entry nhiều nhất
    except Exception:
        pass

    # ---- Frames ----
    style.configure("AP.TFrame", background=SURFACE)
    style.configure("Card.TFrame", background=CARD)

    # ---- Labels ----
    style.configure("AP.TLabel", background=SURFACE, foreground=TEXT, font=T.f_body)
    style.configure("Card.TLabel", background=CARD, foreground=TEXT, font=T.f_body)
    style.configure("H1.TLabel", background=SURFACE, foreground=NAVY, font=T.f_h1)
    style.configure("CardH1.TLabel", background=CARD, foreground=NAVY, font=T.f_h1)
    style.configure("CardH2.TLabel", background=CARD, foreground=NAVY, font=T.f_h2)
    style.configure("Muted.TLabel", background=SURFACE, foreground=TEXT_MUTED, font=T.f_small)
    style.configure("CardMuted.TLabel", background=CARD, foreground=TEXT_MUTED, font=T.f_small)

    # ---- Primary button (crimson) ----
    style.configure("Primary.TButton", background=CRIMSON, foreground="#FFFFFF",
                    font=T.f_body_b, borderwidth=0, focusthickness=0,
                    padding=(16, 9))
    style.map("Primary.TButton",
              background=[("pressed", CRIMSON_DK), ("active", CRIMSON_DK),
                          ("disabled", "#D9A7AF")],
              foreground=[("disabled", "#FBEAED")])

    # ---- Secondary button (navy) ----
    style.configure("Secondary.TButton", background=NAVY, foreground="#FFFFFF",
                    font=T.f_body_b, borderwidth=0, focusthickness=0,
                    padding=(16, 9))
    style.map("Secondary.TButton",
              background=[("pressed", NAVY_DARK), ("active", NAVY_DARK),
                          ("disabled", "#9AA3B5")])

    # ---- Ghost button (viền, nền sáng) ----
    style.configure("Ghost.TButton", background=CARD, foreground=NAVY,
                    font=T.f_body_b, borderwidth=1, bordercolor=BORDER,
                    focusthickness=0, padding=(14, 8))
    style.map("Ghost.TButton",
              background=[("active", "#EEF1F5")],
              bordercolor=[("active", NAVY)])

    # ---- Checkbutton ----
    style.configure("AP.TCheckbutton", background=CARD, foreground=TEXT, font=T.f_body)
    style.map("AP.TCheckbutton", background=[("active", CARD)])

    # ---- Entry / Combobox ----
    style.configure("AP.TEntry", fieldbackground=CARD, bordercolor=BORDER,
                    lightcolor=BORDER, darkcolor=BORDER, foreground=TEXT, padding=6)
    style.configure("AP.TCombobox", fieldbackground=CARD, background=CARD,
                    bordercolor=BORDER, foreground=TEXT, padding=5, arrowcolor=NAVY)

    # ---- Progressbar ----
    style.configure("AP.Horizontal.TProgressbar", background=CRIMSON,
                    troughcolor=BORDER, borderwidth=0, thickness=6)
    style.configure("TProgressbar", background=CRIMSON, troughcolor=BORDER, borderwidth=0)

    # ---- Base classes: widget ttk KHÔNG đặt style vẫn on-brand ----
    # (giúp các widget tạo động trong app cũ tự đồng bộ, không cần sửa code)
    style.configure("TFrame", background=CARD)
    style.configure("TLabel", background=CARD, foreground=TEXT, font=T.f_body)
    style.configure("TLabelframe", background=CARD, bordercolor=BORDER,
                    relief="solid", borderwidth=1)
    style.configure("TLabelframe.Label", background=CARD, foreground=NAVY, font=T.f_body_b)
    style.configure("TEntry", fieldbackground=CARD, bordercolor=BORDER, lightcolor=BORDER,
                    darkcolor=BORDER, foreground=TEXT, insertcolor=TEXT, padding=5)
    style.configure("TCombobox", fieldbackground=CARD, background=CARD, bordercolor=BORDER,
                    foreground=TEXT, arrowcolor=NAVY, padding=4)
    style.map("TCombobox",
              fieldbackground=[("readonly", CARD)], foreground=[("readonly", TEXT)],
              selectbackground=[("readonly", CARD)], selectforeground=[("readonly", TEXT)])
    style.configure("TCheckbutton", background=CARD, foreground=TEXT, font=T.f_body)
    style.map("TCheckbutton", background=[("active", CARD)])
    style.configure("TButton", background=NAVY, foreground="#FFFFFF", font=T.f_body_b,
                    borderwidth=0, focusthickness=0, padding=(12, 7))
    style.map("TButton", background=[("pressed", NAVY_DARK), ("active", NAVY_DARK),
                                     ("disabled", "#9AA3B5")])
    style.configure("Vertical.TScrollbar", background="#CBD2DC", troughcolor=SURFACE,
                    bordercolor=SURFACE, arrowcolor=NAVY)
    style.configure("Horizontal.TScrollbar", background="#CBD2DC", troughcolor=SURFACE,
                    bordercolor=SURFACE, arrowcolor=NAVY)

    # ---- Notebook (tab) ----
    style.configure("TNotebook", background=CARD, borderwidth=0, tabmargins=(2, 6, 2, 0))
    style.configure("TNotebook.Tab", background="#E7EAEF", foreground=TEXT_MUTED,
                    padding=(16, 9), font=T.f_body_b, borderwidth=0)
    style.map("TNotebook.Tab",
              background=[("selected", CARD)], foreground=[("selected", NAVY)])

    # ---- Treeview ----
    style.configure("Treeview", background=CARD, fieldbackground=CARD, foreground=TEXT,
                    borderwidth=0, rowheight=25, font=T.f_body)
    style.configure("Treeview.Heading", background=NAVY, foreground="#FFFFFF",
                    font=T.f_body_b, relief="flat", padding=(6, 5))
    style.map("Treeview", background=[("selected", NAVY)], foreground=[("selected", "#FFFFFF")])
    style.map("Treeview.Heading", background=[("active", NAVY_DARK)])

    # ---- Radiobutton ----
    style.configure("TRadiobutton", background=CARD, foreground=TEXT, font=T.f_body)
    style.map("TRadiobutton", background=[("active", CARD)])

    return T


# ────────────────────────────────────────────────────────── Logo (badge) ──
def _logo_path() -> str:
    return os.path.join(os.path.dirname(os.path.abspath(__file__)),
                        "gui_assets", "LogoAP.png")


_LOGO_PIL_CACHE: dict = {}


def _logo_pil(size: int, radius: int):
    """Dựng ảnh PIL: badge trắng bo góc chứa logo đỏ (nền logo trắng đặc)."""
    key = (size, radius)
    if key in _LOGO_PIL_CACHE:
        return _LOGO_PIL_CACHE[key]
    path = _logo_path()
    if not os.path.exists(path):
        return None
    try:
        from PIL import Image, ImageDraw
        mask = Image.new("L", (size, size), 0)
        ImageDraw.Draw(mask).rounded_rectangle([0, 0, size - 1, size - 1],
                                               radius=radius, fill=255)
        chip = Image.composite(Image.new("RGBA", (size, size), (255, 255, 255, 255)),
                               Image.new("RGBA", (size, size), (0, 0, 0, 0)), mask)
        inner = size - 12
        logo = Image.open(path).convert("RGBA").resize((inner, inner), Image.LANCZOS)
        off = (size - inner) // 2
        chip.alpha_composite(logo, (off, off))
        a = Image.composite(chip.split()[3], Image.new("L", (size, size), 0), mask)
        chip.putalpha(a)
    except Exception:
        return None
    _LOGO_PIL_CACHE[key] = chip
    return chip


def make_logo_photo(master, size: int = 46, radius: int = 11):
    """Trả về PhotoImage badge logo (ưu tiên PIL bo góc; fallback PNG native)."""
    pil = _logo_pil(size, radius)
    if pil is not None:
        try:
            from PIL import ImageTk
            return ImageTk.PhotoImage(pil, master=master)
        except Exception:
            pass
    path = _logo_path()
    if os.path.exists(path):
        try:
            img = tk.PhotoImage(file=path, master=master)
            factor = max(1, round(img.width() / size))
            return img.subsample(factor, factor)
        except Exception:
            pass
    return None


# ─────────────────────────────────────────────────────── Component: Header ──
class Header(tk.Frame):
    """Thanh header thương hiệu: nền navy + logo An Phước + tiêu đề + gạch crimson."""
    def __init__(self, parent, title: str, subtitle: str | None = None,
                 app_tag: str | None = None, family: str | None = None):
        super().__init__(parent, bg=NAVY)
        fam = family or FONT_FAMILY

        inner = tk.Frame(self, bg=NAVY)
        inner.pack(fill="x", padx=20, pady=(14, 12))

        # Logo An Phước (badge). Fallback monogram 'AP' nếu không load được ảnh.
        self._logo_ref = make_logo_photo(self, size=46)
        if self._logo_ref is not None:
            tk.Label(inner, image=self._logo_ref, bg=NAVY, bd=0).pack(side="left")
        else:
            tk.Label(inner, text="AP", bg=CRIMSON, fg="#FFFFFF",
                     font=(fam, 16, "bold"), padx=11, pady=4).pack(side="left")

        txt = tk.Frame(inner, bg=NAVY)
        txt.pack(side="left", padx=12)
        tk.Label(txt, text=title, bg=NAVY, fg=TEXT_ONNAVY,
                 font=(fam, 15, "bold"), anchor="w").pack(anchor="w")
        if subtitle:
            tk.Label(txt, text=subtitle, bg=NAVY, fg=SUBTLE_ON_NAVY,
                     font=(fam, 9), anchor="w").pack(anchor="w")

        if app_tag:
            tk.Label(inner, text=app_tag, bg=NAVY_LIGHT, fg=SUBTLE_ON_NAVY,
                     font=(fam, 8, "bold"), padx=9, pady=3).pack(side="right")

        # Gạch nhấn crimson dưới header
        tk.Frame(self, bg=CRIMSON, height=3).pack(fill="x")


# ─────────────────────────────────────────────────────── Component: Footer ──
class Footer(tk.Frame):
    """Footer mang mark tác giả (bắt buộc cho mọi GUI An Phước)."""
    def __init__(self, parent, extra: str | None = None, family: str | None = None):
        super().__init__(parent, bg=SURFACE)
        fam = family or FONT_FAMILY
        tk.Frame(self, bg=BORDER, height=1).pack(fill="x")
        row = tk.Frame(self, bg=SURFACE)
        row.pack(fill="x", padx=20, pady=7)
        tk.Label(row, text="●", bg=SURFACE, fg=CRIMSON,
                 font=(fam, 8)).pack(side="left")
        tk.Label(row, text="  " + MARK_TEXT, bg=SURFACE, fg=TEXT_MUTED,
                 font=(fam, 9)).pack(side="left")
        if extra:
            tk.Label(row, text=extra, bg=SURFACE, fg=TEXT_MUTED,
                     font=(fam, 9)).pack(side="right")


def make_card(parent, padx: int = 16, pady: int = 14) -> tk.Frame:
    """Tạo 1 'card' trắng có viền nhẹ để nhóm nội dung."""
    outer = tk.Frame(parent, bg=BORDER)          # giả viền 1px
    inner = tk.Frame(outer, bg=CARD)
    inner.pack(fill="both", expand=True, padx=1, pady=1)
    inner._pad = (padx, pady)
    return inner


# ══════════════════════════════════════════════ Component cao cấp (VIP) ══════
def premium_card(parent, accent: str = NAVY) -> tk.Frame:
    """Card trắng có DẢI NHẤN màu ở đỉnh + viền mảnh. Trả về frame nội dung;
    caller pack/grid `frame.master` (wrap)."""
    wrap = tk.Frame(parent, bg=BORDER)
    tk.Frame(wrap, bg=accent, height=3).pack(fill="x")
    inner = tk.Frame(wrap, bg=CARD)
    inner.pack(fill="both", expand=True, padx=1, pady=(0, 1))
    return inner


def _round_rect(cv: tk.Canvas, x1, y1, x2, y2, r, **kw):
    pts = [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2,
           x2 - r, y2, x1 + r, y2, x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]
    return cv.create_polygon(pts, smooth=True, **kw)


_PILL_COLORS = {
    "ok":    (OK_GREEN,  "#E7F6EE"),
    "err":   (ERR_RED,   "#FCE8EC"),
    "info":  (NAVY,      "#E8ECF4"),
    "warn":  (WARN_AMBER,"#FDF3E3"),
    "muted": (TEXT_MUTED,"#EEF1F5"),
}


class Pill(tk.Canvas):
    """Nhãn bo tròn (pill) tự co theo chữ — hiển thị trạng thái/verdict."""
    def __init__(self, parent, text: str, kind: str = "muted",
                 bgcolor: str = CARD, size: int = 9, family: str | None = None):
        fam = family or FONT_FAMILY
        f = tkfont.Font(family=fam, size=size, weight="bold")
        tw, th = f.measure(text), f.metrics("linespace")
        px, py = 11, 5
        W, H = tw + 2 * px, th + 2 * py
        super().__init__(parent, width=W, height=H, bg=bgcolor,
                         highlightthickness=0, bd=0)
        fg, fill = _PILL_COLORS.get(kind, _PILL_COLORS["muted"])
        _round_rect(self, 1, 1, W - 1, H - 1, H // 2, fill=fill, outline="")
        self.create_text(W // 2, H // 2 + 1, text=text, fill=fg, font=f)


class StepBadge(tk.Canvas):
    """Huy hiệu số tròn cho stepper (①②③ dạng vòng tròn màu)."""
    def __init__(self, parent, number, color: str = NAVY, bgcolor: str = CARD,
                 d: int = 34, family: str | None = None):
        super().__init__(parent, width=d, height=d, bg=bgcolor,
                         highlightthickness=0, bd=0)
        self.create_oval(2, 2, d - 2, d - 2, fill=color, outline="")
        self.create_text(d // 2, d // 2 + 1, text=str(number), fill="#FFFFFF",
                         font=(family or FONT_FAMILY, 13, "bold"))
