"""
Interfaz gráfica moderna para el procesador de órdenes Returns-TANET.
Versión actualizada:
- Panel derecho con confirmación de ubicaciones siempre visible
- Tarjeta única para Retiro y Entrega
- Soporte para logo.png en login y pantalla principal
"""

import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime

import customtkinter as ctk
from PIL import Image, ImageDraw, ImageOps

from src.tanet import Tanet
from src.excel_manager import ExcelManager
from src.services import OrderService
from src.models import SiteInfo, MatchStatus
from src.exporters import export_orders_to_excel


# -----------------------------
# Config visual
# -----------------------------
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

APP_BG = "#F5F7FB"
CARD_BG = "#FFFFFF"
CARD_SELECTED = "#F3F8FF"
BORDER = "#E5EAF2"
TEXT = "#0F172A"
MUTED = "#64748B"
PRIMARY = "#2563EB"
PRIMARY_HOVER = "#1D4ED8"
SUCCESS = "#16A34A"
SUCCESS_BG = "#EAF8EF"
DANGER = "#DC2626"
DANGER_BG = "#FDECEC"
WARNING = "#D97706"
WARNING_BG = "#FFF7E8"

LOGO_FILENAME = "logo.png"


# -----------------------------
# Helpers
# -----------------------------
def safe_text(value, fallback="-"):
    if value is None:
        return fallback
    text = str(value).strip()
    return text if text else fallback


def get_logo_path():
    candidates = [
        os.path.join(os.getcwd(), LOGO_FILENAME),
        os.path.join(os.path.dirname(os.path.abspath(__file__)), LOGO_FILENAME),
    ]
    for path in candidates:
        if os.path.exists(path):
            return path
    return None


def try_parse_datetime(value):
    if value is None:
        return None

    if isinstance(value, datetime):
        return value

    text = str(value).strip()
    if not text or text == "-":
        return None

    patterns = [
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%Y-%m-%d",
        "%d/%m/%Y",
    ]

    for fmt in patterns:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            pass

    try:
        return datetime.fromisoformat(text.replace("T", " "))
    except Exception:
        return None


def try_parse_time(value):
    if value is None:
        return None

    text = str(value).strip()
    if not text or text == "-":
        return None

    patterns = ["%H:%M:%S", "%H:%M"]
    for fmt in patterns:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            pass
    return None


def format_single_slot(start_value, end_value, fallback_date=None, fallback_start_time=None, fallback_end_time=None):
    start_dt = try_parse_datetime(start_value)
    end_dt = try_parse_datetime(end_value)

    if start_dt and end_dt:
        if start_dt.date() == end_dt.date():
            return f"{start_dt:%d/%m/%Y %H:%M} - {end_dt:%H:%M}"
        return f"{start_dt:%d/%m/%Y %H:%M} - {end_dt:%d/%m/%Y %H:%M}"

    date_dt = try_parse_datetime(fallback_date)
    start_t = try_parse_time(fallback_start_time)
    end_t = try_parse_time(fallback_end_time)

    if date_dt and start_t and end_t:
        return f"{date_dt:%d/%m/%Y} {start_t:%H:%M} - {end_t:%H:%M}"

    if start_dt:
        return f"{start_dt:%d/%m/%Y %H:%M}"

    return "-"


def get_retiro_text(extra):
    return format_single_slot(
        extra.get("retiro_desde"),
        extra.get("retiro_hasta"),
        fallback_date=extra.get("fecha_retiro"),
        fallback_start_time=extra.get("hora_retiro_desde"),
        fallback_end_time=extra.get("hora_retiro_hasta"),
    )


def get_entrega_text(extra):
    return format_single_slot(
        extra.get("entrega_desde"),
        extra.get("entrega_hasta"),
    )


def get_order_status_meta(order):
    if order.is_confirmed:
        return "Confirmada", SUCCESS_BG, SUCCESS
    if order.is_discarded:
        return "Descartada", DANGER_BG, DANGER
    if order.match_status == MatchStatus.NO_MATCH:
        return "Sin matches", WARNING_BG, WARNING
    if order.match_status == MatchStatus.SINGLE_MATCH:
        return "1 match", "#EAF2FF", PRIMARY
    return f"{order.match_count} matches", "#EEF2FF", "#4338CA"


def get_order_preview(order):
    extra = order.extra_data or {}
    referencia = safe_text(extra.get("referencia"))
    retiro = get_retiro_text(extra)
    if referencia != "-":
        return f"Referencia: {referencia}"
    if retiro != "-":
        return f"Retiro: {retiro}"
    return "Sin detalle adicional"


def bind_click_recursive(widget, callback):
    widget.bind("<Button-1>", callback)
    for child in widget.winfo_children():
        bind_click_recursive(child, callback)


# -----------------------------
# Componentes visuales
# -----------------------------
class LogoWidget(ctk.CTkFrame):
    def __init__(self, parent, size=72, show_caption=False):
        super().__init__(parent, fg_color="transparent")
        self.logo_image = None

        corner_radius = 18

        logo_path = get_logo_path()
        if logo_path:
            try:
                img = Image.open(logo_path).convert("RGBA")

                # Hace que el logo llene todo el cuadro sin deformarse
                img = ImageOps.fit(img, (size, size), method=Image.LANCZOS)

                # Máscara con bordes redondeados
                mask = Image.new("L", (size, size), 0)
                draw = ImageDraw.Draw(mask)
                draw.rounded_rectangle((0, 0, size, size), radius=corner_radius, fill=255)

                rounded = Image.new("RGBA", (size, size), (0, 0, 0, 0))
                rounded.paste(img, (0, 0), mask=mask)

                self.logo_image = ctk.CTkImage(light_image=rounded, dark_image=rounded, size=(size, size))

                ctk.CTkLabel(
                    self,
                    text="",
                    image=self.logo_image
                ).pack()
            except Exception:
                fallback = ctk.CTkFrame(
                    self,
                    width=size,
                    height=size,
                    fg_color="#F8FAFC",
                    corner_radius=corner_radius,
                    border_width=1,
                    border_color=BORDER
                )
                fallback.pack()
                fallback.pack_propagate(False)

                ctk.CTkLabel(
                    fallback,
                    text="LOGO",
                    text_color=MUTED,
                    font=ctk.CTkFont(size=14, weight="bold")
                ).pack(expand=True)
        else:
            fallback = ctk.CTkFrame(
                self,
                width=size,
                height=size,
                fg_color="#F8FAFC",
                corner_radius=corner_radius,
                border_width=1,
                border_color=BORDER
            )
            fallback.pack()
            fallback.pack_propagate(False)

            ctk.CTkLabel(
                fallback,
                text="logo.png",
                text_color=MUTED,
                font=ctk.CTkFont(size=12, weight="bold")
            ).pack(expand=True)


class BrandBlock(ctk.CTkFrame):
    def __init__(self, parent, compact=False):
        super().__init__(parent, fg_color="transparent")

        wrapper = ctk.CTkFrame(self, fg_color="transparent")
        wrapper.pack(anchor="w")

        self.logo = LogoWidget(wrapper, size=56 if compact else 90, show_caption=not compact)
        self.logo.pack(side="left", padx=(0, 14))

        texts = ctk.CTkFrame(wrapper, fg_color="transparent")
        texts.pack(side="left", anchor="center")

        ctk.CTkLabel(
            texts,
            text="RETURNS · TANET",
            font=ctk.CTkFont(size=18 if compact else 26, weight="bold"),
            text_color=TEXT
        ).pack(anchor="w")


class StatCard(ctk.CTkFrame):
    def __init__(self, parent, title, value="0"):
        super().__init__(parent, fg_color=CARD_BG, corner_radius=18, border_width=1, border_color=BORDER)
        self.grid_columnconfigure(0, weight=1)

        self.title_label = ctk.CTkLabel(
            self,
            text=title,
            font=ctk.CTkFont(size=12),
            text_color=MUTED
        )
        self.title_label.grid(row=0, column=0, padx=14, pady=(12, 2), sticky="w")

        self.value_label = ctk.CTkLabel(
            self,
            text=value,
            font=ctk.CTkFont(size=24, weight="bold"),
            text_color=TEXT
        )
        self.value_label.grid(row=1, column=0, padx=14, pady=(0, 12), sticky="w")

    def set_value(self, value):
        self.value_label.configure(text=str(value))


class InfoFieldCard(ctk.CTkFrame):
    def __init__(self, parent, label, value, big=False):
        super().__init__(parent, fg_color="#F8FAFC", corner_radius=14, border_width=1, border_color=BORDER)

        ctk.CTkLabel(
            self,
            text=label,
            font=ctk.CTkFont(size=11),
            text_color=MUTED
        ).pack(anchor="w", padx=12, pady=(10, 2))

        ctk.CTkLabel(
            self,
            text=safe_text(value),
            font=ctk.CTkFont(size=15 if big else 13, weight="bold"),
            text_color=TEXT,
            justify="left",
            wraplength=330
        ).pack(anchor="w", padx=12, pady=(0, 12), fill="x")


class OrderCard(ctk.CTkFrame):
    def __init__(self, parent, order, on_select, selected=False):
        super().__init__(
            parent,
            fg_color=CARD_SELECTED if selected else CARD_BG,
            corner_radius=18,
            border_width=2 if selected else 1,
            border_color=PRIMARY if selected else BORDER
        )
        self.order = order
        self.on_select = on_select

        self.grid_columnconfigure(0, weight=1)

        status_text, status_bg, status_fg = get_order_status_meta(order)

        header = ctk.CTkFrame(self, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=14, pady=(14, 8))
        header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            header,
            text=f"Orden #{order.order_id}",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=TEXT
        ).grid(row=0, column=0, sticky="w")

        ctk.CTkLabel(
            header,
            text=status_text,
            fg_color=status_bg,
            text_color=status_fg,
            corner_radius=999,
            padx=10,
            pady=4,
            font=ctk.CTkFont(size=11, weight="bold")
        ).grid(row=0, column=1, sticky="e")

        ctk.CTkLabel(
            self,
            text=f"{safe_text(order.protocol)}  •  Sitio {safe_text(order.site_number)}",
            font=ctk.CTkFont(size=13),
            text_color=TEXT
        ).grid(row=1, column=0, sticky="w", padx=14)

        ctk.CTkLabel(
            self,
            text=get_order_preview(order),
            font=ctk.CTkFont(size=12),
            text_color=MUTED,
            justify="left",
            wraplength=360
        ).grid(row=2, column=0, sticky="ew", padx=14, pady=(8, 14))

        bind_click_recursive(self, lambda e: self.on_select(self.order.order_id))

    def set_selected(self, selected):
        self.configure(
            fg_color=CARD_SELECTED if selected else CARD_BG,
            border_width=2 if selected else 1,
            border_color=PRIMARY if selected else BORDER
        )


class MatchCard(ctk.CTkFrame):
    def __init__(self, parent, match, variable, on_select, selected=False, disabled=False):
        super().__init__(
            parent,
            fg_color=CARD_SELECTED if selected else CARD_BG,
            corner_radius=18,
            border_width=2 if selected else 1,
            border_color=PRIMARY if selected else BORDER
        )
        self.match = match
        self.variable = variable
        self.on_select = on_select
        self.disabled = disabled

        self.grid_columnconfigure(1, weight=1)

        self.radio = ctk.CTkRadioButton(
            self,
            text="",
            variable=self.variable,
            value=str(match.match_index),
            command=lambda: self.on_select(str(match.match_index)),
            state="disabled" if disabled else "normal"
        )
        self.radio.grid(row=0, column=0, rowspan=3, padx=(14, 8), pady=14, sticky="n")

        ctk.CTkLabel(
            self,
            text=f"{safe_text(match.site_info.protocol)}  •  Sitio {safe_text(match.site_info.site_number)}",
            font=ctk.CTkFont(size=15, weight="bold"),
            text_color=TEXT
        ).grid(row=0, column=1, sticky="w", padx=(0, 14), pady=(14, 4))

        ctk.CTkLabel(
            self,
            text=safe_text(match.site_info.nomdomicilio),
            font=ctk.CTkFont(size=13),
            text_color=TEXT,
            justify="left",
            wraplength=450
        ).grid(row=1, column=1, sticky="w", padx=(0, 14))

        ctk.CTkLabel(
            self,
            text=(
                f"{safe_text(match.site_info.calle)}  •  "
                f"{safe_text(match.site_info.localidad)}  •  "
                f"{safe_text(match.site_info.nomprovincia)}  •  "
                f"{safe_text(match.site_info.nompais)}"
            ),
            font=ctk.CTkFont(size=12),
            text_color=MUTED,
            justify="left",
            wraplength=450
        ).grid(row=2, column=1, sticky="w", padx=(0, 14), pady=(4, 14))

        if not disabled:
            bind_click_recursive(self, lambda e: self._select_from_card())

    def _select_from_card(self):
        self.variable.set(str(self.match.match_index))
        self.on_select(str(self.match.match_index))

    def set_selected(self, selected):
        self.configure(
            fg_color=CARD_SELECTED if selected else CARD_BG,
            border_width=2 if selected else 1,
            border_color=PRIMARY if selected else BORDER
        )


# -----------------------------
# Login
# -----------------------------
class LoginFrame(ctk.CTkFrame):
    def __init__(self, parent, on_login_success):
        super().__init__(parent, fg_color="transparent")
        self.on_login_success = on_login_success
        self._build()

    def _build(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        card = ctk.CTkFrame(
            self,
            width=460,
            fg_color=CARD_BG,
            corner_radius=24,
            border_width=1,
            border_color=BORDER
        )
        card.grid(row=0, column=0, padx=24, pady=24)
        card.grid_columnconfigure(0, weight=1)

        self.brand = BrandBlock(card, compact=False)
        self.brand.grid(row=0, column=0, padx=28, pady=(28, 16), sticky="w")

        ctk.CTkLabel(
            card,
            text="Inicio de sesión",
            font=ctk.CTkFont(size=15),
            text_color=MUTED
        ).grid(row=1, column=0, padx=28, pady=(0, 18), sticky="w")

        ctk.CTkLabel(card, text="Usuario", text_color=TEXT).grid(row=2, column=0, padx=28, pady=(0, 6), sticky="w")
        self.username_entry = ctk.CTkEntry(card, height=42, placeholder_text="Ingresá tu usuario")
        self.username_entry.grid(row=3, column=0, padx=28, pady=(0, 14), sticky="ew")

        ctk.CTkLabel(card, text="Contraseña", text_color=TEXT).grid(row=4, column=0, padx=28, pady=(0, 6), sticky="w")
        self.password_entry = ctk.CTkEntry(card, height=42, placeholder_text="Ingresá tu contraseña", show="•")
        self.password_entry.grid(row=5, column=0, padx=28, pady=(0, 18), sticky="ew")

        self.login_btn = ctk.CTkButton(
            card,
            text="Conectar",
            height=44,
            corner_radius=14,
            fg_color=PRIMARY,
            hover_color=PRIMARY_HOVER,
            command=self._do_login
        )
        self.login_btn.grid(row=6, column=0, padx=28, pady=(0, 12), sticky="ew")

        self.progress = ctk.CTkProgressBar(card, mode="indeterminate")
        self.progress.grid(row=7, column=0, padx=28, pady=(0, 10), sticky="ew")
        self.progress.grid_remove()

        self.status_label = ctk.CTkLabel(
            card,
            text="",
            text_color=MUTED,
            justify="left",
            wraplength=360
        )
        self.status_label.grid(row=8, column=0, padx=28, pady=(0, 24), sticky="w")

        self.password_entry.bind("<Return>", lambda e: self._do_login())

    def _set_status(self, text, color=MUTED):
        self.status_label.configure(text=text, text_color=color)

    def _do_login(self):
        username = self.username_entry.get().strip()
        password = self.password_entry.get().strip()

        if not username or not password:
            messagebox.showwarning("Aviso", "Ingresá usuario y contraseña.")
            return

        self.login_btn.configure(state="disabled")
        self.progress.grid()
        self.progress.start()
        self._set_status("Conectando con TANET...", PRIMARY)

        thread = threading.Thread(target=self._login_thread, args=(username, password), daemon=True)
        thread.start()

    def _login_thread(self, username, password):
        try:
            tanet = Tanet()
            tanet.login(username, password)
            df = tanet.load_site_data()
            sites = [SiteInfo.from_dict(row.to_dict()) for _, row in df.iterrows()]
            self.after(0, lambda: self._login_success(sites))
        except Exception as e:
            self.after(0, lambda: self._login_error(str(e)))

    def _login_success(self, sites):
        self.progress.stop()
        self.progress.grid_remove()
        self._set_status(f"Conectado correctamente. {len(sites)} registros cargados.", SUCCESS)
        self.on_login_success(sites)

    def _login_error(self, error):
        self.progress.stop()
        self.progress.grid_remove()
        self.login_btn.configure(state="normal")
        self._set_status(f"Error: {error}", DANGER)
        messagebox.showerror("Error", f"No se pudo conectar:\n{error}")


# -----------------------------
# Pantalla principal
# -----------------------------
class OrdersFrame(ctk.CTkFrame):
    def __init__(self, parent, sites):
        super().__init__(parent, fg_color="transparent")
        self.sites = sites
        self.service = None
        self.excel_manager = ExcelManager()

        self.current_order = None
        self.current_order_id = None
        self.order_cards = {}
        self.match_cards = {}
        self.selected_match_var = tk.StringVar(value="")

        self.search_var = tk.StringVar(value="")
        self.filter_var = tk.StringVar(value="Pendientes")

        self._build()
        self._render_empty_state()

    def _build(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

        # Header superior
        top = ctk.CTkFrame(self, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", padx=18, pady=(10, 6))
        top.grid_columnconfigure(1, weight=1)

        self.brand = BrandBlock(top, compact=True)
        self.brand.grid(row=0, column=0, sticky="w")

        actions = ctk.CTkFrame(top, fg_color="transparent")
        actions.grid(row=0, column=2, sticky="e")

        self.open_excel_btn = ctk.CTkButton(
            actions,
            text="Abrir Excel de órdenes",
            height=42,
            corner_radius=14,
            fg_color=PRIMARY,
            hover_color=PRIMARY_HOVER,
            command=self._open_excel
        )
        self.open_excel_btn.pack(side="left", padx=(0, 10))

        self.export_btn = ctk.CTkButton(
            actions,
            text="Exportar confirmadas",
            height=42,
            corner_radius=14,
            fg_color="#0F766E",
            hover_color="#115E59",
            command=self._export_orders,
            state="disabled"
        )
        self.export_btn.pack(side="left")

        self.status_pill = ctk.CTkLabel(
            self,
            text="Esperando órdenes...",
            fg_color="#EDF2F7",
            text_color=MUTED,
            corner_radius=999,
            padx=12,
            pady=8,
            font=ctk.CTkFont(size=12, weight="bold")
        )
        self.status_pill.grid(row=1, column=0, sticky="w", padx=18, pady=(0, 6))

        # Stats
        stats = ctk.CTkFrame(self, fg_color="transparent")
        stats.grid(row=2, column=0, sticky="ew", padx=18, pady=(0, 8))
        for i in range(4):
            stats.grid_columnconfigure(i, weight=1)

        self.total_card = StatCard(stats, "Total")
        self.total_card.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        self.pending_card = StatCard(stats, "Pendientes")
        self.pending_card.grid(row=0, column=1, sticky="ew", padx=8)

        self.confirmed_card = StatCard(stats, "Confirmadas")
        self.confirmed_card.grid(row=0, column=2, sticky="ew", padx=8)

        self.discarded_card = StatCard(stats, "Descartadas")
        self.discarded_card.grid(row=0, column=3, sticky="ew", padx=(8, 0))

        # Cuerpo
        body = ctk.CTkFrame(self, fg_color="transparent")
        body.grid(row=3, column=0, sticky="nsew", padx=18, pady=(0, 12))
        body.grid_columnconfigure(0, weight=1, minsize=430)
        body.grid_columnconfigure(1, weight=2, minsize=780)
        body.grid_rowconfigure(0, weight=1)

        # Panel izquierdo: órdenes
        self.left_panel = ctk.CTkFrame(body, fg_color=CARD_BG, corner_radius=24, border_width=1, border_color=BORDER)
        self.left_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        self.left_panel.grid_columnconfigure(0, weight=1)
        self.left_panel.grid_rowconfigure(3, weight=1)

        left_header = ctk.CTkFrame(self.left_panel, fg_color="transparent")
        left_header.grid(row=0, column=0, sticky="ew", padx=18, pady=(18, 10))
        left_header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            left_header,
            text="Órdenes",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color=TEXT
        ).grid(row=0, column=0, sticky="w")

        self.results_label = ctk.CTkLabel(
            left_header,
            text="0 resultados",
            font=ctk.CTkFont(size=12),
            text_color=MUTED
        )
        self.results_label.grid(row=0, column=1, sticky="e")

        self.search_entry = ctk.CTkEntry(
            self.left_panel,
            height=40,
            placeholder_text="Buscar por protocolo, sitio o referencia",
            textvariable=self.search_var
        )
        self.search_entry.grid(row=1, column=0, sticky="ew", padx=18, pady=(0, 12))
        self.search_entry.bind("<KeyRelease>", lambda e: self._refresh_orders_list())

        self.filter_segment = ctk.CTkSegmentedButton(
            self.left_panel,
            values=["Pendientes", "Todas", "Confirmadas", "Descartadas", "Sin match"],
            variable=self.filter_var,
            command=lambda _: self._refresh_orders_list()
        )
        self.filter_segment.grid(row=2, column=0, sticky="ew", padx=18, pady=(0, 12))
        self.filter_segment.set("Pendientes")

        self.orders_scroll = ctk.CTkScrollableFrame(self.left_panel, fg_color="transparent", corner_radius=0)
        self.orders_scroll.grid(row=3, column=0, sticky="nsew", padx=14, pady=(0, 14))

        # Panel derecho: detalle + confirmación visible
        self.right_panel = ctk.CTkFrame(body, fg_color=CARD_BG, corner_radius=24, border_width=1, border_color=BORDER)
        self.right_panel.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        self.right_panel.grid_columnconfigure(0, weight=1)
        self.right_panel.grid_rowconfigure(1, weight=1)

        self.detail_header = ctk.CTkFrame(self.right_panel, fg_color="transparent")
        self.detail_header.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 12))
        self.detail_header.grid_columnconfigure(0, weight=1)

        self.order_title_label = ctk.CTkLabel(
            self.detail_header,
            text="Seleccioná una orden",
            font=ctk.CTkFont(size=22, weight="bold"),
            text_color=TEXT
        )
        self.order_title_label.grid(row=0, column=0, sticky="w")

        self.order_status_badge = ctk.CTkLabel(
            self.detail_header,
            text="",
            fg_color="#EDF2F7",
            text_color=MUTED,
            corner_radius=999,
            padx=12,
            pady=8,
            font=ctk.CTkFont(size=12, weight="bold")
        )
        self.order_status_badge.grid(row=0, column=1, sticky="e")

        self.order_meta_label = ctk.CTkLabel(
            self.detail_header,
            text="",
            font=ctk.CTkFont(size=13),
            text_color=MUTED
        )
        self.order_meta_label.grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))

        # Split del panel derecho
        self.right_content = ctk.CTkFrame(self.right_panel, fg_color="transparent")
        self.right_content.grid(row=1, column=0, sticky="nsew", padx=20, pady=(0, 20))
        self.right_content.grid_columnconfigure(0, weight=1, minsize=360)
        self.right_content.grid_columnconfigure(1, weight=1, minsize=400)
        self.right_content.grid_rowconfigure(0, weight=1)

        # Bloque datos
        self.info_card = ctk.CTkScrollableFrame(
            self.right_content,
            fg_color="#F8FAFC",
            corner_radius=20,
            border_width=1,
            border_color=BORDER
        )
        self.info_card.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        self.info_card.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            self.info_card,
            text="Datos de la orden",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=TEXT
        ).grid(row=0, column=0, sticky="w", padx=16, pady=(16, 10))

        self.info_grid = ctk.CTkFrame(self.info_card, fg_color="transparent")
        self.info_grid.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 16))
        self.info_grid.grid_columnconfigure(0, weight=1)
        self.info_grid.grid_columnconfigure(1, weight=1)

        # Bloque confirmar ubicación
        self.confirm_card = ctk.CTkFrame(
            self.right_content,
            fg_color="#F8FAFC",
            corner_radius=20,
            border_width=1,
            border_color=BORDER
        )
        self.confirm_card.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        self.confirm_card.grid_columnconfigure(0, weight=1)
        self.confirm_card.grid_rowconfigure(2, weight=1)

        confirm_header = ctk.CTkFrame(self.confirm_card, fg_color="transparent")
        confirm_header.grid(row=0, column=0, sticky="ew", padx=16, pady=(16, 10))
        confirm_header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            confirm_header,
            text="Confirmar ubicación",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=TEXT
        ).grid(row=0, column=0, sticky="w")

        self.matches_count_label = ctk.CTkLabel(
            confirm_header,
            text="",
            font=ctk.CTkFont(size=12),
            text_color=MUTED
        )
        self.matches_count_label.grid(row=0, column=1, sticky="e")

        self.selection_hint_label = ctk.CTkLabel(
            self.confirm_card,
            text="Seleccioná una ubicación TANET para confirmar esta orden.",
            font=ctk.CTkFont(size=12),
            text_color=MUTED,
            justify="left",
            wraplength=420
        )
        self.selection_hint_label.grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 10))

        self.matches_scroll = ctk.CTkScrollableFrame(
            self.confirm_card,
            fg_color="transparent",
            corner_radius=0
        )
        self.matches_scroll.grid(row=2, column=0, sticky="nsew", padx=12, pady=(0, 12))

        actions = ctk.CTkFrame(self.confirm_card, fg_color="transparent")
        actions.grid(row=3, column=0, sticky="ew", padx=16, pady=(0, 16))
        actions.grid_columnconfigure((0, 1, 2), weight=1)

        self.confirm_btn = ctk.CTkButton(
            actions,
            text="Confirmar selección",
            height=44,
            corner_radius=14,
            fg_color=SUCCESS,
            hover_color="#15803D",
            command=self._confirm_order,
            state="disabled"
        )
        self.confirm_btn.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        self.discard_btn = ctk.CTkButton(
            actions,
            text="Descartar orden",
            height=44,
            corner_radius=14,
            fg_color=DANGER,
            hover_color="#B91C1C",
            command=self._discard_order,
            state="disabled"
        )
        self.discard_btn.grid(row=0, column=1, sticky="ew", padx=8)

        self.skip_btn = ctk.CTkButton(
            actions,
            text="Saltar",
            height=44,
            corner_radius=14,
            fg_color="#475569",
            hover_color="#334155",
            command=self._skip_order,
            state="disabled"
        )
        self.skip_btn.grid(row=0, column=2, sticky="ew", padx=(8, 0))

    # -----------------------------
    # Estado visual
    # -----------------------------
    def _set_status(self, text, kind="neutral"):
        palette = {
            "neutral": ("#EDF2F7", MUTED),
            "info": ("#EAF2FF", PRIMARY),
            "success": (SUCCESS_BG, SUCCESS),
            "warning": (WARNING_BG, WARNING),
            "danger": (DANGER_BG, DANGER),
        }
        bg, fg = palette.get(kind, palette["neutral"])
        self.status_pill.configure(text=text, fg_color=bg, text_color=fg)

    def _render_empty_state(self):
        self.total_card.set_value(0)
        self.pending_card.set_value(0)
        self.confirmed_card.set_value(0)
        self.discarded_card.set_value(0)
        self.results_label.configure(text="0 resultados")
        self._clear_orders_list("Cargá un Excel para comenzar.")
        self._clear_detail_panel()

    def _clear_orders_list(self, message):
        for widget in self.orders_scroll.winfo_children():
            widget.destroy()

        ctk.CTkLabel(
            self.orders_scroll,
            text=message,
            text_color=MUTED,
            font=ctk.CTkFont(size=14),
            justify="center",
            wraplength=330
        ).pack(fill="x", padx=16, pady=40)

    def _clear_detail_panel(self):
        self.current_order = None
        self.current_order_id = None
        self.selected_match_var.set("")

        self.order_title_label.configure(text="Seleccioná una orden")
        self.order_meta_label.configure(text="Vas a ver el detalle de la orden y a la derecha las ubicaciones para confirmar.")
        self.order_status_badge.configure(text="", fg_color="#EDF2F7", text_color=MUTED)
        self.matches_count_label.configure(text="")
        self.selection_hint_label.configure(text="Seleccioná una ubicación TANET para confirmar esta orden.")

        for widget in self.info_grid.winfo_children():
            widget.destroy()

        for widget in self.matches_scroll.winfo_children():
            widget.destroy()

        ctk.CTkLabel(
            self.matches_scroll,
            text="No hay ubicaciones para mostrar.",
            text_color=MUTED,
            font=ctk.CTkFont(size=14)
        ).pack(fill="x", padx=16, pady=28)

        self.confirm_btn.configure(state="disabled")
        self.discard_btn.configure(state="disabled")
        self.skip_btn.configure(state="disabled")

    def _update_summary(self):
        if not self.service:
            self.total_card.set_value(0)
            self.pending_card.set_value(0)
            self.confirmed_card.set_value(0)
            self.discarded_card.set_value(0)
            self.export_btn.configure(state="disabled")
            return

        summary = self.service.get_summary()
        self.total_card.set_value(summary.total_orders)
        self.pending_card.set_value(summary.pending)
        self.confirmed_card.set_value(summary.confirmed)
        self.discarded_card.set_value(summary.discarded)
        self.export_btn.configure(state="normal" if summary.confirmed > 0 else "disabled")

    # -----------------------------
    # Órdenes
    # -----------------------------
    def _get_filtered_orders(self):
        if not self.service:
            return []

        orders = self.service.orders
        filter_value = self.filter_var.get()
        query = self.search_var.get().strip().lower()

        if filter_value == "Pendientes":
            orders = [o for o in orders if o.is_pending]
        elif filter_value == "Confirmadas":
            orders = [o for o in orders if o.is_confirmed]
        elif filter_value == "Descartadas":
            orders = [o for o in orders if o.is_discarded]
        elif filter_value == "Sin match":
            orders = [o for o in orders if o.match_status == MatchStatus.NO_MATCH]

        if query:
            filtered = []
            for o in orders:
                reference = safe_text(o.extra_data.get("referencia") if o.extra_data else "")
                haystack = " ".join([
                    str(o.order_id),
                    safe_text(o.protocol, ""),
                    safe_text(o.site_number, ""),
                    reference
                ]).lower()
                if query in haystack:
                    filtered.append(o)
            orders = filtered

        return orders

    def _refresh_orders_list(self, select_order_id=None):
        for widget in self.orders_scroll.winfo_children():
            widget.destroy()

        self.order_cards = {}
        filtered_orders = self._get_filtered_orders()
        self.results_label.configure(text=f"{len(filtered_orders)} resultados")

        if not filtered_orders:
            self._clear_orders_list("No hay órdenes para ese filtro.")
            self._update_summary()
            self._clear_detail_panel()
            return

        for order in filtered_orders:
            card = OrderCard(
                self.orders_scroll,
                order=order,
                on_select=self._select_order,
                selected=(order.order_id == self.current_order_id)
            )
            card.pack(fill="x", padx=4, pady=(0, 10))
            self.order_cards[order.order_id] = card

        available_ids = [o.order_id for o in filtered_orders]
        target_id = select_order_id if select_order_id in available_ids else self.current_order_id

        if target_id not in available_ids:
            pending_order = next((o for o in filtered_orders if o.is_pending), None)
            target_id = pending_order.order_id if pending_order else filtered_orders[0].order_id

        self._select_order(target_id)
        self._update_summary()

    def _select_order(self, order_id):
        if not self.service:
            return

        order = next((o for o in self.service.orders if o.order_id == order_id), None)
        if not order:
            return

        self.current_order = order
        self.current_order_id = order_id

        for oid, card in self.order_cards.items():
            card.set_selected(oid == order_id)

        self._show_order_details(order)

    def _show_order_details(self, order):
        extra = order.extra_data or {}
        status_text, status_bg, status_fg = get_order_status_meta(order)

        self.order_title_label.configure(text=f"Orden #{order.order_id}")
        self.order_meta_label.configure(
            text=f"{safe_text(order.protocol)}  •  Sitio {safe_text(order.site_number)}  •  Ref: {safe_text(extra.get('referencia'))}"
        )
        self.order_status_badge.configure(text=status_text, fg_color=status_bg, text_color=status_fg)
        self.matches_count_label.configure(text=f"{order.match_count} ubicaciones")

        self._render_info_fields(order)
        self._render_matches(order)

        if order.is_confirmed or order.is_discarded:
            self.confirm_btn.configure(state="disabled")
            self.discard_btn.configure(state="disabled")
            self.skip_btn.configure(state="disabled")
        else:
            self.discard_btn.configure(state="normal")
            self.skip_btn.configure(state="normal")
            self.confirm_btn.configure(state="normal" if order.matches else "disabled")

    def _render_info_fields(self, order):
        for widget in self.info_grid.winfo_children():
            widget.destroy()

        extra = order.extra_data or {}

        fields = [
            ("Referencia", extra.get("referencia")),
            ("Email", extra.get("email")),
            ("Retiro", get_retiro_text(extra)),
            ("Entrega", get_entrega_text(extra)),
            ("Cantidad de cajas", extra.get("cantidad_cajas")),
            ("Sector", extra.get("sector")),
            ("Contactos", extra.get("contactos")),
            ("Tipo de material", extra.get("tipo_material")),
        ]

        row = 0
        col = 0
        for label, value in fields:
            big = label in ("Retiro", "Entrega")
            card = InfoFieldCard(self.info_grid, label, value, big=big)
            card.grid(row=row, column=col, sticky="ew", padx=6, pady=6)
            col += 1
            if col > 1:
                col = 0
                row += 1

        comments = safe_text(extra.get("comentarios"))
        if comments != "-":
            comments_card = InfoFieldCard(self.info_grid, "Comentarios", comments, big=False)
            comments_card.grid(row=row + 1, column=0, columnspan=2, sticky="ew", padx=6, pady=6)

    def _render_matches(self, order):
        for widget in self.matches_scroll.winfo_children():
            widget.destroy()

        self.match_cards = {}

        if order.is_confirmed or order.is_discarded:
            self.selected_match_var.set("")
            self.selection_hint_label.configure(text="Esta orden ya fue procesada.")
        elif len(order.matches) == 1:
            first = str(order.matches[0].match_index)
            self.selected_match_var.set(first)
            self.selection_hint_label.configure(text="Se detectó una única ubicación posible.")
        else:
            self.selected_match_var.set("")
            self.selection_hint_label.configure(text="Seleccioná una ubicación TANET para confirmar esta orden.")

        if not order.matches:
            self.selection_hint_label.configure(text="Esta orden no tiene ubicaciones sugeridas.")
            ctk.CTkLabel(
                self.matches_scroll,
                text="No se encontraron ubicaciones para esta orden.",
                text_color=MUTED,
                font=ctk.CTkFont(size=14)
            ).pack(fill="x", padx=16, pady=28)
            return

        for match in order.matches:
            selected = self.selected_match_var.get() == str(match.match_index)
            card = MatchCard(
                self.matches_scroll,
                match=match,
                variable=self.selected_match_var,
                on_select=self._on_match_selected,
                selected=selected,
                disabled=(order.is_confirmed or order.is_discarded)
            )
            card.pack(fill="x", padx=4, pady=(0, 10))
            self.match_cards[str(match.match_index)] = card

        if self.selected_match_var.get():
            self._on_match_selected(self.selected_match_var.get())

    def _on_match_selected(self, selected_value):
        for match_index, card in self.match_cards.items():
            card.set_selected(match_index == selected_value)

        if not self.current_order:
            return

        selected_match = next(
            (m for m in self.current_order.matches if str(m.match_index) == str(selected_value)),
            None
        )

        if selected_match:
            self.selection_hint_label.configure(
                text=(
                    f"Ubicación seleccionada: "
                    f"{safe_text(selected_match.site_info.protocol)} / "
                    f"Sitio {safe_text(selected_match.site_info.site_number)}"
                )
            )

    # -----------------------------
    # Flujo Excel
    # -----------------------------
    def _open_excel(self):
        self.open_excel_btn.configure(state="disabled")
        self._set_status("Creando plantilla Excel...", "info")

        try:
            self.excel_manager.create_template()
            self.excel_manager.open_excel()
            self._set_status("Completá el Excel y cerralo para continuar.", "warning")
            thread = threading.Thread(target=self._wait_excel_close, daemon=True)
            thread.start()
        except Exception as e:
            self.open_excel_btn.configure(state="normal")
            self._set_status("No se pudo abrir el Excel.", "danger")
            messagebox.showerror("Error", f"No se pudo abrir el Excel:\n{e}")

    def _wait_excel_close(self):
        closed = self.excel_manager.wait_for_excel_close()
        self.after(0, lambda: self._on_excel_closed(closed))

    def _on_excel_closed(self, closed):
        if not closed:
            self.open_excel_btn.configure(state="normal")
            self._set_status("Carga cancelada.", "danger")
            return

        try:
            orders = self.excel_manager.read_orders_as_models()
        except Exception as e:
            self.open_excel_btn.configure(state="normal")
            self._set_status("Error leyendo el Excel.", "danger")
            messagebox.showerror("Error", f"Error al leer órdenes:\n{e}")
            return

        if not orders:
            self.open_excel_btn.configure(state="normal")
            self._set_status("No se encontraron órdenes válidas.", "warning")
            messagebox.showwarning("Aviso", "No se encontraron órdenes válidas.")
            return

        try:
            self._set_status("Procesando órdenes...", "info")
            self.service = OrderService(self.sites)
            self.service.add_orders(orders)
            self.service.process_all_orders()

            self._set_status(f"{len(orders)} órdenes cargadas correctamente.", "success")
            self.open_excel_btn.configure(state="normal")
            self._refresh_orders_list()
        except Exception as e:
            self.open_excel_btn.configure(state="normal")
            self._set_status("Error procesando órdenes.", "danger")
            messagebox.showerror("Error", f"No se pudieron procesar las órdenes:\n{e}")

    # -----------------------------
    # Acciones
    # -----------------------------
    def _confirm_order(self):
        if not self.current_order or not self.service:
            return

        order = self.current_order
        selected = self.selected_match_var.get().strip()

        if not selected:
            if len(order.matches) == 1:
                match_index = order.matches[0].match_index
            else:
                messagebox.showwarning("Aviso", "Seleccioná una ubicación.")
                return
        else:
            match_index = int(selected)

        try:
            self.service.confirm_order(order, match_index)
            self._set_status(f"Orden #{order.order_id} confirmada.", "success")
            next_id = self._find_next_pending_order_id()
            self._refresh_orders_list(select_order_id=next_id)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo confirmar la orden:\n{e}")

    def _discard_order(self):
        if not self.current_order or not self.service:
            return

        try:
            self.service.discard_order(self.current_order)
            self._set_status(f"Orden #{self.current_order.order_id} descartada.", "warning")
            next_id = self._find_next_pending_order_id()
            self._refresh_orders_list(select_order_id=next_id)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo descartar la orden:\n{e}")

    def _skip_order(self):
        next_id = self._find_next_pending_order_id(skip_current=True)
        if next_id is not None:
            self._select_order(next_id)

    def _find_next_pending_order_id(self, skip_current=False):
        if not self.service or not self.service.orders:
            return None

        orders = self.service.orders

        if self.current_order_id is None:
            pending = next((o for o in orders if o.is_pending), None)
            return pending.order_id if pending else None

        current_index = next((i for i, o in enumerate(orders) if o.order_id == self.current_order_id), -1)
        if current_index == -1:
            pending = next((o for o in orders if o.is_pending), None)
            return pending.order_id if pending else None

        start_offset = 1 if skip_current else 0
        for step in range(start_offset, len(orders) + 1):
            candidate = orders[(current_index + step) % len(orders)]
            if candidate.is_pending:
                return candidate.order_id

        return self.current_order_id

    # -----------------------------
    # Exportar
    # -----------------------------
    def _export_orders(self):
        if not self.service:
            return

        confirmed = self.service.get_confirmed_orders()
        if not confirmed:
            messagebox.showinfo("Info", "No hay órdenes confirmadas para exportar.")
            return

        default_path = os.path.join(os.path.expanduser("~"), "Downloads", "returns_confirmed.xlsx")
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="returns_confirmed.xlsx",
            initialdir=os.path.dirname(default_path)
        )

        if not filepath:
            return

        try:
            export_orders_to_excel(confirmed, filepath)
            self._set_status("Archivo exportado correctamente.", "success")
            messagebox.showinfo("Éxito", f"Exportado: {filepath}")
        except Exception as e:
            self._set_status("Error al exportar.", "danger")
            messagebox.showerror("Error", f"Error al exportar:\n{e}")


# -----------------------------
# App principal
# -----------------------------
class Application(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Returns · TANET")
        self.geometry("1500x900")
        self.minsize(1250, 780)
        self.configure(fg_color=APP_BG)

        self._center_window()

        self.container = ctk.CTkFrame(self, fg_color="transparent")
        self.container.pack(fill="both", expand=True)

        self._show_login()

    def _center_window(self):
        self.update_idletasks()
        width = 1500
        height = 900
        x = (self.winfo_screenwidth() - width) // 2
        y = (self.winfo_screenheight() - height) // 2
        self.geometry(f"{width}x{height}+{x}+{y}")

    def _clear_container(self):
        for widget in self.container.winfo_children():
            widget.destroy()

    def _show_login(self):
        self._clear_container()
        login = LoginFrame(self.container, self._on_login_success)
        login.pack(fill="both", expand=True, padx=20, pady=20)

    def _on_login_success(self, sites):
        self._clear_container()
        orders_frame = OrdersFrame(self.container, sites)
        orders_frame.pack(fill="both", expand=True, padx=6, pady=6)


def run():
    app = Application()
    app.mainloop()


if __name__ == "__main__":
    run()