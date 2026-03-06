"""
Interfaz gráfica con TKinter para el procesador de órdenes Returns-TANET.
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import os

from src.tanet import Tanet
from src.excel_manager import ExcelManager
from src.services import OrderService
from src.models import SiteInfo, MatchStatus, Order
from src.exporters import export_orders_to_excel


class LoginFrame(ttk.Frame):
    """Frame de inicio de sesión."""
    
    def __init__(self, parent, on_login_success):
        super().__init__(parent, padding=20)
        self.on_login_success = on_login_success
        self._create_widgets()
    
    def _create_widgets(self):
        # Título
        ttk.Label(self, text="RETURNS - TANET", font=('Helvetica', 16, 'bold')).pack(pady=(0, 20))
        ttk.Label(self, text="Inicio de Sesión", font=('Helvetica', 12)).pack(pady=(0, 15))
        
        # Usuario
        ttk.Label(self, text="Usuario:").pack(anchor='w')
        self.username_entry = ttk.Entry(self, width=30)
        self.username_entry.pack(pady=(0, 10), fill='x')
        
        # Contraseña
        ttk.Label(self, text="Contraseña:").pack(anchor='w')
        self.password_entry = ttk.Entry(self, width=30, show="*")
        self.password_entry.pack(pady=(0, 15), fill='x')
        
        # Botón login
        self.login_btn = ttk.Button(self, text="Conectar", command=self._do_login)
        self.login_btn.pack(pady=10)
        
        # Status
        self.status_label = ttk.Label(self, text="", foreground='gray')
        self.status_label.pack(pady=10)
        
        # Bind Enter key
        self.password_entry.bind('<Return>', lambda e: self._do_login())
    
    def _do_login(self):
        username = self.username_entry.get().strip()
        password = self.password_entry.get().strip()
        
        if not username or not password:
            messagebox.showwarning("Aviso", "Ingrese usuario y contraseña")
            return
        
        self.login_btn.config(state='disabled')
        self.status_label.config(text="Conectando...", foreground='blue')
        
        # Login en thread separado
        thread = threading.Thread(target=self._login_thread, args=(username, password))
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
        self.status_label.config(text=f"✓ {len(sites)} registros cargados", foreground='green')
        self.on_login_success(sites)
    
    def _login_error(self, error):
        self.login_btn.config(state='normal')
        self.status_label.config(text=f"Error: {error}", foreground='red')
        messagebox.showerror("Error", f"No se pudo conectar:\n{error}")


class OrdersFrame(ttk.Frame):
    """Frame principal para gestión de órdenes."""
    
    def __init__(self, parent, sites):
        super().__init__(parent, padding=10)
        self.sites = sites
        self.service = None
        self.excel_manager = ExcelManager()
        self.current_order_index = 0
        self._create_widgets()
    
    def _create_widgets(self):
        # Panel superior - Acciones
        top_frame = ttk.Frame(self)
        top_frame.pack(fill='x', pady=(0, 10))
        
        self.open_excel_btn = ttk.Button(top_frame, text="📄 Abrir Excel de Órdenes", command=self._open_excel)
        self.open_excel_btn.pack(side='left', padx=5)
        
        self.status_label = ttk.Label(top_frame, text="Esperando órdenes...", foreground='gray')
        self.status_label.pack(side='left', padx=20)
        
        # Panel central - Dividido
        paned = ttk.PanedWindow(self, orient='horizontal')
        paned.pack(fill='both', expand=True, pady=10)
        
        # Lista de órdenes (izquierda)
        left_frame = ttk.LabelFrame(paned, text="Órdenes", padding=5)
        paned.add(left_frame, weight=1)
        
        self.orders_tree = ttk.Treeview(left_frame, columns=('status', 'protocol', 'site'), show='headings', height=15)
        self.orders_tree.heading('status', text='Estado')
        self.orders_tree.heading('protocol', text='Protocolo')
        self.orders_tree.heading('site', text='Sitio')
        self.orders_tree.column('status', width=80)
        self.orders_tree.column('protocol', width=150)
        self.orders_tree.column('site', width=80)
        self.orders_tree.pack(fill='both', expand=True)
        self.orders_tree.bind('<<TreeviewSelect>>', self._on_order_select)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(left_frame, orient='vertical', command=self.orders_tree.yview)
        self.orders_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side='right', fill='y')
        
        # Detalle de orden (derecha)
        right_frame = ttk.LabelFrame(paned, text="Coincidencias", padding=10)
        paned.add(right_frame, weight=2)
        
        # Info de la orden seleccionada
        self.order_info_label = ttk.Label(right_frame, text="Seleccione una orden", font=('Helvetica', 11))
        self.order_info_label.pack(anchor='w', pady=(0, 10))
        
        # Lista de coincidencias
        self.matches_tree = ttk.Treeview(right_frame, columns=('idx', 'protocol', 'site', 'similarity'), show='headings', height=8)
        self.matches_tree.heading('idx', text='#')
        self.matches_tree.heading('protocol', text='Protocolo TANET')
        self.matches_tree.heading('site', text='Sitio TANET')
        self.matches_tree.heading('similarity', text='Similitud')
        self.matches_tree.column('idx', width=30)
        self.matches_tree.column('protocol', width=180)
        self.matches_tree.column('site', width=100)
        self.matches_tree.column('similarity', width=80)
        self.matches_tree.pack(fill='both', expand=True, pady=(0, 10))
        
        # Botones de acción
        actions_frame = ttk.Frame(right_frame)
        actions_frame.pack(fill='x', pady=10)
        
        self.confirm_btn = ttk.Button(actions_frame, text="✓ Confirmar Selección", command=self._confirm_order, state='disabled')
        self.confirm_btn.pack(side='left', padx=5)
        
        self.discard_btn = ttk.Button(actions_frame, text="✗ Descartar Orden", command=self._discard_order, state='disabled')
        self.discard_btn.pack(side='left', padx=5)
        
        self.skip_btn = ttk.Button(actions_frame, text="→ Saltar", command=self._skip_order, state='disabled')
        self.skip_btn.pack(side='left', padx=5)
        
        # Panel inferior - Resumen y exportar
        bottom_frame = ttk.Frame(self)
        bottom_frame.pack(fill='x', pady=10)
        
        self.summary_label = ttk.Label(bottom_frame, text="")
        self.summary_label.pack(side='left')
        
        self.export_btn = ttk.Button(bottom_frame, text="📥 Exportar Confirmadas", command=self._export_orders, state='disabled')
        self.export_btn.pack(side='right', padx=5)
    
    def _open_excel(self):
        self.open_excel_btn.config(state='disabled')
        self.status_label.config(text="Creando plantilla Excel...", foreground='blue')
        
        # Crear y abrir Excel
        self.excel_manager.create_template()
        self.excel_manager.open_excel()
        
        self.status_label.config(text="Complete el Excel y ciérrelo para continuar...", foreground='orange')
        
        # Esperar cierre en thread
        thread = threading.Thread(target=self._wait_excel_close)
        thread.start()
    
    def _wait_excel_close(self):
        closed = self.excel_manager.wait_for_excel_close()
        self.after(0, lambda: self._on_excel_closed(closed))
    
    def _on_excel_closed(self, closed):
        if not closed:
            self.open_excel_btn.config(state='normal')
            self.status_label.config(text="Cancelado", foreground='red')
            return
        
        try:
            orders = self.excel_manager.read_orders_as_models()
        except Exception as e:
            messagebox.showerror("Error", f"Error al leer órdenes:\n{e}")
            self.open_excel_btn.config(state='normal')
            return
        
        if not orders:
            messagebox.showwarning("Aviso", "No se encontraron órdenes válidas")
            self.open_excel_btn.config(state='normal')
            return
        
        # Procesar órdenes
        self.service = OrderService(self.sites)
        self.service.add_orders(orders)
        self.service.process_all_orders()
        
        self._refresh_orders_list()
        self._update_summary()
        self.status_label.config(text=f"✓ {len(orders)} órdenes cargadas", foreground='green')
        self.export_btn.config(state='normal')
    
    def _refresh_orders_list(self):
        self.orders_tree.delete(*self.orders_tree.get_children())
        
        for order in self.service.orders:
            status = self._get_status_text(order)
            self.orders_tree.insert('', 'end', iid=order.order_id, values=(status, order.protocol, order.site_number))
        
        self._update_row_colors()
    
    def _get_status_text(self, order):
        if order.is_confirmed:
            return "✓ Confirmada"
        elif order.is_discarded:
            return "✗ Descartada"
        elif order.match_status == MatchStatus.NO_MATCH:
            return "⚠ Sin matches"
        elif order.match_status == MatchStatus.SINGLE_MATCH:
            return "● 1 match"
        else:
            return f"● {order.match_count} matches"
    
    def _update_row_colors(self):
        for order in self.service.orders:
            tags = ()
            if order.is_confirmed:
                tags = ('confirmed',)
            elif order.is_discarded:
                tags = ('discarded',)
            elif order.match_status == MatchStatus.NO_MATCH:
                tags = ('no_match',)
            self.orders_tree.item(order.order_id, tags=tags)
        
        self.orders_tree.tag_configure('confirmed', background='#d4edda')
        self.orders_tree.tag_configure('discarded', background='#f8d7da')
        self.orders_tree.tag_configure('no_match', background='#fff3cd')
    
    def _on_order_select(self, event):
        selection = self.orders_tree.selection()
        if not selection:
            return
        
        order_id = int(selection[0])
        order = next((o for o in self.service.orders if o.order_id == order_id), None)
        
        if not order:
            return
        
        self._show_order_details(order)
    
    def _show_order_details(self, order):
        self.current_order = order
        
        # Info de la orden
        self.order_info_label.config(text=f"Orden #{order.order_id}: {order.protocol} / Sitio: {order.site_number}")
        
        # Limpiar y llenar matches
        self.matches_tree.delete(*self.matches_tree.get_children())
        
        for match in order.matches:
            self.matches_tree.insert('', 'end', iid=match.match_index, values=(
                match.match_index,
                match.site_info.protocol,
                match.site_info.site_number,
                f"{match.similarity}%"
            ))
        
        # Habilitar/deshabilitar botones según estado
        if order.is_confirmed or order.is_discarded:
            self.confirm_btn.config(state='disabled')
            self.discard_btn.config(state='disabled')
            self.skip_btn.config(state='disabled')
        else:
            self.discard_btn.config(state='normal')
            self.skip_btn.config(state='normal')
            self.confirm_btn.config(state='normal' if order.matches else 'disabled')
    
    def _confirm_order(self):
        if not hasattr(self, 'current_order'):
            return
        
        order = self.current_order
        
        # Obtener match seleccionado
        selection = self.matches_tree.selection()
        if not selection:
            if order.match_count == 1:
                match_index = 1
            else:
                messagebox.showwarning("Aviso", "Seleccione una coincidencia")
                return
        else:
            match_index = int(selection[0])
        
        self.service.confirm_order(order, match_index)
        self._refresh_orders_list()
        self._update_summary()
        self._show_order_details(order)
        self._select_next_pending()
    
    def _discard_order(self):
        if not hasattr(self, 'current_order'):
            return
        
        self.service.discard_order(self.current_order)
        self._refresh_orders_list()
        self._update_summary()
        self._show_order_details(self.current_order)
        self._select_next_pending()
    
    def _skip_order(self):
        self._select_next_pending()
    
    def _select_next_pending(self):
        for order in self.service.orders:
            if order.is_pending:
                self.orders_tree.selection_set(order.order_id)
                self.orders_tree.see(order.order_id)
                self._show_order_details(order)
                return
    
    def _update_summary(self):
        summary = self.service.get_summary()
        text = f"Total: {summary.total_orders} | Confirmadas: {summary.confirmed} | Descartadas: {summary.discarded} | Pendientes: {summary.pending}"
        self.summary_label.config(text=text)
    
    def _export_orders(self):
        confirmed = self.service.get_confirmed_orders()
        
        if not confirmed:
            messagebox.showinfo("Info", "No hay órdenes confirmadas para exportar")
            return
        
        # Diálogo para guardar
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
            messagebox.showinfo("Éxito", f"Exportado: {filepath}")
        except Exception as e:
            messagebox.showerror("Error", f"Error al exportar:\n{e}")


class Application(tk.Tk):
    """Ventana principal de la aplicación."""
    
    def __init__(self):
        super().__init__()
        
        self.title("Returns - TANET")
        self.geometry("900x600")
        self.minsize(800, 500)
        
        # Centrar ventana
        self.update_idletasks()
        x = (self.winfo_screenwidth() - 900) // 2
        y = (self.winfo_screenheight() - 600) // 2
        self.geometry(f"+{x}+{y}")
        
        # Estilo
        style = ttk.Style()
        style.theme_use('clam')
        
        # Frame contenedor
        self.container = ttk.Frame(self)
        self.container.pack(fill='both', expand=True)
        
        # Mostrar login
        self._show_login()
    
    def _show_login(self):
        for widget in self.container.winfo_children():
            widget.destroy()
        
        login_frame = LoginFrame(self.container, self._on_login_success)
        login_frame.pack(expand=True)
    
    def _on_login_success(self, sites):
        for widget in self.container.winfo_children():
            widget.destroy()
        
        orders_frame = OrdersFrame(self.container, sites)
        orders_frame.pack(fill='both', expand=True)


def run():
    """Punto de entrada de la GUI."""
    app = Application()
    app.mainloop()


if __name__ == "__main__":
    run()
