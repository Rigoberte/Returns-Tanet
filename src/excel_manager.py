"""
Módulo para gestionar el Excel temporal donde el usuario ingresa sus órdenes.
"""
import os
import time
import pandas as pd
from pathlib import Path
import tempfile


class ExcelManager:
    """Gestiona la creación y lectura del Excel temporal de órdenes."""
    
    TEMPLATE_COLUMNS = [
        'protocol', 'site_number', 'referencia',
        'retira_desde', 'retira_hasta',
        'entrega_desde', 'entrega_hasta',
        'sector', 'tipo_material', 'temperatura',
        'autorizado', 'telefono', 'cantidad_cajas'
    ]
    
    def __init__(self, excel_path: str = None):
        """
        Inicializa el gestor de Excel.
        
        Args:
            excel_path: Ruta del archivo Excel. Si no se proporciona, se crea uno temporal.
        """
        if excel_path:
            self.excel_path = Path(excel_path)
        else:
            temp_dir = tempfile.gettempdir()
            self.excel_path = Path(temp_dir) / "return_orders_temp.xlsx"
    
    def create_template(self) -> str:
        """
        Crea una plantilla Excel vacía con las columnas necesarias.
        
        Returns:
            Ruta del archivo creado.
        """
        template_df = pd.DataFrame(columns=self.TEMPLATE_COLUMNS)
        
        example_row = pd.DataFrame([{
            'protocol': 'EJEMPLO: MK-6482-011',
            'site_number': '800'
        }])
        template_df = pd.concat([template_df, example_row], ignore_index=True)
        
        template_df.to_excel(self.excel_path, index=False)
        print(f"Plantilla creada en: {self.excel_path}")
        
        return str(self.excel_path)
    
    def open_excel(self) -> None:
        """Abre el archivo Excel con la aplicación predeterminada."""
        if not self.excel_path.exists():
            self.create_template()
        
        # Abrir con la aplicación predeterminada en Windows
        os.startfile(str(self.excel_path))
        print("Excel abierto. Complete los datos y guarde el archivo.")
    
    def wait_for_excel_close(self, check_interval: float = 2.0) -> bool:
        """
        Espera hasta que el usuario cierre el archivo Excel.
        
        Args:
            check_interval: Intervalo en segundos entre cada verificación.
            
        Returns:
            True si el archivo fue cerrado correctamente.
        """
        print("\nEsperando a que cierre el archivo Excel...")
        print("(Presione Ctrl+C para cancelar)")
        
        # Esperar un momento para que Excel abra
        time.sleep(2)
        
        try:
            while self._is_file_locked():
                time.sleep(check_interval)
        except KeyboardInterrupt:
            print("\nOperación cancelada por el usuario.")
            return False
        
        print("Archivo Excel cerrado. Procesando datos...")
        return True
    
    def _is_file_locked(self) -> bool:
        """
        Verifica si el archivo está siendo usado por otro proceso.
        
        Returns:
            True si el archivo está bloqueado.
        """
        if not self.excel_path.exists():
            return False
            
        try:
            # Intentar abrir el archivo en modo exclusivo
            with open(self.excel_path, 'r+b'):
                return False
        except (IOError, PermissionError):
            return True
    
    def read_orders(self) -> pd.DataFrame:
        """
        Lee las órdenes del archivo Excel.
        
        Returns:
            DataFrame con las órdenes del usuario.
        """
        if not self.excel_path.exists():
            raise FileNotFoundError(f"No se encontró el archivo: {self.excel_path}")
        
        df = pd.read_excel(self.excel_path)
        
        # Filtrar filas de ejemplo o vacías
        df = df[~df['protocol'].astype(str).str.startswith('EJEMPLO:')]
        df = df.dropna(subset=['protocol', 'site_number'])
        
        # Limpiar datos
        df['protocol'] = df['protocol'].astype(str).str.strip()
        df['site_number'] = df['site_number'].astype(str).str.strip()
        
        # Agregar índice de orden
        df['ORDER_ID'] = range(1, len(df) + 1)
        
        return df
    
    def read_orders_as_models(self) -> list:
        """
        Lee las órdenes del archivo Excel y las retorna como modelos Order.
        
        Returns:
            Lista de objetos Order.
        """
        from src.models import Order
        
        df = self.read_orders()
        orders = []
        
        for _, row in df.iterrows():
            # Extraer datos extra (todas las columnas excepto las principales)
            extra_data = {}
            for col in df.columns:
                if col not in ['protocol', 'site_number', 'ORDER_ID']:
                    value = row[col]
                    if pd.notna(value):
                        extra_data[col] = value
            
            order = Order(
                order_id=int(row['ORDER_ID']),
                protocol=row['protocol'],
                site_number=row['site_number'],
                extra_data=extra_data
            )
            orders.append(order)
        
        return orders
    
    def cleanup(self) -> None:
        """Elimina el archivo temporal si existe."""
        if self.excel_path.exists():
            try:
                os.remove(self.excel_path)
                print(f"Archivo temporal eliminado: {self.excel_path}")
            except Exception as e:
                print(f"No se pudo eliminar el archivo temporal: {e}")
