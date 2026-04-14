"""
Módulo para gestionar el Excel temporal donde el usuario ingresa sus órdenes.
"""
import os
import time
import pandas as pd
from pathlib import Path
import tempfile
from datetime import datetime, timedelta, time as dt_time
import re


class ExcelManager:
    """Gestiona la creación y lectura del Excel temporal de órdenes."""
    
    TEMPLATE_COLUMNS = [
        'protocol', 'site_number', 'email', 'cantidad_cajas',
        'fecha_retiro', 'hora_retiro_desde', 'hora_retiro_hasta',
        'referencia', 'sector', 'contactos',
        'tipo_material', 'comentarios'
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
        
        Valida y procesa:
        - fecha_retiro: debe ser una fecha válida
        - hora_retiro_desde y hora_retiro_hasta: deben ser horarios válidos (formato hh:mm)
        
        Crea columnas:
        - retiro_desde: fecha_retiro + hora_retiro_desde (formato dd/mm/yyyy hh:mm)
        - retiro_hasta: fecha_retiro + hora_retiro_hasta (formato dd/mm/yyyy hh:mm)
        - entrega_desde: siguiente día hábil a las 09:00 (formato dd/mm/yyyy hh:mm)
        - entrega_hasta: siguiente día hábil a las 16:00 (formato dd/mm/yyyy hh:mm)
        
        Returns:
            DataFrame con las órdenes del usuario.
            
        Raises:
            ValueError: Si hay errores de validación en los datos.
        """
        if not self.excel_path.exists():
            raise FileNotFoundError(f"No se encontró el archivo: {self.excel_path}")
        
        df = pd.read_excel(self.excel_path)
        
        # Filtrar filas de ejemplo o vacías
        df = df[~df['protocol'].astype(str).str.startswith('EJEMPLO:')]
        df = df.dropna(subset=['protocol', 'site_number'])
        
        # Limpiar datos principales
        df['protocol'] = df['protocol'].astype(str).str.strip()
        df['site_number'] = df['site_number'].astype(str).str.strip()
        
        # Procesar fechas y horarios
        errors = []
        
        # Crear listas para las nuevas columnas
        retiro_desde_list = []
        retiro_hasta_list = []
        entrega_desde_list = []
        entrega_hasta_list = []
        
        for idx, row in df.iterrows():
            try:
                # Validar y procesar fecha_retiro
                if pd.isna(row.get('fecha_retiro')):
                    raise ValueError("fecha_retiro es requerida")
                
                fecha_retiro = self.parse_date(row['fecha_retiro'])
                
                # Validar y procesar horarios
                if pd.isna(row.get('hora_retiro_desde')):
                    raise ValueError("hora_retiro_desde es requerida")
                if pd.isna(row.get('hora_retiro_hasta')):
                    raise ValueError("hora_retiro_hasta es requerida")
                
                hora_desde = self.parse_time(row['hora_retiro_desde'])
                hora_hasta = self.parse_time(row['hora_retiro_hasta'])
                
                # Crear columnas de retiro
                retiro_desde = self.format_datetime(fecha_retiro, hora_desde)
                retiro_hasta = self.format_datetime(fecha_retiro, hora_hasta)
                
                # Calcular siguiente día hábil para entrega
                fecha_entrega = self.get_next_business_day(fecha_retiro)
                entrega_desde = self.format_datetime(fecha_entrega, "09:00")
                entrega_hasta = self.format_datetime(fecha_entrega, "16:00")
                
                # Agregar a las listas
                retiro_desde_list.append(retiro_desde)
                retiro_hasta_list.append(retiro_hasta)
                entrega_desde_list.append(entrega_desde)
                entrega_hasta_list.append(entrega_hasta)
                
            except Exception as e:
                protocol = row.get('protocol', 'desconocido')
                error_msg = f"Fila {idx + 2} (protocol={protocol}): {str(e)}"
                errors.append(error_msg)
                # Agregar valores None para mantener la alineación
                retiro_desde_list.append(None)
                retiro_hasta_list.append(None)
                entrega_desde_list.append(None)
                entrega_hasta_list.append(None)
        
        # Si hay errores, lanzar excepción con todos los detalles
        if errors:
            error_details = "\n".join(errors)
            raise ValueError(f"Errores de validación en los datos:\n{error_details}")
        
        # Agregar las nuevas columnas al DataFrame
        df['retiro_desde'] = retiro_desde_list
        df['retiro_hasta'] = retiro_hasta_list
        df['entrega_desde'] = entrega_desde_list
        df['entrega_hasta'] = entrega_hasta_list
        
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


    def parse_date(self, date_value) -> datetime:
        """
        Convierte diversos formatos de fecha a un objeto datetime.
        
        Args:
            date_value: Fecha en varios formatos (str, datetime, etc.)
            
        Returns:
            Objeto datetime
            
        Raises:
            ValueError: Si la fecha no es válida
        """
        if isinstance(date_value, datetime):
            return date_value
        
        if isinstance(date_value, str):
            date_value = date_value.strip()
            # Probar varios formatos comunes
            for fmt in ['%d/%m/%Y', '%d-%m-%Y', '%Y-%m-%d', '%d/%m/%y']:
                try:
                    return datetime.strptime(date_value, fmt)
                except ValueError:
                    continue
            raise ValueError(f"Formato de fecha no válido: {date_value}")
        
        raise ValueError(f"Tipo de fecha no válido: {type(date_value)}")

    def parse_time(self, time_value) -> str:
        """
        Convierte diversos formatos de hora a un string en formato hh:mm.
        Maneja horarios con o sin segundos.
    
        Args:
            time_value: Hora en varios formatos (str, etc.)
            Soporta: hh:mm, h:mm, hh:mm:ss, h:mm:ss, hhmm, hhmmss
        
        Returns:
            String en formato hh:mm (sin segundos)
        
        Raises:
            ValueError: Si la hora no es válida
        """
        if isinstance(time_value, dt_time) or isinstance(time_value, datetime):
            return f"{time_value.hour:02d}:{time_value.minute:02d}"

        if isinstance(time_value, str):
            time_value = time_value.strip()
        
            # Probar formato hh:mm:ss (con segundos)
            if re.match(r'^\d{1,2}:\d{2}:\d{2}$', time_value):
                parts = time_value.split(':')
                hour = int(parts[0])
                minute = int(parts[1])
                second = int(parts[2])
                if 0 <= hour < 24 and 0 <= minute < 60 and 0 <= second < 60:
                    return f"{hour:02d}:{minute:02d}"
        
            # Probar formato hh:mm (sin segundos)
            if re.match(r'^\d{1,2}:\d{2}$', time_value):
                parts = time_value.split(':')
                hour = int(parts[0])
                minute = int(parts[1])
                if 0 <= hour < 24 and 0 <= minute < 60:
                    return f"{hour:02d}:{minute:02d}"
        
            # Probar formato hhmmss (sin separadores, con segundos)
            if re.match(r'^\d{6}$', time_value):
                hour = int(time_value[:2])
                minute = int(time_value[2:4])
                second = int(time_value[4:6])
                if 0 <= hour < 24 and 0 <= minute < 60 and 0 <= second < 60:
                    return f"{hour:02d}:{minute:02d}"
        
            # Probar formato hhmm (sin separadores, sin segundos)
            if re.match(r'^\d{4}$', time_value):
                hour = int(time_value[:2])
                minute = int(time_value[2:])
                if 0 <= hour < 24 and 0 <= minute < 60:
                    return f"{hour:02d}:{minute:02d}"
        
            raise ValueError(f"Formato de hora no válido: {time_value}")
    
        raise ValueError(f"Tipo de hora no válido: {type(time_value)}")
    
    def get_next_business_day(self, date: datetime) -> datetime:
        """
        Calcula el siguiente día hábil (lunes a viernes).
        
        Args:
            date: Fecha base
            
        Returns:
            Fecha del siguiente día hábil
        """
        next_day = date + timedelta(days=1)
        # 5 = sábado, 6 = domingo
        while next_day.weekday() >= 5:
            next_day += timedelta(days=1)
        return next_day

    def format_datetime(self, date: datetime, time_str: str) -> str:
        """
        Combina una fecha y una hora en formato dd/mm/yyyy hh:mm.
        
        Args:
            date: Objeto datetime
            time_str: String en formato hh:mm
            
        Returns:
            String en formato dd/mm/yyyy hh:mm
        """
        return date.strftime('%d/%m/%Y') + f" {time_str}"
