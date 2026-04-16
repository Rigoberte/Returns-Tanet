"""
Exportador de órdenes a Excel.
"""
import pandas as pd
from typing import List, Dict, Any

from src.models import Order


def export_orders_to_excel(orders: List[Order], output_path: str) -> str:
    """
    Exporta las órdenes confirmadas a un archivo Excel.
    
    Args:
        orders: Lista de órdenes a exportar.
        output_path: Ruta del archivo de salida.
        
    Returns:
        Ruta del archivo creado, o vacío si no hay órdenes.
    """
    if not orders:
        return ""
    
    data = [_order_to_dict(order) for order in orders]
    df = pd.DataFrame(data)
    df.to_excel(output_path, index=False)
    
    return output_path


def export_processed_rows_to_excel(rows: List[Dict[str, Any]], output_path: str) -> str:
    """Exporta filas procesadas (incluyendo Tracking Number) a un archivo Excel."""
    if not rows:
        return ""

    df = pd.DataFrame(rows)
    df.to_excel(output_path, index=False)

    return output_path


def _order_to_dict(order: Order) -> Dict[str, Any]:
    """Convierte una orden a diccionario para exportación."""
    row = {
        'ORDER_ID': order.order_id,
        'ORIGINAL_PROTOCOL': order.protocol,
        'ORIGINAL_SITE': order.site_number,
    }
    
    # Agregar datos extra de la orden
    for key, value in order.extra_data.items():
        row[f'INPUT_{key.upper()}'] = value
    
    # Agregar datos del match seleccionado
    if order.selected_match:
        match = order.selected_match
        row['TANET_PROTOCOL'] = match.site_info.protocol
        row['TANET_SITE'] = match.site_info.site_number
        row['TANET_ID_UBICACION'] = match.site_info.id_ubicacion
        row['SIMILARITY'] = match.similarity
        
        # Agregar datos raw del sitio
        for key, value in match.site_info.raw_data.items():
            if key not in ['nomlinea', 'site', 'idubicacion']:
                row[f'TANET_{key.upper()}'] = value
    
    return row
