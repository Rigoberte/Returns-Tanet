"""
Servicios de aplicación y de dominio para procesamiento de órdenes.
"""
from typing import Any, Dict, List, Tuple
from difflib import SequenceMatcher

from src.models import (
    Order, SiteInfo, MatchResult, MatchStatus, 
    OrderStatus, ProcessingSummary
)
from src.tanet import Tanet
from src.excel_manager import ExcelManager


class WorkflowService:
    """Orquesta infraestructura (TANET + Excel) fuera de la GUI."""

    def __init__(self) -> None:
        self._tanet = Tanet()
        self._excel_manager = ExcelManager()

    def login_and_load_sites(self, username: str, password: str) -> List[SiteInfo]:
        self._tanet.login(username, password)
        df = self._tanet.load_site_data()
        return [SiteInfo.from_dict(row.to_dict()) for _, row in df.iterrows()]

    def create_and_open_excel_template(self) -> None:
        self._excel_manager.create_template()
        self._excel_manager.open_excel()

    def wait_for_excel_close(self) -> bool:
        return self._excel_manager.wait_for_excel_close()

    def read_orders_from_excel(self) -> List[Order]:
        return self._excel_manager.read_orders_as_models()

    def process_confirmed_orders(self, order_service: "OrderService") -> Tuple[List[Dict[str, Any]], List[str]]:
        rows_to_export: List[Dict[str, Any]] = []
        errors: List[str] = []
        browser_started = False

        try:
            self._tanet.build_driver_and_login()
            browser_started = True

            for order in order_service.get_confirmed_orders():
                try:
                    id_ubicacion = order.confirmed_site_id
                    if not id_ubicacion:
                        raise ValueError("id_ubicacion vacío en la ubicación confirmada")

                    extra = order.extra_data or {}
                    response = self._tanet.create_return(
                        id_ubicacion=id_ubicacion,
                        referencia=OrderService._safe_text(extra.get("referencia"), ""),
                        retiradde=OrderService._safe_text(extra.get("retiro_desde"), ""),
                        retirahta=OrderService._safe_text(extra.get("retiro_hasta"), ""),
                        entregadde=OrderService._safe_text(extra.get("entrega_desde"), ""),
                        entregahta=OrderService._safe_text(extra.get("entrega_hasta"), ""),
                        obsOper=OrderService._safe_text(extra.get("comentarios"), ""),
                        tipomaterial=OrderService._safe_int(extra.get("tipo_material"), default=1),
                        cajas=OrderService._safe_int(extra.get("cantidad_cajas"), default=1),
                    )

                    tracking_number = OrderService._extract_tracking_number(response)
                    if not tracking_number:
                        raise ValueError("respuesta sin Tracking Number (strJob)")

                    order._tracking_number = tracking_number
                    self._tanet.print_label_document(tracking_number)
                    order._status = OrderStatus.PROCESSED

                    row: Dict[str, Any] = {
                        "ORDER_ID": order.order_id,
                        "protocol": order.protocol,
                        "site_number": order.site_number,
                        **extra,
                        "id_ubicacion": id_ubicacion,
                        "Tracking Number": tracking_number,
                    }
                    rows_to_export.append(row)
                except Exception as e:
                    errors.append(f"Orden #{order.order_id}: {e}")
        except Exception as e:
            errors.append(f"Error de navegador TANET: {e}")
        finally:
            if browser_started:
                try:
                    self._tanet.close_browser()
                except Exception as e:
                    errors.append(f"No se pudo cerrar navegador TANET: {e}")

        return rows_to_export, errors


class OrderService:
    """Servicio principal para gestionar órdenes y buscar coincidencias."""
    
    def __init__(self, sites: List[SiteInfo]) -> None:
        self._sites: List[SiteInfo] = sites
        self._orders: List[Order] = []
    
    def add_orders(self, orders: List[Order]) -> None:
        """Agrega órdenes al servicio."""
        self._orders.extend(orders)
    
    def clear_orders(self) -> None:
        """Limpia todas las órdenes."""
        self._orders.clear()
    
    @property
    def orders(self) -> List[Order]:
        return list(self._orders)
    
    def process_all_orders(self) -> None:
        """Procesa todas las órdenes buscando coincidencias."""
        for order in self._orders:
            matches = self._find_matches(order)
            order._matches = matches
    
    def _find_matches(self, order: Order) -> List[MatchResult]:
        """Busca coincidencias para una orden."""
        matches = []
        protocol = order.protocol.upper().strip()
        site_number = order.site_number.upper().strip()
        
        similarity_threshold = 0.8 # Umbral de similitud para considerar una coincidencia

        for site in self._sites:
            site_protocol = site.protocol.upper().strip()
            site_site_number = site.site_number.upper().strip()

            similarity = SequenceMatcher(None, protocol, site_protocol).ratio()
            
            if similarity >= similarity_threshold and site_number in site_site_number:
                matches.append(MatchResult(
                    site_info=site,
                    similarity=round(similarity * 100, 2),
                    match_index=len(matches) + 1
                ))
        
        return matches
    
    def confirm_order(self, order: Order, match_index: int) -> bool:
        """Confirma una orden con una coincidencia específica."""
        if match_index < 1 or match_index > len(order._matches):
            return False
        
        order._selected_match = order._matches[match_index - 1]
        order._status = OrderStatus.CONFIRMED
        return True
    
    def discard_order(self, order: Order) -> bool:
        """Descarta una orden."""
        order._status = OrderStatus.DISCARDED
        return True
    
    def get_summary(self) -> ProcessingSummary:
        """Obtiene un resumen del estado actual."""
        total = len(self._orders)
        no_match = sum(1 for o in self._orders if o.match_status == MatchStatus.NO_MATCH)
        single = sum(1 for o in self._orders if o.match_status == MatchStatus.SINGLE_MATCH)
        multiple = sum(1 for o in self._orders if o.match_status == MatchStatus.MULTIPLE_MATCHES)
        confirmed = sum(1 for o in self._orders if o.is_confirmed)
        processed = sum(1 for o in self._orders if o.is_processed)
        discarded = sum(1 for o in self._orders if o.is_discarded)
        
        return ProcessingSummary(
            total_orders=total,
            no_matches=no_match,
            single_matches=single,
            multiple_matches=multiple,
            confirmed=confirmed,
            processed=processed,
            discarded=discarded
        )
    
    def get_confirmed_orders(self) -> List[Order]:
        """Obtiene las órdenes confirmadas."""
        return [o for o in self._orders if o.is_confirmed]

    @staticmethod
    def _extract_tracking_number(response: Dict[str, Any]) -> str:
        if not isinstance(response, dict):
            return ""

        data_node = response.get("data")
        if isinstance(data_node, dict):
            tracking = data_node.get("strJob")
            if tracking:
                return str(tracking)

            inner_data = data_node.get("data")
            if isinstance(inner_data, dict):
                tracking = inner_data.get("strJob")
                if tracking:
                    return str(tracking)

        tracking = response.get("strJob")
        return str(tracking) if tracking else ""

    @staticmethod
    def _safe_text(value: Any, fallback: str = "-") -> str:
        if value is None:
            return fallback
        text = str(value).strip()
        return text if text else fallback

    @staticmethod
    def _safe_int(value: Any, default: int = 1) -> int:
        if value is None:
            return default

        text = str(value).strip()
        if not text:
            return default

        try:
            number = float(text.replace(",", "."))
            return int(number)
        except Exception:
            return default
