"""
Servicio de órdenes - Lógica de negocio para procesamiento de órdenes.
"""
from typing import List, Optional
from difflib import SequenceMatcher

from src.models import (
    Order, SiteInfo, MatchResult, MatchStatus, 
    OrderStatus, ProcessingSummary
)


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
        discarded = sum(1 for o in self._orders if o.is_discarded)
        
        return ProcessingSummary(
            total_orders=total,
            no_matches=no_match,
            single_matches=single,
            multiple_matches=multiple,
            confirmed=confirmed,
            discarded=discarded
        )
    
    def get_confirmed_orders(self) -> List[Order]:
        """Obtiene las órdenes confirmadas."""
        return [o for o in self._orders if o.is_confirmed]
