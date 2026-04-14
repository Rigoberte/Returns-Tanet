"""
Modelos de dominio puros para el procesamiento de órdenes.
Estos modelos no tienen dependencias de infraestructura (pandas, I/O, etc.).
"""
from dataclasses import dataclass, field
from typing import List, Optional, Dict, Any
from enum import Enum


class MatchStatus(Enum):
    """Estado de la coincidencia encontrada."""
    NO_MATCH = "no_match"
    SINGLE_MATCH = "single_match"
    MULTIPLE_MATCHES = "multiple"


class OrderStatus(Enum):
    """Estado de una orden en el flujo de confirmación."""
    PENDING = "pending"
    CONFIRMED = "confirmed"
    DISCARDED = "discarded"


@dataclass(frozen=True)
class SiteInfo:
    """Información de un sitio de TANET (inmutable)."""
    id_ubicacion: str
    protocol: str
    site_number: str
    nomdomicilio: str = ""
    calle: str = ""
    localidad: str = ""
    nomprovincia: str = ""
    nompais: str = ""
    raw_data: Dict[str, Any] = field(default_factory=dict)
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'SiteInfo':
        """Crea un SiteInfo desde un diccionario."""
        return cls(
            id_ubicacion=str(data.get('idubicacion', '')),
            protocol=str(data.get('nomlinea', '')),
            site_number=str(data.get('site', '')),
            nomdomicilio=str(data.get('nomdomicilio', '')),
            calle=str(data.get('calle', '')),
            localidad=str(data.get('localidad', '')),
            nomprovincia=str(data.get('nomprovincia', '')),
            nompais=str(data.get('nompais', '')),
            raw_data=dict(data)
        )


@dataclass(frozen=True)
class MatchResult:
    """Resultado de una coincidencia encontrada (inmutable)."""
    site_info: SiteInfo
    similarity: float
    match_index: int


@dataclass
class Order:
    """Representa una orden del usuario."""
    order_id: int
    protocol: str
    site_number: str
    extra_data: Dict[str, Any] = field(default_factory=dict)
    
    # Estado mutable gestionado por el servicio
    _status: OrderStatus = field(default=OrderStatus.PENDING, repr=False)
    _matches: List[MatchResult] = field(default_factory=list, repr=False)
    _selected_match: Optional[MatchResult] = field(default=None, repr=False)
    
    @property
    def status(self) -> OrderStatus:
        return self._status
    
    @property
    def matches(self) -> List[MatchResult]:
        return list(self._matches)  # Retorna copia
    
    @property
    def selected_match(self) -> Optional[MatchResult]:
        return self._selected_match
    
    @property
    def match_status(self) -> MatchStatus:
        """Determina el estado de coincidencias."""
        if not self._matches:
            return MatchStatus.NO_MATCH
        elif len(self._matches) == 1:
            return MatchStatus.SINGLE_MATCH
        else:
            return MatchStatus.MULTIPLE_MATCHES
    
    @property
    def match_count(self) -> int:
        return len(self._matches)
    
    @property
    def is_confirmed(self) -> bool:
        return self._status == OrderStatus.CONFIRMED
    
    @property
    def is_discarded(self) -> bool:
        return self._status == OrderStatus.DISCARDED
    
    @property
    def is_pending(self) -> bool:
        return self._status == OrderStatus.PENDING


@dataclass
class ProcessingSummary:
    """Resumen del procesamiento de órdenes."""
    total_orders: int
    no_matches: int
    single_matches: int
    multiple_matches: int
    confirmed: int
    discarded: int
    
    @property
    def pending(self) -> int:
        return self.total_orders - self.confirmed - self.discarded
