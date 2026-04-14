"""
Aplicación principal para procesar órdenes de Returns usando datos de TANET.

Punto de entrada que lanza la interfaz gráfica.
"""
from src.gui import run

from src.tanet import Tanet

def main():
    """Punto de entrada de la aplicación."""
    run()

    """ta = Tanet()
    ta.login("fcstest", "Fcs1234")

    print("Cargando datos de sitios desde TANET...")
    site_data = ta.load_site_data()
    print(site_data.head())

    print("Creando orden de devolución en TANET...")
    order_data = ta.create_return(
        id_ubicacion="540",
        referencia="REF004 ñáéíóú",
        retiradde="12/12/2026 16:00",
        retirahta="12/12/2026 16:30",
        entregadde="13/12/2026 20:00",
        entregahta="13/12/2026 22:00",
        obsOper="Observación de prueba \n Otra línea de observación \n ñáéíóú",
        tipomaterial=9,
        cajas=1
    )

    print(order_data)"""

if __name__ == "__main__":
    main()