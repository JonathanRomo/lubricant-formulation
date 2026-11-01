from pathlib import Path
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

from reconstruir_base_datos import EVENTO_CATALOGO, notificar_catalogo, suscribir_catalogo
from interfaz_seleccion_especificaciones import Catalogo, Especificacion, SelectorEspecificaciones


class RaizEventos:
    def __init__(self):
        self.enlaces = {}
        self.pendientes = []
        self.contador = 0

    def bind(self, evento, accion, add):
        assert evento == EVENTO_CATALOGO and add == "+"
        self.contador += 1
        enlace = str(self.contador)
        self.enlaces[enlace] = accion
        return enlace

    def unbind(self, evento, enlace):
        assert evento == EVENTO_CATALOGO
        self.enlaces.pop(enlace)

    def event_generate(self, evento, when):
        assert evento == EVENTO_CATALOGO and when == "tail"
        self.pendientes.append(evento)

    def recibir(self):
        while self.pendientes:
            self.pendientes.pop(0)
            for accion in list(self.enlaces.values()):
                accion(SimpleNamespace())


class VentanaEventos:
    def __init__(self, raiz):
        self.raiz = raiz
        self.existe = True
        self.al_destruir = None

    def _root(self):
        return self.raiz

    def winfo_exists(self):
        return self.existe

    def bind(self, evento, accion, add):
        assert evento == "<Destroy>" and add == "+"
        self.al_destruir = accion

    def cerrar(self):
        self.existe = False
        self.al_destruir(SimpleNamespace(widget=self))


class EventosCatalogoTests(unittest.TestCase):
    def test_recarga_todas_las_listas_cada_vez_que_termina_el_proceso(self):
        raiz = RaizEventos()
        recargas = [Mock(), Mock(), Mock()]
        ventanas = [VentanaEventos(raiz) for _ in recargas]
        for ventana, recargar in zip(ventanas, recargas):
            suscribir_catalogo(ventana, recargar)
        for pasada in (1, 2):
            notificar_catalogo(raiz)
            self.assertEqual(recargas[0].call_count, pasada - 1)
            raiz.recibir()
            self.assertEqual([r.call_count for r in recargas], [pasada] * 3)

    def test_cerrar_ventana_libera_solo_su_suscripcion(self):
        raiz = RaizEventos()
        cerrada, abierta = VentanaEventos(raiz), VentanaEventos(raiz)
        recargar_cerrada, recargar_abierta = Mock(), Mock()
        suscribir_catalogo(cerrada, recargar_cerrada)
        suscribir_catalogo(abierta, recargar_abierta)
        notificar_catalogo(raiz)
        cerrada.cerrar()
        raiz.recibir()
        recargar_cerrada.assert_not_called()
        recargar_abierta.assert_called_once_with()
        self.assertEqual(len(raiz.enlaces), 1)

    def test_destruir_un_control_hijo_no_desuscribe_la_ventana(self):
        raiz = RaizEventos()
        ventana = VentanaEventos(raiz)
        recargar = Mock()
        suscribir_catalogo(ventana, recargar)
        ventana.al_destruir(SimpleNamespace(widget=object()))
        notificar_catalogo(raiz)
        raiz.recibir()
        recargar.assert_called_once_with()

    def test_selector_descarta_marcas_y_previa_de_archivos_movidos(self):
        conservada = Especificacion(Path("conservada.json"), "Producto", "20100001", "Cliente", "00", "PRODUCTO")
        movida = Especificacion(Path("movida.json"), "Producto", "20100001", "Cliente", "00", "PRODUCTO")
        panel = SimpleNamespace(catalogo=Catalogo((conservada, movida), ()),
                                documentos_marcados={conservada.ruta, movida.ruta},
                                especificacion_actual=movida, estado_catalogo=Mock())
        selector = SelectorEspecificaciones.__new__(SelectorEspecificaciones)
        selector.carpetas_json = {"nuevas": Path("catalogo")}
        selector.paneles = {"nuevas": panel}
        for metodo in ("_actualizar_sugerencias", "_actualizar_boton_idioma", "_actualizar_boton_contraparte",
                       "_actualizar_estado_accion", "_cerrar_vista_previa", "_mostrar_previa"):
            setattr(selector, metodo, Mock())
        with patch("interfaz_seleccion_especificaciones.cargar_catalogo", return_value=Catalogo((conservada,), ())):
            selector.recargar_catalogos()
        self.assertEqual(panel.documentos_marcados, {conservada.ruta})
        selector._cerrar_vista_previa.assert_called_once_with("nuevas")
        selector._mostrar_previa.assert_not_called()
        selector._actualizar_sugerencias.assert_called_once_with("nuevas")


if __name__ == "__main__":
    unittest.main()
