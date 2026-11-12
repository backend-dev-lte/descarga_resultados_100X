#!/usr/bin/env python3
"""Consulta concurrente de resultados agregados Saber 11 2026.

Usa hasta cien trabajadores con sesiones independientes, limitación global de tasa,
protección ante bloqueos y un único escritor de checkpoint. Consolida solamente
el modelo tabular necesario para Power BI en un Excel reanudable y verificable.
"""

from __future__ import annotations

import argparse
import concurrent.futures
import csv
import hashlib
import html
import io
import json
import logging
import os
import random
import re
import shutil
import sys
import threading
import time
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable
from urllib.parse import urljoin
from xml.etree import ElementTree

import requests
from bs4 import BeautifulSoup, Tag
from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


BASE_URL = "https://resultados.icfes.edu.co/resultados-saber2016-web/"
SEARCH_PATH = (
    "pages/publicacionResultados/agregados/saber11/"
    "consultaAgregadosEstablecimiento.jsf"
)
REPORT_PATH = (
    "pages/publicacionResultados/agregados/saber11/"
    "agregadosEstablecimiento.jsf"
)

AJAX_HEADERS = {
    "Faces-Request": "partial/ajax",
    "X-Requested-With": "XMLHttpRequest",
}


class IcfesError(RuntimeError):
    """Error recuperable al consultar un código."""


class BloqueoIcfesError(IcfesError):
    """Señal global para detener el lote sin rechazar colegios pendientes."""


class ErrorTransporteIcfes(IcfesError):
    """Fallo temporal que no debe convertir un colegio válido en rechazado."""


class ControlPeticiones:
    """Coordina la tasa y el cortacircuitos de todos los trabajadores."""

    def __init__(
        self,
        solicitudes_por_segundo: float,
        max_429: int = 3,
        max_fallos_transporte: int = 8,
    ) -> None:
        if solicitudes_por_segundo <= 0:
            raise IcfesError("--rps debe ser mayor que cero.")
        self.intervalo = 1.0 / solicitudes_por_segundo
        self.max_429 = max_429
        self.max_fallos_transporte = max_fallos_transporte
        self._condicion = threading.Condition()
        self._siguiente = time.monotonic()
        self._enfriamiento_hasta = 0.0
        self._incidentes_429 = 0
        self._fallos_transporte = 0
        self._exitos_desde_fallo = 0
        self._detenido = threading.Event()
        self._motivo = ""

    @property
    def detenido(self) -> bool:
        return self._detenido.is_set()

    @property
    def motivo(self) -> str:
        return self._motivo

    def detener(self, motivo: str) -> None:
        with self._condicion:
            if not self._detenido.is_set():
                self._motivo = motivo
                logging.error("Cortacircuitos activado: %s", motivo)
            self._detenido.set()
            self._condicion.notify_all()

    def antes_de_solicitar(self) -> None:
        while True:
            with self._condicion:
                if self._detenido.is_set():
                    raise BloqueoIcfesError(self._motivo or "Ejecución detenida.")
                ahora = time.monotonic()
                objetivo = max(self._siguiente, self._enfriamiento_hasta)
                if objetivo <= ahora:
                    # El pequeño jitter evita que los cincuenta trabajadores formen ráfagas.
                    self._siguiente = ahora + self.intervalo + random.uniform(0, 0.12)
                    return
                self._condicion.wait(timeout=min(objetivo - ahora, 1.0))

    def registrar_respuesta(
        self, status: int, retry_after: str | None = None
    ) -> None:
        if status == 403:
            self.detener("ICFES respondió 403; posible restricción de acceso o IP.")
            raise BloqueoIcfesError(self.motivo)
        if status == 429:
            with self._condicion:
                self._incidentes_429 += 1
                try:
                    espera = float(retry_after or "")
                except ValueError:
                    espera = min(120.0 * self._incidentes_429, 600.0)
                espera = max(espera, 60.0)
                self._enfriamiento_hasta = max(
                    self._enfriamiento_hasta, time.monotonic() + espera
                )
                logging.warning(
                    "HTTP 429: pausa global de %.0f s (%d/%d).",
                    espera,
                    self._incidentes_429,
                    self.max_429,
                )
                self._condicion.notify_all()
                if self._incidentes_429 >= self.max_429:
                    self.detener(
                        "Se alcanzó el umbral de HTTP 429. Reanude más tarde con el mismo checkpoint."
                    )
                    raise BloqueoIcfesError(self.motivo)
            return
        if status in {502, 503, 504}:
            with self._condicion:
                self._enfriamiento_hasta = max(
                    self._enfriamiento_hasta, time.monotonic() + 15.0
                )
                self._condicion.notify_all()
            return
        if 200 <= status < 400:
            with self._condicion:
                self._exitos_desde_fallo += 1
                if self._exitos_desde_fallo >= 50:
                    self._incidentes_429 = 0
                    self._fallos_transporte = 0
                    self._exitos_desde_fallo = 0

    def registrar_fallo_transporte(self, detalle: str) -> None:
        with self._condicion:
            self._fallos_transporte += 1
            self._exitos_desde_fallo = 0
            logging.warning(
                "Fallo de transporte global %d/%d: %s",
                self._fallos_transporte,
                self.max_fallos_transporte,
                detalle,
            )
            if self._fallos_transporte >= self.max_fallos_transporte:
                self.detener(
                    "Demasiados fallos HTTP/de red consecutivos; se evita rechazar colegios en masa."
                )
                raise BloqueoIcfesError(self.motivo)


@dataclass(frozen=True)
class ColegioEntrada:
    orden: int
    dane: str
    nombre_csv: str
    calendario_csv: str
    dane_original: str = ""
    nombre_original: str = ""
    calendario_original: str = ""


@dataclass
class ResultadoColegio:
    entrada: ColegioEntrada
    nombre_portal: str
    etc: str
    html_reporte: str
    fecha_consulta: str


def texto_limpio(value: object) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", html.unescape(str(value))).strip()


def normalizar_texto(value: object) -> str:
    """Normaliza Unicode, espacios y caracteres invisibles de un campo."""
    return texto_limpio(unicodedata.normalize("NFKC", str(value or "")))


def normalizar_dane(value: object) -> str:
    """Normaliza un DANE sin completar ni inventar dígitos faltantes."""
    raw = normalizar_texto(value)
    if re.fullmatch(r"\d+\.0+", raw):
        raw = raw.split(".", 1)[0]
    return re.sub(r"[\s-]+", "", raw)


def normalizar_encabezado(value: object) -> str:
    raw = unicodedata.normalize("NFKD", normalizar_texto(value)).encode(
        "ascii", "ignore"
    ).decode("ascii")
    return re.sub(r"[^a-z0-9]+", "_", raw.lower()).strip("_")


def decodificar_respuesta_icfes(data: bytes) -> str:
    """Decodifica las respuestas mixtas UTF-8/Windows-1252 del portal.

    Las plantillas JSF están en UTF-8, pero algunos valores provenientes de la
    base de datos llegan como bytes Windows-1252 aislados. ``surrogateescape``
    conserva esos bytes para convertirlos sin dañar el UTF-8 válido.
    """
    value = data.decode("utf-8", errors="surrogateescape")
    repaired: list[str] = []
    for char in value:
        code = ord(char)
        if 0xDC80 <= code <= 0xDCFF:
            raw_byte = bytes([code - 0xDC00])
            try:
                repaired.append(raw_byte.decode("cp1252"))
            except UnicodeDecodeError:
                repaired.append(raw_byte.decode("latin-1"))
        else:
            repaired.append(char)
    return "".join(repaired)


ENCABEZADOS_DANE = (
    "dane",
    "codigo",
    "codigo_dane",
    "codigo_dane_sede",
    "cod_dane",
)
ENCABEZADOS_NOMBRE = (
    "nombre",
    "institucion",
    "nombre_institucion",
    "nombre_del_establecimiento",
    "nombre_establecimiento",
    "establecimiento",
)
ENCABEZADOS_CALENDARIO = ("calendario", "calendar")


def _primer_campo(
    fields: dict[str, object], aliases: tuple[str, ...]
) -> object | None:
    return next((fields[name] for name in aliases if name in fields), None)


def _leer_entradas_csv(ruta: Path) -> list[ColegioEntrada]:
    """Lee CSV y conserva la posición original de cada registro."""
    codificaciones = ("utf-8-sig", "cp1252", "latin-1")
    ultimo_error: Exception | None = None
    for encoding in codificaciones:
        try:
            content = ruta.read_text(encoding=encoding)
            break
        except UnicodeDecodeError as exc:
            ultimo_error = exc
    else:
        raise IcfesError(f"No se pudo decodificar el CSV: {ultimo_error}")

    try:
        dialect = csv.Sniffer().sniff(content[:8192], delimiters=";,\t|")
        reader = csv.DictReader(io.StringIO(content), dialect=dialect)
    except csv.Error:
        reader = csv.DictReader(io.StringIO(content), delimiter=";")
    if not reader.fieldnames:
        raise IcfesError("El CSV no contiene encabezados.")

    fields = {normalizar_encabezado(name): name for name in reader.fieldnames if name}
    dane_field = _primer_campo(fields, ENCABEZADOS_DANE)
    name_field = _primer_campo(fields, ENCABEZADOS_NOMBRE)
    calendar_field = _primer_campo(fields, ENCABEZADOS_CALENDARIO)
    if dane_field is None:
        raise IcfesError(
            "El archivo debe contener 'dane', 'codigo_dane' o un encabezado equivalente."
        )

    entradas: list[ColegioEntrada] = []
    for orden, row in enumerate(reader, start=2):
        original_dane = normalizar_texto(row.get(dane_field))
        original_name = normalizar_texto(row.get(name_field)) if name_field else ""
        original_calendar = (
            normalizar_texto(row.get(calendar_field)) if calendar_field else ""
        )
        entradas.append(
            ColegioEntrada(
                orden=orden,
                dane=normalizar_dane(original_dane),
                nombre_csv=normalizar_texto(original_name),
                calendario_csv=normalizar_texto(original_calendar).upper(),
                dane_original=original_dane,
                nombre_original=original_name,
                calendario_original=original_calendar,
            )
        )
    return entradas


def _indices_encabezado_excel(
    values: tuple[object, ...] | list[object],
) -> tuple[int, int | None, int | None] | None:
    fields = {
        normalizar_encabezado(value): index
        for index, value in enumerate(values)
        if normalizar_texto(value)
    }
    dane_index = _primer_campo(fields, ENCABEZADOS_DANE)
    if dane_index is None:
        return None
    name_index = _primer_campo(fields, ENCABEZADOS_NOMBRE)
    calendar_index = _primer_campo(fields, ENCABEZADOS_CALENDARIO)
    return int(dane_index), (
        int(name_index) if name_index is not None else None
    ), (
        int(calendar_index) if calendar_index is not None else None
    )


def _entrada_desde_excel(
    values: tuple[object, ...] | list[object],
    indices: tuple[int, int | None, int | None],
    orden: int,
) -> ColegioEntrada | None:
    dane_index, name_index, calendar_index = indices

    def value_at(index: int | None) -> object:
        if index is None or index >= len(values):
            return ""
        return values[index]

    original_dane = normalizar_texto(value_at(dane_index))
    original_name = normalizar_texto(value_at(name_index))
    original_calendar = normalizar_texto(value_at(calendar_index))
    if not original_dane and not original_name and not original_calendar:
        return None
    return ColegioEntrada(
        orden=orden,
        dane=normalizar_dane(original_dane),
        nombre_csv=normalizar_texto(original_name),
        calendario_csv=normalizar_texto(original_calendar).upper(),
        dane_original=original_dane,
        nombre_original=original_name,
        calendario_original=original_calendar,
    )


def _leer_entradas_xls(ruta: Path) -> list[ColegioEntrada]:
    """Lee el formato binario Excel 97-2003 mediante xlrd."""
    try:
        import xlrd
    except ImportError as exc:
        raise IcfesError(
            "Para leer .xls instale las dependencias con: "
            "python -m pip install -r requirements.txt"
        ) from exc

    try:
        book = xlrd.open_workbook(str(ruta), on_demand=True)
    except Exception as exc:
        raise IcfesError(f"No fue posible abrir el archivo XLS: {exc}") from exc
    try:
        for sheet in book.sheets():
            indices: tuple[int, int | None, int | None] | None = None
            header_row = -1
            for row_index in range(min(sheet.nrows, 200)):
                candidate = _indices_encabezado_excel(sheet.row_values(row_index))
                if candidate is not None:
                    indices = candidate
                    header_row = row_index
                    break
            if indices is None:
                continue
            entradas: list[ColegioEntrada] = []
            empty_streak = 0
            for row_index in range(header_row + 1, sheet.nrows):
                entrada = _entrada_desde_excel(
                    sheet.row_values(row_index), indices, row_index + 1
                )
                if entrada is None:
                    empty_streak += 1
                    if entradas and empty_streak >= 100:
                        break
                    continue
                empty_streak = 0
                entradas.append(entrada)
            return entradas
    finally:
        book.release_resources()
    raise IcfesError(
        "No se encontró una columna DANE en las primeras 200 filas de ninguna hoja XLS."
    )


def _leer_entradas_xlsx(ruta: Path) -> list[ColegioEntrada]:
    """Lee libros modernos de Excel en modo de bajo consumo de memoria."""
    try:
        wb = load_workbook(ruta, read_only=True, data_only=True)
    except Exception as exc:
        raise IcfesError(f"No fue posible abrir el archivo XLSX: {exc}") from exc
    try:
        for ws in wb.worksheets:
            indices: tuple[int, int | None, int | None] | None = None
            entradas: list[ColegioEntrada] = []
            empty_streak = 0
            for row_number, values in enumerate(
                ws.iter_rows(values_only=True), start=1
            ):
                if indices is None:
                    if row_number > 200:
                        break
                    indices = _indices_encabezado_excel(values)
                    continue
                entrada = _entrada_desde_excel(values, indices, row_number)
                if entrada is None:
                    empty_streak += 1
                    if entradas and empty_streak >= 100:
                        break
                    continue
                empty_streak = 0
                entradas.append(entrada)
            if indices is not None:
                return entradas
    finally:
        wb.close()
    raise IcfesError(
        "No se encontró una columna DANE en las primeras 200 filas de ninguna hoja XLSX."
    )


def leer_entradas(ruta: Path) -> list[ColegioEntrada]:
    """Lee CSV o Excel y devuelve el mismo modelo normalizado de entrada."""
    suffix = ruta.suffix.lower()
    if suffix == ".xls":
        return _leer_entradas_xls(ruta)
    if suffix in {".xlsx", ".xlsm"}:
        return _leer_entradas_xlsx(ruta)
    if suffix in {".csv", ".txt"}:
        return _leer_entradas_csv(ruta)
    raise IcfesError(
        f"Formato de entrada no compatible: {suffix or 'sin extensión'}. "
        "Use .csv, .txt, .xls, .xlsx o .xlsm."
    )


def formato_dane_valido(dane: str) -> bool:
    """Los códigos DANE de sede usados por el portal tienen 12 dígitos."""
    return bool(re.fullmatch(r"\d{12}", dane))


def parse_partial_response(xml_text: str) -> tuple[dict[str, str], str | None]:
    try:
        root = ElementTree.fromstring(xml_text)
    except ElementTree.ParseError as exc:
        raise IcfesError("El portal no devolvió una respuesta JSF válida.") from exc

    updates: dict[str, str] = {}
    for node in root.findall(".//update"):
        updates[node.attrib.get("id", "")] = node.text or ""
    redirect = root.find(".//redirect")
    return updates, redirect.attrib.get("url") if redirect is not None else None


def extraer_viewstate_html(documento: str) -> str:
    soup = BeautifulSoup(documento, "html.parser")
    node = soup.select_one('input[name="javax.faces.ViewState"]')
    if not node or not node.get("value"):
        raise IcfesError("No se encontró javax.faces.ViewState.")
    return str(node["value"])


class ClienteIcfes:
    def __init__(
        self,
        aplicacion: str = "20263",
        grado: str = "11",
        timeout: float = 45,
        pausa: float = 0,
        control_peticiones: ControlPeticiones | None = None,
    ) -> None:
        self.aplicacion = aplicacion
        self.grado = grado
        self.timeout = timeout
        self.pausa = pausa
        self.control_peticiones = control_peticiones
        self.search_url = urljoin(BASE_URL, SEARCH_PATH)
        self.report_url = urljoin(BASE_URL, REPORT_PATH)

    def _request(
        self,
        session: requests.Session,
        method: str,
        url: str,
        etapa: str = "",
        **kwargs: object,
    ) -> requests.Response:
        last_error: requests.RequestException | None = None
        for attempt in range(1, 4):
            if self.control_peticiones:
                self.control_peticiones.antes_de_solicitar()
            inicio = time.monotonic()
            try:
                response = session.request(method, url, timeout=self.timeout, **kwargs)
                elapsed = time.monotonic() - inicio
                if self.control_peticiones:
                    self.control_peticiones.registrar_respuesta(
                        response.status_code, response.headers.get("Retry-After")
                    )
                logging.debug(
                    "HTTP etapa=%s metodo=%s status=%d intento=%d tiempo=%.2fs",
                    etapa or "sin_etiqueta",
                    method,
                    response.status_code,
                    attempt,
                    elapsed,
                )
                response.raise_for_status()
                break
            except BloqueoIcfesError:
                raise
            except requests.RequestException as exc:
                last_error = exc
                if attempt == 3:
                    if self.control_peticiones:
                        self.control_peticiones.registrar_fallo_transporte(str(exc))
                    raise ErrorTransporteIcfes(
                        f"Fallo HTTP en {etapa or url} después de {attempt} intentos: {exc}"
                    ) from exc
                time.sleep(2**attempt)
        else:  # pragma: no cover - defensa para analizadores estáticos
            raise IcfesError(f"Fallo HTTP: {last_error}")
        # Normaliza la mezcla de UTF-8 de las plantillas con Windows-1252 de
        # algunos nombres institucionales. Se reinyecta UTF-8 para que
        # ``response.text`` sea correcto en el resto del programa.
        response._content = decodificar_respuesta_icfes(response.content).encode("utf-8")
        response.encoding = "utf-8"
        contenido = response.text.lower()
        marcadores_bloqueo = (
            "access denied",
            "acceso denegado",
            "request rejected",
            "cf-chl-captcha",
            "verify you are human",
            "verifique que no es un robot",
            "tráfico inusual",
            "trafico inusual",
            "unusual traffic",
        )
        if any(marker in contenido for marker in marcadores_bloqueo):
            motivo = (
                f"El portal devolvió una página de bloqueo/CAPTCHA durante {etapa or method}."
            )
            if self.control_peticiones:
                self.control_peticiones.detener(motivo)
            raise BloqueoIcfesError(motivo)
        if self.pausa:
            time.sleep(self.pausa)
        return response

    def consultar(self, entrada: ColegioEntrada) -> ResultadoColegio:
        """Ejecuta Buscar -> seleccionar -> Continuar y obtiene el reporte."""
        with requests.Session() as session:
            session.headers.update(
                {
                    "User-Agent": (
                        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                        "AppleWebKit/537.36 ICFES-resultados-institucionales/1.0"
                    ),
                    "Accept-Language": "es-CO,es;q=0.9",
                }
            )

            initial = self._request(session, "GET", self.search_url, etapa="abrir_consulta")
            viewstate = extraer_viewstate_html(initial.text)

            # El cambio de aplicación hace que JSF defina el selector de grado
            # correcto (selGradoTodos_input para las aplicaciones recientes).
            change = self._request(
                session,
                "POST",
                self.search_url,
                etapa="preparar_aplicacion_grado",
                headers=AJAX_HEADERS,
                data={
                    "javax.faces.partial.ajax": "true",
                    "javax.faces.source": "campoAplicacion",
                    "javax.faces.partial.execute": "campoAplicacion",
                    "javax.faces.partial.render": "panel2",
                    "javax.faces.behavior.event": "valueChange",
                    "javax.faces.partial.event": "valueChange",
                    "baseForm": "baseForm",
                    "campoAplicacion_input": self.aplicacion,
                    "javax.faces.ViewState": viewstate,
                },
            )
            updates, _ = parse_partial_response(change.text)
            viewstate = updates.get("javax.faces.ViewState", "")
            panel_grado = BeautifulSoup(updates.get("panel2", ""), "html.parser")
            grado_select = panel_grado.select_one('select[id$="_input"]')
            if not viewstate or not grado_select or not grado_select.get("name"):
                raise IcfesError("No fue posible preparar aplicación y grado.")
            grado_field = str(grado_select["name"])

            search = self._request(
                session,
                "POST",
                self.search_url,
                etapa="buscar_dane",
                headers=AJAX_HEADERS,
                data={
                    "javax.faces.partial.ajax": "true",
                    "javax.faces.source": "j_idt82",
                    "javax.faces.partial.execute": "pnlBuqueda",
                    "javax.faces.partial.render": (
                        "pnlResultados panelFiltros panelCaptcha"
                    ),
                    "j_idt82": "j_idt82",
                    "baseForm": "baseForm",
                    "selTipo": "APLICACION",
                    "campoAplicacion_input": self.aplicacion,
                    "txtNombre": "",
                    grado_field: self.grado,
                    "txtDANE": entrada.dane,
                    "javax.faces.ViewState": viewstate,
                },
            )
            updates, _ = parse_partial_response(search.text)
            viewstate = updates.get("javax.faces.ViewState", "")
            result_html = updates.get("pnlResultados", "")
            result_soup = BeautifulSoup(result_html, "html.parser")
            row = result_soup.select_one("tr[data-rk]")
            if row is None:
                raise IcfesError("Sin resultados para aplicación 2026-1 y grado 11.")
            row_key = texto_limpio(row.get("data-rk"))
            cells = [texto_limpio(td.get_text(" ", strip=True)) for td in row.select("td")]
            nombre_portal = cells[1] if len(cells) > 1 else entrada.nombre_csv
            etc = cells[2] if len(cells) > 2 else ""

            selected = self._request(
                session,
                "POST",
                self.search_url,
                etapa="seleccionar_colegio",
                headers=AJAX_HEADERS,
                data={
                    "javax.faces.partial.ajax": "true",
                    "javax.faces.source": "tblResultados",
                    "javax.faces.partial.execute": "tblResultados",
                    "javax.faces.partial.render": "tblResultados btnConsultar",
                    "javax.faces.behavior.event": "rowSelectRadio",
                    "javax.faces.partial.event": "rowSelectRadio",
                    "baseForm": "baseForm",
                    "tblResultados_selection": row_key,
                    "javax.faces.ViewState": viewstate,
                },
            )
            updates, _ = parse_partial_response(selected.text)
            viewstate = updates.get("javax.faces.ViewState", "")
            if not viewstate:
                raise IcfesError("El portal no aceptó la selección del colegio.")

            continued = self._request(
                session,
                "POST",
                self.search_url,
                etapa="continuar_reporte",
                headers=AJAX_HEADERS,
                data={
                    "javax.faces.partial.ajax": "true",
                    "javax.faces.source": "btnConsultar",
                    "javax.faces.partial.execute": "pnlResultados",
                    "javax.faces.partial.render": "panelMensajesError",
                    "btnConsultar": "btnConsultar",
                    "baseForm": "baseForm",
                    "tblResultados_selection": row_key,
                    "javax.faces.ViewState": viewstate,
                },
            )
            _, redirect = parse_partial_response(continued.text)
            if not redirect:
                raise IcfesError("El portal no generó el reporte institucional.")

            report = self._request(
                session,
                "GET",
                urljoin(BASE_URL, redirect),
                etapa="descargar_reporte",
            )
            if "Ficha técnica" not in report.text or "Lectura crítica" not in report.text:
                raise IcfesError("El reporte recibido está incompleto o la sesión caducó.")

            return ResultadoColegio(
                entrada=entrada,
                nombre_portal=nombre_portal,
                etc=etc,
                html_reporte=report.text,
                fecha_consulta=datetime.now().astimezone().isoformat(timespec="seconds"),
            )


def nombre_tabla(table: Tag, index: int) -> str:
    """Busca el título/sección inmediatamente anterior a una tabla."""
    for previous in table.find_all_previous(limit=20):
        if not isinstance(previous, Tag):
            continue
        if previous.name in {"h1", "h2", "h3", "h4", "legend"}:
            value = texto_limpio(previous.get_text(" ", strip=True))
            if value:
                return value
        classes = set(previous.get("class", []))
        if classes.intersection({"titulo", "subtitulo", "title", "ui-fieldset-legend"}):
            value = texto_limpio(previous.get_text(" ", strip=True))
            if value:
                return value
    return f"tabla_{index}"


def filas_tabla(table: Tag) -> list[tuple[list[str], bool]]:
    """Devuelve filas y marca si cada una pertenece al encabezado HTML."""
    rows: list[tuple[list[str], bool]] = []
    # recursive=False evita duplicar filas cuando una tabla se usa para layout
    # y contiene otras tablas de datos.
    for tr in table.find_all("tr", recursive=False):
        cells = tr.find_all(["th", "td"], recursive=False)
        values = [texto_limpio(cell.get_text(" ", strip=True)) for cell in cells]
        if values and any(values):
            rows.append((values, bool(tr.find_all("th", recursive=False))))
    if not rows:
        for section in table.find_all(["thead", "tbody", "tfoot"], recursive=False):
            for tr in section.find_all("tr", recursive=False):
                cells = tr.find_all(["th", "td"], recursive=False)
                values = [texto_limpio(cell.get_text(" ", strip=True)) for cell in cells]
                if values and any(values):
                    rows.append((values, bool(tr.find_all("th", recursive=False))))
    return rows


def extraer_tablas(resultado: ResultadoColegio) -> list[dict[str, object]]:
    """Normaliza cada celda de cada tabla de las siete pestañas."""
    soup = BeautifulSoup(resultado.html_reporte, "html.parser")
    nav = soup.select_one("ul.ui-tabs-nav")
    panels_container = soup.select_one("div.ui-tabs-panels")
    if nav is None or panels_container is None:
        raise IcfesError("No se encontró la estructura de pestañas del reporte.")

    tab_names = [texto_limpio(li.get_text(" ", strip=True)) for li in nav.select(":scope > li")]
    panels = panels_container.select(":scope > div.ui-tabs-panel")
    records: list[dict[str, object]] = []

    for tab_index, panel in enumerate(panels):
        tab_name = tab_names[tab_index] if tab_index < len(tab_names) else f"Pestaña {tab_index + 1}"
        # Se conservan tablas hoja: si una tabla contiene otras, funciona solo
        # como layout y sus datos aparecerán en las tablas hijas.
        leaf_tables = [table for table in panel.find_all("table") if not table.find("table")]
        table_number = 0
        for table in leaf_tables:
            rows = filas_tabla(table)
            if not rows:
                continue
            table_number += 1
            section = nombre_tabla(table, table_number)
            width = max(len(row) for row, _ in rows)
            # Se guardan también las filas de encabezado. Algunos reportes
            # usan encabezados de dos niveles con rowspan/colspan; conservar la
            # posición original evita inventar asociaciones incorrectas.
            for row_number, (row, is_header) in enumerate(rows, start=1):
                for column_number in range(width):
                    value = row[column_number] if column_number < len(row) else ""
                    if not value:
                        continue
                    records.append(
                        {
                            "dane": resultado.entrada.dane,
                            "pestana": tab_name,
                            "seccion": section,
                            "tabla": table_number,
                            "fila": row_number,
                            "tipo_fila": "encabezado" if is_header else "dato",
                            "columna_num": column_number + 1,
                            "valor": value,
                        }
                    )
    return records


def extraer_textos_clave(resultado: ResultadoColegio) -> list[dict[str, object]]:
    """Conserva encabezados y textos de lectura de resultados, sin ayudas extensas."""
    soup = BeautifulSoup(resultado.html_reporte, "html.parser")
    nav = soup.select_one("ul.ui-tabs-nav")
    panels_container = soup.select_one("div.ui-tabs-panels")
    if nav is None or panels_container is None:
        return []
    tab_names = [texto_limpio(li.get_text(" ", strip=True)) for li in nav.select(":scope > li")]
    panels = panels_container.select(":scope > div.ui-tabs-panel")
    records: list[dict[str, object]] = []
    for tab_index, panel in enumerate(panels):
        tab = tab_names[tab_index] if tab_index < len(tab_names) else f"Pestaña {tab_index + 1}"
        order = 0
        for element in panel.find_all(["h1", "h2", "h3", "h4", "legend", "p"]):
            value = texto_limpio(element.get_text(" ", strip=True))
            if not value or len(value) > 1200:
                continue
            if element.name == "p" and not re.search(r"lectura de resultados|obtenido|porcentaje", value, re.I):
                continue
            order += 1
            records.append(
                {
                    "dane": resultado.entrada.dane,
                    "pestana": tab,
                    "orden": order,
                    "tipo": element.name,
                    "texto": value,
                }
            )
    return records


def valor_numerico_y_simbolo(value: object) -> tuple[int | float | None, str]:
    """Separa el número y el símbolo comparativo usado por el reporte."""
    raw = texto_limpio(value)
    if not raw or raw.upper().startswith("N.D"):
        return None, ""
    symbol = next((item for item in ("▲", "▼", "•") if item in raw), "")
    match = re.search(r"-?\d+(?:[.,]\d+)?", raw)
    if not match:
        return None, symbol
    numeric = float(match.group(0).replace(",", "."))
    return (int(numeric) if numeric.is_integer() else numeric), symbol


def matrices_desde_celdas(
    registros: list[dict[str, object]],
) -> dict[tuple[str, str, int], list[list[str]]]:
    """Reconstruye las matrices HTML a partir de la hoja de auditoría."""
    grouped: dict[tuple[str, str, int], dict[int, dict[int, str]]] = {}
    for record in registros:
        key = (
            str(record["pestana"]),
            str(record["seccion"]),
            int(record["tabla"]),
        )
        rows = grouped.setdefault(key, {})
        row = rows.setdefault(int(record["fila"]), {})
        row[int(record["columna_num"])] = str(record["valor"])

    matrices: dict[tuple[str, str, int], list[list[str]]] = {}
    for key, rows in grouped.items():
        width = max((max(columns) for columns in rows.values()), default=0)
        matrices[key] = [
            [columns.get(column, "") for column in range(1, width + 1)]
            for _, columns in sorted(rows.items())
        ]
    return matrices


def extraer_modelo_power_bi(
    resultado: ResultadoColegio,
    celdas: list[dict[str, object]],
) -> dict[str, list[dict[str, object]]]:
    """Convierte las tablas visuales en hechos numéricos listos para Power BI."""
    dane = resultado.entrada.dane
    common = {"dane": dane, "aplicacion": "2026-1", "grado": 11}
    output: dict[str, list[dict[str, object]]] = {
        "poblacion": [],
        "puntajes": [],
        "niveles": [],
        "aprendizajes": [],
    }
    subjects = {
        "Lectura crítica",
        "Matemáticas",
        "Sociales y ciudadanas",
        "Ciencias naturales",
        "Inglés",
    }

    for (tab, _section, table_number), matrix in matrices_desde_celdas(celdas).items():
        if not matrix:
            continue

        if tab == "Ficha técnica" and table_number == 1:
            for row in matrix[2:]:
                if len(row) < 5:
                    continue
                output["poblacion"].append(
                    {
                        **common,
                        "nivel_agregacion": row[0],
                        "matriculados": valor_numerico_y_simbolo(row[1])[0],
                        "inscritos": valor_numerico_y_simbolo(row[2])[0],
                        "presentes": valor_numerico_y_simbolo(row[3])[0],
                        "resultados_publicados": valor_numerico_y_simbolo(row[4])[0],
                    }
                )
            continue

        if table_number == 1 and (tab == "General" or tab in subjects):
            prueba = "Puntaje global" if tab == "General" else tab
            for row in matrix[1:]:
                if len(row) < 3:
                    continue
                average, average_symbol = valor_numerico_y_simbolo(row[1])
                deviation, deviation_symbol = valor_numerico_y_simbolo(row[2])
                output["puntajes"].append(
                    {
                        **common,
                        "prueba": prueba,
                        "nivel_agregacion": row[0],
                        "promedio": average,
                        "simbolo_promedio": average_symbol,
                        "desviacion": deviation,
                        "simbolo_desviacion": deviation_symbol,
                    }
                )
            continue

        if table_number == 2 and tab in subjects and len(matrix[0]) >= 2:
            levels = matrix[0][1:]
            for row in matrix[1:]:
                if len(row) < 2:
                    continue
                for index, level in enumerate(levels, start=1):
                    if index >= len(row):
                        continue
                    percentage, symbol = valor_numerico_y_simbolo(row[index])
                    output["niveles"].append(
                        {
                            **common,
                            "prueba": tab,
                            "nivel_agregacion": row[0],
                            "nivel_desempeno": level,
                            "porcentaje_estudiantes": percentage,
                            "simbolo_comparacion": symbol,
                        }
                    )
            continue

        if table_number == 3 and tab in subjects and len(matrix[0]) >= 2:
            aggregation_levels = matrix[0][1:]
            aggregation_names = {"EE": "Establecimiento educativo (EE)"}
            for row in matrix[1:]:
                if len(row) < 2 or row[0].lower().startswith("no records"):
                    continue
                for index, aggregation in enumerate(aggregation_levels, start=1):
                    if index >= len(row):
                        continue
                    percentage, _ = valor_numerico_y_simbolo(row[index])
                    output["aprendizajes"].append(
                        {
                            **common,
                            "prueba": tab,
                            "aprendizaje": row[0],
                            "nivel_agregacion": aggregation_names.get(aggregation, aggregation),
                            "porcentaje_respuestas_incorrectas": percentage,
                        }
                    )
    return output


def huella_archivo(ruta: Path) -> str:
    digest = hashlib.sha256()
    with ruta.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def preparar_checkpoint(
    ruta: Path,
    entrada: Path,
    aplicacion: str,
    grado: str,
    reiniciar: bool,
) -> None:
    """Crea o valida el archivo reanudable de la ejecución masiva."""
    meta_path = ruta.with_suffix(ruta.suffix + ".meta.json")
    if reiniciar:
        ruta.unlink(missing_ok=True)
        meta_path.unlink(missing_ok=True)

    expected = {
        "entrada": str(entrada.resolve()),
        "sha256_entrada": huella_archivo(entrada),
        "aplicacion": aplicacion,
        "grado": str(grado),
        "version_modelo": 3,
    }
    if meta_path.exists():
        current = json.loads(meta_path.read_text(encoding="utf-8"))
        if current != expected:
            raise IcfesError(
                f"El progreso {ruta} pertenece a otro archivo o configuración. "
                "Use --reiniciar para comenzar una ejecución nueva."
            )
    elif ruta.exists() and ruta.stat().st_size:
        raise IcfesError(
            f"Existe {ruta} sin metadatos verificables. Muévalo o use --reiniciar."
        )
    else:
        ruta.parent.mkdir(parents=True, exist_ok=True)
        meta_path.write_text(
            json.dumps(expected, indent=2, ensure_ascii=False), encoding="utf-8"
        )


def _meta_checkpoint(ruta: Path) -> dict[str, object]:
    meta_path = ruta.with_suffix(ruta.suffix + ".meta.json")
    if not meta_path.is_file():
        raise IcfesError(
            f"El checkpoint de origen {ruta} no tiene metadatos verificables."
        )
    try:
        return json.loads(meta_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise IcfesError(
            f"No fue posible leer los metadatos de {ruta}: {exc}"
        ) from exc


def reutilizar_resultados_validos(
    origen: Path,
    destino: Path,
    entradas: list[ColegioEntrada],
    aplicacion: str,
    grado: str,
) -> int:
    """Copia al nuevo lote resultados válidos coincidentes por DANE.

    La fila y los datos descriptivos se ajustan al listado nuevo. Los rechazados
    no se importan: pueden corresponder a errores temporales o a duplicados del
    archivo anterior y, por tanto, se vuelven a evaluar en el lote actual.
    """
    if origen.resolve() == destino.resolve():
        raise IcfesError(
            "--reutilizar-checkpoint debe apuntar a un checkpoint distinto del actual."
        )
    if not origen.is_file():
        raise IcfesError(f"No existe el checkpoint para reutilizar: {origen}")

    meta = _meta_checkpoint(origen)
    aplicacion_origen = re.sub(r"\D+", "", str(meta.get("aplicacion", "")))
    aplicacion_destino = re.sub(r"\D+", "", str(aplicacion))
    tipo_origen = str(meta.get("tipo_reporte", "APLICACION")).upper()
    if (
        tipo_origen != "APLICACION"
        or aplicacion_origen != aplicacion_destino
        or str(meta.get("grado", "")) != str(grado)
    ):
        raise IcfesError(
            "El checkpoint para reutilizar pertenece a otra aplicación o grado."
        )

    entrada_por_dane = {
        entrada.dane: entrada
        for entrada in entradas
        if formato_dane_valido(entrada.dane)
    }
    ya_guardados = {
        str(event.get("control", {}).get("dane", ""))
        for event in iterar_eventos(destino)
        if event.get("estado") == "VALIDO"
    }
    importados = 0
    for event in iterar_eventos(origen):
        if event.get("estado") != "VALIDO":
            continue
        control = dict(event.get("control", {}))
        dane = str(control.get("dane", ""))
        entrada = entrada_por_dane.get(dane)
        if entrada is None or dane in ya_guardados:
            continue

        control.update(control_base(entrada))
        control["estado"] = "VALIDO"
        event["control"] = control
        institucion = dict(event.get("institucion", {}))
        institucion.update(
            {
                "dane": entrada.dane,
                "nombre_csv": entrada.nombre_csv,
                "calendario_csv": entrada.calendario_csv,
                "fila_csv": entrada.orden,
            }
        )
        event["institucion"] = institucion
        guardar_evento(destino, event)
        ya_guardados.add(dane)
        importados += 1
    return importados


def iterar_eventos(ruta: Path) -> Iterable[dict[str, object]]:
    if not ruta.exists():
        return
    with ruta.open("r", encoding="utf-8") as fh:
        for line_number, line in enumerate(fh, start=1):
            if not line.strip():
                continue
            try:
                yield json.loads(line)
            except json.JSONDecodeError as exc:
                raise IcfesError(
                    f"El checkpoint está dañado en la línea {line_number}: {exc}"
                ) from exc


def guardar_evento(ruta: Path, event: dict[str, object]) -> None:
    with ruta.open("a", encoding="utf-8", newline="\n") as fh:
        fh.write(json.dumps(event, ensure_ascii=False, separators=(",", ":")) + "\n")
        fh.flush()
        os.fsync(fh.fileno())


def control_base(entrada: ColegioEntrada) -> dict[str, object]:
    return {
        "fila_csv": entrada.orden,
        "dane_original": entrada.dane_original,
        "dane": entrada.dane,
        "nombre_original": entrada.nombre_original,
        "nombre_normalizado": entrada.nombre_csv,
        "calendario_original": entrada.calendario_original,
        "calendario_normalizado": entrada.calendario_csv,
    }


def evento_rechazado(
    entrada: ColegioEntrada, estado: str, motivo: str
) -> dict[str, object]:
    return {
        "estado": estado,
        "control": {
            **control_base(entrada),
            "estado": estado,
            "motivo": motivo,
            "nombre_icfes": "",
            "etc": "",
            "fecha_consulta": "",
        },
    }


def evento_valido(resultado: ResultadoColegio) -> dict[str, object]:
    celdas = extraer_tablas(resultado)
    tabs = {str(item["pestana"]) for item in celdas}
    expected_tabs = {
        "Ficha técnica",
        "General",
        "Lectura crítica",
        "Matemáticas",
        "Sociales y ciudadanas",
        "Ciencias naturales",
        "Inglés",
    }
    missing = sorted(expected_tabs - tabs)
    if missing:
        raise IcfesError(f"Reporte incompleto; faltan pestañas: {', '.join(missing)}")

    model = extraer_modelo_power_bi(resultado, celdas)
    for key in ("poblacion", "puntajes", "niveles"):
        if not model[key]:
            raise IcfesError(f"La normalización no produjo registros de {key}.")

    institution = {
        "dane": resultado.entrada.dane,
        "nombre_csv": resultado.entrada.nombre_csv,
        "nombre_icfes": resultado.nombre_portal,
        "calendario_csv": resultado.entrada.calendario_csv,
        "etc": resultado.etc,
        "aplicacion": "2026-1",
        "grado": 11,
        "fecha_consulta": resultado.fecha_consulta,
        "fila_csv": resultado.entrada.orden,
    }
    return {
        "estado": "VALIDO",
        "control": {
            **control_base(resultado.entrada),
            "estado": "VALIDO",
            "motivo": "",
            "nombre_icfes": resultado.nombre_portal,
            "etc": resultado.etc,
            "fecha_consulta": resultado.fecha_consulta,
            "pestanas_verificadas": len(tabs),
            "filas_poblacion": len(model["poblacion"]),
            "filas_puntajes": len(model["puntajes"]),
            "filas_niveles": len(model["niveles"]),
            "filas_aprendizajes": len(model["aprendizajes"]),
        },
        "institucion": institution,
        "poblacion": model["poblacion"],
        "puntajes": model["puntajes"],
        "niveles": model["niveles"],
        "aprendizajes": model["aprendizajes"],
    }


def _crear_hoja_stream(
    wb: Workbook, name: str, headers: list[str], widths: list[int] | None = None
):
    ws = wb.create_sheet(name)
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    if widths:
        for index, width in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(index)].width = width
    fill = PatternFill("solid", fgColor="0B5A8E")
    cells = []
    for header in headers:
        cell = WriteOnlyCell(ws, value=header)
        cell.font = Font(color="FFFFFF", bold=True)
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cells.append(cell)
    ws.append(cells)
    return ws


def _append_stream(ws, headers: list[str], item: dict[str, object]) -> None:
    row = []
    for header in headers:
        cell = WriteOnlyCell(ws, value=item.get(header, ""))
        if header == "dane":
            cell.number_format = "@"
        row.append(cell)
    ws.append(row)


class TablaExcelParticionada:
    """Escribe una tabla lógica en una o más hojas sin superar Excel."""

    def __init__(
        self,
        wb: Workbook,
        nombre: str,
        headers: list[str],
        max_filas: int = 900_000,
    ) -> None:
        self.wb = wb
        self.nombre = nombre
        self.headers = headers
        self.max_filas = max_filas
        self.hojas: list[tuple[str, object, int]] = []

    def _nueva_hoja(self) -> None:
        parte = len(self.hojas) + 1
        nombre = self.nombre if parte == 1 else f"{self.nombre}_{parte:02d}"
        if len(nombre) > 31:
            nombre = f"{self.nombre[:28]}_{parte:02d}"
        ws = _crear_hoja_stream(self.wb, nombre, self.headers)
        self.hojas.append((nombre, ws, 0))

    def append(self, item: dict[str, object]) -> None:
        if not self.hojas or self.hojas[-1][2] >= self.max_filas:
            self._nueva_hoja()
        nombre, ws, count = self.hojas[-1]
        _append_stream(ws, self.headers, item)
        self.hojas[-1] = (nombre, ws, count + 1)

    def finalizar(self) -> dict[str, int]:
        if not self.hojas:
            self._nueva_hoja()
        counts: dict[str, int] = {}
        for nombre, ws, count in self.hojas:
            if count:
                ws.auto_filter.ref = (
                    f"A1:{get_column_letter(len(self.headers))}{count + 1}"
                )
            counts[nombre] = count
        return counts


def exportar_excel_masivo(
    ruta: Path,
    checkpoint: Path,
    entrada: Path,
    total_entrada: int,
) -> tuple[dict[str, int], dict[str, int]]:
    """Genera un Excel particionado con el modelo Power BI y hojas de control."""
    specs: dict[str, list[str]] = {
        "Dim_Instituciones": [
            "dane", "nombre_csv", "nombre_icfes", "calendario_csv", "etc",
            "aplicacion", "grado", "fecha_consulta", "fila_csv",
        ],
        "Hecho_Poblacion": [
            "dane", "aplicacion", "grado", "nivel_agregacion", "matriculados",
            "inscritos", "presentes", "resultados_publicados",
        ],
        "Hecho_Puntajes": [
            "dane", "aplicacion", "grado", "prueba", "nivel_agregacion",
            "promedio", "simbolo_promedio", "desviacion", "simbolo_desviacion",
        ],
        "Hecho_Niveles": [
            "dane", "aplicacion", "grado", "prueba", "nivel_agregacion",
            "nivel_desempeno", "porcentaje_estudiantes", "simbolo_comparacion",
        ],
        "Hecho_Aprendizajes": [
            "dane", "aplicacion", "grado", "prueba", "aprendizaje",
            "nivel_agregacion", "porcentaje_respuestas_incorrectas",
        ],
        "Control_Proceso": [
            "fila_csv", "dane_original", "dane", "nombre_original",
            "nombre_normalizado", "calendario_original", "calendario_normalizado",
            "estado", "motivo", "nombre_icfes", "etc", "fecha_consulta",
            "pestanas_verificadas", "filas_poblacion", "filas_puntajes",
            "filas_niveles", "filas_aprendizajes",
        ],
        "Rechazados": [
            "fila_csv", "dane_original", "dane", "nombre_original",
            "nombre_normalizado", "calendario_original", "calendario_normalizado",
            "estado", "motivo",
        ],
    }
    wb = Workbook(write_only=True)
    tables = {
        name: TablaExcelParticionada(wb, name, headers)
        for name, headers in specs.items()
    }
    logical_counts = {name: 0 for name in specs}
    states: dict[str, int] = {}

    for event in iterar_eventos(checkpoint):
        state = str(event["estado"])
        states[state] = states.get(state, 0) + 1
        control = dict(event["control"])
        tables["Control_Proceso"].append(control)
        logical_counts["Control_Proceso"] += 1
        if state != "VALIDO":
            tables["Rechazados"].append(control)
            logical_counts["Rechazados"] += 1
            continue

        tables["Dim_Instituciones"].append(dict(event["institucion"]))
        logical_counts["Dim_Instituciones"] += 1
        for key, sheet_name in (
            ("poblacion", "Hecho_Poblacion"),
            ("puntajes", "Hecho_Puntajes"),
            ("niveles", "Hecho_Niveles"),
            ("aprendizajes", "Hecho_Aprendizajes"),
        ):
            for item in event.get(key, []):
                tables[sheet_name].append(dict(item))
                logical_counts[sheet_name] += 1

    physical_counts: dict[str, int] = {}
    for table in tables.values():
        physical_counts.update(table.finalizar())

    summary_headers = ["indicador", "valor"]
    summary = _crear_hoja_stream(wb, "Resumen_Ejecucion", summary_headers, [38, 70])
    summary_rows = [
        {"indicador": "archivo_entrada", "valor": str(entrada.resolve())},
        {"indicador": "filas_entrada", "valor": total_entrada},
        {"indicador": "registros_procesados", "valor": logical_counts["Control_Proceso"]},
        {"indicador": "instituciones_validas", "valor": logical_counts["Dim_Instituciones"]},
        {"indicador": "registros_rechazados", "valor": logical_counts["Rechazados"]},
        {"indicador": "trabajadores_configurados", "valor": 100},
        {"indicador": "filas_maximas_por_hoja", "valor": 900000},
        *(
            {"indicador": f"estado_{key}", "valor": value}
            for key, value in sorted(states.items())
        ),
        *(
            {"indicador": f"filas_{key}", "valor": value}
            for key, value in logical_counts.items()
        ),
    ]
    for item in summary_rows:
        _append_stream(summary, summary_headers, item)
    summary.auto_filter.ref = f"A1:B{len(summary_rows) + 1}"
    physical_counts["Resumen_Ejecucion"] = len(summary_rows)

    model_headers = [
        "tabla", "granularidad", "relacion_recomendada", "hojas_fisicas"
    ]
    model = _crear_hoja_stream(
        wb, "Modelo_PowerBI", model_headers, [28, 68, 52, 60]
    )
    model_rows = [
        {"tabla": "Dim_Instituciones", "granularidad": "Una fila por DANE válido", "relacion_recomendada": "Clave primaria: dane"},
        {"tabla": "Hecho_Poblacion", "granularidad": "DANE + aplicación + grado + nivel de agregación", "relacion_recomendada": "Muchos a uno con Dim_Instituciones[dane]"},
        {"tabla": "Hecho_Puntajes", "granularidad": "DANE + prueba + nivel de agregación", "relacion_recomendada": "Muchos a uno con Dim_Instituciones[dane]"},
        {"tabla": "Hecho_Niveles", "granularidad": "DANE + prueba + agregación + nivel", "relacion_recomendada": "Muchos a uno con Dim_Instituciones[dane]"},
        {"tabla": "Hecho_Aprendizajes", "granularidad": "DANE + prueba + aprendizaje + agregación", "relacion_recomendada": "Muchos a uno con Dim_Instituciones[dane]"},
    ]
    for item in model_rows:
        table_name = str(item["tabla"])
        item["hojas_fisicas"] = ", ".join(
            name
            for name in physical_counts
            if name == table_name or name.startswith(f"{table_name}_")
        )
        _append_stream(model, model_headers, item)
    model.auto_filter.ref = f"A1:D{len(model_rows) + 1}"
    physical_counts["Modelo_PowerBI"] = len(model_rows)

    ruta.parent.mkdir(parents=True, exist_ok=True)
    wb.save(ruta)
    return logical_counts, physical_counts


def validar_excel_masivo(ruta: Path, counts: dict[str, int]) -> None:
    """Reabre el libro y verifica todas las hojas y sus conteos físicos."""
    try:
        wb = load_workbook(ruta, read_only=True, data_only=True)
    except Exception as exc:
        raise IcfesError(f"El Excel generado no se puede reabrir: {exc}") from exc
    missing = sorted(set(counts) - set(wb.sheetnames))
    if missing:
        wb.close()
        raise IcfesError(f"Al Excel generado le faltan hojas: {', '.join(missing)}")
    for name, expected in counts.items():
        actual = max(sum(1 for _ in wb[name].iter_rows()) - 1, 0)
        if actual != expected:
            wb.close()
            raise IcfesError(
                f"Verificación fallida en {name}: {actual} filas; se esperaban {expected}."
            )
    wb.close()


def generar_excel_atomico(
    ruta: Path,
    checkpoint: Path,
    entrada: Path,
    total_entrada: int,
    espacio_minimo_gb: float,
) -> dict[str, int]:
    """Construye, valida y publica el libro sin exponer una salida incompleta."""
    ruta.parent.mkdir(parents=True, exist_ok=True)
    free = shutil.disk_usage(ruta.parent).free
    minimum = int(espacio_minimo_gb * 1024**3)
    if free < minimum:
        raise IcfesError(
            f"Espacio insuficiente: {free / 1024**3:.2f} GB libres; "
            f"se requieren al menos {espacio_minimo_gb:.2f} GB."
        )
    temporal = ruta.with_name(f"{ruta.stem}.en_construccion{ruta.suffix}")
    temporal.unlink(missing_ok=True)
    logical_counts, physical_counts = exportar_excel_masivo(
        temporal, checkpoint, entrada, total_entrada
    )
    validar_excel_masivo(temporal, physical_counts)
    try:
        os.replace(temporal, ruta)
    except OSError as exc:
        raise IcfesError(
            f"El Excel fue validado pero no pudo publicarse en {ruta}: {exc}. "
            f"La copia temporal permanece en {temporal}."
        ) from exc
    return logical_counts


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--entrada",
        type=Path,
        default=None,
        help="Listado CSV/XLS/XLSX. Si se omite, se solicitará la ruta.",
    )
    parser.add_argument("--salida", type=Path, default=None)
    parser.add_argument(
        "--cantidad",
        type=int,
        default=0,
        help="Cantidad de colegios válidos; 0 procesa todo.",
    )
    parser.add_argument("--aplicacion", default="20261", help="20261 equivale a 2026-1")
    parser.add_argument("--grado", default="11")
    parser.add_argument(
        "--hilos",
        type=int,
        default=100,
        help="Trabajadores concurrentes. El máximo permitido es 100.",
    )
    parser.add_argument(
        "--rps",
        type=float,
        default=0.5,
        help="Límite global de solicitudes HTTP por segundo entre todos los hilos.",
    )
    parser.add_argument(
        "--max-429",
        type=int,
        default=3,
        help="Cantidad de HTTP 429 antes de detener el lote.",
    )
    parser.add_argument(
        "--max-fallos-transporte",
        type=int,
        default=8,
        help="Fallos HTTP/red acumulados antes de activar el cortacircuitos.",
    )
    parser.add_argument(
        "--espacio-minimo-gb",
        type=float,
        default=2.0,
        help="Espacio libre mínimo requerido antes de construir el Excel.",
    )
    parser.add_argument(
        "--checkpoint",
        type=Path,
        default=None,
        help="JSONL de progreso; por defecto se crea junto al Excel.",
    )
    parser.add_argument(
        "--reutilizar-checkpoint",
        type=Path,
        default=None,
        help=(
            "Checkpoint anterior del que se copian resultados válidos que "
            "coincidan por DANE, aplicación y grado."
        ),
    )
    parser.add_argument(
        "--reiniciar",
        action="store_true",
        help="Descarta el checkpoint compatible y comienza desde cero.",
    )
    parser.add_argument(
        "--solo-exportar",
        action="store_true",
        help="No consulta ICFES; reconstruye el Excel desde el checkpoint.",
    )
    return parser.parse_args(argv)


def solicitar_ruta_entrada() -> Path:
    print("\nIngrese la ruta completa del listado CSV o Excel con los colegios.")
    print("Puede arrastrar el archivo a esta ventana y presionar Enter.\n")
    while True:
        try:
            raw = input("Ruta del listado: ").strip().strip('"').strip("'")
        except EOFError as exc:
            raise IcfesError(
                "No se recibió una ruta. Use --entrada al ejecutar sin consola."
            ) from exc
        path = Path(raw).expanduser()
        if path.is_file():
            return path.resolve()
        print(f"No se encontró el archivo: {path}. Intente nuevamente.")


def procesar_colegio(
    entrada: ColegioEntrada,
    args: argparse.Namespace,
    control: ControlPeticiones,
) -> dict[str, object]:
    """Consulta y normaliza un colegio; nunca escribe archivos desde el hilo."""
    cliente = ClienteIcfes(
        aplicacion=args.aplicacion,
        grado=args.grado,
        pausa=0,
        control_peticiones=control,
    )
    try:
        return evento_valido(cliente.consultar(entrada))
    except (BloqueoIcfesError, ErrorTransporteIcfes):
        raise
    except Exception as exc:
        error_text = texto_limpio(exc)
        state = (
            "SIN_RESULTADOS_ICFES"
            if "sin resultados" in error_text.lower()
            else "ERROR_CONSULTA_O_NORMALIZACION"
        )
        return evento_rechazado(entrada, state, error_text)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    if args.cantidad < 0:
        raise IcfesError("--cantidad no puede ser negativa; use 0 para todo.")
    if not 1 <= args.hilos <= 100:
        raise IcfesError("--hilos debe estar entre 1 y 100.")
    if args.rps <= 0:
        raise IcfesError("--rps debe ser mayor que cero.")
    if args.max_429 < 1 or args.max_fallos_transporte < 1:
        raise IcfesError("Los umbrales de errores deben ser mayores que cero.")
    if args.espacio_minimo_gb < 0:
        raise IcfesError("--espacio-minimo-gb no puede ser negativo.")
    if args.solo_exportar and args.reiniciar:
        raise IcfesError("--solo-exportar y --reiniciar no pueden usarse juntos.")

    entrada_path = args.entrada.resolve() if args.entrada else solicitar_ruta_entrada()
    if not entrada_path.is_file():
        raise IcfesError(f"No existe el archivo de entrada: {entrada_path}")

    base_dir = Path(__file__).resolve().parent
    if args.salida:
        output_path = args.salida.resolve()
    else:
        suffix = f"_{args.cantidad}_colegios" if args.cantidad else "_completo"
        output_path = base_dir / f"resultados_icfes_2026_100_hilos{suffix}.xlsx"
    checkpoint = (
        args.checkpoint.resolve()
        if args.checkpoint
        else output_path.with_suffix(".progreso.jsonl")
    )
    log_path = output_path.with_suffix(".log")
    log_path.parent.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(threadName)s | %(message)s",
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler(log_path, encoding="utf-8"),
        ],
        force=True,
    )

    entradas = leer_entradas(entrada_path)
    if not entradas:
        raise IcfesError("El listado no contiene filas para procesar.")
    preparar_checkpoint(
        checkpoint,
        entrada_path,
        args.aplicacion,
        args.grado,
        args.reiniciar,
    )
    if args.reutilizar_checkpoint:
        origen_reutilizable = args.reutilizar_checkpoint.resolve()
        importados = reutilizar_resultados_validos(
            origen_reutilizable,
            checkpoint,
            entradas,
            args.aplicacion,
            args.grado,
        )
        logging.info(
            "Resultados válidos reutilizados por DANE desde %s: %d.",
            origen_reutilizable,
            importados,
        )

    processed_rows: set[int] = set()
    seen_danes: set[str] = set()
    valid_count = 0
    for event in iterar_eventos(checkpoint):
        control_evento = dict(event["control"])
        processed_rows.add(int(control_evento["fila_csv"]))
        dane = str(control_evento.get("dane", ""))
        if (
            dane
            and formato_dane_valido(dane)
            and event["estado"] != "RECHAZADO_CALENDARIO"
        ):
            seen_danes.add(dane)
        if event["estado"] == "VALIDO":
            valid_count += 1

    logging.info(
        "Listado normalizado: %d filas; %d procesadas; %d válidas previas.",
        len(entradas),
        len(processed_rows),
        valid_count,
    )

    if args.solo_exportar:
        if not checkpoint.exists() or not checkpoint.stat().st_size:
            raise IcfesError("No existe un checkpoint con datos para exportar.")
        counts = generar_excel_atomico(
            output_path,
            checkpoint,
            entrada_path,
            len(entradas),
            args.espacio_minimo_gb,
        )
        logging.info("Excel reconstruido y verificado: %s", output_path)
        logging.info("Conteos: %s", counts)
        return 0

    candidatos: list[ColegioEntrada] = []
    for entrada in entradas:
        if entrada.orden in processed_rows:
            continue
        event: dict[str, object] | None = None
        if not entrada.dane:
            event = evento_rechazado(
                entrada, "RECHAZADO_DANE_VACIO", "El campo DANE está vacío."
            )
        elif not formato_dane_valido(entrada.dane):
            event = evento_rechazado(
                entrada,
                "RECHAZADO_FORMATO_DANE",
                "Formato inválido: se esperaban exactamente 12 dígitos.",
            )
        elif entrada.calendario_csv and entrada.calendario_csv != "B":
            event = evento_rechazado(
                entrada,
                "RECHAZADO_CALENDARIO",
                "La aplicación 2026-1 corresponde al calendario B.",
            )
        elif entrada.dane in seen_danes:
            event = evento_rechazado(
                entrada,
                "RECHAZADO_DUPLICADO",
                "El DANE ya apareció en una fila anterior compatible.",
            )
        else:
            seen_danes.add(entrada.dane)
            candidatos.append(entrada)

        if event is not None:
            guardar_evento(checkpoint, event)
            processed_rows.add(entrada.orden)

    logging.info(
        "Candidatos pendientes: %d. Se usarán %d hilos y máximo %.2f solicitudes/s.",
        len(candidatos),
        args.hilos,
        args.rps,
    )
    control_http = ControlPeticiones(
        args.rps,
        max_429=args.max_429,
        max_fallos_transporte=args.max_fallos_transporte,
    )
    cola: list[tuple[ColegioEntrada, int]] = [(item, 0) for item in candidatos]
    siguiente = 0
    futures: dict[
        concurrent.futures.Future[dict[str, object]], tuple[ColegioEntrada, int]
    ] = {}
    unresolved: list[ColegioEntrada] = []
    bloqueo = False

    def llenar_cola(executor: concurrent.futures.ThreadPoolExecutor) -> None:
        nonlocal siguiente
        if args.cantidad:
            cupo_validos = max(args.cantidad - valid_count, 0)
            limite = min(args.hilos, cupo_validos)
        else:
            limite = args.hilos
        while (
            not control_http.detenido
            and len(futures) < limite
            and siguiente < len(cola)
        ):
            entrada, retry = cola[siguiente]
            siguiente += 1
            future = executor.submit(procesar_colegio, entrada, args, control_http)
            futures[future] = (entrada, retry)
            logging.info(
                "Encolado DANE %s (fila %d, intento de lote %d).",
                entrada.dane,
                entrada.orden,
                retry + 1,
            )

    with concurrent.futures.ThreadPoolExecutor(
        max_workers=args.hilos,
        thread_name_prefix="icfes",
    ) as executor:
        llenar_cola(executor)
        while futures:
            done, _ = concurrent.futures.wait(
                futures, return_when=concurrent.futures.FIRST_COMPLETED
            )
            for future in done:
                entrada, retry = futures.pop(future)
                try:
                    event = future.result()
                except BloqueoIcfesError as exc:
                    bloqueo = True
                    control_http.detener(str(exc))
                    logging.error(
                        "DANE %s queda pendiente; no se registrará como rechazado.",
                        entrada.dane,
                    )
                    continue
                except ErrorTransporteIcfes as exc:
                    if retry < 1 and not control_http.detenido:
                        cola.append((entrada, retry + 1))
                        logging.warning(
                            "DANE %s se reprogramará una vez: %s", entrada.dane, exc
                        )
                    else:
                        unresolved.append(entrada)
                        logging.error(
                            "DANE %s queda pendiente por error temporal: %s",
                            entrada.dane,
                            exc,
                        )
                    continue
                except Exception as exc:
                    # Defensa: un error de programación queda auditado, no rompe el escritor.
                    event = evento_rechazado(
                        entrada,
                        "ERROR_INTERNO_TRABAJADOR",
                        texto_limpio(exc),
                    )

                guardar_evento(checkpoint, event)
                processed_rows.add(entrada.orden)
                if event["estado"] == "VALIDO":
                    valid_count += 1
                    logging.info(
                        "Válido %d%s: %s",
                        valid_count,
                        f"/{args.cantidad}" if args.cantidad else "",
                        entrada.dane,
                    )
                else:
                    logging.warning(
                        "DANE %s registrado con estado %s.",
                        entrada.dane,
                        event["estado"],
                    )

            if control_http.detenido:
                bloqueo = True
                for future in futures:
                    future.cancel()
                break
            llenar_cola(executor)

    if bloqueo:
        logging.error(
            "Ejecución pausada por protección. Checkpoint seguro: %s. Motivo: %s",
            checkpoint,
            control_http.motivo,
        )
        return 4
    if unresolved:
        logging.error(
            "%d colegios quedaron pendientes por red y no fueron marcados como rechazados. "
            "Ejecute nuevamente el mismo comando.",
            len(unresolved),
        )
        return 5

    logging.info("Construyendo y validando Excel desde %s", checkpoint)
    counts = generar_excel_atomico(
        output_path,
        checkpoint,
        entrada_path,
        len(entradas),
        args.espacio_minimo_gb,
    )
    logging.info(
        "Excel generado: %s (%d válidos, %d rechazados)",
        output_path,
        counts["Dim_Instituciones"],
        counts["Rechazados"],
    )
    if args.cantidad and counts["Dim_Instituciones"] < args.cantidad:
        logging.warning(
            "Solo se lograron %d de %d colegios.",
            counts["Dim_Instituciones"],
            args.cantidad,
        )
        return 3
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        logging.warning(
            "Ejecución interrumpida. El progreso ya guardado se reanudará con el mismo comando."
        )
        raise SystemExit(130)
    except IcfesError as exc:
        logging.error("%s", exc)
        raise SystemExit(1)




