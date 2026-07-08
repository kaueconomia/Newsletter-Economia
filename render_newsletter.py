from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import time
import html as html_lib
import urllib.parse
import urllib.request
import zipfile
from copy import deepcopy
from datetime import datetime
from html.parser import HTMLParser
from calendar import monthrange
from pathlib import Path
from typing import Any
from xml.etree import ElementTree
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from jinja2 import Environment, FileSystemLoader, select_autoescape


BASE_DIR = Path(__file__).resolve().parent
DATA_FILE = BASE_DIR / "data" / "newsletter_exemplo.json"
ENV_FILE = BASE_DIR / "env"
SUEP_RELEASE_CALENDAR_FILE = BASE_DIR / "data" / "release_calendar_suep.json"
PRICE_INDEX_WORKBOOK = Path(
    r"C:\Users\SAMSUNG\OneDrive - Sindicato da Ind da Const Civl do Estado de SP\BI\Índices de Preços\INDICES_COMPLETO.xlsx"
)
TEMPLATE_DIR = BASE_DIR / "templates"
TEMPLATE_NAME = "newsletter_web.html"
OUTPUT_FILE = BASE_DIR / "output" / "newsletter_preview.html"
MORE_NEWS_TEMPLATE_NAME = "mais_noticias.html"
MORE_NEWS_OUTPUT_FILE = BASE_DIR / "output" / "mais_noticias.html"
REPORTS_TEMPLATE_NAME = "relatorios.html"
REPORTS_OUTPUT_FILE = BASE_DIR / "output" / "relatorios.html"
DATABASE_TEMPLATE_NAME = "banco_de_dados.html"
DATABASE_OUTPUT_FILE = BASE_DIR / "output" / "banco_de_dados.html"
INDICATOR_CACHE_FILE = BASE_DIR / "data" / "cache" / "market_indicators.json"
DOCS_DIR = BASE_DIR / "docs"
DOCS_STATIC_DIR = DOCS_DIR / "static"
DOCS_INDEX_FILE = DOCS_DIR / "index.html"
DOCS_MORE_NEWS_FILE = DOCS_DIR / "mais_noticias.html"
DOCS_REPORTS_FILE = DOCS_DIR / "relatorios.html"
DOCS_DATABASE_FILE = DOCS_DIR / "banco_de_dados.html"
VALID_DIRECTIONS = {"alta", "queda", "neutra"}
BRAPI_BASE_URL = "https://brapi.dev/api/v2"
AWESOMEAPI_BASE_URL = "https://economia.awesomeapi.com.br/json/last"
BCB_SELIC_URL = "https://api.bcb.gov.br/dados/serie/bcdata.sgs.432/dados/ultimos/1"
BCB_SITE_SERVICE_BASE_URL = "https://www.bcb.gov.br/api/servico/sitebcb"
BLOG_IBRE_URL = "https://blogdoibre.fgv.br/"
FAZENDA_CONJUNTURA_URL = "https://www.gov.br/fazenda/pt-br/central-de-conteudo/publicacoes/conjuntura-economica"
IPEA_CATEGORY_URLS = {
    "Visão Geral": "https://www.ipea.gov.br/cartadeconjuntura/index.php/category/sumario-executivo/",
    "Atividade Econômica": "https://www.ipea.gov.br/cartadeconjuntura/index.php/category/atividade-economica/",
    "Inflação": "https://www.ipea.gov.br/cartadeconjuntura/index.php/category/inflacao/",
    "Moeda e Crédito": "https://www.ipea.gov.br/cartadeconjuntura/index.php/category/moeda-e-credito/",
    "Finanças Públicas": "https://www.ipea.gov.br/cartadeconjuntura/index.php/category/financas-publicas/",
}
IBGE_CATALOG_REPORTS = (
    {
        "titulo": "PIM-PF - Pesquisa Industrial Mensal",
        "categoria": "IBGE",
        "catalog_id": "7228",
        "descricao": "Relatório mais recente da Pesquisa Industrial Mensal - Produção Física do IBGE.",
    },
    {
        "titulo": "PMC - Pesquisa Mensal de Comércio",
        "categoria": "IBGE",
        "catalog_id": "7230",
        "descricao": "Relatório mais recente da Pesquisa Mensal de Comércio do IBGE.",
    },
    {
        "titulo": "PMS - Pesquisa Mensal de Serviços",
        "categoria": "IBGE",
        "catalog_id": "72419",
        "descricao": "Relatório mais recente da Pesquisa Mensal de Serviços do IBGE.",
    },
)
MARKET_TIMEOUT_SECONDS = 8
NETWORK_RETRY_DELAY_SECONDS = 3
DEFAULT_NETWORK_ATTEMPTS = 2
SGS_NETWORK_ATTEMPTS = 3
RELEASE_CALENDAR_YEAR = 2026
RELEASE_CALENDAR_DEFAULT_START_MONTH = 6
CALENDAR_WEEKDAYS = ("S", "T", "Q", "Q", "S", "S", "D")
MONTHS_PT_BR = (
    "janeiro",
    "fevereiro",
    "março",
    "abril",
    "maio",
    "junho",
    "julho",
    "agosto",
    "setembro",
    "outubro",
    "novembro",
    "dezembro",
)
RELEASE_CALENDAR_EVENTS = (
    {"date": "2026-01-29", "label": "Novo Caged - competência dezembro de 2025"},
    {"date": "2026-03-03", "label": "Novo Caged - competência janeiro de 2026"},
    {"date": "2026-03-31", "label": "Novo Caged - competência fevereiro de 2026"},
    {"date": "2026-04-30", "label": "Novo Caged - competência março de 2026"},
    {"date": "2026-05-28", "label": "Novo Caged - competência abril de 2026"},
    {"date": "2026-06-30", "label": "Novo Caged - competência maio de 2026"},
    {"date": "2026-07-10", "label": "INPC - referência 06/2026"},
    {"date": "2026-07-10", "label": "IPCA - referência 06/2026"},
    {"date": "2026-07-10", "label": "PIM Regional - referência 05/2026"},
    {"date": "2026-07-10", "label": "SINAPI - referência 06/2026"},
    {"date": "2026-07-13", "label": "Logística dos Transportes - referência 2024"},
    {"date": "2026-07-14", "label": "LSPA - referência 06/2026"},
    {"date": "2026-07-15", "label": "Pesquisa Mensal de Serviços - referência 05/2026"},
    {"date": "2026-07-16", "label": "Pesquisa Mensal de Comércio - referência 05/2026"},
    {"date": "2026-07-28", "label": "IPCA-15 - referência 07/2026"},
    {"date": "2026-07-30", "label": "Novo Caged - competência junho de 2026"},
    {"date": "2026-07-30", "label": "PNAD Contínua Mensal - referência 06/2026"},
    {"date": "2026-07-31", "label": "IPP - referência 06/2026"},
    {"date": "2026-08-04", "label": "PIM Brasil - referência 06/2026"},
    {"date": "2026-08-11", "label": "INPC - referência 07/2026"},
    {"date": "2026-08-11", "label": "IPCA - referência 07/2026"},
    {"date": "2026-08-11", "label": "PIM Regional - referência 06/2026"},
    {"date": "2026-08-11", "label": "SINAPI - referência 07/2026"},
    {"date": "2026-08-12", "label": "Pesquisa Mensal de Serviços - referência 06/2026"},
    {"date": "2026-08-13", "label": "LSPA - referência 07/2026"},
    {"date": "2026-08-13", "label": "Pesquisa Mensal de Comércio - referência 06/2026"},
    {"date": "2026-08-14", "label": "PNAD Contínua Trimestral - referência 04/2026 a 06/2026"},
    {"date": "2026-08-19", "label": "Pesquisas Trimestrais Agropecuárias - primeiros resultados"},
    {"date": "2026-08-26", "label": "IPCA-15 - referência 08/2026"},
    {"date": "2026-08-27", "label": "PNAD Contínua Mensal - referência 07/2026"},
    {"date": "2026-08-28", "label": "Novo Caged - competência julho de 2026"},
    {"date": "2026-08-28", "label": "IPP - referência 07/2026"},
    {"date": "2026-09-01", "label": "Contas Nacionais Trimestrais - referência 04/2026 a 06/2026"},
    {"date": "2026-09-02", "label": "PIM Brasil - referência 07/2026"},
    {"date": "2026-09-10", "label": "PIM Regional - referência 07/2026"},
    {"date": "2026-09-10", "label": "Pesquisa Mensal de Serviços - referência 07/2026"},
    {"date": "2026-09-11", "label": "INPC - referência 08/2026"},
    {"date": "2026-09-11", "label": "IPCA - referência 08/2026"},
    {"date": "2026-09-11", "label": "SINAPI - referência 08/2026"},
    {"date": "2026-09-15", "label": "LSPA - referência 08/2026"},
    {"date": "2026-09-15", "label": "Pesquisa Mensal de Comércio - referência 07/2026"},
    {"date": "2026-09-15", "label": "Pesquisas Trimestrais Agropecuárias - referência 04/2026 a 06/2026"},
    {"date": "2026-09-25", "label": "IPCA-15 - referência 09/2026"},
    {"date": "2026-09-25", "label": "IPCA-E - referência 07/2026 a 09/2026"},
    {"date": "2026-09-29", "label": "Novo Caged - competência agosto de 2026"},
    {"date": "2026-09-29", "label": "PNAD Contínua Mensal - referência 08/2026"},
    {"date": "2026-09-30", "label": "IPP - referência 08/2026"},
    {"date": "2026-10-02", "label": "PIM Brasil - referência 08/2026"},
    {"date": "2026-10-08", "label": "PIM Regional - referência 08/2026"},
    {"date": "2026-10-09", "label": "INPC - referência 09/2026"},
    {"date": "2026-10-09", "label": "IPCA - referência 09/2026"},
    {"date": "2026-10-09", "label": "SINAPI - referência 09/2026"},
    {"date": "2026-10-14", "label": "Pesquisa Mensal de Serviços - referência 08/2026"},
    {"date": "2026-10-15", "label": "LSPA - referência 09/2026"},
    {"date": "2026-10-15", "label": "Pesquisa Mensal de Comércio - referência 08/2026"},
    {"date": "2026-10-23", "label": "IPCA-15 - referência 10/2026"},
    {"date": "2026-10-27", "label": "IPP - referência 09/2026"},
    {"date": "2026-10-29", "label": "Novo Caged - competência setembro de 2026"},
    {"date": "2026-10-30", "label": "PNAD Contínua Mensal - referência 09/2026"},
    {"date": "2026-11-05", "label": "PIM Brasil - referência 09/2026"},
    {"date": "2026-11-11", "label": "PIM Regional - referência 09/2026"},
    {"date": "2026-11-11", "label": "Pesquisa Mensal de Serviços - referência 09/2026"},
    {"date": "2026-11-12", "label": "INPC - referência 10/2026"},
    {"date": "2026-11-12", "label": "IPCA - referência 10/2026"},
    {"date": "2026-11-12", "label": "SINAPI - referência 10/2026"},
    {"date": "2026-11-13", "label": "LSPA - referência 10/2026"},
    {"date": "2026-11-13", "label": "Pesquisa de Estoques - referência 01/2026 a 06/2026"},
    {"date": "2026-11-13", "label": "Pesquisa Mensal de Comércio - referência 09/2026"},
    {"date": "2026-11-13", "label": "Prognóstico da Safra - 1º prognóstico 2027"},
    {"date": "2026-11-18", "label": "PNAD Contínua Trimestral - referência 07/2026 a 09/2026"},
    {"date": "2026-11-19", "label": "Pesquisas Trimestrais Agropecuárias - primeiros resultados"},
    {"date": "2026-11-26", "label": "IPP - referência 10/2026"},
    {"date": "2026-11-26", "label": "IPCA-15 - referência 11/2026"},
    {"date": "2026-11-27", "label": "PNAD Contínua Mensal - referência 10/2026"},
    {"date": "2026-11-30", "label": "Novo Caged - competência outubro de 2026"},
    {"date": "2026-12-02", "label": "Contas Nacionais Trimestrais - referência 07/2026 a 09/2026"},
    {"date": "2026-12-03", "label": "PIM Brasil - referência 10/2026"},
    {"date": "2026-12-08", "label": "Pesquisa Mensal de Comércio - referência 10/2026"},
    {"date": "2026-12-09", "label": "PIM Regional - referência 10/2026"},
    {"date": "2026-12-10", "label": "Pesquisa Mensal de Serviços - referência 10/2026"},
    {"date": "2026-12-11", "label": "INPC - referência 11/2026"},
    {"date": "2026-12-11", "label": "IPCA - referência 11/2026"},
    {"date": "2026-12-11", "label": "SINAPI - referência 11/2026"},
    {"date": "2026-12-15", "label": "LSPA - referência 11/2026"},
    {"date": "2026-12-15", "label": "Pesquisas Trimestrais Agropecuárias - referência 07/2026 a 09/2026"},
    {"date": "2026-12-15", "label": "Prognóstico da Safra - 2º prognóstico 2027"},
    {"date": "2026-12-23", "label": "IPCA-15 - referência 12/2026"},
    {"date": "2026-12-23", "label": "IPCA-E - referência 10/2026 a 12/2026"},
    {"date": "2026-12-29", "label": "PNAD Contínua Mensal - referência 11/2026"},
    {"date": "2026-12-30", "label": "Novo Caged - competência novembro de 2026"},
)
STATIC_INDICATOR_ORDER = (
    "ipca 12m",
    "ipca",
    "incc",
    "incc-m",
    "igp-m",
)


def load_newsletter_data(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as file:
        return json.load(file)


def load_project_env(path: Path = ENV_FILE) -> None:
    if not path.exists():
        return

    for line in path.read_text(encoding="utf-8").splitlines():
        clean_line = line.strip()
        if not clean_line or clean_line.startswith("#") or "=" not in clean_line:
            continue

        key, value = clean_line.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


def get_current_datetime() -> datetime:
    try:
        return datetime.now(ZoneInfo("America/Sao_Paulo"))
    except ZoneInfoNotFoundError:
        return datetime.now().astimezone()


def format_date_pt_br(date: datetime) -> str:
    month = MONTHS_PT_BR[date.month - 1]
    return f"{date.day} de {month} de {date.year}"


def load_suep_release_events(path: Path = SUEP_RELEASE_CALENDAR_FILE) -> list[dict[str, str]]:
    if not path.exists():
        return []

    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        print(f"Aviso: não foi possível carregar calendário SUEP. Motivo: {error}")
        return []

    events: list[dict[str, str]] = []
    if not isinstance(payload, list):
        return events

    for item in payload:
        if not isinstance(item, dict):
            continue
        event_date = str(item.get("date", "")).strip()
        label = str(item.get("label", "")).strip()
        if re.fullmatch(r"\d{4}-\d{2}-\d{2}", event_date) and label:
            events.append({"date": event_date, "label": label})
    return events


def release_calendar_events() -> list[dict[str, str]]:
    return [*RELEASE_CALENDAR_EVENTS, *load_suep_release_events()]


def group_release_events_by_date(events: list[dict[str, str]] | None = None) -> dict[str, list[str]]:
    grouped_events: dict[str, list[str]] = {}
    for event in events or release_calendar_events():
        date = event["date"]
        label = event["label"]
        grouped_events.setdefault(date, []).append(label)
    return grouped_events


def iter_calendar_months(start_year: int, start_month: int, end_year: int, end_month: int) -> list[tuple[int, int]]:
    months: list[tuple[int, int]] = []
    year, month = start_year, start_month
    while (year, month) <= (end_year, end_month):
        months.append((year, month))
        month += 1
        if month > 12:
            month = 1
            year += 1
    return months


def build_release_calendar(
    start_year: int = RELEASE_CALENDAR_YEAR,
    start_month: int = 1,
    current_date: datetime | None = None,
) -> dict[str, Any]:
    events = release_calendar_events()
    events_by_date = group_release_events_by_date(events)
    months: list[dict[str, Any]] = []
    start_month = max(1, min(start_month, 12))
    end_year, end_month = RELEASE_CALENDAR_YEAR, 12
    event_dates = [datetime.strptime(event["date"], "%Y-%m-%d") for event in events]
    if event_dates:
        latest_event = max(event_dates)
        end_year, end_month = latest_event.year, latest_event.month

    active_month_index = 0
    for month_index, (year, month) in enumerate(iter_calendar_months(start_year, start_month, end_year, end_month)):
        first_weekday, days_in_month = monthrange(year, month)
        weeks: list[list[dict[str, Any]]] = []
        week: list[dict[str, Any]] = [{"empty": True} for _ in range(first_weekday)]
        is_current_month = bool(current_date and current_date.year == year and current_date.month == month)
        if is_current_month:
            active_month_index = month_index

        for day in range(1, days_in_month + 1):
            iso_date = f"{year}-{month:02d}-{day:02d}"
            events = events_by_date.get(iso_date, [])
            week.append(
                {
                    "empty": False,
                    "day": day,
                    "date": iso_date,
                    "events": events,
                    "event_count": len(events),
                    "tooltip": "\n".join(events),
                }
            )

            if len(week) == 7:
                weeks.append(week)
                week = []

        if week:
            week.extend({"empty": True} for _ in range(7 - len(week)))
            weeks.append(week)

        months.append(
            {
                "name": f"{MONTHS_PT_BR[month - 1].capitalize()} {year}",
                "is_current": is_current_month,
                "weeks": weeks,
            }
        )

    return {
        "title": "Calendário de divulgações",
        "subtitle": "Indicadores e pesquisas",
        "weekdays": CALENDAR_WEEKDAYS,
        "months": months,
        "active_month_index": active_month_index,
    }


def add_current_display_dates(data: dict[str, Any]) -> dict[str, Any]:
    display_data = deepcopy(data)
    now = get_current_datetime()
    display_data["data_edicao"] = format_date_pt_br(now)
    display_data["data_exibicao"] = format_date_pt_br(now)
    display_data["hora_atualizacao"] = now.strftime("%H:%M")
    display_data["ano"] = now.year
    calendar_start_month = (
        max(1, now.month - 1)
        if now.year == RELEASE_CALENDAR_YEAR
        else RELEASE_CALENDAR_DEFAULT_START_MONTH
    )
    display_data["release_calendar"] = build_release_calendar(RELEASE_CALENDAR_YEAR, calendar_start_month, now)
    return display_data


def normalize_indicator_name(name: str) -> str:
    return (
        name.strip()
        .lower()
        .replace("ó", "o")
        .replace("ô", "o")
        .replace("í", "i")
        .replace("é", "e")
        .replace("²", "2")
    )


def format_brazilian_number(value: float, decimals: int = 2) -> str:
    formatted = f"{value:,.{decimals}f}"
    return formatted.replace(",", "X").replace(".", ",").replace("X", ".")


def format_brl(value: Any, decimals: int = 2) -> str:
    try:
        return f"R$ {format_brazilian_number(float(value), decimals)}"
    except (TypeError, ValueError):
        return "N/D"


def format_points(value: Any) -> str:
    try:
        return f"{format_brazilian_number(float(value), 0)} pts"
    except (TypeError, ValueError):
        return "N/D"


def format_percentage(value: Any) -> str:
    try:
        number = float(value)
    except (TypeError, ValueError):
        return "N/D"

    sign = "+" if number > 0 else ""
    return f"{sign}{format_brazilian_number(number, 2)}%"


def format_decimal_percentage(value: Any, signed: bool = False) -> str:
    if value in (None, "", "-"):
        return "N/D"

    try:
        number = float(str(value).replace(",", ".")) * 100
    except (TypeError, ValueError):
        return "N/D"

    sign = "+" if signed and number > 0 else ""
    return f"{sign}{format_brazilian_number(number, 2)}%"


def define_indicator_direction(change_percent: Any) -> str:
    try:
        change = float(change_percent)
    except (TypeError, ValueError):
        return "neutra"

    if change > 0:
        return "alta"
    if change < 0:
        return "queda"
    return "neutra"


def define_indicator_class(change_percent: Any) -> str:
    return f"indicator--{define_indicator_direction(change_percent)}"


def define_direction_from_decimal(value: Any) -> str:
    if value in (None, "", "-"):
        return "neutra"

    try:
        number = float(str(value).replace(",", "."))
    except (TypeError, ValueError):
        return "neutra"

    if number > 0:
        return "alta"
    if number < 0:
        return "queda"
    return "neutra"


def normalize_sheet_text(value: str) -> str:
    normalized = normalize_indicator_name(value)
    normalized = re.sub(r"[^a-z0-9\-/ ]+", " ", normalized)
    normalized = normalized.replace("abrir:", " ")
    return " ".join(normalized.split())


def column_index(cell_reference: str) -> int:
    column = re.sub(r"\d", "", cell_reference)
    index = 0
    for char in column:
        index = index * 26 + (ord(char.upper()) - ord("A") + 1)
    return index - 1


def read_xlsx_shared_strings(zip_file: zipfile.ZipFile) -> list[str]:
    try:
        xml_data = zip_file.read("xl/sharedStrings.xml")
    except KeyError:
        return []

    root = ElementTree.fromstring(xml_data)
    namespace = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    strings: list[str] = []

    for item in root.findall("main:si", namespace):
        parts = [node.text or "" for node in item.findall(".//main:t", namespace)]
        strings.append("".join(parts))

    return strings


def find_xlsx_sheet_path(zip_file: zipfile.ZipFile, sheet_name: str) -> str:
    workbook = ElementTree.fromstring(zip_file.read("xl/workbook.xml"))
    relationships = ElementTree.fromstring(zip_file.read("xl/_rels/workbook.xml.rels"))
    workbook_ns = {
        "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    }
    relationship_ns = {"rel": "http://schemas.openxmlformats.org/package/2006/relationships"}
    targets = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in relationships.findall("rel:Relationship", relationship_ns)
    }

    wanted_name = normalize_sheet_text(sheet_name)
    for sheet in workbook.findall("main:sheets/main:sheet", workbook_ns):
        if normalize_sheet_text(sheet.attrib.get("name", "")) != wanted_name:
            continue

        relationship_id = sheet.attrib.get(f"{{{workbook_ns['rel']}}}id")
        target = targets.get(relationship_id or "")
        if not target:
            break
        return f"xl/{target.lstrip('/')}" if not target.startswith("/") else target.lstrip("/")

    raise ValueError(f"Aba '{sheet_name}' não encontrada.")


def read_xlsx_sheet_rows(path: Path, sheet_name: str) -> list[list[str]]:
    with zipfile.ZipFile(path) as zip_file:
        shared_strings = read_xlsx_shared_strings(zip_file)
        sheet_path = find_xlsx_sheet_path(zip_file, sheet_name)
        root = ElementTree.fromstring(zip_file.read(sheet_path))

    namespace = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    rows: list[list[str]] = []

    for row in root.findall(".//main:sheetData/main:row", namespace):
        values: list[str] = []
        for cell in row.findall("main:c", namespace):
            ref = cell.attrib.get("r", "")
            index = column_index(ref)
            while len(values) <= index:
                values.append("")

            value_node = cell.find("main:v", namespace)
            value = value_node.text if value_node is not None else ""
            if cell.attrib.get("t") == "s" and value:
                value = shared_strings[int(value)]
            elif cell.attrib.get("t") == "inlineStr":
                text_node = cell.find(".//main:t", namespace)
                value = text_node.text if text_node is not None else ""
            values[index] = value or ""

        rows.append(values)

    return rows


def parse_market_datetime(raw_value: Any, fallback: datetime) -> datetime:
    if isinstance(raw_value, (int, float)):
        return datetime.fromtimestamp(raw_value, fallback.tzinfo)

    if isinstance(raw_value, str) and raw_value.strip():
        if raw_value.strip().isdigit():
            return datetime.fromtimestamp(int(raw_value.strip()), fallback.tzinfo)
        clean_value = raw_value.strip().replace("Z", "+00:00")
        try:
            parsed = datetime.fromisoformat(clean_value)
        except ValueError:
            return fallback
        if parsed.tzinfo is None:
            return parsed.replace(tzinfo=fallback.tzinfo)
        return parsed.astimezone(fallback.tzinfo)

    return fallback


def format_update_time(dt: datetime) -> str:
    return f"Atualizado às {dt.strftime('%H:%M')}"


def fallback_market_indicator(name: str) -> dict[str, str]:
    return {
        "nome": name,
        "valor": "N/D",
        "variacao": "N/D",
        "direcao": "neutra",
        "classe": "indicator--neutra",
        "data": "Atualização indisponível",
    }


def fallback_selic_indicator() -> dict[str, str]:
    return {
        "nome": "Selic",
        "valor": "N/D",
        "variacao": "N/D",
        "direcao": "neutra",
        "classe": "indicator--neutra",
        "data": "Atualização indisponível",
    }


def sleep_before_retry(attempt: int, attempts: int, retry_delay: float) -> None:
    if attempt < attempts:
        time.sleep(retry_delay)


def request_urlopen_with_retries(
    request: urllib.request.Request,
    attempts: int = DEFAULT_NETWORK_ATTEMPTS,
    retry_delay: float = NETWORK_RETRY_DELAY_SECONDS,
) -> bytes:
    last_error: Exception | None = None
    for attempt in range(1, attempts + 1):
        try:
            with urllib.request.urlopen(request, timeout=MARKET_TIMEOUT_SECONDS) as response:
                return response.read()
        except Exception as error:
            last_error = error
            sleep_before_retry(attempt, attempts, retry_delay)

    assert last_error is not None
    raise last_error


def load_indicator_cache() -> dict[str, Any]:
    if not INDICATOR_CACHE_FILE.exists():
        return {}

    try:
        payload = json.loads(INDICATOR_CACHE_FILE.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}

    return payload if isinstance(payload, dict) else {}


def save_indicator_cache(cache: dict[str, Any]) -> None:
    INDICATOR_CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
    INDICATOR_CACHE_FILE.write_text(
        json.dumps(cache, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def cache_indicator(indicator: dict[str, str]) -> dict[str, str]:
    name = str(indicator.get("nome", "")).strip()
    value = str(indicator.get("valor", "")).strip()
    if not name or value in {"", "N/D"}:
        return indicator

    cache = load_indicator_cache()
    cache[name] = {
        "indicator": indicator,
        "saved_at": get_current_datetime().isoformat(),
    }
    save_indicator_cache(cache)
    return indicator


def cached_indicator(name: str) -> dict[str, str] | None:
    item = load_indicator_cache().get(name)
    if not isinstance(item, dict):
        return None

    indicator = item.get("indicator")
    if not isinstance(indicator, dict):
        return None

    cached = {str(key): str(value) for key, value in indicator.items()}
    cached["data"] = f"{cached.get('data', '').strip()} · Último valor salvo".strip(" ·")
    return cached


def log_cached_indicator(indicator_name: str) -> None:
    print(f"Aviso: usando último valor salvo para {indicator_name}.")


def request_json(
    url: str,
    query: dict[str, str] | None = None,
    use_brapi_token: bool = False,
    attempts: int = DEFAULT_NETWORK_ATTEMPTS,
) -> dict[str, Any]:
    query = query or {}
    headers = {"User-Agent": "newsletter-economica/1.0"}

    token = get_brapi_token()
    if use_brapi_token and not token:
        raise ValueError("BRAPI_TOKEN não configurado no ambiente ou no arquivo env.")
    if use_brapi_token:
        headers["Authorization"] = f"Bearer {token}"

    query_string = urllib.parse.urlencode(query)
    if query_string:
        url = f"{url}?{query_string}"

    request = urllib.request.Request(url, headers=headers)
    raw_data = request_urlopen_with_retries(request, attempts=attempts)
    return json.loads(raw_data.decode("utf-8"))


def request_text(
    url: str,
    query: dict[str, str] | None = None,
    attempts: int = DEFAULT_NETWORK_ATTEMPTS,
) -> str:
    query = query or {}
    query_string = urllib.parse.urlencode(query)
    if query_string:
        url = f"{url}?{query_string}"

    request = urllib.request.Request(url, headers={"User-Agent": "newsletter-economica/1.0"})
    raw_data = request_urlopen_with_retries(request, attempts=attempts)
    return raw_data.decode("utf-8", errors="replace")


def get_brapi_token() -> str:
    return (
        os.getenv("BRAPI_TOKEN", "").strip()
        or os.getenv("BRAPI_API_TOKEN", "").strip()
    )


def log_fetch_error(indicator_name: str, error: Exception) -> None:
    print(f"Aviso: não foi possível atualizar {indicator_name}. Usando fallback. Motivo: {error}")


def request_brapi_stock_quote(symbol: str) -> dict[str, Any]:
    payload = request_json(
        f"{BRAPI_BASE_URL}/stocks/quote",
        {"symbols": symbol},
        use_brapi_token=True,
    )

    results = payload.get("results")
    if not isinstance(results, list) or not results:
        raise ValueError(f"Resposta sem resultados para {symbol}.")

    item = results[0]
    quote = item.get("data") if isinstance(item, dict) else None
    if isinstance(quote, dict):
        return quote
    if isinstance(item, dict):
        return item
    raise ValueError(f"Resposta inválida para {symbol}.")


def request_brapi_currency(currency: str) -> dict[str, Any]:
    payload = request_json(
        f"{BRAPI_BASE_URL}/currency",
        {"currency": currency},
        use_brapi_token=True,
    )

    currencies = payload.get("currency")
    if not isinstance(currencies, list) or not currencies:
        raise ValueError(f"Resposta sem resultados para {currency}.")
    return currencies[0]


def request_brapi_crypto(coin: str, currency: str = "BRL") -> dict[str, Any]:
    payload = request_json(
        f"{BRAPI_BASE_URL}/crypto",
        {"coin": coin, "currency": currency},
        use_brapi_token=True,
    )

    coins = payload.get("coins")
    if not isinstance(coins, list) or not coins:
        raise ValueError(f"Resposta sem resultados para {coin}-{currency}.")
    return coins[0]


def request_awesomeapi_quote(pair: str) -> dict[str, Any]:
    payload = request_json(f"{AWESOMEAPI_BASE_URL}/{pair}")
    key = pair.replace("-", "")
    quote = payload.get(key)
    if not isinstance(quote, dict):
        raise ValueError(f"Resposta sem resultados para {pair} na AwesomeAPI.")
    return quote


def request_bcb_selic() -> dict[str, Any]:
    payload = request_json(BCB_SELIC_URL, {"formato": "json"}, attempts=SGS_NETWORK_ATTEMPTS)

    if not isinstance(payload, list) or not payload:
        raise ValueError("Resposta sem resultados para Selic SGS 432.")
    return payload[0]


def build_market_indicator(
    name: str,
    quote: dict[str, Any],
    value_formatter: Any,
    date_prefix: str = "Atualizado às",
) -> dict[str, str]:
    now = get_current_datetime()
    price = quote.get("regularMarketPrice")
    change_percent = quote.get("regularMarketChangePercent")
    market_time = parse_market_datetime(quote.get("regularMarketTime"), now)
    direction = define_indicator_direction(change_percent)
    update_text = f"{date_prefix} {market_time.strftime('%H:%M')}"

    return {
        "nome": name,
        "valor": value_formatter(price),
        "variacao": format_percentage(change_percent),
        "direcao": direction,
        "classe": f"indicator--{direction}",
        "data": update_text,
    }


def build_currency_indicator(name: str, quote: dict[str, Any]) -> dict[str, str]:
    now = get_current_datetime()
    price = quote.get("bidPrice")
    change_percent = quote.get("percentageChange")
    market_time = parse_market_datetime(
        quote.get("updatedAtDate") or quote.get("updatedAtTimestamp"),
        now,
    )
    direction = define_indicator_direction(change_percent)

    return {
        "nome": name,
        "valor": format_brl(price, 2),
        "variacao": format_percentage(change_percent),
        "direcao": direction,
        "classe": f"indicator--{direction}",
        "data": format_update_time(market_time),
    }


def build_awesomeapi_indicator(
    name: str,
    quote: dict[str, Any],
    value_formatter: Any,
) -> dict[str, str]:
    now = get_current_datetime()
    price = quote.get("bid")
    change_percent = quote.get("pctChange")
    market_time = parse_market_datetime(quote.get("timestamp"), now)
    direction = define_indicator_direction(change_percent)

    return {
        "nome": name,
        "valor": value_formatter(price),
        "variacao": format_percentage(change_percent),
        "direcao": direction,
        "classe": f"indicator--{direction}",
        "data": format_update_time(market_time),
    }


def buscar_cotacao_dolar() -> dict[str, str]:
    try:
        quote = request_brapi_currency("USD-BRL")
        return cache_indicator(build_currency_indicator("Dólar", quote))
    except Exception as brapi_error:
        try:
            quote = request_awesomeapi_quote("USD-BRL")
            return cache_indicator(build_awesomeapi_indicator("Dólar", quote, lambda value: format_brl(value, 2)))
        except Exception as fallback_error:
            log_fetch_error("Dólar", Exception(f"brapi: {brapi_error}; awesomeapi: {fallback_error}"))
            cached = cached_indicator("Dólar")
            if cached:
                log_cached_indicator("Dólar")
                return cached
        return fallback_market_indicator("Dólar")


def buscar_cotacao_ibovespa() -> dict[str, str]:
    try:
        quote = request_brapi_stock_quote("^BVSP")
        return cache_indicator(build_market_indicator("Ibovespa", quote, format_points))
    except Exception as error:
        log_fetch_error("Ibovespa", error)
        cached = cached_indicator("Ibovespa")
        if cached:
            log_cached_indicator("Ibovespa")
            return cached
        return fallback_market_indicator("Ibovespa")


def buscar_cotacao_bitcoin() -> dict[str, str]:
    try:
        quote = request_brapi_crypto("BTC", "BRL")
        return cache_indicator(build_market_indicator("Bitcoin", quote, lambda value: format_brl(value, 0)))
    except Exception as brapi_error:
        try:
            quote = request_awesomeapi_quote("BTC-BRL")
            return cache_indicator(build_awesomeapi_indicator("Bitcoin", quote, lambda value: format_brl(value, 0)))
        except Exception as fallback_error:
            log_fetch_error("Bitcoin", Exception(f"brapi: {brapi_error}; awesomeapi: {fallback_error}"))
            cached = cached_indicator("Bitcoin")
            if cached:
                log_cached_indicator("Bitcoin")
                return cached
        return fallback_market_indicator("Bitcoin")


def buscar_selic() -> dict[str, str]:
    try:
        result = request_bcb_selic()
        value = float(str(result.get("valor", "")).replace(",", "."))
        return cache_indicator({
            "nome": "Selic",
            "valor": f"{format_brazilian_number(value, 2)}% a.a.",
            "variacao": "Atual",
            "direcao": "neutra",
            "classe": "indicator--neutra",
            "data": "Meta Copom",
        })
    except Exception as error:
        log_fetch_error("Selic", error)
        cached = cached_indicator("Selic")
        if cached:
            log_cached_indicator("Selic")
            return cached
        return fallback_selic_indicator()


def fallback_price_index_indicator(name: str) -> dict[str, str]:
    return {
        "nome": name,
        "valor": "N/D",
        "variacao": "Mensal: N/D",
        "direcao": "neutra",
        "classe": "indicator--neutra",
        "data": "Atualização indisponível",
    }


def build_price_index_indicator(display_name: str, row: dict[str, str]) -> dict[str, str]:
    monthly_value = row.get("variacao_mensal")
    direction = define_direction_from_decimal(monthly_value)

    return {
        "nome": display_name,
        "valor": format_decimal_percentage(row.get("doze_meses")),
        "variacao": f"Mensal: {format_decimal_percentage(monthly_value, signed=True)}",
        "direcao": direction,
        "classe": f"indicator--{direction}",
        "data": f"Ref.: {row.get('data_referencia', 'N/D')}",
    }


def extract_price_index_rows(path: Path = PRICE_INDEX_WORKBOOK) -> dict[str, dict[str, str]]:
    rows = read_xlsx_sheet_rows(path, "VISÃO GERAL")
    if not rows:
        raise ValueError("A aba VISÃO GERAL está vazia.")

    indexes: dict[str, dict[str, str]] = {}
    for row in rows[1:]:
        series = row[0] if len(row) > 0 else ""
        normalized_series = normalize_sheet_text(series)
        if not normalized_series:
            continue

        indexes[normalized_series] = {
            "serie": series,
            "data_referencia": row[2] if len(row) > 2 else "",
            "variacao_mensal": row[3] if len(row) > 3 else "",
            "doze_meses": row[5] if len(row) > 5 else "",
        }

    return indexes


def find_price_index_row(rows: dict[str, dict[str, str]], expected_name: str) -> dict[str, str]:
    normalized_expected = normalize_sheet_text(expected_name)
    for key, value in rows.items():
        if key == normalized_expected or key.endswith(f" {normalized_expected}"):
            return value
    raise ValueError(f"Índice '{expected_name}' não encontrado na aba VISÃO GERAL.")


def buscar_indices_planilha() -> list[dict[str, str]]:
    try:
        rows = extract_price_index_rows()
        return [
            build_price_index_indicator("IPCA", find_price_index_row(rows, "IPCA")),
            build_price_index_indicator("INCC-M", find_price_index_row(rows, "INCC-M")),
            build_price_index_indicator("IGP-M", find_price_index_row(rows, "IGP-M")),
        ]
    except Exception as error:
        log_fetch_error("índices da planilha", error)
        return [
            fallback_price_index_indicator("IPCA"),
            fallback_price_index_indicator("INCC"),
            fallback_price_index_indicator("IGP-M"),
        ]


class BlogIbreParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.posts: list[dict[str, str]] = []
        self._current_href = ""
        self._current_text: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag != "a":
            return

        attributes = dict(attrs)
        href = attributes.get("href") or ""
        if "/posts/" not in href:
            return

        self._current_href = urllib.parse.urljoin(BLOG_IBRE_URL, href)
        self._current_text = []

    def handle_data(self, data: str) -> None:
        if self._current_href:
            self._current_text.append(data)

    def handle_endtag(self, tag: str) -> None:
        if tag != "a" or not self._current_href:
            return

        title = " ".join("".join(self._current_text).split())
        if self._is_valid_title(title):
            self.posts.append({"titulo": title, "url": self._current_href})

        self._current_href = ""
        self._current_text = []

    @staticmethod
    def _is_valid_title(title: str) -> bool:
        if len(title) < 20:
            return False
        if title.lower().startswith("leia mais"):
            return False
        if title.lower().endswith("comentário") or title.lower().endswith("comentários"):
            return False
        return True


def buscar_materias_blog_ibre() -> list[dict[str, str]]:
    html = request_text(
        BLOG_IBRE_URL,
        {
            "utm_source": "portal-ibre",
            "utm_medium": "menu-blog",
            "utm_campaign": "portal-ibre-menu-blog",
        },
    )
    parser = BlogIbreParser()
    parser.feed(html)

    unique_posts: list[dict[str, str]] = []
    seen_urls: set[str] = set()
    for post in parser.posts:
        if post["url"] in seen_urls:
            continue
        unique_posts.append(post)
        seen_urls.add(post["url"])
        if len(unique_posts) == 3:
            break

    if len(unique_posts) < 3:
        raise ValueError("Menos de três matérias encontradas no Blog do IBRE.")
    return unique_posts


def add_blog_ibre_posts(data: dict[str, Any]) -> dict[str, Any]:
    display_data = deepcopy(data)
    try:
        display_data["materias_fgv"] = buscar_materias_blog_ibre()
    except Exception as error:
        log_fetch_error("Blog do IBRE", error)
    return display_data


class LinkCollector(HTMLParser):
    def __init__(self, base_url: str) -> None:
        super().__init__(convert_charrefs=True)
        self.base_url = base_url
        self.links: list[dict[str, str]] = []
        self._href = ""
        self._text: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag != "a":
            return

        attributes = dict(attrs)
        href = attributes.get("href") or ""
        if not href:
            return

        self._href = urllib.parse.urljoin(self.base_url, href)
        self._text = []

    def handle_data(self, data: str) -> None:
        if self._href:
            self._text.append(data)

    def handle_endtag(self, tag: str) -> None:
        if tag != "a" or not self._href:
            return

        text = " ".join("".join(self._text).split())
        if text:
            self.links.append({"texto": text, "url": self._href})

        self._href = ""
        self._text = []


def is_ipea_post_link(link: dict[str, str]) -> bool:
    url = link["url"]
    text = link["texto"]
    lower_text = text.lower()

    if "cartadeconjuntura/index.php/" not in url:
        return False
    if "/category/" in url or "#comment" in url:
        return False
    if len(text) < 18:
        return False
    if re.search(r"\d{1,2} de [a-zç]+ de \d{4}", lower_text):
        return False
    if lower_text in {"carta de conjuntura", "acesse o texto completo", "dados xls"}:
        return False
    if lower_text.startswith("deixe um comentário"):
        return False
    return True


def is_ipea_pdf_link(link: dict[str, str]) -> bool:
    return link["url"].lower().endswith(".pdf") or "acesse o texto completo" in link["texto"].lower()


def find_ipea_post_date(links: list[dict[str, str]], post_url: str, start_index: int) -> str:
    for link in links[start_index + 1 : start_index + 6]:
        if link["url"] != post_url:
            continue
        if re.search(r"\d{1,2} de [a-zç]+ de \d{4}", link["texto"].lower()):
            return link["texto"]
    return "Última publicação"


def clean_html_text(value: str) -> str:
    without_tags = re.sub(r"<[^>]+>", " ", value)
    return " ".join(html_lib.unescape(without_tags).split())


def parse_report_datetime(value: str) -> datetime:
    clean_value = value.strip().lower()
    month_reference = re.search(r"(?:ref\.:\s*)?([a-zç]+)/(\d{4})", clean_value)
    if month_reference:
        month_name, year = month_reference.groups()
        month = {name: index + 1 for index, name in enumerate(MONTHS_PT_BR)}.get(month_name)
        if month:
            return datetime(int(year), month, 28)

    iso_date = re.search(r"(\d{4})-(\d{2})-(\d{2})(?:t(\d{2}):(\d{2}))?", clean_value)
    if iso_date:
        year, month, day, hour, minute = iso_date.groups()
        return datetime(int(year), int(month), int(day), int(hour or 0), int(minute or 0))

    slash_date = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})(?:\s+(\d{1,2})h(\d{2}))?", clean_value)
    if slash_date:
        day, month, year, hour, minute = slash_date.groups()
        return datetime(int(year), int(month), int(day), int(hour or 0), int(minute or 0))

    long_date = re.search(r"(\d{1,2}) de ([a-zç]+) de (\d{4})", clean_value)
    if long_date:
        day, month_name, year = long_date.groups()
        month = {name: index + 1 for index, name in enumerate(MONTHS_PT_BR)}.get(month_name)
        if month:
            return datetime(int(year), month, int(day))

    return datetime.min


def format_report_date(value: str) -> str:
    report_date = parse_report_datetime(value)
    if report_date == datetime.min:
        return value or "Última publicação"
    return format_date_pt_br(report_date)


def select_latest_reports(reports: list[dict[str, str]], limit: int = 5) -> list[dict[str, str]]:
    unique_reports: list[dict[str, str]] = []
    seen_keys: set[str] = set()
    for report in reports:
        file_url = str(report.get("arquivo") or "").strip()
        if file_url and file_url != "#":
            key = file_url
        else:
            key = f"{report.get('categoria', '')}|{report.get('fonte', '')}|{report.get('titulo', '')}|{report.get('data', '')}"
        if key in seen_keys:
            continue
        seen_keys.add(key)
        unique_reports.append(report)
    return sorted(unique_reports, key=lambda report: parse_report_datetime(report.get("data", "")), reverse=True)[:limit]


def buscar_relatorio_ipea(categoria: str, url: str) -> dict[str, str]:
    html = request_text(url)
    parser = LinkCollector(url)
    parser.feed(html)
    links = parser.links

    for index, link in enumerate(links):
        if not is_ipea_post_link(link):
            continue

        for pdf_link in links[index + 1 :]:
            if not is_ipea_pdf_link(pdf_link):
                continue

            return {
                "titulo": link["texto"],
                "categoria": categoria,
                "fonte": "Ipea",
                "descricao": f"Relatório mais recente da Carta de Conjuntura do Ipea em {categoria}.",
                "data": find_ipea_post_date(links, link["url"], index),
                "arquivo": pdf_link["url"],
            }

    raise ValueError(f"Nenhum PDF encontrado para {categoria}.")


def buscar_relatorio_fazenda() -> dict[str, str]:
    html = request_text(FAZENDA_CONJUNTURA_URL)
    pattern = re.compile(
        r"<a[^>]+href=[\"'](?P<href>[^\"']+)[\"'][^>]*>\s*"
        r"(?P<title>[^<]*Panorama Macroecon[^<]*)</a>"
        r"(?P<after>.{0,1200})",
        re.IGNORECASE | re.DOTALL,
    )

    for match in pattern.finditer(html):
        href = urllib.parse.urljoin(FAZENDA_CONJUNTURA_URL, html_lib.unescape(match.group("href")))
        if not href.lower().endswith(".pdf"):
            continue

        title = clean_html_text(match.group("title"))
        after_text = clean_html_text(match.group("after"))
        date_match = re.search(r"\d{2}/\d{2}/\d{4}\s+\d{2}h\d{2}", after_text)
        published_at = date_match.group(0) if date_match else "Última publicação"

        return {
            "titulo": title,
            "categoria": "Ministério da Fazenda",
            "fonte": "Ministério da Fazenda",
            "descricao": "Boletim de Conjuntura mais recente publicado pelo Ministério da Fazenda.",
            "data": format_report_date(published_at),
            "arquivo": href,
        }

    raise ValueError("Nenhum PDF de Panorama Macroeconômico encontrado na página da Fazenda.")


def buscar_anos_bcb_rpm() -> list[int]:
    payload = request_json(f"{BCB_SITE_SERVICE_BASE_URL}/apresrelinf/Anos")
    years: set[int] = set()
    for item in payload.get("conteudo", []):
        reference = str(item.get("dataReferencia") or item.get("DataPublicacao") or "")
        report_date = parse_report_datetime(reference)
        if report_date != datetime.min:
            years.add(report_date.year)
    return sorted(years, reverse=True)


def escolher_ano_bcb_rpm() -> int:
    current_year = get_current_datetime().year
    for year in buscar_anos_bcb_rpm():
        if year <= current_year:
            return year
    return current_year


def buscar_relatorio_bcb_rpm() -> dict[str, str]:
    year = escolher_ano_bcb_rpm()
    payload = request_json(f"{BCB_SITE_SERVICE_BASE_URL}/apresrelinf", {"ano": str(year)})
    reports = payload.get("conteudo", [])
    if not reports and year == get_current_datetime().year:
        payload = request_json(f"{BCB_SITE_SERVICE_BASE_URL}/apresrelinf", {"ano": str(year - 1)})
        reports = payload.get("conteudo", [])

    candidates = [
        item
        for item in reports
        if "apresentação do diretor" in str(item.get("titulo", "")).lower()
        and "relatório de política monetária" in str(item.get("titulo", "")).lower()
        and str(item.get("Url") or (item.get("arquivo") or {}).get("ServerRelativeUrl") or "").lower().endswith(".pdf")
    ]
    if not candidates:
        raise ValueError("Nenhuma apresentação do RPM em PDF encontrada na API do BCB.")

    latest = max(
        candidates,
        key=lambda item: parse_report_datetime(
            str(item.get("dataReferencia") or item.get("DataPublicacao") or item.get("DataModificacao") or "")
        ),
    )
    title = clean_html_text(str(latest.get("titulo", "Apresentação do Relatório de Política Monetária")))
    report_date = str(latest.get("dataReferencia") or latest.get("DataPublicacao") or latest.get("DataModificacao") or "")
    file_url = str(latest.get("Url") or (latest.get("arquivo") or {}).get("ServerRelativeUrl") or "")

    return {
        "titulo": title,
        "categoria": "Banco Central",
        "fonte": "Banco Central",
        "descricao": "Apresentação mais recente do Relatório de Política Monetária publicada pelo Banco Central.",
        "data": format_report_date(report_date),
        "arquivo": urllib.parse.urljoin("https://www.bcb.gov.br", file_url),
    }


def buscar_relatorio_bcb_focus() -> dict[str, str]:
    payload = request_json(f"{BCB_SITE_SERVICE_BASE_URL}/focus/principal", {"filtro": ""})
    reports = payload.get("conteudo", [])
    if not reports:
        raise ValueError("Nenhum Relatório Focus encontrado na API do BCB.")

    latest = reports[0]
    title = clean_html_text(str(latest.get("Titulo") or latest.get("titulo") or "Relatório Focus"))
    report_date = str(latest.get("DataReferencia") or latest.get("data") or latest.get("DataPublicacao") or "")
    file_url = str(latest.get("Url") or "")
    if not file_url.lower().endswith(".pdf"):
        raise ValueError("Relatório Focus mais recente não retornou link de PDF.")

    return {
        "titulo": title,
        "categoria": "Banco Central",
        "fonte": "Banco Central",
        "descricao": "Relatório Focus mais recente publicado pelo Banco Central.",
        "data": format_report_date(report_date),
        "arquivo": urllib.parse.urljoin("https://www.bcb.gov.br", file_url),
    }


def parse_ibge_pdf_reference(file_name: str) -> tuple[datetime, str] | None:
    month_aliases = {
        "jan": "janeiro",
        "fev": "fevereiro",
        "mar": "março",
        "abr": "abril",
        "maio": "maio",
        "mai": "maio",
        "jun": "junho",
        "jul": "julho",
        "ago": "agosto",
        "set": "setembro",
        "out": "outubro",
        "nov": "novembro",
        "dez": "dezembro",
    }
    matches = re.findall(r"_(\d{4})_([a-z]+)", file_name.lower())
    if not matches:
        return None

    year_text, month_key = matches[-1]
    month_name = month_aliases.get(month_key)
    if not month_name:
        return None

    month = MONTHS_PT_BR.index(month_name) + 1
    year = int(year_text)
    return datetime(year, month, 28), f"Ref.: {month_name}/{year}"


def buscar_relatorio_ibge(config: dict[str, str]) -> dict[str, str]:
    url = f"https://biblioteca.ibge.gov.br/index.php/biblioteca-catalogo?view=detalhes&id={config['catalog_id']}"
    html = request_text(url)
    pattern = re.compile(
        r"(?:title=\"(?P<title_a>[^\"]+\.pdf)\"[^>]*href=\"(?P<href_a>[^\"]+\.pdf)\"|"
        r"href=\"(?P<href_b>[^\"]+\.pdf)\"[^>]*title=\"(?P<title_b>[^\"]+\.pdf)\")",
        re.IGNORECASE,
    )

    candidates: list[dict[str, Any]] = []
    for match in pattern.finditer(html):
        file_name = html_lib.unescape(match.group("title_a") or match.group("title_b") or "")
        href = html_lib.unescape(match.group("href_a") or match.group("href_b") or "")
        reference = parse_ibge_pdf_reference(file_name)
        if not reference:
            continue
        reference_date, label = reference
        candidates.append(
            {
                "data_referencia": reference_date,
                "label": label,
                "arquivo": urllib.parse.urljoin(url, href),
            }
        )

    if not candidates:
        raise ValueError(f"Nenhum PDF mensal encontrado no catálogo IBGE {config['catalog_id']}.")

    latest = max(candidates, key=lambda item: item["data_referencia"])
    return {
        "titulo": config["titulo"],
        "categoria": config["categoria"],
        "fonte": config["categoria"],
        "descricao": config["descricao"],
        "data": latest["label"],
        "arquivo": latest["arquivo"],
    }


def buscar_relatorios_ibge() -> list[dict[str, str]]:
    reports: list[dict[str, str]] = []
    for config in IBGE_CATALOG_REPORTS:
        try:
            reports.append(buscar_relatorio_ibge(config))
        except Exception as error:
            log_fetch_error(f"IBGE - {config['titulo']}", error)
            reports.append(fallback_ibge_report(config))
    return reports


def fallback_ipea_report(categoria: str) -> dict[str, str]:
    return {
        "titulo": categoria,
        "categoria": categoria,
        "fonte": "Ipea",
        "descricao": "Relatório será preenchido automaticamente a partir da Carta de Conjuntura do Ipea.",
        "data": "Atualização indisponível",
        "arquivo": "#",
    }


def fallback_fazenda_report() -> dict[str, str]:
    return {
        "titulo": "Panorama Macroeconômico",
        "categoria": "Ministério da Fazenda",
        "fonte": "Ministério da Fazenda",
        "descricao": "Boletim de Conjuntura será preenchido automaticamente a partir do Ministério da Fazenda.",
        "data": "Atualização indisponível",
        "arquivo": "#",
    }


def fallback_bcb_rpm_report() -> dict[str, str]:
    return {
        "titulo": "Apresentação do Relatório de Política Monetária",
        "categoria": "Banco Central",
        "fonte": "Banco Central",
        "descricao": "Apresentação do RPM será preenchida automaticamente a partir do Banco Central.",
        "data": "Atualização indisponível",
        "arquivo": "#",
    }


def fallback_bcb_focus_report() -> dict[str, str]:
    return {
        "titulo": "Relatório Focus",
        "categoria": "Banco Central",
        "fonte": "Banco Central",
        "descricao": "Relatório Focus será preenchido automaticamente a partir do Banco Central.",
        "data": "Atualização indisponível",
        "arquivo": "#",
    }


def fallback_ibge_report(config: dict[str, str]) -> dict[str, str]:
    return {
        "titulo": config["titulo"],
        "categoria": config["categoria"],
        "fonte": config["categoria"],
        "descricao": config["descricao"],
        "data": "Atualização indisponível",
        "arquivo": "#",
    }


def buscar_relatorios_ipea() -> list[dict[str, str]]:
    reports: list[dict[str, str]] = []
    for categoria, url in IPEA_CATEGORY_URLS.items():
        try:
            reports.append(buscar_relatorio_ipea(categoria, url))
        except Exception as error:
            log_fetch_error(f"IPEA - {categoria}", error)
            reports.append(fallback_ipea_report(categoria))
    return reports


def buscar_relatorios_monitorados() -> list[dict[str, str]]:
    reports = buscar_relatorios_ipea()
    try:
        reports.append(buscar_relatorio_fazenda())
    except Exception as error:
        log_fetch_error("Ministério da Fazenda - Boletim de Conjuntura", error)
        reports.append(fallback_fazenda_report())
    try:
        reports.append(buscar_relatorio_bcb_rpm())
    except Exception as error:
        log_fetch_error("Banco Central - Relatório de Política Monetária", error)
        reports.append(fallback_bcb_rpm_report())
    try:
        reports.append(buscar_relatorio_bcb_focus())
    except Exception as error:
        log_fetch_error("Banco Central - Relatório Focus", error)
        reports.append(fallback_bcb_focus_report())
    reports.extend(buscar_relatorios_ibge())
    return reports


def add_ipea_reports(data: dict[str, Any]) -> dict[str, Any]:
    display_data = deepcopy(data)
    reports = buscar_relatorios_monitorados()
    ordered_reports = select_latest_reports(reports, len(reports))
    display_data["relatorios"] = ordered_reports
    display_data["relatorios_preview"] = ordered_reports[:5]
    return display_data


def default_indicator_date(name: str) -> str:
    normalized = normalize_indicator_name(name)
    if normalized == "selic":
        return "Atual"
    if normalized in {"ipca 12m", "incc", "cub/m2"}:
        return "Ref.: jun/26"
    if normalized in {"petroleo brent", "ifix"}:
        return "Fechamento: 06/07"
    return "Referência não informada"


def normalize_static_indicator(item: dict[str, Any]) -> dict[str, str]:
    direction = str(item.get("direcao", "")).strip()
    css_class = str(item.get("classe", "")).strip()

    if not direction and css_class.startswith("indicator--"):
        direction = css_class.replace("indicator--", "", 1)
    if direction not in VALID_DIRECTIONS:
        direction = "neutra"
    if not css_class:
        css_class = f"indicator--{direction}"

    name = str(item.get("nome", "")).strip()
    return {
        "nome": name,
        "valor": str(item.get("valor", "N/D")).strip() or "N/D",
        "variacao": str(item.get("variacao", "N/D")).strip() or "N/D",
        "direcao": direction,
        "classe": css_class,
        "data": str(item.get("data", "")).strip() or default_indicator_date(name),
    }


def build_indicator_list(existing_indicators: list[dict[str, Any]]) -> list[dict[str, str]]:
    ordered_indicators = [
        buscar_cotacao_dolar(),
        buscar_cotacao_ibovespa(),
        buscar_cotacao_bitcoin(),
        buscar_selic(),
    ]
    ordered_indicators.extend(buscar_indices_planilha())
    return ordered_indicators


def add_indicator_data(data: dict[str, Any]) -> dict[str, Any]:
    display_data = deepcopy(data)
    indicators = display_data.get("indicadores", [])
    if isinstance(indicators, list):
        display_data["indicadores"] = build_indicator_list(indicators)
    return display_data


def validate_newsletter(data: dict[str, Any]) -> list[str]:
    errors: list[str] = []

    if not str(data.get("titulo", "")).strip():
        errors.append("O título da newsletter não pode estar vazio.")

    news = data.get("noticias")
    if not isinstance(news, list) or not news:
        errors.append("A newsletter deve ter ao menos 1 notícia.")
    elif isinstance(news, list):
        errors.extend(validate_news_items(news, "Notícia"))

    more_news = data.get("mais_noticias", [])
    if more_news and not isinstance(more_news, list):
        errors.append("O campo mais_noticias deve ser uma lista.")
    elif isinstance(more_news, list):
        errors.extend(validate_news_items(more_news, "Mais notícias"))

    indicators = data.get("indicadores")
    if not isinstance(indicators, list):
        errors.append("O campo indicadores deve ser uma lista.")
    else:
        for index, item in enumerate(indicators, start=1):
            if not str(item.get("nome", "")).strip():
                errors.append(f"Indicador {index}: o nome é obrigatório.")
            if not str(item.get("valor", "")).strip():
                errors.append(f"Indicador {index}: o valor é obrigatório.")

            css_class = str(item.get("classe", "")).strip()
            direction = str(item.get("direcao", "")).strip()
            if not direction and css_class.startswith("indicator--"):
                direction = css_class.replace("indicator--", "", 1)
            if direction not in VALID_DIRECTIONS:
                errors.append(
                    f"Indicador {index}: direção inválida '{direction}'. "
                    "Use 'alta', 'queda' ou 'neutra', ou informe classe como 'indicator--alta', "
                    "'indicator--queda' ou 'indicator--neutra'."
                )

    return errors


def validate_news_items(items: list[dict[str, Any]], label: str) -> list[str]:
    errors: list[str] = []
    required_news_fields = ("titulo", "resumo", "url", "data", "topico")

    for index, item in enumerate(items, start=1):
        for field in required_news_fields:
            if not str(item.get(field, "")).strip():
                errors.append(f"{label} {index}: campo obrigatório ausente ou vazio: {field}.")
        if str(item.get("url", "")).strip() == "#":
            errors.append(f"{label} {index}: a URL não pode ser '#'.")

    return errors


def render_template(data: dict[str, Any], template_name: str) -> str:
    environment = Environment(
        loader=FileSystemLoader(TEMPLATE_DIR),
        autoescape=select_autoescape(("html", "xml")),
    )
    template = environment.get_template(template_name)
    return template.render(newsletter=data)


def write_docs_page(source: Path, destination: Path) -> None:
    html = source.read_text(encoding="utf-8")
    html = html.replace("../static/", "static/")
    html = html.replace("newsletter_preview.html", "index.html")
    destination.write_text(html, encoding="utf-8")


def remove_tree(path: Path) -> None:
    for child in path.rglob("*"):
        try:
            os.chmod(child, 0o700)
        except OSError:
            pass
    try:
        os.chmod(path, 0o700)
    except OSError:
        pass
    shutil.rmtree(path)


def build_github_pages_docs() -> None:
    DOCS_DIR.mkdir(parents=True, exist_ok=True)

    if DOCS_STATIC_DIR.exists():
        remove_tree(DOCS_STATIC_DIR)
    shutil.copytree(BASE_DIR / "static", DOCS_STATIC_DIR)

    write_docs_page(OUTPUT_FILE, DOCS_INDEX_FILE)
    write_docs_page(MORE_NEWS_OUTPUT_FILE, DOCS_MORE_NEWS_FILE)
    write_docs_page(REPORTS_OUTPUT_FILE, DOCS_REPORTS_FILE)
    write_docs_page(DATABASE_OUTPUT_FILE, DOCS_DATABASE_FILE)
    (DOCS_DIR / ".nojekyll").write_text("", encoding="utf-8")


def run_git_command(args: list[str]) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        ["git", *args],
        cwd=BASE_DIR,
        capture_output=True,
        text=True,
        encoding="utf-8",
    )


def publish_github_pages(commit_message: str) -> int:
    repo_check = run_git_command(["rev-parse", "--is-inside-work-tree"])
    if repo_check.returncode != 0:
        print("Publicação automática não executada: esta pasta não parece ser um repositório Git local.")
        print("Suba a pasta docs/ pelo GitHub ou clone o repositório nesta pasta para usar --publish.")
        return 1

    add_result = run_git_command(["add", "docs"])
    if add_result.returncode != 0:
        print("Não foi possível adicionar docs/ ao Git.")
        print(add_result.stderr.strip())
        return add_result.returncode

    diff_result = run_git_command(["diff", "--cached", "--quiet"])
    if diff_result.returncode == 0:
        print("Nenhuma mudança nova em docs/ para publicar.")
        return 0

    commit_result = run_git_command(["commit", "-m", commit_message])
    if commit_result.returncode != 0:
        print("Não foi possível criar o commit.")
        print(commit_result.stderr.strip())
        return commit_result.returncode

    push_result = run_git_command(["push"])
    if push_result.returncode != 0:
        print("Não foi possível enviar para o GitHub.")
        print(push_result.stderr.strip())
        return push_result.returncode

    print("GitHub Pages atualizado via git push.")
    return 0


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Gera a newsletter econômica.")
    parser.add_argument(
        "--publish",
        action="store_true",
        help="Após gerar docs/, cria commit e envia para o GitHub.",
    )
    parser.add_argument(
        "--commit-message",
        default="Atualiza newsletter",
        help="Mensagem do commit usado com --publish.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    load_project_env()
    data = load_newsletter_data(DATA_FILE)
    errors = validate_newsletter(data)

    if errors:
        print("Não foi possível gerar a newsletter. Corrija os erros abaixo:")
        for error in errors:
            print(f"- {error}")
        return 1

    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    display_data = add_ipea_reports(add_blog_ibre_posts(add_indicator_data(add_current_display_dates(data))))
    newsletter_html = render_template(display_data, TEMPLATE_NAME)
    OUTPUT_FILE.write_text(newsletter_html, encoding="utf-8")
    MORE_NEWS_OUTPUT_FILE.write_text(render_template(display_data, MORE_NEWS_TEMPLATE_NAME), encoding="utf-8")
    REPORTS_OUTPUT_FILE.write_text(render_template(display_data, REPORTS_TEMPLATE_NAME), encoding="utf-8")
    DATABASE_OUTPUT_FILE.write_text(render_template(display_data, DATABASE_TEMPLATE_NAME), encoding="utf-8")
    build_github_pages_docs()
    print(f"Newsletter gerada em: {OUTPUT_FILE}")
    print(f"Página de mais notícias gerada em: {MORE_NEWS_OUTPUT_FILE}")
    print(f"Página de relatórios gerada em: {REPORTS_OUTPUT_FILE}")
    print(f"Página de banco de dados gerada em: {DATABASE_OUTPUT_FILE}")
    print(f"Pasta do GitHub Pages atualizada em: {DOCS_DIR}")

    if args.publish:
        return publish_github_pages(args.commit_message)

    return 0


if __name__ == "__main__":
    sys.exit(main())


