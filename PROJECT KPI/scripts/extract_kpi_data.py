#!/usr/bin/env python3
"""
Extrai dados de KPI de um arquivo .xlsx em modo somente leitura (via XML interno)
e grava um JSON para consumo do dashboard web.
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import re
import zipfile
from pathlib import Path
from typing import Any
import xml.etree.ElementTree as ET

NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "docrel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

MONTH_FULL_PT = [
    "JANEIRO",
    "FEVEREIRO",
    "MARÇO",
    "ABRIL",
    "MAIO",
    "JUNHO",
    "JULHO",
    "AGOSTO",
    "SETEMBRO",
    "OUTUBRO",
    "NOVEMBRO",
    "DEZEMBRO",
]

MAX_SALE_BEFORE_LEAD_DAYS = 15

ACUMULADO_ALLOWED_METRICS = {
    "META TOTAL",
    "META RECOR.",
    "META PONT.",
    "META PREVISTA RECORRENTES",
    "META PREVISTA PONTUAIS",
    "VENDAS RECORRENTES",
    "VENDAS PONTUAIS",
    "ORIGINADAS SÓCIOS",
    "VENDAS ORIGINADAS SÓCIOS",
    # compatibilidade com estrutura anterior
    "NOVOS RECOR.",
    "PONTUAIS",
}

ACUMULADO_ALLOWED_COLUMNS = set(MONTH_FULL_PT + ["ACUMULADO"])


def col_to_idx(cell_ref: str) -> int:
    match = re.match(r"([A-Z]+)", cell_ref or "")
    if not match:
        return 0
    value = 0
    for ch in match.group(1):
        value = value * 26 + (ord(ch) - 64)
    return value


def excel_serial_to_iso(serial: float) -> str:
    base = dt.datetime(1899, 12, 30)
    value = base + dt.timedelta(days=serial)
    return value.date().isoformat()


def normalize_text(text: Any) -> str:
    return "" if text is None else str(text).strip()


def parse_num(value: Any) -> float | None:
    raw = normalize_text(value).replace(",", ".")
    if raw == "":
        return None
    try:
        return float(raw)
    except ValueError:
        return None


def parse_maybe_date(header: str, value: Any) -> str:
    raw = normalize_text(value)
    if raw == "":
        return ""

    if "data" in header.lower():
        number = parse_num(raw)
        if number is not None and 35000 <= number <= 60000:
            return excel_serial_to_iso(number)
    return raw


def parse_iso_date(value: Any) -> dt.date | None:
    raw = normalize_text(value)
    if not raw:
        return None
    try:
        return dt.date.fromisoformat(raw[:10])
    except ValueError:
        return None


def month_name_from_iso(value: Any) -> str:
    date_value = parse_iso_date(value)
    if not date_value:
        return ""
    return MONTH_FULL_PT[date_value.month - 1]


def week_label_from_iso(value: Any) -> str:
    date_value = parse_iso_date(value)
    if not date_value:
        return ""
    week = min(5, ((date_value.day - 1) // 7) + 1)
    return f"{week}ª semana"


def normalize_key(value: Any) -> str:
    return " ".join(normalize_text(value).upper().split())


def normalize_origin(value: Any) -> str:
    raw = normalize_key(value)
    if raw in {"S", "SÓCIO", "SOCIO", "SÓCIOS", "SOCIOS"}:
        return "Sócio"
    if raw in {"N", "NÃO SÓCIO", "NAO SOCIO", "NÃO SÓCIOS", "NAO SOCIOS"}:
        return "Não Sócio"
    if "NÃO SÓCIO" in raw or "NAO SOCIO" in raw or "NÃO SÓCIOS" in raw or "NAO SOCIOS" in raw:
        return "Não Sócio"
    if "SÓCIO" in raw or "SOCIO" in raw:
        return "Sócio"
    return "Não informado"


def sale_match_score(lead: dict[str, Any], sale: dict[str, Any]) -> tuple[float, int, str]:
    lead_value = parse_num(lead.get("Valor Proposta")) or 0.0
    sale_value = parse_num(sale.get("Valor")) or 0.0
    base_value = lead_value if lead_value > 0 else sale_value
    value_diff = abs(base_value - sale_value)

    lead_date = parse_iso_date(lead.get("Data"))
    sale_date = parse_iso_date(sale.get("Data"))
    if lead_date and sale_date:
        date_diff_days = abs((sale_date - lead_date).days)
    else:
        date_diff_days = 999999

    sale_date_text = normalize_text(sale.get("Data"))
    return (value_diff, date_diff_days, sale_date_text)


def enrich_geral_with_confirmed_sales(geral: list[dict[str, Any]], vendas: list[dict[str, Any]]) -> None:
    sales_by_exact: dict[tuple[str, str, str], list[dict[str, Any]]] = {}
    sales_by_pair: dict[tuple[str, str], list[dict[str, Any]]] = {}
    for sale in vendas:
        sale_company = normalize_key(sale.get("Cliente") or sale.get("Empresa"))
        sale_service = normalize_key(sale.get("Serviço"))
        sale_unit = normalize_key(sale.get("Unidade"))
        if not sale_company or not sale_service:
            continue
        sales_by_exact.setdefault((sale_company, sale_service, sale_unit), []).append(sale)
        sales_by_pair.setdefault((sale_company, sale_service), []).append(sale)

    used_sales: set[int] = set()

    for lead in geral:
        obs = normalize_text(lead.get("Observação Proposta")).upper()
        auto_marked = "VENDA CONFIRMADA AUTOMATICAMENTE POR CRUZAMENTO" in obs
        if auto_marked:
            lead["Vendido?"] = "N"
            lead["Valor Venda"] = 0
            lead["Data da venda"] = ""
            lead["Status Lead"] = "Proposta" if normalize_key(lead.get("Proposta Enviada ?")) == "S" else "Lead"

    leads_by_exact: dict[tuple[str, str, str], list[int]] = {}
    leads_by_pair: dict[tuple[str, str], list[int]] = {}
    for idx, lead in enumerate(geral):
        lead_company = normalize_key(lead.get("Empresa") or lead.get("Cliente"))
        lead_service = normalize_key(lead.get("Serviço"))
        lead_unit = normalize_key(lead.get("Unidade"))
        if not lead_company or not lead_service:
            continue
        leads_by_exact.setdefault((lead_company, lead_service, lead_unit), []).append(idx)
        leads_by_pair.setdefault((lead_company, lead_service), []).append(idx)

    assigned_leads: set[int] = set()
    sorted_sales = sorted(vendas, key=lambda sale: normalize_text(sale.get("Data")))
    for sale in sorted_sales:
        sale_company = normalize_key(sale.get("Cliente") or sale.get("Empresa"))
        sale_service = normalize_key(sale.get("Serviço"))
        sale_unit = normalize_key(sale.get("Unidade"))
        if not sale_company or not sale_service:
            continue
        exact_candidates = leads_by_exact.get((sale_company, sale_service, sale_unit), [])
        pair_candidates = leads_by_pair.get((sale_company, sale_service), [])
        # Se a venda tem unidade informada, só aceita lead da mesma unidade.
        # Fallback por par (empresa+serviço) só vale quando a venda não tem unidade.
        candidates = exact_candidates if sale_unit else pair_candidates
        if not candidates:
            continue

        sale_date = parse_iso_date(sale.get("Data"))
        best_lead_idx = -1
        best_score: tuple[float, int, str] | None = None
        for lead_idx in candidates:
            if lead_idx in assigned_leads:
                continue
            lead = geral[lead_idx]
            lead_date = parse_iso_date(lead.get("Data"))
            if lead_date and sale_date and (lead_date - sale_date).days > MAX_SALE_BEFORE_LEAD_DAYS:
                continue
            score = sale_match_score(lead, sale)
            if best_score is None or score < best_score:
                best_score = score
                best_lead_idx = lead_idx

        if best_lead_idx < 0:
            continue

        assigned_leads.add(best_lead_idx)
        used_sales.add(id(sale))
        lead = geral[best_lead_idx]
        valor = parse_num(sale.get("Valor")) or 0.0
        lead["Proposta Enviada ?"] = "S"
        lead["Vendido?"] = "S"
        lead["Status Lead"] = "Venda"
        lead["Valor Venda"] = round(valor, 2)
        lead["Data da venda"] = normalize_text(lead.get("Data da venda")) or normalize_text(sale.get("Data"))
        lead["Observação Proposta"] = lead.get("Observação Proposta") or "Venda confirmada automaticamente por cruzamento da base (GERAL x VENDAS) com vínculo 1:1."

    for lead in geral:
        if normalize_text(lead.get("Vendido?")).upper() != "S":
            lead["Vendido?"] = "N"
            lead["Status Lead"] = "Proposta" if normalize_key(lead.get("Proposta Enviada ?")) == "S" else "Lead"


def read_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []

    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    values: list[str] = []

    for si in root.findall("main:si", NS):
        direct = si.find("main:t", NS)
        if direct is not None:
            values.append(direct.text or "")
            continue

        parts: list[str] = []
        for run in si.findall("main:r", NS):
            t = run.find("main:t", NS)
            if t is not None:
                parts.append(t.text or "")
        values.append("".join(parts))

    return values


def read_workbook_map(zf: zipfile.ZipFile) -> dict[str, str]:
    wb = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))

    rel_map: dict[str, str] = {}
    for rel in rels.findall("rel:Relationship", NS):
        rel_map[rel.attrib["Id"]] = rel.attrib["Target"]

    sheet_map: dict[str, str] = {}
    for sheet in wb.findall("main:sheets/main:sheet", NS):
        name = sheet.attrib.get("name", "")
        rid = sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", "")
        target = rel_map.get(rid, "")
        if target:
            normalized_target = target.lstrip("/")
            sheet_map[name] = f"xl/{normalized_target}" if not normalized_target.startswith("xl/") else normalized_target

    return sheet_map


def read_sheet_rows(zf: zipfile.ZipFile, sheet_path: str, shared: list[str]) -> list[list[str]]:
    root = ET.fromstring(zf.read(sheet_path))
    sheet_data = root.find("main:sheetData", NS)
    if sheet_data is None:
        return []

    rows_out: list[list[str]] = []
    for row in sheet_data.findall("main:row", NS):
        cols: dict[int, str] = {}
        for cell in row.findall("main:c", NS):
            idx = col_to_idx(cell.attrib.get("r", ""))
            if idx <= 0:
                continue

            cell_type = cell.attrib.get("t")
            value = ""

            if cell_type == "s":
                v = cell.find("main:v", NS)
                if v is not None and v.text is not None:
                    sidx = int(v.text)
                    value = shared[sidx] if 0 <= sidx < len(shared) else ""
            elif cell_type == "inlineStr":
                t = cell.find("main:is/main:t", NS)
                value = t.text if (t is not None and t.text) else ""
            else:
                v = cell.find("main:v", NS)
                if v is not None and v.text is not None:
                    value = v.text

            cols[idx] = value

        if not cols:
            rows_out.append([])
            continue

        max_col = max(cols.keys())
        rows_out.append([cols.get(i, "") for i in range(1, max_col + 1)])

    return rows_out


def trim_row(row: list[str]) -> list[str]:
    idx = len(row)
    while idx > 0 and normalize_text(row[idx - 1]) == "":
        idx -= 1
    return [normalize_text(x) for x in row[:idx]]


def rows_to_records(rows: list[list[str]], header_cols: int) -> list[dict[str, Any]]:
    cleaned = [trim_row(r) for r in rows if any(normalize_text(c) for c in r)]
    if not cleaned:
        return []

    header = cleaned[0][:header_cols]
    records: list[dict[str, Any]] = []

    for row in cleaned[1:]:
        values = row[:header_cols] + [""] * max(0, header_cols - len(row))
        record: dict[str, Any] = {}
        for i, key in enumerate(header):
            key_text = normalize_text(key)
            if key_text == "":
                continue
            record[key_text] = parse_maybe_date(key_text, values[i])

        if any(normalize_text(v) for v in record.values()):
            records.append(record)

    return records


def detect_header_cols(rows: list[list[str]], minimum: int) -> int:
    if not rows:
        return minimum
    first = trim_row(rows[0])
    return max(minimum, len(first))


def normalize_acumulado_header(value: Any) -> str:
    text = normalize_key(value)
    if text == "SETEMBRO2":
        return "SETEMBRO"
    if text == "NOVEMBRO2":
        return "NOVEMBRO"
    if text == "MARCO":
        return "MARÇO"
    return text


def parse_acumulado_rows(acumulado_rows: list[list[str]]) -> list[dict[str, Any]]:
    cleaned = [trim_row(r) for r in acumulado_rows if any(normalize_text(c) for c in r)]
    if not cleaned:
        return []

    header_row = cleaned[0]
    mapped_headers: dict[int, str] = {}
    for i in range(1, len(header_row)):
        header = normalize_acumulado_header(header_row[i])
        if header in ACUMULADO_ALLOWED_COLUMNS:
            mapped_headers[i] = header

    parsed: list[dict[str, Any]] = []
    for row in cleaned[1:]:
        if not row:
            continue
        metric = normalize_text(row[0])
        metric_key = normalize_key(metric)
        if metric_key not in ACUMULADO_ALLOWED_METRICS:
            continue

        entry: dict[str, Any] = {"Métrica": metric}
        for i, header in mapped_headers.items():
            if i >= len(row):
                continue
            value = normalize_text(row[i])
            if value == "":
                continue
            # Em caso de coluna duplicada (ex.: SETEMBRO e SETEMBRO2), mantém o primeiro valor preenchido.
            if normalize_text(entry.get(header)) != "":
                continue
            entry[header] = value

        if any(normalize_text(v) for k, v in entry.items() if k != "Métrica"):
            parsed.append(entry)

    return parsed


def transform(data_by_sheet: dict[str, list[list[str]]]) -> dict[str, Any]:
    dash_rows = data_by_sheet.get("DASH", [])
    geral_rows = data_by_sheet.get("GERAL", [])
    vendas_rows = data_by_sheet.get("VENDAS", [])
    acumulado_rows = data_by_sheet.get("ACUMULADO", [])

    geral = rows_to_records(geral_rows, header_cols=detect_header_cols(geral_rows, minimum=15))
    vendas = rows_to_records(vendas_rows, header_cols=detect_header_cols(vendas_rows, minimum=11))
    acumulado = parse_acumulado_rows(acumulado_rows)

    for row in geral:
        company = normalize_text(row.get("Empresa")) or normalize_text(row.get("Cliente"))
        if company:
            row["Empresa"] = company
            row["Cliente"] = company
        row["Valor Proposta"] = parse_num(row.get("Valor Proposta"))
        if normalize_text(row.get("Mês")) == "":
            row["Mês"] = month_name_from_iso(row.get("Data"))
        if normalize_text(row.get("Semana Entrada")) == "":
            row["Semana Entrada"] = week_label_from_iso(row.get("Data"))
        if normalize_text(row.get("Semana Envio")) == "":
            envio_ref = row.get("Data de envio") or row.get("Data")
            row["Semana Envio"] = week_label_from_iso(envio_ref)
        proposta_flag = normalize_key(row.get("Proposta Enviada ?"))
        has_send_date = normalize_text(row.get("Data de envio")) != ""
        has_value = (row.get("Valor Proposta") or 0) > 0
        # Higienização: se veio como proposta enviada, mas sem data e sem valor, trata como não enviada.
        if proposta_flag == "S" and (not has_send_date and not has_value):
            row["Proposta Enviada ?"] = "N"

    for row in vendas:
        company = normalize_text(row.get("Cliente")) or normalize_text(row.get("Empresa"))
        if company:
            row["Cliente"] = company
            row["Empresa"] = company
        row["Valor"] = parse_num(row.get("Valor"))
        origin_raw = row.get("Origem")
        if normalize_text(origin_raw) == "":
            origin_raw = row.get("ORIGEM")
        row["Origem"] = normalize_origin(origin_raw)

    enrich_geral_with_confirmed_sales(geral, vendas)

    dash_meta = None
    dash_alcancado = None
    for row in dash_rows:
        trimmed = trim_row(row)
        if not trimmed:
            continue
        for i, cell in enumerate(trimmed):
            label = normalize_text(cell).upper()
            if label == "META" and i + 1 < len(trimmed):
                dash_meta = parse_num(trimmed[i + 1])
            if label == "ALCANÇADO" and i + 1 < len(trimmed):
                dash_alcancado = parse_num(trimmed[i + 1])
    dash_pct = (dash_alcancado / dash_meta) if (dash_meta and dash_alcancado is not None) else None

    payload = {
        "generatedAt": dt.datetime.now().isoformat(timespec="seconds"),
        "sourceSheets": ["DASH", "GERAL", "VENDAS", "ACUMULADO"],
        "records": {
            "dash": {
                "meta": dash_meta,
                "alcancado": dash_alcancado,
                "percentual": dash_pct,
            },
            "geral": geral,
            "vendas": vendas,
            "acumulado": acumulado,
        },
    }
    return payload


def main() -> None:
    parser = argparse.ArgumentParser(description="Extrair KPI do Excel para JSON")
    parser.add_argument("--input", required=True, help="Caminho do arquivo .xlsx")
    parser.add_argument("--output", required=True, help="Caminho do arquivo .json de saída")
    args = parser.parse_args()

    input_path = Path(args.input)
    output_path = Path(args.output)

    if not input_path.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {input_path}")

    with zipfile.ZipFile(input_path, "r") as zf:
        shared = read_shared_strings(zf)
        sheet_map = read_workbook_map(zf)

        needed = ["DASH", "GERAL", "VENDAS", "ACUMULADO"]
        raw: dict[str, list[list[str]]] = {}
        for name in needed:
            sheet_path = sheet_map.get(name)
            if not sheet_path:
                raw[name] = []
                continue
            raw[name] = read_sheet_rows(zf, sheet_path, shared)

    payload = transform(raw)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"OK: JSON gerado em {output_path}")


if __name__ == "__main__":
    main()
