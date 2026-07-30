from __future__ import annotations

import hashlib
import io
import re
import unicodedata
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from difflib import SequenceMatcher
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import pandas as pd
from openpyxl import load_workbook

DEFAULT_EXCLUSIONS = ["LAVADO", "NEUMATICO", "AUTOMATIZACION", "CAMARA"]


def normalize_text(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).strip().upper()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(char for char in text if not unicodedata.combining(char))
    text = re.sub(r"\s+", " ", text)
    return text


def safe_text(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip()


def _find_column(columns: Sequence[Any], aliases: Sequence[str]) -> Optional[Any]:
    normalized = {normalize_text(col): col for col in columns}
    for alias in aliases:
        if normalize_text(alias) in normalized:
            return normalized[normalize_text(alias)]
    for col in columns:
        ncol = normalize_text(col)
        if any(normalize_text(alias) in ncol for alias in aliases):
            return col
    return None


def _parse_date(value: Any) -> Optional[date]:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    parsed = pd.to_datetime(value, dayfirst=True, errors="coerce")
    return None if pd.isna(parsed) else parsed.date()


def _parse_time(value: Any) -> time:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return time(0, 0)
    if isinstance(value, datetime):
        return value.time().replace(microsecond=0)
    if isinstance(value, time):
        return value.replace(microsecond=0)
    if isinstance(value, (float, int)) and 0 <= float(value) < 1:
        seconds = int(round(float(value) * 86400)) % 86400
        return (datetime.min + timedelta(seconds=seconds)).time()
    parsed = pd.to_datetime(str(value), errors="coerce")
    return time(0, 0) if pd.isna(parsed) else parsed.time().replace(microsecond=0)


def _identifier(*parts: Any) -> str:
    raw = "|".join(normalize_text(part) for part in parts)
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def read_detentions_excel(file_bytes: bytes, filename: str, exclusions: Sequence[str]) -> pd.DataFrame:
    excel = pd.ExcelFile(io.BytesIO(file_bytes))
    sheet = "Limpio" if "Limpio" in excel.sheet_names else excel.sheet_names[0]
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet)
    df.columns = [safe_text(col) for col in df.columns]

    col_equipo = _find_column(df.columns, ["Equipo"])
    col_fecha = _find_column(df.columns, ["Fecha"])
    col_hora = _find_column(df.columns, ["Hora", "Hora inicio"])
    col_razon = _find_column(df.columns, ["Razón", "Razon", "Descripción", "Descripcion"])
    col_comentario = _find_column(df.columns, ["Comentarios", "Comentario"])
    col_categoria = _find_column(df.columns, ["Categoría", "Categoria"])
    col_tipo_categoria = _find_column(df.columns, ["Tipo Categoría", "Tipo Categoria"])
    col_turno = _find_column(df.columns, ["Turno"])
    col_codigo = _find_column(df.columns, ["Código", "Codigo"])

    missing = []
    if col_equipo is None:
        missing.append("Equipo")
    if col_fecha is None:
        missing.append("Fecha")
    if col_razon is None:
        missing.append("Razón/Descripción")
    if missing:
        raise ValueError("No se encontraron columnas obligatorias: " + ", ".join(missing))

    exclusion_tokens = [normalize_text(item) for item in exclusions if safe_text(item)]
    records: List[Dict[str, Any]] = []

    for _, row in df.iterrows():
        equipo = safe_text(row.get(col_equipo))
        fecha = _parse_date(row.get(col_fecha))
        if not equipo or fecha is None:
            continue
        hora = _parse_time(row.get(col_hora)) if col_hora else time(0, 0)
        fecha_hora = datetime.combine(fecha, hora)
        razon = safe_text(row.get(col_razon))
        comentario = safe_text(row.get(col_comentario)) if col_comentario else ""
        categoria = safe_text(row.get(col_categoria)) if col_categoria else ""
        tipo_categoria = safe_text(row.get(col_tipo_categoria)) if col_tipo_categoria else ""
        turno = safe_text(row.get(col_turno)) if col_turno else ""
        codigo = safe_text(row.get(col_codigo)) if col_codigo else ""
        searchable = normalize_text(" ".join([razon, comentario, categoria, tipo_categoria]))
        exclusion = next((token for token in exclusion_tokens if token and token in searchable), "")
        requiere_ot = not bool(exclusion)

        records.append({
            "identificador": _identifier(equipo, fecha_hora.isoformat(), razon, comentario),
            "equipo": equipo,
            "fecha_detencion": fecha.isoformat(),
            "hora_inicio": hora.strftime("%H:%M:%S"),
            "fecha_hora_inicio": fecha_hora.isoformat(),
            "turno": turno,
            "codigo": codigo,
            "razon": razon,
            "comentario": comentario,
            "categoria": categoria,
            "tipo_categoria": tipo_categoria,
            "descripcion_normalizada": searchable,
            "requiere_ot": requiere_ot,
            "motivo_exclusion": exclusion.title() if exclusion else "",
            "archivo_origen": filename,
            "estado": "PENDIENTE" if requiere_ot else "EXCLUIDA",
        })

    return pd.DataFrame(records)


def work_order_records(validation_results: Sequence[Dict[str, Any]], received_at: Optional[datetime] = None) -> List[Dict[str, Any]]:
    received_at = received_at or datetime.now()
    records: List[Dict[str, Any]] = []
    for result in validation_results:
        ot_number = safe_text(result.get("Orden"))
        if not ot_number:
            ot_number = f"SIN-NUMERO-{_identifier(result.get('Archivo'), result.get('Equipo'))[:12]}"
        description = safe_text(result.get("Descripción OT")) or safe_text(result.get("Motivo detención"))
        records.append({
            "numero_ot": ot_number,
            "equipo": safe_text(result.get("Equipo")),
            "turno": safe_text(result.get("Turno")),
            "descripcion": description,
            "descripcion_normalizada": normalize_text(description),
            "archivo_origen": safe_text(result.get("Archivo")),
            "fecha_recepcion": received_at.date().isoformat(),
            "fecha_hora_recepcion": received_at.isoformat(),
            "estado_validacion": safe_text(result.get("Estado")),
            "campos_faltantes": int(result.get("Campos faltantes", 0) or 0),
        })
    return records


def similarity(detention: Dict[str, Any], ot: Dict[str, Any]) -> float:
    if normalize_text(detention.get("equipo")) != normalize_text(ot.get("equipo")):
        return 0.0
    a = normalize_text(detention.get("descripcion_normalizada") or detention.get("razon"))
    b = normalize_text(ot.get("descripcion_normalizada") or ot.get("descripcion"))
    text_score = SequenceMatcher(None, a, b).ratio() if a and b else 0.0
    return round(0.35 + 0.65 * text_score, 4)


def week_bounds(reference: date) -> Tuple[date, date]:
    start = reference - timedelta(days=reference.weekday())
    return start, start + timedelta(days=6)


def daily_summary(detentions: pd.DataFrame, associations: pd.DataFrame) -> pd.DataFrame:
    columns = ["Fecha", "Turno", "Detenciones válidas", "Detenciones con OT", "Pendientes", "Cumplimiento (%)"]
    if detentions.empty:
        return pd.DataFrame(columns=columns)

    valid = detentions[detentions["requiere_ot"].fillna(False).astype(bool)].copy()
    if valid.empty:
        return pd.DataFrame(columns=columns)

    associated_ids = set()
    if not associations.empty and "detencion_id" in associations.columns:
        associated_ids = set(associations["detencion_id"].dropna().astype(str))
    valid["con_ot"] = valid["id"].astype(str).isin(associated_ids)
    valid["fecha"] = pd.to_datetime(valid["fecha_detencion"], errors="coerce").dt.date
    valid["turno"] = valid.get("turno", "").fillna("").replace("", "—")

    grouped = valid.groupby(["fecha", "turno"], dropna=False).agg(
        detenciones_validas=("id", "count"),
        con_ot=("con_ot", "sum"),
    ).reset_index()
    grouped["pendientes"] = grouped["detenciones_validas"] - grouped["con_ot"]
    grouped["cumplimiento"] = (grouped["con_ot"] / grouped["detenciones_validas"] * 100).round(1)
    grouped = grouped.sort_values(["fecha", "turno"])
    return pd.DataFrame({
        "Fecha": grouped["fecha"].apply(lambda x: x.strftime("%d/%m/%Y") if pd.notna(x) else ""),
        "Turno": grouped["turno"],
        "Detenciones válidas": grouped["detenciones_validas"].astype(int),
        "Detenciones con OT": grouped["con_ot"].astype(int),
        "Pendientes": grouped["pendientes"].astype(int),
        "Cumplimiento (%)": grouped["cumplimiento"],
    })


@dataclass
class SupabaseRepository:
    client: Any

    def upsert_detentions(self, records: Sequence[Dict[str, Any]]) -> int:
        if not records:
            return 0
        self.client.table("detenciones").upsert(list(records), on_conflict="identificador").execute()
        return len(records)

    def upsert_work_orders(self, records: Sequence[Dict[str, Any]]) -> int:
        if not records:
            return 0
        self.client.table("ordenes_trabajo").upsert(list(records), on_conflict="numero_ot").execute()
        return len(records)

    def list_detentions(self, start: date, end: date) -> pd.DataFrame:
        response = self.client.table("detenciones").select("*").gte(
            "fecha_detencion", start.isoformat()
        ).lte("fecha_detencion", end.isoformat()).order("fecha_hora_inicio").execute()
        return pd.DataFrame(response.data or [])

    def list_work_orders(self) -> pd.DataFrame:
        response = self.client.table("ordenes_trabajo").select("*").order(
            "fecha_hora_recepcion", desc=True
        ).execute()
        return pd.DataFrame(response.data or [])

    def list_associations(self, detention_ids: Optional[Sequence[Any]] = None) -> pd.DataFrame:
        query = self.client.table("detencion_ot").select("*")
        if detention_ids:
            query = query.in_("detencion_id", list(detention_ids))
        response = query.execute()
        return pd.DataFrame(response.data or [])

    def associate(self, detention_id: Any, work_order_id: Any, confidence: float, association_type: str) -> None:
        record = {
            "detencion_id": detention_id,
            "ot_id": work_order_id,
            "confianza": float(confidence),
            "tipo_asociacion": association_type,
            "confirmada": True,
        }
        self.client.table("detencion_ot").upsert(
            record, on_conflict="detencion_id,ot_id"
        ).execute()
        self.client.table("detenciones").update({"estado": "CON_OT"}).eq("id", detention_id).execute()

    def auto_associate(self, start: date, end: date, threshold: float = 0.72) -> int:
        detentions = self.list_detentions(start, end)
        if detentions.empty:
            return 0
        associations = self.list_associations(detentions["id"].tolist())
        associated_ids = set(associations.get("detencion_id", pd.Series(dtype=object)).astype(str)) if not associations.empty else set()
        pending = detentions[
            detentions["requiere_ot"].fillna(False).astype(bool)
            & ~detentions["id"].astype(str).isin(associated_ids)
        ]
        ots = self.list_work_orders()
        if pending.empty or ots.empty:
            return 0

        used_pairs = set()
        count = 0
        for _, ot in ots.iterrows():
            candidates = pending[
                pending["equipo"].apply(normalize_text) == normalize_text(ot.get("equipo"))
            ]
            if candidates.empty:
                continue
            scored = [(similarity(det.to_dict(), ot.to_dict()), det) for _, det in candidates.iterrows()]
            scored.sort(key=lambda item: item[0], reverse=True)
            best_score, best_det = scored[0]
            unique_equipment_candidate = len(candidates) == 1
            if best_score >= threshold or unique_equipment_candidate:
                pair = (str(best_det["id"]), str(ot["id"]))
                if pair not in used_pairs:
                    self.associate(best_det["id"], ot["id"], best_score, "AUTOMATICA")
                    used_pairs.add(pair)
                    count += 1
        return count
