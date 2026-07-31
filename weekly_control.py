from __future__ import annotations

import hashlib
import io
import re
import unicodedata
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from difflib import SequenceMatcher
from typing import Any, Dict, List, Optional, Sequence, Tuple

import pandas as pd

DEFAULT_EXCLUSIONS = ["LAVADO", "NEUMATICO", "AUTOMATIZACION", "CAMARA"]

# La jornada operacional de Dispatch comienza a las 08:00 y termina a las
# 07:59:59 del día calendario siguiente.
OPERATIONAL_DAY_START = time(8, 0)
TURN_B_START = time(20, 0)


def operational_date_and_shift(value: datetime) -> Tuple[date, str]:
    """Devuelve la fecha de control y el turno operacional A/B.

    Reglas:
    - 08:00:00 a 19:59:59 -> fecha calendario, turno A.
    - 20:00:00 a 23:59:59 -> fecha calendario, turno B.
    - 00:00:00 a 07:59:59 -> día operacional anterior, turno B.
    """
    current_time = value.time().replace(microsecond=0)
    if current_time < OPERATIONAL_DAY_START:
        return value.date() - timedelta(days=1), "B"
    if current_time < TURN_B_START:
        return value.date(), "A"
    return value.date(), "B"


def _record_operational_date(record: Dict[str, Any]) -> Optional[date]:
    stored = _parse_date(record.get("fecha_operacional"))
    if stored is not None:
        return stored

    dt = _parse_iso_datetime(record.get("fecha_hora_inicio"))
    if dt is None:
        dt = _combine_datetime(
            record.get("fecha_detencion") or record.get("fecha_inicio"),
            record.get("hora_inicio"),
        )
    if dt is None:
        return None
    return operational_date_and_shift(dt)[0]

# Palabras poco discriminantes para comparar descripciones técnicas.
_TEXT_STOPWORDS = {
    "DE", "DEL", "LA", "LAS", "EL", "LOS", "Y", "EN", "A", "AL", "POR",
    "PARA", "CON", "SE", "REALIZA", "REALIZAR", "TRABAJO", "EQUIPO",
}


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


def normalize_equipment(value: Any) -> str:
    """
    Convierte variantes del código de equipo a una clave comparable.

    Ejemplos:
    - TO-28_D10T2 -> TO28
    - TO-28       -> TO28
    - MO-10_2     -> MO10
    - TN 17       -> TN17
    """
    text = normalize_text(value).replace(" ", "")
    if not text:
        return ""

    # El formato OT puede incorporar turno o sufijos después de _ o /.
    text = re.split(r"[_/\\]", text, maxsplit=1)[0]
    text = re.sub(r"[^A-Z0-9]", "", text)

    match = re.match(r"^([A-Z]+)0*(\d+)", text)
    if not match:
        return text

    prefix, numeric = match.groups()
    number = str(int(numeric))
    if len(number) < 2:
        number = number.zfill(2)
    return f"{prefix}{number}"


def _find_column(columns: Sequence[Any], aliases: Sequence[str]) -> Optional[Any]:
    normalized = {normalize_text(col): col for col in columns}
    for alias in aliases:
        match = normalized.get(normalize_text(alias))
        if match is not None:
            return match

    # Solo se usa búsqueda parcial para alias de más de un carácter.
    long_aliases = [normalize_text(alias) for alias in aliases if len(normalize_text(alias)) > 1]
    for col in columns:
        ncol = normalize_text(col)
        if any(alias in ncol for alias in long_aliases):
            return col
    return None


def _parse_date(value: Any) -> Optional[date]:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = safe_text(value)
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", text):
        try:
            return datetime.strptime(text, "%Y-%m-%d").date()
        except ValueError:
            return None
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

    text = safe_text(value)
    for fmt in ("%I:%M %p", "%I:%M:%S %p", "%H:%M", "%H:%M:%S"):
        try:
            return datetime.strptime(text.upper(), fmt).time()
        except ValueError:
            pass

    parsed = pd.to_datetime(text, errors="coerce")
    return time(0, 0) if pd.isna(parsed) else parsed.time().replace(microsecond=0)


def _parse_duration_seconds(value: Any) -> int:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return 0
    if isinstance(value, timedelta):
        return max(0, int(round(value.total_seconds())))
    if isinstance(value, time):
        return value.hour * 3600 + value.minute * 60 + value.second
    if isinstance(value, (float, int)):
        # Excel almacena las duraciones como fracción de día; 1 equivale a 24 h.
        return max(0, int(round(float(value) * 86400)))

    text = safe_text(value)
    match = re.fullmatch(r"(\d+):(\d{1,2}):(\d{1,2})", text)
    if match:
        hours, minutes, seconds = map(int, match.groups())
        return max(0, hours * 3600 + minutes * 60 + seconds)

    parsed = pd.to_timedelta(text, errors="coerce")
    return 0 if pd.isna(parsed) else max(0, int(round(parsed.total_seconds())))


def _combine_datetime(date_value: Any, time_value: Any) -> Optional[datetime]:
    parsed_date = _parse_date(date_value)
    if parsed_date is None:
        return None
    return datetime.combine(parsed_date, _parse_time(time_value))


def _parse_iso_datetime(value: Any) -> Optional[datetime]:
    if value is None or safe_text(value) == "":
        return None
    if isinstance(value, datetime):
        return value.replace(tzinfo=None)
    parsed = pd.to_datetime(value, errors="coerce")
    if pd.isna(parsed):
        return None
    result = parsed.to_pydatetime()
    if result.tzinfo is not None:
        result = result.replace(tzinfo=None)
    return result


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
    col_duracion = _find_column(df.columns, ["Duración", "Duracion"])

    # La columna H del reporte representa la hora de término. Se busca de forma exacta
    # para no confundirla con columnas que simplemente contienen la letra H.
    normalized_columns = {normalize_text(col): col for col in df.columns}
    col_hora_termino = normalized_columns.get("H")
    if col_hora_termino is None:
        col_hora_termino = _find_column(df.columns, ["Hora término", "Hora termino", "Hora fin"])

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
        equipo_normalizado = normalize_equipment(equipo)
        fecha = _parse_date(row.get(col_fecha))
        if not equipo or not equipo_normalizado or fecha is None:
            continue

        hora = _parse_time(row.get(col_hora)) if col_hora else time(0, 0)
        fecha_hora = datetime.combine(fecha, hora)
        fecha_operacional, turno_operacional = operational_date_and_shift(fecha_hora)
        duration_seconds = _parse_duration_seconds(row.get(col_duracion)) if col_duracion else 0

        fecha_hora_termino: Optional[datetime] = None
        if duration_seconds > 0:
            fecha_hora_termino = fecha_hora + timedelta(seconds=duration_seconds)
        elif col_hora_termino:
            hora_termino = _parse_time(row.get(col_hora_termino))
            fecha_hora_termino = datetime.combine(fecha, hora_termino)
            if fecha_hora_termino <= fecha_hora:
                fecha_hora_termino += timedelta(days=1)

        razon = safe_text(row.get(col_razon))
        comentario = safe_text(row.get(col_comentario)) if col_comentario else ""
        categoria = safe_text(row.get(col_categoria)) if col_categoria else ""
        tipo_categoria = safe_text(row.get(col_tipo_categoria)) if col_tipo_categoria else ""
        # Para el control semanal se usa el turno operacional calculado a partir
        # de la hora de inicio. Esto corrige los registros entre 00:00 y 07:59,
        # que Dispatch fecha con el día calendario siguiente pero pertenecen al
        # turno B del día operacional anterior.
        turno_fuente = safe_text(row.get(col_turno)) if col_turno else ""
        turno = turno_operacional or turno_fuente
        codigo = safe_text(row.get(col_codigo)) if col_codigo else ""
        searchable = normalize_text(" ".join([razon, comentario, categoria, tipo_categoria]))
        exclusion = next((token for token in exclusion_tokens if token and token in searchable), "")
        requiere_ot = not bool(exclusion)

        records.append({
            "identificador": _identifier(equipo_normalizado, fecha_hora.isoformat(), razon, comentario),
            "equipo": equipo,
            "equipo_normalizado": equipo_normalizado,
            "fecha_detencion": fecha.isoformat(),
            "fecha_operacional": fecha_operacional.isoformat(),
            "hora_inicio": hora.strftime("%H:%M:%S"),
            "fecha_hora_inicio": fecha_hora.isoformat(),
            "hora_termino": fecha_hora_termino.time().strftime("%H:%M:%S") if fecha_hora_termino else None,
            "fecha_hora_termino": fecha_hora_termino.isoformat() if fecha_hora_termino else None,
            "duracion_segundos": duration_seconds,
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


def work_order_records(
    validation_results: Sequence[Dict[str, Any]],
    received_at: Optional[datetime] = None,
) -> List[Dict[str, Any]]:
    received_at = received_at or datetime.now()
    records: List[Dict[str, Any]] = []

    for result in validation_results:
        ot_number = safe_text(result.get("Orden"))
        if not ot_number:
            ot_number = f"SIN-NUMERO-{_identifier(result.get('Archivo'), result.get('Equipo'))[:12]}"

        equipo = safe_text(result.get("Equipo"))
        description = safe_text(result.get("Descripción OT")) or safe_text(result.get("Motivo detención"))

        start_dt = _combine_datetime(
            result.get("Fecha inicio OT"),
            result.get("Hora inicio OT"),
        )
        end_dt = _combine_datetime(
            result.get("Fecha término OT"),
            result.get("Hora término OT"),
        )
        if start_dt and end_dt and end_dt < start_dt:
            end_dt += timedelta(days=1)

        fecha_operacional_ot: Optional[date] = None
        turno_operacional_ot = safe_text(result.get("Turno"))
        if start_dt is not None:
            fecha_operacional_ot, turno_calculado = operational_date_and_shift(start_dt)
            if not turno_operacional_ot:
                turno_operacional_ot = turno_calculado

        records.append({
            "numero_ot": ot_number,
            "equipo": equipo,
            "equipo_normalizado": normalize_equipment(equipo),
            "turno": turno_operacional_ot,
            "descripcion": description,
            "descripcion_normalizada": normalize_text(description),
            "fecha_inicio": start_dt.date().isoformat() if start_dt else None,
            "fecha_operacional": fecha_operacional_ot.isoformat() if fecha_operacional_ot else None,
            "hora_inicio": start_dt.time().strftime("%H:%M:%S") if start_dt else None,
            "fecha_hora_inicio": start_dt.isoformat() if start_dt else None,
            "fecha_termino": end_dt.date().isoformat() if end_dt else None,
            "hora_termino": end_dt.time().strftime("%H:%M:%S") if end_dt else None,
            "fecha_hora_termino": end_dt.isoformat() if end_dt else None,
            "archivo_origen": safe_text(result.get("Archivo")),
            "fecha_recepcion": received_at.date().isoformat(),
            "fecha_hora_recepcion": received_at.isoformat(),
            "estado_validacion": safe_text(result.get("Estado")),
            "campos_faltantes": int(result.get("Campos faltantes", 0) or 0),
        })

    return records


def _tokenize_description(value: Any) -> set[str]:
    text = normalize_text(value)
    tokens: set[str] = set()
    for token in re.findall(r"[A-Z0-9]+", text):
        if token in _TEXT_STOPWORDS or len(token) <= 1:
            continue
        # Normalización básica de plural para FILTROS/FILTRO, MANGUERAS/MANGUERA, etc.
        if token.endswith("ES") and len(token) > 5:
            token = token[:-2]
        elif token.endswith("S") and len(token) > 4:
            token = token[:-1]
        tokens.add(token)
    return tokens


def _text_similarity(a: Any, b: Any) -> float:
    text_a = normalize_text(a)
    text_b = normalize_text(b)
    if not text_a or not text_b:
        return 0.0

    sequence = SequenceMatcher(None, text_a, text_b).ratio()
    tokens_a = _tokenize_description(text_a)
    tokens_b = _tokenize_description(text_b)
    if not tokens_a or not tokens_b:
        return sequence

    intersection = len(tokens_a & tokens_b)
    union = len(tokens_a | tokens_b)
    jaccard = intersection / union if union else 0.0
    containment = intersection / min(len(tokens_a), len(tokens_b))
    return max(sequence, jaccard, containment)


def _temporal_similarity(detention: Dict[str, Any], ot: Dict[str, Any]) -> Optional[float]:
    det_start = _parse_iso_datetime(detention.get("fecha_hora_inicio"))
    det_end = _parse_iso_datetime(detention.get("fecha_hora_termino"))
    ot_start = _parse_iso_datetime(ot.get("fecha_hora_inicio"))
    ot_end = _parse_iso_datetime(ot.get("fecha_hora_termino"))

    if det_start is None or ot_start is None:
        return None

    if det_end is None:
        det_end = det_start
    if ot_end is None:
        ot_end = ot_start

    # Corrige cruces de medianoche si una fuente solo entregó horas.
    if det_end < det_start:
        det_end += timedelta(days=1)
    if ot_end < ot_start:
        ot_end += timedelta(days=1)

    # Superposición real de intervalos: señal temporal más fuerte.
    if max(det_start, ot_start) <= min(det_end, ot_end):
        return 1.0

    if ot_start > det_end:
        gap_hours = (ot_start - det_end).total_seconds() / 3600
    else:
        gap_hours = (det_start - ot_end).total_seconds() / 3600

    gap_hours = abs(gap_hours)
    if gap_hours <= 1:
        return 0.90
    if gap_hours <= 3:
        return 0.75
    if gap_hours <= 6:
        return 0.55
    if gap_hours <= 12:
        return 0.35
    if det_start.date() == ot_start.date():
        return 0.25
    if abs((det_start.date() - ot_start.date()).days) <= 1:
        return 0.10
    return 0.0


def similarity(detention: Dict[str, Any], ot: Dict[str, Any]) -> float:
    det_equipment = normalize_equipment(
        detention.get("equipo_normalizado") or detention.get("equipo")
    )
    ot_equipment = normalize_equipment(
        ot.get("equipo_normalizado") or ot.get("equipo")
    )
    if not det_equipment or det_equipment != ot_equipment:
        return 0.0

    det_description = detention.get("descripcion_normalizada") or " ".join(
        [safe_text(detention.get("razon")), safe_text(detention.get("comentario"))]
    )
    ot_description = ot.get("descripcion_normalizada") or ot.get("descripcion")
    text_score = _text_similarity(det_description, ot_description)
    temporal_score = _temporal_similarity(detention, ot)

    det_operational_date = _record_operational_date(detention)
    ot_operational_date = _record_operational_date(ot)
    same_operational_day = (
        det_operational_date is not None
        and ot_operational_date is not None
        and det_operational_date == ot_operational_date
    )

    if temporal_score is None:
        # Respaldo para OT antiguas que aún no tienen fecha/hora guardada.
        # Equipo exacto + mismo día operacional + descripción técnica.
        base = 0.45 if same_operational_day else 0.25
        score = base + 0.55 * text_score
    elif temporal_score >= 0.99:
        # Equipo exacto y superposición real de intervalos. La coincidencia
        # temporal es suficiente para que el caso no se pierda por diferencias
        # de redacción (por ejemplo SATURADO vs TAPONADO).
        score = 0.70 + 0.30 * text_score
    elif same_operational_day:
        score = 0.52 * temporal_score + 0.48 * text_score + 0.08
    else:
        score = 0.52 * temporal_score + 0.48 * text_score

    return round(min(1.0, score), 4)


def week_bounds(reference: date) -> Tuple[date, date]:
    start = reference - timedelta(days=reference.weekday())
    return start, start + timedelta(days=6)


def _format_turn_value(value: Any) -> str:
    """Normaliza el turno leído desde la OT, por ejemplo 3.0 -> 3."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return "—"
    if isinstance(value, (int, float)):
        numeric = float(value)
        if numeric.is_integer():
            return str(int(numeric))
    text = safe_text(value)
    if not text:
        return "—"
    return re.sub(r"\.0$", "", text)


def candidate_work_order_turn(
    detention: Dict[str, Any],
    work_orders: pd.DataFrame,
    minimum_score: float = 0.30,
) -> str:
    """Obtiene el turno desde la OT candidata más probable para una detención.

    El turno nunca se toma del reporte de detenciones. Se utiliza exclusivamente
    el valor guardado desde la celda G19 de la OT. Si no existe una OT candidata
    razonable para el mismo equipo y período operacional, se muestra un guion.
    """
    if work_orders is None or work_orders.empty:
        return "—"

    detention_equipment = normalize_equipment(
        detention.get("equipo_normalizado") or detention.get("equipo")
    )
    if not detention_equipment:
        return "—"

    candidates = work_orders[
        work_orders.apply(
            lambda row: normalize_equipment(
                row.get("equipo_normalizado") or row.get("equipo")
            ) == detention_equipment,
            axis=1,
        )
    ].copy()
    if candidates.empty:
        return "—"

    detention_operational_date = _record_operational_date(detention)
    if detention_operational_date is not None:
        candidate_dates = candidates.apply(
            lambda row: _record_operational_date(row.to_dict()), axis=1
        )
        plausible_mask = candidate_dates.apply(
            lambda value: value is None
            or abs((value - detention_operational_date).days) <= 1
        )
        candidates = candidates[plausible_mask].copy()
    if candidates.empty:
        return "—"

    candidates["_candidate_score"] = candidates.apply(
        lambda row: similarity(detention, row.to_dict()), axis=1
    )
    candidates = candidates.sort_values(
        ["_candidate_score", "fecha_hora_recepcion"],
        ascending=[False, False],
        na_position="last",
    )
    best = candidates.iloc[0]
    if float(best.get("_candidate_score", 0) or 0) < minimum_score:
        return "—"
    return _format_turn_value(best.get("turno"))


def daily_summary(detentions: pd.DataFrame, associations: pd.DataFrame) -> pd.DataFrame:
    columns = ["Fecha", "Detenciones válidas", "Detenciones con OT", "Pendientes", "Cumplimiento (%)"]
    if detentions.empty:
        return pd.DataFrame(columns=columns)

    valid = detentions[detentions["requiere_ot"].fillna(False).astype(bool)].copy()
    if valid.empty:
        return pd.DataFrame(columns=columns)

    associated_ids = set()
    if not associations.empty and "detencion_id" in associations.columns:
        associated_ids = set(associations["detencion_id"].dropna().astype(str))
    valid["con_ot"] = valid["id"].astype(str).isin(associated_ids)
    if "fecha_operacional" in valid.columns:
        valid["fecha"] = pd.to_datetime(valid["fecha_operacional"], errors="coerce").dt.date
    else:
        valid["fecha"] = pd.NaT

    missing_operational = valid["fecha"].isna()
    if missing_operational.any():
        valid.loc[missing_operational, "fecha"] = valid.loc[missing_operational].apply(
            lambda row: _record_operational_date(row.to_dict()), axis=1
        )

    grouped = valid.groupby("fecha", dropna=False).agg(
        detenciones_validas=("id", "count"),
        con_ot=("con_ot", "sum"),
    ).reset_index()
    grouped["pendientes"] = grouped["detenciones_validas"] - grouped["con_ot"]
    grouped["cumplimiento"] = (grouped["con_ot"] / grouped["detenciones_validas"] * 100).round(1)
    grouped = grouped.sort_values("fecha")
    return pd.DataFrame({
        "Fecha": grouped["fecha"].apply(lambda x: x.strftime("%d/%m/%Y") if pd.notna(x) else ""),
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
        # Una jornada operacional puede contener registros físicos hasta las 07:59
        # del día calendario siguiente; por eso se consulta hasta end + 1 día y
        # luego se filtra por fecha_operacional en memoria.
        response = self.client.table("detenciones").select("*").gte(
            "fecha_detencion", start.isoformat()
        ).lte("fecha_detencion", (end + timedelta(days=1)).isoformat()).order(
            "fecha_hora_inicio"
        ).execute()
        df = pd.DataFrame(response.data or [])
        if df.empty:
            return df

        if "fecha_operacional" not in df.columns:
            df["fecha_operacional"] = None
        df["_fecha_operacional"] = df.apply(
            lambda row: _record_operational_date(row.to_dict()), axis=1
        )
        mask = df["_fecha_operacional"].apply(
            lambda value: value is not None and start <= value <= end
        )
        df = df[mask].copy()
        df["fecha_operacional"] = df["_fecha_operacional"].apply(
            lambda value: value.isoformat() if value is not None else None
        )
        return df.drop(columns=["_fecha_operacional"], errors="ignore")

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

    def auto_associate(self, start: date, end: date, threshold: float = 0.62) -> int:
        detentions = self.list_detentions(start, end)
        if detentions.empty:
            return 0

        valid_detentions = detentions[
            detentions["requiere_ot"].fillna(False).astype(bool)
        ].copy()
        if valid_detentions.empty:
            return 0

        associations = self.list_associations()
        already_associated_ot_ids = set()
        if not associations.empty and "ot_id" in associations.columns:
            already_associated_ot_ids = set(associations["ot_id"].dropna().astype(str))

        ots = self.list_work_orders()
        if ots.empty:
            return 0
        ots = ots[~ots["id"].astype(str).isin(already_associated_ot_ids)].copy()
        if ots.empty:
            return 0

        count = 0
        for _, ot in ots.iterrows():
            ot_equipment = normalize_equipment(
                ot.get("equipo_normalizado") or ot.get("equipo")
            )
            candidates = valid_detentions[
                valid_detentions.apply(
                    lambda row: normalize_equipment(
                        row.get("equipo_normalizado") or row.get("equipo")
                    ) == ot_equipment,
                    axis=1,
                )
            ]
            if candidates.empty:
                continue

            ot_operational_date = _record_operational_date(ot.to_dict())
            if ot_operational_date is not None:
                candidate_dates = candidates.apply(
                    lambda row: _record_operational_date(row.to_dict()), axis=1
                )
                candidates = candidates[
                    candidate_dates.apply(
                        lambda value: value is not None
                        and abs((value - ot_operational_date).days) <= 1
                    )
                ]
            if candidates.empty:
                continue

            scored = [
                (similarity(det.to_dict(), ot.to_dict()), det)
                for _, det in candidates.iterrows()
            ]
            scored.sort(key=lambda item: item[0], reverse=True)
            best_score, best_det = scored[0]

            should_associate = best_score >= threshold

            # Compatibilidad con OT guardadas por versiones antiguas, donde aún
            # no existían fecha_hora_inicio/termino. Si el equipo tiene una sola
            # detención candidata y la descripción técnica coincide, la relación
            # puede resolverse sin obligar a volver a cargar inmediatamente la OT.
            if not should_associate and len(candidates) == 1:
                ot_has_datetime = _parse_iso_datetime(ot.get("fecha_hora_inicio")) is not None
                det_description = best_det.get("descripcion_normalizada") or " ".join(
                    [safe_text(best_det.get("razon")), safe_text(best_det.get("comentario"))]
                )
                ot_description = ot.get("descripcion_normalizada") or ot.get("descripcion")
                text_score = _text_similarity(det_description, ot_description)
                if not ot_has_datetime and text_score >= 0.30:
                    best_score = max(best_score, 0.62)
                    should_associate = True

            if should_associate:
                self.associate(best_det["id"], ot["id"], best_score, "AUTOMATICA")
                count += 1

        return count
