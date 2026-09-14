import hashlib
import io
import json
import re
import unicodedata
import zipfile

from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from numbers import Number
from pathlib import Path

import pandas as pd
import streamlit as st


# ============================================================
# CONFIGURACIÓN GENERAL
# ============================================================

st.set_page_config(
    page_title="Control presupuestario",
    page_icon="📊",
    layout="wide",
)

APP_VERSION = "3.0"

MOTHER_SHEET = "MAYOR CONSOLIDADO"
EMPTY_OPTION = "— Seleccionar —"

GROUP_KEYS = [
    "Mes",
    "Responsable",
    "Cuenta",
    "Locación",
]

BUDGET_SUM = "Sumar partidas presupuestarias"
BUDGET_REPEATED = (
    "El mismo presupuesto está repetido en los movimientos"
)

EXCEL_MIME = (
    "application/vnd.openxmlformats-officedocument."
    "spreadsheetml.sheet"
)

ALIASES = {
    "Mes": [
        "MES",
        "PERIODO",
        "PERÍODO",
        "FECHA CONTABLE",
        "FECHA",
    ],
    "Responsable": [
        "RESPONSABLE",
        "RESPONSABLE PRESUPUESTO",
        "RESPONSABLE DE PRESUPUESTO",
        "RESPONSABLE PRESUPUESTARIO",
    ],
    "Cuenta": [
        "CUENTA",
        "CUENTA CONTABLE",
        "CUENTA CONT",
        "CTA CBLE",
        "CODIGO CUENTA",
    ],
    "Locación": [
        "LOCACION",
        "LOCACIÓN",
        "LOCALIDAD",
        "SUCURSAL",
    ],
    "Descripción": [
        "DESC CTA CBLE",
        "CUENTA DES",
        "DESC CUENTA",
        "DESCRIPCION CUENTA",
        "DESCRIPCIÓN CUENTA",
        "DESCRIPCION DE CUENTA",
    ],
    "Gasto": [
        "GASTO",
        "REAL",
        "IMPORTE",
        "SALDO",
        "MONTO",
        "NETO",
        "SUMA DE IMPORTE",
    ],
    "Presupuesto": [
        "PRESUPUESTO",
        "PPTO",
        "PRESUP",
        "BUDGET",
        "IMPORTE PRESUPUESTO",
    ],
}

TYPE_ALIASES = {
    "Gasto": {
        "GASTO",
        "GASTOS",
        "REAL",
        "REALES",
        "GASTO REAL",
        "GASTOS REALES",
        "EJECUTADO",
    },
    "Presupuesto": {
        "PRESUPUESTO",
        "PRESUPUESTOS",
        "PRESUP",
        "PPTO",
        "BUDGET",
        "PRESUPUESTADO",
    },
}

MONTHS = {
    "ENE": 1,
    "FEB": 2,
    "MAR": 3,
    "ABR": 4,
    "MAY": 5,
    "JUN": 6,
    "JUL": 7,
    "AGO": 8,
    "SEP": 9,
    "SET": 9,
    "OCT": 10,
    "NOV": 11,
    "DIC": 12,
}

MONEY_COLUMNS = {
    "Gasto",
    "Presupuesto",
    "Desvío $",
    "Importe",
    "Importe omitido",
    "Importe neto omitido",
    "Importe absoluto omitido",
}

st.markdown(
    """
    <style>
    .block-container {
        padding-top: 2rem;
        max-width: 1600px;
    }

    [data-testid="stMetric"] {
        background: #f3f6fa;
        border: 1px solid #dfe6ef;
        border-radius: 12px;
        padding: 18px;
    }

    [data-testid="stMetricLabel"] {
        color: #41506a;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# FORMATO Y CONVERSIÓN DE DATOS
# ============================================================

def text(value):
    if value is None or pd.isna(value):
        return ""

    if isinstance(value, Number) and not isinstance(value, bool):
        try:
            if float(value).is_integer():
                return str(int(value))
        except (ValueError, TypeError, OverflowError):
            pass

    return re.sub(r"\s+", " ", str(value)).strip()


def norm(value):
    value = unicodedata.normalize("NFKD", text(value))
    value = "".join(
        character
        for character in value
        if not unicodedata.combining(character)
    )

    return re.sub(
        r"[_\s]+",
        " ",
        value,
    ).upper().strip()


def header_key(value):
    return re.sub(
        r"[^A-Z0-9]+",
        " ",
        norm(value),
    ).strip()


def identifier(value):
    return text(value).upper()


def currency(value):
    if value is None or pd.isna(value):
        return "No comparable"

    formatted = f"{float(value):,.2f}"

    return (
        "$ "
        + formatted
        .replace(",", "_")
        .replace(".", ",")
        .replace("_", ".")
    )


def percentage(value):
    if value is None or pd.isna(value):
        return "No calculable"

    return f"{float(value):.1%}".replace(".", ",")


def cents(value):
    """
    Convierte importes numéricos o de texto a centavos.
    Conserva el signo del movimiento.
    """
    if value is None or pd.isna(value) or text(value) == "":
        return None

    if isinstance(value, bool):
        raise ValueError("Un importe contiene un valor lógico.")

    if isinstance(value, (Number, Decimal)):
        try:
            number = Decimal(str(value))
        except InvalidOperation:
            raise ValueError(
                f"Importe no reconocido: {value!r}"
            )
    else:
        source = text(value).upper()

        source = (
            source
            .replace("ARS", "")
            .replace("$", "")
            .replace(" ", "")
        )

        if source in {"-", "–", "—"}:
            return 0

        if source.startswith("(") and source.endswith(")"):
            source = "-" + source[1:-1]

        if "," in source and "." in source:
            if source.rfind(",") > source.rfind("."):
                source = (
                    source
                    .replace(".", "")
                    .replace(",", ".")
                )
            else:
                source = source.replace(",", "")

        elif "," in source:
            source = source.replace(",", ".")

        elif re.fullmatch(
            r"-?\d{1,3}(?:\.\d{3})+",
            source,
        ):
            source = source.replace(".", "")

        try:
            number = Decimal(source)
        except InvalidOperation:
            raise ValueError(
                f"Importe no reconocido: {value!r}"
            )

    if not number.is_finite():
        raise ValueError(
            f"Importe no finito: {value!r}"
        )

    rounded = number.quantize(
        Decimal("0.01"),
        rounding=ROUND_HALF_UP,
    )

    return int(rounded * 100)


def period(value, default_year):
    """
    Admite fechas Excel, fechas de texto, julio-26,
    jul-26, 2026-07, 07/2026 y meses numéricos.
    """
    if value is None or pd.isna(value):
        return None

    if isinstance(value, bool):
        return None

    if isinstance(
        value,
        (datetime, date, pd.Timestamp),
    ):
        return pd.Timestamp(value).strftime("%Y-%m")

    if isinstance(value, Number):
        number = float(value)

        if number.is_integer() and 1 <= number <= 12:
            return f"{default_year}-{int(number):02d}"

        if number.is_integer():
            integer = int(number)

            if 200001 <= integer <= 210012:
                year = integer // 100
                month = integer % 100

                if 1 <= month <= 12:
                    return f"{year}-{month:02d}"

        if 20000 <= number <= 100000:
            timestamp = (
                pd.Timestamp("1899-12-30")
                + pd.Timedelta(days=number)
            )

            return timestamp.strftime("%Y-%m")

        return None

    source = norm(value)

    if re.fullmatch(r"\d+(?:\.\d+)?", source):
        return period(float(source), default_year)

    match = re.fullmatch(
        r"([A-Z]+)\.?[\s/-]*(\d{2}|\d{4})?",
        source,
    )

    if match and match.group(1)[:3] in MONTHS:
        month = MONTHS[match.group(1)[:3]]

        year = (
            int(match.group(2))
            if match.group(2)
            else default_year
        )

        if year < 100:
            year += 2000

        return f"{year}-{month:02d}"

    match = re.fullmatch(
        r"(\d{1,2})[/\-](\d{4})",
        source,
    )

    if match:
        month, year = map(int, match.groups())

        if 1 <= month <= 12:
            return f"{year}-{month:02d}"

        return None

    if re.match(r"^\d{4}-", source):
        timestamp = pd.to_datetime(
            source,
            errors="coerce",
        )
    else:
        timestamp = pd.to_datetime(
            source,
            dayfirst=True,
            errors="coerce",
        )

    if pd.isna(timestamp):
        return None

    return timestamp.strftime("%Y-%m")


# ============================================================
# LECTURA DEL EXCEL
# ============================================================

@st.cache_data(
    show_spinner=False,
    max_entries=2,
)
def get_sheet_names(content):
    with pd.ExcelFile(
        io.BytesIO(content),
        engine="openpyxl",
    ) as workbook:
        return workbook.sheet_names


@st.cache_data(
    show_spinner=False,
    max_entries=2,
)
def detect_header(content, sheet):
    preview = pd.read_excel(
        io.BytesIO(content),
        sheet_name=sheet,
        header=None,
        nrows=40,
        dtype=object,
        engine="openpyxl",
    )

    for index, row in preview.iterrows():
        if any(
            header_key(value) == "TIPO"
            for value in row
        ):
            return int(index) + 1

    return 1


@st.cache_data(
    show_spinner=False,
    max_entries=2,
)
def read_table(content, sheet, header):
    frame = pd.read_excel(
        io.BytesIO(content),
        sheet_name=sheet,
        header=header - 1,
        dtype=object,
        engine="openpyxl",
    )

    frame.columns = [
        str(column).strip()
        for column in frame.columns
    ]

    if frame.columns.duplicated().any():
        raise ValueError(
            "Hay encabezados duplicados. "
            "Las columnas deben tener nombres únicos."
        )

    # Se conserva el índice para identificar la fila original.
    return frame.dropna(how="all")


def choose(label, options, default, key):
    if key in st.session_state:
        if st.session_state[key] not in options:
            del st.session_state[key]

    kwargs = {}

    if key not in st.session_state:
        kwargs["index"] = (
            options.index(default)
            if default in options
            else 0
        )

    return st.selectbox(
        label,
        options,
        key=key,
        **kwargs,
    )


def choose_many(label, options, defaults, key):
    if key in st.session_state:
        previous = list(st.session_state[key])

        current = [
            value
            for value in previous
            if value in options
        ]

        if current != previous:
            st.session_state[key] = current

    kwargs = {}

    if key not in st.session_state:
        kwargs["default"] = [
            value
            for value in defaults
            if value in options
        ]

    return st.multiselect(
        label,
        options,
        key=key,
        format_func=lambda value: value or "(vacío)",
        **kwargs,
    )


def detect_column(columns, aliases):
    for alias in aliases:
        target = header_key(alias)

        for column in columns:
            if header_key(column) == target:
                return column

    return None


def source_config(
    content,
    sheet,
    role,
    saved,
    suggested_header,
):
    prefix = f"src_{role}_"

    st.caption(
        f"Hoja de origen: {sheet}. "
        "La separación se realiza mediante la columna Tipo."
    )

    header = st.number_input(
        "Fila que contiene los encabezados",
        min_value=1,
        max_value=200,
        value=int(
            saved.get(
                "header",
                suggested_header,
            )
        ),
        key=prefix + "header",
    )

    frame = read_table(
        content,
        sheet,
        int(header),
    )

    st.dataframe(
        frame.head(6),
        use_container_width=True,
        hide_index=True,
    )

    type_column = detect_column(
        frame.columns,
        ["TIPO"],
    )

    if not type_column:
        raise ValueError(
            f"{role}: no se encontró la columna Tipo "
            f"en la fila de encabezados {header}."
        )

    available_types = sorted(
        frame[type_column]
        .map(text)
        .unique()
        .tolist()
    )

    saved_types = []

    if header_key(
        saved.get("filter_column", "")
    ) == "TIPO":
        saved_types = saved.get(
            "filter_values",
            [],
        )

    if saved_types:
        normalized_defaults = {
            norm(value)
            for value in saved_types
        }
    else:
        normalized_defaults = TYPE_ALIASES[role]

    default_types = [
        value
        for value in available_types
        if norm(value) in normalized_defaults
    ]

    selected_types = choose_many(
        f"Valores de Tipo que corresponden a {role.lower()}",
        available_types,
        default_types,
        prefix + "types",
    )

    options = [EMPTY_OPTION] + list(frame.columns)
    saved_mapping = saved.get("mapping", {})

    mapping = {}
    columns = st.columns(2)

    fields = [
        "Mes",
        "Responsable",
        "Cuenta",
        role,
        "Locación",
        "Descripción",
    ]

    for index, field in enumerate(fields):
        detected = detect_column(
            frame.columns,
            ALIASES[field],
        )

        if field == "Presupuesto" and not detected:
            detected = detect_column(
                frame.columns,
                ALIASES["Gasto"],
            )

        default = (
            saved_mapping.get(field)
            or detected
            or EMPTY_OPTION
        )

        label = field

        if field == role:
            label = f"Columna del importe de {role.lower()}"

        if field in ["Locación", "Descripción"]:
            label += " — opcional"

        with columns[index % 2]:
            selected = choose(
                label,
                options,
                default,
                prefix + field,
            )

        mapping[field] = (
            None
            if selected == EMPTY_OPTION
            else selected
        )

    year = st.number_input(
        "Año para meses que no incluyen año",
        min_value=2000,
        max_value=2100,
        value=int(
            saved.get("year", date.today().year)
        ),
        key=prefix + "year",
    )

    skip_totals = st.checkbox(
        "Excluir filas rotuladas como Total o Subtotal",
        value=bool(
            saved.get("skip_totals", True)
        ),
        key=prefix + "skip_totals",
    )

    invert = st.checkbox(
        "Invertir el signo de los importes de esta fuente",
        value=bool(
            saved.get("invert", False)
        ),
        key=prefix + "invert",
    )

    mode = None

    if role == "Presupuesto":
        mode = choose(
            "Cómo está registrado el presupuesto",
            [BUDGET_SUM, BUDGET_REPEATED],
            saved.get("mode", BUDGET_SUM),
            prefix + "mode",
        )

        st.caption(
            "Usá sumar partidas cuando cada fila de Tipo "
            "Presupuesto sea una asignación independiente."
        )

    configuration = {
        "sheet": sheet,
        "header": int(header),
        "mapping": mapping,
        "year": int(year),
        "filter_column": type_column,
        "filter_values": selected_types,
        "skip_totals": skip_totals,
        "invert": invert,
        "mode": mode,
    }

    return frame, configuration


# ============================================================
# NORMALIZACIÓN Y OMISIÓN DE FILAS INCOMPLETAS
# ============================================================

def normalize_source(frame, config, role):
    mapping = config["mapping"]

    for field in ["Mes", "Responsable", "Cuenta", role]:
        if not mapping.get(field):
            raise ValueError(
                f"{role}: falta seleccionar la columna {field}."
            )

    if not config["filter_values"]:
        raise ValueError(
            f"{role}: seleccioná los valores de Tipo "
            "que pertenecen a esta fuente."
        )

    selected_types = {
        norm(value)
        for value in config["filter_values"]
    }

    # El filtro por Tipo se aplica antes de validar las filas.
    raw = frame[
        frame[config["filter_column"]]
        .map(norm)
        .isin(selected_types)
    ].copy()

    removed_totals = 0

    if config["skip_totals"]:
        total_mask = pd.Series(
            False,
            index=raw.index,
        )

        for field in [
            "Cuenta",
            "Descripción",
            "Responsable",
        ]:
            column = mapping.get(field)

            if column:
                total_mask |= (
                    raw[column]
                    .map(norm)
                    .str.match(r"^(?:TOTAL|SUBTOTAL)")
                )

        removed_totals = int(total_mask.sum())
        raw = raw.loc[~total_mask]

    parsed = []

    for index, value in raw[mapping[role]].items():
        try:
            parsed.append(cents(value))

        except ValueError as error:
            row = int(index) + config["header"] + 1

            raise ValueError(
                f"{role}, fila Excel {row}: {error}"
            ) from error

    amounts = pd.Series(
        parsed,
        index=raw.index,
        dtype=object,
    )

    raw = raw.loc[amounts.notna()]
    amounts = amounts.loc[raw.index]

    output = pd.DataFrame(index=raw.index)

    output["Mes"] = raw[mapping["Mes"]].map(
        lambda value: period(
            value,
            config["year"],
        )
    )

    output["Responsable"] = (
        raw[mapping["Responsable"]]
        .map(identifier)
    )

    output["Cuenta"] = (
        raw[mapping["Cuenta"]]
        .map(identifier)
    )

    location = mapping.get("Locación")
    description = mapping.get("Descripción")

    output["Locación"] = (
        raw[location]
        .map(identifier)
        .replace("", "SIN LOCACIÓN")
        if location
        else "SIN DESGLOSE"
    )

    output["Descripción"] = (
        raw[description].map(text)
        if description
        else ""
    )

    output["_centavos"] = (
        amounts.astype("int64")
        * (-1 if config["invert"] else 1)
    )

    output["Fila Excel"] = (
        raw.index + config["header"] + 1
    )

    output["Hoja origen"] = config["sheet"]

    original = raw.rename(
        columns=lambda column: "Origen · " + column
    )

    output = pd.concat(
        [output, original],
        axis=1,
    ).reset_index(drop=True)

    input_count = len(output)

    invalid = (
        output["Mes"].isna()
        | output["Cuenta"].eq("")
        | output["Responsable"].eq("")
    )

    omitted = output.loc[invalid].copy()

    if (
        not omitted.empty
        and not config.get("skip_incomplete", False)
    ):
        rows = (
            omitted["Fila Excel"]
            .head(12)
            .astype(str)
            .tolist()
        )

        raise ValueError(
            f"{role}: hay {len(omitted)} filas con importe "
            "pero sin mes, cuenta o responsable válido. "
            f"Filas Excel: {', '.join(rows)}. "
            "Para continuar, activá "
            "'Omitir filas incompletas y continuar'."
        )

    omitted["Fuente"] = role
    omitted["Motivo de omisión"] = ""

    if not omitted.empty:
        def omission_reason(row):
            reasons = []

            if pd.isna(row["Mes"]):
                reasons.append(
                    "Mes faltante o no reconocido"
                )

            if row["Cuenta"] == "":
                reasons.append("Cuenta vacía")

            if row["Responsable"] == "":
                reasons.append("Responsable vacío")

            return "; ".join(reasons)

        omitted["Motivo de omisión"] = (
            omitted.apply(
                omission_reason,
                axis=1,
            )
        )

    omitted_net = (
        omitted["_centavos"].sum() / 100
    )

    omitted_absolute = (
        omitted["_centavos"].abs().sum() / 100
    )

    omitted["Importe omitido"] = (
        omitted["_centavos"] / 100
    )

    first_columns = [
        "Fuente",
        "Hoja origen",
        "Fila Excel",
        "Motivo de omisión",
        "Importe omitido",
    ]

    remaining_columns = [
        column
        for column in omitted.columns
        if column not in first_columns
        and column != "_centavos"
    ]

    omitted = omitted[
        first_columns + remaining_columns
    ].reset_index(drop=True)

    # Solo las filas válidas pasan a los cálculos.
    output = output.loc[
        ~invalid
    ].reset_index(drop=True)

    repeated = 0

    if (
        role == "Presupuesto"
        and config["mode"] == BUDGET_REPEATED
    ):
        distinct = (
            output.groupby(GROUP_KEYS)["_centavos"]
            .nunique()
        )

        conflicts = distinct[distinct > 1]

        if not conflicts.empty:
            raise ValueError(
                "Hay presupuestos diferentes para una misma "
                "combinación de mes, responsable, cuenta y "
                "locación. Revisá el origen o elegí sumar "
                "partidas si son asignaciones independientes."
            )

        before_deduplication = len(output)

        output = output.drop_duplicates(
            GROUP_KEYS
        ).reset_index(drop=True)

        repeated = (
            before_deduplication - len(output)
        )

    audit = {
        "Fuente": role,
        "Filas con importe antes de consolidar": input_count,
        "Filas de total excluidas": removed_totals,
        "Filas incompletas omitidas": len(omitted),
        "Importe neto omitido": omitted_net,
        "Importe absoluto omitido": omitted_absolute,
        "Repeticiones presupuestarias excluidas": repeated,
        "Filas utilizadas": len(output),
        "Importe": output["_centavos"].sum() / 100,
    }

    return output, audit, omitted


# ============================================================
# COMPARACIONES Y TABLAS
# ============================================================

def comparison(
    real,
    budget,
    dimensions,
    comparable=True,
):
    actual = (
        real.groupby(
            dimensions,
            dropna=False,
        )["_centavos"]
        .sum()
        .rename("Gasto")
        .reset_index()
    )

    if not comparable:
        result = actual.copy()
        result["Gasto"] = result["Gasto"] / 100

        for column in [
            "Presupuesto",
            "Desvío $",
            "Desvío %",
        ]:
            result[column] = float("nan")

        result["Estado"] = (
            "Presupuesto sin desglose comparable"
        )

        return result

    planned = (
        budget.groupby(
            dimensions,
            dropna=False,
        )["_centavos"]
        .sum()
        .rename("Presupuesto")
        .reset_index()
    )

    result = actual.merge(
        planned,
        on=dimensions,
        how="outer",
    )

    result[["Gasto", "Presupuesto"]] = (
        result[["Gasto", "Presupuesto"]]
        .fillna(0)
    )

    difference_cents = (
        result["Gasto"]
        - result["Presupuesto"]
    )

    result["Desvío $"] = (
        difference_cents / 100
    )

    result["Desvío %"] = (
        difference_cents
        / result["Presupuesto"].where(
            result["Presupuesto"] > 0
        )
    )

    result[["Gasto", "Presupuesto"]] = (
        result[["Gasto", "Presupuesto"]] / 100
    )

    result["Estado"] = "Dentro del presupuesto"

    result.loc[
        result["Desvío $"] > 0,
        "Estado",
    ] = "Exceso de gasto"

    result.loc[
        result["Presupuesto"] == 0,
        "Estado",
    ] = "Sin presupuesto positivo cargado"

    result.loc[
        result["Presupuesto"] < 0,
        "Estado",
    ] = "Presupuesto negativo: revisar"

    return result.sort_values(
        dimensions
    ).reset_index(drop=True)


def account_descriptions(real, budget):
    descriptions = pd.concat(
        [
            real[["Cuenta", "Descripción"]],
            budget[["Cuenta", "Descripción"]],
        ],
        ignore_index=True,
    )

    return descriptions[
        descriptions["Descripción"].ne("")
    ].drop_duplicates("Cuenta")


def detail_export(frame, amount_name):
    output = frame.drop(
        columns="_centavos"
    ).copy()

    output.insert(
        0,
        amount_name,
        frame["_centavos"].to_numpy() / 100,
    )

    return output


def show_summary(table):
    view = table.copy()
    column_config = {}

    for column in MONEY_COLUMNS:
        if column in view.columns:
            column_config[column] = (
                st.column_config.NumberColumn(
                    column,
                    format="$ %.2f",
                )
            )

    if "Desvío %" in view.columns:
        view["Desvío %"] = (
            view["Desvío %"] * 100
        )

        column_config["Desvío %"] = (
            st.column_config.NumberColumn(
                "Desvío %",
                format="%.1f%%",
            )
        )

    st.dataframe(
        view,
        column_config=column_config,
        use_container_width=True,
        hide_index=True,
    )


# ============================================================
# GENERACIÓN DE ARCHIVOS EXCEL
# ============================================================

def create_excel(tables, include_chart=False):
    buffer = io.BytesIO()

    with pd.ExcelWriter(
        buffer,
        engine="xlsxwriter",
        engine_kwargs={
            "options": {
                "strings_to_formulas": False,
                "strings_to_urls": False,
            }
        },
    ) as writer:
        workbook = writer.book

        header_format = workbook.add_format({
            "bold": True,
            "bg_color": "#193451",
            "font_color": "#FFFFFF",
            "text_wrap": True,
            "valign": "vcenter",
        })

        money_format = workbook.add_format({
            "num_format": "#,##0.00;[Red]-#,##0.00"
        })

        percent_format = workbook.add_format({
            "num_format": "0.0%;[Red]-0.0%"
        })

        text_format = workbook.add_format({
            "text_wrap": True,
            "valign": "top",
        })

        excess_format = workbook.add_format({
            "bg_color": "#FCE4E4",
            "font_color": "#9C2020",
        })

        for name, table in tables.items():
            table.to_excel(
                writer,
                sheet_name=name,
                index=False,
            )

            worksheet = writer.sheets[name]
            worksheet.freeze_panes(1, 0)
            worksheet.set_row(0, 32)

            for index, column in enumerate(table.columns):
                worksheet.write(
                    0,
                    index,
                    column,
                    header_format,
                )

                width = min(
                    48,
                    max(16, len(str(column)) + 3),
                )

                cell_format = None

                if column in MONEY_COLUMNS:
                    cell_format = money_format

                elif column == "Desvío %":
                    cell_format = percent_format

                worksheet.set_column(
                    index,
                    index,
                    width,
                    cell_format,
                )

            if not table.empty:
                worksheet.autofilter(
                    0,
                    0,
                    len(table),
                    len(table.columns) - 1,
                )

            if (
                "Desvío $" in table.columns
                and not table.empty
            ):
                column = table.columns.get_loc(
                    "Desvío $"
                )

                worksheet.conditional_format(
                    1,
                    column,
                    len(table),
                    column,
                    {
                        "type": "cell",
                        "criteria": ">",
                        "value": 0,
                        "format": excess_format,
                    },
                )

            if name == "Resumen":
                worksheet.set_column("A:A", 28)
                worksheet.set_column("B:B", 72)

                for row_number, values in enumerate(
                    table.itertuples(
                        index=False,
                        name=None,
                    ),
                    start=1,
                ):
                    label, value = values

                    if label in {
                        "Presupuesto",
                        "Gasto real",
                        "Desvío $",
                    }:
                        worksheet.write_number(
                            row_number,
                            1,
                            float(value),
                            money_format,
                        )

                    elif label == "Desvío %":
                        if value is None or pd.isna(value):
                            worksheet.write(
                                row_number,
                                1,
                                "No calculable",
                                text_format,
                            )
                        else:
                            worksheet.write_number(
                                row_number,
                                1,
                                float(value),
                                percent_format,
                            )

                    else:
                        worksheet.write(
                            row_number,
                            1,
                            text(value),
                            text_format,
                        )

                        worksheet.set_row(
                            row_number,
                            34,
                        )

        if (
            include_chart
            and "Evolucion" in tables
            and not tables["Evolucion"].empty
        ):
            monthly = tables["Evolucion"]

            chart = workbook.add_chart({
                "type": "column"
            })

            for column_name, color in [
                ("Presupuesto", "#8BA6C2"),
                ("Gasto", "#244D78"),
            ]:
                column_index = (
                    monthly.columns.get_loc(
                        column_name
                    )
                )

                chart.add_series({
                    "name": column_name,
                    "categories": [
                        "Evolucion",
                        1,
                        0,
                        len(monthly),
                        0,
                    ],
                    "values": [
                        "Evolucion",
                        1,
                        column_index,
                        len(monthly),
                        column_index,
                    ],
                    "fill": {"color": color},
                    "border": {"color": color},
                })

            chart.set_title({
                "name": "Presupuesto y gasto"
            })

            chart.set_legend({
                "position": "bottom"
            })

            writer.sheets["Resumen"].insert_chart(
                "D2",
                chart,
            )

    return buffer.getvalue()


def report(
    real,
    budget,
    owner,
    months,
    source_name,
    integrity_note,
):
    owner_real = real[
        real["Responsable"].eq(owner)
    ].copy()

    owner_budget = budget[
        budget["Responsable"].eq(owner)
    ].copy()

    accounts = comparison(
        owner_real,
        owner_budget,
        ["Cuenta"],
    )

    accounts = accounts.merge(
        account_descriptions(
            owner_real,
            owner_budget,
        ),
        on="Cuenta",
        how="left",
    )

    monthly = comparison(
        owner_real,
        owner_budget,
        ["Mes"],
    )

    spent_cents = (
        owner_real["_centavos"].sum()
    )

    planned_cents = (
        owner_budget["_centavos"].sum()
    )

    difference_cents = (
        spent_cents - planned_cents
    )

    spent = spent_cents / 100
    planned = planned_cents / 100
    variance = difference_cents / 100

    ratio = (
        difference_cents / planned_cents
        if planned_cents > 0
        else None
    )

    summary = pd.DataFrame(
        [
            ("Responsable", owner),
            ("Períodos incluidos", ", ".join(months)),
            ("Archivo fuente", source_name),
            ("Presupuesto", planned),
            ("Gasto real", spent),
            ("Desvío $", variance),
            ("Desvío %", ratio),
            (
                "Criterio",
                "Desvío = gasto − presupuesto. "
                "Positivo indica exceso.",
            ),
            (
                "Estado de validación",
                integrity_note
                or "Sin omisiones por mes, cuenta o responsable.",
            ),
        ],
        columns=["Indicador", "Valor"],
    )

    tables = {
        "Resumen": summary,
        "Cuentas": accounts,
        "Evolucion": monthly,
        "Detalle contable": detail_export(
            owner_real,
            "Gasto",
        ),
        "Base presupuesto": detail_export(
            owner_budget,
            "Presupuesto",
        ),
    }

    xlsx = create_excel(
        tables,
        include_chart=True,
    )

    without_budget = accounts[
        (accounts["Presupuesto"] == 0)
        & (accounts["Gasto"] != 0)
    ]

    budget_note = ""

    if not without_budget.empty:
        budget_note = (
            f"\nHay {len(without_budget)} cuenta(s) "
            "con movimientos y sin presupuesto positivo "
            "cargado en este reporte.\n"
        )

    integrity_paragraph = (
        f"{integrity_note}\n\n"
        if integrity_note
        else ""
    )

    email = (
        f"Asunto: Control presupuestario | "
        f"{', '.join(months)} | {owner}\n\n"
        "Buen día,\n\n"
        f"{integrity_paragraph}"
        "Adjunto el reporte presupuestario del período "
        "indicado, con el resumen por cuenta y el detalle "
        "de movimientos.\n\n"
        f"Presupuesto: {currency(planned)}\n"
        f"Gasto real: {currency(spent)}\n"
        f"Desvío: {currency(variance)}\n"
        f"Desvío porcentual: {percentage(ratio)}\n"
        f"{budget_note}\n"
        "Un desvío positivo indica gasto por encima del "
        "presupuesto. Un desvío negativo indica una menor "
        "ejecución presupuestaria.\n\n"
        "Por favor, revisen los movimientos e informen "
        "las explicaciones o ajustes de imputación "
        "que correspondan.\n\n"
        "Saludos,\n"
        "Damián"
    )

    return xlsx, email


def safe_name(value):
    cleaned = re.sub(
        r'[<>:"/\\|?*\x00-\x1f]',
        "_",
        value,
    )

    return (
        cleaned.strip(". ")[:100]
        or "reporte"
    )


def period_tag(months):
    ordered = sorted(months)

    if len(ordered) == 1:
        return ordered[0]

    return (
        f"{ordered[0]}_a_{ordered[-1]}"
        f"_{len(ordered)}meses"
    )


def reset_filters(latest_month):
    for key in [
        "f_Responsable",
        "f_Cuenta",
        "f_Locación",
    ]:
        st.session_state[key] = []

    st.session_state["f_Mes"] = [
        latest_month
    ]


# ============================================================
# INTERFAZ PRINCIPAL
# ============================================================

st.title("Control presupuestario")

st.caption(
    "Mayor Consolidado · Gasto y presupuesto por Tipo "
    "· Análisis por responsable · Reportes Excel"
)

uploaded = st.file_uploader(
    "Cargar archivo base",
    type=["xlsx"],
    key="upload",
)

if uploaded is None:
    st.info(
        "Cargá el Excel que contiene la hoja "
        "Mayor Consolidado."
    )
    st.stop()

content = uploaded.getvalue()

file_hash = hashlib.sha256(
    content
).hexdigest()

if (
    st.session_state.get("_file_hash")
    != file_hash
):
    for key in list(st.session_state):
        if key.startswith(
            ("src_", "f_", "out_")
        ):
            del st.session_state[key]

    st.session_state["_file_hash"] = file_hash

try:
    sheets = get_sheet_names(content)

    matching_sheets = [
        sheet
        for sheet in sheets
        if norm(sheet) == MOTHER_SHEET
    ]

    if len(matching_sheets) != 1:
        st.error(
            "No se pudo identificar una única hoja "
            "llamada Mayor Consolidado."
        )

        st.write(
            "Hojas disponibles:",
            sheets,
        )

        st.stop()

    mother_sheet = matching_sheets[0]

    suggested_header = detect_header(
        content,
        mother_sheet,
    )

except Exception as error:
    st.error(
        f"No se pudo abrir el Excel: {error}"
    )
    st.stop()

config_path = Path(__file__).with_name(
    "configuracion.json"
)

saved_config = st.session_state.get(
    "_active_config",
    {},
)

if not saved_config and config_path.exists():
    try:
        loaded_config = json.loads(
            config_path.read_text(
                encoding="utf-8"
            )
        )

        if isinstance(loaded_config, dict):
            saved_config = loaded_config

        else:
            st.warning(
                "configuracion.json no contiene un objeto "
                "de configuración válido. Configurá la lectura."
            )

    except (ValueError, OSError):
        st.warning(
            "No se pudo leer configuracion.json. "
            "Configurá las fuentes desde la aplicación."
        )

with st.expander(
    "Configuración de lectura",
    expanded=not bool(saved_config),
):
    st.write(
        "Ambas fuentes se leen desde Mayor Consolidado. "
        "La columna Tipo determina qué filas corresponden "
        "a gasto y cuáles a presupuesto."
    )

    st.caption(
        "Si ambos tipos usan una misma columna de importes, "
        "seleccioná esa columna en las dos pestañas."
    )

    actual_tab, budget_tab = st.tabs(
        ["Gasto real", "Presupuesto"]
    )

    try:
        with actual_tab:
            actual_table, actual_config = source_config(
                content,
                mother_sheet,
                "Gasto",
                saved_config.get("Gasto", {}),
                suggested_header,
            )

        with budget_tab:
            budget_table, budget_config = source_config(
                content,
                mother_sheet,
                "Presupuesto",
                saved_config.get("Presupuesto", {}),
                suggested_header,
            )

    except Exception as error:
        st.error(
            f"Revisá la configuración de lectura: {error}"
        )
        st.stop()

skip_incomplete = st.checkbox(
    "Omitir filas incompletas y continuar",
    value=bool(
        saved_config.get("Gasto", {}).get(
            "skip_incomplete",
            False,
        )
    ),
    key="src_skip_incomplete",
    help=(
        "Excluye de los cálculos las filas con importe "
        "pero sin mes, cuenta o responsable válido. "
        "Las filas omitidas quedan disponibles para descargar."
    ),
)

actual_config["skip_incomplete"] = (
    skip_incomplete
)

budget_config["skip_incomplete"] = (
    skip_incomplete
)

configuration = {
    "Gasto": actual_config,
    "Presupuesto": budget_config,
    "_app_version": APP_VERSION,
}

st.session_state["_active_config"] = (
    configuration
)

config_json = json.dumps(
    configuration,
    ensure_ascii=False,
    indent=2,
)

st.download_button(
    "Guardar configuración de lectura",
    config_json.encode("utf-8"),
    file_name="configuracion.json",
    mime="application/json",
)

with st.expander(
    "Ver configuración para copiar en GitHub"
):
    st.code(
        config_json,
        language="json",
    )

    st.caption(
        "Para conservar estos ajustes en nuevas sesiones, "
        "actualizá configuracion.json en GitHub con este contenido."
    )

model_hash = hashlib.sha256(
    content
    + json.dumps(
        configuration,
        sort_keys=True,
    ).encode("utf-8")
).hexdigest()

try:
    if (
        st.session_state.get("_model_hash")
        != model_hash
    ):
        if (
            actual_config["header"]
            != budget_config["header"]
        ):
            raise ValueError(
                "Gasto y presupuesto deben usar la misma "
                "fila de encabezados de Mayor Consolidado."
            )

        actual_types = {
            norm(value)
            for value in actual_config["filter_values"]
        }

        budget_types = {
            norm(value)
            for value in budget_config["filter_values"]
        }

        if actual_types & budget_types:
            raise ValueError(
                "Un mismo valor de Tipo está seleccionado "
                "como gasto y como presupuesto. "
                "Asigná cada Tipo a un único escenario."
            )

        with st.spinner(
            "Procesando gasto, presupuesto y controles..."
        ):
            real, real_audit, real_omitted = (
                normalize_source(
                    actual_table,
                    actual_config,
                    "Gasto",
                )
            )

            budget, budget_audit, budget_omitted = (
                normalize_source(
                    budget_table,
                    budget_config,
                    "Presupuesto",
                )
            )

            omitted_rows = pd.concat(
                [
                    real_omitted,
                    budget_omitted,
                ],
                ignore_index=True,
                sort=False,
            )

            new_model = {
                "real": real,
                "budget": budget,
                "audit": pd.DataFrame(
                    [
                        real_audit,
                        budget_audit,
                    ]
                ),
                "omitted": omitted_rows,
            }

        # Actualización conjunta después de validar ambas fuentes.
        st.session_state["_model"] = new_model
        st.session_state["_model_hash"] = model_hash

        for key in [
            "out_individual",
            "out_batch",
            "out_omitted",
        ]:
            st.session_state.pop(
                key,
                None,
            )

except Exception as error:
    st.error(str(error))
    st.stop()

model = st.session_state["_model"]

real = model["real"]
budget = model["budget"]
audit = model["audit"]
omitted_rows = model["omitted"]

integrity_note = ""

if not omitted_rows.empty:
    integrity_note = (
        "Fuente con omisiones; cifras de las filas incluidas."
    )

    st.warning(
        f"Se omitieron {len(omitted_rows)} filas incompletas. "
        "Los indicadores y reportes utilizan las filas válidas "
        "que quedaron incluidas."
    )

    omission_totals = (
        omitted_rows
        .groupby(
            "Fuente",
            as_index=False,
        )
        .agg(
            Filas=("Importe omitido", "size"),
            Importe_neto_omitido=(
                "Importe omitido",
                "sum",
            ),
        )
        .rename(
            columns={
                "Importe_neto_omitido":
                    "Importe neto omitido"
            }
        )
    )

    show_summary(omission_totals)

    with st.expander(
        "Revisar y descargar filas omitidas"
    ):
        st.caption(
            "Este listado corresponde a todo el archivo "
            "cargado. Los filtros del tablero no modifican "
            "las omisiones."
        )

        st.dataframe(
            omitted_rows,
            use_container_width=True,
            hide_index=True,
        )

        omitted_export = st.session_state.get(
            "out_omitted"
        )

        if (
            not omitted_export
            or omitted_export["key"] != model_hash
        ):
            omitted_export = {
                "key": model_hash,
                "data": create_excel({
                    "Filas omitidas": omitted_rows
                }),
            }

            st.session_state["out_omitted"] = (
                omitted_export
            )

        st.download_button(
            "Descargar filas omitidas",
            omitted_export["data"],
            file_name="filas_omitidas_revision.xlsx",
            mime=EXCEL_MIME,
            key="download_omitted_rows",
        )

if real.empty:
    st.info(
        "No quedaron movimientos reales válidos para analizar. "
        "Podés descargar las filas omitidas para revisarlas "
        "o ajustar la configuración de lectura."
    )
    st.stop()

months_available = sorted(
    real["Mes"].unique().tolist()
)

latest_month = months_available[-1]

with st.sidebar:
    st.title("Filtros")

    st.button(
        "Limpiar filtros",
        on_click=reset_filters,
        args=(latest_month,),
    )

    months = choose_many(
        "Mes",
        months_available,
        [latest_month],
        "f_Mes",
    )

    selections = {}

    for dimension in [
        "Responsable",
        "Cuenta",
        "Locación",
    ]:
        options = sorted(
            set(real[dimension].tolist())
            | set(budget[dimension].tolist())
        )

        selections[dimension] = choose_many(
            dimension,
            options,
            [],
            "f_" + dimension,
        )

if not months:
    st.info(
        "Seleccioná al menos un mes."
    )
    st.stop()

selected_months = sorted(months)

period_real = real[
    real["Mes"].isin(selected_months)
].copy()

period_budget = budget[
    budget["Mes"].isin(selected_months)
].copy()

view_real = period_real.copy()
view_budget = period_budget.copy()

for dimension in [
    "Responsable",
    "Cuenta",
]:
    selected = selections[dimension]

    if selected:
        view_real = view_real[
            view_real[dimension].isin(selected)
        ]

        view_budget = view_budget[
            view_budget[dimension].isin(selected)
        ]

location_filter = selections["Locación"]
comparable = True

if location_filter:
    has_location_columns = bool(
        actual_config["mapping"]["Locación"]
        and budget_config["mapping"]["Locación"]
    )

    unallocated_budget = (
        view_budget["Locación"].isin(
            ["SIN LOCACIÓN", "SIN DESGLOSE"]
        )
        & view_budget["_centavos"].ne(0)
    ).any()

    comparable = (
        has_location_columns
        and not unallocated_budget
    )

    view_real = view_real[
        view_real["Locación"].isin(
            location_filter
        )
    ]

    if comparable:
        view_budget = view_budget[
            view_budget["Locación"].isin(
                location_filter
            )
        ]

    else:
        st.info(
            "Se muestra el gasto de las locaciones elegidas. "
            "El presupuesto no tiene un desglose completo "
            "comparable por locación, por lo que se omiten "
            "los indicadores de desvío para esta selección."
        )

spent_cents = (
    view_real["_centavos"].sum()
)

planned_cents = (
    view_budget["_centavos"].sum()
    if comparable
    else None
)

spent = spent_cents / 100

planned = (
    planned_cents / 100
    if comparable
    else None
)

variance = (
    (spent_cents - planned_cents) / 100
    if comparable
    else None
)

ratio = (
    (spent_cents - planned_cents)
    / planned_cents
    if comparable and planned_cents > 0
    else None
)

st.caption(
    "Períodos seleccionados: "
    + ", ".join(selected_months)
)

kpis = st.columns(4)

kpis[0].metric(
    "Presupuesto",
    currency(planned),
)

kpis[1].metric(
    "Gasto real",
    currency(spent),
)

kpis[2].metric(
    "Desvío $",
    currency(variance),
)

kpis[3].metric(
    "Desvío %",
    percentage(ratio),
)

overview, accounts_tab, movements, exports, controls = (
    st.tabs(
        [
            "Panorama",
            "Cuentas",
            "Movimientos",
            "Descargas",
            "Control de importación",
        ]
    )
)


# ============================================================
# PANORAMA
# ============================================================

with overview:
    by_owner = comparison(
        view_real,
        view_budget,
        ["Responsable"],
        comparable,
    )

    st.subheader(
        "Presupuesto y gasto por responsable"
    )

    chart_columns = (
        ["Presupuesto", "Gasto"]
        if comparable
        else ["Gasto"]
    )

    if not by_owner.empty:
        chart = (
            by_owner
            .sort_values(
                "Gasto",
                ascending=False,
            )
            .head(20)
            .set_index("Responsable")[chart_columns]
        )

        st.bar_chart(chart)

    else:
        st.info(
            "No hay datos para esta combinación de filtros."
        )

    show_summary(by_owner)

    st.subheader(
        "Comparación mensual del período seleccionado"
    )

    by_month = comparison(
        view_real,
        view_budget,
        ["Mes"],
        comparable,
    )

    if not by_month.empty:
        st.line_chart(
            by_month
            .set_index("Mes")[chart_columns]
        )

    show_summary(by_month)


# ============================================================
# CUENTAS Y MOVIMIENTOS
# ============================================================

with accounts_tab:
    by_account = comparison(
        view_real,
        view_budget,
        ["Responsable", "Cuenta"],
        comparable,
    )

    by_account = by_account.merge(
        account_descriptions(
            real,
            budget,
        ),
        on="Cuenta",
        how="left",
    )

    show_summary(by_account)

with movements:
    st.caption(
        "Movimientos de gasto incluidos desde Mayor Consolidado, "
        "con sus columnas originales y número de fila de Excel."
    )

    st.dataframe(
        detail_export(
            view_real,
            "Gasto",
        ),
        use_container_width=True,
        hide_index=True,
    )


# ============================================================
# REPORTES INDIVIDUALES Y LOTE
# ============================================================

with exports:
    st.write(
        "Los reportes incluyen todas las cuentas del responsable "
        "en los meses seleccionados."
    )

    st.caption(
        "Los filtros de cuenta, locación y responsable del "
        "tablero no recortan estas descargas. "
        "Las filas incompletas omitidas quedan fuera de los reportes."
    )

    owners = sorted(
        set(period_real["Responsable"])
        | set(period_budget["Responsable"])
    )

    if not owners:
        st.info(
            "No hay responsables para exportar en este período."
        )

    else:
        owner = choose(
            "Responsable a exportar",
            owners,
            owners[0],
            "report_owner",
        )

        export_key = hashlib.sha256(
            json.dumps(
                {
                    "model": model_hash,
                    "months": selected_months,
                    "owner": owner,
                    "source": uploaded.name,
                },
                sort_keys=True,
            ).encode("utf-8")
        ).hexdigest()

        if st.button(
            "Preparar reporte individual",
            type="primary",
        ):
            st.session_state.pop(
                "out_individual",
                None,
            )

            try:
                with st.spinner(
                    "Preparando Excel y texto del correo..."
                ):
                    xlsx, email = report(
                        period_real,
                        period_budget,
                        owner,
                        selected_months,
                        uploaded.name,
                        integrity_note,
                    )

                st.session_state["out_individual"] = {
                    "key": export_key,
                    "xlsx": xlsx,
                    "email": email,
                }

            except Exception as error:
                st.error(
                    f"No se pudo generar el reporte: {error}"
                )

        prepared = st.session_state.get(
            "out_individual",
            {},
        )

        if prepared.get("key") == export_key:
            filename = (
                f"{safe_name(owner)}_"
                f"{period_tag(selected_months)}"
            )

            st.download_button(
                "Descargar Excel",
                prepared["xlsx"],
                file_name=filename + ".xlsx",
                mime=EXCEL_MIME,
            )

            st.text_area(
                "Texto del correo",
                prepared["email"],
                height=330,
                key="email_" + export_key,
            )

            st.download_button(
                "Descargar texto del correo",
                prepared["email"].encode("utf-8"),
                file_name="Correo_" + filename + ".txt",
                mime="text/plain",
            )

        st.divider()

        batch_key = hashlib.sha256(
            json.dumps(
                {
                    "model": model_hash,
                    "months": selected_months,
                    "source": uploaded.name,
                },
                sort_keys=True,
            ).encode("utf-8")
        ).hexdigest()

        if st.button(
            f"Preparar lote de {len(owners)} responsables"
        ):
            st.session_state.pop(
                "out_batch",
                None,
            )

            archive = io.BytesIO()
            progress = st.progress(0)

            try:
                with zipfile.ZipFile(
                    archive,
                    "w",
                    zipfile.ZIP_DEFLATED,
                ) as package:
                    for index, person in enumerate(
                        owners,
                        start=1,
                    ):
                        xlsx, email = report(
                            period_real,
                            period_budget,
                            person,
                            selected_months,
                            uploaded.name,
                            integrity_note,
                        )

                        filename = (
                            f"{index:02d}_"
                            f"{safe_name(person)}"
                        )

                        package.writestr(
                            f"Excel/{filename}.xlsx",
                            xlsx,
                        )

                        package.writestr(
                            f"Correos/{filename}.txt",
                            email.encode("utf-8"),
                        )

                        progress.progress(
                            index / len(owners)
                        )

                st.session_state["out_batch"] = {
                    "key": batch_key,
                    "data": archive.getvalue(),
                }

            except Exception as error:
                st.error(
                    "No se pudo completar el lote de reportes: "
                    f"{error}"
                )

            finally:
                progress.empty()

        prepared_batch = st.session_state.get(
            "out_batch",
            {},
        )

        if prepared_batch.get("key") == batch_key:
            st.download_button(
                "Descargar todos los reportes y correos",
                prepared_batch["data"],
                file_name=(
                    "Reportes_presupuestarios_"
                    f"{period_tag(selected_months)}.zip"
                ),
                mime="application/zip",
            )


# ============================================================
# CONTROL DE IMPORTACIÓN
# ============================================================

with controls:
    st.write(
        "Control de lectura del archivo completo, "
        "antes de aplicar los filtros del tablero."
    )

    show_summary(audit)

    st.caption(
        "Los importes se procesan a centavos. "
        "Las partidas reales se conservan con su signo. "
        "Las repeticiones presupuestarias solo se consolidan "
        "cuando se elige expresamente ese método."
    )

    type_column = actual_config["filter_column"]

    type_counts = (
        actual_table[type_column]
        .map(text)
        .value_counts(dropna=False)
        .rename_axis("Tipo")
        .reset_index(name="Filas")
    )

    selected_actual_types = {
        norm(value)
        for value in actual_config["filter_values"]
    }

    selected_budget_types = {
        norm(value)
        for value in budget_config["filter_values"]
    }

    def type_destination(value):
        normalized = norm(value)

        if normalized in selected_actual_types:
            return "Gasto"

        if normalized in selected_budget_types:
            return "Presupuesto"

        return "Fuera de los tipos seleccionados"

    type_counts["Destino"] = (
        type_counts["Tipo"].map(
            type_destination
        )
    )

    st.write(
        "Distribución de filas por Tipo"
    )

    st.dataframe(
        type_counts,
        use_container_width=True,
        hide_index=True,
    )

    st.caption(
        "Este conteo incluye las filas de la tabla antes "
        "de validar importes y campos. Los tipos no seleccionados "
        "no forman parte del gasto ni del presupuesto importado."
    )
