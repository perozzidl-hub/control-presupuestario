import hashlib
import io
import json
import re
import unicodedata
import zipfile
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path

import pandas as pd
import streamlit as st


st.set_page_config(
    page_title="Control presupuestario",
    page_icon="📊",
    layout="wide",
)

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

KEYS = ["Mes", "Responsable", "Cuenta", "Locación"]
EMPTY = "— Seleccionar —"

ALIASES = {
    "Mes": ["MES", "PERIODO", "FECHA", "FECHA CONTABLE"],
    "Responsable": [
        "RESPONSABLE",
        "RESPONSABLE PRESUPUESTO",
        "RESPONSABLE DE PRESUPUESTO",
    ],
    "Cuenta": ["CUENTA", "CUENTA CONTABLE", "CUENTA CONT", "CTA CBLE"],
    "Locación": ["LOCACION", "LOCALIDAD", "SUCURSAL"],
    "Descripción": [
        "DESC CTA CBLE",
        "CUENTA DES",
        "DESC CUENTA",
        "DESCRIPCION CUENTA",
    ],
    "Gasto": ["GASTO", "REAL", "IMPORTE", "SALDO", "NETO"],
    "Presupuesto": ["PRESUPUESTO", "PPTO", "PRESUP", "BUDGET"],
}

MONTHS = {
    "ENE": 1, "FEB": 2, "MAR": 3, "ABR": 4,
    "MAY": 5, "JUN": 6, "JUL": 7, "AGO": 8,
    "SEP": 9, "SET": 9, "OCT": 10, "NOV": 11, "DIC": 12,
}


def text(value):
    if value is None or pd.isna(value):
        return ""
    if isinstance(value, (int, float)) and float(value).is_integer():
        return str(int(value))
    return re.sub(r"\s+", " ", str(value)).strip()


def norm(value):
    value = unicodedata.normalize("NFKD", text(value))
    value = "".join(c for c in value if not unicodedata.combining(c))
    return re.sub(r"[_\s]+", " ", value).upper().strip()


def identifier(value):
    return text(value).upper()


def currency(value):
    if value is None or pd.isna(value):
        return "No comparable"
    formatted = f"{float(value):,.2f}"
    return "$ " + formatted.replace(",", "_").replace(".", ",").replace("_", ".")


def percentage(value):
    if value is None or pd.isna(value):
        return "No calculable"
    return f"{float(value):.1%}".replace(".", ",")


def cents(value):
    """Convierte números o importes de texto a centavos."""
    if value is None or pd.isna(value) or text(value) == "":
        return None

    if isinstance(value, bool):
        raise ValueError("Un importe contiene un valor lógico.")

    if isinstance(value, (int, float, Decimal)):
        number = Decimal(str(value))
    else:
        s = text(value).upper()
        s = s.replace("ARS", "").replace("$", "").replace(" ", "")

        if s in {"-", "–", "—"}:
            return 0

        if s.startswith("(") and s.endswith(")"):
            s = "-" + s[1:-1]

        if "," in s and "." in s:
            if s.rfind(",") > s.rfind("."):
                s = s.replace(".", "").replace(",", ".")
            else:
                s = s.replace(",", "")
        elif "," in s:
            s = s.replace(",", ".")
        elif re.fullmatch(r"-?\d{1,3}(?:\.\d{3})+", s):
            s = s.replace(".", "")

        try:
            number = Decimal(s)
        except InvalidOperation:
            raise ValueError(f"Importe no reconocido: {value!r}")

    if not number.is_finite():
        raise ValueError(f"Importe no finito: {value!r}")

    rounded = number.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    return int(rounded * 100)


def period(value, default_year):
    """Admite fechas Excel, jul-26, julio-26, 2026-07 y meses numéricos."""
    if value is None or pd.isna(value):
        return None

    if isinstance(value, (datetime, date, pd.Timestamp)):
        return pd.Timestamp(value).strftime("%Y-%m")

    if isinstance(value, (int, float)):
        number = float(value)
        if number.is_integer() and 1 <= number <= 12:
            return f"{default_year}-{int(number):02d}"
        if 20000 <= number <= 100000:
            timestamp = pd.Timestamp("1899-12-30") + pd.Timedelta(days=number)
            return timestamp.strftime("%Y-%m")
        return None

    s = norm(value)
    match = re.fullmatch(r"([A-Z]+)[\s/-]*(\d{2}|\d{4})?", s)
    if match and match.group(1)[:3] in MONTHS:
        month = MONTHS[match.group(1)[:3]]
        year = int(match.group(2)) if match.group(2) else default_year
        if year < 100:
            year += 2000
        return f"{year}-{month:02d}"

    if re.fullmatch(r"\d{1,2}", s):
        month = int(s)
        return f"{default_year}-{month:02d}" if 1 <= month <= 12 else None

    match = re.fullmatch(r"(\d{1,2})[/\-](\d{4})", s)
    if match:
        month, year = map(int, match.groups())
        return f"{year}-{month:02d}" if 1 <= month <= 12 else None

    if re.match(r"^\d{4}-", s):
        timestamp = pd.to_datetime(s, errors="coerce")
    else:
        timestamp = pd.to_datetime(s, dayfirst=True, errors="coerce")

    return None if pd.isna(timestamp) else timestamp.strftime("%Y-%m")


@st.cache_data(show_spinner=False, max_entries=4)
def read_table(content, sheet, header):
    frame = pd.read_excel(
        io.BytesIO(content),
        sheet_name=sheet,
        header=header - 1,
        dtype=object,
        engine="openpyxl",
    )
    frame.columns = [str(c).strip() for c in frame.columns]
    if frame.columns.duplicated().any():
        raise ValueError("Hay encabezados duplicados. Deben ser únicos.")
    return frame.dropna(how="all")


def choose(label, options, default, key):
    if key in st.session_state and st.session_state[key] not in options:
        del st.session_state[key]
    index = options.index(default) if default in options else 0
    return st.selectbox(label, options, index=index, key=key)


def source_config(content, sheets, role, saved):
    prefix = f"src_{role}_"
    sheet = choose(
        "Hoja de origen",
        sheets,
        saved.get("sheet", sheets[0]),
        prefix + "sheet",
    )
    header = st.number_input(
        "Fila que contiene los encabezados",
        min_value=1,
        max_value=100,
        value=int(saved.get("header", 1)),
        key=prefix + "header",
    )
    frame = read_table(content, sheet, int(header))
    st.dataframe(frame.head(6), use_container_width=True)

    options = [EMPTY] + list(frame.columns)
    mapping = {}
    columns = st.columns(2)

    for index, field in enumerate(
        ["Mes", "Responsable", "Cuenta", role, "Locación", "Descripción"]
    ):
        candidates = ALIASES[field]
        detected = next(
            (column for column in frame.columns if norm(column) in candidates),
            EMPTY,
        )
        default = saved.get("mapping", {}).get(field) or detected
        with columns[index % 2]:
            selected = choose(
                field + (" — opcional" if field in ["Locación", "Descripción"] else ""),
                options,
                default,
                prefix + field,
            )
        mapping[field] = None if selected == EMPTY else selected

    year = st.number_input(
        "Año para meses que no incluyen año",
        min_value=2000,
        max_value=2100,
        value=int(saved.get("year", date.today().year)),
        key=prefix + "year",
    )

    filter_column = choose(
        "Filtrar filas por una columna — opcional",
        options,
        saved.get("filter_column") or EMPTY,
        prefix + "filter_column",
    )
    filter_values = []

    if filter_column != EMPTY:
        available = sorted(frame[filter_column].map(text).unique().tolist())
        default_values = [
            value for value in saved.get("filter_values", [])
            if value in available
        ]
        filter_key = prefix + "filter_values"
        if filter_key in st.session_state:
            st.session_state[filter_key] = [
                value for value in st.session_state[filter_key]
                if value in available
            ]
        filter_values = st.multiselect(
            "Valores que pertenecen a esta fuente",
            available,
            default=default_values,
            key=filter_key,
        )

    skip_totals = st.checkbox(
        "Excluir filas rotuladas como Total o Subtotal",
        value=bool(saved.get("skip_totals", True)),
        key=prefix + "skip_totals",
    )
    invert = st.checkbox(
        "Invertir el signo de los importes de esta fuente",
        value=bool(saved.get("invert", False)),
        key=prefix + "invert",
    )

    mode = None
    if role == "Presupuesto":
        modes = [
            "Elegir método",
            "Sumar partidas presupuestarias",
            "El mismo presupuesto está repetido en los movimientos",
        ]
        mode = choose(
            "Cómo está registrado el presupuesto",
            modes,
            saved.get("mode", modes[0]),
            prefix + "mode",
        )
        st.caption(
            "La segunda opción conserva un importe por mes, responsable, "
            "cuenta y locación. Si encuentra importes diferentes para esa "
            "misma combinación, detiene la carga para revisión."
        )

    configuration = {
        "sheet": sheet,
        "header": int(header),
        "mapping": mapping,
        "year": int(year),
        "filter_column": None if filter_column == EMPTY else filter_column,
        "filter_values": filter_values,
        "skip_totals": skip_totals,
        "invert": invert,
        "mode": mode,
    }
    return frame, configuration


def normalize_source(frame, config, role):
    mapping = config["mapping"]

    for field in ["Mes", "Responsable", "Cuenta", role]:
        if not mapping.get(field):
            raise ValueError(f"{role}: falta seleccionar la columna {field}.")

    raw = frame.copy()
    removed_totals = 0

    if config["filter_column"]:
        if not config["filter_values"]:
            raise ValueError(f"{role}: seleccioná los valores del filtro de filas.")
        raw = raw[
            raw[config["filter_column"]].map(text).isin(config["filter_values"])
        ]

    if config["skip_totals"]:
        total_mask = pd.Series(False, index=raw.index)
        for field in ["Cuenta", "Descripción", "Responsable"]:
            column = mapping.get(field)
            if column:
                total_mask |= raw[column].map(norm).str.match(
                    r"^(?:TOTAL|SUBTOTAL)"
                )
        removed_totals = int(total_mask.sum())
        raw = raw[~total_mask]

    parsed = []
    for index, value in raw[mapping[role]].items():
        try:
            parsed.append(cents(value))
        except ValueError as error:
            row = int(index) + config["header"] + 1
            raise ValueError(f"{role}, fila Excel {row}: {error}") from error

    amounts = pd.Series(parsed, index=raw.index, dtype=object)
    raw = raw.loc[amounts.notna()]
    amounts = amounts.loc[raw.index]

    output = pd.DataFrame(index=raw.index)
    output["Mes"] = raw[mapping["Mes"]].map(
        lambda value: period(value, config["year"])
    )
    output["Responsable"] = raw[mapping["Responsable"]].map(identifier)
    output["Cuenta"] = raw[mapping["Cuenta"]].map(identifier)

    location = mapping.get("Locación")
    description = mapping.get("Descripción")

    output["Locación"] = (
        raw[location].map(identifier).replace("", "SIN LOCACIÓN")
        if location else "SIN DESGLOSE"
    )
    output["Descripción"] = raw[description].map(text) if description else ""
    output["_centavos"] = amounts.astype("int64") * (-1 if config["invert"] else 1)
    output["Fila Excel"] = raw.index + config["header"] + 1
    output["Hoja origen"] = config["sheet"]

    invalid = (
        output["Mes"].isna()
        | output["Cuenta"].eq("")
        | output["Responsable"].eq("")
    )
    if invalid.any():
        rows = output.loc[invalid, "Fila Excel"].head(12).astype(str).tolist()
        raise ValueError(
            f"{role}: hay filas con importe pero sin mes, cuenta o responsable "
            f"válido. Revisá las filas Excel: {', '.join(rows)}. "
            "Seleccioná una tabla de datos, sin celdas agrupadas de tabla dinámica."
        )

    original = raw.rename(columns=lambda column: "Origen · " + column)
    output = pd.concat([output, original], axis=1).reset_index(drop=True)

    input_count = len(output)
    repeated = 0

    if role == "Presupuesto":
        if config["mode"] == "Elegir método":
            raise ValueError("Indicá cómo está registrado el presupuesto.")

        if config["mode"] == "El mismo presupuesto está repetido en los movimientos":
            distinct = output.groupby(KEYS)["_centavos"].nunique()
            conflicts = distinct[distinct > 1]
            if not conflicts.empty:
                raise ValueError(
                    "Hay presupuestos diferentes para una misma combinación "
                    "de mes, responsable, cuenta y locación. Revisá el origen "
                    "o elegí sumar partidas si son asignaciones independientes."
                )
            output = output.drop_duplicates(KEYS)
            repeated = input_count - len(output)

    audit = {
        "Fuente": role,
        "Filas con importe antes de consolidar": input_count,
        "Filas de total excluidas": removed_totals,
        "Repeticiones presupuestarias excluidas": repeated,
        "Filas utilizadas": len(output),
        "Importe": output["_centavos"].sum() / 100,
    }
    return output, audit


def comparison(real, budget, dimensions, comparable=True):
    actual = (
        real.groupby(dimensions, dropna=False)["_centavos"]
        .sum().rename("Gasto").reset_index()
    )
    planned = (
        budget.groupby(dimensions, dropna=False)["_centavos"]
        .sum().rename("Presupuesto").reset_index()
    )

    if not comparable:
        result = actual.copy()
        result["Gasto"] = result["Gasto"] / 100
        for column in ["Presupuesto", "Desvío $", "Desvío %"]:
            result[column] = float("nan")
        result["Estado"] = "Presupuesto sin desglose comparable"
        return result

    result = actual.merge(planned, on=dimensions, how="outer")
    result[["Gasto", "Presupuesto"]] = (
        result[["Gasto", "Presupuesto"]].fillna(0) / 100
    )
    result["Desvío $"] = result["Gasto"] - result["Presupuesto"]
    result["Desvío %"] = (
        result["Desvío $"]
        / result["Presupuesto"].where(result["Presupuesto"] > 0)
    )
    result["Estado"] = "Dentro del presupuesto"
    result.loc[result["Desvío $"] > 0, "Estado"] = "Exceso de gasto"
    result.loc[result["Presupuesto"] == 0, "Estado"] = "Sin presupuesto"
    result.loc[result["Presupuesto"] < 0, "Estado"] = "Presupuesto negativo: revisar"
    return result.sort_values(dimensions).reset_index(drop=True)


def detail_export(frame, amount_name):
    output = frame.drop(columns="_centavos").copy()
    output.insert(0, amount_name, frame["_centavos"].to_numpy() / 100)
    return output


def report(real, budget, owner, months):
    real = real[real["Responsable"].eq(owner)].copy()
    budget = budget[budget["Responsable"].eq(owner)].copy()

    accounts = comparison(real, budget, ["Cuenta"])
    monthly = comparison(real, budget, ["Mes"])

    descriptions = pd.concat(
        [real[["Cuenta", "Descripción"]], budget[["Cuenta", "Descripción"]]]
    )
    descriptions = descriptions[
        descriptions["Descripción"].ne("")
    ].drop_duplicates("Cuenta")
    accounts = accounts.merge(descriptions, on="Cuenta", how="left")

    spent = real["_centavos"].sum() / 100
    planned = budget["_centavos"].sum() / 100
    variance = spent - planned
    ratio = variance / planned if planned > 0 else None

    summary = pd.DataFrame({
        "Indicador": [
            "Responsable",
            "Períodos incluidos",
            "Presupuesto",
            "Gasto",
            "Desvío $",
            "Desvío %",
            "Criterio",
        ],
        "Valor": [
            owner,
            ", ".join(months),
            planned,
            spent,
            variance,
            ratio,
            "Desvío = gasto − presupuesto. Positivo indica exceso.",
        ],
    })

    output = io.BytesIO()
    with pd.ExcelWriter(
        output,
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
        number_format = workbook.add_format({"num_format": '#,##0.00;[Red]-#,##0.00'})
        percent_format = workbook.add_format({"num_format": '0.0%;[Red]-0.0%'})
        excess_format = workbook.add_format({
            "bg_color": "#FCE4E4",
            "font_color": "#9C2020",
        })

        tables = {
            "Resumen": summary,
            "Cuentas": accounts,
            "Evolucion": monthly,
            "Detalle contable": detail_export(real, "Gasto"),
            "Base presupuesto": detail_export(budget, "Presupuesto"),
        }

        for name, table in tables.items():
            table.to_excel(writer, sheet_name=name, index=False)
            worksheet = writer.sheets[name]
            worksheet.freeze_panes(1, 0)
            worksheet.set_row(0, 32)

            for index, column in enumerate(table.columns):
                worksheet.write(0, index, column, header_format)
                width = min(48, max(16, len(str(column)) + 3))
                cell_format = None
                if column in ["Gasto", "Presupuesto", "Desvío $"]:
                    cell_format = number_format
                elif column == "Desvío %":
                    cell_format = percent_format
                worksheet.set_column(index, index, width, cell_format)

            if not table.empty:
                worksheet.autofilter(0, 0, len(table), len(table.columns) - 1)

            if "Desvío $" in table.columns and not table.empty:
                column = table.columns.get_loc("Desvío $")
                worksheet.conditional_format(
                    1, column, len(table), column,
                    {
                        "type": "cell",
                        "criteria": ">",
                        "value": 0,
                        "format": excess_format,
                    },
                )

        summary_sheet = writer.sheets["Resumen"]
        summary_sheet.set_column("A:A", 24)
        summary_sheet.set_column("B:B", 65)
        for row in [3, 4, 5]:
            summary_sheet.write_number(row, 1, float(summary.iloc[row - 1]["Valor"]), number_format)
        if ratio is not None:
            summary_sheet.write_number(6, 1, float(ratio), percent_format)

        if not monthly.empty:
            chart = workbook.add_chart({"type": "column"})
            for column_name, color in [
                ("Presupuesto", "#8BA6C2"),
                ("Gasto", "#244D78"),
            ]:
                column_index = monthly.columns.get_loc(column_name)
                chart.add_series({
                    "name": column_name,
                    "categories": ["Evolucion", 1, 0, len(monthly), 0],
                    "values": [
                        "Evolucion", 1, column_index, len(monthly), column_index
                    ],
                    "fill": {"color": color},
                })
            chart.set_title({"name": "Presupuesto y gasto"})
            chart.set_legend({"position": "bottom"})
            summary_sheet.insert_chart("D2", chart)

    relevant = accounts[
        (accounts["Presupuesto"] == 0) & (accounts["Gasto"] != 0)
    ]
    note = (
        f"\nHay {len(relevant)} cuenta(s) con movimientos sin presupuesto asignado.\n"
        if not relevant.empty else ""
    )

    email = (
        f"Asunto: Control presupuestario | {', '.join(months)} | {owner}\n\n"
        "Buen día,\n\n"
        "Adjunto el reporte presupuestario del período indicado, "
        "con el resumen por cuenta y el detalle de movimientos.\n\n"
        f"Presupuesto: {currency(planned)}\n"
        f"Gasto real: {currency(spent)}\n"
        f"Desvío: {currency(variance)}\n"
        f"Desvío porcentual: {percentage(ratio)}\n"
        f"{note}\n"
        "Un desvío positivo indica gasto por encima del presupuesto. "
        "Un desvío negativo indica una menor ejecución presupuestaria.\n\n"
        "Por favor, revisen los movimientos e informen las explicaciones "
        "o ajustes de imputación que correspondan.\n\n"
        "Saludos,\nDamián"
    )
    return output.getvalue(), email


def safe_name(value):
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", value).strip(". ")[:100] or "reporte"


def reset_filters():
    for key in ["f_Responsable", "f_Cuenta", "f_Locación"]:
        st.session_state[key] = []
    st.session_state["f_Mes"] = [st.session_state["_latest_month"]]


st.title("Control presupuestario")
st.caption("Vista general · Análisis por responsable · Reportes mensuales")

uploaded = st.file_uploader("Cargar archivo base", type=["xlsx"], key="upload")

if uploaded is None:
    st.info("Cargá el Excel que contiene la hoja madre y los presupuestos.")
    st.stop()

content = uploaded.getvalue()
file_hash = hashlib.sha256(content).hexdigest()

if st.session_state.get("_file_hash") != file_hash:
    for key in list(st.session_state):
        if key.startswith(("src_", "f_", "out_")):
            del st.session_state[key]
    st.session_state["_file_hash"] = file_hash

try:
    with pd.ExcelFile(io.BytesIO(content), engine="openpyxl") as excel:
        sheets = excel.sheet_names
except Exception as error:
    st.error(f"No se pudo abrir el Excel: {error}")
    st.stop()

config_path = Path(__file__).with_name("configuracion.json")
saved_config = st.session_state.get("_active_config", {})
if not saved_config and config_path.exists():
    try:
        saved_config = json.loads(config_path.read_text(encoding="utf-8"))
    except (ValueError, OSError):
        st.warning("No se pudo leer configuracion.json. Configurá las fuentes.")

with st.expander("Configuración de lectura", expanded=not bool(saved_config)):
    st.write(
        "Elegí la hoja madre para el gasto real. El presupuesto puede estar "
        "en esa misma hoja o en otra del mismo archivo."
    )
    st.caption(
        "La agrupación elegida como Responsable determinará los archivos "
        "individuales. Usá una columna que corresponda a tus destinatarios."
    )

    actual_tab, budget_tab = st.tabs(["Gasto real", "Presupuesto"])
    try:
        with actual_tab:
            actual_table, actual_config = source_config(
                content, sheets, "Gasto", saved_config.get("Gasto", {})
            )
        with budget_tab:
            budget_table, budget_config = source_config(
                content, sheets, "Presupuesto", saved_config.get("Presupuesto", {})
            )
    except Exception as error:
        st.error(f"Revisá la hoja y la fila de encabezados: {error}")
        st.stop()

configuration = {"Gasto": actual_config, "Presupuesto": budget_config}
st.session_state["_active_config"] = configuration

model_hash = hashlib.sha256(
    content + json.dumps(configuration, sort_keys=True).encode("utf-8")
).hexdigest()

try:
    if st.session_state.get("_model_hash") != model_hash:
        real, real_audit = normalize_source(actual_table, actual_config, "Gasto")
        budget, budget_audit = normalize_source(
            budget_table, budget_config, "Presupuesto"
        )

        same_amount_source = (
            actual_config["sheet"] == budget_config["sheet"]
            and actual_config["mapping"]["Gasto"]
            == budget_config["mapping"]["Presupuesto"]
        )
        if same_amount_source:
            overlap = set(real["Fila Excel"]) & set(budget["Fila Excel"])
            if overlap:
                raise ValueError(
                    "Las mismas celdas están siendo usadas como gasto y presupuesto. "
                    "Elegí columnas distintas o filtros de filas que separen ambos."
                )

        if real.empty:
            raise ValueError("No se encontraron movimientos reales con la configuración elegida.")

        st.session_state["_real"] = real
        st.session_state["_budget"] = budget
        st.session_state["_audit"] = pd.DataFrame([real_audit, budget_audit])
        st.session_state["_model_hash"] = model_hash
except Exception as error:
    st.error(str(error))
    st.stop()

real = st.session_state["_real"]
budget = st.session_state["_budget"]

st.download_button(
    "Guardar configuración de lectura",
    json.dumps(configuration, ensure_ascii=False, indent=2).encode("utf-8"),
    file_name="configuracion.json",
    mime="application/json",
)
st.caption(
    "Para reutilizarla en futuras sesiones, colocá configuracion.json junto "
    "a app.py. Los archivos mensuales conservarán sus nombres y estructura."
)

months_available = sorted(real["Mes"].unique().tolist())
st.session_state["_latest_month"] = months_available[-1]

with st.sidebar:
    st.title("Filtros")
    st.button("Limpiar filtros", on_click=reset_filters)

    if "f_Mes" in st.session_state:
        st.session_state["f_Mes"] = [
            value for value in st.session_state["f_Mes"]
            if value in months_available
        ]

    months = st.multiselect(
        "Mes",
        months_available,
        default=[months_available[-1]],
        key="f_Mes",
    )
    selections = {}
    for dimension in ["Responsable", "Cuenta", "Locación"]:
        options = sorted(
            set(real[dimension].tolist()) | set(budget[dimension].tolist())
        )
        key = "f_" + dimension
        if key in st.session_state:
            st.session_state[key] = [
                value for value in st.session_state[key] if value in options
            ]
        selections[dimension] = st.multiselect(dimension, options, key=key)

if not months:
    st.info("Seleccioná al menos un mes.")
    st.stop()

period_real = real[real["Mes"].isin(months)].copy()
period_budget = budget[budget["Mes"].isin(months)].copy()
view_real, view_budget = period_real.copy(), period_budget.copy()

for dimension in ["Responsable", "Cuenta"]:
    selected = selections[dimension]
    if selected:
        view_real = view_real[view_real[dimension].isin(selected)]
        view_budget = view_budget[view_budget[dimension].isin(selected)]

location_filter = selections["Locación"]
comparable = True
if location_filter:
    comparable = bool(
        actual_config["mapping"]["Locación"]
        and budget_config["mapping"]["Locación"]
    )
    view_real = view_real[view_real["Locación"].isin(location_filter)]
    if comparable:
        view_budget = view_budget[view_budget["Locación"].isin(location_filter)]
    else:
        st.info(
            "Se muestra el gasto de las locaciones elegidas. "
            "El presupuesto no tiene un desglose comparable por locación, "
            "por lo que se omiten sus desvíos."
        )

spent = view_real["_centavos"].sum() / 100
planned = view_budget["_centavos"].sum() / 100 if comparable else None
variance = spent - planned if comparable else None
ratio = variance / planned if comparable and planned > 0 else None

kpis = st.columns(4)
kpis[0].metric("Presupuesto", currency(planned))
kpis[1].metric("Gasto real", currency(spent))
kpis[2].metric("Desvío $", currency(variance))
kpis[3].metric("Desvío %", percentage(ratio))

overview, account_tab, movements, exports, controls = st.tabs([
    "Panorama",
    "Cuentas",
    "Movimientos",
    "Descargas",
    "Control de importación",
])

with overview:
    by_owner = comparison(
        view_real, view_budget, ["Responsable"], comparable
    )
    st.subheader("Presupuesto y gasto por responsable")
    chart_columns = ["Presupuesto", "Gasto"] if comparable else ["Gasto"]
    if not by_owner.empty:
        chart = (
            by_owner.sort_values("Gasto", ascending=False)
            .head(20).set_index("Responsable")[chart_columns]
        )
        st.bar_chart(chart)
    st.dataframe(by_owner, use_container_width=True, hide_index=True)

    st.subheader("Comparación mensual del período seleccionado")
    by_month = comparison(view_real, view_budget, ["Mes"], comparable)
    if not by_month.empty:
        st.line_chart(by_month.set_index("Mes")[chart_columns])

with account_tab:
    by_account = comparison(
        view_real, view_budget, ["Responsable", "Cuenta"], comparable
    )
    descriptions = pd.concat([
        real[["Cuenta", "Descripción"]],
        budget[["Cuenta", "Descripción"]],
    ])
    descriptions = descriptions[
        descriptions["Descripción"].ne("")
    ].drop_duplicates("Cuenta")
    by_account = by_account.merge(descriptions, on="Cuenta", how="left")
    st.dataframe(by_account, use_container_width=True, hide_index=True)

with movements:
    st.caption("Importes provenientes de la fuente seleccionada como hoja madre.")
    st.dataframe(
        detail_export(view_real, "Gasto"),
        use_container_width=True,
        hide_index=True,
    )

with exports:
    st.write(
        "Los reportes incluyen todas las cuentas de cada responsable "
        "en los meses seleccionados."
    )
    st.caption(
        "Los filtros de cuenta, locación y responsable del tablero "
        "no recortan estas descargas."
    )

    owners = sorted(
        set(period_real["Responsable"]) | set(period_budget["Responsable"])
    )
    owner = st.selectbox("Responsable a exportar", owners)

    export_key = hashlib.sha256(
        f"{model_hash}|{sorted(months)}|{owner}".encode("utf-8")
    ).hexdigest()

    if st.button("Preparar reporte individual", type="primary"):
        with st.spinner("Preparando Excel y texto del correo..."):
            xlsx, email = report(period_real, period_budget, owner, sorted(months))
            st.session_state["out_individual"] = {
                "key": export_key,
                "xlsx": xlsx,
                "email": email,
            }

    prepared = st.session_state.get("out_individual", {})
    if prepared.get("key") == export_key:
        st.download_button(
            "Descargar Excel",
            prepared["xlsx"],
            file_name=f"{safe_name(owner)}_{max(months)}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.text_area("Texto del correo", prepared["email"], height=330)
        st.download_button(
            "Descargar texto del correo",
            prepared["email"].encode("utf-8"),
            file_name=f"Correo_{safe_name(owner)}.txt",
            mime="text/plain",
        )

    st.divider()
    batch_key = model_hash + "|" + ",".join(sorted(months))

    if st.button(f"Preparar lote de {len(owners)} responsables"):
        archive = io.BytesIO()
        progress = st.progress(0)
        with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED) as package:
            for index, person in enumerate(owners, start=1):
                xlsx, email = report(
                    period_real, period_budget, person, sorted(months)
                )
                filename = f"{index:02d}_{safe_name(person)}"
                package.writestr(f"Excel/{filename}.xlsx", xlsx)
                package.writestr(
                    f"Correos/{filename}.txt", email.encode("utf-8")
                )
                progress.progress(index / len(owners))
        st.session_state["out_batch"] = {
            "key": batch_key,
            "data": archive.getvalue(),
        }

    prepared_batch = st.session_state.get("out_batch", {})
    if prepared_batch.get("key") == batch_key:
        st.download_button(
            "Descargar todos los reportes y correos",
            prepared_batch["data"],
            file_name=f"Reportes_presupuestarios_{max(months)}.zip",
            mime="application/zip",
        )

with controls:
    st.write("Control de lectura del archivo completo, antes de aplicar filtros.")
    st.dataframe(
        st.session_state["_audit"],
        use_container_width=True,
        hide_index=True,
    )
    st.caption(
        "Los importes se procesan a centavos. Las partidas reales repetidas "
        "no se eliminan automáticamente. Este control de lectura debe "
        "contrastarse con los totales del cierre original."
    )
