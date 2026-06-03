"""
Servidor Flask – Sistema de Liquidación de Nómina
"""

import io
import json
import os
import sqlite3
from difflib import get_close_matches
from datetime import datetime, date
from pathlib import Path

from flask import Flask, jsonify, request, send_file, render_template
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from engine import TURNO_CATALOGO, calcular_nomina

app = Flask(__name__)

BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"

# ─────────────────────────────────────────────────────────────
# Base de datos SQLite para persistencia de configuración
# ─────────────────────────────────────────────────────────────

# En producción (Render) se puede montar un disco en /data.
# En local sigue usando la carpeta data/ del proyecto.
DB_PATH = Path(os.environ.get("DB_PATH", str(DATA_DIR / "nomina.db")))


def _get_db():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    return conn


def _init_db():
    with _get_db() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS config (
                key   TEXT PRIMARY KEY,
                value TEXT NOT NULL
            )
        """)
        conn.commit()


def _db_get(key):
    with _get_db() as conn:
        row = conn.execute("SELECT value FROM config WHERE key = ?", (key,)).fetchone()
        return json.loads(row["value"]) if row else None


def _db_set(key, value):
    with _get_db() as conn:
        conn.execute(
            "INSERT INTO config(key, value) VALUES(?, ?) "
            "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
            (key, json.dumps(value, ensure_ascii=False)),
        )
        conn.commit()


# ─────────────────────────────────────────────────────────────
# Utilidades
# ─────────────────────────────────────────────────────────────

def _load_festivos():
    fp = DATA_DIR / "festivos.json"
    if fp.exists():
        return json.loads(fp.read_text(encoding="utf-8"))
    return []


def _load_config():
    """Configuración persistente desde SQLite."""
    return _db_get("user_config") or {}


def _save_config(cfg):
    _db_set("user_config", cfg)


def _load_current():
    data = _load_defaults()
    cfg = _load_config()
    if cfg:
        if isinstance(cfg.get("params"), dict):
            data["params"].update(cfg["params"])
        if isinstance(cfg.get("reglas"), dict):
            data["reglas"].update(cfg["reglas"])
        if isinstance(cfg.get("festivos"), list):
            data["festivos"] = cfg["festivos"]
    return data


def _load_defaults():
    return {
        "params": {
            "anio": 2026,
            "mes": 4,
            "dias_ciclo": 21,
            "horas_objetivo": 132,
            "horas_turno": 12.583333,
            "smmlv": 1423500,
            "auxilio_transporte_mensual": 200000,
            "inicio_noc_h": 19,
            "fin_noc_h": 6,
        },
        "reglas": {
            "REC_DIURNO": 0.0,
            "REC_NOCTURNO": 0.35,
            "REC_DOM_FEST_DIURNO": 0.80,
            "REC_DOM_FEST_NOCTURNO": 1.15,
            "EXT_DIURNA": 1.25,
            "EXT_NOCTURNA": 1.75,
            "EXT_FEST_DIURNA": 2.05,
            "EXT_FEST_NOCTURNA": 2.55,
        },
        "festivos": _load_festivos(),
        "empleados": [],
    }


def _fecha_dmy(iso):
    s = str(iso or "")[:10]
    p = s.split("-")
    return f"{p[2]}/{p[1]}/{p[0]}" if len(p) == 3 else s


def _parse_fecha(val):
    if val is None:
        return ""
    if isinstance(val, datetime):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, date):
        return val.strftime("%Y-%m-%d")
    s = str(val).strip()
    if not s:
        return ""
    s = s.split(" ")[0]
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%Y/%m/%d", "%d/%m/%y", "%d-%m-%y", "%m/%d/%Y"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except Exception:
            pass
    return s[:10]


TURNOS_VALIDOS = set(TURNO_CATALOGO.keys())


def _normalizar_codigo_turno(val):
    if val is None:
        return "", None
    raw = str(val).strip()
    if not raw:
        return "", None
    norm = "".join(raw.upper().split())
    if norm in TURNOS_VALIDOS:
        issue = None
        if norm != raw:
            issue = {
                "type": "warning",
                "message": f"Se normalizó '{raw}' a '{norm}'.",
                "suggestion": norm,
            }
        return norm, issue
    suggestion = get_close_matches(norm, list(TURNOS_VALIDOS), n=1, cutoff=0.75)
    issue = {
        "type": "error",
        "message": f"Turno '{raw}' no reconocido.",
    }
    if suggestion:
        issue["suggestion"] = suggestion[0]
        issue["message"] += f" Sugerencia: '{suggestion[0]}'."
    return norm, issue


def _validar_turnos_empleados(empleados):
    cleaned = []
    warnings = []
    errors = []
    empleados_afectados = set()
    total_celdas = 0
    normalizadas = 0

    for emp in empleados:
        emp_clean = dict(emp)
        turnos = {}
        for d, val in (emp.get("turnos") or {}).items():
            total_celdas += 1
            norm, issue = _normalizar_codigo_turno(val)
            if norm:
                turnos[str(d)] = norm
            if issue:
                base = {
                    "employee_id": emp.get("id"),
                    "employee_name": emp.get("nombre", ""),
                    "row": emp.get("_source_row"),
                    "day": int(d) if str(d).isdigit() else d,
                    "value": val,
                    "normalized": norm,
                }
                base.update(issue)
                if issue["type"] == "warning":
                    warnings.append(base)
                    normalizadas += 1
                else:
                    errors.append(base)
                empleados_afectados.add(emp.get("id"))
        emp_clean["turnos"] = turnos
        cleaned.append(emp_clean)

    validation = {
        "ok": len(errors) == 0,
        "warnings": warnings,
        "errors": errors,
        "summary": {
            "empleados": len(empleados),
            "empleados_afectados": len([x for x in empleados_afectados if x is not None]),
            "celdas_turno": total_celdas,
            "advertencias": len(warnings),
            "errores": len(errors),
            "normalizadas": normalizadas,
        },
    }
    return cleaned, validation


def _excel_to_data(wb):
    data = _load_current()
    sheet_name = "PLANTILLA_CARGA" if "PLANTILLA_CARGA" in wb.sheetnames else "MATRIZ_CARGA"
    turno_col_offset = 7 if sheet_name == "PLANTILLA_CARGA" else 9
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        header_row = None
        empleados = []
        for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
            if row[0] == "ID":
                header_row = i
                continue
            if header_row and row[0] and str(row[0]).isdigit() or (header_row and isinstance(row[0], (int, float)) and row[0]):
                emp_id = row[0]
                if not isinstance(emp_id, (int, float)):
                    continue
                nombre = str(row[1] or "").strip()
                salario = float(row[2] or 0)
                fic = _parse_fecha(row[3])
                saldo = float(row[4] or 0)
                dias_trab = None
                aux_manual = None
                if sheet_name == "PLANTILLA_CARGA":
                    dias_trab = row[5] if len(row) > 5 and row[5] else None
                    aux_manual = row[6] if len(row) > 6 and row[6] else None
                turnos = {}
                for d in range(1, 32):
                    val = row[turno_col_offset + d - 1] if len(row) > turno_col_offset + d - 1 else None
                    if val:
                        turnos[str(d)] = str(val).strip()
                emp_dict = {
                    "id": int(emp_id),
                    "nombre": nombre,
                    "salario_mensual": salario,
                    "fecha_inicio_ciclo": fic,
                    "saldo_inicial_horas": saldo,
                    "turnos": turnos,
                    "_source_row": i,
                }
                if dias_trab is not None:
                    try: emp_dict["dias_trabajados_manual"] = int(dias_trab)
                    except: pass
                if aux_manual is not None:
                    try: emp_dict["auxilio_transporte_manual"] = float(aux_manual)
                    except: pass
                empleados.append(emp_dict)
        data["empleados"], data["validation"] = _validar_turnos_empleados(empleados)
    return data


def _generar_excel(resultados, params, reglas=None):
    reglas = reglas or {}
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "LIQUIDACION"

    CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
    RIGHT  = Alignment(horizontal="right",  vertical="center")
    thin   = Side(style="thin", color="BFBFBF")
    BDR    = Border(left=thin, right=thin, top=thin, bottom=thin)

    C_TITLE = "1F4E79"; C_REC = "1d6fa0"; C_EXT = "2585c0"; C_MON = "1F4E79"
    C_TOT = "c6efce"; C_ALT = "f7fafe"

    def hdr(fg, sz=9, color="FFFFFF"):
        return Font(bold=True, color=color, size=sz), PatternFill("solid", fgColor=fg)

    def pc(k):
        return round(float(reglas.get(k, 0)) * 100)

    mes_nombre = ["", "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
                  "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"][int(params["mes"])]

    cols = [
        ("ID", C_TITLE), ("EMPLEADO", C_TITLE), ("SALARIO BASE", C_TITLE),
        (f"REC. NOCTURNO {pc('REC_NOCTURNO')}%", C_REC),
        (f"REC. DOM/FEST {pc('REC_DOM_FEST_DIURNO')}%", C_REC),
        (f"REC. DOM/FEST NOCT {pc('REC_DOM_FEST_NOCTURNO')}%", C_REC),
        ("TOTAL RECARGOS", C_REC),
        (f"H.E DIURNA {pc('EXT_DIURNA')}%", C_EXT),
        (f"H.E NOCTURNA {pc('EXT_NOCTURNA')}%", C_EXT),
        (f"H.E FEST DIURNA {pc('EXT_FEST_DIURNA')}%", C_EXT),
        (f"H.E FEST NOCT {pc('EXT_FEST_NOCTURNA')}%", C_EXT),
        ("TOTAL H.E", C_EXT),
        ("AUXILIO TRANSP.", C_MON),
        ("TOTAL A PAGAR", C_MON),
    ]
    ncol = len(cols)
    last_col = get_column_letter(ncol)

    ws.merge_cells(f"A1:{last_col}1")
    ws["A1"] = f"LIQUIDACION DE NOMINA – {mes_nombre} {params['anio']} · PERSONAL DE SEGURIDAD"
    f, fill = hdr(C_TITLE, sz=12); ws["A1"].font = f; ws["A1"].fill = fill; ws["A1"].alignment = CENTER
    ws.row_dimensions[1].height = 26

    ws.merge_cells(f"A2:{last_col}2")
    ws["A2"] = (f"Ciclo: {params['dias_ciclo']} días  |  Horas objetivo: {params['horas_objetivo']}h  |  "
                f"IVH = Salario ÷ {params.get('horas_mes', 220)}  |  "
                f"Nocturno: {params.get('inicio_noc_h', 19)}:00–{params.get('fin_noc_h', 6)}:00")
    f2, fill2 = hdr("2E75B6", sz=8); ws["A2"].font = f2; ws["A2"].fill = fill2; ws["A2"].alignment = CENTER
    ws.row_dimensions[2].height = 16

    ws.row_dimensions[3].height = 30
    for c, (lbl, fg) in enumerate(cols, 1):
        cell = ws.cell(row=3, column=c, value=lbl)
        f, fill = hdr(fg); cell.font = f; cell.fill = fill; cell.alignment = CENTER; cell.border = BDR

    FMT_NUM = "#,##0"
    FILL_ALT = PatternFill("solid", fgColor=C_ALT)
    tot = [0] * ncol
    start_row = 4

    for i, r in enumerate(resultados, start=start_row):
        v = r.get("val", {})
        row_data = [
            r["id"], r["nombre"], round(r["salario_mensual"]),
            v.get("rec_noct", 0), v.get("rec_fest_d", 0), v.get("rec_fest_n", 0), r["valor_recargo"],
            v.get("ext_diurna", 0), v.get("ext_noct", 0), v.get("ext_fest_d", 0), v.get("ext_fest_n", 0), r["valor_extra"],
            r["auxilio_transporte"], r["total_pagar"],
        ]
        is_even = (i % 2 == 0)
        for col, val in enumerate(row_data, 1):
            cell = ws.cell(row=i, column=col, value=val)
            cell.border = BDR
            if is_even and col > 1:
                cell.fill = FILL_ALT
            if col == 1:
                cell.alignment = CENTER
            elif col == 2:
                cell.alignment = Alignment(vertical="center")
            else:
                cell.alignment = RIGHT; cell.number_format = FMT_NUM
                if col in (7, 12):
                    cell.font = Font(bold=True, size=9)
                if col == ncol:
                    cell.font = Font(bold=True, color="C00000", size=9)
            if col >= 3 and isinstance(val, (int, float)):
                tot[col - 1] += val

    tot_row = start_row + len(resultados)
    ws.row_dimensions[tot_row].height = 18
    ws.merge_cells(f"A{tot_row}:B{tot_row}")
    cell = ws.cell(row=tot_row, column=1, value=f"TOTALES ({len(resultados)} empleados)")
    cell.font = Font(bold=True, size=9); cell.fill = PatternFill("solid", fgColor=C_TOT)
    cell.alignment = CENTER; cell.border = BDR
    b2 = ws.cell(row=tot_row, column=2)
    b2.fill = PatternFill("solid", fgColor=C_TOT); b2.border = BDR
    for col in range(3, ncol + 1):
        cell = ws.cell(row=tot_row, column=col, value=tot[col - 1])
        cell.font = Font(bold=True, size=9, color=("C00000" if col == ncol else C_TITLE))
        cell.fill = PatternFill("solid", fgColor=C_TOT)
        cell.number_format = FMT_NUM; cell.alignment = RIGHT; cell.border = BDR

    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 30
    for col in range(3, ncol + 1):
        ws.column_dimensions[get_column_letter(col)].width = 16
    ws.freeze_panes = "C4"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def _generar_excel_detalle(resultados, params):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "DETALLE"

    CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
    RIGHT  = Alignment(horizontal="right",  vertical="center")
    thin   = Side(style="thin", color="BFBFBF")
    BDR    = Border(left=thin, right=thin, top=thin, bottom=thin)
    C_TITLE = "1F4E79"; C_REC = "2c6a9e"; C_EXT = "3a7abf"; C_VAL = "155724"
    C_ALT = "f7fafe"; C_TOT = "c6efce"

    def hdr(fg, sz=9):
        return Font(bold=True, color="FFFFFF", size=sz), PatternFill("solid", fgColor=fg)

    mes_nombre = ["", "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
                  "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"][int(params["mes"])]

    cols = [
        ("EMPLEADO", C_TITLE), ("DÍA", C_TITLE), ("FECHA", C_TITLE), ("D.SEM", C_TITLE),
        ("TURNO", C_TITLE), ("CICLO", C_TITLE), ("ACUM.\nINICIO", C_TITLE),
        ("TOT REC", C_REC), ("DIURNO", C_REC), ("NOCT", C_REC), ("FEST D", C_REC), ("FEST N", C_REC),
        ("TOT EXT", C_EXT), ("DIURNA", C_EXT), ("NOCT", C_EXT), ("FEST D", C_EXT), ("FEST N", C_EXT),
        ("TOTAL\nHRS", C_TITLE),
        ("VR. RECARGO", C_VAL), ("VR. EXTRA", C_VAL), ("VR. TOTAL", C_VAL),
    ]
    ncol = len(cols)
    last_col = get_column_letter(ncol)

    ws.merge_cells(f"A1:{last_col}1")
    ws["A1"] = f"DETALLE DE NÓMINA POR TURNO/DÍA – {mes_nombre} {params['anio']} · PERSONAL DE SEGURIDAD"
    f, fill = hdr(C_TITLE, sz=12); ws["A1"].font = f; ws["A1"].fill = fill; ws["A1"].alignment = CENTER
    ws.row_dimensions[1].height = 24

    ws.row_dimensions[2].height = 28
    for c, (lbl, fg) in enumerate(cols, 1):
        cell = ws.cell(row=2, column=c, value=lbl)
        f, fill = hdr(fg); cell.font = f; cell.fill = fill; cell.alignment = CENTER; cell.border = BDR

    CICLO_LABEL = {"ant": "Anterior", "act": "Actual", "pos": "Posterior"}
    FMT_NUM = "#,##0"; FMT_HRS = "0.00"
    FILL_ALT = PatternFill("solid", fgColor=C_ALT)
    tot_keys = ["total_rec", "rec_diurno", "rec_nocturno", "rec_fest_d", "rec_fest_n",
                "total_ext", "ext_diurna", "ext_nocturna", "ext_fest_d", "ext_fest_n", "total_hrs",
                "valor_recargo", "valor_extra", "valor_total"]
    tot = {k: 0 for k in tot_keys}

    row = 3
    for r in resultados:
        for d in r.get("detalle_dias", []):
            vals = [
                r.get("nombre", ""), d.get("dia"), _fecha_dmy(d.get("fecha")), d.get("dia_semana"),
                d.get("turno"), CICLO_LABEL.get(d.get("ciclo"), d.get("ciclo")), d.get("acum_ini", 0),
                d.get("total_rec", 0), d.get("rec_diurno", 0), d.get("rec_nocturno", 0), d.get("rec_fest_d", 0), d.get("rec_fest_n", 0),
                d.get("total_ext", 0), d.get("ext_diurna", 0), d.get("ext_nocturna", 0), d.get("ext_fest_d", 0), d.get("ext_fest_n", 0),
                d.get("total_hrs", 0),
                d.get("valor_recargo", 0), d.get("valor_extra", 0), d.get("valor_total", 0),
            ]
            even = (row % 2 == 0)
            for col, val in enumerate(vals, 1):
                cell = ws.cell(row=row, column=col, value=val)
                cell.border = BDR
                if even:
                    cell.fill = FILL_ALT
                if col == 1:
                    cell.alignment = Alignment(vertical="center")
                elif col in (2, 3, 4, 5, 6):
                    cell.alignment = CENTER
                elif col == 7 or (8 <= col <= 18):
                    cell.alignment = RIGHT; cell.number_format = FMT_HRS
                else:
                    cell.alignment = RIGHT; cell.number_format = FMT_NUM
            for k in tot_keys:
                tot[k] += d.get(k, 0)
            row += 1

    n_turnos = sum(len(r.get("detalle_dias", [])) for r in resultados)
    ws.merge_cells(f"A{row}:G{row}")
    c0 = ws.cell(row=row, column=1, value=f"TOTALES — {n_turnos} turnos")
    c0.font = Font(bold=True, size=9); c0.fill = PatternFill("solid", fgColor=C_TOT)
    c0.alignment = CENTER; c0.border = BDR
    for c in range(2, 8):
        cc = ws.cell(row=row, column=c); cc.fill = PatternFill("solid", fgColor=C_TOT); cc.border = BDR
    order = ["total_rec", "rec_diurno", "rec_nocturno", "rec_fest_d", "rec_fest_n",
             "total_ext", "ext_diurna", "ext_nocturna", "ext_fest_d", "ext_fest_n", "total_hrs",
             "valor_recargo", "valor_extra", "valor_total"]
    for i, k in enumerate(order):
        col = 8 + i
        es_hrs = (col <= 18)
        cell = ws.cell(row=row, column=col, value=round(tot[k], 2) if es_hrs else round(tot[k]))
        cell.font = Font(bold=True, size=9); cell.fill = PatternFill("solid", fgColor=C_TOT)
        cell.alignment = RIGHT; cell.number_format = (FMT_HRS if es_hrs else FMT_NUM); cell.border = BDR

    widths = [26, 5, 12, 7, 7, 10, 9, 8, 8, 8, 8, 8, 8, 8, 8, 8, 8, 9, 13, 13, 13]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A3"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ─────────────────────────────────────────────────────────────
# Rutas
# ─────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/defaults")
def api_defaults():
    return jsonify(_load_current())


@app.route("/api/guardar-params", methods=["POST"])
def api_guardar_params():
    try:
        body = request.get_json(force=True) or {}
        cfg = {
            "params": body.get("params", {}),
            "reglas": body.get("reglas", {}),
            "festivos": body.get("festivos", []),
        }
        _save_config(cfg)
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 400


@app.route("/api/calcular", methods=["POST"])
def api_calcular():
    body = request.get_json(force=True)
    try:
        empleados_limpios, validation = _validar_turnos_empleados(body["empleados"])
        if validation["errors"]:
            return jsonify({"ok": False, "error": "Hay turnos inválidos en la plantilla.", "validation": validation}), 400
        resultados = calcular_nomina(
            params=body["params"],
            empleados=empleados_limpios,
            reglas=body["reglas"],
            festivos=body["festivos"],
        )
        return jsonify({"ok": True, "resultados": resultados, "validation": validation})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 400


@app.route("/api/descargar", methods=["POST"])
def api_descargar():
    body = request.get_json(force=True)
    try:
        empleados_limpios, validation = _validar_turnos_empleados(body["empleados"])
        if validation["errors"]:
            return jsonify({"ok": False, "error": "Hay turnos inválidos en la plantilla.", "validation": validation}), 400
        resultados = calcular_nomina(
            params=body["params"],
            empleados=empleados_limpios,
            reglas=body["reglas"],
            festivos=body["festivos"],
        )
        buf = _generar_excel(resultados, body["params"], body.get("reglas", {}))
        mes = int(body["params"]["mes"])
        anio = int(body["params"]["anio"])
        filename = f"Liquidacion_Nomina_{anio}_{mes:02d}.xlsx"
        return send_file(
            buf,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=filename,
        )
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 400


@app.route("/api/descargar-detalle", methods=["POST"])
def api_descargar_detalle():
    body = request.get_json(force=True)
    try:
        empleados_limpios, validation = _validar_turnos_empleados(body["empleados"])
        if validation["errors"]:
            return jsonify({"ok": False, "error": "Hay turnos inválidos en la plantilla.", "validation": validation}), 400
        resultados = calcular_nomina(
            params=body["params"],
            empleados=empleados_limpios,
            reglas=body["reglas"],
            festivos=body["festivos"],
        )
        buf = _generar_excel_detalle(resultados, body["params"])
        mes = int(body["params"]["mes"])
        anio = int(body["params"]["anio"])
        filename = f"Detalle_Nomina_{anio}_{mes:02d}.xlsx"
        return send_file(
            buf,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=filename,
        )
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 400


@app.route("/api/plantilla")
def api_plantilla():
    fp = BASE_DIR / "Plantilla_Nomina.xlsx"
    if not fp.exists():
        return jsonify({"ok": False, "error": "Plantilla no encontrada"}), 404
    return send_file(str(fp), mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True, download_name="Plantilla_Nomina.xlsx")


@app.route("/api/cargar-excel", methods=["POST"])
def api_cargar_excel():
    if "file" not in request.files:
        return jsonify({"ok": False, "error": "No se recibió archivo"}), 400
    f = request.files["file"]
    try:
        wb = openpyxl.load_workbook(f, data_only=True)
        data = _excel_to_data(wb)
        return jsonify({"ok": True, "data": data})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 400


# Inicializar DB al importar (gunicorn no llama __main__)
_init_db()

if __name__ == "__main__":
    print("\n✅  Servidor de nómina arrancado")
    print("   Abre tu navegador en: http://localhost:5050\n")
    app.run(host="0.0.0.0", port=5050, debug=False)
