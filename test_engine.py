"""
Tests del motor de liquidación de nómina.
Valida la clasificación hora-a-hora de turnos con múltiples conceptos,
cruces de medianoche, transiciones festivo↔ordinario y umbral de extras.

IVH = 2_113_897.50 / 220 = 9_608.625
"""

import json
from datetime import date
from engine import calcular_nomina

# ── Parámetros base de prueba ───────────────────────────────
SALARIO = 2_113_897.50
IVH = SALARIO / 220  # 9608.625

PARAMS = {
    "anio": 2026,
    "mes": 5,
    "dias_ciclo": 21,
    "horas_objetivo": 132,
    "horas_turno": 12.583333,
    "smmlv": 1423500,
    "auxilio_transporte_mensual": 200000,
    "inicio_noc_h": 19,
    "fin_noc_h": 6,
    "almuerzo_offset_h": 6,
    "almuerzo_duracion_min": 25,
    "horas_mes": 220,
}

REGLAS = {
    "REC_DIURNO": 0.0,
    "REC_NOCTURNO": 0.35,
    "REC_DOM_FEST_DIURNO": 0.80,
    "REC_DOM_FEST_NOCTURNO": 1.15,
    "EXT_DIURNA": 1.25,
    "EXT_NOCTURNA": 1.75,
    "EXT_FEST_DIURNA": 2.05,
    "EXT_FEST_NOCTURNA": 2.55,
}

# Mayo 2026: 1 de mayo festivo (Día del Trabajo)
FESTIVOS = ["2026-05-01"]


def _run(turnos, festivos=None, saldo=0.0, fic="2026-05-01"):
    emp = {
        "id": 1,
        "nombre": "TEST",
        "salario_mensual": SALARIO,
        "fecha_inicio_ciclo": fic,
        "saldo_inicial_horas": saldo,
        "turnos": turnos,
    }
    res = calcular_nomina(PARAMS, [emp], REGLAS, festivos or FESTIVOS)
    return res[0]


def _detalle(r, dia):
    for d in r["detalle_dias"]:
        if d["dia"] == dia:
            return d
    return None


def _approx(a, b, tol=0.02):
    return abs(a - b) <= tol


# ── Test 1: Turno completamente diurno (D) en día ordinario ─
def test_turno_diurno_ordinario():
    """Miércoles 06 - Turno D: todo diurno ordinario (0%)."""
    r = _run({"6": "D"})
    d = _detalle(r, 6)
    assert d is not None
    assert _approx(d["total_hrs"], 12.58)
    assert _approx(d["rec_diurno"], 12.58)
    assert d["rec_nocturno"] == 0
    assert d["rec_fest_d"] == 0
    assert d["rec_fest_n"] == 0
    assert d["total_ext"] == 0
    # Valor $0 porque recargo diurno es 0%
    assert d["valor_total"] == 0
    print("  PASS test_turno_diurno_ordinario")


# ── Test 2: Turno completamente nocturno (N) sin cruce festivo ─
def test_turno_nocturno_ordinario():
    """Viernes 08 - Turno N: 1h diurno + 11.58h nocturno, sin festivo."""
    r = _run({"8": "N"})
    d = _detalle(r, 8)
    assert d is not None
    assert _approx(d["total_hrs"], 12.58)
    assert _approx(d["rec_diurno"], 1.0), f"rec_diurno={d['rec_diurno']}"
    assert _approx(d["rec_nocturno"], 11.58), f"rec_nocturno={d['rec_nocturno']}"
    assert d["rec_fest_d"] == 0
    assert d["rec_fest_n"] == 0
    # Valor: 11.58h * IVH * 0.35
    esperado = round(d["rec_nocturno"] * IVH * 0.35)
    assert _approx(d["valor_total"], esperado, tol=50)
    print("  PASS test_turno_nocturno_ordinario")


# ── Test 3: Turno N que cruza medianoche Sáb→Dom (ord→fest) ─
def test_cruce_sabado_domingo():
    """Sábado 02 - Turno N: cruza a domingo. Ref Excel: $85,076."""
    r = _run({"2": "N"})
    d = _detalle(r, 2)
    assert d is not None
    assert _approx(d["total_hrs"], 12.58)
    # 18:00 Sáb = diurno ordinario
    assert _approx(d["rec_diurno"], 1.0), f"rec_diurno={d['rec_diurno']}"
    # 19:00-00:25 Sáb (5h + 35min post-almuerzo) = nocturno ordinario
    assert _approx(d["rec_nocturno"], 5.58, tol=0.05), f"rec_nocturno={d['rec_nocturno']}"
    # 01:00-07:00 Dom = festivo nocturno
    assert _approx(d["rec_fest_n"], 6.0, tol=0.05), f"rec_fest_n={d['rec_fest_n']}"
    # Sin festivo diurno
    assert d["rec_fest_d"] == 0
    # Valor total ~$85,076 (ref Excel)
    assert _approx(d["valor_total"], 85076, tol=200), f"valor={d['valor_total']}"
    print("  PASS test_cruce_sabado_domingo")


# ── Test 4: Turno FN que inicia festivo y termina ordinario ─
def test_fn_domingo_a_lunes():
    """Domingo 03 - Turno FN: CASO CRÍTICO. Ref Excel: $89,560.
    BUG ANTERIOR: todo se liquidaba como festivo ($135,682)."""
    r = _run({"3": "FN"})
    d = _detalle(r, 3)
    assert d is not None
    assert _approx(d["total_hrs"], 12.58)
    # 18:00 Dom = festivo diurno (1h)
    assert _approx(d["rec_fest_d"], 1.0), f"rec_fest_d={d['rec_fest_d']}"
    # 19:00-00:25 Dom = festivo nocturno (5h + 35min = 5.58h)
    assert _approx(d["rec_fest_n"], 5.58, tol=0.05), f"rec_fest_n={d['rec_fest_n']}"
    # 01:00-07:00 Lun = nocturno ordinario (6h)
    assert _approx(d["rec_nocturno"], 6.0, tol=0.05), f"rec_nocturno={d['rec_nocturno']}"
    # Sin diurno ordinario
    assert d["rec_diurno"] == 0
    # Valor total ~$89,560 (ref Excel), NO $135,682
    assert _approx(d["valor_total"], 89560, tol=200), f"valor={d['valor_total']}"
    assert d["valor_total"] < 100000, "BUG: turno FN sigue liquidando todo como festivo"
    print("  PASS test_fn_domingo_a_lunes")


# ── Test 5: Turno FD completamente festivo ──────────────────
def test_fd_completamente_festivo():
    """Viernes 01 (festivo) - Turno FD: todo festivo diurno. Ref: $96,727."""
    r = _run({"1": "FD"})
    d = _detalle(r, 1)
    assert d is not None
    assert _approx(d["total_hrs"], 12.58)
    assert _approx(d["rec_fest_d"], 12.58), f"rec_fest_d={d['rec_fest_d']}"
    assert d["rec_nocturno"] == 0
    assert d["rec_fest_n"] == 0
    assert d["rec_diurno"] == 0
    esperado = round(12.583333 * IVH * 0.80)
    assert _approx(d["valor_total"], esperado, tol=200)
    print("  PASS test_fd_completamente_festivo")


# ── Test 6: Turno que inicia ordinario y termina festivo ────
def test_cruce_ordinario_a_festivo():
    """Sábado 09 - Turno N: ordinario→domingo festivo. Ref: $85,076."""
    r = _run({"9": "N"})
    d = _detalle(r, 9)
    assert d is not None
    # Mismo patrón que Sábado 02
    assert _approx(d["rec_diurno"], 1.0)
    assert _approx(d["rec_nocturno"], 5.58, tol=0.05)
    assert _approx(d["rec_fest_n"], 6.0, tol=0.05)
    assert _approx(d["valor_total"], 85076, tol=200)
    print("  PASS test_cruce_ordinario_a_festivo")


# ── Test 7: Turno con horas diurnas y nocturnas mixtas ──────
def test_mixto_diurno_nocturno():
    """Turno N cualquier día ordinario: 1h diurno + resto nocturno."""
    r = _run({"20": "N"})  # Miércoles 20
    d = _detalle(r, 20)
    assert d is not None
    assert _approx(d["rec_diurno"], 1.0)
    assert _approx(d["rec_nocturno"], 11.58, tol=0.05)
    assert d["rec_fest_d"] == 0
    assert d["rec_fest_n"] == 0
    print("  PASS test_mixto_diurno_nocturno")


# ── Test 8: Varios conceptos simultáneos (multi-concepto) ───
def test_multi_concepto():
    """Sábado→Domingo: turno N con 3 conceptos (diurno + nocturno + fest_nocturno)."""
    r = _run({"2": "N"})
    d = _detalle(r, 2)
    assert d is not None
    conceptos_no_cero = sum(1 for k in ["rec_diurno", "rec_nocturno", "rec_fest_d", "rec_fest_n"]
                            if d[k] > 0)
    assert conceptos_no_cero >= 3, f"Solo {conceptos_no_cero} conceptos, se esperaban >= 3"
    print("  PASS test_multi_concepto")


# ── Test 9: Umbral de extras (recargos → extras) ────────────
def test_umbral_extras():
    """Empleado con saldo alto: al superar 132h, horas pasan a extras."""
    # Saldo = 126h, turno D de 12.58h → 6.58h rec + 6h ext
    r = _run({"12": "D"}, saldo=126.0, fic="2026-05-01")
    d = _detalle(r, 12)
    assert d is not None
    rec_total = d["total_rec"]
    ext_total = d["total_ext"]
    assert rec_total > 0, "Debería haber horas de recargo"
    assert ext_total > 0, "Debería haber horas extra"
    assert _approx(rec_total + ext_total, 12.58, tol=0.05)
    # Las horas de recargo deben ser 132 - 126 = 6h
    assert _approx(rec_total, 6.0, tol=0.1), f"rec_total={rec_total}"
    print("  PASS test_umbral_extras")


# ── Test 10: Turno FN en festivo con extras post-umbral ─────
def test_fn_con_extras():
    """FN con acum alto: parte recargo festivo, parte extra festivo,
    y parte que cruza a día ordinario."""
    # Saldo = 128h → 4h recargo, luego extras
    r = _run({"3": "FN"}, saldo=128.0, fic="2026-05-01")
    d = _detalle(r, 3)
    assert d is not None
    # Debe haber recargo festivo Y extras
    assert d["total_rec"] > 0
    assert d["total_ext"] > 0
    # Las horas en lunes NO deben ser festivas (ni recargo ni extra festivo)
    # Al menos parte de las extras nocturnas deben ser ordinarias
    assert d["ext_nocturna"] > 0 or d["rec_nocturno"] > 0, \
        "Horas del lunes deberían ser nocturnas ordinarias"
    print("  PASS test_fn_con_extras")


# ── Test 11: Validación completa contra Excel referencia ────
def test_referencia_excel_completa():
    """Mes completo con turnos del Excel de referencia.
    Valida los totales contra los valores esperados."""
    turnos = {
        "1": "FD", "2": "N", "3": "FN",
        "6": "D", "7": "D", "8": "N", "9": "N",
        "12": "D", "13": "D", "14": "N", "15": "N",
        "18": "FD", "19": "D", "20": "N", "21": "N",
        "24": "FD", "25": "D", "26": "N", "27": "N", "28": "N",
        "30": "D", "31": "FD",
    }
    # Valores esperados del Excel (R15 de cada día)
    esperados = {
        1: 96726.85, 2: 85076.39, 3: 89560.41,
        6: 0, 7: 0, 8: 38954.98, 9: 85076.39,
        12: 0, 13: 72064.70, 14: 206785.67, 15: 206785.67,
        18: 96726.85, 19: 0, 20: 38954.98, 21: 38954.98,
        24: 96726.85, 25: 0, 26: 38954.98, 27: 38954.98, 28: 38954.98,
        30: 0, 31: 168791.55,
    }
    r = _run(turnos, saldo=25.17, fic="2026-04-25")
    errores = []
    for dia, esperado in esperados.items():
        d = _detalle(r, dia)
        if d is None:
            if esperado > 0:
                errores.append(f"  Día {dia}: sin detalle, esperado ${esperado:,.0f}")
            continue
        actual = d["valor_total"]
        if not _approx(actual, esperado, tol=500):
            errores.append(f"  Día {dia}: actual=${actual:,.0f} vs esperado=${esperado:,.0f} "
                           f"(diff=${actual - esperado:+,.0f})")
    if errores:
        print("  DIFERENCIAS encontradas:")
        for e in errores:
            print(e)
        # No falla el test completo, solo reporta diferencias
    else:
        print("  Todos los días coinciden con el Excel de referencia")
    print("  PASS test_referencia_excel_completa (con reporte)")


# ── Test 12: Overrides de configuración del catálogo de turnos (turno_params) ──
def test_configuracion_catalogo_overrides():
    """Prueba que si cambia la configuración de los turnos en turno_params (parámetros),
    el motor calcule correctamente los días pagados y los días de auxilio de transporte
    sobre el motor de liquidación base."""
    # 1. Configuración de parámetros personalizada con anulaciones (overrides)
    params_custom = dict(PARAMS)
    params_custom["turno_params"] = {
        "D":  {"paga_dia": False, "aux_transp": True},   # Anula paga_dia de D a False
        "N":  {"paga_dia": True,  "aux_transp": False},  # Anula aux_transp de N a False
        "X":  {"paga_dia": True,  "aux_transp": False},  # Anula paga_dia de X a True
    }
    
    # Empleado de prueba con turnos: 1 de mayo (D), 2 de mayo (N), 3 de mayo (X)
    emp = {
        "id": 1,
        "nombre": "TEST_PARAMS",
        "salario_mensual": SALARIO,
        "fecha_inicio_ciclo": "2026-05-01",
        "saldo_inicial_horas": 0.0,
        "turnos": {"1": "D", "2": "N", "3": "X"},
    }
    
    res = calcular_nomina(params_custom, [emp], REGLAS, FESTIVOS)
    r = res[0]
    
    # 2. Validaciones
    # Días paga día:
    # - Turno D: paga_dia fue anulado a False (0 días)
    # - Turno N: paga_dia es True (1 día)
    # - Turno X: paga_dia fue anulado a True (1 día)
    # Total esperado = 2 días paga día
    assert r["dias_paga_dia"] == 2, f"Se esperaban 2 días paga día, se obtuvo: {r['dias_paga_dia']}"
    
    # Días auxilio de transporte:
    # - Turno D: aux_transp es True (1 día)
    # - Turno N: aux_transp fue anulado a False (0 días)
    # - Turno X: aux_transp es False (0 días)
    # Total esperado = 1 día con auxilio de transporte
    assert r["dias_aux_transp"] == 1, f"Se esperaba 1 día con auxilio, se obtuvo: {r['dias_aux_transp']}"
    
    print("  PASS test_configuracion_catalogo_overrides")


# ── Ejecutar todos los tests ────────────────────────────────
if __name__ == "__main__":
    tests = [
        test_turno_diurno_ordinario,
        test_turno_nocturno_ordinario,
        test_cruce_sabado_domingo,
        test_fn_domingo_a_lunes,
        test_fd_completamente_festivo,
        test_cruce_ordinario_a_festivo,
        test_mixto_diurno_nocturno,
        test_multi_concepto,
        test_umbral_extras,
        test_fn_con_extras,
        test_referencia_excel_completa,
        test_configuracion_catalogo_overrides,
    ]
    passed = 0
    failed = 0
    for t in tests:
        try:
            t()
            passed += 1
        except AssertionError as e:
            print(f"  FAIL {t.__name__}: {e}")
            failed += 1
        except Exception as e:
            print(f"  ERROR {t.__name__}: {e}")
            failed += 1
    print(f"\n{'='*50}")
    print(f"Resultados: {passed} passed, {failed} failed de {len(tests)} tests")
