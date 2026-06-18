"""
Motor de cálculo de nómina - Personal de Seguridad

Logica de ciclos:
  ant: dias del mes ANTES de fecha_inicio_ciclo (cierre del ciclo anterior)
  act: dias dentro del ciclo actual [fic, fic + dias_ciclo - 1]
  pos: dias DESPUES del ciclo actual (inicio del siguiente ciclo)

Recargos: hasta horas_objetivo (132h) por ciclo. Extras: sin techo.

saldo_inicial_horas:
  - Si fic > dia 1 del mes: saldo aplica al bucket "ant"
  - Si fic <= dia 1 del mes: saldo aplica al bucket "act"
"""

from datetime import date, datetime, timedelta
from calendar import monthrange
from typing import Dict, List

TURNO_CATALOGO = {
    "D":  (6*60,  False, True,  True,  True),
    "N":  (18*60, False, True,  True,  True),
    "X":  (None,  False, False, False, False),
    "FD": (6*60,  True,  True,  False, True),
    "FN": (18*60, True,  True,  False, True),
    "AD": (6*60,  False, False, False, True),
    "AN": (18*60, False, False, False, True),
    "IN": (None,  False, False, False, False),
    "VC": (None,  False, True,  False, False),
    "P":  (None,  False, True,  False, False),
    "AU": (None,  False, False, False, False),
    "LC": (None,  False, True,  False, False),
    "ID": (6*60,  False, True,  True,  True),
    "RT": (None,  False, False, False, False),
    "CV": (None,  False, True,  False, False),
    "DE": (6*60,  False, True,  True,  True),
    "PS": (None,  False, True,  False, False),
}

TURNO_INICIO  = {k: v[0] for k, v in TURNO_CATALOGO.items() if v[0] is not None}
TURNO_ES_FEST = {k for k, v in TURNO_CATALOGO.items() if v[1]}
TURNO_N_TYPE  = {"N", "FN", "AN"}

DIAS_SEMANA = ["Lun","Mar","Mié","Jue","Vie","Sáb","Dom"]


def _es_nocturno(minuto_dia, ini_noc, fin_noc, fin_ext=0):
    eff = fin_ext if fin_ext else fin_noc
    if ini_noc > eff:
        return minuto_dia >= ini_noc or minuto_dia < eff
    return ini_noc <= minuto_dia < eff


def calcular_nomina(params, empleados, reglas, festivos):
    festivos_dates = set()
    for f in festivos:
        try:
            festivos_dates.add(datetime.strptime(f, "%Y-%m-%d").date())
        except Exception:
            pass

    anio       = int(params["anio"])
    mes        = int(params["mes"])
    dias_ciclo = int(params["dias_ciclo"])
    horas_obj  = float(params["horas_objetivo"])
    min_turno  = round(float(params.get("horas_turno", 12.583333)) * 60)
    ini_noc    = int(params.get("inicio_noc_h", 19)) * 60
    fin_noc    = int(params.get("fin_noc_h", 6)) * 60
    # horas_mes: divisor mensual para IVH (44h/sem ÷ 6 días × 30 días = 220h/mes)
    horas_mes  = float(params.get("horas_mes", 220))
    obj_min    = horas_obj * 60   # umbral del ciclo para recargos (132h)
    smmlv      = float(params.get("smmlv", 1423500))
    aux_mens   = float(params.get("auxilio_transporte_mensual", 200000))
    limite_aux = smmlv * 2
    # Fase 3: almuerzo explícito — minutos desde el inicio del turno donde ocurre el descanso.
    # Por defecto 6h tras el inicio (turno D inicia 6:00 → almuerzo 12:00; turno N inicia 18:00 → 00:00 del día siguiente).
    alm_offset_min = int(round(float(params.get("almuerzo_offset_h", 6)) * 60))
    alm_dur_min    = int(params.get("almuerzo_duracion_min", 25))

    _, dias_mes = monthrange(anio, mes)
    primer_dia_mes = date(anio, mes, 1)

    # Aplicar overrides de CODIGOS_TURNO si fueron cargados desde el Excel
    turno_params = params.get("turno_params", {})
    catalogo = {}
    for k, v in TURNO_CATALOGO.items():
        override = turno_params.get(k, {})
        paga_dia  = override.get("paga_dia",  v[2]) if override.get("paga_dia")  is not None else v[2]
        aux_transp= override.get("aux_transp", v[3]) if override.get("aux_transp") is not None else v[3]
        catalogo[k] = (v[0], v[1], paga_dia, aux_transp, v[4])

    resultados = []

    for emp in empleados:
        nombre  = emp.get("nombre", "")
        salario = float(emp.get("salario_mensual", 0))
        ivh     = salario / horas_mes if horas_mes else 0  # IVH = Salario / 220

        try:
            fic = datetime.strptime(emp["fecha_inicio_ciclo"], "%Y-%m-%d").date()
        except Exception:
            fic = date(anio, mes, 1)

        fic_fin       = fic + timedelta(days=dias_ciclo - 1)
        saldo_ini_min = float(emp.get("saldo_inicial_horas", 0)) * 60
        turnos        = emp.get("turnos", {})

        TIPOS_REC = ["DIURNO", "NOCTURNO", "FEST_DIURNO", "FEST_NOCTURNO"]
        TIPOS_EXT = ["DIURNA", "NOCTURNA", "FEST_DIURNA", "FEST_NOCTURNA"]

        saldo_bucket = "ant" if fic > primer_dia_mes else "act"

        buckets = {
            c: {
                "rec":  {t: 0.0 for t in TIPOS_REC},
                "ext":  {t: 0.0 for t in TIPOS_EXT},
                "acum": saldo_ini_min if c == saldo_bucket else 0.0,
            }
            for c in ("ant", "act", "pos")
        }

        dias_paga_dia   = 0
        dias_aux_transp = 0
        dias_trabajados = emp.get("dias_trabajados_manual", 0) or 0

        for d_num in range(1, dias_mes + 1):
            t = str(turnos.get(str(d_num), "")).strip().upper()
            if not t:
                continue
            cat = catalogo.get(t)
            if cat is None:
                continue
            if cat[2]:
                dias_paga_dia += 1
            if cat[3]:
                dias_aux_transp += 1

        all_segs = []

        for d_num in range(1, dias_mes + 1):
            turno = str(turnos.get(str(d_num), "")).strip().upper()
            if turno not in TURNO_INICIO:
                continue

            if not emp.get("dias_trabajados_manual"):
                dias_trabajados += 1
            td            = date(anio, mes, d_num)
            inicio        = TURNO_INICIO[turno]
            es_fest_base  = turno in TURNO_ES_FEST
            cursor        = inicio
            remaining     = min_turno

            # Fase 3: rango absoluto del almuerzo (minutos desde 00:00 del día base del turno).
            # Los segmentos se cortan en lunch_ini y se reanudan en lunch_fin; el tiempo del
            # almuerzo no genera horas liquidadas, pero el cursor sí avanza para que la
            # clasificación (día, fest, nocturno) se aplique al MOMENTO REAL donde ocurre.
            lunch_ini = inicio + alm_offset_min
            lunch_fin = lunch_ini + alm_dur_min

            # Detectar si el almuerzo absorbe un cambio de día (medianoche).
            # Cuando esto ocurre, los segmentos post-almuerzo dentro de la
            # misma franja horaria conservan la clasificación festiva del
            # día anterior (el trabajador estaba en descanso al cambiar el día).
            first_midnight = ((inicio // (24 * 60)) + 1) * (24 * 60)
            lunch_absorbs_midnight = (alm_dur_min > 0
                                      and lunch_ini <= first_midnight <= lunch_fin)

            while remaining > 0:
                # Saltar la franja del almuerzo (no se acumula ni se clasifica).
                if alm_dur_min > 0 and lunch_ini <= cursor < lunch_fin:
                    cursor = lunch_fin
                    continue

                dur = min(60, remaining)
                # No cruzar el inicio del almuerzo en el mismo segmento
                if alm_dur_min > 0 and cursor < lunch_ini:
                    dur = min(dur, lunch_ini - cursor)
                # Mantener la segmentación alineada al cambio de hora-en-punto
                next_hour_edge = ((cursor // 60) + 1) * 60
                dur = min(dur, next_hour_edge - cursor)

                de  = cursor // (24 * 60)
                sd  = td + timedelta(days=de)
                md  = cursor % (24 * 60)
                fin_ext = (fin_noc + 60) if turno in TURNO_N_TYPE else fin_noc

                # Fecha de clasificación para el estado festivo:
                # - es_fest_base (FN/FD) solo aplica al día de inicio del turno
                # - Si el almuerzo absorbe la medianoche, los segmentos
                #   post-almuerzo hasta la siguiente hora en punto usan
                #   la fecha del día de inicio del turno
                if (lunch_absorbs_midnight
                        and cursor >= lunch_fin
                        and cursor < first_midnight + 60):
                    class_date = td
                else:
                    class_date = sd

                all_segs.append({
                    "order":    (td, cursor),
                    "seg_date": sd,
                    "dur":      dur,
                    "noc":      _es_nocturno(md, ini_noc, fin_noc, fin_ext),
                    "fest_dom": (class_date.weekday() == 6
                                or class_date in festivos_dates
                                or (es_fest_base and class_date == td)),
                    "td":       td,
                    "d_num":    d_num,
                })
                cursor    += dur
                remaining -= dur

        all_segs.sort(key=lambda s: s["order"])

        # Detalle por turno (dia)
        turn_agg = {}  # d_num -> {rec, ext, total_min, ciclo, acum_ini}

        for seg in all_segs:
            sd       = seg["seg_date"]
            td_seg   = seg["td"]
            # Fase 4: el ciclo (y por tanto el umbral de 132h) se decide por la fecha de
            # INICIO del turno, no del segmento. Así un turno N que arranca el último día
            # del ciclo y cruza al ciclo siguiente NO reinicia el contador a 0h, evitando
            # que recargos aparezcan tras superar las horas objetivo.
            ciclo = "ant" if td_seg < fic else ("act" if td_seg <= fic_fin else "pos")
            bk    = buckets[ciclo]
            d_num = seg["d_num"]
            dur, noc, fest = seg["dur"], seg["noc"], seg["fest_dom"]
            acum  = bk["acum"]

            # Inicializar entrada de detalle para este turno
            if d_num not in turn_agg:
                turn_agg[d_num] = {
                    "rec":      {t: 0.0 for t in TIPOS_REC},
                    "ext":      {t: 0.0 for t in TIPOS_EXT},
                    "total_min": 0,
                    "ciclo":    ciclo,  # ciclo del primer segmento
                    "acum_ini": acum,   # acumulado al inicio del turno
                }

            rec_disp = max(0.0, obj_min - acum)
            min_rec  = min(dur, rec_disp)
            min_ext  = dur - min_rec

            if min_rec > 0:
                h = min_rec / 60.0
                if fest and noc:
                    bk["rec"]["FEST_NOCTURNO"] += h
                    turn_agg[d_num]["rec"]["FEST_NOCTURNO"] += h
                elif fest:
                    bk["rec"]["FEST_DIURNO"] += h
                    turn_agg[d_num]["rec"]["FEST_DIURNO"] += h
                elif noc:
                    bk["rec"]["NOCTURNO"] += h
                    turn_agg[d_num]["rec"]["NOCTURNO"] += h
                else:
                    bk["rec"]["DIURNO"] += h
                    turn_agg[d_num]["rec"]["DIURNO"] += h

            if min_ext > 0:
                h = min_ext / 60.0
                if fest and noc:
                    bk["ext"]["FEST_NOCTURNA"] += h
                    turn_agg[d_num]["ext"]["FEST_NOCTURNA"] += h
                elif fest:
                    bk["ext"]["FEST_DIURNA"] += h
                    turn_agg[d_num]["ext"]["FEST_DIURNA"] += h
                elif noc:
                    bk["ext"]["NOCTURNA"] += h
                    turn_agg[d_num]["ext"]["NOCTURNA"] += h
                else:
                    bk["ext"]["DIURNA"] += h
                    turn_agg[d_num]["ext"]["DIURNA"] += h

            turn_agg[d_num]["total_min"] += dur
            bk["acum"] += dur

        def _vr(rec):
            return (rec["DIURNO"] * ivh * reglas.get("REC_DIURNO", 0.0)
                  + rec["NOCTURNO"] * ivh * reglas.get("REC_NOCTURNO", 0.35)
                  + rec["FEST_DIURNO"] * ivh * reglas.get("REC_DOM_FEST_DIURNO", 0.80)
                  + rec["FEST_NOCTURNO"] * ivh * reglas.get("REC_DOM_FEST_NOCTURNO", 1.15))

        def _ve(ext):
            return (ext["DIURNA"] * ivh * reglas.get("EXT_DIURNA", 1.25)
                  + ext["NOCTURNA"] * ivh * reglas.get("EXT_NOCTURNA", 1.75)
                  + ext["FEST_DIURNA"] * ivh * reglas.get("EXT_FEST_DIURNA", 2.05)
                  + ext["FEST_NOCTURNA"] * ivh * reglas.get("EXT_FEST_NOCTURNA", 2.55))

        val_rec = sum(_vr(buckets[c]["rec"]) for c in ("ant", "act", "pos"))
        val_ext = sum(_ve(buckets[c]["ext"]) for c in ("ant", "act", "pos"))

        salario_prop = round(salario * dias_paga_dia / dias_mes) if dias_mes else round(salario)

        aux_manual = emp.get("auxilio_transporte_manual")
        if aux_manual is not None and aux_manual > 0:
            auxilio = round(float(aux_manual))
        elif salario <= limite_aux and dias_mes > 0:
            auxilio = round(aux_mens * dias_aux_transp / dias_mes)
        else:
            auxilio = 0

        hrs_rec = {t: sum(buckets[c]["rec"][t] for c in ("ant","act","pos")) for t in TIPOS_REC}
        hrs_ext = {t: sum(buckets[c]["ext"][t] for c in ("ant","act","pos")) for t in TIPOS_EXT}
        total_hrs = sum(hrs_rec.values()) + sum(hrs_ext.values())

        # Valor en pesos por cada tipo: horas x IVH x factor (desde Reglas de Liquidacion)
        val = {
            "rec_diurno": round(hrs_rec["DIURNO"]        * ivh * reglas.get("REC_DIURNO", 0.0)),
            "rec_noct":   round(hrs_rec["NOCTURNO"]      * ivh * reglas.get("REC_NOCTURNO", 0.35)),
            "rec_fest_d": round(hrs_rec["FEST_DIURNO"]   * ivh * reglas.get("REC_DOM_FEST_DIURNO", 0.80)),
            "rec_fest_n": round(hrs_rec["FEST_NOCTURNO"] * ivh * reglas.get("REC_DOM_FEST_NOCTURNO", 1.15)),
            "ext_diurna": round(hrs_ext["DIURNA"]        * ivh * reglas.get("EXT_DIURNA", 1.25)),
            "ext_noct":   round(hrs_ext["NOCTURNA"]      * ivh * reglas.get("EXT_NOCTURNA", 1.75)),
            "ext_fest_d": round(hrs_ext["FEST_DIURNA"]   * ivh * reglas.get("EXT_FEST_DIURNA", 2.05)),
            "ext_fest_n": round(hrs_ext["FEST_NOCTURNA"] * ivh * reglas.get("EXT_FEST_NOCTURNA", 2.55)),
        }
        val_rec_total = val["rec_diurno"] + val["rec_noct"] + val["rec_fest_d"] + val["rec_fest_n"]
        val_ext_total = val["ext_diurna"] + val["ext_noct"] + val["ext_fest_d"] + val["ext_fest_n"]

        # El salario base se paga COMPLETO (parte fija). Total = suma exacta de columnas.
        total = round(salario) + val_rec_total + val_ext_total + auxilio

        ciclos_detalle = {}
        for c in ("ant", "act", "pos"):
            vr_c  = _vr(buckets[c]["rec"])
            ve_c  = _ve(buckets[c]["ext"])
            rec_c = buckets[c]["rec"]
            ext_c = buckets[c]["ext"]
            hrs_c = sum(rec_c.values()) + sum(ext_c.values())
            # Fase 2.3: agregados por familia de concepto, útiles para el panel "Resumen del ciclo vigente".
            h_diurnas   = rec_c["DIURNO"]        + ext_c["DIURNA"]
            h_nocturnas = rec_c["NOCTURNO"]      + ext_c["NOCTURNA"]
            h_fest_d    = rec_c["FEST_DIURNO"]   + ext_c["FEST_DIURNA"]
            h_fest_n    = rec_c["FEST_NOCTURNO"] + ext_c["FEST_NOCTURNA"]
            ciclos_detalle[c] = {
                "rec":         {t: round(rec_c[t], 2) for t in TIPOS_REC},
                "ext":         {t: round(ext_c[t], 2) for t in TIPOS_EXT},
                "total_hrs":   round(hrs_c, 2),
                "total_rec":   round(sum(rec_c.values()), 2),
                "total_ext":   round(sum(ext_c.values()), 2),
                "h_diurnas":   round(h_diurnas, 2),
                "h_nocturnas": round(h_nocturnas, 2),
                "h_fest_d":    round(h_fest_d, 2),
                "h_fest_n":    round(h_fest_n, 2),
                "h_dom_fest":  round(h_fest_d + h_fest_n, 2),
                "acum":        round(buckets[c]["acum"] / 60.0, 2),
                "val_rec":     round(vr_c),
                "val_ext":     round(ve_c),
                "total_val":   round(vr_c + ve_c),
            }

        # Construir lista de detalle por dia ordenada
        detalle_dias = []
        for d_num in sorted(turn_agg.keys()):
            ag = turn_agg[d_num]
            td = date(anio, mes, d_num)
            total_rec_h = sum(ag["rec"].values())
            total_ext_h = sum(ag["ext"].values())
            total_h = ag["total_min"] / 60.0
            # Fase 2.1: valor monetario por concepto en cada día (mismo cálculo que el resumen).
            vr_rec_d  = ag["rec"]["DIURNO"]        * ivh * reglas.get("REC_DIURNO", 0.0)
            vr_rec_n  = ag["rec"]["NOCTURNO"]      * ivh * reglas.get("REC_NOCTURNO", 0.35)
            vr_rec_fd = ag["rec"]["FEST_DIURNO"]   * ivh * reglas.get("REC_DOM_FEST_DIURNO", 0.80)
            vr_rec_fn = ag["rec"]["FEST_NOCTURNO"] * ivh * reglas.get("REC_DOM_FEST_NOCTURNO", 1.15)
            ve_ext_d  = ag["ext"]["DIURNA"]        * ivh * reglas.get("EXT_DIURNA", 1.25)
            ve_ext_n  = ag["ext"]["NOCTURNA"]      * ivh * reglas.get("EXT_NOCTURNA", 1.75)
            ve_ext_fd = ag["ext"]["FEST_DIURNA"]   * ivh * reglas.get("EXT_FEST_DIURNA", 2.05)
            ve_ext_fn = ag["ext"]["FEST_NOCTURNA"] * ivh * reglas.get("EXT_FEST_NOCTURNA", 2.55)
            vr_d = vr_rec_d + vr_rec_n + vr_rec_fd + vr_rec_fn
            ve_d = ve_ext_d + ve_ext_n + ve_ext_fd + ve_ext_fn
            detalle_dias.append({
                "dia":        d_num,
                "fecha":      td.strftime("%Y-%m-%d"),
                "dia_semana": DIAS_SEMANA[td.weekday()],
                "turno":      turnos.get(str(d_num), ""),
                "ciclo":      ag["ciclo"],
                "acum_ini":   round(ag["acum_ini"] / 60.0, 2),
                "total_hrs":  round(total_h, 2),
                "rec_diurno":     round(ag["rec"]["DIURNO"], 2),
                "rec_nocturno":   round(ag["rec"]["NOCTURNO"], 2),
                "rec_fest_d":     round(ag["rec"]["FEST_DIURNO"], 2),
                "rec_fest_n":     round(ag["rec"]["FEST_NOCTURNO"], 2),
                "total_rec":      round(total_rec_h, 2),
                "ext_diurna":     round(ag["ext"]["DIURNA"], 2),
                "ext_nocturna":   round(ag["ext"]["NOCTURNA"], 2),
                "ext_fest_d":     round(ag["ext"]["FEST_DIURNA"], 2),
                "ext_fest_n":     round(ag["ext"]["FEST_NOCTURNA"], 2),
                "total_ext":      round(total_ext_h, 2),
                "vr_rec_diurno":   round(vr_rec_d),
                "vr_rec_nocturno": round(vr_rec_n),
                "vr_rec_fest_d":   round(vr_rec_fd),
                "vr_rec_fest_n":   round(vr_rec_fn),
                "vr_ext_diurna":   round(ve_ext_d),
                "vr_ext_nocturna": round(ve_ext_n),
                "vr_ext_fest_d":   round(ve_ext_fd),
                "vr_ext_fest_n":   round(ve_ext_fn),
                "valor_recargo":   round(vr_d),
                "valor_extra":     round(ve_d),
                "valor_total":     round(vr_d + ve_d),
            })

        resultados.append({
            "id":                   emp.get("id"),
            "nombre":               nombre,
            "salario_mensual":      salario,
            "salario_proporcional": salario_prop,
            "dias_paga_dia":        dias_paga_dia,
            "dias_aux_transp":      dias_aux_transp,
            "dias_trabajados":      dias_trabajados,
            "total_horas":          round(total_hrs, 2),
            "hrs_diurnas":          round(hrs_rec["DIURNO"], 2),
            "hrs_nocturnas":        round(hrs_rec["NOCTURNO"], 2),
            "hrs_fest_diurnas":     round(hrs_rec["FEST_DIURNO"] + hrs_ext["FEST_DIURNA"], 2),
            "hrs_fest_noc":         round(hrs_rec["FEST_NOCTURNO"] + hrs_ext["FEST_NOCTURNA"], 2),
            "hrs_ext_diurnas":      round(hrs_ext["DIURNA"], 2),
            "hrs_ext_noc":          round(hrs_ext["NOCTURNA"], 2),
            # Horas separadas por concepto exacto (sin mezcla fest_rec con fest_ext).
            "hrs_por_concepto": {
                "rec_diurno":   round(hrs_rec["DIURNO"], 2),
                "rec_nocturno": round(hrs_rec["NOCTURNO"], 2),
                "rec_fest_d":   round(hrs_rec["FEST_DIURNO"], 2),
                "rec_fest_n":   round(hrs_rec["FEST_NOCTURNO"], 2),
                "ext_diurna":   round(hrs_ext["DIURNA"], 2),
                "ext_nocturna": round(hrs_ext["NOCTURNA"], 2),
                "ext_fest_d":   round(hrs_ext["FEST_DIURNA"], 2),
                "ext_fest_n":   round(hrs_ext["FEST_NOCTURNA"], 2),
            },
            "valor_recargo":        val_rec_total,
            "valor_extra":          val_ext_total,
            "val":                  val,
            "auxilio_transporte":   auxilio,
            "total_pagar":          total,
            "fic":                  fic.strftime("%Y-%m-%d"),
            "fic_fin":              fic_fin.strftime("%Y-%m-%d"),
            "fic_siguiente":        (fic_fin + timedelta(days=1)).strftime("%Y-%m-%d"),
            "horas_objetivo":       horas_obj,
            "ciclos":               ciclos_detalle,
            "detalle_dias":         detalle_dias,
        })

    return resultados
