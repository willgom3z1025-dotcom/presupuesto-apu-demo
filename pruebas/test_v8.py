"""Pruebas automáticas de la app v8 (presupuesto APU).

Uso:  python3 pruebas/test_v8.py [ruta de v8/index.html]
Requisitos: playwright (chromium), pypdf y openpyxl. Para el PDF y el Excel, las librerías jsPDF 2.5.1,
jspdf-autotable 3.8.2 y ExcelJS 4.4.0 se sirven desde node_modules (variable LIBS) en lugar del CDN.
Cada comprobación cuenta como una prueba; al final se imprime el total de correctas y fallidas.
"""
import json, os, re, sys, io, math
from pathlib import Path
from playwright.sync_api import sync_playwright
from pypdf import PdfReader
import openpyxl

RAIZ = Path(__file__).resolve().parent.parent
HTML = Path(sys.argv[1]) if len(sys.argv) > 1 else RAIZ / "v8" / "index.html"
URL = HTML.resolve().as_uri()
CAT = json.loads((HTML.parent / "catalogo.json").read_text(encoding="utf-8"))
LIBS = Path(os.environ.get("LIBS", "/home/claude/libs/node_modules"))
CDN = {
    "jspdf.umd.min.js": LIBS / "jspdf/dist/jspdf.umd.min.js",
    "jspdf.plugin.autotable.min.js": LIBS / "jspdf-autotable/dist/jspdf.plugin.autotable.min.js",
    "exceljs.min.js": LIBS / "exceljs/dist/exceljs.min.js",
}
RES = []
GRUPO = [""]


def chk(nombre, cond, detalle=""):
    RES.append((GRUPO[0], nombre, bool(cond), "" if cond else str(detalle)[:300]))


def grupo(n):
    GRUPO[0] = n
    print("…", n, flush=True)


def R2(x):
    return math.floor(x * 100 + 0.5 + 1e-9) / 100 if x >= 0 else -math.floor(-x * 100 + 0.5 + 1e-9) / 100


REC = {r["id"]: r for r in CAT["recursos"]}


def pu_py(a):
    t = 0.0
    for rid, c in a["lineas"]:
        r = REC.get(rid)
        if not r:
            continue
        t += c * r["precio"] * (1 + (r.get("prest") or 0) / 100) if r["tipo"] == "MO" else c * r["precio"]
    t += (a.get("eq") or 0) + (a.get("pdir") or 0)
    return R2(t)


def money_ok(txt, v):
    return f"{v:,.2f}" in txt


def nueva_pagina(b, **kw):
    ctx = b.new_context(accept_downloads=True, **kw)
    for nombre, ruta in CDN.items():
        if ruta.exists():
            ctx.route(re.compile(r".*/" + re.escape(nombre) + r"$"), (lambda f: (lambda route: route.fulfill(path=f, content_type="application/javascript")))(str(ruta)))
    ctx.route(re.compile(r"https://.*supabase\.co/.*"), lambda route: route.abort())
    ctx.set_default_timeout(15000)
    pg = ctx.new_page()
    errs = []
    pg.on("pageerror", lambda e: errs.append(str(e)))
    return ctx, pg, errs


N_ABRIR = [0]


def abrir(pg, vista="presupuesto"):
    N_ABRIR[0] += 1  # cada apertura es una carga nueva de la página
    pg.goto(URL + "?carga=" + str(N_ABRIR[0]) + "#" + vista)
    pg.wait_for_timeout(1200)


with sync_playwright() as p:
    b = p.chromium.launch()

    # ---------- 1. Arranque y proyecto de ejemplo ----------
    grupo("Arranque y ejemplo")
    ctx, pg, errs = nueva_pagina(b, viewport={"width": 1300, "height": 900})
    abrir(pg)
    d = pg.evaluate("(()=>{const C=calcular(); return {total:C.c.total, cd:C.c.cd, mat:C.mat, mo:C.mo, sc:C.sc, otro:C.otro, aj:C.ajPU, ajR:C.ajRed, ver:String(DATA.version), caps:DATA.caps.length, apus:DATA.apus.length, recs:DATA.recursos.length, cat:DB.cat, tabs:[...document.querySelectorAll('nav.tabs .tab')].map(t=>t.textContent.trim())}})()")
    chk("Sin errores de página al abrir", not errs, errs)
    chk("Total del ejemplo = $50,730.74", d["total"] == 50730.74, d["total"])
    chk("Costo directo del ejemplo = $25,365.37", d["cd"] == 25365.37, d["cd"])
    chk("Catálogo versión 5", d["ver"] == "5", d["ver"])
    chk("14 partidas en el catálogo", d["caps"] == 14, d["caps"])
    chk("3,234 subpartidas (166 estándar y 3,068 de referencia)", d["apus"] == 3234, d["apus"])
    chk("196 recursos (205 − 2 de agua − 9 de acero por qq − Autocad − base de concreto + 2 puertas + bañera + tubo 6x2)", d["recs"] == 196, d["recs"])
    chk("DB.cat = 3 (códigos nuevos)", d["cat"] == 3, d["cat"])
    chk("Pestañas: Proyecto, Subpartidas, Presupuesto, Materiales, PU", d["tabs"] == ["Proyecto", "Subpartidas", "Presupuesto", "Materiales", "PU"], d["tabs"])
    comp = pg.evaluate("(()=>{const C=calcular(); return [C.mat, C.mo, C.sc, C.otro, C.ajEd||0, C.ajRed||0, C.c.cd]})()")
    chk("Conciliación: componentes = costo directo", abs(sum(comp[:6]) - comp[6]) < 0.005, comp)

    # ---------- 2. Precios de recursos ----------
    grupo("Precios de recursos (18 nov 2026)")
    esperado = {
        "bloque15": (0.75, "actualizado"), "bloque10": (0.71, "actualizado"), "cemento": (8.05, "actualizado"),
        "arena": (29.26, "actualizado"), "alambre": (0.62, "actualizado"), "m_tubo_pvc_1_2_u": (2.39, "actualizado"),
        "m_tubo_pvc_4_64psi_ml": (2.02, "actualizado"), "m_lamina_zinc_alum_m": (8.67, "actualizado"),
        "m_panel_thermotecho_2_m": (25.66, "actualizado"), "m_porcelanato_m": (13.42, "actualizado"),
        "albanil": (30.0, "actualizado"), "armador": (30.0, "actualizado"), "carpintero": (30.0, "actualizado"),
        "soldador": (30.0, "actualizado"), "electricista": (30.0, "actualizado"), "fontanero": (30.0, "actualizado"),
        "mo_tablaroquero": (30.0, "actualizado"), "mo_mecanico_de_banco": (30.0, "actualizado"),
        "auxiliar": (20.0, "actualizado"), "obrero": (20.0, "actualizado"), "maestro": (40.0, "actualizado"),
        "mo_topografo": (40.0, "actualizado"), "mo_jardinero": (30.0, "actualizado"),
        "m_puerta_metalica_gladiador_u": (123.01, "actualizado"), "m_banera_fibra_vidrio_u": (308.85, "actualizado"), "m_puerta_corrediza_pvc_vidrio_u": (637.17, "actualizado"), "m_tubo_estructural_6x2_ch16_ml": (7.37, "actualizado"),
        "varilla2_6m": (1.24, "actualizado"), "varilla3_6m": (3.19, "actualizado"), "varilla4_6m": (5.66, "actualizado"),
    }
    app_rec = pg.evaluate("Object.fromEntries(DATA.recursos.map(r=>[r.id,[r.precio,(r.val||{}).estado||'',(r.val||{}).fecha||'',r.unidad,r.prest]]))")
    for rid, (pr, est) in esperado.items():
        x = app_rec.get(rid)
        chk(f"{rid}: ${pr:.2f}, {est}", x and x[0] == pr and x[1] == est, x)
    con_fuente = [rid for rid in esperado if app_rec[rid][2] in ("18 de noviembre de 2026", "19 de noviembre de 2026")]
    chk("Todos los recursos actualizados guardan la fecha de consulta", len(con_fuente) == len(esperado), set(esperado) - set(con_fuente))
    chk("Varillas #5 a #10 por validar", all(app_rec[f"varilla{n}_6m"][1] == "pendiente" for n in range(5, 11)))
    chk("Ningún material por quintal: el acero va por varilla de 6 m", not [k for k, v in app_rec.items() if v[3] == "qq"], [k for k, v in app_rec.items() if v[3] == "qq"])
    chk("Cable THHN #12 (12.15) a $0.74/ml", pg.evaluate("BASE_APU.get('12.15').pdir") == 0.74)
    unidos = pg.evaluate("REC_UNIDOS")
    chk("30 recursos anteriores con su reemplazo (17 repetidos, agua, acero por qq, Autocad y base de concreto)", len(unidos) == 30, len(unidos))
    chk("Ninguna subpartida usa un recurso reemplazado", pg.evaluate("DATA.apus.every(a=>a.lineas.every(l=>!(l[0] in REC_UNIDOS) && BASE_REC.has(l[0])))"))
    chk("Los repetidos ya no aparecen en el catálogo", pg.evaluate("Object.keys(REC_UNIDOS).every(k=>!BASE_REC.has(k))"))
    chk("Agua solo por m³", not [k for k, v in app_rec.items() if k.startswith("agua") and k != "agua"] and app_rec["agua"][3] == "m³", [k for k in app_rec if k.startswith("agua")])
    chk("Sin precios distintos para el mismo nombre de material", pg.evaluate("(()=>{const m={}; for(const r of DATA.recursos){const k=r.nombre.toLowerCase().replace(/\\(por [^)]*\\)/,'').trim()+'|'+r.unidad; (m[k]=m[k]||new Set()).add(r.precio)} return Object.values(m).every(s=>s.size===1)})()"))
    chk("Cable #14 y #10 sin cambio", pg.evaluate("[BASE_APU.get('12.14').pdir, BASE_APU.get('12.16').pdir]") == [0.40, 0.85])

    # ---------- 3. Mano de obra ----------
    grupo("Mano de obra")
    mo = {k: v for k, v in app_rec.items() if REC[k]["tipo"] == "MO"}
    for rid, v in mo.items():
        chk(f"{rid}: prestaciones 57% y unidad día/hombre", v[4] == 57 and v[3] == "día/hombre", v)
    chk("Ningún recurso dice 'jornal'", not [r for r in CAT["recursos"] if "jornal" in (r["nombre"] + r["unidad"]).lower()])
    rend = pg.evaluate("Object.fromEntries(['05.07','05.06'].map(k=>[k, Object.fromEntries(BASE_APU.get(k).lineas.filter(l=>['albanil','auxiliar','armador'].includes(l[0])))]))")
    chk("05.07 albañil 0.125 día/hombre por m² (8 m²/día)", rend["05.07"]["albanil"] == 0.125, rend)
    chk("05.06 albañil 0.125 día/hombre por m² (8 m²/día)", rend["05.06"]["albanil"] == 0.125, rend)

    # ---------- 4. P.U. de las 153 subpartidas estándar, recalculados aparte ----------
    grupo("P.U. de las subpartidas estándar")
    app_pu = pg.evaluate("Object.fromEntries(DATA.apus.filter(a=>!a.id.startsWith('R')).map(a=>[a.id,costoAPU(a).total]))")
    for a in CAT["apus"]:
        if a["id"].startswith("R"):
            continue
        chk(f"{a['id']} P.U. = {pu_py(a):.2f}", app_pu[a["id"]] == pu_py(a), (app_pu[a["id"]], pu_py(a)))
    chk("06.08 ThermoTecho con receta (sin precio global)", pg.evaluate("BASE_APU.get('06.08').pdir === 0 && BASE_APU.get('06.08').lineas.length === 4"))
    chk("06.08 P.U. = 29.94", app_pu["06.08"] == 29.94, app_pu["06.08"])

    # ---------- 5. Cierre, conjunto y fórmulas ----------
    grupo("Cierre, conjunto y fórmulas")
    c1 = pg.evaluate("cierre(1000, {modo:'factor', factor:100, iva:true})")
    chk("Recargo 100%: subtotal 2,000 e IVA 260", c1["sub"] == 2000 and c1["iva"] == 260 and c1["total"] == 2260, c1)
    c2 = pg.evaluate("cierre(1000, {modo:'ind', ind:30, util:10, iva:false})")
    chk("Indirectos 30% + utilidad 10%: 1,430", c2["sub"] == 1430 and c2["iva"] == 0, c2)
    c3 = pg.evaluate("cierre(1000, {modo:'aiu', adm:10, imp:5, ua:10, iva:false})")
    chk("Adm. 10% + imp. 5% + util. 10%: 1,250", c3["sub"] == 1250, c3)
    c4 = pg.evaluate("cierre(333.33, {modo:'factor', factor:33.333, iva:true})")
    chk("Redondeo a centavos en cada paso", c4["factor"] == 111.11 and c4["sub"] == 444.44 and c4["iva"] == 57.78, c4)
    cj = pg.evaluate("(()=>{P().nviv=4; const C=calcular(); return [conjunto(C).cd, C.c.cd, conjunto(C).n]})()")
    chk("Conjunto de 4 viviendas: costo directo × 4", cj[0] == R2(cj[1] * 4) and cj[2] == 4, cj)
    for f, v in [("2*6.6*0.4 + 7.1*0.4", 8.12), ("15x13.8", 207.0), ("(2+3)*4", 20.0), ("10/4", 2.5), ("1,5*2", 3.0), ("-3+5", 2.0)]:
        r = pg.evaluate("s=>evalFormula(s)", f)
        chk(f"Fórmula {f} = {v}", r is not None and abs(r - v) < 1e-9, r)
    for f in ["2+", "abc", "3*(4"]:
        chk(f"Fórmula no válida '{f}' se rechaza", pg.evaluate("s=>evalFormula(s)", f) is None)

    # ---------- 6. Acero por metro lineal ----------
    grupo("Acero por metro lineal")
    ac = pg.evaluate("VARILLAS.map(v=>[v.n, v.apu, BASE_APU.get(v.apu).unidad, BASE_APU.get(v.apu).lineas[0]])")
    for n, apu, u, l0 in ac:
        chk(f"#{n} ({apu}) por ml con 1/6 de varilla de 6 m", u == "ml" and l0[0] == f"varilla{n}_6m" and abs(l0[1] - 1 / 6) < 1e-5, (apu, u, l0))
    kg = 0.994 / 45.36
    pu4 = R2(5.66 / 6 + 0.62 * 2 * kg + 0.6 * kg * 30 * 1.57 + 0.6 * kg * 20 * 1.57)
    chk(f"04.23 (#4) P.U. = {pu4:.2f} por ml", app_pu["04.23"] == pu4, app_pu["04.23"])
    tabla = "ARMADURA\nFamilia y tipo\tLongitud\nVarilla #4\t1200\nVarilla #3\t450.5\n"
    t = pg.evaluate("t=>{const r=nuevaTabla(t,'armadura.txt'); return r.imp.partidas.map((p,i)=>[r.imp.tipos[i].nombre,p.apu,p.unidad,cantidadPorUnidad(r.imp.tipos[i],p.unidad)])}", tabla)
    chk("Tabla de armadura: #4 → 04.23 en ml con 1,200 ml", t[0][1:] == ["04.23", "ml", 1200], t)
    chk("Tabla de armadura: #3 → 04.22 en ml con 450.5 ml", t[1][1:] == ["04.22", "ml", 450.5], t)
    pg.evaluate("ACALC.n=4; ACALC.unidad='varillas'; ACALC.valor='100'; ACALC.desp=5")
    cv = pg.evaluate("convAcero(4, 100, 'varillas', 6, 5)")
    chk("Calculadora: 100 varillas #4 = 600 ml = 13.148 qq", abs(cv["ml"] - 600) < 1e-9 and round(cv["qq"], 3) == 13.148, cv)
    chk("Calculadora: con 5% se compran 105 varillas", cv["varCompra"] == 105, cv["varCompra"])
    pg.evaluate("ir('materiales'); render()")
    pg.wait_for_timeout(300)
    n0 = pg.evaluate("P().manual.length")
    pg.click("[data-aceroadd]")
    pg.wait_for_timeout(300)
    ult = pg.evaluate("P().manual[P().manual.length-1]")
    chk("Calculadora agrega 600 ml de 04.23 al presupuesto", pg.evaluate("P().manual.length") == n0 + 1 and ult["apu"] == "04.23" and ult["cant"] == 600 and ult["unidad"] == "ml", ult)
    mig = pg.evaluate("""(()=>{ const p={revit:[{partidas:[{apu:'04.23',unidad:'qq',ov:89}],tipos:[{}]}], manual:[{apu:'04.23',unidad:'qq',cant:13.148,formula:'x',ov:90}]}; migrarProyecto(p); return [p.revit[0].partidas[0], p.manual[0]] })()""")
    chk("Migración: partida de Revit en qq pasa a ml", mig[0]["unidad"] == "ml" and mig[0]["ov"] is None, mig[0])
    chk("Migración: 13.148 qq de #4 pasan a 600 ml", mig[1]["unidad"] == "ml" and abs(mig[1]["cant"] - 600) < 0.01, mig[1])
    chk("Explosión: la varilla de 6 m se compra entera", pg.evaluate("UNIDADES_ENTERAS.has(BASE_REC.get('varilla4_6m').unidad)"))

    # ---------- 7. Códigos de Partida y Subpartida ----------
    grupo("Códigos de Partida y Subpartida")
    chk("Tabla de equivalencias con 3,221 códigos", pg.evaluate("ANT.size") == 3221)
    for v, n in [("05.01.07", "05.07"), ("05.01.06", "05.06"), ("05.02.01", "05.09"), ("04.06.03", "04.23"), ("06.02.05", "06.08"), ("12.04.05", "12.15"), ("R11.04.298", "R11.0674")]:
        chk(f"{v} → {n}", pg.evaluate("c=>codNuevo(c)", v) == n)
    lc = pg.evaluate("['05.01.07 Pared', '05.07 Pared', '5.7 Pared'].map(s=>leerCodigo(s))")
    chk("COD_PARTIDA viejo se lee como nuevo", lc[0]["cod"] == "05.07" and lc[0]["valido"] and lc[0]["ant"] == "05.01.07", lc[0])
    chk("COD_PARTIDA nuevo se lee igual", lc[1]["cod"] == "05.07" and lc[1]["valido"], lc[1])
    chk("COD_PARTIDA sin ceros no es válido", not lc[2]["valido"], lc[2])
    chk("Siguiente código libre: 05.12, R11.0918, 04.34", pg.evaluate("[nextCode('05'), nextCode('11',true), nextCode('04')]") == ["05.12", "R11.0918", "04.34"])
    viejo = {"v": 4, "cat": 2, "actual": "p1", "rec": {}, "recNew": {}, "tema": "noche", "elab": {}, "desp": {},
             "apuOv": {"05.01.12": {"id": "05.01.12", "cc": "05", "nombre": "Propia", "unidad": "m²", "lineas": [["cemento", 1]], "eq": 0, "pdir": 0, "nuevo": 1}},
             "aprendido": {"MURO X": {"nombre": "Muro X", "apu": "05.01.07", "unidad": "m²", "veces": 2, "fecha": 1}},
             "proyectos": {"p1": {"id": "p1", "nombre": "Viejo", "cfg": {"modo": "factor", "factor": 100, "iva": False}, "nviv": 1, "revit": [],
                                  "manual": [{"apu": "05.01.07", "unidad": "m²", "cant": 10, "formula": "", "desc": "", "pu": None, "ov": None},
                                             {"apu": "05.01.12", "unidad": "m²", "cant": 2, "formula": "", "desc": "", "pu": None, "ov": None},
                                             {"apu": "04.06.03", "unidad": "qq", "cant": 1, "formula": "", "desc": "", "pu": None, "ov": None}], "mod": 1}}}
    ctx2, pg2, errs2 = nueva_pagina(b)
    abrir(pg2)
    clave = pg2.evaluate("KEY")
    # se escribe antes de que cargue la página, para que la app no lo sobrescriba al salir
    ctx2.add_init_script("if(!sessionStorage.getItem('prueba')){sessionStorage.setItem('prueba','1'); localStorage.setItem(" + json.dumps(clave) + ", " + json.dumps(json.dumps(viejo)) + ");}")
    abrir(pg2)
    m = pg2.evaluate("[P().manual.map(x=>[x.apu,x.unidad,x.cant]), Object.keys(DB.apuOv), DB.apuOv[Object.keys(DB.apuOv)[0]].ant, DB.aprendido['MURO X'].apu, DB.cat]")
    chk("Proyecto viejo: 05.01.07 → 05.07", m[0][0][0] == "05.07", m)
    chk("Proyecto viejo: PU propio 05.01.12 → 05.12 con su código anterior", m[1] == ["05.12"] and m[2] == "05.01.12" and m[0][1][0] == "05.12", m)
    chk("Proyecto viejo: acero 04.06.03 en qq → 04.23 en ml", m[0][2][0] == "04.23" and m[0][2][1] == "ml" and abs(m[0][2][2] - 45.63) < 0.01, m[0][2])
    chk("Proyecto viejo: lo aprendido → 05.07", m[3] == "05.07", m)
    chk("PU propio con un material repetido pasa al material único", pg2.evaluate("(()=>{DB.apuOv['05.95']={id:'05.95',cc:'05',nombre:'x',unidad:'m²',lineas:[['acero3_false',1],['arena_arena',2]],nuevo:1}; convertirCodigos(); return DB.apuOv['05.95'].lineas})()") == [["varilla3_6m", 13.5], ["arena", 2]])
    chk("Proyecto viejo: queda en DB.cat 3 sin errores", m[4] == 3 and not errs2, errs2)
    ctx2.close()

    # ---------- 8. Catálogo de la base de datos ----------
    grupo("Catálogo de la base de datos")
    chk("Un catálogo más viejo no reemplaza al de la app", not pg.evaluate("catMasNueva(2, 4)") and pg.evaluate("catMasNueva(5, 4)"))
    conv = pg.evaluate("d=>{ const x=catalogoCodigosNuevos({apus:[{id:'05.01.07'},{id:'R11.04.298'},{id:'05.07'}]}); return x.apus.map(a=>a.id) }", None)
    chk("Catálogo con códigos viejos se convierte", conv == ["05.07", "R11.0674", "05.07"], conv)
    ctx3, pg3, errs3 = nueva_pagina(b)
    cat2 = {"version": 2, "caps": CAT["caps"], "recursos": [dict(r, precio=99) for r in CAT["recursos"]], "apus": CAT["apus"]}
    ctx3.add_init_script("localStorage.setItem('apu-catalogo', " + json.dumps(json.dumps(cat2)) + ");")
    abrir(pg3)
    chk("Caché del navegador con versión 2 se ignora", pg3.evaluate("[CAT.version, BASE_REC.get('cemento').precio]") == ["5", 8.05])
    ctx3.close()

    # ---------- 8b. Datos del 19 nov: PU nuevos, reglas, fuera del presupuesto y acero por varilla ----------
    grupo("PU nuevos, reglas y acero por varilla (19 nov)")
    abrir(pg)
    for cod, pu, u in [("09.09", 164.22, "u"), ("09.10", 719.60, "u"), ("11.18", 391.28, "u"), ("13.08", 61.49, "m³"), ("11.19", 1755.0, "SG")]:
        x = pg.evaluate("c=>[costoAPU(BASE_APU.get(c)).total, BASE_APU.get(c).unidad]", cod)
        chk(f"{cod} P.U. = {pu:.2f} por {u}", x == [pu, u], x)
    eq = pg.evaluate("['09.09','09.10','11.18','13.08'].map(c=>{const a=BASE_APU.get(c), mo=a.lineas.reduce((s,l)=>{const r=BASE_REC.get(l[0]); return s+(r.tipo==='MO'?l[1]*r.precio*(1+(r.prest||0)/100):0)},0); return Math.abs(a.eq-R2(mo*0.05))<0.006})")
    chk("Herramienta menor = 5% de la mano de obra en 09.09, 09.10, 11.18 y 13.08", all(eq), eq)
    chk("Arena puesta en obra: cantera + flete de $40 por camión de 3 m³", "flete $13.33" in pg.evaluate("BASE_REC.get('arena').val.nota"))
    chk("Lámina Zinc Alum de Galvanissa", pg.evaluate("BASE_REC.get('m_lamina_zinc_alum_m').val.fuente") == "Galvanissa")
    chk("Bloque 15 con precio de tienda San Miguel (Freund 806183)", pg.evaluate("BASE_REC.get('bloque15').val.codigo") == "806183")
    chk("Sin Autocad ni base de concreto SG en el catálogo", pg.evaluate("!BASE_REC.has('m_autocad_m') && !BASE_REC.has('m_base_de_concreto_sg') && DATA.apus.every(a=>a.lineas.every(l=>l[0]!=='m_autocad_m' && l[0]!=='m_base_de_concreto_sg'))"))
    chk("R06.0115 (cubierta zinc alum) usa la lámina en lugar de 'Autocad'", pg.evaluate("BASE_APU.get('R06.0115').lineas[0][0]") == "m_lamina_zinc_alum_m")
    chk("Grava y selecto: precio puesto en obra", all("puesto en obra" in pg.evaluate("k=>BASE_REC.get(k).val.nota", k) for k in ["grava", "selecto"]))
    fz = pg.evaluate("[BASE_REC.get('sc_fosa_septica_y_tuberia_pozo_de_absorcion_sg').precio, BASE_REC.get('sc_fosa_septica_y_tuberia_pozo_de_absorcion_sg').val, BASE_APU.get('R11.0258').unidad]")
    chk("Fosa con pozo: $1,755.00 estimado, sin fuente, R11.0258 por SG", fz[0] == 1755.0 and fz[1] == {"estado": "provisional", "nota": "Precio estimado."} and fz[2] == "SG", fz)
    cz = pg.evaluate("R2(['cemento','arena','grava','agua'].reduce((s,k,i)=>s+BASE_REC.get(k).precio*[9.5,0.52,0.52,0.21][i],0))")
    chk(f"Concreto f'c 210 = ${cz:.2f}/m³ con la dosificación de la app (antes $5.93)", pg.evaluate("BASE_REC.get('m_concreto_f_c_210kg_cm2_m').precio") == cz == 116.05, cz)
    chk("Cemento: mismo precio en todas las tiendas", "todas las tiendas" in pg.evaluate("BASE_REC.get('cemento').val.nota"))
    chk("Cable #12: IVA confirmado", "con IVA" in pg.evaluate("BASE_APU.get('12.15').est") and "por confirmar" not in pg.evaluate("BASE_APU.get('12.15').est"))
    chk("Ninguna receta usa acero por quintal", pg.evaluate("DATA.apus.every(a=>a.lineas.every(l=>!/^acero\\d+$/.test(l[0])))"))
    tabla = ("Familia y tipo\tÁrea\tLongitud\tVolumen\tRecuento\nPuerta metálica 0.95x2.10\t\t\t\t3\nPuerta de madera cedro\t\t\t\t1\nPuerta corrediza PVC vidrio 2 hojas\t\t\t\t1\nBañera 1.50\t\t\t\t1\n"
             "Suelo grava\t20\t\t3\t1\nImposta\t5\t10\t\t1\nFosa séptica + pozo\t\t\t\t1\nCama matrimonial\t\t\t\t2\nCodo PVC 4\t\t\t\t6\nTee PVC 4\t\t\t\t2\nMueble de cocina\t\t3\t\t1\n")
    t = pg.evaluate("t=>{const r=nuevaTabla(t,'casa.txt'); return Object.fromEntries(r.imp.partidas.map((p,i)=>[r.imp.tipos[i].nombre,[p.apu,p.unidad,p.excluido||'',!!p.sgPend]]))}", tabla)
    for nom, esp in [("Puerta metálica 0.95x2.10", ["09.09", "u", "", False]), ("Puerta de madera cedro", ["09.01", "u", "", False]), ("Puerta corrediza PVC vidrio 2 hojas", ["09.10", "u", "", False]), ("Bañera 1.50", ["11.18", "u", "", False]),
                     ("Suelo grava", ["13.08", "m³", "", False]), ("Imposta", ["05.09", "m²", "", False]), ("Fosa séptica + pozo", ["11.19", None, "", True]),
                     ("Cama matrimonial", ["", None, "mobiliario", False]), ("Codo PVC 4", ["", None, "accesorio", False]), ("Tee PVC 4", ["", None, "accesorio", False])]:
        chk(f"Regla: '{nom}' → {esp[0] or 'fuera del presupuesto (' + esp[2] + ')'}", t[nom] == esp, t[nom])
    reglas = [("Ventana corredera de 2 hojas 1:V-2", "09.06"), ("Fijo:V-1", "09.07"), ("Puerta corredera, 2 hojas:P-5", "09.10"), ("Portão_metalico:Portão_metalico", "09.04"),
              ("Asiento de WC con cisterna:INODORO", "11.10"), ("Lavabo moderno:LAVABO EMPOTRADO", "11.11"), ("Shower-column:DUCHA", "11.12"),
              ("Lámpara empotrada - LED:LUMINARIA EMPOTRADA 24W", "12.04"), ("M_Toma - 220V:Estándar", "12.06"), ("M_Interruptores de iluminación:Tridireccional", "12.07"),
              ("M_Cuadro de control de accesorios e iluminación - 208V MCB:TABLERO DISTRIBUCION 100A", "12.03"), ("Tipos de tubería:PVC SANITARIO", "11.07"),
              ("Tipos de tubería:PVC SANITARIO 2''", "11.08"), ("Tipos de tubería:PVC AGUA POTABLE", "11.02"), ("Tipos de tubería:PVC A.L.L", "11.16"), ("Canalón:Canalón", "06.09"),
              ("Cubierta básica:THERMO TECHO GFX", "06.08"), ("Cubierta básica:LAMINA ZINCALUM CAL24", "06.04"), ("Techo compuesto:CIELO FALSO PVC", "10.02"), ("Suelo:GRAMA", "13.04"),
              ("Suelo:PORCELANATO 60X60", "07.05"), ("Suelo:PISO ANTIDESLIZANTE", "07.06"), ("Suelo:LOSA DE ENTREPISO VT1-20", "04.19"), ("Hormigón-Viga rectangular:S-F 1", "03.06"),
              ("Hormigón-Viga rectangular:VIGA V-2", "04.08"), ("RHS-Sección hueca rectangular:TUBO ESTRUCTURAL 6*2", "")]
    got = pg.evaluate("ns=>ns.map(n=>sugerirAPU(n))", [n for n, _ in reglas])
    for (nom, esp), g in zip(reglas, got):
        chk(f"Regla del modelo ARGOS: '{nom}' → {esp or 'sin regla'}", g == esp, g)
    chk("El tablero eléctrico no se toma como accesorio de tubería", pg.evaluate("excluirTipo('M_Cuadro de control de accesorios e iluminación - 208V MCB')") == "")
    chk("Área de puertas y ventanas = alto × ancho del IFC", "OverallHeight" in HTML.read_text(encoding="utf-8") and pg.evaluate("FUENTE_IFC.dimensiones") == "Alto × ancho de puertas y ventanas")
    for cod, pu in [("03.13", 51.98), ("04.31", 43.58), ("04.32", 36.49), ("04.33", 26.91), ("07.09", 45.76), ("11.21", 17.0), ("13.09", 27.02), ("06.13", 10.59)]:
        chk(f"Subpartida del modelo ARGOS {cod} = {pu:.2f}", pg.evaluate("c=>costoAPU(BASE_APU.get(c)).total", cod) == pu)
    t2 = pg.evaluate("""t=>{const r=nuevaTabla(t,'argos.txt'); return r.imp.partidas.map((p,i)=>[r.imp.tipos[i].nombre,p.apu,p.origen,!!p.codNoCat,p.excluido||''])}""",
        "COD_PARTIDA\tFamilia y tipo\tLongitud\tÁrea\tRecuento\n04.34\tHormigón-Viga rectangular: VIGA V-1\t10\t\t1\n09.13\tVentana corredera de 2 hojas 1: V-2\t\t4.48\t2\n04.99\tElemento sin regla\t5\t\t1\n07.88\tMesa-Café\t\t\t6\n")
    chk("Código del modelo fuera del catálogo: se propone el APU por el nombre (por confirmar)", t2[0][1:4] == ["04.07", "regla", True] and t2[1][1:4] == ["09.06", "regla", True], t2)
    chk("Código fuera del catálogo y sin regla: queda para crear el APU", t2[2][1:4] == ["", None, True], t2[2])
    chk("Código fuera del catálogo en mobiliario: fuera del presupuesto", t2[3][4] == "mobiliario", t2[3])
    chk("Mueble de cocina no se trata como mobiliario", t["Mueble de cocina"][2] == "", t["Mueble de cocina"])
    pg.evaluate("t=>{cargarTexto(t,'casa.txt')}", tabla)
    pg.wait_for_timeout(300)
    ls = pg.evaluate("lineas().map(l=>l.titulo)")
    chk("Cama, codos y tees no entran al presupuesto ni a los avisos", not any(x in ls for x in ["Cama matrimonial", "Codo PVC 4", "Tee PVC 4"]), ls)
    pg.evaluate("ir('partidas'); render()")
    pg.wait_for_timeout(300)
    txt = pg.inner_text("#revitGroups")
    chk("Subpartidas: aviso 'Fuera del presupuesto' y conteo", "Fuera del presupuesto" in txt and "3 fuera del presupuesto" in txt, txt[:300])
    ref = pg.evaluate("(()=>{const R=P().revit; for(let i=0;i<R.length;i++){const j=R[i].tipos.findIndex(t=>t.nombre==='Puerta metálica 0.95x2.10'); if(j>=0) return 'r:'+i+':'+j;} return ''})()")
    pg.click(f'[data-quitarpres="{ref}"]'); pg.wait_for_timeout(200)
    q = pg.evaluate("r=>{const p=getItem(r); return [p.excluido, p.apu, lineas().some(l=>l.titulo==='Puerta metálica 0.95x2.10')]}", ref)
    chk("Quitar del presupuesto: la subpartida sale del total", q == ["usuario", "", False], q)
    rq = pg.evaluate("r=>{const i=+r.split(':')[1], st=statsTabla(P().revit[i]); return [st.quitadas, st.enPresupuesto + st.quitadas + st.fuera === P().revit[i].partidas.length, REG_COLS.includes('Quitadas a mano'), REG_COLS.includes('Lectura del IFC (s)')]}", ref)
    chk("Registro: cuenta las quitadas a mano y la confirmación solo pide las que quedan en el presupuesto", rq == [1, True, True, True], rq)
    chk("Hojas: las tablas no se encogen (resumen del IFC visible)", ".sheet>*{flex-shrink:0}" in HTML.read_text(encoding="utf-8"))
    pg.click(f'[data-incluir="{ref}"]'); pg.wait_for_timeout(200)
    q = pg.evaluate("r=>{const p=getItem(r); return [p.excluido||'', p.apu, lineas().some(l=>l.titulo==='Puerta metálica 0.95x2.10')]}", ref)
    chk("Volver a incluir: recupera su APU", q == ["", "09.09", True], q)
    mig = pg.evaluate("""(()=>{ const p={revit:[{partidas:[{apu:'',unidad:'u'}],tipos:[{nombre:'Silla plástica'}]}], manual:[]}; migrarProyecto(p); return p.revit[0].partidas[0] })()""")
    chk("Migración: lo ya importado sin APU se marca fuera del presupuesto", mig.get("excluido") == "mobiliario" and mig.get("unidad") is None, mig)
    abrir(pg)
    ex = pg.evaluate("explosion().mat.filter(x=>x.varCompra).map(x=>[x.r.id, x.r.unidad, x.conDesp, x.varCompra, x.qqCompra, x.comercial])")
    chk("Explosión: acero en varillas de 6 m (#2, #3 y #4)", sorted(x[0] for x in ex) == ["varilla2_6m", "varilla3_6m", "varilla4_6m"] and all(x[1] == "varilla" for x in ex), ex)
    kgv = {2: 0.249, 3: 0.560, 4: 0.994}
    ok = all(x[3] == math.ceil(x[2] - 1e-9) and abs(x[4] - x[3] * 6 * kgv[int(x[0][7:-3])] / 45.36) < 1e-9 for x in ex)
    chk("Explosión: varillas enteras y su equivalente en quintales", ok, ex)
    pg.evaluate("ir('materiales'); render()")
    pg.wait_for_timeout(300)
    chk("Materiales a comprar muestra los quintales del acero", re.search(r"\d+ varillas de 6 m = [\d.,]+ qq", pg.inner_text("#v-materiales")) is not None)
    qq = pg.evaluate("R2(BASE_REC.get('varilla4_6m').precio/6/0.994*45.36)")
    chk(f"Tabla de acero: #4 a ${qq:.2f}/qq calculado del precio de la varilla", money_ok(pg.inner_text("#cardAcero"), qq), qq)
    pg.evaluate("SEL='05.07'; ir('biblioteca'); render()")
    pg.wait_for_timeout(300)
    chk("Editor de PU: la varilla muestra ml, kg y qq", re.search(r"ml ≈ [\d.,]+ kg ≈ [\d.,]+ qq por m²", pg.inner_text("#libDetalle")) is not None)
    ov = pg.evaluate("(()=>{DB.apuOv['05.96']={id:'05.96',cc:'05',nombre:'x',unidad:'m²',lineas:[['acero4',0.1],['m_autocad_m',1],['agua_litros',20],['agua_bolsa',1]],nuevo:1}; convertirCodigos(); const r=DB.apuOv['05.96'].lineas; delete DB.apuOv['05.96']; return r})()")
    chk("PU propio: acero en qq pasa a varillas (× 7.6056), agua en litros a m³ y se quitan Autocad y agua en bolsa", ov == [["varilla4_6m", 0.760563], ["agua", 0.02]], ov)
    abrir(pg)

    # ---------- 9. Pantallas ----------
    grupo("Pantallas y nombres")
    for v in ["proyecto", "partidas", "presupuesto", "materiales", "biblioteca"]:
        abrir(pg, v)
        txt = pg.evaluate("document.querySelector('.view.on').innerText")
        chk(f"Pestaña {v}: no dice 'jornal' ni 'capítulo'", not re.search(r"jornal|cap[ií]tulo", txt, re.I), re.findall(r".{20}(?:jornal|cap[ií]tulo).{20}", txt, re.I)[:3])
    abrir(pg, "materiales")
    chk("Tabla de mano de obra con 'Días/hombre'", "Días/hombre" in pg.inner_text("#tblMO"))
    chk("Tabla de acero con precio por varilla y por ml", "Precio por varilla de 6 m" in pg.inner_text("#cardAcero") and "Precio por ml" in pg.inner_text("#cardAcero"))
    abrir(pg, "biblioteca")
    chk("Precios de recursos: 'Precio actualizado' sin fecha", "Precio actualizado" in pg.inner_text("#tblRecursos") and "noviembre 2026" not in pg.inner_text("#tblRecursos") and "18 nov" not in pg.inner_text("#tblRecursos"))
    pg.evaluate("SEL='05.07'; render()")
    chk("Editor de PU: rendimiento en día/hombre", "día/hombre por m²" in pg.inner_text("#libDetalle"), pg.inner_text("#libDetalle")[:200])
    abrir(pg, "presupuesto")
    chk("Presupuesto por partidas", "Presupuesto por partidas" in pg.inner_text("#v-presupuesto"))
    pg.evaluate("P().revit.forEach(i=>i.partidas.forEach(x=>{x.apu='';})); render()")
    av = pg.inner_text("#avisos")
    chk("Aviso resumido de subpartidas fuera del total", "subpartida(s) fuera del total" in av and "Ir a Subpartidas" in av, av[:200])
    chk("Lista plegable cuando son más de 4", pg.locator("#avisos details").count() >= 1)
    abrir(pg)

    # ---------- 10. Tamaño de letra ----------
    grupo("Tamaño de letra")
    for v, px in [("0.9", 13.5), ("1", 15), ("1.15", 17.25), ("1.3", 19.5)]:
        pg.evaluate(f"DB.letra={v}; aplicarTema()")
        chk(f"Letra {v}: texto base {px} px", pg.evaluate("parseFloat(getComputedStyle(document.body).fontSize)") == px)
    pg.evaluate("guardar()")
    pg.wait_for_timeout(300)
    abrir(pg)
    chk("El tamaño se mantiene al recargar", pg.evaluate("parseFloat(getComputedStyle(document.body).fontSize)") == 19.5)
    pg.evaluate("DB.letra=1; aplicarTema(); guardar()")

    # ---------- 11. Borrar PU y materiales ----------
    grupo("Borrar PU y materiales")
    pg.evaluate("DB.apuOv['05.90']={id:'05.90',cc:'05',nombre:'PU de prueba',unidad:'m²',lineas:[],eq:0,pdir:0,nuevo:1}; render()")
    pg.evaluate("OPEN='borrarapu:05.90'; renderSheet()")
    chk("PU propio sin uso: pide confirmación", "¿Borrar" in pg.inner_text("#sheet"))
    pg.click("[data-confirmborrarcat]")
    pg.wait_for_timeout(300)
    chk("PU propio borrado", not pg.evaluate("!!DB.apuOv['05.90']"))
    pg.evaluate("OPEN='borrarapu:05.07'; renderSheet()")
    chk("PU usado en un proyecto: no se borra", "No se puede borrar" in pg.inner_text("#sheet"))
    pg.evaluate("OPEN=null; renderSheet(); SEL='01.01'; ir('biblioteca'); render()")
    chk("Sin cuenta de administrador no aparece 'Borrar del catálogo'", pg.locator("[data-borrarapu]").count() == 0)

    # ---------- 12. PDF y Excel ----------
    grupo("PDF y Excel")
    abrir(pg)
    pg.evaluate("DB.expo = {...(DB.expo||{}), version:'interna', detalle:true, memoria:true, materiales:true, apu:true}; P().cfg.iva=false; guardar()")
    with pg.expect_download(timeout=30000) as dl:
        pg.evaluate("exportar('pdf', true)")
    tot = pg.evaluate("calcular().c.total")
    tot_txt = f"{tot:,.2f}"
    pdf = PdfReader(dl.value.path())
    texto = "\n".join(pg_.extract_text() or "" for pg_ in pdf.pages)
    chk("PDF generado", len(pdf.pages) >= 2, len(pdf.pages))
    chk("PDF: RESUMEN POR PARTIDAS", "RESUMEN POR PARTIDAS" in texto)
    chk("PDF: columna Subpartida", "Subpartida" in texto)
    chk(f"PDF: total igual al de la pantalla (${tot_txt})", tot_txt in texto, tot_txt)
    chk("PDF: mano de obra en días/hombre", "Días/hombre" in texto and not re.search(r"jornal", texto, re.I))
    with pg.expect_download(timeout=30000) as dl2:
        pg.evaluate("exportar('xlsx', true)")
    ruta_x = Path("/tmp") / dl2.value.suggested_filename
    dl2.value.save_as(ruta_x)
    wb = openpyxl.load_workbook(ruta_x, data_only=False)
    todo = []
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            todo.extend(str(v) for v in row if v is not None)
    s = "\n".join(todo)
    chk("Excel generado con varias hojas", len(wb.worksheets) >= 4, wb.sheetnames)
    chk("Excel: columna Subpartida", "Subpartida" in s)
    chk("Excel: mano de obra en días/hombre", "Días/hombre" in s and not re.search(r"jornal", s, re.I))
    chk(f"Excel: total igual al de la pantalla ({tot})", any(abs(float(v) - tot) < 0.005 for v in todo if re.fullmatch(r"-?\d+(\.\d+)?", v)), tot)
    chk("Sin errores de página al exportar", not errs, errs)
    ctx.close()

    # ---------- 13. Celular y vista previa ----------
    grupo("Celular y vista previa")
    ctx4, pg4, errs4 = nueva_pagina(b, viewport={"width": 390, "height": 844}, is_mobile=True, has_touch=True)
    abrir(pg4)
    r = pg4.evaluate("(()=>{const r=document.querySelector('nav.tabs').getBoundingClientRect(); return [r.bottom, innerHeight, document.documentElement.scrollWidth<=innerWidth+1]})()")
    chk("Celular: pestañas abajo", abs(r[0] - r[1]) < 2, r)
    chk("Celular: sin desborde lateral", r[2], r)
    ctx4.close()
    ctx5, pg5, errs5 = nueva_pagina(b)
    pg5.set_content('<iframe id="f" sandbox="allow-scripts allow-same-origin" style="width:1200px;height:800px"></iframe>')
    pg5.evaluate("s=>document.getElementById('f').srcdoc=s", HTML.read_text(encoding="utf-8"))
    pg5.wait_for_timeout(2500)
    fr = pg5.frames[1]
    fr.click("text=Subpartidas")
    pg5.wait_for_timeout(300)
    chk("Vista previa (iframe srcdoc): sin error de replaceState", not [e for e in errs5 if "replaceState" in e], errs5)
    ctx5.close()
    b.close()

ok = sum(1 for r in RES if r[2])
mal = [r for r in RES if not r[2]]
grupos = {}
for g, _, c, _ in RES:
    grupos.setdefault(g, [0, 0])[0 if c else 1] += 1
for g, (a, f) in grupos.items():
    print(f"{g}: {a} correctas{', ' + str(f) + ' fallidas' if f else ''}")
for g, n, _, det in mal:
    print(f"FALLA [{g}] {n}: {det}")
print(f"TOTAL: {ok} correctas, {len(mal)} fallidas, {len(RES)} pruebas")
sys.exit(1 if mal else 0)
