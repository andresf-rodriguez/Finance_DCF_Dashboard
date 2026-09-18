import json
import os
import re
import sys
import tomllib

import requests
import pandas as pd

SECRETS = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".streamlit", "secrets.toml")


def user_agent():
    """La SEC exige identificarse en el User-Agent ("Nombre Apellido correo@dominio.com").
    Se busca en la variable de entorno SEC_USER_AGENT y, si no, en .streamlit/secrets.toml."""
    valor = os.environ.get("SEC_USER_AGENT", "").strip()
    if not valor and os.path.exists(SECRETS):
        with open(SECRETS, "rb") as f:
            valor = str(tomllib.load(f).get("SEC_USER_AGENT", "")).strip()
    if valor:
        return valor
    print(
        "ERROR: falta el User-Agent para la SEC (exige identificarse con nombre y correo).\n"
        "Configuralo de una de estas dos formas:\n"
        "  1) Copia .streamlit/secrets.toml.example a .streamlit/secrets.toml y pon tus datos.\n"
        '  2) Exporta la variable de entorno:  export SEC_USER_AGENT="Tu Nombre tu-correo@ejemplo.com"',
        file=sys.stderr,
    )
    sys.exit(1)


# ============ CONFIGURACION (lo unico que cambias) ============
TICKER = sys.argv[1].strip().upper() if len(sys.argv) > 1 else "META"  # python extraer_datos.py AAPL
if not re.fullmatch(r"[A-Z0-9.\-]{1,6}", TICKER):
    print(f"ERROR: ticker no valido: {TICKER!r}. Usa solo letras, numeros, punto o guion, maximo 6 caracteres "
          "(ej.: META, BRK.B).", file=sys.stderr)
    sys.exit(1)
ANIOS = 5
H = {"User-Agent": user_agent()}
EPS_MINIMO = 1.0  # solo se deriva acciones = utilidad / EPS si |EPS| >= 1 (el redondeo a centavos lo hace inestable)
# ==============================================================

# Cada dato puede venir con distintas etiquetas segun la empresa o el año.
# Se prueban en orden: la primera que tenga valor para ese año es la que se usa.
# Una opcion ("cociente", A, B) calcula A / B (ej. utilidad neta / EPS = acciones).
ETIQUETAS = {
    "Ingresos": ["RevenueFromContractWithCustomerExcludingAssessedTax", "Revenues", "SalesRevenueNet"],
    "CFO": ["NetCashProvidedByUsedInOperatingActivities"],
    "Capex": ["PaymentsToAcquirePropertyPlantAndEquipment", "PaymentsToAcquireProductiveAssets"],
    "Arrendamientos financieros": ["FinanceLeasePrincipalPayments"],
    "SBC": ["ShareBasedCompensation", "AllocatedShareBasedCompensationExpense"],
    "Caja": ["CashAndCashEquivalentsAtCarryingValue"],
    "Valores negociables": ["MarketableSecuritiesCurrent", "AvailableForSaleSecuritiesDebtSecuritiesCurrent", "ShortTermInvestments"],
    "Deuda": ["LongTermDebt", "LongTermDebtNoncurrent"],
    "Acciones diluidas": [
        "WeightedAverageNumberOfDilutedSharesOutstanding",
        ("cociente", "NetIncomeLoss", "EarningsPerShareDiluted"),
        "WeightedAverageNumberOfSharesOutstandingBasic",
        ("cociente", "NetIncomeLoss", "EarningsPerShareBasic"),
        "CommonStockSharesOutstanding",
    ],
}

AJUSTES = []  # reexpresiones aplicadas (splits) para el reporte y el JSON
AVISOS = []   # cosas raras que conviene revisar


def hechos_anuales(gaap, etiqueta):
    """Lista [(fecha_fin, hecho)] con datos anuales de 10-K de una etiqueta, y su unidad."""
    if etiqueta not in gaap:
        return [], None
    unidades = gaap[etiqueta]["units"]
    unidad = next((u for u in ("USD", "shares", "USD/shares") if u in unidades), None)
    if unidad is None:
        return [], None
    hechos = []
    for x in unidades[unidad]:
        if x.get("form") != "10-K":
            continue
        fin = pd.Timestamp(x["end"])
        # Datos de flujo (ingresos, CFO...) tienen fecha de inicio: exigir ~1 año
        if "start" in x:
            dias = (fin - pd.Timestamp(x["start"])).days
            if not 350 <= dias <= 380:
                continue
        hechos.append((fin, x))
    return hechos, unidad


def mas_reciente(hechos):
    """Si el mismo año aparece en varios reportes, quedarse con el mas reciente."""
    mejores = {}
    for fin, x in hechos:
        if fin not in mejores or x["filed"] > mejores[fin]["filed"]:
            mejores[fin] = x
    return {fin: x["val"] for fin, x in mejores.items()}


def encadenar(hechos, etiqueta):
    """Para acciones y cifras por accion: parte del 10-K mas reciente y va agregando
    años de 10-K anteriores. Si un 10-K viejo comparte años con la base y todos sus
    valores difieren por el mismo factor (ej. x10 tras un split), ese factor se
    aplica a los años que solo estan en el 10-K viejo. Asi toda la serie queda en
    la base mas reciente aunque el split haya sido despues de parte del historico."""
    por_reporte = {}
    for fin, x in hechos:
        r = por_reporte.setdefault(x["accn"], {"filed": x["filed"], "vals": {}})
        r["vals"][fin] = x["val"]
    reportes = sorted(por_reporte.values(), key=lambda r: r["filed"], reverse=True)
    base = dict(reportes[0]["vals"])
    for rep in reportes[1:]:
        comunes = [f for f in rep["vals"] if f in base and rep["vals"][f] and base[f]
                   and (rep["vals"][f] > 0) == (base[f] > 0)]
        factor = 1.0
        if comunes:
            cocientes = sorted(base[f] / rep["vals"][f] for f in comunes)
            mediana = cocientes[len(cocientes) // 2]
            consistente = cocientes[-1] / cocientes[0] < 1.02
            if consistente and abs(mediana - 1) > 0.2:
                factor = mediana
        nuevos = [f for f in rep["vals"] if f not in base]
        if nuevos and comunes and not consistente:
            AVISOS.append({"etiqueta": etiqueta, "anios": sorted(f.year for f in nuevos),
                           "msg": f"{etiqueta}: el 10-K del {rep['filed']} no cuadra con el mas reciente "
                                  f"(cocientes {cocientes[0]:.3f} a {cocientes[-1]:.3f}); se uso sin ajustar."})
        for f in nuevos:
            base[f] = rep["vals"][f] * factor
        if nuevos and factor != 1.0:
            AJUSTES.append({"etiqueta": etiqueta, "anios": sorted(f.year for f in nuevos),
                            "factor": round(factor, 4), "base_10k": reportes[0]["filed"]})
    return base


def relevantes(lista, anios, etiquetas):
    """Filtra ajustes/avisos a los años de la tabla y a las etiquetas que de verdad se usaron."""
    return [x for x in lista if x["etiqueta"] in etiquetas and set(x["anios"]) & set(anios)]


def valores_anuales(gaap, etiqueta):
    """Devuelve {fecha_fin: valor} solo con datos anuales de reportes 10-K.
    USD: se toma el 10-K mas reciente (los balances solo solapan un año, no hay con que contrastar).
    Acciones y por accion: se encadenan los 10-K para dejar todo en la misma base (splits)."""
    hechos, unidad = hechos_anuales(gaap, etiqueta)
    if not hechos:
        return {}
    if unidad == "USD":
        return mas_reciente(hechos)
    return encadenar(hechos, etiqueta)


def resolver(gaap, opcion):
    """Convierte una opcion de ETIQUETAS en (nombre, {fecha: valor}, es_aproximado)."""
    if isinstance(opcion, tuple):
        _, num, den = opcion
        n, d = valores_anuales(gaap, num), valores_anuales(gaap, den)
        vals = {f: n[f] / d[f] for f in n if f in d and abs(d[f]) >= EPS_MINIMO}
        return f"{num} / {den}", vals, True
    return opcion, valores_anuales(gaap, opcion), False


# 1. Buscar el CIK (numero de la empresa en la SEC)
r = requests.get("https://www.sec.gov/files/company_tickers.json", headers=H, timeout=30)
r.raise_for_status()
#    Las acciones con clase se escriben BRK.B en bolsa, pero la SEC las lista como BRK-B: se prueban ambas formas.
variantes = list(dict.fromkeys([TICKER.replace(".", "-"), TICKER, TICKER.replace("-", ".")]))
por_ticker = {v["ticker"]: v["cik_str"] for v in r.json().values()}
simbolo_sec = next((t for t in variantes if t in por_ticker), None)
cik = por_ticker.get(simbolo_sec)
if simbolo_sec and simbolo_sec != TICKER:
    print(f"Nota: la SEC lista {TICKER} como {simbolo_sec}.")
if cik is None:
    print(f"El ticker {TICKER} no está en la lista de la SEC; puede que no cotice en EE. UU. o presente "
          "formulario 20-F en vez de 10-K.", file=sys.stderr)
    sys.exit(1)

# 2. Bajar todos los datos financieros
r = requests.get(f"https://data.sec.gov/api/xbrl/companyfacts/CIK{cik:010d}.json", headers=H, timeout=60)
r.raise_for_status()
gaap = r.json()["facts"]["us-gaap"]

# 3. Fechas de cierre de los ultimos años fiscales (tomadas del CFO) y ultimo 10-K presentado
hechos_cfo, _ = hechos_anuales(gaap, "NetCashProvidedByUsedInOperatingActivities")
fechas = sorted({fin for fin, _ in hechos_cfo})[-ANIOS:]
ultimo_10k = max((x for fin, x in hechos_cfo if fin == fechas[-1]), key=lambda x: x["filed"])

# 4. Armar la tabla
tabla = {}
usadas = {}
aproximados = []
for dato, opciones in ETIQUETAS.items():
    series = [resolver(gaap, op) for op in opciones]
    fila = {}
    for fecha in fechas:
        for nombre, valores, aprox in series:
            if fecha in valores:
                fila[fecha.year] = valores[fecha] / 1e9  # en miles de millones
                usadas.setdefault(dato, set()).add(nombre)
                if aprox:
                    aproximados.append({"dato": dato, "anio": fecha.year, "fuente": nombre})
                break
        else:
            fila[fecha.year] = None
    tabla[dato] = fila

df = pd.DataFrame(tabla).T
df.loc["Arrendamientos financieros"] = df.loc["Arrendamientos financieros"].fillna(0)
df.loc["FCF"] = df.loc["CFO"] - df.loc["Capex"] - df.loc["Arrendamientos financieros"]
df.loc["FCF - SBC"] = df.loc["FCF"] - df.loc["SBC"]

# Quedarse solo con los ajustes/avisos que tocan los años de la tabla y etiquetas usadas
anios_tabla = [f.year for f in fechas]
etq_usadas = {e for nombres in usadas.values() for n in nombres for e in n.split(" / ")}
ajustes = []
for aj in relevantes(AJUSTES, anios_tabla, etq_usadas):
    previo = next((a for a in ajustes if a["etiqueta"] == aj["etiqueta"] and a["factor"] == aj["factor"]), None)
    if previo:
        previo["anios"] = sorted(set(previo["anios"]) | set(aj["anios"]))
    else:
        ajustes.append(dict(aj))
avisos = [av["msg"] for av in relevantes(AVISOS, anios_tabla, etq_usadas)]

# Revisar la serie de acciones: NaN o saltos > 50% entre años suelen ser splits sin ajustar
acc = df.loc["Acciones diluidas"]
for anio, v in acc.items():
    if pd.isna(v):
        avisos.append(f"Acciones diluidas {anio}: sin dato en ninguna etiqueta.")
saltos = acc.pct_change()
for anio, s in saltos.items():
    if pd.notna(s) and abs(s) > 0.5:
        avisos.append(f"Acciones diluidas {anio}: cambio de {s:+.0%} frente al año anterior. ¿Split sin ajustar?")

# 5. Mostrar y guardar
pd.set_option("display.width", 200)
print(f"\n{TICKER} - cifras en miles de millones (acciones: en miles de millones de acciones)\n")
print(df.round(2))

print("\nEtiquetas usadas:")
for dato in ETIQUETAS:
    print(f"  {dato}: {', '.join(sorted(usadas.get(dato, ['NO ENCONTRADA'])))}")
for a in aproximados:
    print(f"  ~ {a['dato']} {a['anio']}: aproximado como {a['fuente']}")
for aj in ajustes:
    anios_txt = ", ".join(str(a) for a in aj["anios"] if a in anios_tabla)
    print(f"  Reexpresion {aj['etiqueta']} ({anios_txt}): x{aj['factor']:g} para dejarlo en la base del 10-K del {aj['base_10k']}")

faltantes = df.isna().sum().sum()
if faltantes:
    print(f"\nOJO: faltan {faltantes} datos. Revisa las filas con NaN.")
for av in avisos:
    print(f"OJO: {av}")

df.to_csv(f"datos_{TICKER}.csv")
meta = {
    "ticker": TICKER,
    "filed": ultimo_10k["filed"],          # fecha de presentacion del ultimo 10-K
    "cierre": fechas[-1].strftime("%Y-%m-%d"),
    "fuentes": {d: sorted(u) for d, u in usadas.items()},
    "aproximados": aproximados,
    "ajustes": ajustes,
    "avisos": avisos,
}
with open(f"datos_{TICKER}.json", "w") as f:
    json.dump(meta, f, indent=2, ensure_ascii=False)
print(f"\nGuardado en datos_{TICKER}.csv y datos_{TICKER}.json")
print("El Excel del DCF se descarga desde la app (boton 'Descargar Excel').")
