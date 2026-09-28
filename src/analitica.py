"""Fase 1-2: limpia renovaciones.xlsx -> SQLite + tablero HTML (calendario y retención).
Corre local, sin API. Volver a correr cada vez que se actualice el Excel."""
import json, os, re, sqlite3, sys, unicodedata
from datetime import date
from pathlib import Path
import pandas as pd
from dotenv import load_dotenv

RAIZ = Path(__file__).resolve().parent.parent
load_dotenv(RAIZ / ".env")
WORK_DIR = Path(os.getenv("DATA_DIR", RAIZ / "datos"))
EXCEL = Path(sys.argv[1]) if len(sys.argv) > 1 else WORK_DIR / "renovaciones.xlsx"
DB = WORK_DIR / "cartera.db"
HTML = WORK_DIR / "tablero_renovaciones.html"
TIPO_CAMBIO = {"MXN": 1.0, "USD": 18.5, "UDI": 8.5}  # ajustar a mano
VENTANA_RENOVACION = 60  # días de tolerancia para considerar que una póliza se renovó
HOY = pd.Timestamp(date.today())

RAMOS = [  # (regex sobre tipo_seguro sin acentos, ramo) — el orden importa
    (r"fianza", "Fianzas"),
    (r"moto", "Motos"),
    (r"flotilla", "Autos Flotilla"),
    (r"auto|vehic|camion|amplia|motorizate|movilidad", "Autos"),
    (r"colectiv|grupo|salud de grupo", "GMM Colectivo"),
    (r"gastos medicos|salud|medica|golden|linea azul|plan seguro", "GMM Individual"),
    (r"accidente|viajes", "Accidentes Personales"),
    (r"vida|capitaliza|persona clave", "Vida"),
    (r"hogar|habitt|hipoteca|incendio contenidos", "Hogar"),
    (r"transport|carga|mercanc", "Transportes"),
    (r"responsabilidad|\brc\b|rcpa|rcmm|rc ", "Responsabilidad Civil"),
    (r"empresarial|pyme|multipoliza|riesgos|danos|gasolinera|estacion|incendio", "Daños / Empresarial"),
]


def sin_acentos(s):
    return "".join(c for c in unicodedata.normalize("NFD", str(s)) if unicodedata.category(c) != "Mn").lower()


def ramo_de(t):
    t = sin_acentos(t)
    for rx, r in RAMOS:
        if re.search(rx, t):
            return r
    return "Otros"


def forma_de(f):
    f = sin_acentos(f)
    for k in ("mensual", "trimestral", "semestral"):
        if k in f:
            return k.capitalize()
    return "Anual / Contado" if f and f != "nan" else ""


def clave_cliente(r):
    rfc = re.sub(r"[^A-Z0-9]", "", str(r.rfc_curp).upper())
    if rfc and rfc != "NULL" and len(rfc) >= 10:
        return rfc[:10]
    nom = sin_acentos(r.cliente)
    nom = re.sub(r"\b(s\.?a\.?|de|c\.?v\.?|s\. de r\.l\.|sapi|sc|el|la|los)\b|[^a-z ]", " ", nom)
    return "N:" + " ".join(nom.split())


def cargar():
    x = pd.read_excel(EXCEL, sheet_name=None, dtype=str, keep_default_na=False)
    d = pd.concat([x["Vigentes"], x["Pólizas anteriores"]], ignore_index=True)
    d["cliente"] = d.contratante.where(d.contratante.str.strip() != "", d.nombre_asegurado).str.strip().str.upper()
    d["cliente_id"] = d.apply(clave_cliente, axis=1)
    # nombre más reciente por cliente
    d["ini"] = pd.to_datetime(d.vigencia_inicio, errors="coerce")
    d["fin"] = pd.to_datetime(d.vigencia_fin, errors="coerce")
    d = d.dropna(subset=["ini", "fin"]).sort_values("ini")
    d["cliente"] = d.groupby("cliente_id").cliente.transform("last")
    d["ramo"] = d.tipo_seguro.map(ramo_de)
    d["forma"] = d.forma_pago.map(forma_de)
    d["moneda"] = d.moneda.replace("", "MXN")
    prima = pd.to_numeric(d.prima_neta.str.replace(",", ""), errors="coerce")
    d["prima_mxn"] = (prima * d.moneda.map(TIPO_CAMBIO).fillna(1)).round(2)
    return d


def marcar_renovaciones(d):
    """Una póliza se renovó si el mismo cliente tiene otra del mismo ramo que inicia cerca de su fin."""
    d = d.reset_index(drop=True)
    d["renovada"] = pd.NA
    d["id_renovacion"] = pd.NA
    for (_, _), g in d.groupby(["cliente_id", "ramo"]):
        for i, r in g.iterrows():
            if r.fin > HOY:
                continue
            sig = g[(g.index != i) & ((g.ini - r.fin).abs().dt.days <= VENTANA_RENOVACION)]
            if len(sig):  # si hay varias, preferir mismo número de póliza o mismo bien
                mejor = sig[sig.no_poliza == r.no_poliza]
                d.at[i, "id_renovacion"] = (mejor if len(mejor) else sig).index[0]
                d.at[i, "renovada"] = 1
            else:
                d.at[i, "renovada"] = 0
    d["tiene_sucesora"] = d.index.isin(d.id_renovacion.dropna().astype(int))
    return d


def a_sqlite(d):
    con = sqlite3.connect(DB)
    cli = (d.sort_values("ini").groupby("cliente_id")
           .agg(cliente=("cliente", "last"), rfc_curp=("rfc_curp", "last"),
                telefono=("telefono", lambda s: next((v for v in s[::-1] if v), "")),
                correo=("correo", lambda s: next((v for v in s[::-1] if v), "")))
           .reset_index())
    cli.to_sql("clientes", con, if_exists="replace", index=False)
    cols = ["cliente_id", "no_poliza", "poliza_anterior", "aseguradora", "ramo", "tipo_seguro",
            "bien_asegurado", "vigencia_inicio", "vigencia_fin", "prima_neta", "moneda", "prima_mxn",
            "forma", "renovada", "id_renovacion", "grupo", "ruta_relativa"]
    p = d[cols].copy(); p.insert(0, "id", d.index)
    p.to_sql("polizas", con, if_exists="replace", index=False)
    con.executescript("""
    DROP VIEW IF EXISTS vigentes;
    CREATE VIEW vigentes AS
      SELECT p.*, c.cliente, c.rfc_curp, c.telefono, c.correo,
             CAST(julianday(p.vigencia_fin) - julianday('now') AS INT) AS dias_para_vencer
      FROM polizas p JOIN clientes c USING(cliente_id)
      WHERE date(p.vigencia_fin) >= date('now');
    DROP VIEW IF EXISTS retencion;
    CREATE VIEW retencion AS
      SELECT aseguradora, ramo, strftime('%Y', vigencia_fin) AS anio_vencimiento,
             COUNT(*) AS vencidas, SUM(renovada) AS renovadas,
             ROUND(100.0 * SUM(renovada) / COUNT(*), 1) AS pct_retencion
      FROM polizas WHERE renovada IS NOT NULL GROUP BY 1, 2, 3;
    """)
    con.commit(); con.close()


def tablero(d):
    vig = d[d.fin >= HOY].copy()
    vig["dias"] = (vig.fin - HOY).dt.days
    filas = [dict(cliente=r.cliente, rfc=r.rfc_curp, poliza=r.no_poliza, aseg=r.aseguradora, ramo=r.ramo,
                  bien=r.bien_asegurado[:60], fin=r.fin.strftime("%Y-%m-%d"), dias=int(r.dias),
                  prima=None if pd.isna(r.prima_mxn) else float(r.prima_mxn), forma=r.forma,
                  tel=r.telefono, correo=r.correo) for r in vig.itertuples()]
    ven = d[d.renovada.notna() & (d.fin >= "2024-01-01")].copy()
    ven["renovada"] = ven.renovada.astype(int)
    perdidas = [dict(cliente=r.cliente, poliza=r.no_poliza, aseg=r.aseguradora, ramo=r.ramo,
                     fin=r.fin.strftime("%Y-%m-%d"),
                     prima=None if pd.isna(r.prima_mxn) else float(r.prima_mxn))
                for r in ven[ven.renovada == 0].sort_values("fin", ascending=False).itertuples()]
    ret = [dict(aseg=r.aseguradora, ramo=r.ramo, anio=r.fin.year, ren=int(r.renovada)) for r in ven.itertuples()]
    datos = dict(hoy=HOY.strftime("%Y-%m-%d"), vigentes=filas, retencion=ret, perdidas=perdidas,
                 sin_prima=int(vig.prima_mxn.isna().sum()), tc=TIPO_CAMBIO)
    plantilla = (Path(__file__).parent / "tablero_plantilla.html").read_text(encoding="utf-8")
    HTML.write_text(plantilla.replace("/*DATOS*/null", json.dumps(datos, ensure_ascii=False)), encoding="utf-8")


if __name__ == "__main__":
    d = marcar_renovaciones(cargar())
    a_sqlite(d)
    tablero(d)
    vig = d[d.fin >= HOY]
    print(f"Pólizas: {len(d)} | clientes: {d.cliente_id.nunique()} | vigentes hoy: {len(vig)} "
          f"({vig.cliente_id.nunique()} clientes, prima neta ${vig.prima_mxn.sum():,.0f})")
    print("Ramos:", d.ramo.value_counts().to_dict())
    r = d[d.renovada.notna()]
    print(f"Retención histórica: {r.renovada.astype(int).mean():.1%} de {len(r)} vencidas")
    print("->", DB, "\n->", HTML)
