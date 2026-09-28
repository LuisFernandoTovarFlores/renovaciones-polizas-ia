"""Genera una cartera 100% FICTICIA con el mismo formato que produce extraer_caratulas.py
y construye con ella el tablero de demostración (docs/index.html).

    python scripts/generar_demo.py
"""
import os, random, shutil, string, sys
from datetime import date, timedelta
from pathlib import Path
import pandas as pd

RAIZ = Path(__file__).resolve().parent.parent
DEMO = RAIZ / "demo"
random.seed(42)

NOMBRES = ["Ana", "Luis", "María", "Jorge", "Sofía", "Carlos", "Lucía", "Diego", "Elena", "Pablo",
           "Valeria", "Andrés", "Camila", "Ricardo", "Paula", "Héctor", "Daniela", "Tomás", "Regina", "Iván"]
APELLIDOS = ["Ejemplo", "Prueba", "Demo", "Ficticio", "Muestra", "Modelo", "Simulado", "Genérico",
             "Ensayo", "Patrón", "Plantilla", "Borrador"]
EMPRESAS = ["Comercializadora Alfa", "Logística Beta", "Industrias Gamma", "Servicios Delta",
            "Constructora Épsilon", "Distribuidora Zeta", "Clínica Eta", "Transportes Theta"]
ASEG = {"Autos": ["Zurich", "Qualitas", "AXA", "GNP", "Chubb", "HDI"],
        "GMM Individual": ["AXA", "GNP", "Plan Seguro", "Mapfre", "Panamerican"],
        "Hogar": ["Zurich", "GMX", "Chubb"],
        "Daños / Empresarial": ["Chubb", "SURA", "AXA"],
        "GMM Colectivo": ["AXA", "GNP", "MetLife"],
        "Responsabilidad Civil": ["Chubb", "AIG", "GMX"],
        "Transportes": ["Chubb", "ANA", "SURA"]}
TIPO = {"Autos": "Automóvil", "GMM Individual": "Gastos Médicos Mayores Individual",
        "Hogar": "Hogar", "Daños / Empresarial": "Múltiple Empresarial",
        "GMM Colectivo": "Gastos Médicos Mayores Colectivo",
        "Responsabilidad Civil": "Responsabilidad Civil General", "Transportes": "Transporte de Mercancías"}
PRIMA = {"Autos": (6_000, 35_000), "GMM Individual": (15_000, 120_000), "Hogar": (2_500, 12_000),
         "Daños / Empresarial": (20_000, 250_000), "GMM Colectivo": (150_000, 900_000),
         "Responsabilidad Civil": (5_000, 40_000), "Transportes": (15_000, 90_000)}
AUTOS = ["Nissan Versa", "VW Jetta", "Toyota RAV4", "Mazda 3", "Honda CR-V", "Kia Rio", "Ford Ranger"]


def rfc(empresa=False):
    l = "".join(random.choices(string.ascii_uppercase, k=3 if empresa else 4))
    f = f"{random.randint(60, 99):02d}{random.randint(1, 12):02d}{random.randint(1, 28):02d}"
    return "XX" + l[2:] + f + "".join(random.choices(string.ascii_uppercase + string.digits, k=3))


def cartera(n_clientes=260, hoy=date.today()):
    filas = []
    for i in range(n_clientes):
        empresa = random.random() < 0.15
        nom = (random.choice(EMPRESAS) + f" {i}, S.A. de C.V." if empresa
               else f"{random.choice(NOMBRES)} {random.choice(APELLIDOS)} {random.choice(APELLIDOS)}").upper()
        cliente = dict(nombre_asegurado=nom, contratante=nom, rfc_curp=rfc(empresa) if random.random() > .1 else "null",
                       telefono=f"55 0000 {random.randint(1000, 9999)}", correo=f"cliente{i}@ejemplo.com")
        ramos = (random.sample(["Daños / Empresarial", "GMM Colectivo", "Responsabilidad Civil", "Transportes", "Autos"],
                               random.randint(1, 3)) if empresa else
                 random.sample(["Autos", "Autos", "GMM Individual", "Hogar"], random.choice([1, 1, 1, 2, 3])))
        for ramo in set(ramos):
            aseg = random.choice(ASEG[ramo]); poliza = str(random.randint(10**6, 10**8))
            ini = hoy - timedelta(days=random.randint(0, 365) + 365 * random.randint(1, 5))
            prima = random.uniform(*PRIMA[ramo])
            bien = (f"{random.choice(AUTOS)} {random.randint(2015, 2025)}" if ramo == "Autos"
                    else "Domicilio de ejemplo" if ramo in ("Hogar", "Daños / Empresarial") else "Plan de ejemplo")
            while ini <= hoy + timedelta(days=30):
                fin = ini.replace(year=ini.year + 1)
                if fin < hoy and random.random() < 0.12:  # póliza que no se renueva
                    filas.append(dict(fila_base(cliente, aseg, ramo, poliza, bien, ini, fin, prima)))
                    break
                filas.append(fila_base(cliente, aseg, ramo, poliza, bien, ini, fin, prima))
                ini, prima = fin, prima * random.uniform(1.02, 1.15)
    return pd.DataFrame(filas)


def fila_base(c, aseg, ramo, poliza, bien, ini, fin, prima):
    return {**c, "no_poliza": poliza, "poliza_anterior": "", "aseguradora": aseg, "tipo_seguro": TIPO[ramo],
            "bien_asegurado": bien, "vigencia_inicio": ini.isoformat(), "vigencia_fin": fin.isoformat(),
            "prima_neta": round(prima, 2), "moneda": "MXN",
            "forma_pago": random.choice(["Anual", "Contado", "Semestral", "Mensual", "Trimestral"]),
            "grupo": c["nombre_asegurado"], "ruta_relativa": "", "archivo_origen": "demo.pdf",
            "fecha_procesado": "demo"}


if __name__ == "__main__":
    DEMO.mkdir(exist_ok=True)
    d = cartera()
    anio = pd.to_datetime(d.vigencia_inicio).dt.year
    with pd.ExcelWriter(DEMO / "renovaciones.xlsx") as xw:
        d[anio >= 2025].to_excel(xw, sheet_name="Vigentes", index=False)
        d[anio < 2025].to_excel(xw, sheet_name="Pólizas anteriores", index=False)
    os.environ["DATA_DIR"] = str(DEMO)
    sys.path.insert(0, str(RAIZ / "src"))
    import analitica as a
    dd = a.marcar_renovaciones(a.cargar()); a.a_sqlite(dd); a.tablero(dd)
    (RAIZ / "docs").mkdir(exist_ok=True)
    shutil.copy(a.HTML, RAIZ / "docs" / "index.html")
    print(f"Demo: {len(d)} pólizas ficticias -> docs/index.html")
