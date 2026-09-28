"""
Base de RENOVACIONES a partir de carátulas de pólizas (PDF / escaneos).

Cómo se usa:  python src/extraer_caratulas.py
  1ª corrida: extrae el texto (gratis, local), te muestra el costo estimado y
              te pregunta si envía el lote a la API (Batch = 50% más barato).
  Siguientes corridas: revisa si el lote ya terminó; si sí, descarga los
              resultados y actualiza el Excel. Si faltan archivos nuevos
              (porque agregaste carátulas), prepara otro lote solo con ellos.
Nunca reenvía un archivo que ya se procesó o que está en un lote pendiente.
"""
import os
import re
import json
import hashlib
from pathlib import Path
from datetime import datetime, date

from dotenv import load_dotenv
import pandas as pd
import anthropic
from anthropic.types.message_create_params import MessageCreateParamsNonStreaming
from anthropic.types.messages.batch_create_params import Request

RAIZ = Path(__file__).resolve().parent.parent
load_dotenv(RAIZ / ".env")

# ---------------------------------------------------------------------------
# CONFIGURACIÓN (rutas en .env; ver .env.example)
# ---------------------------------------------------------------------------

ROOT_DIR = Path(os.getenv("CARATULAS_DIR", RAIZ / "caratulas"))
WORK_DIR = Path(os.getenv("DATA_DIR", RAIZ / "datos"))
EXCEL_PATH = WORK_DIR / "renovaciones.xlsx"
ESTADO_PATH = WORK_DIR / "renovaciones_estado.json"   # lotes enviados y resultados
TEXTOS_PATH = WORK_DIR / "renovaciones_textos.json"   # caché de texto extraído
ERROR_LOG_PATH = WORK_DIR / "renovaciones_errores.log"

MODELO = "claude-haiku-4-5"
PRECIO_IN, PRECIO_OUT = 1.00, 5.00   # USD por millón de tokens (Batch cobra la mitad)
PAGINAS = 2                  # solo la carátula; las demás páginas son condiciones
MAX_CHARS = 6000
LIMITE = None                # p. ej. 30 para una prueba; None = todos
ANIO_VIGENTE = 2025          # vigencia que inicia en 2025 o después -> hoja "Vigentes"

# Opcional: carpeta con los binarios de poppler/tesseract si no están en el PATH
if os.getenv("OCR_BIN_DIR"):
    os.environ["PATH"] = os.environ["OCR_BIN_DIR"] + os.pathsep + os.environ["PATH"]

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp"}
PDF_EXT = ".pdf"

CAMPOS = [
    "nombre_asegurado", "contratante", "rfc_curp", "no_poliza", "poliza_anterior",
    "aseguradora", "tipo_seguro", "bien_asegurado", "vigencia_inicio",
    "vigencia_fin", "prima_neta", "moneda", "forma_pago", "telefono", "correo",
]
META = ["grupo", "ruta_relativa", "archivo_origen", "fecha_procesado"]

SCHEMA = {
    "type": "object",
    "properties": {
        "es_caratula_poliza": {"type": "boolean"},
        **{c: {"type": "string"} for c in CAMPOS},
    },
    "required": ["es_caratula_poliza"] + CAMPOS,
    "additionalProperties": False,
}

SYSTEM = """Extraes datos de carátulas de pólizas de seguros en México para una base
de control de RENOVACIONES. El texto puede venir de OCR y tener errores.

Reglas:
- es_caratula_poliza: true solo si el documento es la carátula/especificación de
  una póliza. false si es aviso de cobro, recibo, cotización, endoso, factura,
  identificación, condiciones generales u otro documento.
- Pólizas colectivas o de grupo: NO listes asegurados ni certificados. Usa la
  empresa/contratante como nombre_asegurado y contratante.
- contratante: quien contrata y paga la póliza (puede ser distinto del asegurado).
- rfc_curp: RFC del contratante (o del asegurado si no hay contratante); si no
  hay RFC pero sí CURP, pon la CURP. Es muy importante; sin guiones ni espacios.
- bien_asegurado: autos -> "marca modelo año, placas X, serie Y" (lo que venga);
  hogar/daños -> domicilio o ubicación; gastos médicos/vida -> nombre del plan.
- poliza_anterior: número de póliza que esta renueva, solo si el documento lo dice.
- vigencia_inicio y vigencia_fin: formato AAAA-MM-DD.
- prima_neta: SOLO el monto etiquetado "Prima Neta" (o variantes), como número
  sin comas ni signo de pesos. NO la prima total, NI con recargos, derechos o
  IVA. Si no viene desglosada, deja "".
- moneda: MXN o USD.
- Cualquier dato que no aparezca: "" (cadena vacía). No adivines."""


# ---------------------------------------------------------------------------
# UTILERÍAS
# ---------------------------------------------------------------------------


def cargar_json(path, defecto):
    if path.exists():
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    return defecto


def guardar_json(path, data):
    tmp = path.with_suffix(".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    tmp.replace(path)


def log_error(ruta, error):
    with open(ERROR_LOG_PATH, "a", encoding="utf-8") as f:
        f.write(f"{datetime.now()}\t{ruta}\t{error}\n")


def firma(path: Path) -> str:
    st = path.stat()
    return f"{st.st_size}|{int(st.st_mtime)}"


def custom_id(ruta: str) -> str:
    return hashlib.md5(ruta.encode("utf-8")).hexdigest()


def grupo_de(ruta: str) -> str:
    """Carpeta que identifica cliente + bien (ej. 'Cliente X (Auto 2019) - FEB').
    Las carpetas 'AA...' agrupan varios clientes, ahí se usa la subcarpeta."""
    partes = Path(ruta).parts
    if partes[0].upper().startswith("AA") and len(partes) > 2:
        return f"{partes[0]}\\{partes[1]}"
    return partes[0]


# ---------------------------------------------------------------------------
# EXTRACCIÓN DE TEXTO (local, sin costo)
# ---------------------------------------------------------------------------


def extraer_texto(path: Path) -> str:
    import pdfplumber
    import pytesseract
    from pdf2image import convert_from_path
    from PIL import Image

    if path.suffix.lower() == PDF_EXT:
        texto = ""
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages[:PAGINAS]:
                texto += (page.extract_text() or "") + "\n"
        if len(texto.strip()) < 30:
            imgs = convert_from_path(str(path), dpi=200, first_page=1, last_page=PAGINAS)
            texto = "\n".join(pytesseract.image_to_string(i, lang="spa") for i in imgs)
        return texto
    img = Image.open(path)
    return pytesseract.image_to_string(img, lang="spa")


# ---------------------------------------------------------------------------
# LOTES (Batch API)
# ---------------------------------------------------------------------------


def revisar_lotes(client, estado) -> bool:
    """Descarga resultados de lotes terminados. Devuelve True si queda alguno en proceso."""
    en_proceso = False
    for lote in estado["lotes"]:
        if lote.get("descargado"):
            continue
        b = client.messages.batches.retrieve(lote["id"])
        if b.processing_status != "ended":
            c = b.request_counts
            print(f"Lote {lote['id']} en proceso: {c.succeeded} listos, {c.processing} pendientes.")
            en_proceso = True
            continue
        ok = err = 0
        tin = tout = 0
        for r in client.messages.batches.results(lote["id"]):
            ruta = lote["archivos"].get(r.custom_id)
            if ruta is None:
                continue
            if r.result.type != "succeeded":
                log_error(ruta, f"Lote: {r.result.type}")
                err += 1
                continue
            msg = r.result.message
            tin += msg.usage.input_tokens
            tout += msg.usage.output_tokens
            try:
                texto = next(b.text for b in msg.content if b.type == "text")
                data = json.loads(texto)
            except Exception as e:
                log_error(ruta, f"Respuesta inválida: {e}")
                err += 1
                continue
            data["fecha_procesado"] = datetime.now().strftime("%Y-%m-%d %H:%M")
            estado["resultados"][ruta] = data
            ok += 1
        costo = (tin * PRECIO_IN + tout * PRECIO_OUT) / 1e6 / 2
        lote["descargado"] = True
        lote["costo_usd"] = round(costo, 4)
        guardar_json(ESTADO_PATH, estado)
        print(f"Lote {lote['id']} descargado: {ok} ok, {err} con error. Costo real: ${costo:.4f} USD")
    return en_proceso


def preparar_lote(client, estado, textos):
    archivos = sorted(
        f for f in ROOT_DIR.rglob("*")
        if f.is_file() and f.suffix.lower() in IMAGE_EXTS | {PDF_EXT}
    )
    en_lote = {r for l in estado["lotes"] if not l.get("descargado") for r in l["archivos"].values()}
    pendientes = [
        f for f in archivos
        if str(f.relative_to(ROOT_DIR)) not in estado["resultados"]
        and str(f.relative_to(ROOT_DIR)) not in en_lote
    ]
    if LIMITE:
        pendientes = pendientes[:LIMITE]
    print(f"Documentos en carpeta: {len(archivos)} | ya procesados: {len(estado['resultados'])} "
          f"| por enviar: {len(pendientes)}")
    if not pendientes:
        return

    requests, mapa, chars = [], {}, 0
    for i, f in enumerate(pendientes, 1):
        ruta = str(f.relative_to(ROOT_DIR))
        fir = firma(f)
        if textos.get(ruta, {}).get("firma") != fir:
            print(f"  [{i}/{len(pendientes)}] extrayendo texto: {ruta}")
            try:
                textos[ruta] = {"firma": fir, "texto": extraer_texto(f)[:MAX_CHARS]}
            except Exception as e:
                log_error(ruta, f"Extracción: {e}")
                continue
            if i % 25 == 0:
                guardar_json(TEXTOS_PATH, textos)
        texto = textos[ruta]["texto"]
        if not texto.strip():
            log_error(ruta, "Sin texto extraído (documento ilegible)")
            continue
        cid = custom_id(ruta)
        mapa[cid] = ruta
        chars += len(texto) + len(SYSTEM)
        requests.append(Request(
            custom_id=cid,
            params=MessageCreateParamsNonStreaming(
                model=MODELO,
                max_tokens=1024,
                system=SYSTEM,
                messages=[{"role": "user", "content": f"Archivo: {f.name}\n---\n{texto}\n---"}],
                output_config={"format": {"type": "json_schema", "schema": SCHEMA}},
            ),
        ))
    guardar_json(TEXTOS_PATH, textos)
    if not requests:
        return

    # ~3 caracteres por token en español con números; 250 tokens de respuesta
    est = (chars / 3 * PRECIO_IN + len(requests) * 250 * PRECIO_OUT) / 1e6 / 2
    print(f"\nListos para enviar {len(requests)} documentos. Costo estimado: ~${est:.2f} USD (Batch).")
    if input("¿Enviar a la API? (s/n): ").strip().lstrip("﻿").lower() != "s":
        print("No se envió nada.")
        return
    try:
        lote = client.messages.batches.create(requests=requests)
    except anthropic.APIStatusError as e:
        print(f"La API rechazó el lote ({e.status_code}). Si es por saldo, recarga créditos "
              f"en console.anthropic.com y vuelve a correr el script.\n{e}")
        return
    estado["lotes"].append({"id": lote.id, "creado": datetime.now().isoformat(), "archivos": mapa})
    guardar_json(ESTADO_PATH, estado)
    print(f"Lote enviado: {lote.id}. Suele tardar menos de 1 hora (máx. 24 h). "
          f"Vuelve a correr el script más tarde para descargar los resultados.")


# ---------------------------------------------------------------------------
# EXCEL
# ---------------------------------------------------------------------------


def a_fecha(s):
    try:
        return datetime.strptime(str(s).strip()[:10], "%Y-%m-%d").date()
    except ValueError:
        return None


def a_numero(s):
    s = re.sub(r"[^\d.]", "", str(s or ""))
    try:
        return float(s) if s else None
    except ValueError:
        return None


ASEGURADORAS = [  # (patrón en el nombre, nombre corto) — el orden importa
    (r"santander", "Zurich Santander"), (r"zurich", "Zurich"), (r"qualitas|quálitas", "Qualitas"),
    (r"\bgnp\b|nacional provincial", "GNP"), (r"\baxa\b", "AXA"), (r"chubb|\bace\b", "Chubb"),
    (r"mapfre", "Mapfre"), (r"\bsura\b", "SURA"), (r"\bhdi\b", "HDI"), (r"banorte", "Banorte"),
    (r"inbursa", "Inbursa"), (r"metlife", "MetLife"), (r"\bbupa\b", "BUPA"),
    (r"monterrey|new york life", "Seguros Monterrey"), (r"pan.?american|palig", "Panamerican"),
    (r"plan seguro", "Plan Seguro"), (r"\bgmx\b|grupo mexicano de seguros", "GMX"),
    (r"\baig\b", "AIG"), (r"\ba\.?n\.?a\.?\b|ana compa", "ANA"), (r"primero", "Primero Seguros"),
    (r"allianz", "Allianz"), (r"atlas", "Seguros Atlas"), (r"afirme", "Afirme"), (r"\bhir\b", "HIR"),
    (r"general de seguros", "General de Seguros"), (r"sofimex", "Sofimex"), (r"el potosi|potosí", "El Potosí"),
]


def normaliza_aseguradora(s):
    if pd.isna(s) or not str(s).strip():
        return None
    t = str(s).lower()
    for patron, corto in ASEGURADORAS:
        if re.search(patron, t):
            return corto
    return str(s).strip()


def escribir_excel(estado):
    filas, no_caratulas = [], []
    for ruta, d in estado["resultados"].items():
        if not (ROOT_DIR / ruta).exists():  # archivo movido o borrado de la carpeta
            continue
        fila = {c: (d.get(c) or None) for c in CAMPOS}
        fila.update(grupo=grupo_de(ruta), ruta_relativa=ruta, archivo_origen=Path(ruta).name,
                    fecha_procesado=d.get("fecha_procesado"))
        # Algunas pólizas traen un recibo como primera página y el modelo las
        # marca como "no carátula"; si el nombre del archivo dice póliza/carátula
        # y sí hay número de póliza y vigencia, se rescatan.
        nombre = Path(ruta).name.lower()
        rescatable = (re.search(r"pol|caratula|carátula|emision|emisión", nombre)
                      and not re.search(r"recibo|rec\d|aviso|pago|nota|endoso|end_|cot|propuesta|tarjeta|notice", nombre)
                      and d.get("no_poliza") and d.get("vigencia_inicio"))
        (filas if d.get("es_caratula_poliza", True) or rescatable else no_caratulas).append(fila)
    if not filas:
        return
    df = pd.DataFrame(filas, columns=CAMPOS + META)
    df["vigencia_inicio"] = df["vigencia_inicio"].map(a_fecha)
    df["vigencia_fin"] = df["vigencia_fin"].map(a_fecha)
    df["prima_neta"] = df["prima_neta"].map(a_numero)
    df["forma_pago"] = df["forma_pago"].str.strip().str.capitalize()
    df["aseguradora"] = df["aseguradora"].map(normaliza_aseguradora)
    df["rfc_curp"] = df["rfc_curp"].map(
        lambda s: re.sub(r"[\s\-]", "", str(s)).upper() if pd.notna(s) and str(s).strip() else "null")

    # Duplicados: misma póliza y misma vigencia (p. ej. versiones "Correcta"/"Final")
    df = df.sort_values("fecha_procesado").drop_duplicates(
        subset=["no_poliza", "vigencia_inicio"], keep="last")

    hoy = date.today()
    df["mes_vencimiento"] = df["vigencia_fin"].map(lambda d: d.strftime("%m-%b") if d else None)
    df["dias_para_vencer"] = df["vigencia_fin"].map(lambda d: (d - hoy).days if d else None)
    ultima = df.dropna(subset=["vigencia_inicio"]).groupby("grupo")["vigencia_inicio"].transform("max")
    df["es_ultima_renovacion"] = df["vigencia_inicio"].eq(ultima.reindex(df.index))

    anio = df["vigencia_inicio"].map(lambda d: d.year if d else None)
    vigentes = df[anio >= ANIO_VIGENTE].sort_values("vigencia_fin")
    anteriores = df[anio < ANIO_VIGENTE].sort_values(["grupo", "vigencia_inicio"])

    # Hoja "Revisar": documentos que no parecen carátula y pólizas sin vigencia
    revisar = pd.concat([
        pd.DataFrame(no_caratulas, columns=CAMPOS + META).assign(motivo="No parece carátula"),
        df[anio.isna()][CAMPOS + META].assign(motivo="Sin fecha de vigencia"),
    ], ignore_index=True)
    revisar = revisar[["motivo"] + CAMPOS + META].astype(str).replace({"None": "", "nan": "", "NaT": ""})

    with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as xw:
        vigentes.to_excel(xw, sheet_name="Vigentes", index=False)
        anteriores.to_excel(xw, sheet_name="Pólizas anteriores", index=False)
        if len(revisar):
            revisar.to_excel(xw, sheet_name="Revisar", index=False)
        for ws in xw.book.worksheets:
            ws.freeze_panes = "B2"
            ws.auto_filter.ref = ws.dimensions
            for col in ws.columns:
                ancho = max(len(str(c.value or "")) for c in col[:200])
                ws.column_dimensions[col[0].column_letter].width = min(max(ancho, 10), 45)
    print(f"Excel actualizado: {EXCEL_PATH}  (Vigentes: {len(vigentes)}, "
          f"Anteriores: {len(anteriores)}, Revisar: {len(revisar)})")


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------


def main():
    if not os.environ.get("ANTHROPIC_API_KEY"):
        print("Falta ANTHROPIC_API_KEY en el archivo .env")
        return
    if not ROOT_DIR.exists():
        print(f"No encuentro {ROOT_DIR}. Configura CARATULAS_DIR en .env")
        return
    WORK_DIR.mkdir(parents=True, exist_ok=True)
    client = anthropic.Anthropic()
    estado = cargar_json(ESTADO_PATH, {"lotes": [], "resultados": {}})
    textos = cargar_json(TEXTOS_PATH, {})

    en_proceso = revisar_lotes(client, estado)
    if en_proceso:
        print("Hay un lote en proceso; espera a que termine antes de enviar otro.")
    else:
        preparar_lote(client, estado, textos)
    escribir_excel(estado)
    gastado = sum(l.get("costo_usd", 0) for l in estado["lotes"])
    print(f"Gasto acumulado en lotes descargados: ${gastado:.2f} USD")


if __name__ == "__main__":
    main()

