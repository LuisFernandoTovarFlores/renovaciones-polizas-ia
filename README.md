# Renovaciones de pólizas con IA

Pipeline que convierte ~3,000 documentos de pólizas de seguros (PDFs y escaneos, en carpetas desordenadas) en una **base de datos de renovaciones** y un **tablero interactivo** para una agencia de seguros en México.

**[Ver tablero de demostración →](docs/index.html)** *(datos 100% ficticios)*

## El problema
Una agencia guardaba años de carátulas de pólizas en carpetas por cliente, mezcladas con recibos, endosos, cotizaciones y condiciones generales. No existía una base central para saber **qué vence cada mes, cuánta prima neta está en juego y qué clientes se renovaron o se perdieron**.

## La solución

```
PDFs / escaneos ──► 1. Curaduría local ──► 2. Texto (pdfplumber / OCR) ──► 3. Claude Haiku 4.5 (Batch API)
                                                                                     │  JSON estructurado
                                            5. Tablero HTML ◄── 4. Limpieza + SQLite ◄┘
```

1. **Curaduría sin costo:** clasificador por reglas (nombre de archivo, palabras clave, estructura) que separa carátulas de recibos, endosos y cotizaciones. Nada se borra: lo descartado se mueve con registro reversible.
2. **Extracción de texto local:** `pdfplumber` para PDFs digitales; `Tesseract` OCR en español para escaneos. Solo las primeras páginas (la carátula) para minimizar tokens.
3. **Extracción estructurada con LLM:** `claude-haiku-4-5` vía **Message Batches API** (50 % de descuento) con **structured outputs** (JSON Schema) → 15 campos por póliza: contratante, RFC, número de póliza, aseguradora, ramo, bien asegurado, vigencia, **prima neta** (no total), forma de pago, contacto. El modelo también decide si el documento es realmente una carátula.
4. **Normalización y modelo de datos:** aseguradoras y ~170 variantes de tipo de seguro unificadas en 13 ramos; clientes agrupados por RFC; deduplicación; conversión de moneda; detección de renovaciones (misma cliente + ramo con inicio ±60 días del fin anterior). Salida a Excel y **SQLite** (tablas `clientes`, `polizas`; vistas `vigentes`, `retencion`).
5. **Tablero** HTML autocontenido (Chart.js): calendario de vencimientos a 30/60/90 días, retención por aseguradora y ramo, concentración de cartera, oportunidades de venta cruzada y pólizas no renovadas. Exporta a CSV.

## Resultados (proyecto real)
| | |
|---|---|
| Documentos revisados | ~3,300 |
| No-carátulas descartadas automáticamente | ~1,900 |
| Pólizas extraídas | ~1,430 |
| **Costo total de API** | **≈ $3.60 USD** |

Decisiones de costo: simulación previa de costo entre modelos (Haiku vs Sonnet vs Opus), piloto con muestra antes del lote completo, Batch API, límite de páginas y caché de texto para no reenviar nada.

## Estructura
```
src/extraer_caratulas.py   # texto + OCR + lote a Claude + Excel
src/analitica.py           # limpieza, renovaciones, SQLite y tablero
src/tablero_plantilla.html # tablero (los datos se inyectan como JSON)
scripts/generar_demo.py    # cartera ficticia para la demo
docs/index.html            # tablero de demostración
```

## Uso
```bash
pip install -r requirements.txt      # + Tesseract (idioma spa) y Poppler instalados
cp .env.example .env                 # API key y rutas
python src/extraer_caratulas.py      # 1ª vez: estima costo y pregunta antes de enviar el lote
python src/extraer_caratulas.py      # después: descarga resultados y genera el Excel
python src/analitica.py              # SQLite + tablero
python scripts/generar_demo.py       # demo con datos ficticios
```

## Privacidad
Este repositorio **no contiene datos reales**. `.gitignore` excluye documentos, Excel, bases y cachés. Los datos de la demo son generados aleatoriamente (nombres, RFC, teléfonos y correos inventados).

## Stack
Python · pandas · pdfplumber · Tesseract OCR · Anthropic API (Claude Haiku 4.5, Batch, Structured Outputs) · SQLite · Chart.js
