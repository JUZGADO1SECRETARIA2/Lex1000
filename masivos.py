# -*- coding: utf-8 -*-
import os
import sys
import argparse
import time
from typing import List, Dict, Optional
from multiprocessing import Process

# -----------------------
# Path del proyecto
# -----------------------
PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

# -----------------------
# Utils
# -----------------------
try:
    from utils_mini import (
        autenticar_google_sheets,
        configurar_selenium,
        iniciar_sesion,
        abrir_menu_masivos_documentos_digitales,
        masivo_confirmar_seleccion_final,
        seleccionar_modelo,
        masivo_marcar_a_la_firma,
        estacionar_mouse,
        seleccionar_modelo_por_texto,
    )
except ImportError:
    print("📥 utils_mini no encontrado localmente, descargando desde GitHub...")
    import urllib.request
    import importlib.util
    
    # Descargar utils_mini desde GitHub
    url = "https://raw.githubusercontent.com/JUZGADO1SECRETARIA2/Lex1000/main/utils_mini.py"
    response = urllib.request.urlopen(url)
    utils_code = response.read().decode('utf-8')
    
    # Crear módulo temporal
    spec = importlib.util.spec_from_loader('utils_mini', loader=None)
    utils_mini = importlib.util.module_from_spec(spec)
    exec(utils_code, utils_mini.__dict__)
    sys.modules['utils_mini'] = utils_mini
    
    # Re-importar las funciones
    from utils_mini import (
        autenticar_google_sheets,
        configurar_selenium,
        iniciar_sesion,
        abrir_menu_masivos_documentos_digitales,
        masivo_confirmar_seleccion_final,
        seleccionar_modelo,
        masivo_marcar_a_la_firma,
        estacionar_mouse,
        seleccionar_modelo_por_texto,
    )
    print("✅ utils_mini cargado desde GitHub")

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, StaleElementReferenceException,
    ElementNotInteractableException, NoSuchElementException
)

# =======================
# CONFIG (con defaults por ENV; overridable por CLI)
# =======================
SHEET_NAME = os.getenv("SHEET_NAME", "Juzgado Personal")
SHEET_TAB  = os.getenv("SHEET_TAB",  "masivos")
FILA_INICIO = int(os.getenv("FILA_INICIO", 3))
CHROME_DRIVER_PATH = os.getenv("CHROME_DRIVER_PATH")
KEEP_BROWSER_OPEN = True  # puede override por CLI

# input de la pantalla de "Documentos digitales"
NAME_CODIGOBARRAS = 'despachoDocumentosMasivoDecorate:searchFilters:search1:filterFormVisible:codigoBarras'

# =======================
# OPCIONES (columna + modelo)
# =======================
OPCIONES: List[Dict] = [
    {"id": "A - Primera Liquidacion", "col_letra": "A", "clave": "CONTESTA TRASLADO - PASE A RESOLVER EJECUCION -ETQ-",
     "modelo_texto": "CONTESTA TRASLADO - PASE A RESOLVER EJECUCION -ETQ-"},
    {"id": "C - Liquidacion Actualizada", "col_letra": "C", "clave": "CONTESTA TRASLADO - PASE A RESOLVER LIQUIDACION ACTUALIZADA -ETQ-",
     "modelo_texto": "CONTESTA TRASLADO - PASE A RESOLVER LIQUIDACION ACTUALIZADA -ETQ-"},
    {"id": "E - Embargos", "col_letra": "E", "clave": "PASE A RESOLVER PEDIDO DE EMBARGO -ETQ-",
     "modelo_texto": "PASE A RESOLVER PEDIDO DE EMBARGO -ETQ-"},
    {"id": "G - TRALIQ", "col_letra": "G", "clave": "TRALIQ -TRASLADO DE LIQUIDACION -ETQ-",
     "modelo_texto": "TRALIQ -TRASLADO DE LIQUIDACION -ETQ-"},
    {"id": "I - Impugnacion", "col_letra": "I", "clave": "TRASLADO DE LAS IMPUGNACIONES",
     "modelo_texto": "TRASLADO DE LAS IMPUGNACIONES"},
    {"id": "K - Impugna + Previo", "col_letra": "K", "clave": "TRASLADO DE LAS IMPUGNACIONES + PREVIO",
     "modelo_texto": "TRASLADO DE LAS IMPUGNACIONES + PREVIO"},
    {"id": "M - Honorarios Intimacion", "col_letra": "M", "clave": "HONORARIOS INTIMACION BAJO APERCIBIMIENTO DE EMBARGO",
     "modelo_texto": "HONORARIOS INTIMACION BAJO APERCIBIMIENTO DE EMBARGO"},
    {"id": "O - TRF Capital", "col_letra": "O", "clave": "SE LIBRA OFICIO DE TRANSFERENCIA -ETQ-",
     "modelo_texto": "SE LIBRA OFICIO DE TRANSFERENCIA -ETQ-"},
    {"id": "P - TRF Honorarios", "col_letra": "P", "clave": "CUMPLASE DEOX HONORARIOS",
     "modelo_texto": "CUMPLASE DEOX HONORARIOS"},
]

# =======================
# HELPERS LOCALES
# =======================
def letra_a_indice(letra: str) -> int:
    """Convierte letra de columna (A=1, B=2, ...) a índice 1-indexed."""
    letra = (letra or "").strip().upper()
    if not letra or not letra.isalpha():
        raise ValueError(f"Letra inválida de columna: {letra}")
    idx = 0
    for ch in letra:
        idx = idx * 26 + (ord(ch) - ord('A') + 1)
    return idx

def normalizar_expediente(expediente: str) -> str:
    return (expediente or "").replace("/", "").strip()

def seleccionar_mejor_opcion(filas) -> bool:
    """Tilda checkbox evitando 'incidente' / 'recurso de queja' y eligiendo el más 'corto'."""
    mejor = None
    mejor_len = float('inf')
    for fila in filas:
        try:
            txt = fila.find_element(By.XPATH, ".//td[4]").text
            tnorm = txt.replace("CSS ", "").strip().lower()
            if ("incidente" in tnorm) or ("recurso de queja" in tnorm):
                continue
            if len(tnorm) < mejor_len:
                mejor = fila
                mejor_len = len(tnorm)
        except Exception:
            continue
    if not mejor:
        return False
    try:
        cb = mejor.find_element(By.XPATH, ".//input[@type='checkbox']")
        if not cb.is_selected():
            cb.click()
        return True
    except Exception:
        return False

def confirmar_seleccion(driver, wait: WebDriverWait) -> None:
    """Primer 'Confirmar selección' (de la grilla)."""
    xpaths = [
        '//span[normalize-space(text())="Confirmar selección"]/ancestor::div[@role="button"]',
        '//input[@value="Confirmar selección" or @title="Confirmar selección"]',
        '//button[normalize-space(text())="Confirmar selección"]',
    ]
    for xp in xpaths:
        try:
            btn = wait.until(EC.element_to_be_clickable((By.XPATH, xp)))
            driver.execute_script("arguments[0].click();", btn)
            time.sleep(0.3)
            try: estacionar_mouse(driver)
            except Exception: pass
            return
        except Exception:
            continue
    raise TimeoutException("No se encontró el botón 'Confirmar selección' (grilla).")

def xpath_modelo_clickable(texto: str) -> str:
    """Arma el XPATH robusto de la sugerencia clickeable para el texto dado."""
    frag = texto.replace('"', '\\"')
    return ('//*[@id="despacho:modeloDecoration:modeloSuggestion"]'
            f'//*[contains(normalize-space(.), "{frag}")]'
            '/ancestor-or-self::*[self::a or self::span or self::div or self::td or self::tr or self::li][1]')

def leer_columna_por_letra(worksheet, letra: str, fila_inicio: int) -> List[str]:
    """Lee valores de Google Sheets desde la columna 'letra', empezando en fila_inicio."""
    idx = letra_a_indice(letra)
    vals = worksheet.col_values(idx)
    return [v for v in vals[fila_inicio - 1:] if (v or "").strip()]

# =======================
# FLUJO POR OPCIÓN
# =======================
def ejecutar_opcion(conf: Dict, *, sheet_name: str, sheet_tab: str, fila_inicio: int,
                    chromedriver_path: Optional[str], keep_browser_open: bool):
    nombre = conf["id"]
    col = conf["col_letra"]
    clave = conf["clave"]
    modelo_txt = conf["modelo_texto"]

    print(f"\n>>> [{nombre}] Iniciando…  (Columna {col}, clave='{clave}', modelo='{modelo_txt}')")

    # 1) Leer expedientes de Sheets
    hoja = autenticar_google_sheets(sheet_name, sheet_tab)
    expedientes = leer_columna_por_letra(hoja, col, fila_inicio)
    print(f"    · {len(expedientes)} expedientes en columna {col} (desde fila {fila_inicio})")
    if not expedientes:
        print("    · No hay expedientes. Fin de esta opción.")
        return

    # 2) Selenium
    driver, wait, _actions = configurar_selenium(chromedriver_path)
    try:
        iniciar_sesion(driver, wait)
        input_field = abrir_menu_masivos_documentos_digitales(driver, wait, verificar_input=True)
        try:
            estacionar_mouse(driver, input_field)
        except Exception:
            pass

        # 3) Loteo
        for exp in expedientes:
            exp_norm = normalizar_expediente(exp)
            print(f"    - {exp} → {exp_norm}")
            try:
                try:
                    input_field.clear()
                except (StaleElementReferenceException, ElementNotInteractableException, NoSuchElementException):
                    input_field = wait.until(EC.element_to_be_clickable((By.NAME, NAME_CODIGOBARRAS)))
                    try: input_field.click()
                    except Exception: pass
                    try: estacionar_mouse(driver, input_field)
                    except Exception: pass
                    input_field.clear()

                input_field.send_keys(exp_norm)
                input_field.send_keys(Keys.ENTER)
                time.sleep(1.5)  # ← Aumentado ligeramente

                wait.until(EC.presence_of_element_located((By.XPATH, "//tr[contains(@class, 'rich-table-row')]")))
                filas = driver.find_elements(By.XPATH, "//tr[contains(@class, 'rich-table-row')]")

                if filas:
                    try: estacionar_mouse(driver, filas[0])
                    except Exception: pass

                if not filas:
                    print("      ⚠️ Sin filas para este código.")
                    continue

                if len(filas) == 1:
                    try:
                        cb = filas[0].find_element(By.XPATH, ".//input[@type='checkbox']")
                        if not cb.is_selected():
                            cb.click()
                    except Exception:
                        print("      ⚠️ No se pudo tildar la única fila.")
                else:
                    ok = seleccionar_mejor_opcion(filas)
                    if not ok:
                        print("      ⚠️ No se pudo elegir opción válida (¿todas eran incidente/queja?).")
            except Exception as e:
                print(f"      ❌ Error con {exp}: {type(e).__name__} - {e}")

        # 4) Confirmaciones + Modelo + Firma + Estado
        confirmar_seleccion(driver, wait)
        masivo_confirmar_seleccion_final(driver, wait)
        try: estacionar_mouse(driver)
        except Exception: pass

        seleccionar_modelo_por_texto(
            driver, wait,
            clave=clave,
            texto_objetivo=modelo_txt,
            frag_fallback=modelo_txt
        )

        try: estacionar_mouse(driver)
        except Exception: pass

        masivo_marcar_a_la_firma(driver, wait, marcar=True)

        try:
            from utils_mini import seleccionar_estado_proyecto
            seleccionar_estado_proyecto(driver, wait)
        except Exception:
            pass

        print(f"    ✅ [{nombre}] Finalizado OK.")
    finally:
        if not keep_browser_open:
            try:
                driver.quit()
            except Exception:
                pass

# =======================
# INPUT / PARSER
# =======================
def parse_ops_string(ops: str, max_n: int) -> List[int]:
    """
    Convierte "1,3,5" -> [0,2,4] (índices 0-based).
    Ignora tokens inválidos o fuera de rango.
    """
    if not ops:
        return []
    out: List[int] = []
    for tok in ops.replace(" ", "").split(","):
        if tok.isdigit():
            j = int(tok)
            if 1 <= j <= max_n:
                out.append(j - 1)
    # eliminar duplicados preservando orden
    seen = set()
    filtered = []
    for x in out:
        if x not in seen:
            filtered.append(x); seen.add(x)
    return filtered

def pedir_opciones_interactivo() -> List[int]:
    """Modo clásico por consola (si no se usa --ops)."""
    print("\n=== Opciones disponibles (columna + modelo) ===")
    for i, o in enumerate(OPCIONES, 1):
        print(f"{i}. {o['id']}  —  clave='{o['clave']}'  —  modelo='{o['modelo_texto']}'")
    sel = input("\nElegí opciones a ejecutar (ej: 1,3,5): ").strip()
    return parse_ops_string(sel, len(OPCIONES))

# =======================
# MAIN
# =======================
def ejecutar_agente_masivos(
    *,
    ops_indices: Optional[List[int]] = None,
    sheet_name: str = SHEET_NAME,
    sheet_tab: str = SHEET_TAB,
    fila_inicio: int = FILA_INICIO,
    chromedriver_path: Optional[str] = CHROME_DRIVER_PATH,
    keep_browser_open: bool = KEEP_BROWSER_OPEN,
):
    """
    Ejecuta el agente de Masivos. Si ops_indices es None, pregunta por consola.
    Si ops_indices tiene 1 elemento → ejecución directa.
    Si tiene >1 → arranca un proceso por opción (en paralelo).
    """
    print(">>> Iniciando agente_masivos (multi-opción).")

    if ops_indices is None:
        idxs = pedir_opciones_interactivo()
    else:
        idxs = ops_indices

    if not idxs:
        print("No seleccionaste opciones. Fin.")
        return

    if len(idxs) == 1:
        ejecutar_opcion(
            OPCIONES[idxs[0]],
            sheet_name=sheet_name,
            sheet_tab=sheet_tab,
            fila_inicio=fila_inicio,
            chromedriver_path=chromedriver_path,
            keep_browser_open=keep_browser_open,
        )
        print("\n✅ Listo.")
        return

    # >1 opción → ejecutar en paralelo (un Chrome por opción)
    procs = []
    for i in idxs:
        p = Process(
            target=ejecutar_opcion,
            args=(OPCIONES[i],),
            kwargs=dict(
                sheet_name=sheet_name,
                sheet_tab=sheet_tab,
                fila_inicio=fila_inicio,
                chromedriver_path=chromedriver_path,
                keep_browser_open=keep_browser_open,
            ),
        )
        p.daemon = False
        p.start()
        procs.append(p)
        time.sleep(0.5)  # ← Aumentado el desfase de arranque

    for p in procs:
        p.join()

    print("\n✅ Todas las opciones seleccionadas finalizaron.")

def build_arg_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="Agente Masivos — ejecución por opciones.")
    p.add_argument("--ops", type=str, default=None,
                   help='Opciones a ejecutar separadas por comas (ej: "1,3,5"). '
                        'Si se omite, se pedirá por consola.')
    p.add_argument("--sheet-name", type=str, default=SHEET_NAME)
    p.add_argument("--sheet-tab", type=str, default=SHEET_TAB)
    p.add_argument("--start-row", type=int, default=FILA_INICIO,
                   help="Fila de inicio (1-indexed). Default toma de FILA_INICIO.")
    p.add_argument("--chromedriver", type=str, default=CHROME_DRIVER_PATH)
    grp = p.add_mutually_exclusive_group()
    grp.add_argument("--keep-browser-open", action="store_true", help="Dejar Chrome abierto al final.")
    grp.add_argument("--close-browser", action="store_true", help="Cerrar Chrome al final.")
    return p

if __name__ == "__main__":
    # Nota: En Windows, multiprocessing usa 'spawn', por eso mantenemos el guard.
    parser = build_arg_parser()
    args = parser.parse_args()

    # Resolver keep_browser_open
    keep_open = KEEP_BROWSER_OPEN
    if args.keep_browser_open:
        keep_open = True
    if args.close_browser:
        keep_open = False

    # Parsear --ops → índices 0-based
    ops_idxs = parse_ops_string(args.ops, len(OPCIONES)) if args.ops else None

    ejecutar_agente_masivos(
        ops_indices=ops_idxs,
        sheet_name=args.sheet_name,
        sheet_tab=args.sheet_tab,
        fila_inicio=args.start_row,
        chromedriver_path=args.chromedriver,
        keep_browser_open=keep_open,
    )
