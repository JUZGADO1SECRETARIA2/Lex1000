# utils_mini.py
# Versión reducida con solo las funciones esenciales

import time
import os
import gspread
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from oauth2client.service_account import ServiceAccountCredentials
from dotenv import load_dotenv

# === CONFIGURACIÓN BÁSICA ===
PROJECT_ROOT = Path(__file__).resolve().parent.parent
load_dotenv(PROJECT_ROOT / ".env")

LEX100_URL = os.getenv("LEX100_URL")
CUIT = os.getenv("CUIT")
PASSWORD = os.getenv("PASSWORD")

DOWNLOAD_DIR = PROJECT_ROOT / "downloads"
DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

CREDENCIALES_PATH = PROJECT_ROOT / "service_account.json"

# ============== FUNCIONES ESENCIALES ==============

def autenticar_google_sheets(sheet_name, pestaña):
    """
    Autentica con Google Sheets usando service_account.json
    """
    if not CREDENCIALES_PATH.exists():
        raise FileNotFoundError(
            f"❌ No encuentro la credencial en: {CREDENCIALES_PATH}"
        )

    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = ServiceAccountCredentials.from_json_keyfile_name(str(CREDENCIALES_PATH), scope)
    cliente = gspread.authorize(creds)
    return cliente.open(sheet_name).worksheet(pestaña)

def configurar_selenium(chrome_driver_path: str = None):
    """
    Configura Selenium con Chrome
    """
    options = webdriver.ChromeOptions()
    
    prefs = {
        "download.default_directory": str(DOWNLOAD_DIR),
        "download.prompt_for_download": False,
        "directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    options.add_experimental_option("prefs", prefs)
    options.add_experimental_option("detach", True)
    options.add_argument("--start-maximized")
    options.add_argument("--log-level=3")

    try:
        if chrome_driver_path:
            service = Service(chrome_driver_path)
        else:
            service = Service()

        print("🧭 Abriendo Google Chrome...")
        driver = webdriver.Chrome(service=service, options=options)
    except Exception as e:
        raise RuntimeError(f"No pude iniciar Chrome: {e}")

    wait = WebDriverWait(driver, 10)
    actions = ActionChains(driver)
    return driver, wait, actions

def iniciar_sesion(driver, wait):
    """
    Inicia sesión en Lex100
    """
    driver.get(LEX100_URL)

    # Paso 1: Login básico
    wait.until(EC.presence_of_element_located((By.ID, 'username'))).send_keys(CUIT)
    driver.find_element(By.ID, 'password').send_keys(PASSWORD)
    driver.find_element(By.ID, 'kc-login').click()

    # Paso 2: Verificar selección de perfil
    try:
        perfil_css = "#kc-perfil-login-form > ul > li.collection-item.avatar.perfil-item.item-color-2 > p"
        element = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, perfil_css)))
        driver.execute_script("arguments[0].scrollIntoView(true);", element)
        driver.execute_script("arguments[0].click();", element)
        print("🔁 SECRETARÍA N°2 seleccionada correctamente.")
        time.sleep(2)
    except TimeoutException:
        print("✅ No se mostró selección de perfil. Continuando normalmente.")
    
    # Paso 3: Esperar carga del sistema
    try:
        print("⏳ Esperando que cargue el sistema luego del login...")
        wait.until(EC.presence_of_element_located((By.XPATH, '//div[text()="Expedientes"]')))
        print("✅ Sistema cargado correctamente.")
    except TimeoutException:
        print("⚠️ No se detectó el menú 'Expedientes'. Verificar si se cargó bien el sistema.")

def abrir_menu_masivos_documentos_digitales(driver, wait, verificar_input=True):
    """
    Navega a: Masivos -> Despacho de Documentos -> Documentos digitales
    """
    # 1) Click en 'Masivos'
    try:
        masivos_container = wait.until(EC.presence_of_element_located((By.ID, "toolbarForm:j_id244")))
        ActionChains(driver).move_to_element(masivos_container).pause(0.15).perform()
        masivos_btn = masivos_container.find_element(By.XPATH, "./div[1]")
        ActionChains(driver).move_to_element(masivos_btn).pause(0.1).click().perform()
        time.sleep(0.3)
    except Exception:
        # Fallback
        driver.find_element(By.XPATH, '//*[@id="toolbarForm:j_id244"]/div[1]').click()

    # 2) Click en 'Despacho de Documentos'
    try:
        ActionChains(driver).move_to_element(masivos_container).pause(0.15).perform()
    except Exception:
        pass

    # Buscar por texto
    despacho_xpath = (
        '//*[@id="toolbarForm:j_id244"]//span[contains(@class,"rich-menu-item-label") and '
        'contains(translate(normalize-space(.),'
        '"ABCDEFGHIJKLMNOPQRSTUVWXYZÁÉÍÓÚÜ","abcdefghijklmnopqrstuvwxyzáéíóúü"),'
        '"despacho de documentos")]'
    )
    
    try:
        despacho_btn = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, despacho_xpath)))
        driver.execute_script("arguments[0].scrollIntoView(true);", despacho_btn)
        driver.execute_script("arguments[0].click();", despacho_btn)
        time.sleep(0.3)
    except Exception:
        raise TimeoutException("No pude clickear 'Despacho de Documentos'")

    # 3) Click en 'Documentos digitales'
    documentos_xpath = '//*[@id="masivoDespachoExpedientes"]/div/div/table/tbody/tr/td[2]/span/h2/a'
    driver.find_element(By.XPATH, documentos_xpath).click()

    if not verificar_input:
        return True

    # 4) Verificar input de código de barras
    NAME_CODIGOBARRAS = 'despachoDocumentosMasivoDecorate:searchFilters:search1:filterFormVisible:codigoBarras'
    el = wait.until(EC.presence_of_element_located((By.NAME, NAME_CODIGOBARRAS)))
    return el

def masivo_confirmar_seleccion_final(driver, wait):
    """
    Segundo 'Confirmar selección' en pantalla de parámetros masivos
    """
    intentos = [
        (By.NAME, 'parametrosMasivoDespacho:j_id479'),
        (By.CSS_SELECTOR, '#parametrosMasivoDespacho\\:j_id477 > input:nth-child(2)'),
        (By.XPATH, "//input[@type='submit' and contains(@value,'Confirmar') and contains(@value,'selección')]"),
    ]
    
    for how, sel in intentos:
        try:
            el = wait.until(EC.element_to_be_clickable((how, sel)))
            driver.execute_script("arguments[0].scrollIntoView(true);", el)
            time.sleep(0.1)
            driver.execute_script("arguments[0].click();", el)
            time.sleep(0.3)
            return True
        except Exception:
            continue
    
    raise TimeoutException("No pude hacer el segundo 'Confirmar selección'.")

def seleccionar_modelo(driver, wait, clave: str, sugerencia_xpath: str):
    """
    Selecciona un modelo del autosuggest
    """
    INPUT_ID = 'despacho:modeloDecoration:modelo'
    inp = wait.until(EC.presence_of_element_located((By.ID, INPUT_ID)))
    
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", inp)
    try: 
        inp.click()
    except: 
        driver.execute_script("arguments[0].click();", inp)
    
    try: 
        inp.clear()
    except: 
        pass

    # Escribir clave
    try:
        ActionChains(driver).move_to_element(inp).click().pause(0.05).send_keys(clave).perform()
    except:
        try: 
            inp.send_keys(clave)
        except: 
            pass
    
    time.sleep(0.25)

    # Forzar valor si no quedó
    val = inp.get_attribute("value") or ""
    if val.strip().lower() != clave.strip().lower():
        driver.execute_script("""
            const el=arguments[0], txt=arguments[1];
            el.focus(); el.value=txt;
            el.dispatchEvent(new Event('input',{bubbles:true}));
            el.dispatchEvent(new Event('change',{bubbles:true}));
        """, inp, clave)

    # Esperar sugerencias
    try:
        wait.until(EC.presence_of_element_located((By.ID, "despacho:modeloDecoration:modeloSuggestion")))
    except:
        try: 
            inp.send_keys("r"); time.sleep(0.08); inp.send_keys(Keys.BACK_SPACE)
        except: 
            pass
    
    # Click en sugerencia
    sug = wait.until(EC.visibility_of_element_located((By.XPATH, sugerencia_xpath)))
    
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", sug)
        time.sleep(0.05)
        sug.click()
        return True
    except:
        try:
            driver.execute_script("arguments[0].click();", sug)
            return True
        except:
            try:
                ActionChains(driver).move_to_element(sug).pause(0.05).click().perform()
                return True
            except:
                raise TimeoutException("No pude clickear la opción del modelo.")

def masivo_marcar_a_la_firma(driver, wait, marcar=True, timeout=8):
    """
    Marca/desmarca el checkbox 'A la firma'
    """
    intentos = [
        (By.CSS_SELECTOR, '#despacho\\:despachoMasivoDiv > input[type="checkbox"]'),
        (By.NAME, 'despacho:j_id5689'),
        (By.XPATH, '/html/body/div[4]/div[2]/div/form[2]/div/div/input'),
    ]

    chk = None
    for how, sel in intentos:
        try:
            chk = wait.until(EC.presence_of_element_located((how, sel)))
            break
        except Exception:
            continue

    if chk is None:
        raise TimeoutException("No encontré el checkbox 'A la firma'.")

    # Si ya está en el estado deseado, salir
    try:
        if chk.is_selected() == marcar:
            return True
    except Exception:
        pass

    # Click
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", chk)
        time.sleep(0.05)
        chk.click()
    except Exception:
        try:
            driver.execute_script("arguments[0].click();", chk)
        except Exception:
            ActionChains(driver).move_to_element(chk).pause(0.05).click().perform()

    # Esperar cambio de estado
    fin = time.time() + timeout
    while time.time() < fin:
        try:
            if chk.is_selected() == marcar:
                return True
        except Exception:
            pass
        time.sleep(0.1)

    raise TimeoutException("El checkbox 'A la firma' no quedó en el estado esperado.")

def estacionar_mouse(driver, target=None):
    """
    Mueve el cursor lejos del toolbar
    """
    try:
        if target is not None:
            ActionChains(driver).move_to_element(target).pause(0.05).perform()
        else:
            body = driver.find_element(By.TAG_NAME, "body")
            ActionChains(driver).move_to_element_with_offset(body, 5, 5).pause(0.05).perform()
    except Exception:
        pass

def seleccionar_modelo_por_texto(driver, wait, clave: str, texto_objetivo: str, frag_fallback: str = None):
    """
    Selecciona modelo por texto visible en las sugerencias
    """
    def _norm_txt(s: str) -> str:
        return " ".join((s or "").strip().lower().split())

    # 1) Input
    inp = None
    for how, sel in [
        (By.ID, 'despacho:modeloDecoration:modeloSuggestionInput'),
        (By.ID, 'despacho:modeloDecoration:modelo'),
    ]:
        try:
            inp = wait.until(EC.element_to_be_clickable((how, sel)))
            break
        except Exception:
            continue
    
    if inp is None:
        raise TimeoutException("No encontré el input de modelo")

    # Escribir clave
    try: 
        inp.clear()
    except Exception: 
        pass
    
    inp.click()
    inp.send_keys(clave)

    # 2) Esperar sugerencias
    try:
        cont = wait.until(EC.presence_of_element_located((By.ID, "despacho:modeloDecoration:modeloSuggestion")))
    except TimeoutException:
        try:
            inp.send_keys(" ")
            time.sleep(0.05)
            inp.send_keys(Keys.BACK_SPACE)
        except Exception:
            pass
    
    # 3) Buscar opción por texto
    tgt_full = _norm_txt(texto_objetivo)
    tgt_frag = _norm_txt(frag_fallback or texto_objetivo)

    elegido = None
    # Buscar en elementos visibles
    for xp in [".//tr", ".//li", ".//div", ".//span", ".//a", ".//td"]:
        try:
            for n in cont.find_elements(By.XPATH, xp):
                try:
                    if n.is_displayed():
                        t = (n.text or "").strip()
                        if t:
                            if _norm_txt(t) == tgt_full:
                                elegido = n
                                break
                            elif tgt_frag and tgt_frag in _norm_txt(t):
                                elegido = n
                                break
                except Exception:
                    continue
            if elegido:
                break
        except Exception:
            continue

    if elegido is None:
        raise TimeoutException("No encontré la opción deseada en las sugerencias")

    # 4) Click
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elegido)
        time.sleep(0.06)
    except Exception:
        pass

    try:
        ActionChains(driver).move_to_element(elegido).pause(0.10).click().perform()
        return True
    except Exception:
        try:
            driver.execute_script("arguments[0].click();", elegido)
            return True
        except Exception:
            raise TimeoutException("No pude seleccionar el modelo")

# === FUNCIONES AUXILIARES ===
def _safe_click(driver, element):
    """Click seguro con JavaScript"""
    try:
        element.click()
    except:
        driver.execute_script("arguments[0].click();", element)

def _scroll_to(driver, element):
    """Scroll seguro a elemento"""
    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", element)
        time.sleep(0.1)
    except:
        pass
