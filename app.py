import os
import sys

# ===== CONFIGURACION CRITICA PARA STREAMLIT CLOUD - MEJORADA =====
os.environ['STREAMLIT_SERVER_FILE_WATCHER_TYPE'] = 'none'
os.environ['STREAMLIT_CI'] = 'true'
os.environ['STREAMLIT_SERVER_HEADLESS'] = 'true'
os.environ['STREAMLIT_SERVER_ENABLE_STATIC_SERVING'] = 'true'
os.environ['STREAMLIT_SERVER_ENABLE_XSRF_PROTECTION'] = 'false'

# Monkey patch para evitar problemas de watcher
import streamlit.web.bootstrap
import streamlit.watcher

def no_op_watch(*args, **kwargs):
    return lambda: None

def no_op_watch_file(*args, **kwargs):
    return

streamlit.watcher.path_watcher.watch_file = no_op_watch_file
streamlit.watcher.path_watcher._watch_path = no_op_watch
streamlit.watcher.event_based_path_watcher.EventBasedPathWatcher.__init__ = lambda *args, **kwargs: None
streamlit.web.bootstrap._install_config_watchers = lambda *args, **kwargs: None

# ===== IMPORTS NORMALES =====
import streamlit as st
import pandas as pd
import numpy as np
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
import re
import tempfile

# Configuracion adicional para Streamlit
st.set_page_config(
    page_title="Validador Power BI - APP ALMA",
    page_icon="💳",
    layout="wide"
)

# ===== CSS Sidebar =====
st.markdown("""
<style>
/* ===== Sidebar ===== */
[data-testid="stSidebar"] {
    background-color: #1E1E2F !important;
    color: white !important;
    width: 300px !important;
    padding: 20px 10px 20px 10px !important;
    border-right: 1px solid #333 !important;
}

/* Texto general en blanco */
[data-testid="stSidebar"] h1, 
[data-testid="stSidebar"] h2, 
[data-testid="stSidebar"] h3,
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] .stMarkdown p,
[data-testid="stSidebar"] .stCheckbox label {
    color: white !important; 
}

/* SOLO el label del file_uploader en blanco */
[data-testid="stSidebar"] .stFileUploader > label {
    color: white !important;
    font-weight: bold;
}

/* Mantener en negro el resto del uploader */
[data-testid="stSidebar"] .stFileUploader .uppy-Dashboard-AddFiles-title,
[data-testid="stSidebar"] .stFileUploader .uppy-Dashboard-AddFiles-subtitle,
[data-testid="stSidebar"] .stFileUploader .uppy-Dashboard-AddFiles-list button,
[data-testid="stSidebar"] .stFileUploader .uppy-Dashboard-Item-name,
[data-testid="stSidebar"] .stFileUploader .uppy-Dashboard-Item-status,
[data-testid="stSidebar"] .stFileUploader span,
[data-testid="stSidebar"] .stFileUploader div {
    color: black !important;
}

/* ===== Boton de expandir/cerrar sidebar ===== */
[data-testid="stSidebarNav"] button {
    background: #2E2E3E !important;
    color: white !important;
    border-radius: 6px !important;
}

/* ===== Encabezados del sidebar ===== */
[data-testid="stSidebar"] h1, 
[data-testid="stSidebar"] h2, 
[data-testid="stSidebar"] h3 {
    color: #00CFFF !important;
}

/* ===== Inputs de texto en el sidebar ===== */
[data-testid="stSidebar"] input[type="text"],
[data-testid="stSidebar"] input[type="password"] {
    color: black !important;
    background-color: white !important;
    border-radius: 6px !important;
    padding: 5px !important;
}

/* ===== BOTON "BROWSE FILES" ===== */
[data-testid="stSidebar"] .uppy-Dashboard-AddFiles-list button {
    color: black !important;
    background-color: #f0f0f0 !important;
    border: 1px solid #ccc !important;
}
[data-testid="stSidebar"] .uppy-Dashboard-AddFiles-list button:hover {
    background-color: #e0e0e0 !important;
}

/* ===== Texto en multiselect ===== */
[data-testid="stSidebar"] .stMultiSelect label,
[data-testid="stSidebar"] .stMultiSelect div[data-baseweb="select"] {
    color: white !important;
}
[data-testid="stSidebar"] .stMultiSelect div[data-baseweb="tag"] {
    color: black !important;
    background-color: #e0e0e0 !important;
}

/* ===== ICONOS DE AYUDA (?) EN EL SIDEBAR ===== */
[data-testid="stSidebar"] svg.icon {
    stroke: white !important;
    color: white !important;
    fill: none !important;
    opacity: 1 !important;
}

/* ===== MEJORAS PARA STREAMLIT CLOUD ===== */
.stSpinner > div > div {
    border-color: #00CFFF !important;
}

.stProgress > div > div > div > div {
    background-color: #00CFFF !important;
}
</style>
""", unsafe_allow_html=True)

# Logo de ALMA
st.markdown("""
<div style="display: flex; justify-content: center; margin-bottom: 30px;">
    <img src="https://i.imgur.com/z9xt46F.jpeg"
         style="width: 50%; border-radius: 10px; display: block; margin: 0 auto;" 
         alt="Logo ALMA">
</div>
""", unsafe_allow_html=True)

# ===== FUNCIONES DE EXTRACCION DE EXCEL (ALMA) =====

def extract_date_from_excel(df):
    """Extraer fecha de la fila 2 del Excel formato 'REPORTE IP/REV 24 DE SEPTIEMBRE DEL 2025'
       Devuelve fecha en formato YYYY-MM-DD o None.
    """
    try:
        if df.shape[0] < 2:
            return None
        fila_2 = df.iloc[1]
        for celda in fila_2:
            if pd.notna(celda) and isinstance(celda, str):
                texto = celda.upper()
                patron = r'(\d{1,2})\s+DE\s+([A-ZÁÉÍÓÚÑ]+)\s+DEL\s+(\d{4})'
                match = re.search(patron, texto)
                if match:
                    dia, mes_texto, anio = match.groups()
                    meses = {
                        'ENERO': '01', 'FEBRERO': '02', 'MARZO': '03', 'ABRIL': '04',
                        'MAYO': '05', 'JUNIO': '06', 'JULIO': '07', 'AGOSTO': '08',
                        'SEPTIEMBRE': '09', 'SETIEMBRE': '09', 'OCTUBRE': '10',
                        'NOVIEMBRE': '11', 'DICIEMBRE': '12'
                    }
                    mes = meses.get(mes_texto.strip(), '')
                    if mes:
                        return f"{anio}-{mes}-{str(dia).zfill(2)}"
        return None
    except Exception as e:
        st.error(f"Error extrayendo fecha: {e}")
        return None

def _parse_currency_to_float(value):
    """Parsea un string tipo '$12.345.678,90' o '12.345.678,90' a float 12345678.9"""
    try:
        if value is None or (isinstance(value, (float, np.floating)) and pd.isna(value)):
            return None
        if isinstance(value, (int, float, np.integer, np.floating)):
            return float(value)
        s = str(value).strip()
        s = s.replace(' ', '').replace('\xa0', '')
        s = re.sub(r'[^\d,.-]', '', s)
        if '.' in s and ',' in s:
            s = s.replace('.', '').replace(',', '.')
        else:
            if s.count('.') > 1:
                s = s.replace('.', '')
            if ',' in s and s.count(',') == 1:
                s = s.replace(',', '.')
            if s.count(',') > 1:
                s = s.replace(',', '')
        if s == '' or s == '-':
            return None
        return float(s)
    except Exception:
        return None

def extract_excel_values_alma(uploaded_file):
    """Extraer TOTAL y NUMERO DE REGISTROS del Excel unico de ALMA"""
    try:
        df = pd.read_excel(uploaded_file, sheet_name=0, header=None)
        fecha = extract_date_from_excel(df)
        valor_total = None
        numero_registros = None

        rows = df.values.tolist()

        for i, row in enumerate(rows):
            fila_textos = []
            for v in row:
                if pd.isna(v):
                    fila_textos.append('')
                else:
                    fila_textos.append(str(v).upper())

            if any('TOTAL' in t for t in fila_textos):
                for offset in range(1, 4):
                    for j, cell in enumerate(row):
                        if pd.isna(cell):
                            continue
                        if 'TOTAL' in fila_textos[j]:
                            right_idx = j + offset
                            if right_idx < len(row):
                                candidato = row[right_idx]
                                parsed = _parse_currency_to_float(candidato)
                                if parsed is not None:
                                    valor_total = parsed
                                    break
                    if valor_total is not None:
                        break

            if any('REGISTRO' in t or 'REGISTROS' in t or 'NUMERO' in t or 'N°' in t for t in fila_textos):
                for offset in range(1, 4):
                    for j, cell in enumerate(row):
                        if pd.isna(cell):
                            continue
                        if 'REGISTRO' in fila_textos[j] or 'REGISTROS' in fila_textos[j] or 'NUMERO' in fila_textos[j] or 'N°' in fila_textos[j]:
                            right_idx = j + offset
                            if right_idx < len(row):
                                candidato = row[right_idx]
                                if pd.isna(candidato):
                                    continue
                                if isinstance(candidato, (int, np.integer)):
                                    numero_registros = int(candidato)
                                    break
                                if isinstance(candidato, (float, np.floating)):
                                    if pd.isna(candidato):
                                        continue
                                    numero_registros = int(candidato)
                                    break
                                cand_str = str(candidato).replace('.', '').replace(',', '').strip()
                                if cand_str.isdigit():
                                    numero_registros = int(cand_str)
                                    break
                    if numero_registros is not None:
                        break

            if valor_total is not None and numero_registros is not None:
                break

        if valor_total is None:
            for r in rows[-6:]:
                for candidate in r:
                    if pd.isna(candidate):
                        continue
                    val = _parse_currency_to_float(candidate)
                    if val is not None:
                        valor_total = val
                        break
                if valor_total is not None:
                    break

        return valor_total, numero_registros, fecha

    except Exception as e:
        st.error(f"Error procesando archivo Excel: {str(e)}")
        return None, None, None

# ===== FUNCIONES DE EXTRACCION DE POWER BI (ALMA) =====

def setup_driver():
    """Configurar ChromeDriver para Selenium"""
    try:
        chrome_options = Options()
        
        chrome_options.add_argument("--headless=new")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        
        chrome_options.add_argument("--user-agent=Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
        
        try:
            driver = webdriver.Chrome(options=chrome_options)
            driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            return driver
        except Exception as e:
            st.error(f"Error al configurar ChromeDriver: {e}")
            return None
            
    except Exception as e:
        st.error(f"Error critico al configurar ChromeDriver: {e}")
        return None

def click_conciliacion_alma(driver, fecha_objetivo):
    """Hacer clic en la conciliacion ALMA especifica por fecha formato YYYY-MM-DD"""
    try:
        fecha_partes = fecha_objetivo.split('-')
        anio = fecha_partes[0]
        mes_num = fecha_partes[1]
        dia = fecha_partes[2]
        
        formatos_especificos = [
            f"ConciliaciónALMAdel{anio}−{mes_num}−{dia}00:00al{anio}−{mes_num}−{dia}11:59",
            f"Conciliación ALMA del {anio}-{mes_num}-{dia} 00:00 al {anio}-{mes_num}-{dia} 11:59",
            f"CONCILIACIÓN ALMA DEL {anio}-{mes_num}-{dia} 00:00 AL {anio}-{mes_num}-{dia} 11:59",
            f"ALMA {anio}-{mes_num}-{dia}",
            f"ALMA {dia}/{mes_num}/{anio}",
            f"Conciliación ALMA {anio}-{mes_num}-{dia}",
        ]
        
        elemento_conciliacion = None
        
        for formato in formatos_especificos:
            try:
                formato_busqueda = formato.replace('−', '-')
                selector = f"//*[contains(text(), '{formato_busqueda}')]"
                elementos = driver.find_elements(By.XPATH, selector)
                
                for elemento in elementos:
                    if elemento.is_displayed():
                        elemento_conciliacion = elemento
                        st.success(f"✅ Encontrada conciliacion con formato especifico: {elemento.text}")
                        break
                if elemento_conciliacion:
                    break
            except Exception as e:
                continue
        
        if not elemento_conciliacion:
            try:
                elementos_alma = driver.find_elements(By.XPATH, "//*[contains(text(), 'ALMA') or contains(text(), 'Alma')]")
                
                for elemento in elementos_alma:
                    if elemento.is_displayed():
                        texto_elemento = elemento.text
                        if (f"{anio}-{mes_num}-{dia}" in texto_elemento or 
                            f"{anio}−{mes_num}−{dia}" in texto_elemento or
                            f"{dia}/{mes_num}/{anio}" in texto_elemento):
                            elemento_conciliacion = elemento
                            st.success(f"✅ Encontrada conciliacion por partes: {texto_elemento}")
                            break
            except Exception as e:
                st.warning(f"Busqueda por partes no exitosa: {e}")
        
        if not elemento_conciliacion:
            patrones_fecha = [
                f"{anio}-{mes_num}-{dia}",
                f"{anio}−{mes_num}−{dia}",
                f"{dia}/{mes_num}/{anio}",
            ]
            
            for patron in patrones_fecha:
                try:
                    selector = f"//*[contains(text(), '{patron}')]"
                    elementos = driver.find_elements(By.XPATH, selector)
                    for elemento in elementos:
                        if elemento.is_displayed():
                            texto_elemento = elemento.text.upper()
                            if 'ALMA' in texto_elemento or 'CONCILIACIÓN' in texto_elemento:
                                elemento_conciliacion = elemento
                                st.success(f"✅ Encontrada por fecha y ALMA: {elemento.text}")
                                break
                    if elemento_conciliacion:
                        break
                except:
                    continue
        
        if not elemento_conciliacion:
            try:
                elementos_alma = driver.find_elements(By.XPATH, "//*[contains(text(), 'ALMA') or contains(text(), 'Alma')]")
                for elemento in elementos_alma:
                    if elemento.is_displayed() and elemento.is_enabled():
                        elemento_conciliacion = elemento
                        st.success(f"✅ Encontrado elemento ALMA: {elemento.text}")
                        break
            except:
                pass
        
        if elemento_conciliacion:
            driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", elemento_conciliacion)
            time.sleep(2)
            
            try:
                elemento_conciliacion.click()
                st.success("✅ Clic realizado con exito")
            except Exception as e:
                try:
                    driver.execute_script("arguments[0].click();", elemento_conciliacion)
                    st.success("✅ Clic realizado con JavaScript")
                except Exception as e2:
                    try:
                        from selenium.webdriver.common.action_chains import ActionChains
                        actions = ActionChains(driver)
                        actions.move_to_element(elemento_conciliacion).click().perform()
                        st.success("✅ Clic realizado con ActionChains")
                    except Exception as e3:
                        st.error(f"❌ No se pudo hacer clic: {e3}")
                        return False
            
            time.sleep(5)
            return True
        else:
            st.error(f"❌ No se encontro la conciliacion para la fecha {fecha_objetivo}")
            st.info("🔍 Elementos disponibles en la pagina que contienen 'ALMA':")
            try:
                elementos_texto = driver.find_elements(By.XPATH, "//*[text()]")
                textos_alma = []
                for elem in elementos_texto:
                    if elem.is_displayed() and elem.text.strip():
                        texto = elem.text.strip()
                        if 'ALMA' in texto.upper() or 'CONCILIACIÓN' in texto.upper():
                            textos_alma.append(texto)
                
                if textos_alma:
                    for texto in textos_alma[:10]:
                        st.write(f"• {texto}")
                else:
                    st.write("No se encontraron elementos con 'ALMA' o 'CONCILIACIÓN'")
                    
            except Exception as e:
                st.write(f"Error al buscar elementos: {e}")
            
            return False
            
    except Exception as e:
        st.error(f"❌ Error al hacer clic en conciliacion: {str(e)}")
        return False

def find_valor_a_pagar_alma(driver):
    """Buscar 'VALOR A PAGAR A COMERCIO' en Power BI ALMA"""
    try:
        elementos = driver.find_elements(By.XPATH, "//*[text()]")
        
        for elemento in elementos:
            if elemento.is_displayed():
                texto = elemento.text.strip()
                
                if 'VALOR A PAGAR A COMERCIO' in texto and 'CANTIDADPASOS' in texto:
                    st.info(f"📝 Texto completo encontrado: '{texto}'")
                    
                    patron_valor = r'VALOR A PAGAR A COMERCIO[^\d]*([\d,]+)'
                    match_valor = re.search(patron_valor, texto)
                    
                    if match_valor:
                        valor_extraido = match_valor.group(1)
                        st.success(f"✅ Valor extraido correctamente: {valor_extraido}")
                        return valor_extraido
                    
                    patron_valor_alt = r'VALOR A PAGAR A COMERCIO.*?(\d{1,3}(?:,\d{3})*)'
                    match_valor_alt = re.search(patron_valor_alt, texto)
                    
                    if match_valor_alt:
                        valor_extraido = match_valor_alt.group(1)
                        st.success(f"✅ Valor extraido (alternativo): {valor_extraido}")
                        return valor_extraido
        
        st.warning("No se encontro el texto combinado, buscando por separado...")
        
        elementos_numericos = driver.find_elements(By.XPATH, "//*[text()]")
        for elemento in elementos_numericos:
            if elemento.is_displayed():
                texto = elemento.text.strip()
                if texto and re.match(r'^\$?[\d,]+$', texto.replace(' ', '')):
                    numero_limpio = texto.replace('$', '').replace(',', '').replace(' ', '')
                    if numero_limpio.isdigit():
                        valor_num = int(numero_limpio)
                        if 1000000 <= valor_num <= 50000000:
                            st.success(f"💰 Valor candidato encontrado: {texto}")
                            match = re.search(r'([\d,]+)', texto)
                            if match:
                                return match.group(1)
        
        st.error("No se pudo encontrar el valor numerico")
        return None
        
    except Exception as e:
        st.error(f"Error buscando valor: {str(e)}")
        return None

def find_cantidad_pasos_alma(driver):
    """Buscar la tarjeta/table 'CANTIDAD PASOS' a la derecha de 'VALOR A PAGAR A COMERCIO'"""
    try:
        titulo_selectors = [
            "//*[contains(text(), 'CANTIDAD PASOS')]",
            "//*[contains(text(), 'Cantidad Pasos')]",
            "//*[contains(text(), 'CANTIDAD DE PASOS')]",
            "//*[contains(text(), 'Cantidad de Pasos')]",
            "//*[contains(text(), 'CANTIDAD') and contains(text(), 'PASOS')]",
            "//*[text()='CANTIDAD PASOS']",
            "//*[text()='Cantidad Pasos']",
        ]
        
        titulo_element = None
        for selector in titulo_selectors:
            try:
                elementos = driver.find_elements(By.XPATH, selector)
                for elemento in elementos:
                    if elemento.is_displayed():
                        texto = elemento.text.strip()
                        if any(palabra in texto.upper() for palabra in ['CANTIDAD', 'PASOS']):
                            titulo_element = elemento
                            st.success(f"✅ Titulo encontrado: {texto}")
                            break
                if titulo_element:
                    break
            except Exception as e:
                continue
        
        if not titulo_element:
            st.warning("❌ No se encontro el titulo 'CANTIDAD PASOS'")
            return None
        
        try:
            container = titulo_element.find_element(By.XPATH, "./..")
            all_elements = container.find_elements(By.XPATH, ".//*")
            
            for elem in all_elements:
                texto = elem.text.strip()
                if (texto and 
                    any(char.isdigit() for char in texto) and 
                    len(texto) < 20 and 
                    texto != titulo_element.text and
                    not any(word in texto.upper() for word in ['TOTAL', 'VALOR', 'PAGAR', 'COMERCIO', 'CANTIDAD', 'PASOS'])):
                    
                    digit_count = sum(char.isdigit() for char in texto)
                    if digit_count >= 1:
                        st.success(f"✅ Valor numerico encontrado: {texto}")
                        return texto
                        
        except Exception as e:
            st.warning(f"⚠️ Estrategia 1 fallo: {e}")
        
        try:
            parent = titulo_element.find_element(By.XPATH, "./..")
            siblings = parent.find_elements(By.XPATH, "./*")
            
            for sibling in siblings:
                if sibling != titulo_element:
                    texto = sibling.text.strip()
                    if (texto and 
                        any(char.isdigit() for char in texto) and 
                        len(texto) < 20 and
                        not any(word in texto.upper() for word in ['TOTAL', 'VALOR', 'PAGAR', 'COMERCIO', 'CANTIDAD', 'PASOS'])):
                        
                        digit_count = sum(char.isdigit() for char in texto)
                        if digit_count >= 1:
                            st.success(f"✅ Valor encontrado en hermano: {texto}")
                            return texto
        except Exception as e:
            st.warning(f"⚠️ Estrategia 2 fallo: {e}")
        
        try:
            following_elements = driver.find_elements(By.XPATH, f"//*[contains(text(), 'CANTIDAD PASOS')]/following::*")
            
            for i, elem in enumerate(following_elements[:20]):
                texto = elem.text.strip()
                if (texto and 
                    any(char.isdigit() for char in texto) and 
                    len(texto) < 20 and
                    not any(word in texto.upper() for word in ['TOTAL', 'VALOR', 'PAGAR', 'COMERCIO', 'CANTIDAD', 'PASOS'])):
                    
                    digit_count = sum(char.isdigit() for char in texto)
                    if digit_count >= 1:
                        st.success(f"✅ Valor encontrado en elemento siguiente: {texto}")
                        return texto
        except Exception as e:
            st.warning(f"⚠️ Estrategia 3 fallo: {e}")
        
        try:
            valor_element = driver.find_element(By.XPATH, "//*[contains(text(), 'VALOR A PAGAR A COMERCIO')]")
            if valor_element:
                container_valor = valor_element.find_element(By.XPATH, "./..")
                all_nearby = container_valor.find_elements(By.XPATH, ".//*")
                
                for elem in all_nearby:
                    texto = elem.text.strip()
                    if (texto and 
                        any(char.isdigit() for char in texto) and 
                        len(texto) < 20 and
                        'CANTIDAD' in texto.upper() and 'PASOS' in texto.upper()):
                        continue
                    
                    if (texto and 
                        any(char.isdigit() for char in texto) and 
                        len(texto) < 20 and
                        not any(word in texto.upper() for word in ['TOTAL', 'VALOR', 'PAGAR', 'COMERCIO'])):
                        
                        digit_count = sum(char.isdigit() for char in texto)
                        if digit_count >= 1:
                            st.success(f"✅ Valor encontrado cerca de VALOR A PAGAR: {texto}")
                            return texto
        except Exception as e:
            st.warning(f"⚠️ Estrategia 4 fallo: {e}")
        
        st.error("❌ No se pudo encontrar el valor numerico de CANTIDAD PASOS")
        return None
        
    except Exception as e:
        st.error(f"❌ Error buscando cantidad de pasos: {str(e)}")
        return None

def extract_powerbi_data_alma(fecha_objetivo):
    """Funcion principal para extraer datos de Power BI ALMA"""
    
    REPORT_URL = "https://app.powerbi.com/view?r=eyJrIjoiMWExM2JkMzctMDgyMi00ZWZhLTgxODUtNGNlZGViYTcyM2NiIiwidCI6ImY5MTdlZDFiLWI0MDMtNDljNS1iODBiLWJhYWUzY2UwMzc1YSJ9"
    
    driver = setup_driver()
    if not driver:
        return None
    
    try:
        with st.spinner("🌐 Conectando con Power BI ALMA..."):
            driver.get(REPORT_URL)
            time.sleep(12)
        
        st.info("📊 Pagina de Power BI cargada")
        driver.save_screenshot("powerbi_alma_inicial.png")
        
        with st.spinner("🔍 Buscando conciliacion..."):
            if not click_conciliacion_alma(driver, fecha_objetivo):
                return None
        
        time.sleep(5)
        driver.save_screenshot("powerbi_alma_despues_seleccion.png")
        st.success("✅ Conciliacion seleccionada")
        
        with st.spinner("💰 Buscando valor a pagar..."):
            valor_texto = find_valor_a_pagar_alma(driver)
        
        with st.spinner("👣 Buscando cantidad de pasos..."):
            cantidad_pasos_texto = find_cantidad_pasos_alma(driver)
        
        driver.save_screenshot("powerbi_alma_final.png")
        
        return {
            'valor_texto': valor_texto,
            'cantidad_pasos_texto': cantidad_pasos_texto,
            'screenshots': {
                'inicial': 'powerbi_alma_inicial.png',
                'seleccion': 'powerbi_alma_despues_seleccion.png',
                'final': 'powerbi_alma_final.png'
            }
        }
        
    except Exception as e:
        st.error(f"Error durante la extraccion: {str(e)}")
        import traceback
        st.error(f"Detalle del error: {traceback.format_exc()}")
        return None
    finally:
        try:
            driver.quit()
        except:
            pass

# ===== FUNCIONES DE COMPARACION =====

def convert_currency_to_float(currency_string):
    """Convierte string de moneda a float"""
    try:
        if isinstance(currency_string, (int, float)):
            return float(currency_string)
            
        if isinstance(currency_string, str):
            cleaned = currency_string.strip()
            cleaned = cleaned.replace(', '')
            cleaned = cleaned.replace(' ', '')
            
            if '.' in cleaned and ',' in cleaned:
                cleaned = cleaned.replace('.', '')
                cleaned = cleaned.replace(',', '.')
            elif '.' in cleaned and cleaned.count('.') > 1:
                cleaned = cleaned.replace('.', '')
            elif ',' in cleaned:
                if cleaned.count(',') == 2 and '.' in cleaned:
                    cleaned = cleaned.replace(',', '')
                elif cleaned.count(',') == 1:
                    cleaned = cleaned.replace(',', '.')
                else:
                    cleaned = cleaned.replace(',', '')
            
            return float(cleaned) if cleaned else 0.0
            
        return float(currency_string)
        
    except Exception as e:
        st.error(f"Error convirtiendo moneda: '{currency_string}' - {e}")
        return 0.0

def compare_values_alma(valor_powerbi, valor_excel):
    """Comparar valores de Power BI y Excel"""
    try:
        powerbi_numero = convert_currency_to_float(valor_powerbi)
        excel_numero = float(valor_excel) if valor_excel else 0
        
        tolerancia = 0.01
        coinciden = abs(powerbi_numero - excel_numero) <= tolerancia
        diferencia = abs(powerbi_numero - excel_numero)
        
        return powerbi_numero, excel_numero, str(valor_powerbi), coinciden, diferencia
        
    except Exception as e:
        st.error(f"Error comparando valores: {e}")
        return None, None, str(valor_powerbi), False, 0

def compare_pasos_alma(pasos_powerbi, pasos_excel):
    """Comparar pasos de Power BI y Excel"""
    try:
        if isinstance(pasos_powerbi, str):
            pasos_powerbi_limpio = re.sub(r'[^\d]', '', pasos_powerbi)
            powerbi_numero = int(pasos_powerbi_limpio) if pasos_powerbi_limpio else 0
        else:
            powerbi_numero = int(pasos_powerbi) if pasos_powerbi else 0
        
        excel_numero = int(pasos_excel) if pasos_excel else 0
        
        coinciden = powerbi_numero == excel_numero
        diferencia = abs(powerbi_numero - excel_numero)
        
        return powerbi_numero, excel_numero, str(pasos_powerbi), coinciden, diferencia
        
    except Exception as e:
        st.error(f"Error comparando pasos: {e}")
        return 0, 0, str(pasos_powerbi), False, 0

# ===== INTERFAZ PRINCIPAL =====

def main():
    st.title("💳 Validador Power BI - Conciliaciones APP ALMA")
    st.markdown("---")
    
    st.sidebar.header("📋 Informacion del Reporte")
    st.sidebar.info("""
    **Objetivo:**
    - Cargar archivo Excel de ALMA
    - Extraer fecha, TOTAL y NUMERO DE REGISTROS
    - Comparar con Power BI automaticamente
    
    **Estado:** ✅ ChromeDriver Compatible
    **Version:** v1.4 - ALMA Automatico
    """)
    
    st.sidebar.header("🛠️ Estado del Sistema")
    st.sidebar.success(f"✅ Python {sys.version_info.major}.{sys.version_info.minor}")
    st.sidebar.info(f"✅ Pandas {pd.__version__}")
    st.sidebar.info(f"✅ Streamlit {st.__version__}")
    
    st.subheader("📁 Cargar Archivo Excel")
    uploaded_file = st.file_uploader(
        "Selecciona el archivo Excel de ALMA", 
        type=['xlsx', 'xls']
    )
    
    if uploaded_file is not None:
        with st.spinner("📊 Procesando archivo Excel..."):
            valor_total, numero_registros, fecha_extraida = extract_excel_values_alma(uploaded_file)
        
        if valor_total is not None and numero_registros is not None:
            
            st.markdown("### 📊 Valores Extraidos del Excel")
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.metric("TOTAL", f"${valor_total:,.0f}".replace(",", "."))
            
            with col2:
                st.metric("NUMERO DE REGISTROS", f"{numero_registros:,}".replace(",", "."))
            
            with col3:
                if fecha_extraida:
                    st.metric("FECHA", fecha_extraida)
                else:
                    st.warning("Fecha no detectada")
            
            st.markdown("---")
            
            if fecha_extraida:
                ejecutar_extraccion = True
            else:
                st.warning("No se pudo extraer la fecha automaticamente del Excel")
                ejecutar_extraccion = False
            
            if ejecutar_extraccion:
                with st.spinner("🌐 Extrayendo datos de Power BI ALMA... Esto puede tomar 1-2 minutos"):
                    resultados = extract_powerbi_data_alma(fecha_extraida)
                    
                    if resultados and (resultados.get('valor_texto') or resultados.get('cantidad_pasos_texto')):
                        valor_powerbi_texto = resultados.get('valor_texto')
                        cantidad_pasos_powerbi = resultados.get('cantidad_pasos_texto')
                        
                        st.markdown("---")
                        
                        st.markdown("### 📊 Valores Extraidos de Power BI")
                        
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            if valor_powerbi_texto:
                                st.metric("VALOR A PAGAR A COMERCIO (BI)", f"${valor_powerbi_texto}")
                            else:
                                st.warning("Valor no encontrado en Power BI")
                        
                        with col2:
                            if cantidad_pasos_powerbi:
                                st.metric("CANTIDAD DE PASOS (BI)", cantidad_pasos_powerbi)
                            else:
                                st.warning("Pasos no encontrados en Power BI")
                        
                        st.markdown("---")
                        
                        if valor_powerbi_texto:
                            st.markdown("### 💰 Validacion: Valores")
                            
                            powerbi_valor, excel_valor, valor_formateado, coinciden_valor, dif_valor = compare_values_alma(
                                valor_powerbi_texto, valor_total
                            )
                            
                            if powerbi_valor is not None:
                                col1, col2, col3 = st.columns([2, 2, 1])
                                
                                with col1:
                                    st.metric("Power BI", f"${powerbi_valor:,.0f}".replace(",", "."))
                                with col2:
                                    st.metric("Excel", f"${excel_valor:,.0f}".replace(",", "."))
                                with col3:
                                    if coinciden_valor:
                                        st.markdown("#### ✅")
                                        st.success("COINCIDE")
                                    else:
                                        st.markdown("#### ❌")
                                        st.error("DIFERENCIA")
                                        st.caption(f"${dif_valor:,.0f}".replace(",", "."))
                        
                        if cantidad_pasos_powerbi:
                            st.markdown("### 👣 Validacion: Numero de Registros/Pasos")
                            
                            powerbi_pasos, excel_pasos, pasos_formateado, coinciden_pasos, dif_pasos = compare_pasos_alma(
                                cantidad_pasos_powerbi, numero_registros
                            )
                            
                            col1, col2, col3 = st.columns([2, 2, 1])
                            
                            with col1:
                                st.metric("Power BI", f"{powerbi_pasos:,}".replace(",", "."))
                            with col2:
                                st.metric("Excel", f"{excel_pasos:,}".replace(",", "."))
                            with col3:
                                if coinciden_pasos:
                                    st.markdown("#### ✅")
                                    st.success("COINCIDE")
                                else:
                                    st.markdown("#### ❌")
                                    st.error("DIFERENCIA")
                                    st.caption(f"{dif_pasos:,}")
                        
                        st.markdown("---")
                        
                        st.markdown("### 📋 Resultado Final")
                        
                        if valor_powerbi_texto and cantidad_pasos_powerbi:
                            if coinciden_valor and coinciden_pasos:
                                st.success("🎉 VALIDACION EXITOSA - Valores y pasos coinciden")
                                st.balloons()
                            elif coinciden_valor and not coinciden_pasos:
                                st.warning("⚠️ VALIDACION PARCIAL - Valores coinciden, pero hay diferencias en pasos")
                            elif not coinciden_valor and coinciden_pasos:
                                st.warning("⚠️ VALIDACION PARCIAL - Pasos coinciden, pero hay diferencias en valores")
                            else:
                                st.error("❌ VALIDACION FALLIDA - Existen diferencias en valores y pasos")
                        else:
                            st.warning("⚠️ VALIDACION INCOMPLETA - No se pudieron extraer todos los datos de Power BI")
                    
                    elif resultados:
                        st.error("Se accedio al reporte pero no se encontraron los valores especificos")
                    else:
                        st.error("No se pudieron extraer datos del reporte Power BI")
        
        else:
            st.error("No se pudieron extraer valores del archivo Excel")
            with st.expander("💡 Sugerencias para solucionar el problema"):
                st.markdown("""
                - Verifica que el Excel sea de ALMA y tenga una unica hoja
                - Asegurate que la fila 2 contenga el texto con la fecha (ej: "REPORTE IP/REV 24 DE SEPTIEMBRE DEL 2025")
                - El archivo debe contener las palabras "TOTAL" y "NUMERO DE REGISTROS" o que en sus ultimas filas aparezcan el total y el conteo.
                - Los valores pueden estar 1 o 2 columnas a la derecha de la etiqueta.
                """)
    
    else:
        st.info("📁 Por favor, carga un archivo Excel para comenzar la validacion automatica")

    st.markdown("---")
    with st.expander("ℹ️ Instrucciones de Uso"):
        st.markdown("""
        **Proceso Automatico:**
        1. **Cargar Excel**: Archivo ALMA con una unica hoja
        2. **Extraccion automatica**: 
           - Busca la fecha en la fila 2
           - Busca "TOTAL" y trae el valor a la derecha (1-3 columnas)
           - Busca "NUMERO DE REGISTROS" y trae el valor a la derecha (1-3 columnas)
        3. **Extraccion Power BI**: Navega automaticamente a la conciliacion ALMA de la fecha extraida
        4. **Comparacion**: Compara automaticamente VALOR A PAGAR A COMERCIO y CANTIDAD DE PASOS
        
        **Mejoras en esta version:**
        - ✅ Extraccion automatica al cargar el archivo
        - ✅ Sin necesidad de hacer clic en botones
        - ✅ Proceso completamente automatizado
        """)

if __name__ == "__main__":
    main()

    st.markdown("---")
    st.markdown('<div class="footer">💻 Desarrollado por Angel Torres | 🚀 Powered by Streamlit | v1.4 ALMA Automatico</div>', unsafe_allow_html=True)
