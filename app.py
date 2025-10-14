import os
import sys

# ===== CONFIGURACIÓN CRÍTICA PARA STREAMLIT CLOUD - MEJORADA =====
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

# Configuración adicional para Streamlit
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

/* ===== Botón de expandir/cerrar sidebar ===== */
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

/* ===== BOTÓN "BROWSE FILES" ===== */
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

# ===== FUNCIONES DE EXTRACCIÓN DE EXCEL (ALMA) =====

def extract_date_from_excel(df):
    """Extraer fecha de la fila 2 del Excel formato 'REPORTE IP/REV 24 DE SEPTIEMBRE DEL 2025'
       Devuelve fecha en formato YYYY-MM-DD o None.
    """
    try:
        # fila 2 = índice 1
        if df.shape[0] < 2:
            return None
        fila_2 = df.iloc[1]
        for celda in fila_2:
            if pd.notna(celda) and isinstance(celda, str):
                texto = celda.upper()
                # Buscar patrón día DE MES DEL YYYY
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
        # VERIFICAR SI ES NaN PRIMERO
        if value is None or (isinstance(value, (float, np.floating)) and pd.isna(value)):
            return None
        if isinstance(value, (int, float, np.integer, np.floating)):
            return float(value)
        s = str(value).strip()
        # eliminar espacios y símbolos
        s = s.replace(' ', '').replace('\xa0', '')
        # si tiene $ o COP, quitarlo
        s = re.sub(r'[^\d,.-]', '', s)
        # Si hay tanto punto como coma, asumimos que punto es separador de miles
        if '.' in s and ',' in s:
            s = s.replace('.', '').replace(',', '.')
        else:
            # si sólo tiene puntos y más de 1 punto, quitar puntos (miles)
            if s.count('.') > 1:
                s = s.replace('.', '')
            # si sólo tiene coma como decimal
            if ',' in s and s.count(',') == 1:
                s = s.replace(',', '.')
            # si tiene coma múltiples, quitar comas
            if s.count(',') > 1:
                s = s.replace(',', '')
        if s == '' or s == '-':
            return None
        return float(s)
    except Exception:
        return None

def extract_excel_values_alma(uploaded_file):
    """Extraer TOTAL y NUMERO DE REGISTROS del Excel único de ALMA
       Retorna (valor_total, numero_registros, fecha) donde:
         - valor_total: float (o None)
         - numero_registros: int (o None)
         - fecha: 'YYYY-MM-DD' (o None)
    """
    try:
        # Leer la hoja 0 sin encabezados
        df = pd.read_excel(uploaded_file, sheet_name=0, header=None)
        fecha = extract_date_from_excel(df)
        valor_total = None
        numero_registros = None

        # Normalizar texto de búsqueda
        rows = df.values.tolist()

        for i, row in enumerate(rows):
            # construir versión en mayúsculas para búsquedas
            fila_textos = []
            for v in row:
                if pd.isna(v):
                    fila_textos.append('')
                else:
                    fila_textos.append(str(v).upper())

            # Si fila contiene 'TOTAL' buscar valor en columnas a la derecha
            if any('TOTAL' in t for t in fila_textos):
                # buscar valor numérico en la misma fila: derecha 1..3 columnas
                for offset in range(1, 4):
                    for j, cell in enumerate(row):
                        if pd.isna(cell):
                            continue
                        # Si la celda actual (j) contiene TOTAL (chequeamos en fila_textos)
                        if 'TOTAL' in fila_textos[j]:
                            # comprobar j+offset
                            right_idx = j + offset
                            if right_idx < len(row):
                                candidato = row[right_idx]
                                parsed = _parse_currency_to_float(candidato)
                                if parsed is not None:
                                    valor_total = parsed
                                    break
                    if valor_total is not None:
                        break

            # Si fila contiene 'REGISTRO' (o variantes) buscar número de registros
            if any('REGISTRO' in t or 'REGISTROS' in t or 'NUMERO' in t or 'N°' in t for t in fila_textos):
                for offset in range(1, 4):
                    for j, cell in enumerate(row):
                        if pd.isna(cell):
                            continue
                        if 'REGISTRO' in fila_textos[j] or 'REGISTROS' in fila_textos[j] or 'NUMERO' in fila_textos[j] or 'N°' in fila_textos[j]:
                            right_idx = j + offset
                            if right_idx < len(row):
                                candidato = row[right_idx]
                                # VERIFICAR SI ES NaN ANTES DE CONVERTIR
                                if pd.isna(candidato):
                                    continue
                                # limpiar y convertir a int si es posible
                                if isinstance(candidato, (int, np.integer)):
                                    numero_registros = int(candidato)
                                    break
                                # si viene como float convertible (y no es NaN)
                                if isinstance(candidato, (float, np.floating)):
                                    # VERIFICAR EXPLÍCITAMENTE QUE NO SEA NaN
                                    if pd.isna(candidato):
                                        continue
                                    numero_registros = int(candidato)
                                    break
                                # si viene como string con separadores
                                cand_str = str(candidato).replace('.', '').replace(',', '').strip()
                                if cand_str.isdigit():
                                    numero_registros = int(cand_str)
                                    break
                    if numero_registros is not None:
                        break

            # si ambos encontrados, salir
            if valor_total is not None and numero_registros is not None:
                break

        # Caso especial: si no se encontró valor_total buscando 'TOTAL', buscar en últimas filas
        if valor_total is None:
            # inspeccionar últimas filas por una celda que parezca moneda
            for r in rows[-6:]:
                for candidate in r:
                    # VERIFICAR SI ES NaN ANTES DE PROCESAR
                    if pd.isna(candidate):
                        continue
                    val = _parse_currency_to_float(candidate)
                    if val is not None:
                        # asumimos el primer candidato encontrado en las últimas filas es el total
                        valor_total = val
                        break
                if valor_total is not None:
                    break

        return valor_total, numero_registros, fecha

    except Exception as e:
        st.error(f"Error procesando archivo Excel: {str(e)}")
        return None, None, None

# ===== FUNCIONES DE EXTRACCIÓN DE POWER BI (ALMA) =====

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
        st.error(f"Error crítico al configurar ChromeDriver: {e}")
        return None

def click_conciliacion_alma(driver, fecha_objetivo):
    """Hacer clic en la conciliación ALMA específica por fecha formato YYYY-MM-DD"""
    try:
        # Convertir fecha de YYYY-MM-DD a DD-MM-YYYY para la búsqueda
        fecha_partes = fecha_objetivo.split('-')
        fecha_formatos = [
            f"Conciliación ALMA de {fecha_objetivo}",
            f"Conciliación ALMA {fecha_objetivo}",
            f"CONCILIACIÓN ALMA DE {fecha_objetivo}",
            f"CONCILIACIÓN ALMA {fecha_objetivo}",
            f"Conciliación ALMA de {fecha_partes[2]}/{fecha_partes[1]}/{fecha_partes[0]}",
        ]
        
        elemento_conciliacion = None
        for formato in fecha_formatos:
            try:
                selector = f"//*[contains(text(), '{formato}')]"
                elementos = driver.find_elements(By.XPATH, selector)
                for elemento in elementos:
                    if elemento.is_displayed():
                        elemento_conciliacion = elemento
                        break
                if elemento_conciliacion:
                    break
            except:
                continue
        
        if elemento_conciliacion:
            driver.execute_script("arguments[0].scrollIntoView(true);", elemento_conciliacion)
            time.sleep(1)
            driver.execute_script("arguments[0].click();", elemento_conciliacion)
            time.sleep(3)
            return True
        else:
            st.error(f"No se encontró la conciliación para la fecha {fecha_objetivo}")
            return False
            
    except Exception as e:
        st.error(f"Error al hacer clic en conciliación: {str(e)}")
        return False

def find_valor_a_pagar_alma(driver):
    """Buscar 'VALOR A PAGAR A COMERCIO' en Power BI ALMA"""
    try:
        titulo_selectors = [
            "//*[contains(text(), 'VALOR A PAGAR A COMERCIO')]",
            "//*[contains(text(), 'Valor a pagar a comercio')]",
            "//*[contains(text(), 'VALOR A PAGAR') and contains(text(), 'COMERCIO')]",
        ]
        
        titulo_element = None
        for selector in titulo_selectors:
            try:
                elementos = driver.find_elements(By.XPATH, selector)
                for elemento in elementos:
                    if elemento.is_displayed():
                        titulo_element = elemento
                        break
                if titulo_element:
                    break
            except:
                continue
        
        if not titulo_element:
            st.error("No se encontró 'VALOR A PAGAR A COMERCIO'")
            return None
        
        # Buscar en el contenedor
        try:
            container = titulo_element.find_element(By.XPATH, "./..")
            numeric_elements = container.find_elements(By.XPATH, ".//*")
            
            for elem in numeric_elements:
                texto = elem.text.strip()
                if texto and any(char.isdigit() for char in texto) and len(texto) < 50:
                    if texto != titulo_element.text:
                        return texto
        except:
            pass
        
        # Estrategia 2: elementos hermanos
        try:
            parent = titulo_element.find_element(By.XPATH, "./..")
            siblings = parent.find_elements(By.XPATH, "./*")
            
            for sibling in siblings:
                if sibling != titulo_element:
                    texto = sibling.text.strip()
                    if texto and any(char.isdigit() for char in texto):
                        return texto
        except:
            pass
        
        st.error("No se pudo encontrar el valor numérico")
        return None
        
    except Exception as e:
        st.error(f"Error buscando valor: {str(e)}")
        return None

def find_cantidad_pasos_alma(driver):
    """Buscar 'CANTIDAD DE PASOS' en Power BI ALMA"""
    try:
        titulo_selectors = [
            "//*[contains(text(), 'CANTIDAD DE PASOS')]",
            "//*[contains(text(), 'Cantidad de pasos')]",
            "//*[contains(text(), 'CANTIDAD') and contains(text(), 'PASOS')]",
        ]
        
        titulo_element = None
        for selector in titulo_selectors:
            try:
                elementos = driver.find_elements(By.XPATH, selector)
                for elemento in elementos:
                    if elemento.is_displayed():
                        titulo_element = elemento
                        break
                if titulo_element:
                    break
            except:
                continue
        
        if not titulo_element:
            st.error("No se encontró 'CANTIDAD DE PASOS'")
            return None
        
        # Buscar en el contenedor
        try:
            container = titulo_element.find_element(By.XPATH, "./..")
            numeric_elements = container.find_elements(By.XPATH, ".//*")
            
            for elem in numeric_elements:
                texto = elem.text.strip()
                if texto and any(char.isdigit() for char in texto) and len(texto) < 20:
                    if texto != titulo_element.text:
                        return texto
        except:
            pass
        
        # Estrategia 2: elementos hermanos
        try:
            parent = titulo_element.find_element(By.XPATH, "./..")
            siblings = parent.find_elements(By.XPATH, "./*")
            
            for sibling in siblings:
                if sibling != titulo_element:
                    texto = sibling.text.strip()
                    if texto and any(char.isdigit() for char in texto):
                        return texto
        except:
            pass
        
        st.error("No se pudo encontrar el valor de pasos")
        return None
        
    except Exception as e:
        st.error(f"Error buscando pasos: {str(e)}")
        return None

def extract_powerbi_data_alma(fecha_objetivo):
    """Función principal para extraer datos de Power BI ALMA"""
    
    REPORT_URL = "https://app.powerbi.com/view?r=eyJrIjoiMWExM2JkMzctMDgyMi00ZWZhLTgxODUtNGNlZGViYTcyM2NiIiwidCI6ImY5MTdlZDFiLWI0MDMtNDljNS1iODBiLWJhYWUzY2UwMzc1YSJ9"
    
    driver = setup_driver()
    if not driver:
        return None
    
    try:
        # Navegar al reporte
        with st.spinner("Conectando con Power BI ALMA..."):
            driver.get(REPORT_URL)
            time.sleep(10)
        
        driver.save_screenshot("powerbi_alma_inicial.png")
        
        # Hacer clic en la conciliación específica
        if not click_conciliacion_alma(driver, fecha_objetivo):
            return None
        
        time.sleep(3)
        driver.save_screenshot("powerbi_alma_despues_seleccion.png")
        
        # Buscar VALOR A PAGAR A COMERCIO
        valor_texto = find_valor_a_pagar_alma(driver)
        
        # Buscar CANTIDAD DE PASOS
        cantidad_pasos_texto = find_cantidad_pasos_alma(driver)
        
        driver.save_screenshot("powerbi_alma_final.png")
        
        return {
            'valor_texto': valor_texto,
            'cantidad_pasos_texto': cantidad_pasos_texto or 'No encontrado',
            'screenshots': {
                'inicial': 'powerbi_alma_inicial.png',
                'seleccion': 'powerbi_alma_despues_seleccion.png',
                'final': 'powerbi_alma_final.png'
            }
        }
        
    except Exception as e:
        st.error(f"Error durante la extracción: {str(e)}")
        return None
    finally:
        driver.quit()

# ===== FUNCIONES DE COMPARACIÓN =====

def convert_currency_to_float(currency_string):
    """Convierte string de moneda a float"""
    try:
        if isinstance(currency_string, (int, float)):
            return float(currency_string)
            
        if isinstance(currency_string, str):
            cleaned = currency_string.strip().replace('$', '').replace(' ', '')
            
            if '.' in cleaned and ',' in cleaned:
                cleaned = cleaned.replace('.', '').replace(',', '.')
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
    
    # Información del reporte
    st.sidebar.header("📋 Información del Reporte")
    st.sidebar.info("""
    **Objetivo:**
    - Cargar archivo Excel de ALMA
    - Extraer fecha, TOTAL y NUMERO DE REGISTROS
    - Comparar con Power BI automáticamente
    
    **Estado:** ✅ ChromeDriver Compatible
    **Versión:** v1.0 - ALMA
    """)
    
    # Estado del sistema
    st.sidebar.header("🛠️ Estado del Sistema")
    st.sidebar.success(f"✅ Python {sys.version_info.major}.{sys.version_info.minor}")
    st.sidebar.info(f"✅ Pandas {pd.__version__}")
    st.sidebar.info(f"✅ Streamlit {st.__version__}")
    
    # Cargar archivo Excel
    st.subheader("📁 Cargar Archivo Excel")
    uploaded_file = st.file_uploader(
        "Selecciona el archivo Excel de ALMA", 
        type=['xlsx', 'xls']
    )
    
    if uploaded_file is not None:
        # Extraer valores del Excel
        with st.spinner("📊 Procesando archivo Excel..."):
            valor_total, numero_registros, fecha_extraida = extract_excel_values_alma(uploaded_file)
        
        if valor_total is not None and numero_registros is not None:
            
            # Mostrar valores extraídos del Excel
            st.markdown("### 📊 Valores Extraídos del Excel")
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.metric("TOTAL", f"${valor_total:,.0f}".replace(",", "."))
            
            with col2:
                st.metric("NÚMERO DE REGISTROS", f"{numero_registros:,}".replace(",", "."))
            
            with col3:
                if fecha_extraida:
                    st.metric("FECHA", fecha_extraida)
                else:
                    st.warning("Fecha no detectada")
            
            st.markdown("---")
            
            # Extraer de Power BI
            if fecha_extraida:
                ejecutar_extraccion = st.button("🎯 Extraer de Power BI y Comparar", type="primary", use_container_width=True)
            else:
                st.warning("No se pudo extraer la fecha automáticamente del Excel")
                ejecutar_extraccion = False
            
            if ejecutar_extraccion:
                with st.spinner("🌐 Extrayendo datos de Power BI ALMA... Esto puede tomar 1-2 minutos"):
                    resultados = extract_powerbi_data_alma(fecha_extraida)
                    
                    if resultados and resultados.get('valor_texto'):
                        valor_powerbi_texto = resultados['valor_texto']
                        cantidad_pasos_powerbi = resultados.get('cantidad_pasos_texto', 'No encontrado')
                        
                        st.markdown("---")
                        
                        # Mostrar valores de Power BI
                        st.markdown("### 📊 Valores Extraídos de Power BI")
                        
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            st.metric("VALOR A PAGAR A COMERCIO (BI)", valor_powerbi_texto)
                        
                        with col2:
                            st.metric("CANTIDAD DE PASOS (BI)", cantidad_pasos_powerbi)
                        
                        st.markdown("---")
                        
                        # Comparación de Valores
                        st.markdown("### 💰 Validación: Valores")
                        
                        powerbi_valor, excel_valor, valor_formateado, coinciden_valor, dif_valor = compare_values_alma(
                            valor_powerbi_texto, valor_total
                        )
                        
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
                        
                        st.markdown("---")
                        
                        # Comparación de Pasos
                        st.markdown("### 👣 Validación: Número de Registros/Pasos")
                        
                        if cantidad_pasos_powerbi and cantidad_pasos_powerbi != 'No encontrado':
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
                        else:
                            st.warning("No se pudo extraer el número de pasos de Power BI")
                            coinciden_pasos = False
                        
                        st.markdown("---")
                        
                        # Resultado Final
                        st.markdown("### 📋 Resultado Final")
                        
                        if coinciden_valor and coinciden_pasos:
                            st.success("🎉 VALIDACIÓN EXITOSA - Valores y pasos coinciden")
                            st.balloons()
                        elif coinciden_valor and not coinciden_pasos:
                            st.warning("⚠️ VALIDACIÓN PARCIAL - Valores coinciden, pero hay diferencias en pasos")
                        elif not coinciden_valor and coinciden_pasos:
                            st.warning("⚠️ VALIDACIÓN PARCIAL - Pasos coinciden, pero hay diferencias en valores")
                        else:
                            st.error("❌ VALIDACIÓN FALLIDA - Existen diferencias en valores y pasos")
                        
                        # Botón para ver detalles
                        with st.expander("🔍 Ver Detalles Completos"):
                            st.markdown("#### 📊 Tabla Detallada de Comparación")
                            
                            resumen_data = []
                            
                            # Fila de valores
                            resumen_data.append({
                                'Concepto': 'VALOR TOTAL',
                                'Power BI': f"${powerbi_valor:,.0f}".replace(",", "."),
                                'Excel': f"${excel_valor:,.0f}".replace(",", "."),
                                'Estado': '✅ Coincide' if coinciden_valor else '❌ No coincide',
                                'Diferencia': f"${dif_valor:,.0f}".replace(",", "."),
                                'Dif. %': f"{(dif_valor/excel_valor*100):.2f}%" if excel_valor > 0 else "N/A"
                            })
                            
                            # Fila de pasos
                            if cantidad_pasos_powerbi and cantidad_pasos_powerbi != 'No encontrado':
                                resumen_data.append({
                                    'Concepto': 'REGISTROS/PASOS',
                                    'Power BI': f"{powerbi_pasos:,}".replace(",", "."),
                                    'Excel': f"{excel_pasos:,}".replace(",", "."),
                                    'Estado': '✅ Coincide' if coinciden_pasos else '❌ No coincide',
                                    'Diferencia': f"{dif_pasos:,}",
                                    'Dif. %': f"{(dif_pasos/excel_pasos*100):.2f}%" if excel_pasos > 0 else "N/A"
                                })
                            
                            df_resumen = pd.DataFrame(resumen_data)
                            st.dataframe(df_resumen, use_container_width=True, hide_index=True)
                            
                            # Screenshots
                            st.markdown("#### 📸 Capturas del Proceso")
                            col1, col2, col3 = st.columns(3)
                            screenshots = resultados.get('screenshots', {})
                            
                            if 'inicial' in screenshots and os.path.exists(screenshots['inicial']):
                                with col1:
                                    st.image(screenshots['inicial'], caption="Vista Inicial", use_column_width=True)
                            
                            if 'seleccion' in screenshots and os.path.exists(screenshots['seleccion']):
                                with col2:
                                    st.image(screenshots['seleccion'], caption="Tras Selección", use_column_width=True)
                            
                            if 'final' in screenshots and os.path.exists(screenshots['final']):
                                with col3:
                                    st.image(screenshots['final'], caption="Vista Final", use_column_width=True)
                    
                    elif resultados:
                        st.error("Se accedió al reporte pero no se encontraron los valores específicos")
                    else:
                        st.error("No se pudieron extraer datos del reporte Power BI")
        
        else:
            st.error("No se pudieron extraer valores del archivo Excel")
            with st.expander("💡 Sugerencias para solucionar el problema"):
                st.markdown("""
                - Verifica que el Excel sea de ALMA y tenga una única hoja
                - Asegúrate que la fila 2 contenga el texto con la fecha (ej: "REPORTE IP/REV 24 DE SEPTIEMBRE DEL 2025")
                - El archivo debe contener las palabras "TOTAL" y "NUMERO DE REGISTROS" o que en sus últimas filas aparezcan el total y el conteo.
                - Los valores pueden estar 1 o 2 columnas a la derecha de la etiqueta.
                """)
    
    else:
        st.info("📁 Por favor, carga un archivo Excel para comenzar la validación")

    # Información de ayuda
    st.markdown("---")
    with st.expander("ℹ️ Instrucciones de Uso"):
        st.markdown("""
        **Proceso:**
        1. **Cargar Excel**: Archivo ALMA con una única hoja
        2. **Extracción automática**: 
           - Busca la fecha en la fila 2
           - Busca "TOTAL" y trae el valor a la derecha (1-3 columnas)
           - Busca "NUMERO DE REGISTROS" y trae el valor a la derecha (1-3 columnas)
        3. **Extracción Power BI**: Navega a la conciliación ALMA de la fecha extraída
        4. **Comparación**: Compara VALOR A PAGAR A COMERCIO y CANTIDAD DE PASOS
        """)

if __name__ == "__main__":
    main()

    # Footer
    st.markdown("---")
    st.markdown('<div class="footer">💻 Desarrollado por Angel Torres | 🚀 Powered by Streamlit | v1.0 ALMA</div>', unsafe_allow_html=True)
