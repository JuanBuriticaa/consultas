import customtkinter as ctk
import threading
import os
import time
import csv
from tkinter import filedialog
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ========== CONFIGURACIÓN RÁPIDA DE LA INTERFAZ ========== #
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

root = ctk.CTk()
root.title("Consulta Express de Antecedentes y EPS")
root.geometry("500x550")  # Tamaño óptimo

# Marco principal (optimizado para carga rápida)
frame = ctk.CTkFrame(root)
frame.pack(pady=15, padx=15, fill="both", expand=True)

# Título simplificado
title_label = ctk.CTkLabel(frame, text="Consultas Automáticas", font=("Arial", 18, "bold"))
title_label.pack(pady=10)

# ========== FUNCIONALIDAD ACELERADA ========== #
def seleccionar_archivo():
    """Carga CSV sin retrasos"""
    filepath = filedialog.askopenfilename(filetypes=[("CSV", "*.csv")])
    if filepath:
        csv_entry.delete(0, "end")
        csv_entry.insert(0, filepath)
        log_message(f"✓ Archivo: {os.path.basename(filepath)}")

# Componentes de UI optimizados
csv_frame = ctk.CTkFrame(frame)
csv_frame.pack(pady=5, fill="x")

ctk.CTkLabel(csv_frame, text="CSV con cédulas:").grid(row=0, column=0, sticky="w")
csv_entry = ctk.CTkEntry(csv_frame)
csv_entry.grid(row=1, column=0, padx=5, sticky="ew")

ctk.CTkButton(
    csv_frame, 
    text="Examinar", 
    width=80,
    command=seleccionar_archivo
).grid(row=1, column=1, padx=5)

csv_frame.grid_columnconfigure(0, weight=1)

# Área de log (renderizado rápido)
log_text = ctk.CTkTextbox(frame, height=180, width=450)
log_text.pack(pady=10)

def log_message(msg):
    log_text.insert("end", f"{msg}\n")
    log_text.see("end")

# ========== CÓDIGO DE CONSULTAS TURBO ========== #
def leer_csv(archivo):
    """Carga cédulas con mínimo overhead"""
    try:
        with open(archivo, mode='r', encoding='utf-8') as f:
            return [row['cedula'].strip() for row in csv.DictReader(f) if row.get('cedula')]
    except Exception as e:
        log_message(f"✗ Error CSV: {str(e)}")
        return None

def antecedentes_express(driver, cedula, carpeta):
    """Flujo optimizado para antecedentes"""
    try:
        # Navegación rápida con esperas críticas
        driver.get("https://antecedentes.policia.gov.co:7005/WebJudicial/")
        WebDriverWait(driver, 8).until(  # Reducido de 20 a 8s
            EC.element_to_be_clickable((By.ID, "aceptaOption:0"))
        ).click()
        
        WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.ID, "continuarBtn"))
        ).click()
        
        campo = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.ID, "cedulaInput"))
        )
        campo.send_keys(cedula)
        
        # CAPTCHA (tiempo preservado)
        log_message("⌛ Resuelva CAPTCHA (10s)...")
        time.sleep(10)  
        
        WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//span[text()='Consultar']"))
        ).click()
        
        # Espera mínima para resultados
        time.sleep(3)  
        driver.save_screenshot(f"{carpeta}/antecedentes_{cedula}.png")
        return True
        
    except Exception as e:
        log_message(f"✗ Antecedentes: {str(e)[:50]}...")
        return False

def eps_express(driver, cedula, carpeta):
    """Consulta EPS ultrarrápida"""
    try:
        main_window = driver.current_window_handle
        driver.get("https://aplicaciones.adres.gov.co/bdua_internet/Pages/ConsultarAfiliadoWeb.aspx")
        
        # Flujo acelerado
        WebDriverWait(driver, 8).until(
            EC.presence_of_element_located((By.ID, "txtNumDoc"))
        ).send_keys(cedula)
        
        log_message("⌛ CAPTCHA EPS (8s)...")
        time.sleep(8)  # Mínimo razonable para CAPTCHA
        
        WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@value='Consultar']"))
        ).click()
        
        # Cambio a nueva pestaña rápido
        WebDriverWait(driver, 5).until(lambda d: len(d.window_handles) > 1)
        driver.switch_to.window([w for w in driver.window_handles if w != main_window][0])
        
        # Captura instantánea al cargar
        time.sleep(2)  # Espera crítica reducida
        driver.save_screenshot(f"{carpeta}/eps_{cedula}.png")
        driver.close()
        driver.switch_to.window(main_window)
        return True
        
    except Exception as e:
        log_message(f"✗ EPS: {str(e)[:50]}...")
        return False

def ejecutar_consultas_express():
    """Núcleo de ejecución optimizado"""
    if not (archivo := csv_entry.get()):
        log_message("✗ Seleccione CSV primero")
        return
    
    if not (cedulas := leer_csv(archivo)):
        return
    
    log_message("⚡ Iniciando navegador...")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    
    for cedula in cedulas:
        carpeta = os.path.join(os.path.expanduser("~/Downloads"), f"consulta_{cedula}")
        os.makedirs(carpeta, exist_ok=True)
        
        if antecedentes_express(driver, cedula, carpeta):
            eps_express(driver, cedula, carpeta)
    
    driver.quit()
    log_message("✅ Proceso completado")

# ========== LANZAMIENTO EN HILO ========== #
def iniciar_hilo():
    if threading.active_count() == 1:
        threading.Thread(target=ejecutar_consultas_express).start()
    else:
        log_message("⚠ Ya en ejecución")

ctk.CTkButton(
    frame,
    text="🚀 Ejecutar Consultas",
    command=iniciar_hilo
).pack(pady=10)

root.mainloop()