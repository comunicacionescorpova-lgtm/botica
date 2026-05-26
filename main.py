import os
import time
import pandas as pd
import json
import requests
import hashlib
import re
from datetime import datetime, date
from dotenv import load_dotenv


class SafeEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, (datetime, date, pd.Timestamp)):
            return obj.strftime("%Y-%m-%d") if not hasattr(obj, 'hour') else obj.strftime("%Y-%m-%d %H:%M:%S")
        if pd.isna(obj):
            return None
        return super().default(obj)

load_dotenv()

def normalize_url(url):
    """Convierte links de compartir de Google Drive, OneDrive o SharePoint en links de descarga directa."""
    if not isinstance(url, str):
        return url
        
    url = url.strip()
        
    # --- GOOGLE DRIVE ---
    gd_pattern = r'https://docs\.google\.com/spreadsheets/d/([a-zA-Z0-9-_]+)'
    gd_match = re.search(gd_pattern, url)
    if gd_match:
        file_id = gd_match.group(1)
        return f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx"
    
    # --- ONEDRIVE / SHAREPOINT ---
    # Si es un link de OneDrive/SharePoint y no tiene el parámetro de descarga
    if ("onedrive.live.com" in url or "1drv.ms" in url or "sharepoint.com" in url):
        if "download=1" not in url:
            separator = "&" if "?" in url else "?"
            return f"{url}{separator}download=1"
    
    return url

# Configuración
MASTER_EXCEL_URL = normalize_url(os.getenv("MASTER_EXCEL_URL")) # Auto-normalizar master
SOURCES_FILE = "sources.json"
REGISTRY_FILE = "registry.json"
CHANGELOG_FILE = "changelog.md"
OUTPUT_DIR = "outputs"

def get_file_hash(content):
    return hashlib.md5(content).hexdigest()

def load_registry():
    if os.path.exists(REGISTRY_FILE):
        with open(REGISTRY_FILE, 'r') as f:
            return json.load(f)
    return {}

def save_registry(registry):
    with open(REGISTRY_FILE, 'w') as f:
        json.dump(registry, f, indent=4)

def update_changelog(source_name, message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"| {timestamp} | {source_name} | {message} |\n"
    
    if not os.path.exists(CHANGELOG_FILE):
        with open(CHANGELOG_FILE, 'w', encoding='utf-8') as f:
            f.write("# 📜 Historial de Cambios (Changelog)\n\n")
            f.write("| Fecha/Hora | Proyecto | Descripción |\n")
            f.write("| --- | --- | --- |\n")
    
    with open(CHANGELOG_FILE, 'a', encoding='utf-8') as f:
        f.write(line)

def get_sources():
    """Obtiene la lista de fuentes desde el Excel Maestro o desde el archivo local."""
    if MASTER_EXCEL_URL:
        print(f"📥 Cargando configuración desde el Excel Maestro...")
        try:
            resp = requests.get(MASTER_EXCEL_URL)
            resp.raise_for_status()
            with open("master_config.xlsx", "wb") as f:
                f.write(resp.content)
            
            df_master = pd.read_excel("master_config.xlsx")
            sources = df_master.to_dict(orient='records')
            
            os.remove("master_config.xlsx")
            return sources
        except Exception as e:
            print(f"⚠️ Error cargando Excel Maestro: {e}. Usando backup local...")
    
    if os.path.exists(SOURCES_FILE):
        with open(SOURCES_FILE, 'r') as f:
            return json.load(f)
    
    return []

def process_source(source, registry):
    # Validar que los campos esenciales existan y no sean NaN
    if pd.isna(source.get('id')) or pd.isna(source.get('url')):
        return False
        
    source_id = str(source['id']).strip()
    source_name = source.get('name', source_id)
    url = normalize_url(source['url'])
    sheet = source.get('sheet', 'Catálogo')
    
    # Manejar SkipRows de forma segura
    try:
        skip_val = source.get('skiprows', 0)
        skiprows = int(skip_val) if pd.notna(skip_val) and str(skip_val).isdigit() else 0
    except:
        skiprows = 0
    
    print(f"\n🔍 Procesando: {source_name}...")
    
    try:
        # Parámetro anti-cache para forzar descarga fresca desde Google Drive / OneDrive
        separator = "&" if "?" in url else "?"
        fresh_url = f"{url}{separator}_t={int(time.time())}"
        response = requests.get(fresh_url, headers={"Cache-Control": "no-cache"})
        response.raise_for_status()
        content = response.content
        
        current_hash = get_file_hash(content)

        temp_file = f"temp_{source_id}.xlsx"
        with open(temp_file, 'wb') as f:
            f.write(content)
            
        # Intento inicial con skiprows
        df = pd.read_excel(temp_file, sheet_name=sheet, skiprows=skiprows)
        df.columns = [str(c).strip().upper() for c in df.columns]
        
        # Función para buscar la columna de código
        def find_codigo_col(columns):
            return next((c for c in columns if 'CODIGO' in c or 'CÓDIGO' in c), None)

        col_codigo = find_codigo_col(df.columns)
        
        # Si no se encuentra, buscar automáticamente en las primeras 20 filas
        if not col_codigo:
            print(f"⚠️ No se encontró 'CODIGO' en la fila {skiprows}. Buscando automáticamente...")
            full_df = pd.read_excel(temp_file, sheet_name=sheet, header=None, nrows=20)
            for i, row in full_df.iterrows():
                row_values = [str(v).strip().upper() for v in row.values]
                if find_codigo_col(row_values):
                    print(f"✨ ¡Encontrado! La cabecera real está en la fila {i + 1}.")
                    df = pd.read_excel(temp_file, sheet_name=sheet, skiprows=i)
                    df.columns = [str(c).strip().upper() for c in df.columns]
                    col_codigo = find_codigo_col(df.columns)
                    break
        
        if not col_codigo:
            print(f"❌ Error: No se pudo localizar la columna 'CODIGO' en las primeras 20 filas.")
            return False

        df = df[df[col_codigo].notna()]

        target_dir = os.path.join(OUTPUT_DIR, source_id)
        os.makedirs(target_dir, exist_ok=True)

        output_path = os.path.join(target_dir, "data.json")
        data = df.to_dict(orient='records')
        new_json = json.dumps(data, ensure_ascii=False, indent=4, cls=SafeEncoder)

        # Comparar con el JSON actual para evitar commits innecesarios
        existing_json = ""
        if os.path.exists(output_path):
            with open(output_path, 'r', encoding='utf-8') as f:
                existing_json = f.read()

        if new_json == existing_json:
            print(f"⏩ Sin cambios en los datos para {source_name}.")
            if os.path.exists(temp_file):
                os.remove(temp_file)
            return False

        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(new_json)

        registry[source_id] = {
            "name": source_name,
            "hash": current_hash,
            "last_updated": datetime.now().isoformat()
        }
        update_changelog(source_name, f"Actualización de stock ({len(data)} registros)")

        print(f"✅ ¡Actualizado! {output_path}")

        if os.path.exists(temp_file):
            os.remove(temp_file)
        return True

    except Exception as e:
        print(f"❌ Error en {source_name}: {str(e)}")
        return False

def main():
    sources = get_sources()
    if not sources:
        print("❌ No se encontraron fuentes de datos para procesar.")
        return

    registry = load_registry()
    changes_made = False

    for source in sources:
        source = {k.lower(): v for k, v in source.items()}
        if process_source(source, registry):
            changes_made = True

    if changes_made:
        save_registry(registry)
        print("\n✨ Proceso finalizado con actualizaciones.")
    else:
        print("\n😴 Sin cambios en los datos.")

if __name__ == "__main__":
    main()
