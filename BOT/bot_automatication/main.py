from shopify_automation import ShopifyAutomation
from sharepoint_uploader import SharePointConfig, SharePointUploader
import logging
import os
from dotenv import load_dotenv
import pandas as pd
from pathlib import Path

def setup_base_logging():
    """Configura el logging base para todo el proyecto"""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler('automation.log'),
            logging.StreamHandler()
        ]
    )

def validate_environment_vars():
    """Valida que todas las variables de entorno necesarias estén presentes"""
    required_vars = [
        'SHOPIFY_EMAIL',
        'SHOPIFY_PASSWORD',
        'SHAREPOINT_EMAIL',
        'SHAREPOINT_PASSWORD',
        'SHAREPOINT_BASE_URL',
        'SHAREPOINT_MAYORCA_URL',
        'SHAREPOINT_MONTERIA_URL'
    ]
    
    missing_vars = [var for var in required_vars if not os.getenv(var)]
    if missing_vars:
        raise ValueError(f"Faltan las siguientes variables de entorno: {', '.join(missing_vars)}")

def process_csv_file(file_path: Path) -> bool:
    """
    Procesa un archivo CSV de Shopify para eliminar filas donde todos los valores numéricos son 0
    o están vacíos.
    
    Args:
        file_path (Path): Ruta al archivo CSV
    Returns:
        bool: True si el proceso fue exitoso, False en caso contrario
    """
    try:
        logging.info(f"Procesando archivo: {file_path}")
        

        df = pd.read_csv(file_path, keep_default_na=True)
        
   
        numeric_columns = [
            'Quantity ordered',
            'Gross sales',
            'Discounts',
            'Total sales',
            'Net sales'
        ]
        
    
        mask = df[numeric_columns].replace({0: None}).notna().any(axis=1)
        

        filtered_df = df[mask]

        filtered_df.to_csv(file_path, index=False)
        

        rows_removed = len(df) - len(filtered_df)
        logging.info(f"Procesamiento completado para {file_path.name}:")
        logging.info(f"- Filas originales: {len(df)}")
        logging.info(f"- Filas después del filtrado: {len(filtered_df)}")
        logging.info(f"- Filas removidas: {rows_removed}")
        
        return True
        
    except Exception as e:
        logging.error(f"Error procesando {file_path}: {str(e)}")
        return False

def run_automation():
    """Ejecuta el proceso completo de automatización"""
    try:
        setup_base_logging()
        logging.info("Iniciando proceso de automatización")
   
        load_dotenv()
        validate_environment_vars()
        

        logging.info("Iniciando descarga de archivos de Shopify")
        

        logging.info("Procesando tienda POS (Monteria)")
        pos_automation = ShopifyAutomation('pos')
        pos_automation.run()
        

        logging.info("Procesando tienda Outlet (Mayorca)")
        outlet_automation = ShopifyAutomation('outlet')
        outlet_automation.run()
        
        logging.info("Iniciando procesamiento de archivos CSV para eliminar filas con valores en 0")
        csv_files = [
            Path("Monteria/2025.csv"),
            Path("Mayorca/2025.csv")
        ]
        
        for csv_file in csv_files:
            if csv_file.exists():
                if not process_csv_file(csv_file):
                    raise Exception(f"Error al procesar el archivo {csv_file}")
            else:
                logging.error(f"Archivo no encontrado: {csv_file}")
                raise FileNotFoundError(f"No se encontró el archivo: {csv_file}")
        

        logging.info("Iniciando proceso de carga a SharePoint")
        sharepoint_config = SharePointConfig(
            email=os.getenv('SHAREPOINT_EMAIL'),
            password=os.getenv('SHAREPOINT_PASSWORD'),
            base_url=os.getenv('SHAREPOINT_BASE_URL'),
            mayorca_folder_url=os.getenv('SHAREPOINT_MAYORCA_URL'),
            monteria_folder_url=os.getenv('SHAREPOINT_MONTERIA_URL')
        )
        
        uploader = SharePointUploader(sharepoint_config)
        uploader.upload_files()
        
        logging.info("Proceso completo finalizado exitosamente")
        
    except Exception as e:
        logging.error(f"Error en el proceso de automatización: {str(e)}")
        raise

if __name__ == "__main__":
    run_automation()