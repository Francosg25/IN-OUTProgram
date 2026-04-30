import pandas as pd
import logging
from typing import Dict, Any

logger = logging.getLogger(__name__)

class LogisticsPipelineOrchestrator:
    """
    Orquesta la extracción, transformación (estandarización) y carga (cálculo)
    de múltiples fuentes de datos logísticas.
    """
    def __init__(self, allocation_engine):
        self.allocation_engine = allocation_engine
        
        # Mapa de estandarización: Transforma columnas de origen al Modelo Canónico
        self.schema_map = {
            "expos": {
                "Transport Name": "carrier",
                "Container Number": "reference",
                "BU": "bu",
                "Gross Weight (Kgs)": "gross_weight" # Ajustar si hay espacios extra
            },
            "impos_land": {
                "Reference": "reference",
                "BU": "bu",
                "Peso Bruto (Kgs)": "gross_weight"
            },
            "china_sea": {
                "Container Number": "reference",
                "BU": "bu",
                "Total Gross Weight": "gross_weight"
            }
        }

    def _extract_and_standardize(self, file_path: str, source_type: str, transport_label: str) -> pd.DataFrame:
        """
        Lee el archivo, extrae las columnas configuradas y las renombra al estándar.
        """
        try:
            # NOTA: Aquí usarías el SmartExcelExtractor o MarkerDrivenExtractor 
            # que diseñamos previamente para limpiar la basura superior del Excel.
            # Para este ejemplo, asumimos lectura directa de la hoja principal.
            df_raw = pd.read_excel(file_path) 
            
            mapping = self.schema_map.get(source_type)
            if not mapping:
                raise ValueError(f"Tipo de fuente '{source_type}' no configurada.")

            # Filtrar solo las columnas que existen y necesitamos
            cols_to_keep = [col for col in mapping.keys() if col in df_raw.columns]
            df_std = df_raw[cols_to_keep].copy()
            
            # Renombrar al Modelo Canónico
            df_std.rename(columns=mapping, inplace=True)
            
            # Etiquetar el tipo de transporte (Sea, Land, Outbound)
            df_std['transport_type'] = transport_label
            
            # Limpieza básica: Asegurar que el peso es numérico
            if 'gross_weight' in df_std.columns:
                df_std['gross_weight'] = pd.to_numeric(df_std['gross_weight'], errors='coerce').fillna(0)

            return df_std

        except Exception as e:
            logger.error(f"Error estandarizando {file_path} ({source_type}): {e}")
            return pd.DataFrame()

    def run_pipeline(self, files: Dict[str, Dict[str, str]], df_costs: pd.DataFrame) -> pd.DataFrame:
        """
        Ejecuta el pipeline completo y retorna el Summary.
        :param files: Dict con formato { 'ruta_archivo': {'type': 'expos', 'label': 'Outbound'} }
        """
        standardized_dfs = []
        
        # 1. Extracción y Adaptación
        for file_path, config in files.items():
            df = self._extract_and_standardize(
                file_path, 
                source_type=config['type'], 
                transport_label=config['label']
            )
            if not df.empty:
                standardized_dfs.append(df)

        if not standardized_dfs:
            raise ValueError("No se pudieron extraer datos válidos de ningún archivo.")

        # 2. Unificación (Creación del Modelo Canónico completo)
        # pd.concat alinea automáticamente las columnas con el mismo nombre
        df_unified = pd.concat(standardized_dfs, ignore_index=True)
        
        logger.info(f"Datos unificados: {len(df_unified)} transacciones logísticas listas para procesar.")

        # 3. Procesamiento y Asignación de Costos
        # Pasamos el bloque maestro al motor que creamos en el paso anterior
        summary_df = self.allocation_engine.calculate_outbound(df_unified, df_costs)
        
        return summary_df