import pandas as pd
import logging
import re
from typing import Dict, Any

logger = logging.getLogger(__name__)

class LogisticsPipelineOrchestrator:
    def __init__(self, allocation_engine):
        self.allocation_engine = allocation_engin
          
        self.aliases = {
            'reference': ['REFERENCE', 'CONTAINER NUMBER', 'WAYBILL NUMBER', 'REFERENCIA', 'CONTAINER'],
            'bu': ['BU', 'OU', 'BUSINESS UNIT', 'UNIDAD DE NEGOCIO'],
            'gross_weight': ['GROSS WEIGHT (KGS)', 'TOTAL GROSS WEIGHT', 'PESO BRUTO (KGS)', 'WEIGHT'],
            'inbound': ['INBOUND'],
            'outbound': ['OUTBOUND'],
            'method': ['METHOD'],
            'part_number': ['NO DE PARTE', 'PART NUMBER', 'ITEM CODE']
        }
        
        self.valid_bus = {'M01', 'M02', 'M19', 'M23', 'M45'} 
    
    def clean_bu_code(self, raw_bu: Any) -> str:
             if pd.isna(raw_bu):
                 return 'DEFAULT_BU'
    
             # 1. Normalización básica
             clean_str = str(raw_bu).strip().upper()
             clean_str = re.sub(r'[^A-Z0-9]', '', clean_str)
    
             # 2. Validación de rango M00 - M1000
             match = re.match(r'^M(\d{1,4})$', clean_str)
             if match:
                num = int(match.group(1))
                if 0 <= num <= 1000:
                    return clean_str # Es una BU válida en el rango

             return clean_str if clean_str else 'DEFAULT_BU'
            
             clean_str = str(raw_bu).strip().upper()
        
             clean_str = re.sub(r'[^A-Z0-9]', '', clean_str)
        
             return clean_str if clean_str else 'DEFAULT_BU'
    

    def _extract_and_standardize(self, file_path: Any, source_type: str, transport_label: str) -> pd.DataFrame:
        try:
            df_raw = pd.read_excel(file_path) 
            df_raw.columns = df_raw.columns.astype(str).str.strip().str.upper()

            df_std = pd.DataFrame()
            mandatory_cols = ['reference', 'bu', 'gross_weight']
            
            for canon_col, possible_names in self.aliases.items():
                found = False
                for name in possible_names:
                    matched_cols = [c for c in df_raw.columns if name in c]
                    if matched_cols:
                        df_std[canon_col] = df_raw[matched_cols[0]]
                        found = True
                        break
                
                if not found:
                    if canon_col in mandatory_cols:
                        if canon_col == 'bu':
                            df_std['bu'] = 'DEFAULT_BU'
                        else:
                            logger.error(f"Falta columna crítica '{canon_col}' en {transport_label}")
                            return pd.DataFrame()
                    else:
                        # Columnas opcionales (como inbound, method, part_number)
                        df_std[canon_col] = None

            if 'bu' in df_std.columns:
                df_std['bu'] = df_std['bu'].apply(self.clean_bu_code)

            def apply_business_rules(row):
                part_val = str(row['part_number']).upper() if row['part_number'] else ""
                
                # 1. REGLA CAPEX: Si contiene "CAPEX" o NO tiene números (solo letras/guiones)
                if "CAPEX" in part_val or (part_val and part_val != 'NONE' and not any(char.isdigit() for char in part_val)):
                    return "Capex"
                
                # 2. REGLA MISCELANEUS (Heurística de Descripción):
                # Si el "Número de Parte" es en realidad una descripción larga (ej. ZOWIETEK 5-50MM...)
                # Criterios: Más de 3 espacios O longitud > 25 caracteres con al menos 2 espacios.
                num_spaces = part_val.count(' ')
                if num_spaces > 3 or (len(part_val) > 25 and num_spaces >= 2):
                    return 'Miscelaneus'
                
                # 3. REGLA MISCELANEUS (Palabras clave tradicionales)
                miscelaneus_keywords = ['TAPA PLASTICA', 'CHAROLA', 'BASE PLASTICA', 'TAPA', 'BASE']
                for kw in miscelaneus_keywords:
                    if kw in part_val:
                        return 'Miscelaneus'
                
                return row['bu']
            
            if 'part_number' in df_std.columns and 'bu' in df_std.columns:
                df_std['bu'] = df_std.apply(apply_business_rules, axis=1)

            df_std['transport_type'] = transport_label
            df_std['gross_weight'] = pd.to_numeric(df_std['gross_weight'], errors='coerce').fillna(0)

            return df_std

        except Exception as e:
            logger.error(f"Error estandarizando fuente {transport_label}: {e}")
            return pd.DataFrame()

    def run_pipeline(self, files: Dict[str, Dict[str, str]], df_costs: pd.DataFrame) -> pd.DataFrame:
        standardized_dfs = []
        
        for file_obj, config in files.items():
            df = self._extract_and_standardize(
                file_path=file_obj, 
                source_type=config['type'], 
                transport_label=config['label']
            )
            if not df.empty:
                standardized_dfs.append(df)

        if not standardized_dfs:
            raise ValueError("No se pudieron extraer datos válidos de ningún archivo. Verifica los nombres de las columnas.")

        # Unificación
        df_unified = pd.concat(standardized_dfs, ignore_index=True)
        
        # Procesamiento final (Prorrateo)
        summary_df = self.allocation_engine.calculate_outbound(df_unified, df_costs)
        
        return summary_df