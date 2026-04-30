import pandas as pd
import numpy as np
from typing import List, Dict, Any, Callable

class DataValidatorEngine:
    def __init__(self, config: Dict[str, List[Dict[str, Any]]]):
        """
        Inicia el motor con una configuración de reglas.
        :param config: Diccionario donde la llave es el nombre de la columna 
                       y el valor es una lista de diccionarios con la función y argumentos.
        """
        self.config = config

    def process(self, df: pd.DataFrame) -> (pd.DataFrame, pd.DataFrame):
        """
        Ejecuta todas las validaciones definidas en la configuración.
        """
        # Trabajamos sobre una copia para no alterar los datos originales del usuario
        processed_df = df.copy()
        
        # Inicializamos columnas de control
        processed_df['is_valid'] = True
        processed_df['error_details'] = ""

        for column, rules in self.config.items():
            if column not in processed_df.columns:
                continue # O podrías registrar que falta una columna obligatoria

            for rule in rules:
                func: Callable = rule['func']
                args: list = rule.get('args', [])
                kwargs: dict = rule.get('kwargs', {})

                try:
                    # La función de validación debe devolver una Serie Booleana
                    # True = Pasa / False = Error
                    mask, error_msg = func(processed_df, column, *args, **kwargs)
                    
                    # Actualizar registros que fallaron
                    if not mask.all():
                        # Solo actualizamos las filas donde mask es False
                        processed_df.loc[~mask, 'is_valid'] = False
                        processed_df.loc[~mask, 'error_details'] += f"[{column}]: {error_msg} | "
                
                except Exception as e:
                    # Manejo de errores robusto para no detener el motor por una regla mal programada
                    processed_df['is_valid'] = False
                    processed_df['error_details'] += f"Error crítico en regla de {column}: {str(e)} | "

        # Separar los datos
        valid_data = processed_df[processed_df['is_valid']].drop(columns=['is_valid', 'error_details'])
        invalid_data = processed_df[~processed_df['is_valid']]

        return valid_data, invalid_data

