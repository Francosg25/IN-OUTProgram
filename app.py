import streamlit as st
import sys
from pathlib import Path

# --- ANCLAJE DE RUTAS ROBUSTO ---
# Obtenemos la ruta absoluta de la carpeta raíz (IN-OUTProgram)
ROOT_DIR = Path(__file__).resolve().parent

# Agregamos la raíz al principio de sys.path
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

try:
    from services.orchestrator import LogisticsPipelineOrchestrator
except ModuleNotFoundError as e:
    st.error(f"No se encontró el módulo: {e.name}")
    st.info(f"Buscando en: {ROOT_DIR}")
    # Listamos directorios para debug
    st.code(f"Directorios en la raíz: {[d.name for d in ROOT_DIR.iterdir() if d.is_dir()]}")
    st.stop()

# Asegúrate de tener estas importaciones al inicio de tu app.py
from engine.allocation import CostAllocationEngine  # Ajusta el nombre de la clase/ruta si es distinto
from services.orchestrator import LogisticsPipelineOrchestrator

def main():
    st.title("Logistics Data Orchestrator")
    st.markdown("---")

    try:
        # Instanciamos la dependencia (el motor matemático)
        # Puedes cambiar 'weight' por 'full_container' según tu lógica de negocio
        engine_instance = CostAllocationEngine(allocation_type='weight') 

        # 2. Inyectamos la dependencia al constructor del orquestador
        orchestrator = LogisticsPipelineOrchestrator(allocation_engine=engine_instance)
        
    except Exception as e:
        st.error(f"Error inicializando los servicios core: {e}")
        st.stop()
    
    orchestrator = LogisticsPipelineOrchestrator()

    # Sidebar para carga de archivos maestros (Bases de Datos / Consolidados)
    with st.sidebar:
        st.header("Bases de Datos Maestras")
        cost_file = st.file_uploader("Consolidado de Facturación (Costos)", type=["xlsx"])
        
    # Área principal para archivos de operación
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        st.subheader("Sea / China")
        sea_file = st.file_uploader("Reporte Marítimo", type=["xlsx"], key="sea")
    with col_b:
        st.subheader("Land / Impos")
        land_file = st.file_uploader("Impos Land", type=["xlsx"], key="land")
    with col_c:
        st.subheader("Outbound / Expos")
        outbound_file = st.file_uploader("Reporte Expos", type=["xlsx", "xls"], key="out")

    st.markdown("---")

    # Lógica de procesamiento
    if st.button("⚡ Ejecutar Procesamiento y Prorrateo", use_container_width=True):
        if not cost_file:
            st.warning("Por favor, sube el archivo de Consolidado de Facturación en la barra lateral.")
            return

        # Construcción del diccionario de archivos para el orquestador
        # Filtramos solo los que el usuario haya subido
        files_to_process = {}
        if sea_file: files_to_process['Sea'] = sea_file
        if land_file: files_to_process['Land'] = land_file
        if outbound_file: files_to_process['Outbound'] = outbound_file

        if not files_to_process:
            st.error("No hay archivos de operación cargados para procesar.")
            return

        try:
            with st.spinner("🔄 Procesando reglas de negocio y reconciliación de costos..."):
                # 1. Cargar el consolidado de costos
                df_costs = pd.read_excel(cost_file)
                
                # 2. Ejecutar Pipeline completo a través del Orquestador
                final_summary = orchestrator.run_full_process(files_to_process, df_costs)
                
                # 3. Mostrar resultados
                st.success("✅ Procesamiento completado con éxito.")
                
                # Pestañas para organizar la visualización
                tab_summary, tab_details = st.tabs(["📊 Resumen (Summary)", "📝 Detalles por Transacción"])
                
                with tab_summary:
                    st.table(final_summary)
                    st.download_button(
                        "Descargar Summary (CSV)",
                        data=final_summary.to_csv(index=False),
                        file_name="summary_logistica.csv",
                        mime="text/csv"
                    )
                
                with tab_details:
                    st.write("Vista previa de datos procesados:")
                    # Aquí podrías mostrar el master_df si el orquestador lo devuelve
                    st.info("Desglose detallado disponible para auditoría.")

        except Exception as e:
            st.error(f"Error durante la ejecución: {str(e)}")
            st.exception(e) # Solo usar durante desarrollo/debugging

if __name__ == "__main__":
    main()