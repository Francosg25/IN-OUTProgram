import streamlit as st
import pandas as pd
import requests

# Configuración del Backend
API_URL = "http://localhost:8000/process-logistics"

def main():
    st.set_page_config(page_title="In-Out Logistics", layout="wide")
    st.title("Logistics Data Orchestrator (Frontend)")
    st.markdown("---")

    with st.sidebar:
        st.header("Bases de Datos Maestras")
        cost_file = st.file_uploader("Consolidado de Facturación (Costos)", type=["xlsx"])
        
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

    if st.button("Ejecutar Procesamiento (vía API)", use_container_width=True):
        if not cost_file:
            st.warning("Por favor, sube el archivo de Consolidado de Facturación.")
            return

        try:
            with st.spinner("🚀 Llamando al Microservicio de Procesamiento..."):
                # Preparar archivos para enviar por HTTP
                files = {
                    "cost_file": (cost_file.name, cost_file.getvalue(), cost_file.type)
                }
                if sea_file: files["sea_file"] = (sea_file.name, sea_file.getvalue(), sea_file.type)
                if land_file: files["land_file"] = (land_file.name, land_file.getvalue(), land_file.type)
                if outbound_file: files["outbound_file"] = (outbound_file.name, outbound_file.getvalue(), outbound_file.type)

                # Petición al Backend
                response = requests.post(API_URL, files=files)
                
                if response.status_code == 200:
                    data = response.json()
                    final_summary = pd.DataFrame(data["summary"])
                    recon = data["reconciliation"]
                    
                    st.success("✅ Procesamiento completado con éxito.")
                    
                    # --- AUDITORÍA ---
                    if recon:
                        st.markdown("### 🔍 Estado de Conciliación")
                        c1, c2, c3, c4 = st.columns(4)
                        c1.metric("Total Facturado", f"${recon['total_facturado']:,.2f}")
                        c2.metric("Total Asignado", f"${recon['total_asignado']:,.2f}")
                        diff = recon['diferencia']
                        c3.metric("Diferencia", f"${diff:,.2f}", delta=f"{diff:,.2f}", delta_color="normal" if abs(diff) < 1 else "inverse")
                        c4.metric("Match Rate", f"{recon['match_rate']:.1f}%")

                    # --- VISUALIZACIÓN ---
                    tabs = st.tabs(["📊 Resumen General", "🚢 Sea", "🚛 Land", "📦 Outbound"])
                    
                    with tabs[0]:
                        st.dataframe(final_summary, use_container_width=True)
                    
                    for i, label in enumerate(["Sea", "Land", "Outbound"], 1):
                        with tabs[i]:
                            df_f = final_summary[final_summary['Transport'] == label]
                            if not df_f.empty:
                                st.table(df_f)
                            else:
                                st.info(f"Sin datos para {label}")
                else:
                    st.error(f"Error en el Backend: {response.text}")

        except Exception as e:
            st.error(f"Fallo de conexión con la API: {e}")

if __name__ == "__main__":
    main()
