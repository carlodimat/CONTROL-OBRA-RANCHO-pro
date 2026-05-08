import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io
import os
from datetime import datetime

# 1. CONFIGURACIÓN DE PÁGINA
st.set_page_config(page_title="DIMAQUINAS C.A. - RANCHO FLAMBOYANT PRO", layout="wide", page_icon="🚀")

# 2. DISEÑO CSS
st.markdown("""
    <style>
    .stMetric { border: 1px solid #1e3a8a; padding: 15px; border-radius: 12px; background: #f8fafc; }
    .header-box { background-color: #1e3a8a; color: white; padding: 80px 20px; border-radius: 15px; margin-bottom: 30px; text-align: center; }
    .title-text { font-weight: 900; margin: 0; text-transform: uppercase; line-height: 1.1; }

    @media (max-width: 600px) {
        .title-text { font-size: 35px !important; }
        .subtitle-text { font-size: 18px !important; }
        .header-box { padding: 40px 10px !important; }
    }
    @media (min-width: 601px) {
        .title-text { font-size: 100px; }
        .subtitle-text { font-size: 35px; }
    }
    html, body, [class*="st-"] { color: #000000; font-weight: 500; }
    </style>
    """, unsafe_allow_html=True)

# ─── Funciones de Utilidad ───
def create_excel(df_report):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_report.to_excel(writer, index=False, sheet_name='Reporte')
    return output.getvalue()

def load_all_data():
    csv_name = "RANCHO.csv"
    if not os.path.exists(csv_name):
        # Intentar con otras opciones si el principal no existe
        options = ["DIMAQUINAS C.A._RANCHO FLAMBOYANT (3).csv", "DIMAQUINAS_C.A._RANCHO_FLAMBOYANT.csv"]
        for opt in options:
            if os.path.exists(opt):
                csv_name = opt
                break
    
    try:
        df = pd.read_csv(csv_name)
        df['FECHA'] = pd.to_datetime(df['FECHA'])
        cols_fin = ['MONTO BASE USD', 'MONTO PAGADO', 'HONORARIOS', 'COSTO TOTAL', '% ADMIN', 'MONTO ORIG', 'TASA']
        for col in cols_fin:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        return df, csv_name
    except Exception as e:
        st.error(f"Error cargando datos: {e}")
        return None, None

df, active_csv = load_all_data()

if df is not None:
    # --- VARIABLES INICIALES ---
    empresa_default = df['EMPRESA'].iloc[0] if 'EMPRESA' in df.columns else "DIMAQUINAS C.A."
    obra_default    = df['OBRA'].iloc[0]    if 'OBRA'    in df.columns else "RANCHO FLAMBOYANT"
    
    df_gastos_base = df[df['CLASE'] == 'GASTO'].copy()
    df_ingresos    = df[df['CLASE'] == 'INGRESO'].copy()

    # --- BARRA LATERAL DE FILTROS ---
    st.sidebar.header("🎯 FILTROS DE OBRA")
    tipos_sel = st.sidebar.multiselect("Filtrar por TIPO:",      options=sorted(df_gastos_base['TIPO'].unique()))
    areas_sel = st.sidebar.multiselect("Filtrar por ÁREA:",      options=sorted(df_gastos_base['AREA'].unique()))
    prov_sel  = st.sidebar.multiselect("Filtrar por PROVEEDOR:", options=sorted(df_gastos_base['PROVEEDOR'].unique()))

    filtro_activo = bool(tipos_sel or areas_sel or prov_sel)

    # Aplicar filtros
    df_gastos = df_gastos_base.copy()
    if tipos_sel: df_gastos = df_gastos[df_gastos['TIPO'].isin(tipos_sel)]
    if areas_sel: df_gastos = df_gastos[df_gastos['AREA'].isin(areas_sel)]
    if prov_sel:  df_gastos = df_gastos[df_gastos['PROVEEDOR'].isin(prov_sel)]

    # --- CÁLCULOS ---
    total_ing        = df_ingresos['MONTO BASE USD'].sum()
    total_neto       = df_gastos['MONTO BASE USD'].sum()
    total_honorarios = df_gastos['HONORARIOS'].sum()
    gasto_total_real = total_neto + total_honorarios

    _hon_base       = df_gastos_base['HONORARIOS'].sum() if 'HONORARIOS' in df_gastos_base.columns else 0
    neto_base       = df_gastos_base['MONTO BASE USD'].sum()
    gasto_base_real = neto_base + _hon_base
    saldo_base_real = total_ing - gasto_base_real

    # --- ENCABEZADO ---
    st.markdown(
        f'<div class="header-box">'
        f'<p class="title-text">{empresa_default}</p>'
        f'<p class="subtitle-text">OBRA: {obra_default} <span style="font-size:15px; color:#ffdd00;">[PRO]</span></p>'
        f'</div>',
        unsafe_allow_html=True
    )

    # --- MÉTRICAS ---
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("TOTAL INGRESOS",   f"$ {total_ing:,.2f}")
    m2.metric("GASTOS NETOS",     f"$ {neto_base:,.2f}")
    m3.metric("ADMIN. DELEGADA",  f"$ {_hon_base:,.2f}")
    m4.metric("SALDO REAL",       f"$ {saldo_base_real:,.2f}")

    st.divider()

    # --- RESUMEN GLOBAL DE FILTROS ---
    if filtro_activo:
        filtros_desc = []
        if tipos_sel: filtros_desc.append(f"Tipos: {', '.join(tipos_sel)}")
        if areas_sel: filtros_desc.append(f"Áreas: {', '.join(areas_sel)}")
        if prov_sel:  filtros_desc.append(f"Proveedores: {', '.join(prov_sel)}")
        lbl_filtros = " | ".join(filtros_desc)
        st.info(f"🔍 **Filtro activo:** {lbl_filtros} | {len(df_gastos)} registros")

    # ──────────────────────────────────────────────────────────
    # TABS
    # ──────────────────────────────────────────────────────────
    t1, t2, t3, t4, t5 = st.tabs(["📊 GRÁFICOS", "💸 EGRESOS", "💰 INGRESOS", "🔍 BUSCADOR", "🚀 CARGA MASIVA"])

    with t1:
        # Gráficos (Lógica idéntica a versión estándar)
        st.write("### 📌 Inversión por Tipo")
        df_t = df_gastos.groupby('TIPO')['MONTO BASE USD'].sum().reset_index()
        df_t = pd.concat([df_t, pd.DataFrame({'TIPO': ['ADMINISTRACIÓN DELEGADA'], 'MONTO BASE USD': [total_honorarios]})], ignore_index=True)
        fig_t = px.bar(df_t, x='MONTO BASE USD', y='TIPO', orientation='h', color='MONTO BASE USD', color_continuous_scale='Viridis')
        st.plotly_chart(fig_t, use_container_width=True)

    with t2:
        st.subheader("📝 Detalle de Gastos")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Registros", f"{len(df_gastos)}")
        c2.metric("Total Neto", f"$ {total_neto:,.2f}")
        c3.metric("Admin. Delegada", f"$ {total_honorarios:,.2f}")
        c4.metric("TOTAL FILTRADO", f"$ {gasto_total_real:,.2f}")
        st.divider()
        cols_show = [c for c in ['FECHA', 'TIPO', 'AREA', 'PROVEEDOR', 'DESCRIPCION', 'MONTO ORIG', '% ADMIN', 'HONORARIOS', 'COSTO TOTAL'] if c in df_gastos.columns]
        st.dataframe(df_gastos[cols_show].sort_values('FECHA', ascending=False), use_container_width=True)

    with t3:
        st.subheader("💰 Detalle de Ingresos")
        st.dataframe(df_ingresos.sort_values('FECHA', ascending=False), use_container_width=True)

    with t5:
        st.subheader("🚀 Migrador Mágico: Carga Masiva desde Excel")
        st.markdown("""
        1. Copia tus datos desde Excel (sin encabezados).
        2. El orden de columnas debe ser: **FECHA | CLASE | TIPO | AREA | PROVEEDOR | DESCRIPCION | MONTO ORIG | TASA | % ADMIN | EMPRESA | OBRA**
        3. Pega los datos abajo (Ctrl+V).
        """)
        
        # Plantilla vacía para el editor
        template = pd.DataFrame(columns=[
            'FECHA', 'CLASE', 'TIPO', 'AREA', 'PROVEEDOR', 'DESCRIPCION', 'MONTO ORIG', 'TASA', '% ADMIN', 'EMPRESA', 'OBRA'
        ])
        
        # Editor de datos
        edited_df = st.data_editor(
            template,
            num_rows="dynamic",
            use_container_width=True,
            column_config={
                "CLASE": st.column_config.SelectboxColumn(options=["GASTO", "INGRESO"]),
                "FECHA": st.column_config.DateColumn(),
                "MONTO ORIG": st.column_config.NumberColumn(format="$ %.2f"),
                "TASA": st.column_config.NumberColumn(format="%.4f"),
                "% ADMIN": st.column_config.NumberColumn(format="%.2f%%"),
            }
        )
        
        if st.button("💾 Procesar y Guardar en " + active_csv):
            if not edited_df.empty:
                try:
                    # Validar y Limpiar
                    new_rows = edited_df.copy()
                    new_rows['FECHA'] = pd.to_datetime(new_rows['FECHA'])
                    new_rows['MONTO ORIG'] = pd.to_numeric(new_rows['MONTO ORIG'])
                    new_rows['TASA'] = pd.to_numeric(new_rows['TASA'])
                    new_rows['% ADMIN'] = pd.to_numeric(new_rows['% ADMIN'])
                    
                    # Cálculos automáticos
                    new_rows['MONTO BASE USD'] = new_rows['MONTO ORIG'] / new_rows['TASA']
                    new_rows['HONORARIOS'] = new_rows['MONTO BASE USD'] * (new_rows['% ADMIN'] / 100)
                    new_rows['COSTO TOTAL'] = new_rows['MONTO BASE USD'] + new_rows['HONORARIOS']
                    new_rows['MONTO PAGADO'] = new_rows['MONTO BASE USD'] # Asunción por defecto
                    
                    # Asegurar que todas las columnas del CSV original existan
                    for col in df.columns:
                        if col not in new_rows.columns:
                            new_rows[col] = 0
                    
                    # Concatenar y guardar
                    updated_df = pd.concat([df, new_rows[df.columns]], ignore_index=True)
                    updated_df.to_csv(active_csv, index=False)
                    
                    st.success(f"✅ ¡{len(new_rows)} registros añadidos exitosamente!")
                    st.balloons()
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Error al procesar datos: {e}")
            else:
                st.warning("⚠️ No hay datos para procesar. Pega información en la tabla arriba.")

else:
    st.error("No se pudieron cargar los datos.")
