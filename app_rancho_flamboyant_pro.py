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
def wrap_label(text, width=18):
    if not isinstance(text, str): return str(text)
    words = text.split()
    lines, current = [], []
    for word in words:
        if sum(len(w) for w in current) + len(current) + len(word) > width:
            if current: lines.append(" ".join(current))
            current = [word]
        else: current.append(word)
    if current: lines.append(" ".join(current))
    return "<br>".join(lines)

def create_excel(df_report):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_report.to_excel(writer, index=False, sheet_name='Reporte')
    return output.getvalue()

def horizontal_bar_chart(df_plot, x_col, y_col, color_scale, title, height=500):
    if df_plot.empty:
        st.info("No hay datos para este gráfico.")
        return
    df_sorted = df_plot.sort_values(x_col, ascending=True).copy()
    labels = [f"$ {v:,.2f}" for v in df_sorted[x_col]]
    vals = df_sorted[x_col].values
    max_v = vals.max() if vals.max() > 0 else 1
    norm  = vals / max_v
    import plotly.colors as pc
    palette = pc.get_colorscale(color_scale)
    bar_colors = pc.sample_colorscale(palette, norm)
    fig = go.Figure(go.Bar(
        x=df_sorted[x_col], y=df_sorted[y_col], orientation='h',
        marker_color=bar_colors, text=labels, textposition='outside',
        textfont=dict(size=11, color='#1e3a8a', family='Arial Black'),
        cliponaxis=False,
    ))
    fig.update_layout(
        title=dict(text=title, font=dict(size=14, color='#1e3a8a'), x=0.01),
        height=height, margin=dict(l=10, r=130, t=50, b=30),
        xaxis=dict(showticklabels=False, showgrid=False, zeroline=False, range=[0, max_v * 1.35]),
        yaxis=dict(tickfont=dict(size=11, color='#000000'), showgrid=False),
        plot_bgcolor='#f8fafc', paper_bgcolor='#ffffff', showlegend=False,
    )
    st.plotly_chart(fig, use_container_width=True)

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
        st.write("### 📌 Inversión por Tipo")
        df_t = df_gastos.groupby('TIPO')['MONTO BASE USD'].sum().reset_index()
        df_t = pd.concat([df_t, pd.DataFrame({'TIPO': ['ADMINISTRACIÓN DELEGADA'], 'MONTO BASE USD': [total_honorarios]})], ignore_index=True)
        horizontal_bar_chart(df_t, 'MONTO BASE USD', 'TIPO', 'Viridis', '📌 Inversión total por Tipo de Gasto', height=max(350, len(df_t) * 45))

        st.divider()
        st.write("### 📐 Inversión por Área")
        df_a = df_gastos.groupby('AREA')['MONTO BASE USD'].sum().reset_index()
        df_a = pd.concat([df_a, pd.DataFrame({'AREA': ['ADMINISTRACIÓN DELEGADA'], 'MONTO BASE USD': [total_honorarios]})], ignore_index=True)
        horizontal_bar_chart(df_a, 'MONTO BASE USD', 'AREA', 'Blues', '📐 Inversión total por Área de Obra', height=max(400, len(df_a) * 42))

        st.divider()
        st.write("### 👥 Top Proveedores")
        df_p = (df_gastos.groupby('PROVEEDOR')['MONTO BASE USD'].sum().sort_values(ascending=False).head(20).reset_index())
        df_p = pd.concat([df_p, pd.DataFrame({'PROVEEDOR': ['ADMINISTRACIÓN DELEGADA'], 'MONTO BASE USD': [total_honorarios]})], ignore_index=True)
        horizontal_bar_chart(df_p, 'MONTO BASE USD', 'PROVEEDOR', 'Reds', '👥 Top 20 Proveedores por Gasto', height=max(500, len(df_p) * 40))

        st.divider()
        st.write("### 📅 Evolución Acumulativa")
        fecha_inicio = df_ingresos['FECHA'].min() if not df_ingresos.empty else df_gastos_base['FECHA'].min()
        idx_full = pd.date_range(start=fecha_inicio, end=pd.Timestamp.today(), freq='D')
        s_gastos = (df_gastos_base.set_index('FECHA')['COSTO TOTAL'].resample('D').sum().reindex(idx_full, fill_value=0).cumsum())
        s_ingresos = (df_ingresos.set_index('FECHA')['MONTO BASE USD'].resample('D').sum().reindex(idx_full, fill_value=0).cumsum())
        fig_time = go.Figure()
        fig_time.add_trace(go.Scatter(x=idx_full, y=s_ingresos, name='Ingresos', fill='tozeroy', line=dict(color='#22c55e')))
        fig_time.add_trace(go.Scatter(x=idx_full, y=s_gastos, name='Gastos', fill='tozeroy', line=dict(color='#1e3a8a')))
        fig_time.update_layout(height=400, margin=dict(l=10, r=10, t=30, b=30), plot_bgcolor='#f8fafc')
        st.plotly_chart(fig_time, use_container_width=True)

    with t2:
        st.subheader("📝 Detalle de Gastos")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Registros", f"{len(df_gastos)}")
        c2.metric("Total Neto", f"$ {total_neto:,.2f}")
        c3.metric("Admin. Delegada", f"$ {total_honorarios:,.2f}")
        c4.metric("TOTAL FILTRADO", f"$ {gasto_total_real:,.2f}")
        st.divider()
        cols_show = ['FECHA', 'TIPO', 'AREA', 'PROVEEDOR', 'FORMA DE PAGO', 'DESCRIPCION', 'MONTO ORIG', '% ADMIN', 'HONORARIOS', 'COSTO TOTAL']
        cols_show = [c for c in cols_show if c in df_gastos.columns]
        st.dataframe(df_gastos[cols_show].sort_values('FECHA', ascending=False).style.format({
            'MONTO ORIG': "{:,.2f}", 'HONORARIOS': "${:,.2f}", 'COSTO TOTAL': "${:,.2f}", '% ADMIN': "{:.1f}%"
        }), use_container_width=True)

    with t3:
        st.subheader("💰 Detalle de Ingresos")
        st.dataframe(df_ingresos.sort_values('FECHA', ascending=False), use_container_width=True)

    with t5:
        st.subheader("🚀 Carga de Gastos: Pegar desde Excel")
        st.markdown("""
        1. Copia tus datos desde Excel (sin encabezados).
        2. Orden: **FECHA | TIPO | AREA | PROVEEDOR | DESCRIPCION | MONTO ORIG | TASA | % ADMIN | ESTADO | FORMA DE PAGO**
        3. Pega abajo (Ctrl+V). *Estado debe ser 'PAGADO' o 'POR PAGAR'.*
        """)
        
        template = pd.DataFrame(columns=[
            'FECHA', 'TIPO', 'AREA', 'PROVEEDOR', 'DESCRIPCION', 'MONTO ORIG', 'TASA', '% ADMIN', 'ESTADO', 'FORMA DE PAGO'
        ])
        
        edited_df = st.data_editor(
            template, num_rows="dynamic", use_container_width=True,
            column_config={
                "FECHA": st.column_config.DateColumn(format="DD/MM/YYYY"),
                "MONTO ORIG": st.column_config.NumberColumn(format="$ %.2f"),
                "TASA": st.column_config.NumberColumn(format="%.4f"),
                # Quitamos los Selectbox restrictivos para que el pegado sea más fluido
                "% ADMIN": st.column_config.TextColumn(help="Ingresa solo el número (ej: 15)"),
                "ESTADO": st.column_config.TextColumn(help="PAGADO o POR PAGAR"),
                "FORMA DE PAGO": st.column_config.TextColumn(help="TRANSFERENCIA BANCARIA, EFECTIVO, etc."),
            }
        )
        
        if st.button("💾 Procesar y Guardar Gastos"):
            if not edited_df.empty:
                try:
                    new_rows = edited_df.copy()
                    # 1. Limpieza de Fechas
                    new_rows['FECHA'] = pd.to_datetime(new_rows['FECHA'], errors='coerce')
                    new_rows = new_rows.dropna(subset=['FECHA'])
                    
                    # 2. Cálculos y Valores por Defecto
                    new_rows['MONTO ORIG'] = pd.to_numeric(new_rows['MONTO ORIG'], errors='coerce').fillna(0)
                    new_rows['TASA'] = pd.to_numeric(new_rows['TASA'], errors='coerce').fillna(1)
                    
                    # Manejo de % ADMIN (Default 15%)
                    new_rows['% ADMIN'] = new_rows['% ADMIN'].astype(str).str.replace('%', '').replace('nan', '')
                    new_rows['% ADMIN'] = pd.to_numeric(new_rows['% ADMIN'], errors='coerce').fillna(15)
                    
                    # Manejo de ESTADO (Default PAGADO)
                    new_rows['ESTADO'] = new_rows['ESTADO'].fillna('').replace('', 'PAGADO').str.upper()
                    
                    # Manejo de FORMA DE PAGO (Default TRANSFERENCIA BANCARIA)
                    new_rows['FORMA DE PAGO'] = new_rows['FORMA DE PAGO'].fillna('').replace('', 'TRANSFERENCIA BANCARIA').str.upper()
                    
                    new_rows['MONTO BASE USD'] = new_rows['MONTO ORIG'] / new_rows['TASA']
                    new_rows['HONORARIOS'] = new_rows['MONTO BASE USD'] * (new_rows['% ADMIN'] / 100)
                    new_rows['COSTO TOTAL'] = new_rows['MONTO BASE USD'] + new_rows['HONORARIOS']
                    
                    # MONTO PAGADO: Si está pagado, es el monto base
                    new_rows['MONTO PAGADO'] = new_rows.apply(lambda r: r['MONTO BASE USD'] if 'PAGADO' in r['ESTADO'] else 0, axis=1)
                    new_rows['SALDO PENDIENTE'] = new_rows['MONTO BASE USD'] - new_rows['MONTO PAGADO']
                    
                    # 3. Campos Automáticos
                    new_rows['CLASE'] = 'GASTO'
                    new_rows['MES'] = new_rows['FECHA'].dt.strftime('%Y-%m')
                    new_rows['SEMANA'] = new_rows['FECHA'].dt.strftime('%Y-S%V')
                    new_rows['EMPRESA'] = empresa_default
                    new_rows['OBRA'] = obra_default
                    new_rows['MONEDA'] = 'USD'
                    
                    # 4. Asegurar todas las columnas del CSV original en el orden correcto
                    csv_cols = list(df.columns)
                    for col in csv_cols:
                        if col not in new_rows.columns:
                            new_rows[col] = ""
                    
                    # Concatenar manteniendo el orden de columnas del original
                    updated_df = pd.concat([df, new_rows[csv_cols]], ignore_index=True)
                    updated_df.to_csv(active_csv, index=False)
                    
                    st.success(f"✅ ¡{len(new_rows)} gastos añadidos exitosamente!")
                    st.balloons()
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ Error al procesar: {e}")
            else:
                st.warning("⚠️ No hay datos para procesar.")

else:
    st.error("No se pudieron cargar los datos.")
