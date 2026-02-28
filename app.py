import os
import io
import pandas as pd
import streamlit as st
from datetime import date
import plotly.express as px

# ============================
# CONFIG
# ============================
st.set_page_config(
    page_title="Med On Fit | Dashboard de Puntaje de Actitud",
    layout="wide",
    initial_sidebar_state="expanded",
)

EXCEL_PATH = "data.xlsx"          # Opción B: 1 sola hoja
LOGO_PATH  = "logo_medonfit.png"
SHEET_NAME = "Registro"

REQ_COLS = ["Fecha", "Alumno", "Tipo_Entrenamiento", "Puntaje"]


# ============================
# PREMIUM STYLES
# ============================
st.markdown("""
<style>
:root{
  --brand:#7a0f16;
  --card:#0f1724;
  --card2:#101a2b;
  --muted:#93a4b8;
  --text:#e8eef7;
  --line:rgba(255,255,255,.10);
}
.block-container{padding-top:1.0rem;padding-bottom:2.4rem;}
.card{
  background: linear-gradient(145deg, var(--card), var(--card2));
  border: 1px solid var(--line);
  border-radius: 18px;
  padding: 16px 16px;
  box-shadow: 0 10px 22px rgba(0,0,0,.22);
}
.badge{
  display:inline-block; padding:4px 10px; border-radius:999px;
  border:1px solid var(--line); color:var(--text);
  background:rgba(122,15,22,.18);
  font-size: 12px;
}
.small-muted{color:var(--muted);font-size:12px;}
.hr{height:1px;background:var(--line);margin:14px 0;}
/* métricas */
div[data-testid="stMetric"]{
  background: linear-gradient(145deg, var(--card), var(--card2));
  border: 1px solid var(--line);
  border-radius: 18px;
  padding: 14px 16px;
  box-shadow: 0 10px 22px rgba(0,0,0,.20);
}
</style>
""", unsafe_allow_html=True)

def fmt_int_es(x) -> str:
    try:
        return f"{int(x):,}".replace(",", ".")
    except Exception:
        return str(x)

def fmt_num_es(x) -> str:
    try:
        return f"{float(x):,.0f}".replace(",", ".")
    except Exception:
        return str(x)

def df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Export")
    return out.getvalue()


# ============================
# DATA LOAD / SAVE (Opción B)
# ============================
@st.cache_data(show_spinner=False)
def load_registro(excel_path: str) -> pd.DataFrame:
    if not os.path.exists(excel_path):
        base = pd.DataFrame(columns=REQ_COLS)
        base["__rowid__"] = []
        return base

    try:
        df = pd.read_excel(excel_path, sheet_name=0)
        if df is None or df.empty:
            base = pd.DataFrame(columns=REQ_COLS)
            base["__rowid__"] = []
            return base

        # Normaliza nombres comunes
        rename_map = {}
        for c in df.columns:
            cl = str(c).strip().lower()
            if cl in ["tipo de entrenamiento", "tipo_entrenamiento", "tipo entrenamiento"]:
                rename_map[c] = "Tipo_Entrenamiento"
        if rename_map:
            df = df.rename(columns=rename_map)

        for c in REQ_COLS:
            if c not in df.columns:
                df[c] = pd.NA

        df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce")
        df["Alumno"] = df["Alumno"].astype(str).str.strip()
        df["Tipo_Entrenamiento"] = df["Tipo_Entrenamiento"].astype(str).str.strip()
        df["Puntaje"] = pd.to_numeric(df["Puntaje"], errors="coerce")

        df = df.dropna(how="all")
        df = df[REQ_COLS].copy()

        # ID interno (NO se guarda en Excel; solo para selección/eliminación)
        df["__rowid__"] = range(1, len(df) + 1)
        return df
    except Exception:
        base = pd.DataFrame(columns=REQ_COLS)
        base["__rowid__"] = []
        return base

def add_time_fields(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    if out.empty:
        out["Mes"] = pd.Series(dtype="object")
        out["Semana"] = pd.Series(dtype="object")
        return out
    out["Mes"] = out["Fecha"].dt.to_period("M").astype(str)
    iso = out["Fecha"].dt.isocalendar()
    out["Semana"] = iso["year"].astype(str) + "-W" + iso["week"].astype(str).str.zfill(2)
    return out

def save_full_excel(excel_path: str, df: pd.DataFrame) -> None:
    df_save = df.copy()
    for c in REQ_COLS:
        if c not in df_save.columns:
            df_save[c] = pd.NA
    df_save = df_save[REQ_COLS].copy()

    with pd.ExcelWriter(excel_path, engine="openpyxl", mode="w") as writer:
        df_save.to_excel(writer, index=False, sheet_name=SHEET_NAME)


# ============================
# HEADER
# ============================
hL, hR = st.columns([1, 5], vertical_alignment="center")
with hL:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, use_container_width=True)
    else:
        st.markdown(
            '<div class="card"><span class="badge">Logo</span><div class="hr"></div>'
            '<div class="small-muted">Guarda el logo como <b>logo_medonfit.png</b> en la carpeta del proyecto.</div></div>',
            unsafe_allow_html=True
        )

with hR:
    st.markdown("""
    <div class="card">
      <div style="display:flex;justify-content:space-between;align-items:center;gap:16px;">
        <div>
          <div style="font-size:28px;font-weight:900;color:#e8eef7;line-height:1.2;text-transform:uppercase;">
            DASHBOARD DE PUNTAJE DE ACTITUD MED ON FIT
          </div>
          <div class="small-muted">
            Top 10 • Top 1 • Distribución por tipo • Matriz por periodo • Gestión de registros (alta/baja)
          </div>
        </div>
        <div style="text-align:right">
          <span class="badge">Opción B (1 hoja)</span>
          &nbsp;
          <span class="badge">Excel como fuente</span>
        </div>
      </div>
    </div>
    """, unsafe_allow_html=True)

st.write("")


# ============================
# LOAD DATA
# ============================
df_base = load_registro(EXCEL_PATH)
df_base = add_time_fields(df_base)

# ============================
# SIDEBAR (Filtros + Ingreso)
# ============================
with st.sidebar:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, use_container_width=True)

    st.markdown("### Filtros avanzados")

    if df_base.empty or df_base["Fecha"].dropna().empty:
        min_date = date.today()
        max_date = date.today()
    else:
        min_date = df_base["Fecha"].min().date()
        max_date = df_base["Fecha"].max().date()

    date_range = st.date_input("Rango de fechas", value=(min_date, max_date))
    if isinstance(date_range, tuple) and len(date_range) == 2:
        d0, d1 = date_range
    else:
        d0, d1 = min_date, max_date

    granularidad = st.radio("Agrupar por", ["Mes", "Semana"], horizontal=True)
    period_col = "Mes" if granularidad == "Mes" else "Semana"

    tipos = sorted([t for t in df_base["Tipo_Entrenamiento"].dropna().unique().tolist() if t and t != "nan"])
    tipo_sel = st.multiselect("Tipo de entrenamiento", options=tipos, default=[])

    alumnos = sorted([a for a in df_base["Alumno"].dropna().unique().tolist() if a and a != "nan"])
    alumno_sel = st.multiselect("Alumno", options=alumnos, default=[])

    st.markdown("---")
    st.markdown("### Ingreso de registro")

    fecha_new = st.date_input("Fecha (nuevo)", value=date.today(), key="fecha_new")
    alumno_pick = st.selectbox("Alumno", (alumnos + ["(Nuevo alumno)"]) if alumnos else ["(Nuevo alumno)"])
    alumno_new = st.text_input("Nombre alumno (si es nuevo)", value="", placeholder="Ej: Juan Pérez").strip() if alumno_pick == "(Nuevo alumno)" else alumno_pick

    tipo_pick = st.selectbox("Tipo", (tipos + ["(Nuevo tipo)"]) if tipos else ["(Nuevo tipo)"])
    tipo_new = st.text_input("Nombre tipo (si es nuevo)", value="", placeholder="Ej: HIIT").strip() if tipo_pick == "(Nuevo tipo)" else tipo_pick

    puntaje_new = st.number_input("Puntaje", min_value=0.0, step=1.0)

    cA, cB = st.columns(2)
    with cA:
        if st.button("💾 Guardar", use_container_width=True):
            if not alumno_new:
                st.error("Ingresa un nombre de alumno válido.")
            elif not tipo_new:
                st.error("Ingresa un tipo válido.")
            else:
                new_row = pd.DataFrame([{
                    "Fecha": pd.to_datetime(fecha_new),
                    "Alumno": alumno_new,
                    "Tipo_Entrenamiento": tipo_new,
                    "Puntaje": float(puntaje_new),
                }])

                base_for_save = df_base.copy()
                base_for_save = base_for_save.drop(columns=["__rowid__", "Mes", "Semana"], errors="ignore")

                df_save = pd.concat([base_for_save[REQ_COLS], new_row], ignore_index=True)
                save_full_excel(EXCEL_PATH, df_save)
                st.success("Guardado en data.xlsx ✅")
                st.cache_data.clear()
                st.rerun()

    with cB:
        st.download_button(
            "⬇️ Plantilla",
            data=df_to_excel_bytes(pd.DataFrame(columns=REQ_COLS)),
            file_name="plantilla_registro_medonfit.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )


# ============================
# APPLY FILTERS
# ============================
df = df_base.copy()
if not df.empty:
    df = df[(df["Fecha"].dt.date >= d0) & (df["Fecha"].dt.date <= d1)]
if tipo_sel:
    df = df[df["Tipo_Entrenamiento"].isin(tipo_sel)]
if alumno_sel:
    df = df[df["Alumno"].isin(alumno_sel)]

df = add_time_fields(df)

# ============================
# KPIs (SIN "último mes vs anterior")
# ============================
total_puntos = float(df["Puntaje"].fillna(0).sum()) if not df.empty else 0.0
total_registros = int(len(df)) if not df.empty else 0
n_alumnos = int(df["Alumno"].nunique()) if not df.empty else 0

agg = (
    df.groupby("Alumno", as_index=False)["Puntaje"]
      .sum()
      .sort_values("Puntaje", ascending=False)
) if not df.empty else pd.DataFrame(columns=["Alumno", "Puntaje"])

top1_name = agg.iloc[0]["Alumno"] if not agg.empty else "-"
top1_score = float(agg.iloc[0]["Puntaje"]) if not agg.empty else 0.0

k1, k2, k3, k4 = st.columns(4)
k1.metric("Puntaje total", fmt_num_es(total_puntos))
k2.metric("Registros", fmt_int_es(total_registros))
k3.metric("Alumnos únicos", fmt_int_es(n_alumnos))
k4.metric("🏆 Top 1", top1_name, fmt_num_es(top1_score))

st.write("")


# ============================
# TABS
# ============================
tab1, tab2 = st.tabs(["📊 Dashboard", "🧾 Gestión de registros (Eliminar)"])


# ----------------------------
# TAB 1: DASHBOARD
# (Se elimina la sección "Vista por mes sin tendencia" y se deja solo Donut + Matriz en bloque limpio)
# ----------------------------
with tab1:
    c1, c2 = st.columns([1.15, 1.85], vertical_alignment="top")

    with c1:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown("#### Ranking Top 10 (acumulado)")

        if agg.empty:
            st.info("No hay datos con los filtros actuales.")
        else:
            top10 = agg.head(10).copy().reset_index(drop=True)
            medals = ["🥇", "🥈", "🥉"] + ["🏅"] * 7
            top10.insert(0, "Rank", [f"{medals[i]} {i+1}" for i in range(len(top10))])
            top10_show = top10.rename(columns={"Puntaje": "Puntaje acumulado"})

            st.dataframe(top10_show, use_container_width=True, hide_index=True)

            fig_bar = px.bar(
                top10_show.sort_values("Puntaje acumulado"),
                x="Puntaje acumulado",
                y="Alumno",
                orientation="h",
            )
            fig_bar.update_layout(
                height=360,
                margin=dict(l=10, r=10, t=10, b=10),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                font=dict(color="#e8eef7"),
            )
            st.plotly_chart(fig_bar, use_container_width=True)

        st.markdown('</div>', unsafe_allow_html=True)

    with c2:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown("#### Distribución por tipo + Matriz por periodo")

        if df.empty:
            st.info("No hay datos para graficar.")
        else:
            # --- Donut por tipo ---
            st.markdown("#### Distribución por tipo (donut)")
            by_type = (
                df.groupby("Tipo_Entrenamiento", as_index=False)["Puntaje"]
                  .sum()
                  .sort_values("Puntaje", ascending=False)
            )
            fig_pie = px.pie(by_type, values="Puntaje", names="Tipo_Entrenamiento", hole=0.55)
            fig_pie.update_layout(
                height=320,
                margin=dict(l=10, r=10, t=10, b=10),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor="rgba(0,0,0,0)",
                font=dict(color="#e8eef7"),
                legend_title_text="Tipo",
            )
            st.plotly_chart(fig_pie, use_container_width=True)

            st.markdown('<div class="hr"></div>', unsafe_allow_html=True)

            # --- Matriz alumno x periodo ---
            st.markdown(f"#### Matriz: acumulado por alumno ({granularidad})")
            pivot = df.pivot_table(
                index="Alumno",
                columns=period_col,
                values="Puntaje",
                aggfunc="sum",
                fill_value=0,
            )
            pivot["Total"] = pivot.sum(axis=1)
            pivot = pivot.sort_values("Total", ascending=False)
            st.dataframe(pivot, use_container_width=True)

        st.markdown('</div>', unsafe_allow_html=True)

    st.write("")
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("#### Detalle y exportación")

    if df.empty:
        st.info("Sin registros para mostrar.")
    else:
        detalle_cols = ["__rowid__", "Fecha", "Mes", "Semana", "Alumno", "Tipo_Entrenamiento", "Puntaje"]
        st.dataframe(
            df[detalle_cols].sort_values("Fecha", ascending=False),
            use_container_width=True,
            hide_index=True
        )

        colA, colB = st.columns(2)
        with colA:
            st.download_button(
                "⬇️ Descargar datos filtrados (Excel)",
                data=df_to_excel_bytes(df[["Fecha", "Alumno", "Tipo_Entrenamiento", "Puntaje"]].copy()),
                file_name="medonfit_datos_filtrados.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with colB:
            st.download_button(
                "⬇️ Descargar ranking Top 10 (Excel)",
                data=df_to_excel_bytes(agg.head(10).copy()),
                file_name="medonfit_ranking_top10.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    st.markdown('</div>', unsafe_allow_html=True)


# ----------------------------
# TAB 2: ELIMINAR REGISTROS
# ----------------------------
with tab2:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("### Eliminación de registros (por error de ingreso)")
    st.markdown(
        '<div class="small-muted">Selecciona las filas a eliminar y confirma. Esto actualizará <b>data.xlsx</b>.</div>',
        unsafe_allow_html=True
    )
    st.markdown('<div class="hr"></div>', unsafe_allow_html=True)

    if df_base.empty:
        st.info("No hay registros en el Excel para eliminar.")
    else:
        view_cols = ["__rowid__", "Fecha", "Alumno", "Tipo_Entrenamiento", "Puntaje"]
        to_manage = df_base[view_cols].copy()
        to_manage.insert(0, "Eliminar", False)

        edited = st.data_editor(
            to_manage,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Eliminar": st.column_config.CheckboxColumn("Eliminar", help="Marca filas para borrar"),
                "__rowid__": st.column_config.NumberColumn("ID", help="Identificador interno"),
                "Fecha": st.column_config.DatetimeColumn("Fecha"),
            },
            disabled=["__rowid__", "Fecha", "Alumno", "Tipo_Entrenamiento", "Puntaje"],
            height=420,
        )

        selected_ids = edited.loc[edited["Eliminar"] == True, "__rowid__"].tolist()

        st.write("")
        c1, c2, c3 = st.columns([1.2, 1, 1])

        with c1:
            st.markdown(f"<span class='badge'>Seleccionados: {len(selected_ids)}</span>", unsafe_allow_html=True)

        with c2:
            confirm = st.checkbox("Confirmo que deseo eliminar", value=False)

        with c3:
            if st.button("🗑️ Eliminar seleccionados", use_container_width=True,
                         disabled=(not confirm or len(selected_ids) == 0)):
                raw = df_base.copy()
                raw = raw.drop(columns=["Mes", "Semana"], errors="ignore")

                remaining = raw[~raw["__rowid__"].isin(selected_ids)].copy()
                remaining = remaining.drop(columns=["__rowid__"], errors="ignore")

                save_full_excel(EXCEL_PATH, remaining[REQ_COLS].copy())

                st.success("Registros eliminados y Excel actualizado ✅")
                st.cache_data.clear()
                st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)