from pathlib import Path
import pandas as pd
import calendar
import streamlit as st
import matplotlib.pyplot as plt

st.set_page_config(page_title="Dashboard de Mantenimiento", layout="wide")
st.markdown(""""
    <style>
    .block-container {
        padding-top: 1rem;
        padding-bottom: 1rem;
    }
    </style>
    """, unsafe_allow_html=True)

st.title("Dashboard de Mantenimiento")
st.caption("MTTR / MTBF con resumen mensual y análisis por comedor")

BASE_DIR = Path(__file__).resolve().parent

ruta = BASE_DIR/ "mttr_mtbf.xlsx"

# -----------------------------
# LEER ARCHIVO
# -----------------------------
df = pd.read_excel(ruta, engine="openpyxl")
# -----------------------------
# LEER ARCHIVO OTs (Maximo)
# ------------------------------
ruta_carpeta_ot = BASE_DIR / "OTS_MAXIMO"

archivos_ot = list(ruta_carpeta_ot.glob("*.xlsx"))

lista_df_ot = []

for archivo in archivos_ot:

    df_temp = pd.read_excel(archivo, engine="openpyxl")
    df_temp.columns = df_temp.columns.str.strip().str.upper()
    lista_df_ot.append(df_temp)

df_ot = pd.concat(lista_df_ot, ignore_index=True)

#CONVERTIR FECHA
df_ot["SCHEDULED FINISH"] = pd.to_datetime(
    df_ot["SCHEDULED FINISH"],
    format="%m/%d/%y %I:%M %p",
    errors="coerce"
)

#SACAR MES
meses_num = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
    9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE"
}

df_ot["MES"] = df_ot["SCHEDULED FINISH"].dt.month.map(meses_num)

#CLASIFICAR AREA
def clasificar_area_pcon(texto):
    texto = str(texto).upper()

    if "IFSI" in texto:
        return "SCI"
    elif "IFCO" in texto:
        return "COMEDORES"
    elif "IFDE" in texto:
        return "DESASOLVE"
    elif "IFTE" in texto:
        return "TECHOS"
    elif "IFFA" in texto:
        return "CONSERVACION"
    else:
        return "VIAS"
    
df_ot["AREA"] = df_ot["PCON LOCATION"].apply(clasificar_area_pcon)

df.columns = df.columns.str.strip().str.upper()

#VALIDAR QUE EXISTA COMEDOR
if "COMEDOR" not in df.columns:
    st.error(f"Columnas disponibles: {df.columns.tolist()}")
    st.stop()

# Normalizar texto
for col in ["AREA", "DESCRIPCION", "TIPO_DE_FALLA", "RESPONSABLE", "TECNICO", "COMEDOR", "MES"]:
    if col in df.columns:
        df[col] = df[col].astype(str).str.strip().str.upper()

meses = {
    "ENERO": 1, "FEBRERO": 2, "MARZO": 3, "ABRIL": 4,
    "MAYO": 5, "JUNIO": 6, "JULIO": 7, "AGOSTO": 8,
    "SEPTIEMBRE": 9, "OCTUBRE": 10, "NOVIEMBRE": 11, "DICIEMBRE": 12
}

df["NUM_MES"] = df["MES"].map(meses)

# -----------------------------
# FILTROS
# -----------------------------
st.sidebar.header("Filtros")

meses_disponibles = sorted(df["MES"].dropna().unique(), key=lambda x: meses.get(x, 99))
areas_disponibles = sorted(df["AREA"].dropna().unique())
comedores_disponibles = sorted(df["COMEDOR"].dropna().unique())

meses_sel = st.sidebar.multiselect("Mes", meses_disponibles, default=meses_disponibles)
areas_sel = st.sidebar.multiselect("Área", areas_disponibles, default=areas_disponibles)
comedores_sel = st.sidebar.multiselect("Comedor", comedores_disponibles, default=comedores_disponibles)

df_filtrado = df[
    df["MES"].isin(meses_sel) &
    df["AREA"].isin(areas_sel) &
    df["COMEDOR"].isin(comedores_sel)
].copy()

df_ot_filtrado = df_ot[df_ot["MES"].isin(meses_sel)].copy()

if df_filtrado.empty:
    st.warning("No hay datos con los filtros seleccionados.")
    st.stop()

# -----------------------------
# RESUMEN MENSUAL
# -----------------------------
resumen = df_filtrado.groupby(["MES", "NUM_MES"]).agg(
    FALLAS=("MES", "count"),
    tiempo_total_min=("TIEMPO_REPARACION_(MIN)", "sum")
).reset_index()

resumen["MTTR_MIN"] = resumen["tiempo_total_min"] / resumen["FALLAS"]
resumen["MTTR_HR"] = resumen["MTTR_MIN"] / 60
resumen["DIAS_MES"] = resumen["NUM_MES"].apply(lambda x: calendar.monthrange(2026, int(x))[1])
resumen["HORAS_OPERACION"] = resumen["DIAS_MES"] * 24
resumen["MTBF_HR"] = resumen["HORAS_OPERACION"] / resumen["FALLAS"]
resumen["MTBF_MIN"] = resumen["MTBF_HR"] * 60

resumen = resumen.sort_values("NUM_MES")

for col in ["tiempo_total_min", "MTTR_MIN", "MTTR_HR", "MTBF_HR", "MTBF_MIN"]:
    resumen[col] = resumen[col].round(2)

reporte = resumen[[
    "MES", "FALLAS", "tiempo_total_min", "MTTR_MIN", "MTTR_HR", "MTBF_HR", "MTBF_MIN"
]]

reporte["MTTR_TENDENCIA"] = reporte["MTTR_MIN"].rolling(3).mean().round(2)
reporte["MTBF_TENDENCIA"] = reporte["MTBF_MIN"].rolling(3).mean().round(2)

# -----------------------------
# KPIS GENERALES
# -----------------------------
fallas_totales = int(len(df_filtrado))
tiempo_total = float(df_filtrado["TIEMPO_REPARACION_(MIN)"].sum())
mttr_general_min = round(tiempo_total / fallas_totales, 2)

horas_operacion_total = 0
for num_mes in resumen["NUM_MES"].unique():
    dias = calendar.monthrange(2026, int(num_mes))[1]
    horas_operacion_total += dias * 24

mtbf_general_min = round((horas_operacion_total / fallas_totales) * 60, 2)

st.subheader("Indicadores Generales")
k1, k2, k3, k4 = st.columns(4)
total_ots = int(len(df_ot_filtrado))
k1.metric("Fallas totales", fallas_totales)
k2.metric("MTTR promedio (min)", mttr_general_min)
k3.metric("MTBF promedio (min)", mtbf_general_min)
k4.metric("OTs totales", total_ots)

ots_area = df_ot_filtrado["AREA"].value_counts().reset_index()
ots_area.columns = ["AREA", "TOTAL_OT"]

#Semaforo MTTR
if mttr_general_min > 60:
    st.error("MTTR alto 🚨")
elif mttr_general_min > 50:
    st.warning("MTTR medio ⚠️")
else:
    st.success("MTTR controlado ✅")

#Semáforo MTBF
if mtbf_general_min < 900:
    st.error("MTBF bajo 🚨")
elif mtbf_general_min < 1050:
    st.warning("MTBF medio ⚠️")
else:
    st.success("MTBF controlado ✅")

#Insights automáticos
mes_peor_mttr = reporte.sort_values("MTTR_MIN", ascending=False).iloc[0]
mes_peor_mtbf = reporte.sort_values("MTBF_MIN", ascending=False).iloc[0]

st.info(
    f"El mes con mayor MTTR fue {mes_peor_mttr ['MES']} con {mes_peor_mttr ['MTTR_MIN']} min."
)

st.info(
    f"El mes con menor MTBF fue {mes_peor_mtbf ['MES']} con {mes_peor_mtbf ['MTBF_MIN']} min."
)

# -----------------------------
# RESUMEN MENSUAL
# -----------------------------
st.subheader("Resumen mensual")
st.dataframe(reporte, use_container_width=True)

# -----------------------------
# MTTR POR ÁREA
# -----------------------------
resumen_area = df_filtrado.groupby("AREA").agg(
    FALLAS=("AREA", "count"),
    tiempo_total_min=("TIEMPO_REPARACION_(MIN)", "sum")
).reset_index()

resumen_area["MTTR_MIN"] = resumen_area["tiempo_total_min"] / resumen_area["FALLAS"]
resumen_area["MTTR_MIN"] = resumen_area["MTTR_MIN"].round(2)
resumen_area = resumen_area.sort_values("FALLAS", ascending=False)

# -----------------------------
# MTBF POR COMEDOR
# -----------------------------
resumen_comedor = df_filtrado.groupby("COMEDOR").agg(
    FALLAS=("COMEDOR", "count"),
    tiempo_total_min=("TIEMPO_REPARACION_(MIN)", "sum")
).reset_index()

resumen_comedor["HORAS_OPERACION"] = horas_operacion_total
resumen_comedor["MTBF_MIN"] = ((resumen_comedor["HORAS_OPERACION"] / resumen_comedor["FALLAS"]) * 60).round(2)
resumen_comedor["MTTR_MIN"] = (resumen_comedor["tiempo_total_min"] / resumen_comedor["FALLAS"]).round(2)
resumen_comedor = resumen_comedor.sort_values("FALLAS", ascending=False)

c1, c2 = st.columns(2)

with c1:
    st.subheader("MTTR por área")
    st.dataframe(
        resumen_area[["AREA", "FALLAS", "tiempo_total_min", "MTTR_MIN"]],
        use_container_width=True
    )

with c2:
    st.subheader("Indicadores por comedor")
    st.dataframe(
        resumen_comedor[["COMEDOR", "FALLAS", "MTTR_MIN", "MTBF_MIN"]],
        use_container_width=True
    )

# -----------------------------
# TOP FALLAS
# -----------------------------
top_fallas = df_filtrado["TIPO_DE_FALLA"].value_counts().reset_index()
top_fallas.columns = ["TIPO_DE_FALLA", "TOTAL"]

st.subheader("Top fallas")
st.dataframe(top_fallas, use_container_width=True)

# -----------------------------
# GRAFICO DE DONA
# -----------------------------
st.subheader ("Ordenes de trabajo por area")

st. markdown("""
<style>
[data-testid="stDataFrame"] td {
    font-size: 20px !important;
    font-weight: 600 !important;
}
             
[data-testid="stDataFrame"] th {
    font-size: 22px !important;
}
<style>
""", unsafe_allow_html=True)
c5, c6, = st.columns([1,1])
with c5:

    total = ots_area["TOTAL_OT"].sum()

    fila_total = pd.DataFrame({
        "AREA": ["TOTAL"],
        "TOTAL_OT": [total]
    })

    ots_area_final = pd.concat([ots_area, fila_total], ignore_index=True)

    def resaltar_total(row):
        if row["AREA"] == "TOTAL":
            return ["background-color: #e6f0ff; font-weight: bold"] * len(row)
        else:
            return [""] * len(row)

    tabla_ots = (
        ots_area_final.style
        .apply(resaltar_total, axis=1)
        .set_table_styles([
            {
                "selector": "th",
                "props": [
                    ("background-color", "#003366"),
                    ("color", "white"),
                    ("font-size", "22px"),
                    ("font-weight", "bold"),
                    ("text-align", "left")
                ]
            },
            {
                "selector": "td",
                "props": [
                    ("font-size", "20px"),
                    ("font-weight", "bold")
                ]
            }
        ])
        .hide(axis="index")
    )

    st.markdown(tabla_ots.to_html(), unsafe_allow_html=True)
with c6:
    fig_dona, ax_dona = plt.subplots(figsize=(5,3.8))

    def autopct_func(pct):
        return f"{pct:.1f}%" if pct >= 5 else ""
    
    wedges, texts, autotexts = ax_dona.pie(
        ots_area["TOTAL_OT"],
        labels=None,
        autopct=autopct_func,
        startangle=90,
        wedgeprops={"width": 0.4},
        pctdistance=0.72
    )

    ax_dona.set_title("Distribución de OTs")

    ax_dona.legend(
        wedges,
        ots_area["AREA"],
        title="Área",
        loc="center left",
        bbox_to_anchor=(1.0, 0.5)
    )

    plt.tight_layout()
    st.pyplot(fig_dona, use_container_width=False)
# -----------------------------
# GRÁFICAS MÁS PEQUEÑAS
# -----------------------------
g1, g2 = st.columns(2)

with g1:
    st.subheader("MTTR por mes")
    fig1, ax1 = plt.subplots(figsize=(4, 2.5))
    ax1.plot(reporte["MES"], reporte["MTTR_MIN"], marker="o", label="MTTR")
    ax1.plot(reporte["MES"], reporte["MTTR_TENDENCIA"], linestyle="--", marker="s", label= "Tendencia 3M")
    ax1.set_ylabel("Min")
    ax1.grid(True)
    plt.tight_layout()
    st.pyplot(fig1, use_container_width=False)

with g2:
    st.subheader("MTBF por mes")
    fig2, ax2 = plt.subplots(figsize=(4, 2.5))
    ax2.plot(reporte["MES"], reporte["MTBF_MIN"], marker="o", label="MTBF")
    ax2.plot(reporte ["MES"], reporte["MTBF_TENDENCIA"], linestyle="--", marker="s", label="Tendencia 3M")
    ax2.set_ylabel("Min")
    ax2.grid(True)
    plt.tight_layout()
    st.pyplot(fig2, use_container_width=False)

g3, g4 = st.columns(2)

with g3:
    st.subheader("Fallas por área")
    fig3, ax3 = plt.subplots(figsize=(4, 2.5))
    ax3.bar(resumen_area["AREA"], resumen_area["FALLAS"])
    ax3.set_ylabel("Fallas")
    plt.tight_layout()
    st.pyplot(fig3, use_container_width=False)

with g4:
    st.subheader("Fallas por comedor")
    fig4, ax4 = plt.subplots(figsize=(4, 2.5))
    ax4.bar(resumen_comedor["COMEDOR"], resumen_comedor["FALLAS"])
    ax4.set_ylabel("Fallas")
    plt.xticks(rotation=45, ha="right")
    plt.tight_layout()
    st.pyplot(fig4, use_container_width=False)