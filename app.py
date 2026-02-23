import streamlit as st
import pandas as pd
import numpy as np
import random
import math
import io
from copy import deepcopy
from collections import Counter
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ============================================================
# CONFIGURACIÓN DE PÁGINA
# ============================================================
st.set_page_config(
    page_title="Formador de Familias",
    page_icon="👥",
    layout="wide",
)

st.title("👥 Formador de Grupos")
st.caption("Ingresa los participantes, configura las restricciones y genera los grupos automáticamente.")

# ============================================================
# FUNCIONES DE OPTIMIZACIÓN
# ============================================================
def calcular_score(grupos, edades, es_hombre, carreras, pares_idx, piso_h, techo_h):
    score = 0
    varianzas = []
    for g_indices in grupos:
        edades_g  = edades[g_indices]
        hombres_g = es_hombre[g_indices].sum()
        carreras_g = carreras[g_indices]
        idx_set   = set(g_indices)

        for idx1, idx2 in pares_idx:
            if idx1 in idx_set and idx2 in idx_set:
                score += 1_000_000

        for conteo in Counter(carreras_g).values():
            if conteo > 1:
                score += (conteo - 1) * 5_000

        if hombres_g < piso_h or hombres_g > techo_h:
            score += 10_000

        v = edades_g.var()
        varianzas.append(v)
        score += v * 150

    score += max(varianzas) * 600
    score += (max(varianzas) - min(varianzas)) * 300
    return score


def inicializar_grupos(df, lideres, n_grupos):
    idx_lideres = df[df["Nombre"].isin(lideres)].index.tolist()
    idx_resto   = df[~df["Nombre"].isin(lideres)].index.tolist()
    random.shuffle(idx_resto)
    grupos = [[] for _ in range(n_grupos)]
    for i, idx in enumerate(idx_lideres):
        grupos[i].append(idx)
    ptr = 0
    for g_idx in range(n_grupos):
        while len(grupos[g_idx]) < len(df) // n_grupos:
            grupos[g_idx].append(idx_resto[ptr])
            ptr += 1
    # Distribuir sobrantes
    while ptr < len(idx_resto):
        for g_idx in range(n_grupos):
            if ptr < len(idx_resto):
                grupos[g_idx].append(idx_resto[ptr])
                ptr += 1
    return [list(g) for g in grupos]


def optimizar(df, lideres, pares_prohibidos, n_grupos, n_iter=80_000, T_inicial=500.0, T_final=0.1):
    edades    = df["Edad"].values.astype(float)
    es_hombre = (df["Sexo"] == "Hombre").values
    carreras  = df["Carrera"].values

    n_hombres_total = es_hombre.sum()
    piso_h  = math.floor(n_hombres_total / n_grupos)
    techo_h = math.ceil(n_hombres_total / n_grupos)

    pares_idx = []
    for p1, p2 in pares_prohibidos:
        r1 = df[df["Nombre"] == p1]
        r2 = df[df["Nombre"] == p2]
        if not r1.empty and not r2.empty:
            pares_idx.append((r1.index[0], r2.index[0]))

    grupos = inicializar_grupos(df, lideres, n_grupos)
    score_actual = calcular_score(grupos, edades, es_hombre, carreras, pares_idx, piso_h, techo_h)
    mejor_grupos = deepcopy(grupos)
    mejor_score  = score_actual

    alpha = (T_final / T_inicial) ** (1.0 / n_iter)
    T = T_inicial
    n_lideres = len(lideres)
    tam_grupo = len(grupos[0]) - 1  # índice máximo dentro del grupo

    for _ in range(n_iter):
        g1, g2 = random.sample(range(n_grupos), 2)
        start1 = 1 if g1 < n_lideres else 0
        start2 = 1 if g2 < n_lideres else 0
        p1 = random.randint(start1, len(grupos[g1]) - 1)
        p2 = random.randint(start2, len(grupos[g2]) - 1)

        grupos[g1][p1], grupos[g2][p2] = grupos[g2][p2], grupos[g1][p1]
        nuevo_score = calcular_score(grupos, edades, es_hombre, carreras, pares_idx, piso_h, techo_h)
        delta = nuevo_score - score_actual

        if delta < 0 or random.random() < math.exp(-delta / T):
            score_actual = nuevo_score
            if score_actual < mejor_score:
                mejor_score  = score_actual
                mejor_grupos = deepcopy(grupos)
        else:
            grupos[g1][p1], grupos[g2][p2] = grupos[g2][p2], grupos[g1][p1]
        T *= alpha

    return mejor_grupos, mejor_score


def solucion_a_frozenset(grupos):
    return frozenset(frozenset(g) for g in grupos)


def correr_optimizacion(df, lideres, pares_prohibidos, n_grupos, n_corridas=10):
    mejores = []
    huellas = set()
    mejor_score_global = float("inf")

    bar = st.progress(0, text="Iniciando optimización...")
    for semilla in range(n_corridas):
        bar.progress((semilla + 1) / n_corridas, text=f"Corrida {semilla + 1} de {n_corridas}...")
        random.seed(semilla)
        grupos, score = optimizar(df, lideres, pares_prohibidos, n_grupos)
        huella = solucion_a_frozenset(grupos)

        if score < mejor_score_global - 0.01:
            mejor_score_global = score
            mejores = [grupos]
            huellas = {huella}
        elif abs(score - mejor_score_global) < 0.01 and huella not in huellas:
            mejores.append(grupos)
            huellas.add(huella)

    bar.empty()
    return mejores, mejor_score_global


# ============================================================
# FUNCIÓN: GENERAR EXCEL DE RESULTADOS
# ============================================================
def generar_excel_resultados(df, mejores_resultados):
    wb = Workbook()
    wb.remove(wb.active)  # Eliminar hoja vacía inicial

    AZUL     = "1F4E79"
    AZUL_CLA = "D6E4F0"
    COLORES_GRUPO = ["E8F5E9","FFF3E0","FCE4EC","E3F2FD","F3E5F5","E0F7FA"]

    borde = Border(
        left=Side(style="thin", color="BFBFBF"),
        right=Side(style="thin", color="BFBFBF"),
        top=Side(style="thin", color="BFBFBF"),
        bottom=Side(style="thin", color="BFBFBF"),
    )

    for n_op, grupos in enumerate(mejores_resultados):
        ws = wb.create_sheet(title=f"Opción {n_op + 1}")
        fila = 1

        for i, g_indices in enumerate(grupos):
            g = df.loc[g_indices]
            h = (g["Sexo"] == "Hombre").sum()
            m = (g["Sexo"] == "Mujer").sum()
            prom = g["Edad"].mean()
            vari = g["Edad"].var()
            color_g = COLORES_GRUPO[i % len(COLORES_GRUPO)]

            # Título del grupo
            ws.merge_cells(start_row=fila, start_column=1, end_row=fila, end_column=4)
            titulo = ws.cell(row=fila, column=1,
                             value=f"  FAMILIA {i+1}   |   {h}H / {m}M   |   Edad prom: {prom:.1f}   |   Varianza: {vari:.2f}")
            titulo.font      = Font(bold=True, name="Arial", size=11, color="FFFFFF")
            titulo.fill      = PatternFill("solid", start_color=AZUL)
            titulo.alignment = Alignment(horizontal="left", vertical="center")
            ws.row_dimensions[fila].height = 20
            fila += 1

            # Cabeceras de columnas
            for col, header in enumerate(["Nombre", "Sexo", "Edad", "Carrera"], 1):
                cell = ws.cell(row=fila, column=col, value=header)
                cell.font      = Font(bold=True, name="Arial", size=10, color=AZUL)
                cell.fill      = PatternFill("solid", start_color=AZUL_CLA)
                cell.alignment = Alignment(horizontal="center")
                cell.border    = borde
            fila += 1

            # Datos de cada persona
            for _, persona in g.iterrows():
                for col, val in enumerate([persona["Nombre"], persona["Sexo"], persona["Edad"], persona["Carrera"]], 1):
                    cell = ws.cell(row=fila, column=col, value=val)
                    cell.font      = Font(name="Arial", size=10)
                    cell.fill      = PatternFill("solid", start_color=color_g)
                    cell.alignment = Alignment(horizontal="left" if col != 3 else "center")
                    cell.border    = borde
                fila += 1

            fila += 1  # Espacio entre grupos

        # Anchos de columna
        ws.column_dimensions["A"].width = 42
        ws.column_dimensions["B"].width = 10
        ws.column_dimensions["C"].width = 8
        ws.column_dimensions["D"].width = 32

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


# ============================================================
# PASO 1: INGRESAR PARTICIPANTES
# ============================================================
st.header("1️⃣ Participantes")

modo = st.radio("¿Cómo deseas ingresar los datos?",
                ["📂 Subir archivo Excel", "✏️ Captura manual"],
                horizontal=True)

df_participantes = None

if modo == "📂 Subir archivo Excel":
    # Botón para descargar plantilla
    with open("plantilla_participantes.xlsx", "rb") as f:
        st.download_button(
            label="⬇️ Descargar plantilla Excel",
            data=f,
            file_name="plantilla_participantes.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Descarga esta plantilla, llénala con los datos y súbela aquí.",
        )

    archivo = st.file_uploader("Sube el archivo Excel con los participantes", type=["xlsx"])
    if archivo:
        try:
            df_leido = pd.read_excel(archivo)
            # Validar columnas
            cols_requeridas = {"Nombre", "Sexo", "Edad", "Carrera"}
            if not cols_requeridas.issubset(df_leido.columns):
                st.error(f"El archivo debe tener exactamente estas columnas: {cols_requeridas}")
            else:
                # Validar valores de Sexo
                sexos_invalidos = df_leido[~df_leido["Sexo"].isin(["Hombre", "Mujer"])]
                if not sexos_invalidos.empty:
                    st.error(f"La columna 'Sexo' solo acepta 'Hombre' o 'Mujer'. Revisa estas filas: {sexos_invalidos.index.tolist()}")
                elif df_leido["Edad"].isnull().any() or not pd.api.types.is_numeric_dtype(df_leido["Edad"]):
                    st.error("La columna 'Edad' debe contener solo números enteros.")
                elif df_leido["Nombre"].duplicated().any():
                    st.error("Hay nombres duplicados en el archivo. Cada participante debe tener un nombre único.")
                else:
                    df_participantes = df_leido[["Nombre", "Sexo", "Edad", "Carrera"]].copy()
                    df_participantes["Edad"] = df_participantes["Edad"].astype(int)
                    st.success(f"✅ {len(df_participantes)} participantes cargados correctamente.")
                    st.dataframe(df_participantes, use_container_width=True, hide_index=True)
        except Exception as e:
            st.error(f"Error al leer el archivo: {e}")

else:  # Captura manual
    st.info("Ingresa los participantes uno por uno. Presiona **Agregar** después de cada uno.")

    if "participantes" not in st.session_state:
        st.session_state.participantes = []

    with st.form("form_agregar", clear_on_submit=True):
        col1, col2, col3, col4 = st.columns([3, 1, 1, 2])
        nombre  = col1.text_input("Nombre completo")
        sexo    = col2.selectbox("Sexo", ["Mujer", "Hombre"])
        edad    = col3.number_input("Edad", min_value=15, max_value=80, value=20, step=1)
        carrera = col4.text_input("Carrera")
        agregar = st.form_submit_button("➕ Agregar participante")

        if agregar:
            if not nombre.strip() or not carrera.strip():
                st.warning("El nombre y la carrera no pueden estar vacíos.")
            elif any(p["Nombre"] == nombre.strip() for p in st.session_state.participantes):
                st.warning("Ya existe un participante con ese nombre.")
            else:
                st.session_state.participantes.append({
                    "Nombre": nombre.strip(),
                    "Sexo": sexo,
                    "Edad": int(edad),
                    "Carrera": carrera.strip(),
                })

    if st.session_state.participantes:
        df_manual = pd.DataFrame(st.session_state.participantes)
        st.dataframe(df_manual, use_container_width=True, hide_index=True)
        st.caption(f"{len(df_manual)} participantes agregados.")

        col_borrar, _ = st.columns([1, 3])
        if col_borrar.button("🗑️ Borrar todos"):
            st.session_state.participantes = []
            st.rerun()

        df_participantes = df_manual.copy()

# ============================================================
# PASO 2: CONFIGURACIÓN
# ============================================================
if df_participantes is not None and len(df_participantes) > 0:
    st.divider()
    st.header("2️⃣ Configuración de grupos")

    nombres_disponibles = sorted(df_participantes["Nombre"].tolist())

    col_cfg1, col_cfg2 = st.columns(2)

    with col_cfg1:
        n_grupos = st.number_input(
            "Número de grupos",
            min_value=2,
            max_value=len(df_participantes) // 2,
            value=min(6, len(df_participantes) // 5),
            step=1,
        )

        st.subheader("🏅 Líderes (uno por grupo)")
        st.caption("Selecciona las personas que deben quedar en grupos distintos. Máximo uno por grupo.")
        lideres = st.multiselect(
            "Líderes",
            options=nombres_disponibles,
            max_selections=n_grupos,
            label_visibility="collapsed",
        )

    with col_cfg2:
        st.subheader("🚫 Pares prohibidos")
        st.caption("Personas que NO pueden estar en el mismo grupo.")

        if "pares" not in st.session_state:
            st.session_state.pares = []

        with st.form("form_par", clear_on_submit=True):
            pc1, pc2 = st.columns(2)
            par1 = pc1.selectbox("Persona 1", nombres_disponibles, key="par1")
            par2 = pc2.selectbox("Persona 2", nombres_disponibles, key="par2")
            agregar_par = st.form_submit_button("➕ Agregar par prohibido")

            if agregar_par:
                if par1 == par2:
                    st.warning("Selecciona dos personas distintas.")
                elif (par1, par2) in st.session_state.pares or (par2, par1) in st.session_state.pares:
                    st.warning("Ese par ya está registrado.")
                else:
                    st.session_state.pares.append((par1, par2))

        if st.session_state.pares:
            for idx, (p1, p2) in enumerate(st.session_state.pares):
                col_p, col_x = st.columns([4, 1])
                col_p.write(f"❌ {p1}  ↔  {p2}")
                if col_x.button("Quitar", key=f"del_par_{idx}"):
                    st.session_state.pares.pop(idx)
                    st.rerun()

    # ============================================================
    # PASO 3: GENERAR GRUPOS
    # ============================================================
    st.divider()
    st.header("3️⃣ Generar grupos")

    # Advertencias
    if len(lideres) > n_grupos:
        st.error(f"Hay más líderes ({len(lideres)}) que grupos ({n_grupos}). Reduce los líderes.")
    elif len(df_participantes) % n_grupos != 0:
        st.warning(
            f"⚠️ {len(df_participantes)} participantes no se dividen exactamente en {n_grupos} grupos. "
            f"Algunos grupos tendrán {len(df_participantes) // n_grupos} personas y otros "
            f"{len(df_participantes) // n_grupos + 1}."
        )

    if st.button("🚀 Generar grupos", type="primary", use_container_width=True):
        if len(df_participantes) < n_grupos:
            st.error("Hay menos participantes que grupos.")
        else:
            with st.spinner("Optimizando configuración de grupos..."):
                mejores, score_global = correr_optimizacion(
                    df_participantes,
                    lideres,
                    st.session_state.pares,
                    int(n_grupos),
                )
            st.session_state.mejores   = mejores
            st.session_state.score     = score_global
            st.session_state.df_result = df_participantes

    # ============================================================
    # PASO 4: MOSTRAR RESULTADOS
    # ============================================================
    if "mejores" in st.session_state and st.session_state.mejores:
        st.divider()
        st.header("4️⃣ Resultados")

        mejores = st.session_state.mejores
        df_res  = st.session_state.df_result

        if len(mejores) > 1:
            st.info(f"Se encontraron **{len(mejores)} configuraciones distintas** con la misma calidad. Elige la que prefieras.")
            tabs_opciones = st.tabs([f"Opción {i+1}" for i in range(len(mejores))])
        else:
            tabs_opciones = [st.container()]

        for tab, grupos in zip(tabs_opciones, mejores):
            with tab:
                cols = st.columns(min(3, int(n_grupos)))
                for i, g_indices in enumerate(grupos):
                    g = df_res.loc[g_indices]
                    h = (g["Sexo"] == "Hombre").sum()
                    m = (g["Sexo"] == "Mujer").sum()
                    prom = g["Edad"].mean()
                    vari = g["Edad"].var()
                    unicas = g["Carrera"].nunique()

                    with cols[i % len(cols)]:
                        st.markdown(f"### Familia {i+1}")
                        st.caption(
                            f"👥 {h}H / {m}M &nbsp;|&nbsp; 🎂 Prom. {prom:.1f} &nbsp;|&nbsp; "
                            f"📊 Var. {vari:.2f} &nbsp;|&nbsp; 🎓 {unicas}/{len(g)} carreras únicas"
                        )
                        st.dataframe(
                            g[["Nombre", "Sexo", "Edad", "Carrera"]].reset_index(drop=True),
                            use_container_width=True,
                            hide_index=True,
                        )

        # Descarga Excel
        st.divider()
        excel_buffer = generar_excel_resultados(df_res, mejores)
        st.download_button(
            label="⬇️ Descargar resultados en Excel",
            data=excel_buffer,
            file_name="grupos_resultado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )