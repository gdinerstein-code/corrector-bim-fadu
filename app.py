"""
Corrector BIM FADU — App Streamlit
Flujo: subís planilla + archivos del alumno → IA evalúa → revisás → descargás planilla actualizada
"""

import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
from evaluador import (
    EIR_CRITERIA, BEP_CRITERIA,
    extract_text, clasificar_archivos,
    evaluar_documento, get_student_list, escribir_resultados,
)

st.set_page_config(page_title="Corrector BIM FADU", page_icon="📐", layout="wide")

# ── CSS mínimo ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
.step-header { font-size: 1.1rem; font-weight: 600; margin-bottom: 0.25rem; }
</style>
""", unsafe_allow_html=True)


# ── Sidebar ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.title("📐 Corrector BIM FADU")
    st.caption("AS10 · Procesos Metodológicos")
    st.divider()

    # La API key se lee desde Streamlit Secrets (cuando está hosteada)
    # o desde un campo manual (cuando se corre local)
    api_key = st.secrets.get("GEMINI_API_KEY", "") if hasattr(st, "secrets") else ""
    if not api_key:
        st.markdown("**API Key de Gemini (Google)**")
        st.caption("Gratis en [aistudio.google.com](https://aistudio.google.com) → Get API Key")
        api_key = st.text_input("API Key", type="password", label_visibility="collapsed",
                                placeholder="AIza...")
    else:
        st.success("✓ API Key configurada")

    st.divider()
    st.markdown("**¿Primera vez?**")
    with st.expander("Ver instrucciones"):
        st.markdown("""
1. Obtené tu API key gratuita en [aistudio.google.com](https://aistudio.google.com)
2. Pegala en el campo de arriba
3. Subí la planilla Excel de la cohorte
4. Seleccioná el alumno
5. Subí todos sus archivos (Ctrl+A en la carpeta)
6. Evaluá y descargá la planilla actualizada
7. La próxima vez subí la planilla descargada (ya tiene las correcciones anteriores)
""")


# ── Estado de sesión ───────────────────────────────────────────────────────────
if "planilla_bytes" not in st.session_state:
    st.session_state.planilla_bytes   = None
if "eir_results"    not in st.session_state:
    st.session_state.eir_results      = None
if "bep_results"    not in st.session_state:
    st.session_state.bep_results      = None
if "alumno_apellido" not in st.session_state:
    st.session_state.alumno_apellido  = None


# ── PASO 1: Planilla ───────────────────────────────────────────────────────────
st.markdown('<p class="step-header">① Subí la planilla Excel de la cohorte</p>', unsafe_allow_html=True)

planilla_file = st.file_uploader(
    "Planilla",
    type=["xlsx"],
    label_visibility="collapsed",
    key="planilla_upload",
    help="COHORTE 12_ LISTADO ACTIVOS.xlsx — o la versión con correcciones anteriores",
)

if planilla_file:
    st.session_state.planilla_bytes = planilla_file.read()
    st.success(f"✓ {planilla_file.name}")

planilla_ok = st.session_state.planilla_bytes is not None

st.divider()


# ── PASO 2: Alumno ─────────────────────────────────────────────────────────────
st.markdown('<p class="step-header">② Seleccioná el alumno</p>', unsafe_allow_html=True)

if not planilla_ok:
    st.info("Primero subí la planilla Excel.")
    alumno = None
else:
    @st.cache_data
    def cargar_alumnos(planilla_hash: int):
        return get_student_list(st.session_state.planilla_bytes)

    students = cargar_alumnos(hash(bytes(st.session_state.planilla_bytes)))
    opciones  = [s["display"] for s in students]
    seleccion = st.selectbox("Alumno", opciones, label_visibility="collapsed")
    alumno    = next((s for s in students if s["display"] == seleccion), None)

    if alumno:
        eir_ok = alumno["col_eir"][0] is not None
        bep_ok = alumno["col_bep"][0] is not None
        if eir_ok and bep_ok:
            st.caption(f"✓ Encontrado en la planilla — EIR col {alumno['col_eir'][0]}, BEP col {alumno['col_bep'][0]}")
        else:
            st.warning("⚠️ Este alumno no tiene columnas en la planilla.")

st.divider()


# ── PASO 3: Archivos del alumno ────────────────────────────────────────────────
st.markdown('<p class="step-header">③ Subí los archivos del alumno</p>', unsafe_allow_html=True)
st.caption("Seleccioná la carpeta del alumno o arrastrá los archivos acá.")

doc_files = st.file_uploader(
    "Archivos",
    type=["pdf", "docx", "doc", "xlsx", "xlsm", "xls"],
    accept_multiple_files=True,
    label_visibility="collapsed",
    key="docs_upload",
)

# Inyectar JS para que el input acepte selección de carpeta
components.html("""
<script>
(function() {
    const parentDoc = window.parent.document;
    function patchInput() {
        const uploaders = parentDoc.querySelectorAll(
            'section[data-testid="stFileUploaderDropzone"] input[type="file"]'
        );
        uploaders.forEach(function(inp, i) {
            // Solo el segundo uploader (índice 1) es el de documentos
            if (i === 1) {
                inp.setAttribute('webkitdirectory', '');
                inp.setAttribute('multiple', '');
            }
        });
    }
    patchInput();
    const obs = new MutationObserver(patchInput);
    obs.observe(parentDoc.body, { childList: true, subtree: true });
})();
</script>
""", height=0)

if doc_files:
    eir_files, bep_files = clasificar_archivos(doc_files)
    col1, col2 = st.columns(2)
    with col1:
        st.caption(f"📋 EIR ({len(eir_files)} archivos): " + ", ".join(f.name for f in eir_files[:3]) + ("..." if len(eir_files) > 3 else ""))
    with col2:
        st.caption(f"📋 BEP ({len(bep_files)} archivos): " + ", ".join(f.name for f in bep_files[:3]) + ("..." if len(bep_files) > 3 else ""))
else:
    eir_files, bep_files = [], []

st.divider()


# ── PASO 4: Evaluar ────────────────────────────────────────────────────────────
st.markdown('<p class="step-header">④ Evaluación con IA</p>', unsafe_allow_html=True)

can_eval = api_key and planilla_ok and alumno and doc_files

if not api_key:
    st.info("👈 Ingresá tu API Key de Gemini en el panel lateral.")
elif not planilla_ok:
    st.info("⬆️ Subí la planilla Excel en el paso 1.")
elif not alumno:
    st.info("Seleccioná un alumno.")
elif not doc_files:
    st.info("⬆️ Subí los archivos del alumno en el paso 3.")
else:
    if st.button("🚀 Evaluar con IA", type="primary", use_container_width=True):
        with st.spinner("Leyendo documentos..."):
            eir_text = "\n\n".join(extract_text(f) for f in eir_files) if eir_files else ""
            bep_text = "\n\n".join(extract_text(f) for f in bep_files) if bep_files else ""

        col_a, col_b = st.columns(2)
        with col_a:
            with st.spinner("Evaluando EIR..."):
                try:
                    st.session_state.eir_results = evaluar_documento(api_key, "EIR", eir_text, EIR_CRITERIA)
                    st.success("✓ EIR evaluado")
                except Exception as e:
                    st.error(f"Error EIR: {e}")
                    st.session_state.eir_results = None

        with col_b:
            with st.spinner("Evaluando BEP..."):
                try:
                    st.session_state.bep_results = evaluar_documento(api_key, "BEP", bep_text, BEP_CRITERIA)
                    st.success("✓ BEP evaluado")
                except Exception as e:
                    st.error(f"Error BEP: {e}")
                    st.session_state.bep_results = None

        st.session_state.alumno_apellido = alumno["apellido"] if alumno else ""

st.divider()


# ── PASO 5: Revisar y descargar ────────────────────────────────────────────────
st.markdown('<p class="step-header">⑤ Revisá, editá y descargá</p>', unsafe_allow_html=True)

ESTADO_OPTIONS = ["SI", "NO", "REV", "INC", "N/C"]


def results_to_df(results, criteria):
    rows = []
    for r in results:
        idx = r.get("id", -1)
        criterio = criteria[idx][0] if 0 <= idx < len(criteria) else r.get("criterio", "")
        rows.append({"Criterio": criterio, "ESTADO": r.get("estado", ""), "Observación": r.get("obs", "")})
    return pd.DataFrame(rows)


def df_to_results(df, criteria):
    return [
        {"id": i, "criterio": criteria[i][0] if i < len(criteria) else row["Criterio"],
         "estado": row["ESTADO"], "obs": row["Observación"]}
        for i, row in df.iterrows()
    ]


if not (st.session_state.eir_results or st.session_state.bep_results):
    st.info("Los resultados aparecerán acá después de evaluar.")
else:
    tab_eir, tab_bep = st.tabs(["📋 EIR", "📋 BEP"])

    with tab_eir:
        if st.session_state.eir_results:
            df_eir = results_to_df(st.session_state.eir_results, EIR_CRITERIA)
            edited_eir = st.data_editor(
                df_eir,
                column_config={
                    "Criterio":    st.column_config.TextColumn(width="large", disabled=True),
                    "ESTADO":      st.column_config.SelectboxColumn(options=ESTADO_OPTIONS, width="small"),
                    "Observación": st.column_config.TextColumn(width="large"),
                },
                hide_index=True, use_container_width=True, key="edit_eir",
            )
        else:
            st.info("Sin resultados EIR.")
            edited_eir = None

    with tab_bep:
        if st.session_state.bep_results:
            df_bep = results_to_df(st.session_state.bep_results, BEP_CRITERIA)
            edited_bep = st.data_editor(
                df_bep,
                column_config={
                    "Criterio":    st.column_config.TextColumn(width="large", disabled=True),
                    "ESTADO":      st.column_config.SelectboxColumn(options=ESTADO_OPTIONS, width="small"),
                    "Observación": st.column_config.TextColumn(width="large"),
                },
                hide_index=True, use_container_width=True, key="edit_bep",
            )
        else:
            st.info("Sin resultados BEP.")
            edited_bep = None

    # Descargar planilla actualizada
    st.divider()
    apellido_eval = st.session_state.alumno_apellido or ""
    eir_final = df_to_results(edited_eir, EIR_CRITERIA) if edited_eir is not None else (st.session_state.eir_results or [])
    bep_final = df_to_results(edited_bep, BEP_CRITERIA) if edited_bep is not None else (st.session_state.bep_results or [])

    try:
        xlsx_bytes = escribir_resultados(
            st.session_state.planilla_bytes, apellido_eval, eir_final, bep_final
        )
        st.download_button(
            label=f"⬇️ Descargar planilla con correcciones de {apellido_eval}",
            data=xlsx_bytes,
            file_name=f"COHORTE12_CORRECCIONES_{apellido_eval.upper()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
        st.caption("💡 Guardá este archivo. La próxima corrección subilo en el Paso 1 para acumular todas las correcciones.")
    except Exception as e:
        st.error(f"Error generando la planilla: {e}")
