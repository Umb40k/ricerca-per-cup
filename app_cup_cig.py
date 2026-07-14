"""
App Streamlit per la ricerca inversa CUP -> CIG.

Dato uno o più CUP, l'app:
1. Interroga il dataset open data ANAC "cup" (mappatura CIG<->CUP) per
   trovare tutti i CIG collegati a ciascun CUP indicato.
2. Per ciascun CIG trovato, interroga l'API ANAC getSmartCig per
   recuperare i dettagli: stazione appaltante, CPV, date di pubblicazione,
   codice risposta.
3. Presenta i risultati in una tabella navigabile e scaricabile.

Avvio:
    pip install -r requirements.txt
    streamlit run app_cup_cig.py
"""

import io
import time
from io import BytesIO

import pandas as pd
import requests
import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------
# Configurazione
# --------------------------------------------------------------------------

CKAN_API_BASE = "https://dati.anticorruzione.it/opendata/api/3/action/"
CIG_API_BASE = "https://api.anticorruzione.it/apicig/1.0.0/getSmartCig/"
RICHIESTA_TIMEOUT = 30
DOWNLOAD_TIMEOUT = 120
PAUSA_TRA_RICHIESTE = 0.5
MAX_TENTATIVI = 5

COLONNE_RISULTATO = [
    "CUP",
    "CIG",
    "Codice risposta",
    "Stazione Appaltante",
    "Citta",
    "Regione",
    "Codice AUSA",
    "Data pubblicazione bando",
    "Data creazione CIG",
    "CPV",
]

COLORE_PRIMARIO = "#0B3D62"
COLORE_SECONDARIO = "#EEF3F8"

st.set_page_config(
    page_title="Ricerca CIG collegati a un CUP — ANAC",
    page_icon="🔎",
    layout="wide",
)

# --------------------------------------------------------------------------
# Stile (stesso stile dell'app CIG -> CUP)
# --------------------------------------------------------------------------

st.markdown(
    f"""
    <style>
    .main {{
        background-color: #FAFBFC;
    }}
    .app-header {{
        background-color: {COLORE_PRIMARIO};
        padding: 1.6rem 2rem;
        border-radius: 8px;
        margin-bottom: 1.5rem;
    }}
    .app-header h1 {{
        color: white;
        font-size: 1.5rem;
        margin: 0;
        font-weight: 600;
    }}
    .app-header p {{
        color: #D6E2EE;
        margin: 0.3rem 0 0 0;
        font-size: 0.95rem;
    }}
    .info-box {{
        background-color: {COLORE_SECONDARIO};
        border-left: 4px solid {COLORE_PRIMARIO};
        padding: 0.9rem 1.2rem;
        border-radius: 4px;
        margin-bottom: 1rem;
        font-size: 0.9rem;
    }}
    div[data-testid="stMetricValue"] {{
        font-size: 1.6rem;
        color: {COLORE_PRIMARIO};
    }}
    .stButton > button {{
        background-color: {COLORE_PRIMARIO};
        color: white;
        border-radius: 6px;
        padding: 0.5rem 1.5rem;
        border: none;
        font-weight: 600;
    }}
    .stButton > button:hover {{
        background-color: #0A3352;
        color: white;
    }}
    footer {{visibility: hidden;}}
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="app-header">
        <h1>Ricerca CIG collegati a un CUP</h1>
        <p>Trova tutti i CIG associati a uno o più CUP e recupera i relativi dettagli dall'API ANAC</p>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="info-box">
    Questa app cerca i CIG collegati a un CUP incrociando il dataset open data ANAC
    "cup" (mappatura CIG↔CUP degli appalti ordinari), poi arricchisce ogni CIG trovato
    con i dettagli restituiti dall'API pubblica <b>getSmartCig</b> (stazione appaltante,
    categoria merceologica CPV, date di pubblicazione).<br><br>
    <b>Nota sui dati economico-contabili</b>: informazioni di dettaglio su impegni,
    pagamenti e obbligazioni finanziarie non sono esposte da questa API pubblica in
    tempo reale; richiedono l'accesso ai dataset ANAC dedicati (es. "quadro-economico",
    "stati-avanzamento") o al portale ReGiS per i progetti PNRR. Contattaci se ti serve
    ampliare l'app anche su questi dati.
    </div>
    """,
    unsafe_allow_html=True,
)


# --------------------------------------------------------------------------
# Funzioni: recupero dataset CIG<->CUP
# --------------------------------------------------------------------------


@st.cache_data(show_spinner=False, ttl=60 * 60 * 24)
def risolvi_url_dataset_cup() -> str:
    """Interroga l'API CKAN di ANAC per trovare l'URL di download aggiornato
    del file CSV del dataset 'cup'."""
    resp = requests.get(
        CKAN_API_BASE + "package_show",
        params={"id": "cup"},
        timeout=RICHIESTA_TIMEOUT,
    )
    resp.raise_for_status()
    dati = resp.json()
    risorse = dati.get("result", {}).get("resources", [])
    for r in risorse:
        if r.get("format", "").upper() == "CSV" and "log" not in r.get("name", "").lower():
            return r["url"]
    # fallback: prima risorsa CSV disponibile
    for r in risorse:
        if r.get("format", "").upper() == "CSV":
            return r["url"]
    raise RuntimeError("Nessuna risorsa CSV trovata nel dataset 'cup' di ANAC.")


@st.cache_data(show_spinner=False, ttl=60 * 60 * 24)
def scarica_dataset_cup() -> pd.DataFrame:
    """Scarica e carica in memoria il dataset CIG<->CUP di ANAC."""
    url = risolvi_url_dataset_cup()
    resp = requests.get(url, timeout=DOWNLOAD_TIMEOUT, stream=True)
    resp.raise_for_status()
    contenuto = resp.content
    df = pd.read_csv(io.BytesIO(contenuto), dtype=str, sep=None, engine="python")
    df.columns = [c.strip().upper() for c in df.columns]
    return df


def trova_cig_per_cup(df_cup: pd.DataFrame, cup_list: list) -> pd.DataFrame:
    """Filtra il dataset CIG<->CUP per i CUP richiesti."""
    colonna_cup = next((c for c in df_cup.columns if "CUP" in c), None)
    colonna_cig = next((c for c in df_cup.columns if "CIG" in c), None)
    if colonna_cup is None or colonna_cig is None:
        raise RuntimeError(
            f"Colonne CUP/CIG non riconosciute nel dataset. Colonne trovate: {list(df_cup.columns)}"
        )
    cup_set = set(c.strip().upper() for c in cup_list)
    df_filtrato = df_cup[df_cup[colonna_cup].str.strip().str.upper().isin(cup_set)]
    return df_filtrato[[colonna_cup, colonna_cig]].rename(
        columns={colonna_cup: "CUP", colonna_cig: "CIG"}
    )


# --------------------------------------------------------------------------
# Funzioni: arricchimento dettagli CIG via getSmartCig
# --------------------------------------------------------------------------


def interroga_cig(cig: str) -> dict:
    import random

    url = CIG_API_BASE + cig
    ultimo_errore = None
    for tentativo in range(1, MAX_TENTATIVI + 1):
        try:
            resp = requests.get(url, timeout=RICHIESTA_TIMEOUT)
            if resp.status_code >= 500:
                ultimo_errore = f"{resp.status_code} Server Error"
                time.sleep(min(2 ** tentativo, 20) + random.uniform(0, 1))
                continue
            resp.raise_for_status()
            return resp.json()
        except Exception as e:
            ultimo_errore = e
            time.sleep(min(2 ** tentativo, 20) + random.uniform(0, 1))
    return {"errore": str(ultimo_errore)}


def estrai_dati(cup: str, cig: str, risposta: dict) -> dict:
    riga = {
        "CUP": cup,
        "CIG": cig,
        "Codice risposta": risposta.get("codice_risposta", ""),
        "Stazione Appaltante": "",
        "Citta": "",
        "Regione": "",
        "Codice AUSA": "",
        "Data pubblicazione bando": "",
        "Data creazione CIG": "",
        "CPV": "",
    }

    if "errore" in risposta:
        riga["Codice risposta"] = f"ERRORE RETE: {risposta['errore']}"
        return riga

    bando = risposta.get("bando") or {}
    cpv_list = bando.get("CPV") or []
    cpv_desc = []
    for c in cpv_list:
        cod = c.get("COD_CPV", "")
        desc = c.get("DESCRIZIONE_CPV", "")
        if cod or desc:
            cpv_desc.append(f"{cod} - {desc}")
    riga["CPV"] = "; ".join(cpv_desc)

    sa = risposta.get("stazione_appaltante") or {}
    riga["Stazione Appaltante"] = sa.get("DENOMINAZIONE_AMMINISTRAZIONE_APPALTANTE", "")
    riga["Citta"] = sa.get("CITTA", "")
    riga["Regione"] = sa.get("REGIONE", "")
    riga["Codice AUSA"] = sa.get("CODICE_AUSA", "")

    pub = risposta.get("pubblicazioni") or {}
    riga["Data pubblicazione bando"] = pub.get("DATA_PUBBLICAZIONE", "")
    riga["Data creazione CIG"] = pub.get("DATA_CREAZIONE", "")

    return riga


def genera_excel_risultato(df: pd.DataFrame) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "CIG collegati a CUP"

    for col_idx, col_name in enumerate(COLONNE_RISULTATO, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="0B3D62", end_color="0B3D62", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row_idx, row in enumerate(df.itertuples(index=False), start=2):
        for col_idx, val in enumerate(row, start=1):
            ws.cell(row=row_idx, column=col_idx, value=val)

    larghezze = [16, 16, 14, 42, 20, 18, 12, 20, 20, 45]
    for i, larg in enumerate(larghezze, start=1):
        ws.column_dimensions[get_column_letter(i)].width = larg

    ws.freeze_panes = "A2"
    buffer = BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


# --------------------------------------------------------------------------
# Interfaccia
# --------------------------------------------------------------------------

col_input, col_info = st.columns([2, 1])

with col_input:
    tab_testo, tab_file = st.tabs(["Incolla i CUP", "Carica file Excel"])

    cup_list = []

    with tab_testo:
        testo_cup = st.text_area(
            "Inserisci uno o più CUP (uno per riga)",
            height=140,
            placeholder="F16H21000000008\nF55G21000000001\n...",
        )
        if testo_cup.strip():
            cup_list = [c.strip() for c in testo_cup.splitlines() if c.strip()]

    with tab_file:
        file_caricato = st.file_uploader("File Excel con una colonna CUP", type=["xlsx"])
        if file_caricato is not None:
            try:
                wb_up = openpyxl.load_workbook(file_caricato, data_only=True)
                ws_up = wb_up[wb_up.sheetnames[0]]
                intestazioni = [c.value for c in next(ws_up.iter_rows(min_row=1, max_row=1))]
                colonna_idx = None
                for idx, val in enumerate(intestazioni, start=1):
                    if val and "CUP" in str(val).upper():
                        colonna_idx = idx
                        break
                if colonna_idx is None:
                    colonna_idx = 1
                for row in ws_up.iter_rows(min_row=2, values_only=True):
                    val = row[colonna_idx - 1]
                    if val:
                        cup_list.append(str(val).strip())
            except Exception as e:
                st.error(f"Impossibile leggere il file: {e}")

with col_info:
    st.markdown("**Informazioni**")
    st.caption("Fonte mappatura CIG↔CUP: dataset open data ANAC 'cup'")
    st.caption("Fonte dettagli CIG: API pubblica ANAC — getSmartCig")
    st.caption("Il dataset CUP viene scaricato una volta e riutilizzato per 24 ore")

if cup_list:
    cup_list = list(dict.fromkeys(cup_list))  # rimuove duplicati mantenendo l'ordine
    st.success(f"**{len(cup_list)}** CUP pronti per la ricerca.")

    avvia = st.button("Cerca CIG collegati", type="primary")

    if avvia:
        with st.spinner("Recupero il dataset CIG↔CUP di ANAC (può richiedere qualche secondo)..."):
            try:
                df_cup_dataset = scarica_dataset_cup()
            except Exception as e:
                st.error(
                    f"Impossibile scaricare il dataset ANAC 'cup': {e}\n\n"
                    "Puoi comunque scaricarlo manualmente da "
                    "https://dati.anticorruzione.it/opendata/dataset/cup e ricaricarlo qui sotto."
                )
                df_cup_dataset = None

        if df_cup_dataset is not None:
            corrispondenze = trova_cig_per_cup(df_cup_dataset, cup_list)

            cup_trovati = set(corrispondenze["CUP"].str.upper())
            cup_non_trovati = [c for c in cup_list if c.upper() not in cup_trovati]

            if corrispondenze.empty:
                st.warning("Nessun CIG trovato per i CUP indicati nel dataset ANAC.")
            else:
                st.info(
                    f"Trovati **{len(corrispondenze)}** CIG collegati a "
                    f"**{len(cup_trovati)}** CUP su {len(cup_list)} richiesti."
                )

                barra_avanzamento = st.progress(0, text="Recupero dettagli CIG in corso...")
                placeholder_tabella = st.empty()
                risultati = []

                righe = corrispondenze.to_dict("records")
                for i, riga_match in enumerate(righe, start=1):
                    cup = riga_match["CUP"]
                    cig = riga_match["CIG"]
                    risposta = interroga_cig(cig)
                    risultati.append(estrai_dati(cup, cig, risposta))

                    barra_avanzamento.progress(
                        i / len(righe),
                        text=f"CIG {i} di {len(righe)}: {cig} (CUP {cup})",
                    )

                    df_parziale = pd.DataFrame(risultati, columns=COLONNE_RISULTATO)
                    placeholder_tabella.dataframe(df_parziale, use_container_width=True, height=350)

                    time.sleep(PAUSA_TRA_RICHIESTE)

                barra_avanzamento.empty()
                df_finale = pd.DataFrame(risultati, columns=COLONNE_RISULTATO)

                st.markdown("### Riepilogo per CUP")
                riepilogo = (
                    df_finale.groupby("CUP")["CIG"]
                    .nunique()
                    .reset_index()
                    .rename(columns={"CIG": "Numero CIG collegati"})
                )
                if cup_non_trovati:
                    df_non_trovati = pd.DataFrame(
                        {"CUP": cup_non_trovati, "Numero CIG collegati": 0}
                    )
                    riepilogo = pd.concat([riepilogo, df_non_trovati], ignore_index=True)
                st.dataframe(riepilogo, use_container_width=True, height=min(400, 40 + 35 * len(riepilogo)))

                st.markdown("### Dettaglio completo")
                st.dataframe(df_finale, use_container_width=True, height=400)

                excel_bytes = genera_excel_risultato(df_finale)
                csv_bytes = df_finale.to_csv(index=False).encode("utf-8-sig")

                col_dl1, col_dl2 = st.columns(2)
                with col_dl1:
                    st.download_button(
                        label="Scarica risultati (Excel)",
                        data=excel_bytes,
                        file_name="cig_collegati_a_cup.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                with col_dl2:
                    st.download_button(
                        label="Scarica risultati (CSV)",
                        data=csv_bytes,
                        file_name="cig_collegati_a_cup.csv",
                        mime="text/csv",
                        use_container_width=True,
                    )

            if cup_non_trovati:
                st.warning(
                    "I seguenti CUP non risultano nel dataset ANAC (potrebbero riguardare "
                    "affidamenti non ordinari, essere di recente emissione o non ancora "
                    "propagati nel dataset open data): " + ", ".join(cup_non_trovati)
                )
else:
    st.info("Incolla uno o più CUP, oppure carica un file Excel, per iniziare.")

st.markdown("---")
st.caption(
    "Dati recuperati dai dataset open data e dall'API pubblica dell'Autorità Nazionale "
    "Anticorruzione (ANAC). L'applicazione non memorizza né trasmette i dati a terzi: "
    "l'elaborazione avviene solo durante la sessione attiva."
)
