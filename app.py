import streamlit as st
import pandas as pd
import re

st.title("Llistats: (1) grup vàlid i (2) sense grup vàlid excloent els del llistat principal")

# --- Configuració de columnes (índexs 0-based) ---
COL_ALUMNE   = 1    # "Alumno/Alumne"
COL_GRUP     = 47   # Grup
COL_NOM      = 6    # Nom
COL_COGNOM1  = 7    # Primer cognom
COL_COGNOM2  = 8    # Segon cognom
COL_DNI      = 9    # DNI (ajusta si cal)
COL_CORREU   = 11   # Correu corporatiu (ajusta si cal)

# Llista de grups exactes
GRUPS = [
    "G1A","G1B","G2a","G2b","G3a","G3b","G4a","G4b",
    "IN1a","IN1b","IN2a","IN2b","IN3a","IN3b","In3c",
    "IN4a","IN4b","M1","M2","M3","M4","P1","P2","P3","P4"
]
GRUPS_SET = set(GRUPS)

# Regex per detectar "alumne/alumno"
RE_ALUMNE = re.compile(r"\b(alumno|alumne)\b", flags=re.IGNORECASE)

def make_student_key(row: pd.Series) -> str:
    """Clau d'alumne per deduplicar i excloure: DNI si existeix; en cas contrari Nom+Cognoms."""
    dni = (row.get("DNI", "") or "").strip()
    if dni:
        return f"DNI:{dni}"
    nom  = (row.get("Nom", "") or "").strip()
    c1   = (row.get("Primer Cognom", "") or "").strip()
    c2   = (row.get("Segon Cognom", "") or "").strip()
    return f"NOM:{nom}|C1:{c1}|C2:{c2}"

def dedup_first_appearance(df_in: pd.DataFrame) -> pd.DataFrame:
    """Conserva la primera aparició per alumne (segons ordre original) usant la clau d'alumne."""
    df = df_in.copy()
    if "__ordre__" not in df.columns:
        df["__ordre__"] = range(len(df))
    df["__key__"] = df.apply(make_student_key, axis=1)
    df = df.drop_duplicates(subset="__key__", keep="first")
    df = df.sort_values("__ordre__", kind="stable")
    return df

arxiu = st.file_uploader("Puja el teu Excel (.xls o .xlsx)", type=["xls", "xlsx"])

if arxiu:
    # Llegim sense encapçalaments perquè els índexs siguin 0..N
    if arxiu.name.endswith(".xls"):
        df = pd.read_excel(arxiu, engine="xlrd", header=None)
    else:
        df = pd.read_excel(arxiu, engine="openpyxl", header=None)

    st.caption(f"Arxiu carregat: {df.shape[0]} files × {df.shape[1]} columnes")
    st.write("Vista prèvia:")
    st.dataframe(df.head().astype(str))

    # Comprovació d'índexs necessaris
    necessaries = [COL_ALUMNE, COL_GRUP, COL_NOM, COL_COGNOM1, COL_COGNOM2, COL_DNI, COL_CORREU]
    fora_rang = [c for c in necessaries if c < 0 or c >= df.shape[1]]
    if fora_rang:
        st.error(f"Columnes fora de rang: {fora_rang}. L'arxiu té {df.shape[1]} columnes (0..{df.shape[1]-1}).")
        st.stop()

    # Normalitzar a text
    df = df.fillna("").astype(str)

    # --- Filtres base ---
    filtre_alumne = df.iloc[:, COL_ALUMNE].str.strip().str.contains(RE_ALUMNE)
    col_grup = df.iloc[:, COL_GRUP].str.strip()
    filtre_grup_valid = col_grup.isin(GRUPS_SET)

    st.write(f"Files amb 'alumne' a col {COL_ALUMNE}: **{int(filtre_alumne.sum())}**")
    st.write(f"Files amb grup vàlid a col {COL_GRUP}: **{int(filtre_grup_valid.sum())}**")

    # ========== (1) LLISTAT PRINCIPAL: ALUMNE + GRUP VÀLID ==========
    filtre_principal = filtre_alumne & filtre_grup_valid
    st.write(f"(1) Files per llistat principal (alumne + grup vàlid): **{int(filtre_principal.sum())}**")

    llistat_final = pd.DataFrame()
    keys_principal = set()

    if int(filtre_principal.sum()) == 0:
        st.warning("No hi ha files per al llistat principal (alumne + grup vàlid).")
    else:
        llistat = df.iloc[filtre_principal.values,
                          [COL_NOM, COL_COGNOM1, COL_COGNOM2, COL_DNI, COL_CORREU, COL_GRUP]].copy()
        llistat.columns = ["Nom", "Primer Cognom", "Segon Cognom", "DNI", "Correu corporatiu", "Grup"]
        llistat = llistat.apply(lambda s: s.str.strip())
        llistat["Verificació DNI"] = ""
        llistat["__ordre__"] = llistat.index

        # Dedup per alumne (primera aparició)
        llistat_final = dedup_first_appearance(llistat).copy()
        # Guarda les claus d'alumnes que SÍ tenen grup vàlid (per excloure del 2n llistat)
        llistat_final["__key__"] = llistat_final.apply(make_student_key, axis=1)
        keys_principal = set(llistat_final["__key__"].tolist())
        # Neteja columna auxiliar abans de mostrar/descarregar
        llistat_final = llistat_final.drop(columns=["__key__"])

        st.subheader("(1) Llistat principal — primera aparició per alumne")
        st.caption(f"Files: {llistat_final.shape[0]}")
        st.dataframe(llistat_final)

        csv1 = llistat_final.to_csv(index=False).encode("utf-8")
        st.download_button(
            "Descarregar CSV (principal — grup vàlid)",
            data=csv1,
            file_name="llistat_filtrat_primera_aparicio.csv",
            mime="text/csv"
        )

    # ========== (2) SEGON LLISTAT: ALUMNE + SENSE GRUP VÀLID, EXCLOENT ELS DEL (1) ==========
    filtre_sense_grup = filtre_alumne & (~filtre_grup_valid)
    st.write(f"(2) Files amb 'alumne' però sense grup vàlid (abans d'excloure els del principal): **{int(filtre_sense_grup.sum())}**")

    if int(filtre_sense_grup.sum()) == 0:
        st.info("No hi ha files per al segon llistat (alumne sense grup vàlid).")
    else:
        llistat_sense = df.iloc[filtre_sense_grup.values,
                                [COL_NOM, COL_COGNOM1, COL_COGNOM2, COL_DNI, COL_CORREU, COL_GRUP]].copy()
        llistat_sense.columns = ["Nom", "Primer Cognom", "Segon Cognom", "DNI", "Correu corporatiu", "Grup"]
        llistat_sense = llistat_sense.apply(lambda s: s.str.strip())
        llistat_sense["Verificació DNI"] = ""
        llistat_sense["__ordre__"] = llistat_sense.index

        # Calcula la clau d'alumne i EXCLOU els que ja estan al llistat principal
        llistat_sense["__key__"] = llistat_sense.apply(make_student_key, axis=1)
        if keys_principal:
            llistat_sense = llistat_sense[~llistat_sense["__key__"].isin(keys_principal)]

        if llistat_sense.empty:
            st.info("Tots els alumnes sense grup vàlid ja apareixen al llistat principal; no hi ha cap pendent.")
        else:
            # Dedup per alumne (primera aparició) i neteja de columnes auxiliars
            llistat_sense_final = dedup_first_appearance(llistat_sense).drop(columns=["__key__"])

            st.subheader("(2) Llistat sense grup vàlid — excloent els del principal (primera aparició)")
            st.caption(f"Files: {llistat_sense_final.shape[0]}")
            st.dataframe(llistat_sense_final)

            csv2 = llistat_sense_final.to_csv(index=False).encode("utf-8")
            st.download_button(
                "Descarregar CSV (alumne sense grup vàlid, excloent principal)",
                data=csv2,
                file_name="llistat_sense_grup_valid_excloent_principal.csv",
                mime="text/csv"
            )
