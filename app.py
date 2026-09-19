import re
import altair as alt
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="🏆 Liga Lluis Galvart GB",
    layout="wide",
    initial_sidebar_state="collapsed",
)

BYE_NAMES = {"descansa", "libre"}
PPD_JUEGOS = {"501", "701"}
MPR_JUEGOS = {"cricket"}
ORO = "#FCE68E"
PLATA = "#F0F0F0"
BRONCE = "#E6C2A5"
AZUL_ELECTRICO = "#247CFF"
NARANJA_MPR = "#FF6B00"
HAZANAS = ["LOW TON", "HAT TRICK", "6M", "7M", "8M", "9M"]

st.markdown(
    """
    <style>
      .stApp {
        background-image:
          linear-gradient(rgba(255, 255, 255, 0.85), rgba(255, 255, 255, 0.85)),
          url("https://images.unsplash.com/photo-1574169208507-843761580d19?q=80&w=2000&auto=format&fit=crop");
        background-size: cover;
        background-position: center center;
        background-attachment: fixed;
        background-repeat: no-repeat;
      }
      .block-container {
        padding-top: 1.4rem;
        padding-bottom: 2.4rem;
        max-width: 1280px;
        background: rgba(255, 255, 255, 0.78);
        border-radius: 20px;
        box-shadow: 0 12px 40px rgba(0, 0, 0, 0.08);
      }
      h1, h2, h3 { letter-spacing: 0.04em; color: #111 !important; }
      [data-testid="stDataFrame"],
      [data-testid="stDataFrameResizable"] {
        background: rgba(255, 255, 255, 0.96) !important;
        border-radius: 12px;
      }
      div[data-testid="stMetric"] {
        background: linear-gradient(180deg, #1a1a1a 0%, #111 100%);
        border: 1px solid #c9a22755;
        border-radius: 16px;
        padding: 12px 16px;
      }
      div[data-testid="stMetric"] label { color: #d4af37 !important; font-weight: 700 !important; }
      div[data-testid="stMetric"] [data-testid="stMetricValue"] { color: #fff8dc !important; }
      .liga-kicker {
        color: #d4af37; font-weight: 700; letter-spacing: 0.28em;
        text-transform: uppercase; font-size: 0.78rem; margin-bottom: 0.2rem;
      }
    </style>
    """,
    unsafe_allow_html=True,
)


def _strip_series(s: pd.Series) -> pd.Series:
    return s.astype(str).str.strip()


def _norm_name(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def _name_key(value) -> str:
    return _norm_name(value).casefold()


def _is_bye(name: str) -> bool:
    return _name_key(name) in BYE_NAMES


def _jornada_num(value) -> int | None:
    if pd.isna(value):
        return None
    match = re.search(r"\d+", str(value))
    if match:
        return int(match.group())
    return None


def _juego_tipo(value) -> str:
    text = _norm_name(value).casefold()
    if "501" in text or "701" in text:
        return "PPD"
    if "cricket" in text or "cr." in text or text == "cr" or "standard cr" in text:
        return "MPR"
    return ""


def _to_int(value) -> int:
    n = pd.to_numeric(value, errors="coerce")
    if pd.isna(n):
        return 0
    return int(n)


def _rename_resultados(df: pd.DataFrame) -> pd.DataFrame:
    mapa = {
        "DIVISION": "División",
        "DIVISIÓN": "División",
        "JORNADA": "Jornada",
        "SET": "Set",
        "LEG": "Leg",
        "JUEGO": "Juego",
        "JUGADOR 1": "Jugador 1",
        "JUGADOR 2": "Jugador 2",
        "MEDIA JUGADOR 1": "Media J1",
        "MEDIA JUGADOR 2": "Media J2",
        "MEDIA J1": "Media J1",
        "MEDIA J2": "Media J2",
        "GANADOR LEG": "Ganador Leg",
    }
    return df.rename(columns={c: mapa.get(str(c).strip().upper(), str(c).strip()) for c in df.columns})


def _rename_calendario(df: pd.DataFrame) -> pd.DataFrame:
    mapa = {
        "DIVISION": "División",
        "DIVISIÓN": "División",
        "JORNADA": "Jornada",
        "JUGADOR 1": "Jugador 1",
        "JUGADOR 2": "Jugador 2",
    }
    return df.rename(columns={c: mapa.get(str(c).strip().upper(), str(c).strip()) for c in df.columns})


@st.cache_data
def cargar_datos():
    try:
        df_resultados = pd.read_excel("datos_liga.xlsx", sheet_name="Resultados", header=1)
        df_calendario = pd.read_excel("datos_liga.xlsx", sheet_name="Calendario", header=1)
        
        df_resultados = _rename_resultados(df_resultados)
        df_calendario = _rename_calendario(df_calendario)

        for col in ["División", "Jornada", "Set", "Leg", "Juego", "Jugador 1", "Jugador 2", "Ganador Leg"]:
            if col in df_resultados.columns:
                df_resultados[col] = _strip_series(df_resultados[col])
                
        for col in ["División", "Jornada", "Jugador 1", "Jugador 2"]:
            if col in df_calendario.columns:
                df_calendario[col] = _strip_series(df_calendario[col])
                
        for col in ["Media J1", "Media J2"]:
            if col in df_resultados.columns:
                df_resultados[col] = pd.to_numeric(df_resultados[col], errors="coerce")
                
        for h in HAZANAS:
            for prefijo in ("J1", "J2"):
                col = f"{prefijo} {h}"
                if col in df_resultados.columns:
                    df_resultados[col] = pd.to_numeric(df_resultados[col], errors="coerce").fillna(0)

        df_resultados = df_resultados.dropna(subset=["Jugador 1", "Jugador 2"], how="any")
        df_calendario = df_calendario.dropna(subset=["Jugador 1", "Jugador 2"], how="any")
        
        return df_resultados, df_calendario
    except Exception as e:
        st.error(f"Error al leer datos_liga.xlsx: {e}")
        return None, None


def partidos_desde_resultados(df_resultados: pd.DataFrame) -> pd.DataFrame:
    if df_resultados.empty:
        return pd.DataFrame()

    df = df_resultados.copy()
    df["jornada_n"] = df["Jornada"].map(_jornada_num)
    df["p1"] = df["Jugador 1"].map(_name_key)
    df["p2"] = df["Jugador 2"].map(_name_key)
    df["ganador_k"] = df["Ganador Leg"].map(_name_key)
    df["pareja"] = df.apply(lambda r: tuple(sorted((r["p1"], r["p2"]))), axis=1)
    
    if "Set" not in df.columns:
        df["Set"] = 1

    filas = []
    for (division, jornada_n, pareja), g_partido in df.groupby(["División", "jornada_n", "pareja"], dropna=False):
        sets_p = {pareja[0]: 0, pareja[1]: 0}
        
        for _, g_set in g_partido.groupby("Set"):
            legs = g_set["ganador_k"].value_counts()
            if legs.empty:
                continue
            max_legs = legs.max()
            ganadores = legs[legs == max_legs].index.tolist()
            if len(ganadores) == 1 and ganadores[0] in sets_p:
                sets_p[ganadores[0]] += 1

        a, b = pareja
        sa, sb = sets_p[a], sets_p[b]
        
        if sa > sb:
            ganador = a
        elif sb > sa:
            ganador = b
        else:
            ganador = None
            
        filas.append({
            "División": division,
            "jornada_n": jornada_n,
            "pareja": pareja,
            "sets": sets_p,
            "ganador": ganador,
        })
        
    return pd.DataFrame(filas)


def _legs_por_jugador(df_resultados: pd.DataFrame) -> pd.DataFrame:
    registros = []
    for _, row in df_resultados.iterrows():
        tipo = _juego_tipo(row.get("Juego"))
        if not tipo:
            continue
            
        jornada_n = _jornada_num(row.get("Jornada"))
        p1 = _name_key(row.get("Jugador 1"))
        p2 = _name_key(row.get("Jugador 2"))
        
        if not p1 or not p2 or _is_bye(p1) or _is_bye(p2):
            continue
            
        pareja = tuple(sorted((p1, p2)))
        
        for jugador_col, media_col in (("Jugador 1", "Media J1"), ("Jugador 2", "Media J2")):
            jugador = _name_key(row.get(jugador_col))
            media = row.get(media_col)
            
            if not jugador or _is_bye(jugador) or pd.isna(media):
                continue
                
            registros.append({
                "División": row["División"],
                "jornada_n": jornada_n,
                "pareja": pareja,
                "jugador": jugador,
                "tipo": tipo,
                "media": float(media),
            })
            
    return pd.DataFrame(registros)


def medias_por_partido(df_resultados: pd.DataFrame) -> pd.DataFrame:
    legs = _legs_por_jugador(df_resultados)
    if legs.empty:
        return pd.DataFrame(columns=["División", "jornada_n", "pareja", "jugador", "tipo", "media"])
        
    return legs.groupby(["División", "jornada_n", "pareja", "jugador", "tipo"], as_index=False)["media"].mean()


def medias_jugadores(df_resultados: pd.DataFrame) -> pd.DataFrame:
    legs = _legs_por_jugador(df_resultados)
    if legs.empty:
        return pd.DataFrame(columns=["jugador", "PPD", "MPR"])
        
    resumen = legs.groupby(["jugador", "tipo"])["media"].mean().unstack("tipo").reset_index()
    
    for col in ("PPD", "MPR"):
        if col not in resumen.columns:
            resumen[col] = pd.NA
            
    return resumen[["jugador", "PPD", "MPR"]]


def contar_legs(df_resultados: pd.DataFrame, division: str) -> dict:
    df_div = df_resultados[df_resultados["División"].astype(str).str.casefold() == division.casefold()]
    stats = {}
    
    for _, row in df_div.iterrows():
        p1 = _name_key(row.get("Jugador 1"))
        p2 = _name_key(row.get("Jugador 2"))
        ganador = _name_key(row.get("Ganador Leg"))
        
        if not p1 or not p2 or _is_bye(p1) or _is_bye(p2):
            continue
            
        for p in (p1, p2):
            if p not in stats:
                stats[p] = {"LF": 0, "LC": 0, "Legs Jugados": 0}
            
            stats[p]["Legs Jugados"] += 1
            
            if p == ganador:
                stats[p]["LF"] += 1
            else:
                stats[p]["LC"] += 1
                
    return stats


def hazañas_jugadores(df_resultados: pd.DataFrame, division: str) -> pd.DataFrame:
    totales = {}
    df = df_resultados[df_resultados["División"].astype(str).str.casefold() == division.casefold()]
    
    lados = (
        ("Jugador 1", {h: f"J1 {h}" for h in HAZANAS}),
        ("Jugador 2", {h: f"J2 {h}" for h in HAZANAS})
    )
    
    for _, row in df.iterrows():
        for jugador_col, cols in lados:
            clave = _name_key(row.get(jugador_col))
            if not clave or _is_bye(clave):
                continue
                
            if clave not in totales:
                totales[clave] = {h: 0 for h in HAZANAS}
                
            for hazaña, col in cols.items():
                if col in df.columns:
                    totales[clave][hazaña] += _to_int(row.get(col))
                    
    if not totales:
        return pd.DataFrame(columns=["jugador", *HAZANAS])
        
    out = pd.DataFrame.from_dict(totales, orient="index").reset_index().rename(columns={"index": "jugador"})
    return out[["jugador", *HAZANAS]]


def jugadores_division(df_calendario: pd.DataFrame, division: str) -> dict:
    cal = df_calendario[df_calendario["División"].str.casefold() == division.casefold()]
    nombres = {}
    
    for col in ("Jugador 1", "Jugador 2"):
        for n in cal[col].dropna().unique():
            if _is_bye(n):
                continue
            nombres[_name_key(n)] = _norm_name(n)
            
    return nombres


def mapa_nombres(df_calendario: pd.DataFrame) -> dict:
    nombres = {}
    for division in df_calendario["División"].dropna().unique():
        nombres.update(jugadores_division(df_calendario, division))
    return nombres


def clasificacion_general(df_calendario, df_partidos, df_resultados, division: str) -> pd.DataFrame:
    nombres = jugadores_division(df_calendario, division)
    tabla = {
        k: {"Jugador": v, "PJ": 0, "PG": 0, "PP": 0, "SF": 0, "SC": 0, "LF": 0, "LC": 0} 
        for k, v in nombres.items()
    }

    if not df_partidos.empty:
        part = df_partidos[df_partidos["División"].astype(str).str.casefold() == division.casefold()]
        for _, p in part.iterrows():
            a, b = p["pareja"]
            if _is_bye(a) or _is_bye(b):
                continue
                
            sets = p["sets"]
            for jugador, rival in ((a, b), (b, a)):
                if jugador not in tabla:
                    continue
                tabla[jugador]["PJ"] += 1
                tabla[jugador]["SF"] += int(sets.get(jugador, 0))
                tabla[jugador]["SC"] += int(sets.get(rival, 0))
                
            ganador = p["ganador"]
            if ganador == a and a in tabla and b in tabla:
                tabla[a]["PG"] += 1
                tabla[b]["PP"] += 1
            elif ganador == b and a in tabla and b in tabla:
                tabla[b]["PG"] += 1
                tabla[a]["PP"] += 1

    legs_stats = contar_legs(df_resultados, division)
    for k, stat in legs_stats.items():
        if k in tabla:
            tabla[k]["LF"] = stat["LF"]
            tabla[k]["LC"] = stat["LC"]

    out = pd.DataFrame(list(tabla.values()))
    if out.empty:
        return out

    out["diff_sets"] = out["SF"] - out["SC"]
    out["diff_legs"] = out["LF"] - out["LC"]
    
    out = out.sort_values(
        ["PG", "diff_sets", "SF", "diff_legs", "LF"], 
        ascending=[False, False, False, False, False], 
        kind="mergesort"
    )
    
    out.insert(0, "Pos", range(1, len(out) + 1))
    
    return out[["Pos", "Jugador", "PJ", "PG", "PP", "SF", "SC", "LF", "LC"]].reset_index(drop=True)


def clasificacion_estadisticas(df_calendario, df_resultados, df_medias, division: str) -> pd.DataFrame:
    nombres = jugadores_division(df_calendario, division)
    tabla = {
        k: {"Jugador": v, **{h: 0 for h in HAZANAS}, "Legs Jugados": 0} 
        for k, v in nombres.items()
    }

    legs_stats = contar_legs(df_resultados, division)
    for k, stat in legs_stats.items():
        if k in tabla:
            tabla[k]["Legs Jugados"] = stat["Legs Jugados"]

    out = pd.DataFrame(list(tabla.values()))
    if out.empty:
        return out

    hazañas = hazañas_jugadores(df_resultados, division)
    if not hazañas.empty:
        out = out.merge(
            hazañas.rename(columns={"jugador": "_k"}), 
            how="left", 
            left_on=out["Jugador"].map(_name_key), 
            right_on="_k", 
            suffixes=("", "_sum")
        )
        out = out.drop(columns=["_k"], errors="ignore")
        
        for h in HAZANAS:
            col_sum = f"{h}_sum"
            if col_sum in out.columns:
                out[h] = pd.to_numeric(out[col_sum], errors="coerce").fillna(0).astype(int)
                out = out.drop(columns=[col_sum])
            elif h in out.columns:
                out[h] = pd.to_numeric(out[h], errors="coerce").fillna(0).astype(int)

    out = out.merge(
        df_medias.rename(columns={"jugador": "_k"}), 
        how="left", 
        left_on=out["Jugador"].map(_name_key), 
        right_on="_k"
    )
    out = out.drop(columns=["_k"], errors="ignore")
    
    for col in ("PPD", "MPR"):
        if col not in out.columns:
            out[col] = pd.NA
        out[col] = pd.to_numeric(out[col], errors="coerce").round(2)

    out = out.sort_values(["MPR"], ascending=[False], kind="mergesort")
    
    return out[["Jugador", *HAZANAS, "Legs Jugados", "PPD", "MPR"]].reset_index(drop=True)


def estilo_clasificacion_general(df: pd.DataFrame):
    podium = {
        1: f"background-color: {ORO}; color: #000000; font-weight: 700;",
        2: f"background-color: {PLATA}; color: #000000; font-weight: 700;",
        3: f"background-color: {BRONCE}; color: #000000; font-weight: 700;",
    }
    
    def _pinta(data: pd.DataFrame) -> pd.DataFrame:
        styles = pd.DataFrame("", index=data.index, columns=data.columns)
        for rank, i in enumerate(data.index, start=1):
            if rank in podium:
                styles.loc[i, :] = podium[rank]
        return styles
        
    return df.style.apply(_pinta, axis=None).hide(axis="index")


def estilo_clasificacion_estadisticas(df: pd.DataFrame):
    maximo = "color: red; font-weight: bold;"
    
    def _pinta(data: pd.DataFrame) -> pd.DataFrame:
        styles = pd.DataFrame("", index=data.index, columns=data.columns)
        for col in ("PPD", "MPR"):
            if col in data.columns:
                serie = pd.to_numeric(data[col], errors="coerce")
                if serie.notna().any():
                    i = serie.idxmax()
                    styles.loc[i, col] = f"{styles.loc[i, col]} {maximo}"
        return styles
        
    return df.style.apply(_pinta, axis=None).format({"PPD": "{:.2f}", "MPR": "{:.2f}"}, na_rep="—").hide(axis="index")


def partidos_jugados_keys(df_partidos: pd.DataFrame, division: str) -> set:
    keys = set()
    if df_partidos.empty:
        return keys
        
    part = df_partidos[df_partidos["División"].astype(str).str.casefold() == division.casefold()]
    for _, p in part.iterrows():
        keys.add((p["jornada_n"], p["pareja"]))
        
    return keys


def calendario_vista(df_calendario, df_partidos, division: str) -> pd.DataFrame:
    cal = df_calendario[df_calendario["División"].str.casefold() == division.casefold()].copy()
    jugados = partidos_jugados_keys(df_partidos, division)
    filas = []
    
    for _, row in cal.iterrows():
        j1 = _norm_name(row["Jugador 1"])
        j2 = _norm_name(row["Jugador 2"])
        jornada_n = _jornada_num(row["Jornada"])
        bye = _is_bye(j1) or _is_bye(j2)
        
        if bye:
            descanso = j1 if _is_bye(j1) else j2
            activo = j2 if _is_bye(j1) else j1
            etiqueta = "Descansa" if _name_key(descanso) == "descansa" else "Libre"
            estado = f"🛌 {activo} {etiqueta}"
        elif (jornada_n, tuple(sorted((_name_key(j1), _name_key(j2))))) in jugados:
            estado = "✅ JUGADO"
        else:
            estado = "📅 PENDIENTE"
            
        filas.append({
            "Jornada": f"Jornada {jornada_n}" if jornada_n else row["Jornada"],
            "Enfrentamiento": f"{j1} vs {j2}",
            "Estado": estado,
        })
        
    return pd.DataFrame(filas)


def record_partido(df_medias_partido: pd.DataFrame, division: str, tipo: str, nombres: dict):
    if df_medias_partido.empty:
        return None, None
        
    sub = df_medias_partido[
        (df_medias_partido["División"].astype(str).str.casefold() == division.casefold()) & 
        (df_medias_partido["tipo"] == tipo)
    ]
    
    if sub.empty:
        return None, None
        
    fila = sub.loc[sub["media"].idxmax()]
    nombre = nombres.get(fila["jugador"], fila["jugador"])
    
    return nombre, round(float(fila["media"]), 2)


def stats_jugador(df_partidos, df_medias, jugador_k: str) -> dict:
    vacio = {"PJ": 0, "PG": 0, "PP": 0, "SF": 0, "SC": 0, "PPD": None, "MPR": None}
    
    if df_partidos.empty:
        return vacio
        
    for _, p in df_partidos.iterrows():
        a, b = p["pareja"]
        if jugador_k not in (a, b) or _is_bye(a) or _is_bye(b):
            continue
            
        rival = b if jugador_k == a else a
        vacio["PJ"] += 1
        vacio["SF"] += int(p["sets"].get(jugador_k, 0))
        vacio["SC"] += int(p["sets"].get(rival, 0))
        
        if p["ganador"] == jugador_k:
            vacio["PG"] += 1
        elif p["ganador"] == rival:
            vacio["PP"] += 1
            
    if not df_medias.empty:
        row = df_medias[df_medias["jugador"] == jugador_k]
        if not row.empty:
            vacio["PPD"] = row.iloc[0].get("PPD")
            vacio["MPR"] = row.iloc[0].get("MPR")
            
    return vacio


def evolucion_jugador(df_medias_partido: pd.DataFrame, jugador_k: str) -> pd.DataFrame:
    if df_medias_partido.empty:
        return pd.DataFrame(columns=["Jornada", "PPD", "MPR"])
        
    sub = df_medias_partido[df_medias_partido["jugador"] == jugador_k].copy()
    if sub.empty:
        return pd.DataFrame(columns=["Jornada", "PPD", "MPR"])
        
    pivot = (
        sub.pivot_table(index="jornada_n", columns="tipo", values="media", aggfunc="mean")
        .reset_index()
        .rename(columns={"jornada_n": "Jornada"})
        .sort_values("Jornada")
    )
    
    for col in ("PPD", "MPR"):
        if col not in pivot.columns:
            pivot[col] = pd.NA
            
    return pivot[["Jornada", "PPD", "MPR"]]


def grafica_linea(df: pd.DataFrame, y_col: str, titulo: str, domain: list, color: str):
    data = df[["Jornada", y_col]].dropna(subset=[y_col]).copy()
    if data.empty:
        st.info(f"Sin datos de {y_col} para este jugador.")
        return
        
    data["Jornada"] = data["Jornada"].astype(int)
    data[y_col] = data[y_col].astype(float).round(2)
    
    base = alt.Chart(data).encode(
        x=alt.X("Jornada:O", title="Jornada", axis=alt.Axis(labelAngle=0)),
        y=alt.Y(f"{y_col}:Q", title=y_col, scale=alt.Scale(domain=domain, clamp=True)),
        tooltip=[alt.Tooltip("Jornada:O", title="Jornada"), alt.Tooltip(f"{y_col}:Q", title=y_col, format=".2f")],
    )
    
    chart = (
        (base.mark_line(color=color, strokeWidth=3) + base.mark_point(color=color, size=90, filled=True))
        .properties(title=titulo, height=280)
        .interactive()
    )
    
    st.altair_chart(chart, use_container_width=True)


def pintar_records(df_medias_partido, division, nombres):
    c1, c2 = st.columns(2)
    n_ppd, v_ppd = record_partido(df_medias_partido, division, "PPD", nombres)
    n_mpr, v_mpr = record_partido(df_medias_partido, division, "MPR", nombres)
    
    with c1:
        st.metric("Mejor PPD en un Partido", f"{v_ppd:.2f}" if v_ppd is not None else "—", n_ppd or "Sin datos 501/701")
    with c2:
        st.metric("Mejor MPR en un Partido", f"{v_mpr:.2f}" if v_mpr is not None else "—", n_mpr or "Sin datos Cricket")


df_resultados, df_calendario = cargar_datos()

st.markdown('<div class="liga-kicker">Season dashboard</div>', unsafe_allow_html=True)
st.title("🏆 Liga Lluis Galvart GB")
st.caption("Clasificación general, estadísticas y ficha de cada jugador.")

if df_resultados is None or df_calendario is None:
    st.stop()

df_partidos = partidos_desde_resultados(df_resultados)
df_medias_partido = medias_por_partido(df_resultados)
df_medias = medias_jugadores(df_resultados)
nombres_all = mapa_nombres(df_calendario)

tab1, tab2, tab3 = st.tabs(["División 1", "División 2", "🎯 Ficha Individual"])

for tab, division in ((tab1, "Division 1"), (tab2, "Division 2")):
    with tab:
        nombres = jugadores_division(df_calendario, division)
        pintar_records(df_medias_partido, division, nombres)
        st.divider()

        st.subheader("Clasificación General")
        tabla_gen = clasificacion_general(df_calendario, df_partidos, df_resultados, division)
        if tabla_gen.empty:
            st.info("Aún no hay resultados para generar la clasificación general.")
        else:
            st.dataframe(estilo_clasificacion_general(tabla_gen), use_container_width=True, hide_index=True)

        st.divider()
        st.subheader("Estadísticas y Hazañas")
        tabla_est = clasificacion_estadisticas(df_calendario, df_resultados, df_medias, division)
        if tabla_est.empty:
            st.info("Aún no hay estadísticas para esta división.")
        else:
            st.dataframe(estilo_clasificacion_estadisticas(tabla_est), use_container_width=True, hide_index=True)

        st.divider()
        st.subheader("Calendario")
        cal = calendario_vista(df_calendario, df_partidos, division)
        if cal.empty:
            st.info("No hay calendario para esta división.")
        else:
            st.dataframe(cal, use_container_width=True, hide_index=True)

with tab3:
    jugadores = sorted(nombres_all.values(), key=lambda x: x.casefold())
    if not jugadores:
        st.info("No hay jugadores en el calendario.")
    else:
        elegido = st.selectbox("Elige jugador", jugadores, index=0)
        clave = _name_key(elegido)
        stats = stats_jugador(df_partidos, df_medias, clave)
        st.divider()

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("PJ", stats["PJ"])
        m2.metric("PG", stats["PG"])
        m3.metric("PP", stats["PP"])
        m4.metric("Sets +/−", stats["SF"] - stats["SC"])

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Sets a favor", stats["SF"])
        c2.metric("Sets en contra", stats["SC"])
        c3.metric("Media PPD", f"{stats['PPD']:.2f}" if pd.notna(stats["PPD"]) else "—")
        c4.metric("Media MPR", f"{stats['MPR']:.2f}" if pd.notna(stats["MPR"]) else "—")

        st.divider()
        st.subheader("Evolución de medias por jornada")
        evo = evolucion_jugador(df_medias_partido, clave)
        if evo.empty or (evo[["PPD", "MPR"]].isna().all().all()):
            st.info("Este jugador aún no tiene medias registradas.")
        else:
            grafica_linea(evo, "PPD", "Evolución PPD", [10, 45], AZUL_ELECTRICO)
            grafica_linea(evo, "MPR", "Evolución MPR", [1, 6], NARANJA_MPR)
            st.dataframe(
                evo.style.format({"PPD": "{:.2f}", "MPR": "{:.2f}"}, na_rep="—").hide(axis="index"),
                use_container_width=True,
                hide_index=True,
            )
