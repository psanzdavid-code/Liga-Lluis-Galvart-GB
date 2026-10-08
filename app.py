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
      .block-container h1, .block-container h2, .block-container h3 { letter-spacing: 0.04em; color: #111 !important; }
      [data-testid="stDialog"] h2 { color: #f0f0f0 !important; }
      [data-testid="stDataFrame"],
      [data-testid="stDataFrameResizable"] {
        background: rgba(255, 255, 255, 0.96) !important;
        border-radius: 12px;
      }
      [data-testid="stDataFrame"] [data-testid="stElementToolbar"] {
        display: none !important;
      }
      [data-testid="stDataFrame"] [role="columnheader"] {
        pointer-events: none !important;
        justify-content: center !important;
        text-align: center !important;
      }
      [data-testid="stDataFrame"] .gdg-resizer,
      [data-testid="stDataFrame"] [class*="resizer"] {
        display: none !important;
        pointer-events: none !important;
      }
      div[data-testid="stMetric"] {
        background: linear-gradient(180deg, #1a1a1a 0%, #111 100%);
        border: 1px solid #c9a22755;
        border-radius: 16px;
        padding: 12px 16px;
      }
      div[data-testid="stMetric"] label { color: #d4af37 !important; font-weight: 700 !important; }
      div[data-testid="stMetric"] [data-testid="stMetricValue"] { color: #fff8dc !important; font-size: 1.6rem !important; }
      div[data-testid="stMetricDelta"] svg { display: none !important; }
      div[data-testid="stMetricDelta"] { color: #f3e5ab !important; }
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
    codigo = _juego_codigo(value)
    if codigo in {"501", "701"}:
        return "PPD"
    if codigo == "CRK":
        return "MPR"
    return ""


def _juego_codigo(value) -> str:
    text = _norm_name(value).casefold()
    if "501" in text:
        return "501"
    if "701" in text:
        return "701"
    if "cricket" in text or "cr." in text or text == "cr" or "standard cr" in text:
        return "CRK"
    return ""


def _col_x01(division: str) -> str:
    return "501" if "2" in str(division).casefold() else "701"


def _nombre_visible(nombres: dict, clave) -> str:
    return nombres.get(clave, _norm_name(clave))


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


@st.cache_data(ttl=60)
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
                "Set": row.get("Set"),
                "pareja": pareja,
                "jugador": jugador,
                "rival": p2 if jugador == p1 else p1,
                "tipo": tipo,
                "media": float(media),
            })
            
    return pd.DataFrame(registros)


def medias_por_partido(df_resultados: pd.DataFrame) -> pd.DataFrame:
    legs = _legs_por_jugador(df_resultados)
    if legs.empty:
        return pd.DataFrame(columns=["División", "jornada_n", "pareja", "jugador", "rival", "tipo", "media"])
        
    return legs.groupby(
        ["División", "jornada_n", "pareja", "jugador", "rival", "tipo"],
        as_index=False,
    )["media"].mean()


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
            
        codigo = _juego_codigo(row.get("Juego"))
        for p in (p1, p2):
            if p not in stats:
                stats[p] = {"LF": 0, "LC": 0, "Legs": 0, "501": 0, "701": 0, "CRK": 0}

            stats[p]["Legs"] += 1
            if codigo in stats[p]:
                stats[p][codigo] += 1

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
        k: {"Jugador": v, "Forma": "", "PJ": 0, "PG": 0, "PP": 0, "SF": 0, "SC": 0, "LF": 0, "LC": 0} 
        for k, v in nombres.items()
    }

    if not df_partidos.empty:
        part = df_partidos[df_partidos["División"].astype(str).str.casefold() == division.casefold()].copy()
        part = part.sort_values("jornada_n")
        
        for k in tabla:
            match_k = part[part["pareja"].apply(lambda p: k in p)]
            last_3 = match_k.tail(3)
            forma = []
            for _, row in last_3.iterrows():
                if row["ganador"] == k:
                    forma.append("✅")
                else:
                    forma.append("❌")
            tabla[k]["Forma"] = "".join(forma)
            
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
    
    return out[["Pos", "Jugador", "Forma", "PJ", "PG", "PP", "SF", "SC", "LF", "LC"]].reset_index(drop=True)


def clasificacion_estadisticas(df_calendario, df_resultados, df_medias, division: str) -> pd.DataFrame:
    nombres = jugadores_division(df_calendario, division)
    col_x01 = _col_x01(division)
    tabla = {
        k: {"Jugador": v, **{h: 0 for h in HAZANAS}, "Legs": 0, col_x01: 0, "CRK": 0}
        for k, v in nombres.items()
    }

    legs_stats = contar_legs(df_resultados, division)
    for k, stat in legs_stats.items():
        if k in tabla:
            tabla[k]["Legs"] = stat["Legs"]
            tabla[k][col_x01] = stat.get(col_x01, 0)
            tabla[k]["CRK"] = stat.get("CRK", 0)

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

    out["COMB"] = (out["MPR"] * 10) + out["PPD"]
    out["COMB"] = pd.to_numeric(out["COMB"], errors="coerce").round(2)
    out = out.sort_values(["COMB"], ascending=[False], kind="mergesort", na_position="last")

    return out[["Jugador", *HAZANAS, "Legs", col_x01, "CRK", "PPD", "MPR", "COMB"]].reset_index(drop=True)


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
        for col in ("PPD", "MPR", "COMB"):
            if col in data.columns:
                serie = pd.to_numeric(data[col], errors="coerce")
                if serie.notna().any():
                    i = serie.idxmax()
                    styles.loc[i, col] = f"{styles.loc[i, col]} {maximo}"
        return styles

    return (
        df.style.apply(_pinta, axis=None)
        .format({"PPD": "{:.2f}", "MPR": "{:.2f}", "COMB": "{:.2f}"}, na_rep="—")
        .hide(axis="index")
    )


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
    marcadores = {}
    if not df_partidos.empty:
        part = df_partidos[df_partidos["División"].astype(str).str.casefold() == division.casefold()]
        for _, p in part.iterrows():
            marcadores[(p["jornada_n"], p["pareja"])] = p["sets"]
    filas = []

    for _, row in cal.iterrows():
        j1 = _norm_name(row["Jugador 1"])
        j2 = _norm_name(row["Jugador 2"])
        jornada_n = _jornada_num(row["Jornada"])
        bye = _is_bye(j1) or _is_bye(j2)
        clave = (jornada_n, tuple(sorted((_name_key(j1), _name_key(j2)))))
        resultado = "-"

        if bye:
            descanso = j1 if _is_bye(j1) else j2
            activo = j2 if _is_bye(j1) else j1
            etiqueta = "Descansa" if _name_key(descanso) == "descansa" else "Libre"
            estado = f"🛌 {activo} {etiqueta}"
        elif clave in jugados:
            estado = "✅ JUGADO"
            sets = marcadores.get(clave) or {}
            k1, k2 = _name_key(j1), _name_key(j2)
            resultado = f"{int(sets.get(k1, 0))}-{int(sets.get(k2, 0))}"
        else:
            estado = "📅 PENDIENTE"

        filas.append({
            "Jornada": f"Jornada {jornada_n}" if jornada_n else row["Jornada"],
            "Enfrentamiento": f"{j1} vs {j2}",
            "Resultado": resultado,
            "Estado": estado,
        })

    return pd.DataFrame(filas)


def _medias_agrupadas(legs: pd.DataFrame, group_cols: list[str]) -> pd.DataFrame:
    if legs.empty:
        return pd.DataFrame()

    g = legs.groupby([*group_cols, "tipo"], as_index=False)["media"].mean()
    wide = g.pivot_table(index=group_cols, columns="tipo", values="media", aggfunc="mean").reset_index()
    for col in ("PPD", "MPR"):
        if col not in wide.columns:
            wide[col] = pd.NA
    wide["COMB"] = (pd.to_numeric(wide["MPR"], errors="coerce") * 10) + pd.to_numeric(wide["PPD"], errors="coerce")
    return wide


def _mejor_comb(wide: pd.DataFrame, nombres: dict):
    if wide.empty or "COMB" not in wide.columns:
        return None
    sub = wide.dropna(subset=["COMB"])
    if sub.empty:
        return None
    fila = sub.loc[sub["COMB"].idxmax()]
    return {
        "jugador": _nombre_visible(nombres, fila["jugador"]),
        "rival": _nombre_visible(nombres, fila["rival"]),
        "COMB": round(float(fila["COMB"]), 2),
        "PPD": round(float(fila["PPD"]), 2),
        "MPR": round(float(fila["MPR"]), 2),
    }


def _mejor_leg(legs: pd.DataFrame, tipo: str, nombres: dict):
    if legs.empty:
        return None
    sub = legs[legs["tipo"] == tipo]
    if sub.empty:
        return None
    fila = sub.loc[sub["media"].idxmax()]
    return {
        "jugador": _nombre_visible(nombres, fila["jugador"]),
        "rival": _nombre_visible(nombres, fila["rival"]),
        "valor": round(float(fila["media"]), 2),
    }


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


def evolucion_jugador(df_medias_partido: pd.DataFrame, jugador_k: str, nombres: dict) -> pd.DataFrame:
    if df_medias_partido.empty:
        return pd.DataFrame(columns=["Jornada", "Rival", "PPD", "MPR", "COMB"])
        
    sub = df_medias_partido[df_medias_partido["jugador"] == jugador_k].copy()
    if sub.empty:
        return pd.DataFrame(columns=["Jornada", "Rival", "PPD", "MPR", "COMB"])
        
    pivot = (
        sub.pivot_table(index=["jornada_n", "rival"], columns="tipo", values="media", aggfunc="mean")
        .reset_index()
        .rename(columns={"jornada_n": "Jornada", "rival": "Rival"})
        .sort_values("Jornada")
    )
    
    pivot["Rival"] = pivot["Rival"].apply(lambda r: _nombre_visible(nombres, r))
    
    for col in ("PPD", "MPR"):
        if col not in pivot.columns:
            pivot[col] = pd.NA
            
    pivot["COMB"] = (pd.to_numeric(pivot["MPR"], errors="coerce") * 10) + pd.to_numeric(pivot["PPD"], errors="coerce")
    pivot["COMB"] = pivot["COMB"].round(2)
    
    return pivot[["Jornada", "Rival", "PPD", "MPR", "COMB"]]


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
    )
    
    st.altair_chart(chart, width="stretch")


def _columnas_fijas(data) -> dict:
    frame = data.data if hasattr(data, "data") else data
    cfg = {}
    for col in frame.columns:
        serie = frame[col]
        if col == "Pos":
            cfg[col] = st.column_config.NumberColumn(
                col, width=44, format="%d", alignment="center", disabled=True, pinned=True
            )
        elif col == "Jugador":
            cfg[col] = st.column_config.TextColumn(
                col, width=90, alignment="left", disabled=True
            )
        elif col == "Forma":
            cfg[col] = st.column_config.TextColumn(
                col, width=70, alignment="center", disabled=True
            )
        elif col == "Rival":
            cfg[col] = st.column_config.TextColumn(
                col, width=118, alignment="left", disabled=True
            )
        elif col == "Enfrentamiento":
            cfg[col] = st.column_config.TextColumn(col, width=168, alignment="center", disabled=True)
        elif col == "Estado":
            cfg[col] = st.column_config.TextColumn(col, width=148, alignment="center", disabled=True)
        elif col in {"Resultado", "Jornada"} and not pd.api.types.is_numeric_dtype(serie):
            cfg[col] = st.column_config.TextColumn(col, width=78, alignment="center", disabled=True)
        elif col in {"PPD", "MPR", "COMB"}:
            cfg[col] = st.column_config.NumberColumn(
                col, width=58, format="%.2f", alignment="center", disabled=True
            )
        else:
            cfg[col] = st.column_config.NumberColumn(
                col, width=46, format="%d", alignment="center", disabled=True
            )
    return cfg


def mostrar_tabla(data, key: str):
    frame = data.data if hasattr(data, "data") else data
    st.dataframe(
        data,
        width="stretch",
        height="content",
        hide_index=True,
        column_order=list(frame.columns),
        column_config=_columnas_fijas(data),
        key=key,
        row_height=30,
        selection_mode="single-column",
    )


def pintar_records(df_resultados, division, nombres):
    legs = _legs_por_jugador(df_resultados)
    if not legs.empty:
        legs = legs[legs["División"].astype(str).str.casefold() == division.casefold()]

    rec_partido = _mejor_comb(
        _medias_agrupadas(legs, ["División", "jornada_n", "pareja", "jugador", "rival"]),
        nombres,
    )
    rec_set = _mejor_comb(
        _medias_agrupadas(legs, ["División", "jornada_n", "pareja", "Set", "jugador", "rival"]),
        nombres,
    )
    rec_ppd = _mejor_leg(legs, "PPD", nombres)
    rec_mpr = _mejor_leg(legs, "MPR", nombres)

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        if rec_partido:
            st.metric(
                "Mejor COMB (Partido)",
                f"{rec_partido['COMB']:.2f} (PPD {rec_partido['PPD']:.2f} - MPR {rec_partido['MPR']:.2f})",
                f"{rec_partido['jugador']} vs {rec_partido['rival']}",
                delta_color="off",
            )
        else:
            st.metric("Mejor COMB (Partido)", "—", "Sin datos", delta_color="off")
    with c2:
        if rec_set:
            st.metric(
                "Mejor COMB (Set)",
                f"{rec_set['COMB']:.2f} (PPD {rec_set['PPD']:.2f} - MPR {rec_set['MPR']:.2f})",
                f"{rec_set['jugador']} vs {rec_set['rival']}",
                delta_color="off",
            )
        else:
            st.metric("Mejor COMB (Set)", "—", "Sin datos", delta_color="off")
    with c3:
        if rec_ppd:
            st.metric(
                "Mejor PPD (Leg)",
                f"{rec_ppd['valor']:.2f}",
                f"{rec_ppd['jugador']} vs {rec_ppd['rival']}",
                delta_color="off",
            )
        else:
            st.metric("Mejor PPD (Leg)", "—", "Sin datos 501/701", delta_color="off")
    with c4:
        if rec_mpr:
            st.metric(
                "Mejor MPR (Leg)",
                f"{rec_mpr['valor']:.2f}",
                f"{rec_mpr['jugador']} vs {rec_mpr['rival']}",
                delta_color="off",
            )
        else:
            st.metric("Mejor MPR (Leg)", "—", "Sin datos Cricket", delta_color="off")


def pintar_records_individual(df_resultados, jugador_k, nombres):
    legs = _legs_por_jugador(df_resultados)
    if not legs.empty:
        legs = legs[legs["jugador"] == jugador_k]

    rec_partido = _mejor_comb(
        _medias_agrupadas(legs, ["División", "jornada_n", "pareja", "jugador", "rival"]),
        nombres,
    )
    rec_set = _mejor_comb(
        _medias_agrupadas(legs, ["División", "jornada_n", "pareja", "Set", "jugador", "rival"]),
        nombres,
    )
    rec_ppd = _mejor_leg(legs, "PPD", nombres)
    rec_mpr = _mejor_leg(legs, "MPR", nombres)

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        if rec_partido:
            st.metric(
                "Mejor COMB (Partido)",
                f"{rec_partido['COMB']:.2f} (PPD {rec_partido['PPD']:.2f} - MPR {rec_partido['MPR']:.2f})",
                f"vs {rec_partido['rival']}",
                delta_color="off",
            )
        else:
            st.metric("Mejor COMB (Partido)", "—", "Sin datos", delta_color="off")
    with c2:
        if rec_set:
            st.metric(
                "Mejor COMB (Set)",
                f"{rec_set['COMB']:.2f} (PPD {rec_set['PPD']:.2f} - MPR {rec_set['MPR']:.2f})",
                f"vs {rec_set['rival']}",
                delta_color="off",
            )
        else:
            st.metric("Mejor COMB (Set)", "—", "Sin datos", delta_color="off")
    with c3:
        if rec_ppd:
            st.metric(
                "Mejor PPD (Leg)",
                f"{rec_ppd['valor']:.2f}",
                f"vs {rec_ppd['rival']}",
                delta_color="off",
            )
        else:
            st.metric("Mejor PPD (Leg)", "—", "Sin datos", delta_color="off")
    with c4:
        if rec_mpr:
            st.metric(
                "Mejor MPR (Leg)",
                f"{rec_mpr['valor']:.2f}",
                f"vs {rec_mpr['rival']}",
                delta_color="off",
            )
        else:
            st.metric("Mejor MPR (Leg)", "—", "Sin datos", delta_color="off")


@st.dialog("📋 Acta del Partido", width="large")
def acta_dialog(j1_name, j2_name, jornada_n, division, df_resultados):
    k1, k2 = _name_key(j1_name), _name_key(j2_name)
    df_partido = df_resultados[
        (df_resultados["División"].astype(str).str.casefold() == division.casefold()) &
        (df_resultados["Jornada"].map(_jornada_num) == jornada_n)
    ]
    
    mask1 = (df_partido["Jugador 1"].map(_name_key) == k1) & (df_partido["Jugador 2"].map(_name_key) == k2)
    mask2 = (df_partido["Jugador 1"].map(_name_key) == k2) & (df_partido["Jugador 2"].map(_name_key) == k1)
    df_partido = df_partido[mask1 | mask2]
    
    if df_partido.empty:
        st.warning("No se encontraron los datos detallados (legs) de este partido.")
        return

    match_data = {1: {}, 2: {}, 3: {}}
    col_x01 = _col_x01(division)
    
    haz_totales = {k1: {h: 0 for h in HAZANAS}, k2: {h: 0 for h in HAZANAS}}
    
    # Nuevas variables para recopilar los PPD y MPR de este partido concreto
    sum_x01 = {k1: [], k2: []}
    sum_ck = {k1: [], k2: []}
    
    for _, row in df_partido.iterrows():
        s = _to_int(row.get("Set"))
        l = _to_int(row.get("Leg"))
        if s not in match_data:
            match_data[s] = {}
            
        juego_str = _juego_codigo(row.get("Juego"))
        if juego_str in {"501", "701"}: label_juego = col_x01
        elif juego_str == "CRK": label_juego = "CK"
        else: label_juego = ""
        
        ganador_k = _name_key(row.get("Ganador Leg"))
        
        r_j1_k = _name_key(row.get("Jugador 1"))
        r_j2_k = _name_key(row.get("Jugador 2"))
        
        val_k1 = row.get("Media J1") if r_j1_k == k1 else (row.get("Media J2") if r_j2_k == k1 else pd.NA)
        val_k2 = row.get("Media J2") if r_j2_k == k2 else (row.get("Media J1") if r_j1_k == k2 else pd.NA)
        
        match_data[s][l] = {
            "juego": label_juego,
            "ganador": ganador_k,
            k1: val_k1,
            k2: val_k2
        }
        
        # Guardar datos para hacer la media del partido
        if label_juego == col_x01:
            if pd.notna(val_k1): sum_x01[k1].append(val_k1)
            if pd.notna(val_k2): sum_x01[k2].append(val_k2)
        elif label_juego == "CK":
            if pd.notna(val_k1): sum_ck[k1].append(val_k1)
            if pd.notna(val_k2): sum_ck[k2].append(val_k2)
        
        for h in HAZANAS:
            haz_totales[r_j1_k][h] += _to_int(row.get(f"J1 {h}"))
            haz_totales[r_j2_k][h] += _to_int(row.get(f"J2 {h}"))

    # Calcular las medias finales del partido
    stats_partido = {}
    for p in [k1, k2]:
        m_x01 = sum(sum_x01[p])/len(sum_x01[p]) if sum_x01[p] else 0
        m_ck = sum(sum_ck[p])/len(sum_ck[p]) if sum_ck[p] else 0
        comb = (m_ck * 10) + m_x01
        stats_partido[p] = {"x01": m_x01, "ck": m_ck, "comb": comb}
            
    sets_to_show = [1, 2, 3] if match_data.get(3) else [1, 2]
    max_leg_per_set = {s: max(match_data[s].keys()) if match_data.get(s) else 0 for s in sets_to_show}
    
    css = """
    <style>
    .acta-table { width: 100%; border-collapse: collapse; font-family: 'Segoe UI', sans-serif; text-align: center; font-size: 0.95rem; margin-top: 10px; }
    .acta-table th, .acta-table td { border: 2px solid #222; padding: 8px 4px; }
    .acta-header { background-color: #1a1a1a; color: #d4af37; font-weight: 800; text-transform: uppercase; letter-spacing: 1px; }
    .acta-set-header { background-color: #333; color: white; font-weight: 700; letter-spacing: 1.5px; }
    .acta-stats-header { background-color: #444; color: #fff; font-weight: 800; font-size: 0.85em; }
    .acta-leg-header { background-color: #555; color: #ffd700; font-size: 0.8em; font-weight: 800; line-height: 1.2; }
    .cell-win { background-color: #D32F2F; color: white; font-weight: 900; }
    .cell-loss { background-color: #d9d9d9; color: #555; font-weight: 700; }
    .cell-empty { background-color: #f0f0f0; }
    .player-name { background-color: #111; color: white; text-align: center; font-weight: 800; font-size: 1.25rem; }
    .hazanas-cell { background-color: #222; color: #d4af37; font-weight: 700; font-size: 1rem; text-align: center; }
    .stats-cell { background-color: #2b2b2b; color: #fff; font-weight: 800; font-size: 1.05rem; text-align: center; border-left: 2px solid #111; border-right: 2px solid #111; }
    .starter-dot { font-size: 1.4em; line-height: 0; vertical-align: middle; margin-right: 4px; }
    </style>
    """
    
    html = [css, "<div style='overflow-x: auto;'><table class='acta-table'>"]
    
    html.append("<tr><th rowspan='2' class='acta-header'>JUGADOR</th>")
    for s in sets_to_show:
        max_l = max_leg_per_set[s]
        if max_l > 0:
            html.append(f"<th colspan='{max_l}' class='acta-set-header'>SET {s}</th>")
            
    # Añadimos la cabecera general de las MEDIAS
    html.append("<th colspan='3' class='acta-set-header' style='background-color:#2b2b2b;'>MEDIAS</th>")
    html.append(f"<th colspan='{len(HAZANAS)}' class='acta-set-header'>HAZAÑAS</th></tr>")
    
    html.append("<tr>")
    for s in sets_to_show:
        max_l = max_leg_per_set[s]
        for l in range(1, max_l + 1):
            juego_label = match_data.get(s, {}).get(l, {}).get("juego", "")
            if not juego_label:
                html.append(f"<th class='acta-leg-header'>—<br>{l}</th>")
            else:
                html.append(f"<th class='acta-leg-header'>{juego_label}<br>{l}</th>")
                
    # Añadimos las sub-cabeceras de cada media
    html.append(f"<th class='acta-stats-header'>COMB</th>")
    html.append(f"<th class='acta-stats-header'>{col_x01}</th>")
    html.append(f"<th class='acta-stats-header'>CK</th>")

    for h in HAZANAS:
        h_label = h.replace(" ", "<br>")
        html.append(f"<th class='acta-leg-header'>{h_label}</th>")
    html.append("</tr>")
    
    for p_key, p_name in [(k1, j1_name), (k2, j2_name)]:
        html.append("<tr>")
        html.append(f"<td class='player-name'>{p_name}</td>")
        for s in sets_to_show:
            max_l = max_leg_per_set[s]
            for l in range(1, max_l + 1):
                leg_data = match_data.get(s, {}).get(l)
                
                is_starter = False
                if s in [1, 3]:
                    if l in [1, 3, 5] and p_key == k1: is_starter = True
                    if l in [2, 4] and p_key == k2: is_starter = True
                elif s == 2:
                    if l in [1, 3, 5] and p_key == k2: is_starter = True
                    if l in [2, 4] and p_key == k1: is_starter = True
                    
                dot = "<span class='starter-dot'>•</span>" if is_starter else ""
                
                if not leg_data:
                    html.append("<td class='cell-loss'></td>")
                else:
                    val = leg_data.get(p_key)
                    val_str = f"{val:.2f}" if pd.notna(val) else "-"
                    clase = "cell-win" if leg_data.get("ganador") == p_key else "cell-loss"
                    html.append(f"<td class='{clase}'>{dot}{val_str}</td>")
        
        # Inyectamos las celdas con los valores de las medias del partido
        html.append(f"<td class='stats-cell' style='color:#f3e5ab;'>{stats_partido[p_key]['comb']:.2f}</td>")
        html.append(f"<td class='stats-cell' style='color:#247CFF;'>{stats_partido[p_key]['x01']:.2f}</td>")
        html.append(f"<td class='stats-cell' style='color:#FF6B00;'>{stats_partido[p_key]['ck']:.2f}</td>")

        for h in HAZANAS:
            count = haz_totales[p_key][h]
            html.append(f"<td class='hazanas-cell'>{count}</td>")
        html.append("</tr>")
        
    html.append("</table></div>")
    
    st.markdown("".join(html), unsafe_allow_html=True)


def mostrar_calendario_interactivo(df_cal, key, division, df_resultados):
    event = st.dataframe(
        df_cal,
        width="stretch",
        height="content",
        hide_index=True,
        column_order=list(df_cal.columns),
        column_config=_columnas_fijas(df_cal),
        key=key,
        row_height=30,
        on_select="rerun",
        selection_mode="single-row",
    )
    
    if getattr(event, "selection", None) and getattr(event.selection, "rows", None):
        if len(event.selection.rows) > 0:
            idx = event.selection.rows[0]
            row = df_cal.iloc[idx]
            if row["Estado"] == "✅ JUGADO":
                enf = row["Enfrentamiento"]
                if " vs " in enf:
                    j1, j2 = enf.split(" vs ")
                    jornada_n = _jornada_num(row["Jornada"])
                    acta_dialog(j1, j2, jornada_n, division, df_resultados)


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
        pintar_records(df_resultados, division, nombres)
        st.divider()

        st.subheader("Clasificación General")
        tabla_gen = clasificacion_general(df_calendario, df_partidos, df_resultados, division)
        if tabla_gen.empty:
            st.info("Aún no hay resultados para generar la clasificación general.")
        else:
            mostrar_tabla(estilo_clasificacion_general(tabla_gen), f"gen-{division}")

        st.divider()
        st.subheader("Estadísticas y Hazañas")
        tabla_est = clasificacion_estadisticas(df_calendario, df_resultados, df_medias, division)
        if tabla_est.empty:
            st.info("Aún no hay estadísticas para esta división.")
        else:
            mostrar_tabla(estilo_clasificacion_estadisticas(tabla_est), f"est-{division}")

        st.divider()
        st.subheader("Calendario (Haz clic en un partido jugado para ver el acta)")
        cal = calendario_vista(df_calendario, df_partidos, division)
        if cal.empty:
            st.info("No hay calendario para esta división.")
        else:
            mostrar_calendario_interactivo(cal, f"cal-{division}", division, df_resultados)

with tab3:
    jugadores = sorted(nombres_all.values(), key=lambda x: x.casefold())
    if not jugadores:
        st.info("No hay jugadores en el calendario.")
    else:
        elegido = st.selectbox("Elige jugador", jugadores, index=0)
        clave = _name_key(elegido)
        
        div_serie = df_calendario[(df_calendario["Jugador 1"].map(_name_key) == clave) | (df_calendario["Jugador 2"].map(_name_key) == clave)]["División"]
        jugador_division = div_serie.iloc[0] if not div_serie.empty else "1"
        is_div_1 = "1" in str(jugador_division)
        
        st.divider()
        
        pintar_records_individual(df_resultados, clave, nombres_all)

        st.divider()
        st.subheader("Evolución de medias por jornada")
        evo = evolucion_jugador(df_medias_partido, clave, nombres_all)
        if evo.empty or (evo[["PPD", "MPR"]].isna().all().all()):
            st.info("Este jugador aún no tiene medias registradas.")
        else:
            if is_div_1:
                dom_ppd = [25, 46]
                dom_mpr = [2.40, 5.40]
            else:
                dom_ppd = [15, 35]
                dom_mpr = [1.30, 4.30]
                
            grafica_linea(evo, "PPD", "Evolución PPD", dom_ppd, AZUL_ELECTRICO)
            grafica_linea(evo, "MPR", "Evolución MPR", dom_mpr, NARANJA_MPR)
            mostrar_tabla(
                evo.style.format({"PPD": "{:.2f}", "MPR": "{:.2f}", "COMB": "{:.2f}"}, na_rep="—").hide(axis="index"),
                "evo-jugador",
            )
