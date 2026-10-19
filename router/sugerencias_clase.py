"""Sugerencias de clasificación (segmento, familia, subfamilia y grupo de equivalencia) para un producto,
a partir de los datos del catálogo:

1. Equivalencias: productos cuyo número de parte aparece en las equivalencias del producto (o al revés),
   o que tienen el mismo número de parte en otra marca. Si ya están clasificados se sugiere lo mismo y su grupo.
2. Texto: si no hay equivalentes clasificados, la familia/subfamilia más común entre productos con el mismo
   inicio de nombre ("RODAMIENTO DE BOLAS…") y el segmento según las palabras de la aplicación (Descripción Larga).
"""
import re
import time
from collections import Counter, defaultdict

from utils.db import get_connection

FALTANTES = {"", "por clasificar", "sin definir", "no aplica"}
MAX_POR_TOKEN = 25        # un número de parte que aparece en más productos que esto es genérico y no se usa
PALABRAS_VACIAS = {"DE", "DEL", "LA", "EL", "LOS", "LAS", "Y", "CON", "SIN", "PARA", "P", "C", "S", "A", "EN", "TIPO",
                   "AL", "POR", "O", "E", "MM", "PULG", "X"}
_MODELO = {"t": 0, "m": None}
DURACION = 600


def falta(valor):
    return (valor or "").strip().lower() in FALTANTES


def normal(texto):
    return re.sub(r"[^A-Z0-9]", "", (texto or "").upper())


def partes_de(numero_parte):
    """'BB1-4830A/6204-2RSC3' -> {'BB14830A', '62042RSC3', 'BB14830A62042RSC3'}"""
    p = {normal(numero_parte)} | {normal(x) for x in re.split(r"[/;,]", numero_parte or "")}
    return {x for x in p if len(x) >= 3}


def tokens_equivalencias(texto):
    return {t for t in (normal(x) for x in re.split(r"[/;,|\n]", texto or "")) if len(t) >= 4 and re.search(r"\d", t)}


def palabras(texto):
    return [w for w in re.findall(r"[A-ZÑÁÉÍÓÚ]{2,}", (texto or "").upper()) if w not in PALABRAS_VACIAS]


def olvidar():
    """Después de guardar clasificaciones, el modelo se reconstruye en la siguiente consulta."""
    _MODELO["t"] = 0


def modelo():
    if _MODELO["m"] is not None and time.time() - _MODELO["t"] < DURACION:
        return _MODELO["m"]
    conn = get_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT `Codigo Producto`, Nombre, `Numero Parte`, Marca, Equivalencias, `Descripcion Larga` "
                       "FROM catalogo_productos WHERE `Codigo Producto` IS NOT NULL")
        productos = {r[0]: {"nombre": r[1], "parte": r[2], "marca": r[3], "equiv": r[4], "aplic": r[5]} for r in cursor.fetchall()}
        cursor.execute("SELECT codigo, segmento, familia, subfamilia, grupo_equival FROM clase_prod")
        clase = {r[0]: {"segmento": r[1], "familia": r[2], "subfamilia": r[3], "grupo": (r[4] or "").strip() or None}
                 for r in cursor.fetchall()}
    finally:
        cursor.close()
        conn.close()

    por_parte, por_equiv = defaultdict(set), defaultdict(set)
    for cod, p in productos.items():
        p["partes"] = partes_de(p["parte"])
        p["tokens"] = tokens_equivalencias(p["equiv"])
        for x in p["partes"]:
            por_parte[x].add(cod)
        for t in p["tokens"]:
            por_equiv[t].add(cod)

    # Texto -> clasificación, aprendido de los productos ya clasificados
    por_nombre = defaultdict(Counter)       # prefijo del nombre -> (familia, subfamilia)
    seg_nombre = defaultdict(Counter)       # prefijo del nombre -> segmento
    seg_palabra = defaultdict(Counter)      # palabra de la aplicación -> segmento
    tamano_grupo = Counter(c["grupo"] for c in clase.values() if c["grupo"])
    for cod, c in clase.items():
        p = productos.get(cod)
        if not p or falta(c["familia"]):
            continue
        ws = palabras(p["nombre"])
        for n in (1, 2, 3):
            if len(ws) >= n:
                clave = " ".join(ws[:n])
                por_nombre[clave][(c["familia"], c["subfamilia"] or "")] += 1
                if not falta(c["segmento"]):
                    seg_nombre[clave][c["segmento"]] += 1
        if not falta(c["segmento"]):
            for w in set(palabras(p["aplic"])):
                if len(w) >= 3:
                    seg_palabra[w][c["segmento"]] += 1
    m = {"productos": productos, "clase": clase, "por_parte": por_parte, "por_equiv": por_equiv,
         "por_nombre": por_nombre, "seg_nombre": seg_nombre, "seg_palabra": seg_palabra, "tamano_grupo": tamano_grupo,
         "max_grupo": max((int(g[3:]) for g in tamano_grupo if re.fullmatch(r"EQ-\d+", g)), default=0)}
    _MODELO.update(t=time.time(), m=m)
    return m


def equivalentes(m, codigo):
    """Otros productos relacionados por número de parte / equivalencias, con cuántas coincidencias tienen."""
    p = m["productos"].get(codigo)
    if not p:
        return Counter()
    puntos = Counter()
    for t in p["tokens"] | p["partes"]:          # sus equivalencias (y su propio nº de parte en otras marcas)
        cods = m["por_parte"].get(t, set())
        if len(cods) <= MAX_POR_TOKEN:
            for c in cods:
                puntos[c] += 1
    for t in p["partes"]:                         # productos que lo listan como equivalente
        cods = m["por_equiv"].get(t, set())
        if len(cods) <= MAX_POR_TOKEN:
            for c in cods:
                puntos[c] += 1
    puntos.pop(codigo, None)
    # Un solo número en común puede ser casualidad (ARO 19519 vs. una correa que lista 19519):
    # en ese caso se exige que el nombre empiece con la misma palabra.
    primera = (palabras(p["nombre"]) or [""])[0]
    for c in [c for c, n in puntos.items() if n == 1]:
        if (palabras(m["productos"][c]["nombre"]) or [""])[0] != primera:
            del puntos[c]
    return puntos


def votar(contador):
    if not contador:
        return None, 0.0, 0
    total = sum(contador.values())
    valor, n = contador.most_common(1)[0]
    return valor, n / total, total


def sugerir(codigos):
    """{codigo: {segmento, familia, subfamilia, grupo, razon, confianza, equivalentes:[...], agrupar:[...]}}"""
    m = modelo()
    resultado = {}
    nuevos = {}   # grupo nuevo compartido por productos relacionados de la misma página
    for codigo in codigos:
        p = m["productos"].get(codigo)
        if not p:
            continue
        s = {"segmento": None, "familia": None, "subfamilia": None, "grupo": None, "razon": [], "confianza": None,
             "equivalentes": [], "agrupar": []}
        eq = equivalentes(m, codigo)
        clasificados = {c: n for c, n in eq.items() if c in m["clase"] and not falta(m["clase"][c]["familia"])}
        s["equivalentes"] = [{"codigo": c, "marca": m["productos"][c]["marca"], "parte": m["productos"][c]["parte"],
                              "grupo": (m["clase"].get(c) or {}).get("grupo")} for c, _ in eq.most_common(8)]

        # 1) Por equivalencias
        if clasificados:
            votos_fam, votos_seg = Counter(), Counter()
            for c, n in clasificados.items():
                votos_fam[(m["clase"][c]["familia"], m["clase"][c]["subfamilia"] or "")] += n
                if not falta(m["clase"][c]["segmento"]):
                    votos_seg[m["clase"][c]["segmento"]] += n
            fam, conf_f, _ = votar(votos_fam)
            seg, conf_s, _ = votar(votos_seg)
            s["familia"], s["subfamilia"] = fam
            s["segmento"] = seg
            s["confianza"] = "alta" if conf_f >= .75 else "media"
            mejor = max(clasificados, key=lambda c: clasificados[c])
            mp = m["productos"][mejor]
            s["razon"].append(f"equivalente de {mejor} ({mp['marca'] or 's/m'} {mp['parte'] or ''})".strip())

        # Grupo: el de los equivalentes (el más grande si hay varios); si ninguno tiene, uno nuevo para todos
        grupos = Counter()
        for c, n in eq.items():
            g = (m["clase"].get(c) or {}).get("grupo")
            if g:
                grupos[g] += n * 10 + m["tamano_grupo"][g]
        if grupos:
            s["grupo"] = grupos.most_common(1)[0][0]
            s["agrupar"] = [c for c in eq if not (m["clase"].get(c) or {}).get("grupo")]
            if len(grupos) > 1:
                s["razon"].append(f"los equivalentes están en {len(grupos)} grupos distintos")
        elif eq:
            clave = frozenset(eq) | {codigo}
            existente = next((g for k, g in nuevos.items() if k & clave), None)
            if not existente:
                m["max_grupo"] += 1
                existente = f"EQ-{m['max_grupo']:05d}"
            nuevos[clave] = existente
            s["grupo"] = existente
            s["agrupar"] = sorted(eq)
            s["grupo_nuevo"] = True

        # 2) Por texto (nombre y aplicación) para lo que falte
        if not s["familia"]:
            ws = palabras(p["nombre"])
            for n in (3, 2, 1):
                if len(ws) >= n:
                    valor, conf, total = votar(m["por_nombre"].get(" ".join(ws[:n])))
                    if valor and total >= 3 and conf >= (.8 if n == 1 else .6):   # una sola palabra es poca evidencia
                        s["familia"], s["subfamilia"] = valor
                        s["confianza"] = "media" if conf >= .85 and total >= 10 else "baja"
                        s["razon"].append(f"{conf:.0%} de {total} productos \"{' '.join(ws[:n]).lower()}…\" son {valor[0]}")
                        break
        if not s["segmento"]:
            votos = Counter()
            for w in set(palabras(p["aplic"])):
                dist = m["seg_palabra"].get(w)
                if dist and 2 <= sum(dist.values()) <= 3000:
                    total = sum(dist.values())
                    for seg, n in dist.items():
                        votos[seg] += n / total
            valor, conf, _ = votar(votos)
            if valor and conf >= .5:
                s["segmento"] = valor
                s["razon"].append(f"segmento por la aplicación ({conf:.0%})")
            else:
                ws = palabras(p["nombre"])
                for n in (3, 2, 1):
                    if len(ws) >= n:
                        valor, conf, total = votar(m["seg_nombre"].get(" ".join(ws[:n])))
                        if valor and total >= 3 and conf >= .6:
                            s["segmento"] = valor
                            s["razon"].append(f"segmento por el nombre ({conf:.0%})")
                            break
        if s["subfamilia"] == "":
            s["subfamilia"] = None
        if any(s[k] for k in ("segmento", "familia", "subfamilia", "grupo")):
            s["razon"] = "; ".join(s["razon"])
            resultado[codigo] = s
    return resultado
