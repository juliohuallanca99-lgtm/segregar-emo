#!/usr/bin/env python3
"""Segrega PDFs de EMO hacia el proyecto 0171 (TSF 3+) usando el Excel de Fuerza Laboral.
Solo extrae los del TSF; lo demas NO se toca. Modo por defecto: simulacion (no copia nada).

Uso (Windows, PowerShell):
  python segregar_emo.py --lista "P0171_...xlsx" --emo "C:\\ruta\\EMO" --destino "C:\\ruta\\Proyectos Medicos" --meses "1. ENE"
  (agrega --copiar para copiar de verdad)
"""
import argparse, csv, difflib, hashlib, re, shutil, sys, unicodedata
from pathlib import Path
import openpyxl

PROYECTO = "0171 - Proyecto Movimiento de tierras TSF 3+"
UMBRAL = 0.90

def norm(s):
    s = unicodedata.normalize("NFD", str(s).upper())
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")  # Ñ -> N en ambos lados
    s = re.sub(r"[\s_\-.,]+", " ", s)
    return s.strip()

def nombre_desde_archivo(stem):
    s = re.sub(r"[\s_\-]*20\d\d.*$", "", stem, flags=re.I)  # corta desde el año (-2026, _2026_CON...)
    return norm(s)

def cargar_lista(xlsx):
    ws = openpyxl.load_workbook(xlsx, read_only=True, data_only=True).active
    filas = list(ws.iter_rows(values_only=True))
    enc = [str(c or "").strip().upper() for c in filas[0]]
    i_nom = next(i for i, c in enumerate(enc) if "APELLIDOS" in c)
    i_dni = next((i for i, c in enumerate(enc) if c == "DNI"), None)
    lista = {}
    for f in filas[1:]:
        if not f[i_nom]: continue
        k = norm(f[i_nom])
        if k in lista:
            sys.exit(f"ERROR: nombre repetido en la lista: {k}")
        lista[k] = f[i_dni] if i_dni is not None else ""
    return lista

def sha(p):
    h = hashlib.sha256()
    with open(p, "rb") as fh:
        for b in iter(lambda: fh.read(1 << 20), b""): h.update(b)
    return h.hexdigest()

def mes_dest(nombre):  # "1. ENE" -> "01 ENE"
    m = re.match(r"\s*(\d+)\.?\s*(.*)", nombre)
    return f"{int(m.group(1)):02d} {m.group(2).strip()}" if m else nombre

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--lista", required=True); ap.add_argument("--emo", required=True)
    ap.add_argument("--destino", required=True)
    ap.add_argument("--meses", nargs="*", help='ej: "1. ENE"  (vacio = todas)')
    ap.add_argument("--anio", default="2026"); ap.add_argument("--copiar", action="store_true")
    a = ap.parse_args()

    lista = cargar_lista(a.lista)
    emo, dest = Path(a.emo), Path(a.destino)
    carpetas = [c for c in sorted(emo.iterdir()) if c.is_dir() and (not a.meses or c.name in a.meses)]
    reporte, vistos = [], set()

    for c in carpetas:
        for pdf in sorted(c.glob("*.pdf")):
            clave = nombre_desde_archivo(pdf.stem)
            estado, cand = None, ""
            if len(clave.split()) < 2:
                estado = "NOMBRE_MALFORMADO"
            elif clave in lista:
                estado = "TSF_EXACTO"; cand = clave
            else:
                toks = sorted(clave.split())
                cs = [k for k in lista if sorted(k.split()) == toks]            # mismo set de palabras
                if not cs:
                    cs = [k for k in lista if difflib.SequenceMatcher(None, clave, k).ratio() >= UMBRAL
                          or (len(clave) >= 12 and (k.startswith(clave) or clave.startswith(k)))]  # truncados/typos
                if cs: estado = "TSF_REVISAR"; cand = " | ".join(cs)
                else:  estado = "NO_TSF"
            destino = ""
            if estado == "TSF_EXACTO":
                vistos.add(clave)
                carpeta = dest / PROYECTO / a.anio / mes_dest(c.name)
                final = carpeta / pdf.name
                if final.exists():
                    if sha(final) == sha(pdf): estado = "DUPLICADO_IDENTICO"; final = None
                    else:
                        n = 2
                        while (carpeta / f"{pdf.stem}_v{n}.pdf").exists(): n += 1
                        final = carpeta / f"{pdf.stem}_v{n}.pdf"; estado = "TSF_EXACTO_VERSION"
                if final:
                    destino = str(final)
                    if a.copiar:
                        carpeta.mkdir(parents=True, exist_ok=True)
                        shutil.copy2(pdf, final)
                        if final.stat().st_size != pdf.stat().st_size: estado = "ERROR_TAMANO"
            reporte.append([c.name, pdf.name, clave, estado, cand, destino])

    dest.mkdir(parents=True, exist_ok=True)
    nombre = "reporte_" + ("copia" if a.copiar else "simulacion") + ".csv"
    with open(dest / nombre, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh); w.writerow(["mes", "archivo", "clave", "estado", "candidato_lista", "destino"]); w.writerows(reporte)
    faltan = sorted(set(lista) - vistos)
    with open(dest / "lista_sin_pdf.csv", "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh); w.writerow(["nombre_en_lista", "dni"]); w.writerows([[k, lista[k]] for k in faltan])

    from collections import Counter
    print("Modo:", "COPIA" if a.copiar else "SIMULACION", "| PDFs revisados:", len(reporte))
    for k, v in Counter(r[3] for r in reporte).items(): print(f"  {k}: {v}")
    print(f"Personas de la lista SIN PDF en las carpetas revisadas: {len(faltan)} (ver lista_sin_pdf.csv)")
    print("Reportes en:", dest)

if __name__ == "__main__":
    main()
