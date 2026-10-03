# -*- coding: utf-8 -*-
"""
Informe de BASE HISTORICA.xlsx (encuesta laboral bimestral de Catch Consulting).

Genera información útil de la base, estructurada por importancia, usando prints.
Uso:  python informe_base_historica.py

No requiere entornos virtuales: usa el Python del sistema con pandas y openpyxl.
"""
import sys

import pandas as pd

ARCHIVO = "BASE HISTORICA.xlsx"

# Frases (case-insensitive) para localizar cada columna clave dentro de la hoja.
METRICAS = {
    "empleados_operativos":       "número de empleados actual de personal operativo",
    "empleados_administrativos":  "número de empleados personal administrativo",
    "otro_personal":              "otro tipo de personal que labora",
    "salario_diario_inicial":     "salario diario inicial",
    "salario_diario_inter_bajo":  "salario diario intermedio - bajo",
    "salario_diario_inter_alto":  "salario diario intermedio - alto",
    "salario_diario_maximo":      "salario diario máximo",
    "salario_min_mensual":        "salario mínimo mensual",
    "salario_prom_mensual":       "salario promedio mensual",
    "salario_max_mensual":        "salario máximo mensual",
    "contrataciones":             "total de contrataciones",
    "bajas":                      "número de bajas que tuvo",
    "faltas":                     "total de faltas que tuvo",
    "incapacidades":              "total de casos de incapacidad",
    "capacitacion":               "total de horas de capacitación",
}

CATEGORICAS = {
    "tiene_sindicato":     "cuenta con alguna confederación sindical",
    "cct_status":          "estatus actual de su negociación",
    "relacion_sindicato":  "calificaría la relación actual",
}


def configurar_consola():
    """Imprimir acentos correctamente en la terminal de Windows."""
    for flujo in (sys.stdout, sys.stderr):
        try:
            flujo.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def find_col(df, frase):
    """Primera columna cuyo nombre contiene la frase (sin distinguir mayúsculas)."""
    hits = [c for c in df.columns if frase in str(c).lower()]
    return hits[0] if hits else None


def limpiar(df):
    """Elimina la fila de sub-encabezados de Qualtrics y filas vacías."""
    compania = df["Compañía"].astype(str)
    mascara_encabezado = compania.str.contains(
        "Open-Ended|Response", case=False, na=False
    )
    limpio = df[~mascara_encabezado].copy()
    # Elimina filas sin empresa (renglones vacíos al final de la hoja).
    limpio = limpio[limpio["Compañía"].notna()]
    limpio["Compañía"] = limpio["Compañía"].astype(str).str.strip()
    return limpio[limpio["Compañía"] != ""]


def to_num(serie):
    """Convierte a número: quita comas, trata '-'/vacío como NaN."""
    s = serie.astype(str).str.strip()
    s = s.str.replace(",", "", regex=False)
    s = s.replace(["", "-", "nan", "None", "N/A", "No Aplica", "Desconocido"], None)
    return pd.to_numeric(s, errors="coerce")


def fmt_num(x, dec=0):
    if pd.isna(x):
        return "-"
    return f"{x:,.{dec}f}"


def cargar_periodos():
    xl = pd.ExcelFile(ARCHIVO)
    periodos = {}
    for hoja in xl.sheet_names:
        df = pd.read_excel(xl, sheet_name=hoja)
        periodos[hoja] = limpiar(df)
    return periodos


def calcular_metricas(periodos):
    """Devuelve un DataFrame con los agregados de cada periodo."""
    filas = []
    for nombre, df in periodos.items():
        fila = {"periodo": nombre, "empresas": len(df)}

        emp_op = to_num(df[find_col(df, METRICAS["empleados_operativos"])])
        emp_ad = to_num(df[find_col(df, METRICAS["empleados_administrativos"])])
        emp_ot = to_num(df[find_col(df, METRICAS["otro_personal"])])

        fila["headcount"] = emp_op.sum() + emp_ad.sum() + emp_ot.sum()
        hc_por_empresa = emp_op.fillna(0) + emp_ad.fillna(0) + emp_ot.fillna(0)
        fila["hc_max_empresa"] = hc_por_empresa.max()
        fila["sal_diario_max"] = to_num(df[find_col(df, METRICAS["salario_diario_maximo"])]).mean()
        fila["sal_diario_ini"] = to_num(df[find_col(df, METRICAS["salario_diario_inicial"])]).mean()
        fila["contrataciones"] = to_num(df[find_col(df, METRICAS["contrataciones"])]).sum()
        fila["bajas"] = to_num(df[find_col(df, METRICAS["bajas"])]).sum()
        fila["faltas"] = to_num(df[find_col(df, METRICAS["faltas"])]).sum()
        fila["incapacidades"] = to_num(df[find_col(df, METRICAS["incapacidades"])]).sum()
        fila["capacitacion"] = to_num(df[find_col(df, METRICAS["capacitacion"])]).sum()
        fila["rotacion"] = (fila["bajas"] / fila["headcount"] * 100) if fila["headcount"] else float("nan")
        filas.append(fila)

    resumen = pd.DataFrame(filas)
    return resumen


def seccion(titulo, prioridad):
    linea = "=" * 78
    print(f"\n{linea}")
    print(f"[{prioridad}] {titulo}")
    print(linea)


def main():
    configurar_consola()

    print("=" * 78)
    print("INFORME DE BASE HISTORICA — Encuesta laboral bimestral")
    print("=" * 78)

    periodos = cargar_periodos()
    resumen = calcular_metricas(periodos)

    # ---------------------------------------------------------------
    # 1) VOLUMEN Y COBERTURA
    # ---------------------------------------------------------------
    seccion("VOLUMEN Y COBERTURA", "PRIORIDAD CRÍTICA")

    total_registros = sum(len(d) for d in periodos.values())
    empresas_por_periodo = {p: len(d) for p, d in periodos.items()}
    conjuntos_empresas = [set(d["Compañía"]) for d in periodos.values()]
    panel_comun = set.intersection(*conjuntos_empresas)
    panel_total = set.union(*conjuntos_empresas)

    print(f"Periodos bimestrales capturados : {len(periodos)}")
    print(f"Registros (respuestas) totales : {total_registros}")
    print(f"Rango de periodos               : {list(periodos)[0]} → {list(periodos)[-1]}")
    print(f"Empresas únicas en total        : {len(panel_total)}")
    print(f"Empresas presentes en TODOS los periodos (panel): {len(panel_comun)}")

    print("\nRespuestas por periodo:")
    for p, n in empresas_por_periodo.items():
        print(f"   {p:<14} : {n:>3} registros")

    # ---------------------------------------------------------------
    # 2) CALIDAD DE DATOS
    # ---------------------------------------------------------------
    seccion("CALIDAD DE DATOS", "PRIORIDAD ALTA")

    # Cobertura de las columnas clave usando la primera hoja como referencia.
    hoja_ref = periodos[list(periodos)[0]]
    print("Cobertura (respuestas no vacías) de los campos clave:")
    for etiqueta, frase in METRICAS.items():
        col = find_col(hoja_ref, frase)
        if col is None:
            print(f"   {etiqueta:<24} : columna NO encontrada")
            continue
        n = to_num(hoja_ref[col]).notna().sum()
        print(f"   {etiqueta:<24} : {n:>3}/{len(hoja_ref)}")

    vacios_global = sum(
        d.isna().sum().sum() for d in periodos.values()
    )
    celdas_global = sum(d.shape[0] * d.shape[1] for d in periodos.values())
    print(f"\nCeldas vacías en toda la base : {vacios_global:,} de {celdas_global:,} "
          f"({vacios_global / celdas_global * 100:.1f}%)")
    print("   (incluye TODAS las columnas de cada hoja; los campos clave tienen")
    print("    una cobertura mucho mayor, como se muestra arriba).")

    # ---------------------------------------------------------------
    # 3) HEADCOUNT
    # ---------------------------------------------------------------
    seccion("HEADCOUNT (empleo)", "PRIORIDAD ALTA")

    head_total = resumen["headcount"].sum()
    print(f"Headcount total (suma de las 7 mediciones): {fmt_num(head_total)}")
    print(f"Headcount promedio por periodo            : {fmt_num(resumen['headcount'].mean())}")
    print(f"Headcount promedio por empresa/periodo    : {fmt_num(head_total / total_registros)}")

    print("\nEvolución por periodo:")
    resumen_hc = resumen[["periodo", "empresas", "headcount"]].copy()
    resumen_hc["headcount"] = resumen_hc["headcount"].map(lambda x: fmt_num(x))
    print(resumen_hc.to_string(index=False))

    # Señalar periodos con una sola empresa concentrando mucho headcount.
    atipicos = resumen[
        resumen["hc_max_empresa"] > resumen["headcount"] * 0.30
    ]
    for _, fila in atipicos.iterrows():
        print(f"   * {fila['periodo']}: una sola empresa reportó {fmt_num(fila['hc_max_empresa'])} "
              f"empleados, lo que eleva el headcount de ese periodo.")

    # ---------------------------------------------------------------
    # 4) SALARIOS
    # ---------------------------------------------------------------
    seccion("SALARIOS", "PRIORIDAD ALTA")

    print("Salarios diarios (personal operativo) — promedio de toda la base:")
    for etiqueta, frase in [
        ("Inicial", METRICAS["salario_diario_inicial"]),
        ("Intermedio - Bajo", METRICAS["salario_diario_inter_bajo"]),
        ("Intermedio - Alto", METRICAS["salario_diario_inter_alto"]),
        ("Máximo", METRICAS["salario_diario_maximo"]),
    ]:
        serie = pd.concat([to_num(d[find_col(d, frase)]) for d in periodos.values()])
        print(f"   Salario diario {etiqueta:<18}: ${fmt_num(serie.mean(), 2)} MXN")

    print("\nSalarios mensuales (posiciones técnicas) — promedio de toda la base:")
    for etiqueta, frase in [
        ("Mínimo", METRICAS["salario_min_mensual"]),
        ("Promedio", METRICAS["salario_prom_mensual"]),
        ("Máximo", METRICAS["salario_max_mensual"]),
    ]:
        serie = pd.concat([to_num(d[find_col(d, frase)]) for d in periodos.values()])
        print(f"   Salario mensual {etiqueta:<10}: ${fmt_num(serie.mean(), 0)} MXN")

    print("\nEvolución del salario diario (inicial → máximo):")
    resumen_sal = resumen[["periodo", "sal_diario_ini", "sal_diario_max"]].copy()
    resumen_sal["sal_diario_ini"] = resumen_sal["sal_diario_ini"].map(lambda x: fmt_num(x, 2))
    resumen_sal["sal_diario_max"] = resumen_sal["sal_diario_max"].map(lambda x: fmt_num(x, 2))
    print(resumen_sal.to_string(index=False))

    # ---------------------------------------------------------------
    # 5) CONTRATACIÓN, ROTACIÓN Y AUSENTISMO
    # ---------------------------------------------------------------
    seccion("CONTRATACIÓN, ROTACIÓN Y AUSENTISMO", "PRIORIDAD ALTA")

    print(f"Contrataciones totales (7 bimestres) : {fmt_num(resumen['contrataciones'].sum())}")
    print(f"Bajas totales (7 bimestres)           : {fmt_num(resumen['bajas'].sum())}")
    print(f"Faltas totales (7 bimestres)          : {fmt_num(resumen['faltas'].sum())}")
    print(f"Incapacidades totales (7 bimestres)   : {fmt_num(resumen['incapacidades'].sum())}")

    rotacion_global = (
        resumen["bajas"].sum() / resumen["headcount"].sum() * 100
        if resumen["headcount"].sum() else float("nan")
    )
    print(f"Rotación bimestral estimada (bajas/headcount): {rotacion_global:.1f}%")

    print("\nEvolución por periodo:")
    resumen_rot = resumen[
        ["periodo", "contrataciones", "bajas", "faltas", "incapacidades", "rotacion"]
    ].copy()
    for c in ["contrataciones", "bajas", "faltas", "incapacidades"]:
        resumen_rot[c] = resumen_rot[c].map(lambda x: fmt_num(x))
    resumen_rot["rotacion"] = resumen_rot["rotacion"].map(lambda x: f"{x:.1f}%" if pd.notna(x) else "-")
    print(resumen_rot.to_string(index=False))

    # ---------------------------------------------------------------
    # 6) CAPACITACIÓN
    # ---------------------------------------------------------------
    seccion("CAPACITACIÓN", "PRIORIDAD MEDIA")

    print(f"Horas de capacitación totales (7 bimestres): {fmt_num(resumen['capacitacion'].sum())}")
    print(f"Promedio por periodo                       : {fmt_num(resumen['capacitacion'].mean())}")
    print("\nEvolución por periodo:")
    resumen_cap = resumen[["periodo", "capacitacion"]].copy()
    resumen_cap["capacitacion"] = resumen_cap["capacitacion"].map(lambda x: fmt_num(x))
    print(resumen_cap.to_string(index=False))

    # ---------------------------------------------------------------
    # 7) SINDICALIZACIÓN Y CCT
    # ---------------------------------------------------------------
    seccion("SINDICALIZACIÓN Y NEGOCIACIÓN (CCT)", "PRIORIDAD MEDIA")

    for etiqueta, frase in CATEGORICAS.items():
        col = find_col(hoja_ref, frase)
        if col is None:
            print(f"   {etiqueta:<20}: columna NO encontrada")
            continue
        conteos = pd.concat([d[col] for d in periodos.values()])
        conteos = conteos.dropna().astype(str).str.strip()
        conteos = conteos[~conteos.isin(["", "-", "nan", "NaN", "None", "Open-Ended Response", "Response"])]
        if conteos.empty:
            continue
        print(f"\n{etiqueta} ({len(conteos)} respuestas en total):")
        for valor, n in conteos.value_counts().head(6).items():
            pct = n / len(conteos) * 100
            print(f"   {str(valor)[:60]:<60} : {n:>3}  ({pct:.0f}%)")

    # ---------------------------------------------------------------
    # 8) RESUMEN GLOBAL POR PERIODO
    # ---------------------------------------------------------------
    seccion("TABLA RESUMEN POR PERIODO", "RESUMEN GENERAL")

    tabla = resumen.copy()
    tabla["headcount"] = tabla["headcount"].map(lambda x: fmt_num(x))
    tabla["sal_diario_max"] = tabla["sal_diario_max"].map(lambda x: fmt_num(x, 2))
    tabla["contrataciones"] = tabla["contrataciones"].map(lambda x: fmt_num(x))
    tabla["bajas"] = tabla["bajas"].map(lambda x: fmt_num(x))
    tabla["faltas"] = tabla["faltas"].map(lambda x: fmt_num(x))
    tabla["incapacidades"] = tabla["incapacidades"].map(lambda x: fmt_num(x))
    tabla["capacitacion"] = tabla["capacitacion"].map(lambda x: fmt_num(x))
    tabla = tabla.drop(columns=["sal_diario_ini", "rotacion", "hc_max_empresa"])
    print(tabla.to_string(index=False))

    print("\nFin del informe.")


if __name__ == "__main__":
    main()
