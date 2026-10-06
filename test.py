import re

import pandas as pd

ARCHIVO = 'BASE HISTORICA.xlsx'
SALIDA = 'columnas_por_tipo.xlsx'

df = pd.read_excel(ARCHIVO, sheet_name='JUL-AGO_2024')

print(df.info())
print('==============================================')
print(df.describe())
print('==============================================')

selection = df.iloc[:, 0:20]
print(selection)


def _nombre_hoja_valido(texto):
    """Nombre de hoja de Excel válido (sin caracteres prohibidos y máx. 31 chars)."""
    limpio = re.sub(r"[\[\]:*?/\\]", "_", str(texto))
    return limpio[:31] or "Tipo"


def exportar_columnas_por_tipo(archivo=ARCHIVO, salida=SALIDA, hojas=None):
    """Exporta a Excel los nombres de columna agrupados por tipo de dato.

    Recorre las hojas indicadas (todas por defecto) y escribe:
      - Hoja "Resumen": cuántas columnas hay de cada tipo de dato.
      - Una hoja por tipo de dato: nombre de la columna y hojas donde aparece.
    """
    xl = pd.ExcelFile(archivo)
    hojas = list(hojas) if hojas else xl.sheet_names

    registros = []
    for hoja in hojas:
        datos = pd.read_excel(xl, sheet_name=hoja)
        for columna, tipo in datos.dtypes.astype(str).items():
            registros.append({"tipo": tipo, "columna": str(columna), "hoja": hoja})

    catalogo = pd.DataFrame(registros)

    agrupado = (
        catalogo.groupby(["tipo", "columna"], sort=True)["hoja"]
        .agg(
            n_hojas=lambda s: len(set(s)),
            hojas=lambda s: ", ".join(sorted(set(s))),
        )
        .reset_index()
        .rename(columns={
            "tipo": "Tipo de dato",
            "columna": "Columna",
            "n_hojas": "Nº de hojas",
            "hojas": "Hojas",
        })
    )

    resumen = (
        catalogo.groupby("tipo")
        .agg(columnas=("columna", "nunique"), apariciones=("columna", "size"))
        .reset_index()
        .rename(columns={
            "tipo": "Tipo de dato",
            "columnas": "Columnas únicas",
            "apariciones": "Columnas (una por hoja)",
        })
    )

    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        resumen.to_excel(writer, sheet_name="Resumen", index=False)
        for tipo, grupo in agrupado.groupby("Tipo de dato", sort=True):
            grupo[["Columna", "Nº de hojas", "Hojas"]].to_excel(
                writer, sheet_name=_nombre_hoja_valido(tipo), index=False
            )

    print(f"\nCatálogo de columnas exportado a: {salida}")
    print(resumen.to_string(index=False))
    return salida


if __name__ == "__main__":
    exportar_columnas_por_tipo()