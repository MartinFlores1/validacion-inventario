"""
Validador de inventario contra la base de costos y pesos.

Antes de usar un inventario para costear recetas o calcular márgenes, conviene
revisar que los datos estén limpios. Este script cruza ambos archivos y genera
un reporte en Excel con todo lo que hay que corregir:

1. Códigos duplicados en el inventario
2. Productos del inventario que no existen en la base de costos (y al revés)
3. Productos sin costo o con costo en cero
4. Pesos en cero o negativos
5. Nombres que no coinciden entre archivos para un mismo código
6. Posibles duplicados por nombre (mismo producto con distinto código)

Uso:
    python validar_inventario.py
    python validar_inventario.py --inventario otra_ruta.xlsx --costos otra_ruta.xlsx
"""

from __future__ import annotations

import argparse
import re
import unicodedata
from difflib import SequenceMatcher
from pathlib import Path

import pandas as pd

BASE = Path(__file__).parent


def normalizar(texto: str) -> str:
    """Pasa a minúsculas, quita tildes y espacios repetidos."""
    if pd.isna(texto):
        return ""
    texto = unicodedata.normalize("NFKD", str(texto))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return " ".join(texto.lower().split())


def cargar(ruta_inv: Path, ruta_cos: Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    inv = pd.read_excel(ruta_inv)
    cos = pd.read_excel(ruta_cos)
    for df in (inv, cos):
        df["codigo"] = df["codigo"].astype(str).str.strip().str.upper()
        df["nombre_norm"] = df["nombre"].map(normalizar)
    return inv, cos


def codigos_duplicados(inv: pd.DataFrame) -> pd.DataFrame:
    dup = inv[inv.duplicated("codigo", keep=False)]
    return dup.sort_values("codigo")[["codigo", "nombre", "categoria", "stock"]]


def faltantes(inv: pd.DataFrame, cos: pd.DataFrame) -> pd.DataFrame:
    cruce = inv[["codigo", "nombre"]].drop_duplicates("codigo").merge(
        cos[["codigo", "nombre"]], on="codigo", how="outer",
        suffixes=("_inventario", "_costos"), indicator=True,
    )
    cruce = cruce[cruce["_merge"] != "both"].copy()
    cruce["problema"] = cruce["_merge"].map({
        "left_only": "Está en inventario pero no tiene costo registrado",
        "right_only": "Está en costos pero no en inventario",
    })
    return cruce.drop(columns="_merge")


def costos_invalidos(cos: pd.DataFrame) -> pd.DataFrame:
    malo = cos[cos["costo_unitario"].isna() | (cos["costo_unitario"] <= 0)].copy()
    malo["problema"] = malo["costo_unitario"].map(lambda c: "Sin costo" if pd.isna(c) else "Costo en cero")
    return malo[["codigo", "nombre", "costo_unitario", "problema"]]


def pesos_invalidos(cos: pd.DataFrame) -> pd.DataFrame:
    malo = cos[cos["peso_kg"].isna() | (cos["peso_kg"] <= 0)].copy()
    malo["problema"] = malo["peso_kg"].map(lambda p: "Peso negativo" if p < 0 else "Peso vacío o en cero")
    return malo[["codigo", "nombre", "peso_kg", "problema"]]


def nombres_distintos(inv: pd.DataFrame, cos: pd.DataFrame) -> pd.DataFrame:
    """Mismo código, pero el nombre no coincide ni después de normalizar."""
    cruce = inv.drop_duplicates("codigo").merge(cos, on="codigo", suffixes=("_inv", "_cos"))
    distinto = cruce[cruce["nombre_norm_inv"] != cruce["nombre_norm_cos"]]
    return distinto[["codigo", "nombre_inv", "nombre_cos"]].rename(
        columns={"nombre_inv": "nombre_en_inventario", "nombre_cos": "nombre_en_costos"})


def presentacion(nombre_norm: str) -> str:
    """Números del nombre (500 ml, x 12 und...). Si difieren, son productos distintos."""
    return " ".join(re.findall(r"\d+", nombre_norm))


def posibles_duplicados(inv: pd.DataFrame, umbral: float = 0.93) -> pd.DataFrame:
    """Productos con distinto código cuyo nombre es casi igual y tienen la misma presentación."""
    unicos = inv.drop_duplicates("codigo")[["codigo", "nombre", "nombre_norm", "categoria"]].copy()
    unicos["comparable"] = unicos["nombre_norm"].str.replace(r"[^\w\s]", "", regex=True).str.strip()
    unicos["presentacion"] = unicos["nombre_norm"].map(presentacion)
    pares = []
    for _, grupo in unicos.groupby(["categoria", "presentacion"]):
        filas = grupo.to_dict("records")
        for i, a in enumerate(filas):
            for b in filas[i + 1:]:
                parecido = SequenceMatcher(None, a["comparable"], b["comparable"]).ratio()
                if parecido >= umbral:
                    pares.append({
                        "codigo_1": a["codigo"], "nombre_1": a["nombre"],
                        "codigo_2": b["codigo"], "nombre_2": b["nombre"],
                        "parecido": round(parecido, 3),
                    })
    return pd.DataFrame(pares).sort_values("parecido", ascending=False) if pares else pd.DataFrame(
        columns=["codigo_1", "nombre_1", "codigo_2", "nombre_2", "parecido"])


def guardar_reporte(hojas: dict[str, pd.DataFrame], salida: Path) -> None:
    resumen = pd.DataFrame(
        [{"revision": nombre, "casos": len(df)} for nombre, df in hojas.items()]
    )
    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        resumen.to_excel(writer, sheet_name="Resumen", index=False)
        for nombre, df in hojas.items():
            df.to_excel(writer, sheet_name=nombre[:31], index=False)
        for hoja in writer.sheets.values():
            for columna in hoja.columns:
                ancho = max(len(str(celda.value or "")) for celda in columna)
                hoja.column_dimensions[columna[0].column_letter].width = min(ancho + 2, 60)


def main() -> None:
    parser = argparse.ArgumentParser(description="Valida el inventario contra la base de costos.")
    parser.add_argument("--inventario", type=Path, default=BASE / "data" / "inventario.xlsx")
    parser.add_argument("--costos", type=Path, default=BASE / "data" / "base_costos.xlsx")
    parser.add_argument("--salida", type=Path, default=BASE / "salida" / "reporte_validacion.xlsx")
    args = parser.parse_args()

    inv, cos = cargar(args.inventario, args.costos)
    hojas = {
        "Codigos duplicados": codigos_duplicados(inv),
        "Faltantes": faltantes(inv, cos),
        "Costos invalidos": costos_invalidos(cos),
        "Pesos invalidos": pesos_invalidos(cos),
        "Nombres distintos": nombres_distintos(inv, cos),
        "Posibles duplicados": posibles_duplicados(inv),
    }

    args.salida.parent.mkdir(exist_ok=True)
    guardar_reporte(hojas, args.salida)

    print(f"Productos en inventario: {inv['codigo'].nunique()} | en base de costos: {cos['codigo'].nunique()}")
    for nombre, df in hojas.items():
        print(f"  {nombre:<22} {len(df):>4} casos")
    print(f"Reporte guardado en {args.salida}")


if __name__ == "__main__":
    main()
