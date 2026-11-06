"""
Genera datos ficticios para probar el validador de inventario.

Crea dos archivos en data/:
- inventario.xlsx   -> lo que registra almacén (código, nombre, unidad, stock)
- base_costos.xlsx  -> lo que tiene costos (código, nombre, costo, peso)

A propósito se meten errores típicos que aparecen en la vida real:
códigos duplicados, nombres escritos de distintas formas, productos sin costo,
pesos en cero o negativos y productos que existen en un archivo y no en el otro.

Todos los datos son inventados.
"""

from pathlib import Path
import random

import numpy as np
import pandas as pd

SEMILLA = 42
random.seed(SEMILLA)
np.random.seed(SEMILLA)

CARPETA = Path(__file__).parent / "data"

CATEGORIAS = {
    "ABA": ("Abarrotes", ["Arroz extra", "Azúcar rubia", "Aceite vegetal", "Fideo spaghetti",
                          "Harina sin preparar", "Sal de mesa", "Lenteja", "Pallar", "Avena"]),
    "LAC": ("Lácteos", ["Leche evaporada", "Queso fresco", "Mantequilla", "Yogurt natural",
                        "Crema de leche", "Queso parmesano"]),
    "CAR": ("Carnes", ["Pechuga de pollo", "Lomo fino", "Carne molida", "Chorizo",
                       "Tocino ahumado", "Filete de pescado"]),
    "VER": ("Verduras", ["Cebolla roja", "Tomate italiano", "Papa amarilla", "Ají amarillo",
                         "Limón", "Culantro", "Zanahoria", "Choclo"]),
    "BEB": ("Bebidas", ["Gaseosa 500 ml", "Agua sin gas 625 ml", "Chicha morada 1 L",
                        "Jugo de naranja 1 L", "Cerveza 330 ml"]),
    "DES": ("Descartables", ["Vaso 12 oz", "Tapa vaso 12 oz", "Bolsa kraft", "Servilleta",
                             "Envase 1 L", "Cubierto descartable"]),
}

PRESENTACIONES = ["x 1 kg", "x 5 kg", "x 500 g", "x 250 g", "x 12 und", "x 24 und", "x 50 und", "granel"]
UNIDADES = ["KG", "UND", "LT", "PAQ"]


def armar_catalogo(n: int = 310) -> pd.DataFrame:
    # Cada combinación producto + presentación aparece una sola vez
    combinaciones = [
        (pref, f"{nombre} {presentacion}")
        for pref, (_, nombres) in CATEGORIAS.items()
        for nombre in nombres
        for presentacion in PRESENTACIONES
    ]
    random.shuffle(combinaciones)

    filas = []
    correlativo = {pref: 1 for pref in CATEGORIAS}
    for pref, nombre in combinaciones[:n]:
        categoria = CATEGORIAS[pref][0]
        codigo = f"{pref}-{correlativo[pref]:04d}"
        correlativo[pref] += 1
        filas.append({
            "codigo": codigo,
            "nombre": nombre,
            "categoria": categoria,
            "unidad": random.choice(UNIDADES),
            "stock": int(np.random.gamma(2, 25)),
            "costo_unitario": round(float(np.random.uniform(0.3, 85)), 2),
            "peso_kg": round(float(np.random.uniform(0.05, 25)), 3),
        })
    return pd.DataFrame(filas)


def ensuciar_nombre(nombre: str) -> str:
    """Simula cómo se escribe a mano el mismo producto en distintos registros."""
    cambios = [
        lambda s: s.upper(),
        lambda s: s.lower(),
        lambda s: "  " + s + " ",
        lambda s: s.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u"),
        lambda s: s.replace(" x ", " X "),
        lambda s: s.replace(" ", "  ", 1),
    ]
    return random.choice(cambios)(nombre)


def main() -> None:
    CARPETA.mkdir(exist_ok=True)
    catalogo = armar_catalogo()

    inventario = catalogo[["codigo", "nombre", "categoria", "unidad", "stock"]].copy()
    costos = catalogo[["codigo", "nombre", "costo_unitario", "peso_kg"]].copy()

    # 1) Nombres escritos distinto en el inventario (~15 %)
    idx = inventario.sample(frac=0.15, random_state=SEMILLA).index
    inventario.loc[idx, "nombre"] = inventario.loc[idx, "nombre"].map(ensuciar_nombre)

    # 2) Códigos duplicados en inventario: mismo código con otro producto
    dup = inventario.sample(6, random_state=1).copy()
    dup["nombre"] = dup["nombre"].map(lambda s: s + " (reingreso)")
    dup["stock"] = np.random.randint(1, 40, len(dup))
    inventario = pd.concat([inventario, dup], ignore_index=True)

    # 2b) El mismo producto creado dos veces con otro código y el nombre un poco distinto
    repetidos = inventario.drop_duplicates("codigo").sample(5, random_state=7).copy()
    repetidos["codigo"] = repetidos["codigo"].str[:4] + (9000 + np.arange(len(repetidos))).astype(str)
    repetidos["nombre"] = repetidos["nombre"].str.replace(" x ", " x  ", regex=False) + "."
    repetidos["stock"] = np.random.randint(1, 20, len(repetidos))
    inventario = pd.concat([inventario, repetidos], ignore_index=True)

    # 3) Productos sin costo o con costo en cero
    sin_costo = costos.sample(9, random_state=2).index
    costos.loc[sin_costo[:5], "costo_unitario"] = np.nan
    costos.loc[sin_costo[5:], "costo_unitario"] = 0

    # 4) Pesos inválidos
    malos = costos.sample(7, random_state=3).index
    costos.loc[malos[:4], "peso_kg"] = 0
    costos.loc[malos[4:], "peso_kg"] = -1.5

    # 5) Productos que solo están en un archivo
    costos = costos.drop(costos.sample(8, random_state=4).index)
    nuevos = pd.DataFrame({
        "codigo": [f"XXX-{i:04d}" for i in range(1, 5)],
        "nombre": ["Producto descontinuado A", "Producto descontinuado B",
                   "Insumo de prueba", "Promo temporal"],
        "costo_unitario": [4.5, 12.0, 1.2, 9.9],
        "peso_kg": [0.5, 1.0, 0.2, 0.8],
    })
    costos = pd.concat([costos, nuevos], ignore_index=True)

    inventario.sample(frac=1, random_state=5).to_excel(CARPETA / "inventario.xlsx", index=False)
    costos.sample(frac=1, random_state=6).to_excel(CARPETA / "base_costos.xlsx", index=False)
    print(f"Inventario: {len(inventario)} filas | Base de costos: {len(costos)} filas")
    print(f"Archivos guardados en {CARPETA}")


if __name__ == "__main__":
    main()
