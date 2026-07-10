# -*- coding: utf-8 -*-
"""
Demostración: definís los kits una vez y el motor genera todas las variantes
por categoría, agrupadas en hojas. Replica tu pedido manual a GPT.
"""
from generador import exportar_excel

# Un KIT es una lista de (componente, cantidad). Se define UNA vez.
HOJAS = [
    {
        "nombre": "01_KIT_2BALIZAS",
        "kit": [("crique", 1), ("llave_cruz", 1), ("guantes", 1),
                ("chaleco", 1), ("baliza", 2)],
        "categorias": ["kit_seguridad", "crique", "llave_cruz", "chaleco", "baliza"],
    },
    {
        "nombre": "02_KIT_CONO",
        "kit": [("crique", 1), ("llave_cruz", 1), ("guantes", 1),
                ("chaleco", 1), ("baliza", 2), ("cono", 1)],
        "categorias": ["kit_seguridad", "crique", "llave_cruz",
                       "chaleco", "baliza", "cono"],
    },
    {
        "nombre": "03_CRIQUE_COMPRESOR",
        "kit": [("crique", 1), ("compresor", 1)],
        "categorias": ["kit_seguridad", "crique", "compresor"],
    },
]

if __name__ == "__main__":
    ruta = exportar_excel(HOJAS, "/home/claude/generador/DESCRIPCIONES_GENERADAS.xlsx")
    print("OK ->", ruta)

    # muestra por consola una variante para revisar el texto
    from generador import generar_descripcion
    print("\n" + "=" * 70)
    print("MUESTRA: kit crique+compresor en categoría 'compresor'")
    print("=" * 70)
    print(generar_descripcion(HOJAS[2]["kit"], "compresor"))
