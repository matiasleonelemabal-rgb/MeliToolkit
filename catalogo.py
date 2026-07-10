# -*- coding: utf-8 -*-
"""
Catálogo de componentes y categorías para el generador de descripciones.

IDEA CENTRAL
------------
Cada componente es un bloque de texto FIJO (con sus medidas exactas).
El generador solo ELIGE y ORDENA bloques; nunca reescribe ni cambia una medida.
Por eso es imposible que "se le escape" un dato como le pasó a GPT
(150 PSI -> 140 PSI, 25 L/min -> 35 L/min, etc.).

Si mañana cambia una medida real, se edita UNA sola vez acá y se
regeneran todas las publicaciones que usan ese componente.
"""

# ============================================================
# COMPONENTES
# ============================================================
# Cada componente define:
#   nombre_corto  -> para el título y la lista "EL KIT PUBLICADO INCLUYE"
#   inclusion     -> función(cantidad) que devuelve la línea de la lista
#   titulo_bloque -> encabezado del bloque descriptivo
#   bloque        -> función(cantidad) que devuelve el párrafo descriptivo
#   uso           -> fragmento corto para "USOS RECOMENDADOS"

COMPONENTES = {
    "crique": {
        "nombre_corto": "crique carrito hidráulico 2 Toneladas",
        "inclusion": lambda c: f"{c} unidad de crique carrito hidráulico 2 Toneladas.",
        "titulo_bloque": "CRIQUE HIDRÁULICO TIPO CARRITO 2 TONELADAS",
        "bloque": lambda c: (
            "El crique hidráulico tipo carrito 2 Toneladas es una herramienta esencial para "
            "elevar el vehículo de forma estable y controlada durante tareas de auxilio, "
            "mantenimiento básico o cambio de rueda.\n\n"
            "Su capacidad de levante es de 2 Toneladas, equivalente a 2000 kg. Cuenta con "
            "mecanismo hidráulico cerrado con válvula de liberación de precisión, lo que ayuda "
            "a controlar el descenso durante el uso. Posee ruedas giratorias traseras de acero "
            "para facilitar el posicionamiento bajo el chasis.\n\n"
            "Sus medidas aproximadas son 42 cm de largo, 13,5 cm de ancho delantero y 19 cm de "
            "ancho trasero. El rango de elevación aproximado va desde 13 cm hasta 32 cm. El peso "
            "aproximado es de 6,5 kg.\n\n"
            "La palanca de accionamiento posee un largo aproximado de 38 cm y diámetro de agarre "
            "aproximado de 2,3 cm."
        ),
        "uso": "cambio de neumáticos, recambio de rueda de emergencia, mantenimiento básico del vehículo",
    },

    "llave_cruz": {
        "nombre_corto": "llave cruz niquelada multimedida",
        "inclusion": lambda c: f"{c} unidad de llave cruz niquelada multimedida.",
        "titulo_bloque": "LLAVE CRUZ NIQUELADA MULTIMEDIDA",
        "bloque": lambda c: (
            "La llave cruz niquelada incluida complementa el uso del crique durante el recambio "
            "del neumático. Es una llave de rueda tipo cruz, también conocida como cruceta o "
            "llave para bulones, ideal para aflojar y ajustar bulones en autos, camionetas "
            "livianas y vehículos de uso particular.\n\n"
            "Está fabricada en acero forjado de alta resistencia mecánica, con tratamiento "
            "niquelado anticorrosión y antióxido. Cuenta con 4 bocas integradas milimétricas de "
            "17 mm, 19 mm, 21 mm y 23 mm, cubriendo la mayoría de las medidas de bulones de "
            "rueda más utilizadas en el mercado argentino.\n\n"
            "Su largo total aproximado es de 35,5 cm, lo que permite lograr buen torque de "
            "trabajo durante el cambio de rueda."
        ),
        "uso": "aflojar y ajustar bulones de rueda",
    },

    "baliza": {
        # nombre_corto y textos dependen de la cantidad -> se resuelven con funciones
        "nombre_corto": lambda c: (
            "baliza reflectiva triangular" if c == 1 else "balizas reflectivas triangulares"
        ),
        "inclusion": lambda c: (
            "1 unidad de baliza reflectiva triangular homologada con CHAS."
            if c == 1 else
            "2 unidades de baliza reflectiva triangular."
        ),
        "titulo_bloque": lambda c: (
            "BALIZA REFLECTIVA TRIANGULAR HOMOLOGADA CON CHAS" if c == 1
            else "BALIZAS REFLECTIVAS TRIANGULARES"
        ),
        "bloque": lambda c: (
            (
                "La baliza reflectiva triangular incluida es ideal para señalizar el vehículo "
                "ante una detención, emergencia, desperfecto mecánico o cambio de neumático en "
                "ruta, banquina, calle, cochera o vía pública.\n\n"
                "Está fabricada en material acrílico rojo con tela reflectiva, diseñada para "
                "aportar mayor visibilidad nocturna a distancia. Su formato triangular permite "
                "una señalización clara en situaciones de emergencia vial.\n\n"
                "Cuenta con diseño compacto, de fácil y rápido armado. Posee 4 patas o soportes "
                "metálicos extensibles, 2 delanteros y 2 traseros, que ayudan a mejorar la "
                "estabilidad durante el uso.\n\n"
                "Incluye caja protectora roja para un almacenamiento seguro, práctico y "
                "ordenado dentro del vehículo.\n\n"
                "Medidas aproximadas de la baliza: ancho 45 cm, alto 39 cm y largo con patas "
                "extendidas de extremo a extremo 74,5 cm.\n\n"
                "Medidas aproximadas de la caja de guardado: largo 45 cm, alto 4 cm y ancho 3 cm."
            ) if c == 1 else (
                "Las 2 balizas reflectivas triangulares incluidas son ideales para señalizar el "
                "vehículo ante una detención, emergencia, desperfecto mecánico o cambio de "
                "neumático en ruta, banquina, calle, cochera o vía pública.\n\n"
                "Están fabricadas en material acrílico rojo con tela reflectiva, diseñadas para "
                "aportar mayor visibilidad nocturna a distancia. Su formato triangular permite "
                "una señalización clara en situaciones de emergencia vial.\n\n"
                "Cada baliza cuenta con diseño compacto, de fácil y rápido armado. Posee 4 patas "
                "o soportes metálicos extensibles, 2 delanteros y 2 traseros, que ayudan a "
                "mejorar la estabilidad durante el uso.\n\n"
                "Cada unidad incluye caja protectora roja para un almacenamiento seguro, "
                "práctico y ordenado dentro del vehículo.\n\n"
                "Medidas aproximadas de cada baliza: ancho 45 cm, alto 39 cm y largo con patas "
                "extendidas de extremo a extremo 74,5 cm.\n\n"
                "Medidas aproximadas de cada caja de guardado: largo 45 cm, alto 4 cm y ancho 3 cm."
            )
        ),
        "uso": "señalización de emergencia vehicular",
    },

    "chaleco": {
        "nombre_corto": "chaleco reflectivo de seguridad vial",
        "inclusion": lambda c: f"{c} unidad de chaleco reflectivo de seguridad vial.",
        "titulo_bloque": "CHALECO REFLECTIVO DE SEGURIDAD VIAL",
        "bloque": lambda c: (
            "El chaleco reflectivo de seguridad vial ayuda a mejorar la visibilidad del usuario "
            "en situaciones de emergencia, auxilio vial, trabajos en exteriores, detenciones "
            "sobre banquina, ruta, calle, garaje o zonas de baja visibilidad.\n\n"
            "Cuenta con color flúor vibrante y bandas reflectantes de 360 grados. Está fabricado "
            "en poliéster de malla ligera y respirable, ideal para colocarlo rápidamente sobre "
            "la ropa.\n\n"
            "Posee sistema de abrojo frontal para una apertura y cierre simple. Es talle "
            "universal, con diseño amplio y cómodo que se adapta a diferentes contexturas "
            "físicas.\n\n"
            "El color puede variar entre amarillo flúor y naranja flúor según disponibilidad de "
            "stock. No es seleccionable.\n\n"
            "Medidas aproximadas del chaleco reflectivo: largo 67,5 cm y ancho 56 cm."
        ),
        "uso": "mejorar la visibilidad del usuario",
    },

    "guantes": {
        "nombre_corto": "par de guantes de trabajo",
        "inclusion": lambda c: f"{c} par de guantes de trabajo.",
        "titulo_bloque": "GUANTES DE TRABAJO",
        "bloque": lambda c: (
            "Los guantes de trabajo son ideales para proteger las manos durante tareas de "
            "mecánica ligera, carga y descarga, uso de crique, uso de llave cruz, armado de "
            "balizas, despliegue del cono y manipulación de piezas o herramientas.\n\n"
            "Cuentan con puntos de PVC antideslizantes en la palma para lograr un mejor agarre. "
            "Están fabricados con tejido de algodón y poliéster, brindando protección sin "
            "sacrificar comodidad durante el uso.\n\n"
            "Poseen puños elásticos que ayudan a lograr un ajuste correcto y reducir el ingreso "
            "de residuos. Son de tamaño universal, único talle.\n\n"
            "Medidas aproximadas de los guantes de trabajo: largo 22 cm y ancho 13 cm."
        ),
        "uso": "proteger las manos durante la maniobra",
    },

    "cono": {
        "nombre_corto": "cono reflectivo plegable",
        "inclusion": lambda c: f"{c} unidad de cono reflectivo plegable.",
        "titulo_bloque": "CONO REFLECTIVO PLEGABLE 40 CM",
        "bloque": lambda c: (
            "El cono reflectivo plegable incluido es ideal para reforzar la señalización visual "
            "en emergencias viales, delimitación de áreas, estacionamientos, cocheras, "
            "mantenimiento en vía pública o situaciones de auxilio vehicular.\n\n"
            "Cuenta con diseño plegable y telescópico que se pliega totalmente, optimizando el "
            "espacio de guardado en el baúl del vehículo. Está fabricado con tela impermeable "
            "color naranja flúor y banda reflectiva de alta intensidad.\n\n"
            "Posee base cuadrada plástica, diseñada para brindar apoyo firme, estabilidad y "
            "resistencia durante el uso.\n\n"
            "Medidas aproximadas del cono reflectivo plegable: alto total 40 cm.\n\n"
            "Medidas aproximadas de la base plástica: 24,5 cm x 24,5 cm. Altura aproximada de "
            "base: 4 cm."   # <- corrección: base 4 cm, NO 40 cm (dato incoherente en la ficha)
        ),
        "uso": "reforzar la señalización visual y delimitación de áreas",
    },

    "compresor": {
        "nombre_corto": "compresor de aire portátil 12V",
        "inclusion": lambda c: f"{c} unidad de compresor de aire portátil 12V (incluye 3 picos adaptadores de inflado).",
        "titulo_bloque": "COMPRESOR DE AIRE PORTÁTIL 12V",
        "bloque": lambda c: (
            "El compresor de aire portátil 12V permite realizar tareas de inflado de forma "
            "práctica utilizando la alimentación eléctrica del vehículo.\n\n"
            "Cuenta con cuerpo fabricado en metal, aleación de aluminio y componentes plásticos. "
            "Posee manómetro incorporado en la parte superior del equipo para controlar la "
            "presión durante el inflado. Incluye botón de encendido y apagado ON/OFF, manija "
            "superior para facilitar el agarre y transporte, manguera de aire y cable de "
            "alimentación con conexión para ficha de encendedor 12V. El kit incorpora 3 picos "
            "adaptadores de inflado.\n\n"
            "Características: voltaje DC 12V; amperaje máximo 10 Amp; presión máxima 150 PSI; "
            "desplazamiento de aire 25 L/min (+/- 10%).\n\n"
            "Medidas aproximadas del compresor: largo 16 cm, alto 14 cm y ancho 8 cm. "
            "Cable de alimentación 12V: largo aproximado 260 cm. "
            "Manguera de aire: largo aproximado 90 cm. Manómetro: diámetro aproximado 5 cm."
        ),
        "uso": "inflado de neumáticos y control de presión",
    },
}


def nombre_corto(comp_id, cantidad):
    """Devuelve el nombre corto, resolviendo si es función (caso baliza)."""
    nc = COMPONENTES[comp_id]["nombre_corto"]
    return nc(cantidad) if callable(nc) else nc


def titulo_bloque(comp_id, cantidad):
    tb = COMPONENTES[comp_id]["titulo_bloque"]
    return tb(cantidad) if callable(tb) else tb


# ============================================================
# CATEGORÍAS
# ============================================================
# Cada categoría define cuál componente es el "principal" (va primero y ancla
# el título / la intro) y una plantilla de intro. Las intros son plantillas
# FIJAS: no las genera una IA, así no hay variación inventada.

CATEGORIAS = {
    "kit_seguridad": {
        "etiqueta": "Kit de seguridad para auto",
        "principal": None,   # el kit no tiene un principal: ancla en el conjunto
    },
    "crique": {"etiqueta": "Crique hidráulico", "principal": "crique"},
    "llave_cruz": {"etiqueta": "Llave cruz / rueda", "principal": "llave_cruz"},
    "chaleco": {"etiqueta": "Chaleco reflectivo", "principal": "chaleco"},
    "baliza": {"etiqueta": "Baliza triangular", "principal": "baliza"},
    "cono": {"etiqueta": "Cono de seguridad vial", "principal": "cono"},
    "compresor": {"etiqueta": "Compresor / inflador", "principal": "compresor"},
}
