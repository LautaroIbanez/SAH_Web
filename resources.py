"""
Archivo de recursos y constantes para el Sistema de Adelantos Haberes.
"""

# Códigos de conceptos remunerativos (código: {'concepto': ..., 'clave': ...})
CODIGOS_BRUTO = {
    "20":  {"concepto": "Básico de Convenio 264/95", "clave": "Básico"},
    "30":  {"concepto": "Ex Adicional NoRem.Volunt.Var.porCom", "clave": "Adicional"},
    "40":  {"concepto": "A Cuenta Futuros Aumentos Anteriores", "clave": "Aumentos"},
    "97":  {"concepto": "A Cuenta Futuros Aumentos 09-2024", "clave": "Aumentos"},
    "103": {"concepto": "A Cuenta Futuros Aumentos 01-2025", "clave": "Aumentos"},
    "280": {"concepto": "Presentismo", "clave": "Presentismo"},
    "281": {"concepto": "Productividad", "clave": "Productividad"},
    "330": {"concepto": "L26341Canasta Pass 100% 10 de10 bimestres", "clave": "Canasta"},
    "350": {"concepto": "Almuerzo Acuerdo Salarial CCT 264/95", "clave": "Almuerzo"}
}

# Códigos de deducciones (código: {'concepto': ..., 'clave': ...})
CODIGOS_DEDUCCIONES = {
    "7000": {"concepto": "Jubilacion", "clave": "Jubilacion"},
    "7005": {"concepto": "INSSJP", "clave": "INSSJP"},
    "7010": {"concepto": "OSSEG", "clave": "OSSEG"},
    "8005": {"concepto": "Impuesto Ganancias 4ªCat.", "clave": "Ganancias"}
}

MOTIVOS = [
    "Gastos médicos",
    "Reparación del hogar",
    "Compra de equipamiento",
    "Vacaciones",
    "Educación",
    "Otro"
]

# Tope máximo para préstamos (en pesos)
TOPE_MAXIMO_PRESTAMO = 5_000_000  # 5 millones de pesos


