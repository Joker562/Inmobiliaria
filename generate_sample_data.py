"""Genera data/propiedades_ejemplo.xlsx con propiedades ficticias para pruebas.

Uso:
    python generate_sample_data.py
"""
from pathlib import Path

import pandas as pd

OUTPUT_PATH = Path(__file__).parent / "data" / "propiedades_ejemplo.xlsx"

# Fotos reales de Unsplash (uso libre, CDN images.unsplash.com) para simular
# el campo "Enlace de Imagen" que traería un Excel real.
UNSPLASH_HOUSES = [
    "https://images.unsplash.com/photo-1570129477492-45c003edd2be?w=900&q=80&auto=format&fit=crop",
    "https://images.unsplash.com/photo-1568605114967-8130f3a36994?w=900&q=80&auto=format&fit=crop",
    "https://images.unsplash.com/photo-1600596542815-ffad4c1539a9?w=900&q=80&auto=format&fit=crop",
    "https://images.unsplash.com/photo-1600585154340-be6161a56a0c?w=900&q=80&auto=format&fit=crop",
    "https://images.unsplash.com/photo-1600607687939-ce8a6c25118c?w=900&q=80&auto=format&fit=crop",
    "https://images.unsplash.com/photo-1512917774080-9991f1c4c750?w=900&q=80&auto=format&fit=crop",
    "https://images.unsplash.com/photo-1583608205776-bfd35f0d9f83?w=900&q=80&auto=format&fit=crop",
    "https://images.unsplash.com/photo-1564013799919-ab600027ffc6?w=900&q=80&auto=format&fit=crop",
    "https://images.unsplash.com/photo-1524758631624-e2822e304c36?w=900&q=80&auto=format&fit=crop",
    "https://images.unsplash.com/photo-1580587771525-78b9dba3b914?w=900&q=80&auto=format&fit=crop",
    "https://images.unsplash.com/photo-1523217582562-09d0def993a6?w=900&q=80&auto=format&fit=crop",
    "https://images.unsplash.com/photo-1576941089067-2de3c901e126?w=900&q=80&auto=format&fit=crop",
    "https://images.unsplash.com/photo-1592595896551-12b371d546d5?w=900&q=80&auto=format&fit=crop",
    "https://images.unsplash.com/photo-1613977257363-707ba9348227?w=900&q=80&auto=format&fit=crop",
    "https://images.unsplash.com/photo-1605146769289-440113cc3d00?w=900&q=80&auto=format&fit=crop",
    # URL intencionalmente rota para probar el fallback a imagen placeholder
    "https://images.unsplash.com/photo-0000000000000-brokenlink?w=900",
]

ZONAS = [
    "Zona Central", "El Pescadero", "Fidepaz", "Colonia Esterito",
    "Costa Baja", "San Rafael", "Benito Juarez", "El Comitan",
    "Palmira", "Zona Aeropuerto", "Colonia Manuel Marquez de Leon",
    "Colonia Los Olivos", "El Centenario", "8 de Octubre", "Colonia Rodriguez",
]

ESTADOS = ["Mercado Abierto", "Desarrollo"]

DATA = []
for i in range(1, 16):
    zona = ZONAS[(i - 1) % len(ZONAS)]
    estado = ESTADOS[i % 2]
    precio = 1_200_000 + (i * 350_000) + (50_000 if estado == "Desarrollo" else 0)
    m2_construccion = 70 + (i * 12) % 220
    m2_terreno = m2_construccion + 40 + (i * 7) % 150
    recamaras = 1 + (i % 4)
    banos = 1 + (i % 3)
    cocheras = i % 4
    imagen = UNSPLASH_HOUSES[(i - 1) % len(UNSPLASH_HOUSES)]

    DATA.append(
        {
            "ID": i,
            "Nombre": f"Casa {zona} Modelo {i}",
            "Ciudad": "La Paz",
            "Colonia / Zona": zona,
            "Direccion": f"Calle Ejemplo #{100 + i}, {zona}",
            "Precio (MXN)": precio,
            "Estado": estado,
            "Recamaras": recamaras,
            "Banos": banos,
            "Estacionamientos": cocheras,
            "m2 Construccion": m2_construccion,
            "m2 Terreno": m2_terreno,
            "Descripcion": (
                f"Propiedad ficticia de prueba en {zona}, La Paz. "
                f"{recamaras} recamaras, {banos} banos, {cocheras} cocheras. "
                "Datos generados unicamente para probar la aplicacion."
            ),
            "Enlace Imagen": imagen,
        }
    )

df = pd.DataFrame(DATA)

OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
df.to_excel(OUTPUT_PATH, index=False, sheet_name="Propiedades")

print(f"Archivo de ejemplo generado en: {OUTPUT_PATH}")
print(f"Total de propiedades: {len(df)}")
