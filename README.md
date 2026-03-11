# Registro MTG — Meta Technology Global

Plataforma de registro de instalaciones **Emporia Vue 3** para la **Plataforma PAOER**.  
Funciona como app en Android, iPhone e iPad — sin instalar nada.

---

## 🚀 Cómo usar en el teléfono

### Android (Chrome)
1. Abre el link de GitHub Pages en Chrome
2. Menú `⋮` → **Añadir a pantalla de inicio**

### iPhone / iPad (Safari)
1. Abre el link en Safari
2. Botón compartir `⬆` → **Añadir a pantalla de inicio**

---

## 📋 Qué registra

- Tipo de propiedad (Casa / Apartamento / Comercial / Industrial)
- Datos del propietario y técnico instalador
- Sistema energético: Solo red / Panel Solar / Solar + Batería
- Factura eléctrica de los últimos 2 meses
- Medidor Emporia Vue 3: serie, ubicación, canales CT, conectividad
- Fotos por categoría: fachada, panel, breakers, medidor, sensores CT
- Mapa completo de circuitos / breakers con sensor CT asignado
- Estado eléctrico y observaciones

---

## 💾 Exportar datos

### Backup (recomendado)
Desde la app: **Exportar → Copia de Seguridad (.json)**  
Guarda todos los datos **incluyendo fotos**.

### Restaurar en otro dispositivo
Desde la app: **Exportar → Restaurar Backup** → selecciona el `.json`

---

## 📊 Generar Excel profesional

El Excel con colores, headers y formato profesional se genera con Python en tu PC.

### Requisitos
```bash
pip install openpyxl
```

### Uso
```bash
python3 json_to_excel.py  MTG_Backup_2024-11-20.json
```

Genera automáticamente `MTG_PAOER_FECHA.xlsx` con 4 hojas:
- **Portada** — resumen general con totales
- **Propiedades** — todos los datos de instalación
- **Circuitos** — mapa de breakers por propiedad
- **Facturas** — facturas con promedio calculado

---

## 🗂 Estructura del proyecto

```
registro-mtg/
├── index.html          ← App principal (abrir en el teléfono)
├── json_to_excel.py    ← Script para generar Excel profesional
└── README.md           ← Este archivo
```

---

## ⚙️ Tecnologías

- HTML5 + CSS3 + JavaScript puro (sin frameworks)
- Almacenamiento: `localStorage` del navegador
- Exportación Excel: Python + openpyxl
- Fuentes: Playfair Display + Outfit (Google Fonts)

---

**Meta Technology Global** · Plataforma PAOER · Emporia Vue 3
