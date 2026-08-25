# 📊 Banco Pichincha - Limpiador y Formateador de Excel (100% Offline)

[🇺🇸 English](README.md) | [🇪🇸 Español](README_ES.md)

[![Licencia: MIT](https://img.shields.io/badge/Licencia-MIT-blue.svg)](LICENSE)
[![Plataformas](https://img.shields.io/badge/Plataformas-Windows%20%7C%20Android%20%7C%20iOS%20%7C%20macOS-brightgreen.svg)]()
[![Seguridad](https://img.shields.io/badge/Seguridad-100%25%20Local%20y%20Privado-success.svg)]()
[![Descargar Release](https://img.shields.io/badge/Descargar-v1.0.0-blue.svg)](https://github.com/MaizaJoel/pichincha-excel-cleaner/releases)

Una aplicación de escritorio y móvil ligera, independiente y **100% offline** que transforma reportes de movimientos complejos de **Banco Pichincha** (que vienen con filas alternadas, datos desplazados y formatos no tabulares) en una **Tabla oficial de Excel (`ListObject`) con filtros activos por columna, fechas con formato `dd-mm-yyyy hh:mm` y montos formateados como moneda**.

---

> ### ⚡ Aviso de Transparencia y Vibecoding
> **Este proyecto fue desarrollado localmente mediante *vibecoding* con asistencia de IA.**
> - 🛡️ **100% Local y Seguro**: Se ejecuta en tu propio dispositivo con **cero conexión a internet, cero telemetría y cero procesamiento en la nube**. Tus datos y movimientos bancarios nunca salen de tu computadora o teléfono.
> - ⚠️ **Revisión Humana Recomendada**: Aunque el motor de extracción y formateo ha sido rigurosamente probado con pruebas unitarias, los estados de cuenta bancarios pueden cambiar de formato con el tiempo o tener filas no estándar. Siempre verifica los totales con tu estado de cuenta oficial. Si encuentras algún caso no reconocido, ¡abre un Issue en GitHub!

---

## 🎯 ¿Qué problema soluciona?

Al descargar el historial de movimientos de Banco Pichincha:
1. Cada transacción se divide en **filas alternadas/espaciadoras** (el número de documento y el beneficiario aparecen en la fila inferior al concepto y fecha).
2. Las fechas y montos usan formatos con símbolos (`$14.342,22`) que impiden ordenar cronológicamente o hacer fórmulas automáticas como `=SUMA()`.
3. El archivo no es una Tabla oficial de Excel, lo que dificulta filtrar por columnas o buscar transacciones específicas.

**Esta aplicación realiza automáticamente:**
- Unión de filas de metadatos con su transacción correspondiente.
- Eliminación de filas vacías y espaciadoras.
- Formato de fecha uniforme: **`dd-mm-yyyy hh:mm`**.
- Formato de documento como **Texto (`@`)** para no perder ceros a la izquierda.
- Formato de montos y saldos como **Moneda (`$#,##0.00;[Red]-$#,##0.00;"-"`)**.
- Creación de una **Tabla de Excel (`ListObject`)** con botones de autofiltro activos en los encabezados y la primera fila congelada.

---

## 🔒 100% Local, Offline y Privado

| Garantía | Descripción |
| :--- | :--- |
| **Sin peticiones de red** | No se utiliza HTTP, HTTPS, WebSockets ni analíticas. Funciona sin internet. |
| **Sin procesamiento en la nube** | La lectura y generación de Excel ocurren enteramente en la memoria de tu dispositivo. |
| **Sin contraseñas ni accesos** | No necesitas ingresar credenciales bancarias ni claves. |
| **Sin bases de datos externas** | Los archivos se procesan al instante y se guardan directamente donde tú elijas. |

---

## 🚀 Cómo Empezar

### Opción 1: Descargar el Ejecutable para Windows (`.exe`)
¡No requiere instalar Python ni ningún otro programa!
1. Entra a la sección de [Releases](https://github.com/MaizaJoel/pichincha-excel-cleaner/releases/latest).
2. Descarga **`PichinchaExcelCleaner.exe`**.
3. Haz doble clic para abrir la aplicación.

---

### Opción 2: Ejecutar desde el código fuente en Python (Windows / macOS / Linux)

#### 1. Clonar el repositorio
```bash
git clone https://github.com/MaizaJoel/pichincha-excel-cleaner.git
cd pichincha-excel-cleaner
```

#### 2. Instalar dependencias
```bash
pip install -r requirements.txt
```

#### 3. Iniciar la aplicación de escritorio
```bash
python app_gui.py
```

---

### Opción 3: Compilar tu propio ejecutable `.EXE` en Windows
```bash
python build_exe.py
```
El archivo ejecutable se generará en `dist/PichinchaExcelCleaner.exe`.

---

## 📱 Versiones Móviles (Android & iOS)

El repositorio incluye la implementación multiplataforma en **Flutter** dentro de la carpeta `mobile/`:
- **Android**: Compila a APK autónomo para teléfonos y tablets.
- **iOS**: Compila a app nativa para iPhone / iPad.

Para más información técnica, revisa la [Guía Multiplataforma (CROSS_PLATFORM_GUIDE.md)](CROSS_PLATFORM_GUIDE.md).

---

## 🧪 Pruebas Unitarias

Para validar el motor de extracción y generación de Excel:
```bash
pytest test_engine.py -v
```

---

## 📄 Licencia

Este proyecto está distribuido bajo la licencia [MIT](LICENSE).
