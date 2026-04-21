# Guía de Uso: Automatización de Libro Quarto

Esta guía detalla los pasos necesarios para configurar el entorno y ejecutar los scripts de procesamiento para convertir documentos Word en un libro digital interactivo utilizando Quarto.

---

## 1. Requisitos Previos

Antes de comenzar, asegúrate de tener instalado lo siguiente en tu sistema:

1.  **Python 3.8+**: El lenguaje de programación base.
2.  **Quarto CLI**: La herramienta para generar el libro. Puedes descargarla en [quarto.org](https://quarto.org/docs/get-started/).
3.  **Visual Studio Code** (Recomendado): Con las extensiones de Quarto y Python.

---

## 2. Configuración del Entorno (Entorno Virtual)

Para evitar conflictos entre versiones de software, utilizamos un "entorno virtual" (venv). Sigue estos pasos en tu terminal (Linux/macOS):

### Paso 1: Crear el entorno virtual
Si aún no existe la carpeta `venv`, créala con:
```bash
python3 -m venv venv
```

### Paso 2: Activar el entorno virtual
Debes activarlo **cada vez** que abras una nueva terminal para trabajar en el proyecto:
```bash
source venv/bin/activate
```
*(Sabrás que está activo porque aparecerá `(venv)` al inicio de tu línea de comandos).*

### Paso 3: Instalar dependencias
Instala las librerías necesarias (solo la primera vez):
```bash
pip install -r requirements.txt
```

---

## 3. Flujo de Trabajo Paso a Paso

El proceso se divide en tres etapas principales:

### Etapa 1: Separación de Documentos (`separador_mvp.py`)
Este script toma los archivos Word de la carpeta maestra y genera los archivos `.qmd` estructurados.

1.  Asegúrate de que tus archivos Word estén en la carpeta `6.1-ARCHIVOS-MAESTROS-LGRD22` (organizados por carpetas C0, C01, C02, etc.).
2.  Ejecuta el script:
    ```bash
    python3 separador_mvp.py
    ```
3.  **Resultado**: Se creará o actualizará la carpeta `proyecto_libro_quarto` con los archivos `.qmd`, el archivo `_quarto.yml` y una carpeta `media` con las imágenes.

### Etapa 2: Estilización de Cajas (`process_cajas.py`)
Este script busca las tablas de "Caja" y las convierte en bloques visuales modernos (estilo azul con bordes).

1.  Entra en la carpeta del proyecto generado:
    ```bash
    cd proyecto_libro_quarto
    ```
2.  Ejecuta el script de procesamiento:
    ```bash
    python3 process_cajas.py
    ```
3.  **Resultado**: Los archivos `.qmd` ahora tendrán un formato de caja profesional en lugar de tablas simples.

### Etapa 3: Visualización del Libro (`quarto preview`)
Para ver cómo está quedando el libro en tiempo real:

1.  Desde la carpeta `proyecto_libro_quarto`, ejecuta:
    ```bash
    quarto preview
    ```
2.  Se abrirá una ventana en tu navegador. Cada cambio que guardes en los archivos se reflejará automáticamente allí.

---

## 4. Comandos de Mantenimiento

### Limpiar archivos antiguos
Si quieres empezar de cero y borrar todo lo generado anteriormente:
```bash
rm -rf proyecto_libro_quarto/*.qmd proyecto_libro_quarto/media/*
```

### Generar la versión final (Producción)
Cuando el libro esté listo para publicarse:
```bash
quarto render
```
Esto generará una carpeta `_book` con el sitio web completo listo para subir a un servidor.

---

## 5. Solución de Problemas Comunes

*   **Error "ModuleNotFoundError"**: Asegúrate de haber activado el entorno virtual (`source venv/bin/activate`).
*   **Imágenes no aparecen**: Verifica que la carpeta `media` dentro de `proyecto_libro_quarto` contenga los archivos `.png` o `.jpg`. El script `separador_mvp.py` se encarga de extraerlas automáticamente.
*   **Error en Quarto**: Si el comando `quarto` no funciona, verifica que lo hayas instalado y reiniciado la terminal.
