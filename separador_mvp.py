import docx
import os
import re
import unicodedata
import fitz  # PyMuPDF para manejar imágenes PDF
from docx.document import Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph

# ==========================================
# FUNCIONES DE EXTRACCIÓN AVANZADA
# ==========================================

def iter_block_items(parent):
    if isinstance(parent, Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("Elemento no soportado")

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

def extraer_imagenes_del_parrafo(parrafo, doc_part, dir_media):
    imagenes_md = ""
    # Ampliamos la red para atrapar formatos modernos, antiguos y formas agruadas
    nodos_imagen = parrafo._element.xpath('.//*[local-name()="blip" or local-name()="imagedata" or local-name()="shape" or local-name()="drawing"]')
    
    for nodo in nodos_imagen:
        rIds = nodo.xpath('.//@*[local-name()="embed" or local-name()="id" or local-name()="link"]')
        
        for rId in rIds:
            if rId in doc_part.related_parts:
                try:
                    image_part = doc_part.related_parts[rId]
                    if 'image' in image_part.content_type:
                        ext = image_part.content_type.split('/')[-1]
                        filename = f"img_docx_{image_part.sha1[:8]}.{ext}"
                        filepath = os.path.join(dir_media, filename)
                        
                        if not os.path.exists(filepath):
                            with open(filepath, 'wb') as f:
                                f.write(image_part.blob)
                                
                        img_str = f"\n\n![](media/{filename})\n\n"
                        if img_str not in imagenes_md:
                            imagenes_md += img_str
                except Exception:
                    pass
                
    return imagenes_md

def extraer_texto_integrado(parrafo):
    """
    Lee el párrafo de izquierda a derecha agrupando fragmentos con el mismo formato.
    Si encuentra texto normal, le aplica negritas/cursivas al final del grupo.
    Si encuentra una ecuación, la extrae en la misma línea.
    """
    texto_md = ""
    current_text = ""
    current_bold = False
    current_italic = False

    def flush_run():
        nonlocal texto_md, current_text
        if not current_text:
            return
            
        txt = current_text
        if not txt.strip():
            texto_md += txt
        else:
            lspace = len(txt) - len(txt.lstrip(' \t\n\r'))
            rspace = len(txt) - len(txt.rstrip(' \t\n\r'))
            
            l_str = txt[:lspace] if lspace else ""
            r_str = txt[-rspace:] if rspace else ""
            clean_txt = txt.strip()
            
            if current_bold and current_italic:
                clean_txt = f"***{clean_txt}***"
            elif current_bold:
                clean_txt = f"**{clean_txt}**"
            elif current_italic:
                clean_txt = f"*{clean_txt}*"
            
            texto_md += l_str + clean_txt + r_str
            
        current_text = ""

    for child in parrafo._element:
        # 1. Si es un bloque de texto normal (w:r)
        if child.tag.endswith('r'):
            run = docx.text.run.Run(child, parrafo)
            txt = run.text if run.text is not None else ""
            
            if not txt:
                continue
                
            if (bool(run.bold) != current_bold) or (bool(run.italic) != current_italic):
                flush_run()
                current_bold = bool(run.bold)
                current_italic = bool(run.italic)
                
            current_text += txt
            
        # 2. Si es una ecuación matemática de Word (m:oMath)
        elif child.tag.endswith('oMath') or child.tag.endswith('oMathPara'):
            flush_run()
            textos = child.xpath('.//*[local-name()="t"]')
            eq_text = "".join([t.text for t in textos if t.text]).replace('\n', ' ').strip()
            
            if eq_text:
                if 'oMathPara' in child.tag:
                    texto_md += f"\n\n$${eq_text}$$\n\n"
                else:
                    texto_md += f" ${eq_text}$ "
                    
    flush_run()
    return texto_md

def tabla_a_markdown(tabla):
    if not tabla.rows: return ""
    md = "\n"
    for i, fila in enumerate(tabla.rows):
        fila_texto = [celda.text.replace('\n', ' ').strip() for celda in fila.cells]
        md += "| " + " | ".join(fila_texto) + " |\n"
        if i == 0:
            md += "| " + " | ".join(['---'] * len(fila.cells)) + " |\n"
    return md + "\n"

# ==========================================
# UTILIDADES Y RENDERIZACIÓN
# ==========================================

def limpiar_nombre_archivo(texto):
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('utf-8')
    limpio = re.sub(r'[^\w\s-]', '', texto).strip()
    return re.sub(r'[-\s]+', '-', limpio).lower()[:45]

def limpiar_titulo(texto):
    """
    Limpia el título eliminando prefijos de códigos, números iniciales y guiones.
    """
    # 1. Quitar extensiones si existen
    texto = re.sub(r'\.docx$', '', texto, flags=re.IGNORECASE)
    # 2. Quitar códigos de carpeta/archivo como C01, LGRD-14, VF-2022...
    texto = re.sub(r'(?i)\b(C\d+|LGRD-?\d+|VF-?\d{4}-\d{2}-\d{2})\b', '', texto)
    # 3. Quitar numeración inicial (ej: "1. ", "2.1-", "1 ")
    # Quitamos números, puntos, guiones y espacios al inicio
    texto = re.sub(r'^[\d\.\-\s]+', '', texto)
    # 4. Reemplazar guiones y guiones bajos por espacios
    texto = texto.replace('-', ' ').replace('_', ' ')
    # 5. Limpiar espacios múltiples y extremos
    texto = re.sub(r'\s+', ' ', texto).strip()
    return texto

def generar_yaml(dir_salida, lista_archivos):
    ruta_yaml = os.path.join(dir_salida, '_quarto.yml')
    with open(ruta_yaml, 'w', encoding='utf-8') as f:
        f.write("project:\n  type: book\n  output-dir: _book\n\n")
        f.write("book:\n  title: \"Investigaciones en Gestión del Riesgo\"\n  sidebar:\n    logo: media/Logo-UNGRD-Horizontal-removebg-preview.png\n  chapters:\n    - index.qmd\n")
        for archivo in lista_archivos: f.write(f"    - {archivo}\n")
        f.write("\nformat:\n  html:\n    theme: \n      - lumen\n      - theme.scss\n    css: styles.css\n    toc: true\n    toc-depth: 4\n    toc-expand: true\n    number-sections: true\n    html-math-method: mathjax\n    reader-mode: true\n    page-navigation: true\n    back-to-top-navigation: true\n")

    ruta_css = os.path.join(dir_salida, 'styles.css')
    with open(ruta_css, 'w', encoding='utf-8') as f:
        f.write("/* CSS para tablas y estilos generales */\n")
        f.write("table {\n  display: block !important;\n  max-width: 100% !important;\n  overflow-x: auto !important;\n}\n")

def convertir_pdf_a_png(ruta_pdf, dir_media):
    """
    Convierte un PDF a imágenes PNG de alta calidad y las guarda en dir_media.
    Retorna una lista con los nombres de los archivos PNG generados.
    """
    imagenes_generadas = []
    try:
        doc = fitz.open(ruta_pdf)
        nombre_base = os.path.splitext(os.path.basename(ruta_pdf))[0]
        # Limpiamos nombre base por seguridad
        nombre_base = re.sub(r'[^\w\s-]', '', nombre_base).replace(' ', '_')
        
        for page_num in range(len(doc)):
            page = doc.load_page(page_num)
            # Zoom para alta resolución (escala x4 aproxima ~300+ DPI)
            mat = fitz.Matrix(4, 4)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            
            if len(doc) == 1:
                filename = f"{nombre_base}.png"
            else:
                filename = f"{nombre_base}_p{page_num+1}.png"
                
            filepath = os.path.join(dir_media, filename)
            pix.save(filepath)
            imagenes_generadas.append(filename)
        doc.close()
    except Exception as e:
        print(f"Error convirtiendo PDF {ruta_pdf}: {e}")
    return imagenes_generadas

# ==========================================
# LÓGICA DE PROCESAMIENTO DIRECTORIO
# ==========================================

def procesar_directorio(dir_maestro, dir_salida):
    print(f"Iniciando extracción por directorios desde: {dir_maestro}")
    
    dir_media = os.path.join(dir_salida, 'media')
    os.makedirs(dir_media, exist_ok=True)
    
    lista_archivos_qmd = []
    
    # Crear index inicial
    index_path = os.path.join(dir_salida, 'index.qmd')
    with open(index_path, 'w', encoding='utf-8') as f:
        f.write("# Preliminares {.unnumbered}\n\nBienvenido al libro de Investigaciones en Gestión del Riesgo.\n")

    try:
        subcarpetas = [d for d in os.listdir(dir_maestro) if os.path.isdir(os.path.join(dir_maestro, d))]
        subcarpetas.sort()
    except Exception as e:
        print(f"Error al leer el directorio base: {e}")
        return

    # Usaremos el número de carpeta para el nombre de los QMD para mantener el orden
    # pero el contador global se usará solo si no se encuentra número en la carpeta
    contador_respaldo = 1
    
    for carpeta in subcarpetas:
        ruta_carpeta = os.path.join(dir_maestro, carpeta)
        print(f"\n--- Procesando carpeta: {carpeta} ---")

        # Detectar si es un capítulo numerado (C01, C02...) o preliminar (C0)
        match_cap = re.search(r'C(\d+)', carpeta, re.IGNORECASE)
        num_carpeta_int = int(match_cap.group(1)) if match_cap else None
        
        prefijo_titulo = ""
        # Solo poner "Capítulo X" si el número es mayor a 0
        if num_carpeta_int is not None and num_carpeta_int > 0:
            prefijo_titulo = f"Capítulo {num_carpeta_int}: "
            num_orden = num_carpeta_int
        else:
            num_orden = 0 if num_carpeta_int == 0 else contador_respaldo
            contador_respaldo += 1
        
        archivos_docx = [f for f in os.listdir(ruta_carpeta) if f.lower().endswith('.docx') and not f.startswith('~')]
        archivos_docx.sort()
        
        archivos_pdf = [f for f in os.listdir(ruta_carpeta) if f.lower().endswith('.pdf')]
        archivos_pdf.sort()
        
        if not archivos_docx:
            print(f"  Saltando carpeta {carpeta}: No se encontraron archivos .docx")
            continue
            
        # 1. Convertir imágenes PDF originales en Alta Calidad a PNG
        pngs_del_capitulo = []
        for pdf_file in archivos_pdf:
            ruta_pdf = os.path.join(ruta_carpeta, pdf_file)
            print(f"  Convirtiendo PDF de alta resolución: {pdf_file}")
            pngs_generados = convertir_pdf_a_png(ruta_pdf, dir_media)
            pngs_del_capitulo.extend(pngs_generados)

        # 2. Procesar cada archivo Word en la carpeta
        for docx_file in archivos_docx:
            ruta_docx = os.path.join(ruta_carpeta, docx_file)
            print(f"  Procesando documento: {docx_file}")
            
            nombre_limpio = limpiar_nombre_archivo(docx_file.replace('.docx', ''))
            # Se usa el numero de capítulo (o el de respaldo) para ordenar los QMD generados
            nom_qmd = f"{num_orden:02d}-{nombre_limpio}.qmd"
            ruta_qmd = os.path.join(dir_salida, nom_qmd)
            lista_archivos_qmd.append(nom_qmd)
            
            with open(ruta_qmd, 'w', encoding='utf-8') as archivo_actual:
                doc = docx.Document(ruta_docx)
                
                # --- PRE-EXTRACCIÓN DEL TÍTULO REAL ---
                titulo_real = ""
                textos_a_omitir = []
                for p in doc.paragraphs:
                    txt = p.text.strip()
                    if not txt: continue
                    # Si es solo un número o algo como "1.", lo omitimos del cuerpo
                    if re.match(r'^[\d\.]+$', txt):
                        textos_a_omitir.append(txt)
                        continue
                    # El primer párrafo con texto sustancial es nuestro título
                    if len(txt) > 10:
                        titulo_real = txt
                        textos_a_omitir.append(txt)
                        break
                
                if not titulo_real:
                    titulo_real = docx_file.replace('.docx', '')
                
                titulo_final = limpiar_titulo(titulo_real)
                # Las preliminares (num_orden == 0) van sin número
                sufijo_unnumbered = " {.unnumbered}" if num_orden == 0 else ""
                
                # Escribir el título principal (quitamos el prefijo "Capítulo X" a pedido del usuario)
                archivo_actual.write(f"# {titulo_final}{sufijo_unnumbered}\n\n")
                
                try:
                    for block in iter_block_items(doc):
                        if isinstance(block, Table):
                            archivo_actual.write(tabla_a_markdown(block))
                            continue
                            
                        parrafo = block
                        texto_base = parrafo.text.strip()
                        
                        # Omitir los textos que ya usamos para el título o que son numeración suelta al inicio
                        if textos_a_omitir and texto_base in textos_a_omitir:
                            textos_a_omitir.remove(texto_base)
                            continue

                        # Extraer imagenes integradas como fallback en caso de no venir en PDF
                        imagenes_md = extraer_imagenes_del_parrafo(parrafo, doc.part, dir_media)
                        texto_md = extraer_texto_integrado(parrafo)
                        
                        if not texto_md.strip() and not imagenes_md: 
                            continue
                            
                        texto_limpio = texto_md.lstrip()
                        if texto_limpio.startswith('['): texto_limpio = '\\' + texto_limpio
                        
                        estilo = parrafo.style.name.lower() if parrafo.style else ''
                        
                        # Manejo de encabezados secundarios
                        if 'heading 1' in estilo or 'título 1' in estilo:
                            # Si por casualidad hay otro Heading 1, lo bajamos de nivel o lo mantenemos como ##
                            archivo_actual.write(f"## {limpiar_titulo(texto_limpio)}\n\n")
                        elif 'heading 2' in estilo or 'título 2' in estilo:
                            archivo_actual.write(f"### {limpiar_titulo(texto_limpio)}\n\n")
                        else:
                            # Heurística para detectar subtítulos que los autores pusieron en negrita en vez de estilo 'Heading'
                            match_falso_encabezado = re.match(r'^\*\*([^\*]+)\*\*$', texto_limpio.strip())
                            es_subtitulo = False
                            if match_falso_encabezado and len(texto_limpio) < 150:
                                txt_interno = match_falso_encabezado.group(1).strip()
                                # Si es todo mayúsculas o empieza con numeral (ej. 1.1 )
                                if txt_interno.isupper() or re.match(r'^\d+[\.\d]*\s+', txt_interno):
                                    # Evitar convertir los nombres de los autores o citas si por error son todo mayúsculas pero muy largos
                                    if len(txt_interno) < 80:
                                        archivo_actual.write(f"## {txt_interno}\n\n")
                                        es_subtitulo = True
                            
                            if not es_subtitulo:
                                # Prevenir que un '#' accidental en el texto genere un nuevo capítulo
                                if texto_limpio.startswith('# '):
                                    texto_limpio = '#' + texto_limpio
                                archivo_actual.write(f"{texto_limpio}\n\n")
                            
                        # Detectar leyenda de figura y agregar la imagen de alta resolución abajo
                        match_fig = re.search(r'\**Fig(?:ura)?s?\s*\**(\d+)', texto_limpio, re.IGNORECASE)
                        if match_fig:
                            fig_num = match_fig.group(1)
                            png_to_remove = None
                            for png in pngs_del_capitulo:
                                if re.search(rf'(?i)fig(?:ura)?s?[-_\s]*0*{fig_num}.*\.png$', png):
                                    archivo_actual.write(f"![Imagen {png}](media/{png})\n\n")
                                    png_to_remove = png
                                    break
                            if png_to_remove:
                                pngs_del_capitulo.remove(png_to_remove)
                
                except Exception as e:
                    print(f"  Error procesando {docx_file}: {e}")

        # 3. Anexar las imágenes originales convertidas desde PDF (Alta Calidad) al final del último docx del capítulo
        # Comentado por solicitud para no listar las imágenes al final
        # if pngs_del_capitulo and lista_archivos_qmd:
        #     ultimo_qmd = os.path.join(dir_salida, lista_archivos_qmd[-1])
        #     with open(ultimo_qmd, 'a', encoding='utf-8') as f:
        #         f.write("\n## Figuras y Anexos del Capítulo (Alta Resolución)\n\n")
        #         f.write("A continuación se presentan las imágenes originales en alta resolución correspondientes a este capítulo para asegurar la mejor calidad visual en todos los navegadores.\n\n")
        #         for png in pngs_del_capitulo:
        #             f.write(f"![Imagen {png}](media/{png})\n\n")

    # Generar el archivo final para Quarto
    generar_yaml(dir_salida, lista_archivos_qmd)
    print(f"\n¡Procesamiento finalizado! Se han generado {len(lista_archivos_qmd)} archivos en '{dir_salida}'.")

if __name__ == "__main__":
    directorio_entrada = '6.1-ARCHIVOS-MAESTROS-LGRD22'
    carpeta_salida = 'proyecto_libro_quarto'
    
    if os.path.exists(directorio_entrada):
        procesar_directorio(directorio_entrada, carpeta_salida)
    else:
        print(f"El directorio maestro '{directorio_entrada}' no se encontró en la ruta actual.")