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

    capitulo_num = 1
    
    for carpeta in subcarpetas:
        ruta_carpeta = os.path.join(dir_maestro, carpeta)
        print(f"\n--- Procesando carpeta: {carpeta} ---")
        
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
            # Se usa el numero de capítulo para ordenar los QMD generados
            nom_qmd = f"{capitulo_num:02d}-{nombre_limpio}.qmd"
            ruta_qmd = os.path.join(dir_salida, nom_qmd)
            lista_archivos_qmd.append(nom_qmd)
            
            with open(ruta_qmd, 'w', encoding='utf-8') as archivo_actual:
                titulo_capitulo = docx_file.replace('.docx', '')
                titulo_encontrado = False
                
                try:
                    doc = docx.Document(ruta_docx)
                    for block in iter_block_items(doc):
                        if isinstance(block, Table):
                            archivo_actual.write(tabla_a_markdown(block))
                            continue
                            
                        parrafo = block
                        texto_base = parrafo.text.strip()
                        
                        # Extraer imagenes integradas como fallback en caso de no venir en PDF
                        imagenes_md = extraer_imagenes_del_parrafo(parrafo, doc.part, dir_media)
                        texto_md = extraer_texto_integrado(parrafo)
                        
                        if not texto_md.strip() and not imagenes_md: 
                            continue
                            
                        texto_limpio = texto_md.lstrip()
                        if texto_limpio.startswith('['): texto_limpio = '\\' + texto_limpio
                        
                        estilo = parrafo.style.name.lower() if parrafo.style else ''
                        
                        # Manejo del título del capítulo
                        if not titulo_encontrado:
                            if 'heading 1' in estilo or 'título 1' in estilo or texto_limpio.startswith('# '):
                                texto_titulo = texto_limpio.lstrip('#* \t')
                                # Remover numero inicial si lo tiene para evitar redundancia
                                texto_titulo = re.sub(r'^[\d\.\-\s]*', '', texto_titulo)
                                archivo_actual.write(f"# Capítulo {capitulo_num}: {texto_titulo}\n\n")
                                titulo_encontrado = True
                                continue
                            elif texto_limpio and not texto_limpio.startswith('**') and len(texto_limpio) > 20:
                                archivo_actual.write(f"# Capítulo {capitulo_num}: {titulo_capitulo}\n\n")
                                titulo_encontrado = True
                                # No hacer continue para que se imprima este párrafo abajo
                        
                        # Estilos de encabezado
                        if 'heading 1' in estilo or 'título 1' in estilo:
                            archivo_actual.write(f"## {texto_limpio}\n\n")
                        elif 'heading 2' in estilo or 'título 2' in estilo:
                            archivo_actual.write(f"### {texto_limpio}\n\n")
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
                                
                        # if imagenes_md:
                        #     archivo_actual.write(imagenes_md)
                
                except Exception as e:
                    print(f"  Error procesando {docx_file}: {e}")

            capitulo_num += 1

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