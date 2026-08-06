# Portal automático de servicios

Este repositorio genera y publica un catálogo web a partir de la versión más reciente del Excel maestro cuyo nombre comience por:

```text
datos/Servicios*.xlsx
```

El diseño y la lógica siguen el mismo patrón del portal de oportunidades suministrado como ejemplo: los datos se incrustan dentro de `index.html`, el generador reemplaza únicamente ese bloque y GitHub Pages publica una página completamente estática.

## Estructura del repositorio

```text
.
├── index.html
├── datos/
│   └── ServiciosV1.xlsx
├── scripts/
│   └── generar_portal.py
├── .github/
│   └── workflows/
│       └── actualizar-portal.yml
├── requirements.txt
├── .gitignore
├── .nojekyll
└── README.md
```

## Qué hace automáticamente

Cada vez que se confirma en la rama `main` un archivo que coincida con `datos/Servicios*.xlsx`, GitHub Actions:

1. instala Python y `openpyxl`;
2. identifica la versión más reciente del Excel usando orden natural;
3. lee la hoja `servicios` en modo de solo lectura;
4. valida los diez encabezados requeridos;
5. regenera únicamente el bloque de datos dentro de `index.html`;
6. confirma y sube el `index.html` actualizado;
7. empaqueta la página como artefacto de GitHub Pages;
8. publica directamente la nueva versión del portal en la misma ejecución.

Por ejemplo, `ServiciosV10.xlsx` se selecciona después de `ServiciosV9.xlsx`. También puede sustituirse un archivo conservando el mismo nombre.

## Publicar por primera vez en GitHub

1. Cree un repositorio nuevo en GitHub.
2. Suba **todo el contenido de esta carpeta**, incluida la carpeta oculta `.github`.
3. Confirme los archivos en la rama `main`.
4. Abra **Settings → Actions → General**.
5. En **Workflow permissions**, seleccione **Read and write permissions** y guarde.
6. Abra **Settings → Pages**.
7. En **Build and deployment → Source**, seleccione **GitHub Actions**.
8. Regrese a la pestaña **Actions** y ejecute **Actualizar portal desde Excel** mediante **Run workflow**, o confirme una nueva versión del Excel.
9. Al finalizar, GitHub mostrará la dirección pública del portal.

La URL tendrá normalmente esta forma:

```text
https://USUARIO.github.io/NOMBRE-DEL-REPOSITORIO/
```

## Incorporar una nueva versión del Excel

1. Mantenga la hoja con el nombre `servicios`.
2. Guarde la nueva versión dentro de la carpeta `datos`, por ejemplo:

   ```text
   datos/ServiciosV2.xlsx
   ```

3. Confirme el archivo en la rama `main`.
4. Abra la pestaña **Actions** para observar la ejecución de **Actualizar portal desde Excel**.
5. Cuando finalice, GitHub Pages mostrará los datos nuevos.

Puede conservar versiones anteriores. El generador elegirá la de numeración más alta. Para evitar ambigüedades, use nombres consecutivos como `ServiciosV1.xlsx`, `ServiciosV2.xlsx` y `ServiciosV3.xlsx`.

## Encabezados obligatorios

La hoja `servicios` debe conservar el significado de estas diez columnas:

1. `Laboratorio`
2. `Servicio`
3. `Requiere Equipo del Herbario`
4. `Descripcion y Alcance`
5. `Tipo de usuario al que podría dirigirse.`
6. `Equipos requeridos para su prestación.`
7. `Materiales, reactivos o insumos necesarios.`
8. `Personal técnico o especializado requerido`
9. `Tiempo estimado para la realización del servicio.`
10. `Cualquier requisito, limitación o condición especial que deba considerarse para el arrendamiento del servicio.`

El generador tolera diferencias en acentos, mayúsculas, signos, espacios y saltos de línea de los encabezados. Las filas completamente vacías se ignoran.

### Columna opcional `Publicar`

Puede añadirse una columna llamada `Publicar`. Cuando existe, solamente se incluyen las filas que contengan `Sí`, `Si`, `1`, `X` o `True`. Si la columna no existe, se publican todas las filas válidas.

## Ejecución local

Desde la raíz del proyecto:

```bash
python -m pip install -r requirements.txt
python scripts/generar_portal.py
python -m http.server 8000
```

Luego abra:

```text
http://localhost:8000
```

Para forzar un Excel específico:

```bash
python scripts/generar_portal.py --excel datos/ServiciosV1.xlsx
```

## Archivos que deben editarse

- `datos/ServiciosV*.xlsx`: fuente maestra.
- `index.html`: diseño y comportamiento visual.
- `scripts/generar_portal.py`: reglas de lectura y transformación.

No elimine del HTML estos marcadores:

```javascript
/* BEGIN_GENERATED_DATA */
/* END_GENERATED_DATA */
```

El generador los utiliza para reemplazar los datos sin modificar el resto de la página.

## Protección del Excel

El proceso abre el libro con `read_only=True` y `data_only=True`. No guarda ni reescribe el archivo Excel, por lo que no altera fórmulas, formatos, validaciones de datos ni listas desplegables.


## Publicación de GitHub Pages

El workflow usa las acciones oficiales:

- `actions/configure-pages@v5`
- `actions/upload-pages-artifact@v3`
- `actions/deploy-pages@v4`

Por ello, el origen de publicación debe configurarse como **GitHub Actions**, no como despliegue desde una rama. El sitio se publica dentro del mismo workflow que regenera el HTML.
