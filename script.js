let workbookGlobal;  /** Variable global para trabajar con Excel */

/** Funcion 1 */
/** Valida la estructura de un archivo JSON. Revisa si hay campos vacíos, nulos o indefinidos en los datos y muestra advertencias en la consola si encuentra alguna irregularidad. */
function validateJSONStructure(jsonData) {
  let hasEmptyFields = false;  /**Declara una variable para rastrear si hay campos vacíos en los datos JSON */
  jsonData.forEach((row, rowIndex) => {  /**Recorre cada fila del array JSON */
    for (const [key, value] of Object.entries(row)) {  /**Itera sobre las claves y valores de cada fila del JSON */
      if (value === null || value === '' || value === undefined) { /**Verifica si el valor de la clave es nulo, vacío o undefined */
        console.warn(`Fila ${rowIndex + 1}, columna "${key}" está vacía.`); /**Si se encuentra un campo vacío, muestra un mensaje de advertencia en la consola */
        hasEmptyFields = true;
      }
    }
  });

  if (!hasEmptyFields) { /**verifica si la funcion hasEmptyField es falsa, lo caual indica que no hay campos vacios*/
    console.log('Todos los campos del JSON tienen valores.'); /**Si no tiene campos vacios genera este ensaje en cosola*/
  } 
}

/** Funcion 2 */
/** Carga un archivo de Excel desde una URL, convierte los datos a formato workbook, y llena un selector con las hojas disponibles del archivo. Muestra una notificación de éxito. */
function handleExcelLoad() {
  /** Variables */
  const uploadExcel = document.getElementById('uploadExcel');   /** Obtiene el documento HTML con el Id uploadExcel */
  const sheetSelector = document.getElementById('sheetSelector'); /** Selector para seleccionar las hojas del archivo cargado */

  uploadExcel.style.display = 'inline-block'; /** Muestra en pantalla el Upload excel */


  //**ACTUALIZAR O SUBIR CATALOGO */
  const fileName = 'Datexce/Catálogo actualizado 12 de nov.xlsx'; /** Constante con url del catalogo */
  
  fetch(fileName)  // ** Realiza solicitud para cargar archivo */
    .then(response => { 
      // ** Si la solicitud falla genera error */
      if (!response.ok) throw new Error('El archivo JSON no está disponible o tiene un formato incorrecto');
      return response.arrayBuffer();   /** Convierte el archivo en array */
    })
    .then(data => {    
      const workbook = XLSX.read(data, { type: 'array' });  /** Lee el archivo excel ya convertido con la libreria Xlsx */
      
      const extractedFileName = fileName.split('/').pop();  /** Extrae el nombre del archivo separando la frase y tomando la ultima parte */

      /* Envía mensaje de archivo cargado */
      Swal.fire({
        title: '!Archivo cargado!',
        text: `El archivo "${extractedFileName}" se cargó correctamente`,
        icon: 'success',
        showConfirmButton: false,
        timer: 2500,
      });
      

      workbookGlobal = workbook; /** Guarda en la variable global, los objetos convertidos */
      
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const jsonData = XLSX.utils.sheet_to_json(sheet); // Convierte la hoja a JSON
  
      validateJSONStructure(jsonData); // Llama la validación

      sheetSelector.style.display = 'inline-block'; /* Hace visible en pantalla el selector de hojas */
      sheetSelector.innerHTML = '<option value="">Selecciona un Producto</option>'; /* Selector que permite elegir una hoja */

      workbook.SheetNames.forEach(function (sheetName, index) {  /** Recorre las hojas del excel y agrega una opción para cada una */
        /** Crea una opción con índice y nombre de la hoja y lo agrega en el selector */
        const option = document.createElement('option');
        option.value = index;
        option.text = sheetName;
        sheetSelector.appendChild(option);
      });
    })
    /** Genera un mensaje de error en caso de no completar correctamente el proceso */
    .catch(error => {
      console.error('Error al cargar el archivo:', error); /**Muestra el error en Consola */

      swal.fire({ /**Muestra el error en pantalla con un alert */
        
        title: 'Error al cargar el archivo',
        text: error.message, 
        icon: 'error',
        showConfirmButton: true,
        confirmButtonText: 'Cómo cargar el archivo correctamente',
        confirmButtonColor: 'red',
        timer:10000,
    }).then((result) => {
        if (result.isConfirmed) {
          
            const url = 'Datexce/Instrucciones.text';
            const tabName = 'instructionsTab'; 
            window.open(url, tabName);  /**No abre nuevamnete la pestaña del manual si ya esta abierta */
        }
    });
    
});
}

/** Manejador de evento para la carga de archivos */
document.getElementById('uploadExcel').addEventListener('change', (event) => {
  const file = event.target.files[0];   /** Almacena y muestra el número de hojas del archivo seleccionado, en una lista */

  if (file) {  /** Verifica si selecciona un archivo */
    console.log(`Archivo seleccionado: ${file.name}`);  /** Si selecciona un archivo muestra en consola el nombre */
  }
});

/** Funcion 3 */
/** Maneja la carga de un archivo de Excel desde un input de tipo archivo, lo lee y convierte a un formato workbook, luego llena un selector con las hojas disponibles. */
function handleFile(e) {  /** Función con un parámetro */
  const file = e.target.files[0];    /** Constante que almacena el archivo para ejecutarlo más adelante, archivo que se muestra en una lista */
  if (!file) return;                 /** Condición que comprueba si hay un archivo cargado, si no hay archivo la condición sería verdadera y detiene la ejecución del código */

  document.getElementById('uploadExcel').style.display = 'none';  /** Oculta de la pantalla elemento con "uploadExcel" */

  /** Se encarga de convertir los datos del archivo cargado, para poder ser leídos por el navegador, guardados en el libro global y poder ser llamados más adelante */
  const reader = new FileReader();  /** Constante con API FileReader, para leer archivos cargados */
  
  reader.onload = function (e) {  /** Al archivo cargarse, cumple la función asignada */
    const data = new Uint8Array(e.target.result);  /** Guarda en datos binarios el resultado de la lectura del archivo */
    const workbook = XLSX.read(data, { type: 'array' });  /** Constante con librería Xlsx para convertir los datos en workbook */
    
    workbookGlobal = workbook;  /** Guarda los datos convertidos en la variable local */
    
    
   

    const sheetSelector = document.getElementById('sheetSelector');  /** Selecciona el elemento HTML con el Id para mostrar las hojas del archivo */
    sheetSelector.style.display = 'inline-block';  /** Hace que el selector sea visible en pantalla */
    sheetSelector.innerHTML = '<option value="">Selecciona un producto</option>';  /** Muestra el selector con la opción de seleccionar producto */
    
    workbook.SheetNames.forEach(function (sheetName, index) {  /** Hace llamado al archivo convertido guardado en el libro global y acceder a sus hojas */
      /** Crea la opción de seleccionar cada hoja del archivo, mostrando su nombre y contenido */
      const option = document.createElement('option');
      option.value = index;  /** Lista las hojas */
      option.text = sheetName;  /** Nombra las hojas */
      sheetSelector.appendChild(option);
    });
  };

  reader.readAsArrayBuffer(file);  /** El método 'readAsArrayBuffer' se utiliza cuando se necesita acceder a los datos binarios del archivo, como imágenes o archivos en formatos binarios. */
}

/** Funcion 4 */
/** Verifica si un string es una imagen en formato base64. Devuelve true si es una imagen, de lo contrario, false. */
function isBase64Image(base64String) {  /** Función que verifica si un string es una imagen en lenguaje base64 */
  return typeof base64String === "string" && base64String.startsWith("data:image/");
}

/** Funcion 5 */
/** Obtiene el índice de la hoja seleccionada en un selector y carga los datos correspondientes de esa hoja en la interfaz. */
function loadSheet() {                      
  const sheetIndex = document.getElementById('sheetSelector').value;

  /** Si el selector está vacío no genera salida */
  if (sheetIndex === '') {
    document.getElementById('output').innerHTML = '';
    return;  /** Si no hay nada seleccionado, termina la función */
  }

  const sheetName = workbookGlobal.SheetNames[sheetIndex];  /** Con el número de la hoja obtiene el nombre */
  const sheet = workbookGlobal.Sheets[sheetName];  /** Obtiene la hoja del libro global */
  const sheetRange = XLSX.utils.decode_range(sheet['!ref']);  /** Utiliza la librería para obtener las filas y las columnas del archivo */
  createCardsFromExcel(sheet, sheetRange);  /** Llama una función con la hoja, las filas y las columnas */
}

/** Funcion 6 */
/** Crea tarjetas HTML a partir de los datos de una hoja de Excel. Cada tarjeta representa un producto con su nombre, precios e imágenes. Además, genera un botón para interactuar con el contenido. */
function createCardsFromExcel(sheet, data) {  /** Función para crear tarjetas a partir de los datos de Excel */
  const output = document.getElementById('output');  /** Contenedor donde se mostrarán las carpetas */
  output.innerHTML = '';  /** Limpia el contenedor */
  
  let rowHtml = '<div class="row">';  /** Organiza las tarjetas en el contenedor */
  let cardCount = 0;  /** Contador para las Tarjetas */

  for (let rowNum = data.s.r + 1; rowNum <= data.e.r; rowNum++) {
    const productName = sheet[XLSX.utils.encode_cell({ r: rowNum, c: 1 })]; 
    const productValue = productName ? productName.v : 'Sin nombre'; 

    const imageName = sheet[XLSX.utils.encode_cell({ r: rowNum, c: 7 })];  // Suponiendo que la imagen está en la columna 7 (Índice H)
    const imageUrl = imageName ? `img/${imageName.v}` : 'https://via.placeholder.com/150';  // Ruta de la imagen
    

    //**Contenedor, Diseño y contenido */
    rowHtml += `  
      <div class="col-md-4 mt-3">
        <div class="card" style="width: 18rem;">
          <img src="${imageUrl}" class="card-img-top" alt="Imagen de producto">
          <div class="card-body"> 
            <h5 class="card-title">${productValue}</h5>
            <p class="card-text" id="cardText${rowNum}">
    `; 

    let pricesHtml = '<div style="line-height: 1.5;">';  /**Variable Precios, para crear su contenido */
    for (let colNum = 1; colNum <= data.e.c; colNum++) {  /**Bucle que recorre todas las columnas de del excel hasta la ultima que contenga datos */
      const cellAddress = XLSX.utils.encode_cell({ r: rowNum, c: colNum }); /**Codifica la direccion de la celda, para acceder a la correspondiente  */
      const cell = sheet[cellAddress]; /**Accede a la celda codificada */
      const cellValue = cell ? cell.v : ''; /**Extrae valor de la celda si existe; de lo contrario, asigna una cadena vacia*/

      if ([4, 6].includes(colNum) && !isNaN(cellValue)) { 
        const formattedPrice = parseFloat(cellValue).toLocaleString('es-CO', { style: 'currency', currency: 'COP' }); /**Convierte el valor de la celda a moneda colombiana */
        pricesHtml += `<strong style="color: green;">${formattedPrice}</strong><br>`;  /**Agrega el precio ya convertido; con un color verde */
      } else if (colNum === 5 && !isNaN(cellValue)) {  /**verifica si la columna es la 5  y si el valor es numero*/
        const ivaPercentage = (cellValue * 100).toFixed(0); /**Multiplica el valor * 100 y redondea el valor a entero */
        pricesHtml += `<strong style="color: orange;">IVA ${ivaPercentage}%</strong><br>`; /**Agrega el IVA al contenedor en color naranja */
      } else if (isBase64Image(cellValue)) {  /**Verifica si el valor de la celda es base 64 */
        rowHtml += `<img src="${cellValue}" width="100" class="mb-2"/>` /**Agrega la inagen a la targeta */
      } else {
        rowHtml += `${cellValue} `; /**Si no se cumple las caracteristicas anteriores agrega el valor de la celda */
      }
    }

    pricesHtml += '</div>';
    rowHtml += `${pricesHtml}</p>
      <span class="show-more" id="showMoreBtn${rowNum}">Leer más</span>
      <a href="https://wa.me/573163615434" class="btn btn-primary mt-2 btn-whatsapp" target="_blank">Hablar con un Asesor</a>  
    </div>
  </div>
</div>`; /**Boton "leer mas" donde despliega mas inforacion sobre el producto*/ /**Enlace a whataApp para comunicarse con un asesor */

    cardCount++; /**Contador de tarjetas*/

    if (cardCount % 3 === 0) {   
      rowHtml += '</div><div class="row">';
    } /**Condicionale que hace que se agrupen de a tres tarjetas, si son mas de 3.*/
  }

  rowHtml += '</div>';
  output.innerHTML = rowHtml;

  document.querySelectorAll('.show-more').forEach((button, index) => { /** Selecciona todos los elementos que tienen la clase "show-more" y recorre cada uno */
    button.addEventListener('click', function() { /** Agrega un evento click a cada botón seleccionado */
      const cardText = document.getElementById(`cardText${index + 1}`); /**Obtiene el elemento cuyo ID corresponde a "cardText" seguido por el índice  */
      cardText.classList.toggle('expanded');
      button.textContent = cardText.classList.contains('expanded') ? 'Leer menos' : 'Leer más'; /**Cambia el texto del botón según el estado de la clase "expanded" en "cardText" */
    });
  });
}



/** Funcion 7 */
/** Al cargar la página, ejecuta la función handleExcelLoad() para iniciar el proceso de carga de un archivo Excel. */
window.onload = function() {
  handleExcelLoad();
};
