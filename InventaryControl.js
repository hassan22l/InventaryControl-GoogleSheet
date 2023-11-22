//Guardado de productos nuevos

function guardarProductos() {

  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hojaProductos = libro.getSheetByName("Ingr.Productos")

  var codigo = hojaProductos.getRange("C3").getValue();
  var nombre = hojaProductos.getRange("C5").getValue();
  var descripcion = hojaProductos.getRange("E3").getValue();
  var precio = hojaProductos.getRange("G3").getValue();
  var precioVenta = hojaProductos.getRange("I3").getValue();
  var equivalencias = hojaProductos.getRange("E5").getValue();
  var lugar = hojaProductos.getRange("G5").getValue();
  var cantidad = hojaProductos.getRange("I5").getValue();
  var gananciafija = hojaProductos.getRange("H7").getValue();
  var proveedor = hojaProductos.getRange("C7").getValue();
  var tipoProducto = hojaProductos.getRange("E7").getValue();


  var hojaProductosInfo = libro.getSheetByName("Productos.Info")
  hojaProductosInfo.appendRow([codigo, nombre, tipoProducto, descripcion, precio, precioVenta, equivalencias, lugar, gananciafija, proveedor])

  var hojaStock = libro.getSheetByName("Stock")
  var dataStock = hojaStock.getDataRange().getValues();

  for (var i = 1; i < dataStock.length; i++) {
    if (dataStock[i][0] === codigo) {
      var stockActual = dataStock[i][4];
      var nuevoStock = stockActual + cantidad;
      hojaStock.getRange(i + 1, 5).setValue(nuevoStock);
      return;
    }
  }


};

function borrarProducto() {

  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hojaProductos = libro.getSheetByName("Ingr.Productos")

  hojaProductos.getRange("C3").clearContent();
  hojaProductos.getRange("C5").clearContent();
  hojaProductos.getRange("E3").clearContent();
  hojaProductos.getRange("G3").clearContent();
  hojaProductos.getRange("E5").clearContent();
  hojaProductos.getRange("G5").clearContent();
  hojaProductos.getRange("I5").clearContent();


};



function guardarEntrada() {

  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hojaCompras = libro.getSheetByName("Compras")
  var hojaHistorial = libro.getSheetByName("Info.Compras")

  var codigo = hojaCompras.getRange("C5").getValue();
  var nombre = hojaCompras.getRange("C7").getValue();
  var cantidad = hojaCompras.getRange("E6").getValue();
  var precio = hojaCompras.getRange("F5").getValue();
  var precioVenta = hojaCompras.getRange("H5").getValue();
  var proveedor = hojaCompras.getRange("D3").getValue();
  var fecha = hojaCompras.getRange("G3").getValue();

  hojaHistorial.appendRow([codigo, nombre, cantidad, precio, precioVenta, proveedor, fecha])

 // Actualización de Costo por porcentaje en Hoja de Productos

      var hojaCompras = libro.getSheetByName("Compras");
      var hojaProductos = libro.getSheetByName("Productos.Info")
      var productos = hojaProductos.getDataRange().getValues();
      var aumento = hojaCompras.getRange("F7").getValue();

    if (isNaN(aumento) || aumento === "") {
        return;
      }

    for(var i = 1; i <productos.length; i++) {
        var marca = productos[i][9]
        var marcaACambiar = hojaCompras.getRange("D3").getValue()

          if (marcaACambiar === marca){
            
              var precioActual = productos[i] [4];
              var nuevoPrecio = precioActual * (1 + aumento / 100);
              hojaProductos.getRange(i+ 1, 5).setValue(nuevoPrecio);
          }
      }

     var hojaStock = libro.getSheetByName("Stock")
      var dataStock = hojaStock.getDataRange().getValues();

      for (var i = 1; i < dataStock.length; i++) {
        if (dataStock[i][0] === codigo) {
          var stockActual = dataStock[i][4];
          var nuevoStock = stockActual + cantidad;
          hojaStock.getRange(i + 1, 5).setValue(nuevoStock);
          return;
        }
      }

};

function borrarCompra() {

  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hojaCompras = libro.getSheetByName("Compras")

  hojaCompras.getRange("C5").clearContent();
  hojaCompras.getRange("C7").clearContent();
  hojaCompras.getRange("F5").clearContent();
  hojaCompras.getRange("E6").clearContent();
}

//Guaradado de ventas en historial de ventas y actualizacion de stock

function guardarVenta() {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hojaVentas = libro.getSheetByName("Ventas")
  var hojaHistorial = libro.getSheetByName("Info.Ventas")
  var hojaStock = libro.getSheetByName("Stock")

  var codigo = hojaVentas.getRange("D3").getValue();
  var nombre = hojaVentas.getRange("D4").getValue();
  var equivlencias = hojaVentas.getRange("G3").getValue();
  var descripcion = hojaVentas.getRange("G4").getValue();
  var cantidad = hojaVentas.getRange("G6").getValue();

  hojaHistorial.appendRow([codigo, nombre, equivlencias, descripcion, cantidad])

  var hojaStock = libro.getSheetByName("Stock")
  var dataStock = hojaStock.getDataRange().getValues();

  for (var i = 1; i < dataStock.length; i++) {
    if (dataStock[i][0] === codigo) {
      var stockActual = dataStock[i][5];
      var nuevoStock = stockActual + cantidad;
      hojaStock.getRange(i + 1, 6).setValue(nuevoStock);
      return;
    }
  }
}

//buscar producto

function buscarProducto() {
  var hojaProductos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Productos.Info');
  var hojaBusqueda = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Buscar');
  var valorBuscado = hojaBusqueda.getRange('B2').getValue(); 
  var datosProductos = hojaProductos.getDataRange().getValues();

  for (var i = 0; i < datosProductos.length; i++) {
    var codigo = datosProductos[i][0]; 
    var nombre = datosProductos[i][1]; 

    if (codigo.toString().toLowerCase() === valorBuscado.toString().toLowerCase() ||
      nombre.toString().toLowerCase() === valorBuscado.toString().toLowerCase()) {
      var precio = datosProductos[i][4]; 
      var equivalencia = datosProductos[i][5]; 
      var descripcion = datosProductos[i][2]; 

      hojaBusqueda.getRange('E2').setValue(codigo);
      hojaBusqueda.getRange('E3').setValue(nombre);
      hojaBusqueda.getRange('E7').setValue(precio);
      hojaBusqueda.getRange('E5').setValue(equivalencia);
      hojaBusqueda.getRange('E4').setValue(descripcion);

      return; 
    }
  }
  hojaBusqueda.getRange('E2:E6').clearContent(); 
  hojaBusqueda.getRange('E2').setValue('Producto no encontrado');
}


function buscarProductoPorFiltro() {
  var hojaProductos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Productos.Info');
  var hojaBusqueda = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Buscar');

  var valorBuscado = hojaBusqueda.getRange('B4').getValue(); // Supongamos que el valor de búsqueda se encuentra en la celda A1 de la hoja "Buscar Productos".

  // Obtén los datos de la hoja de productos.
  var datosProductos = hojaProductos.getDataRange().getValues();

  for (var i = 0; i < datosProductos.length; i++) {
    var auto = datosProductos[i][3]; 

    if (auto.toString().toLowerCase() === valorBuscado.toString().toLowerCase()) {
      // Si se encuentra una coincidencia, muestra los datos en la hoja de búsqueda.
      var codigo = datosProductos[i][0];
      var nombre = datosProductos[i][1];
      var precio = datosProductos[i][5]; // Supongamos que el precio está en la quita columna.
      var equivalencia = datosProductos[i][6]; // Supongamos que la equivalencia está en la sexta columna.
      var descripcion = datosProductos[i][3]; // Supongamos que la descripción está en la tercera columna.

      // Llena los datos en la hoja de búsqueda.
      hojaBusqueda.getRange('D2').setValue(codigo);
      hojaBusqueda.getRange('D3').setValue(nombre);
      hojaBusqueda.getRange('D7').setValue(precio);
      hojaBusqueda.getRange('D5').setValue(equivalencia);
      hojaBusqueda.getRange('D4').setValue(descripcion);

      return; // Sal del bucle una vez que se haya encontrado una coincidencia.
    }
  }

  // Si no se encontró una coincidencia, puedes mostrar un mensaje de error.
  hojaBusqueda.getRange('D2:D6').clearContent(); // Borra los datos anteriores.
  hojaBusqueda.getRange('B5').setValue('Producto no encontrado');
}
