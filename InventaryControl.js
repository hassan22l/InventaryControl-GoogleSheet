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
  var hojaProductosInfo = libro.getSheetByName("Productos.Info")
  hojaProductosInfo.appendRow([codigo, nombre, descripcion, precio, precioVenta, equivalencias, lugar])

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

  // Actualización de Costo y Precio de Venta en Hoja de Productos

  var hojaProductosInfo = libro.getSheetByName("Productos.Info")
  var dataProductos = hojaProductosInfo.getDataRange().getValues();
  var costoYVentaActualizados = false;

  for (var i = 1; i < dataProductos.length; i++) {
    if (dataProductos[i][0] === codigo) {
      if (precio !== 0 && precio !== "") { // Solo actualiza si el precio no es igual a 0 ni está vacío
        if (dataProductos[i][1] !== precio || dataProductos[i][2] !== precioVenta) {
          hojaProductosInfo.getRange(i + 1, 4).setValue(precio);
          hojaProductosInfo.getRange(i + 1, 5).setValue(precioVenta); 
          costoYVentaActualizados = true;
        }
      }

      break;
    }
  }
  // Actualizacion de Stock
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
