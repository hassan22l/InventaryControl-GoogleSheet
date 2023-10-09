//Guardado de productos nuevos

function guardarProductos() {

    var libro = SpreadsheetApp.getActiveSpreadsheet();
    var hojaProductos = libro.getSheetByName("Ingr.Productos")
  
    var codigo = hojaProductos.getRange("C3").getValue();
    var nombre = hojaProductos.getRange("C5").getValue();
    var descripcion = hojaProductos.getRange("E3").getValue();
    var precio = hojaProductos.getRange("G3").getValue();
    var precioVenta = hojaProductos.getRange("I3").getValue();
    var equivlencias = hojaProductos.getRange("E5").getValue();
    var lugar = hojaProductos.getRange("G5").getValue();
  
    var hojaProductosInfo = libro.getSheetByName("Productos.Info")
    hojaProductosInfo.appendRow([codigo, nombre, descripcion, precio, precioVenta, equivlencias, lugar])
  };
  
  
  //Guardado de compras realizadas en Historial de compras
  
  function guardarEntrada() {
  
    var libro = SpreadsheetApp.getActiveSpreadsheet();
    var hojaCompras = libro.getSheetByName("Compras")
    var hojaHistorial = libro.getSheetByName("Info.Compras")
  
    var codigo = hojaCompras.getRange("C5").getValue();
    var nombre = hojaCompras.getRange("C7").getValue();
    var descripcion = hojaCompras.getRange("E5").getValue();
    var cantidad = hojaCompras.getRange("E7").getValue();
    var precio = hojaCompras.getRange("G5").getValue();
    var precioVenta = hojaCompras.getRange("I5").getValue();
    var equivlencias = hojaCompras.getRange("C9").getValue();
    var lugar = hojaCompras.getRange("E9").getValue();
    var proveedor = hojaCompras.getRange("D3").getValue();
    var fecha = hojaCompras.getRange("G3").getValue();
  
    hojaHistorial.appendRow([codigo, nombre, descripcion, cantidad, precio, precioVenta, equivlencias, lugar, proveedor, fecha])
  
  //Actualizacion de Stock en Hoja de Stock
  
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
    hojaCompras.getRange("E5").clearContent();
    hojaCompras.getRange("E7").clearContent();
    hojaCompras.getRange("G5").clearContent();
    hojaCompras.getRange("G7").clearContent();
    hojaCompras.getRange("I5").clearContent();
    hojaCompras.getRange("C9").clearContent();
    hojaCompras.getRange("E9").clearContent();
  
  }
  
  //Guaradado de ventas en historial de ventas
  
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
  