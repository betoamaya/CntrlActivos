//*****************************************P R O D U C C I O N Ultimo cambio: 06/11/2014 /12:50***********************************************************//
// Variables globales
var LABEL_WIDTH = "500px";
var TEXTBOX_WIDTH = "400px";
var BUTTON_WIDTH = "125px";
var tResultado;
var Id_Fila;
var txtMostrar;
var FilaUsuario = 0;

// Spreadsheet de google
var SPSHT_Base = "1qpzvqfRU3fgOFb0VT0mU5VY8FNu9LkuJcelm5u5Jd64";
var SHEET_NAME = "MovRegistrados"
var SPSHT_Temp = "1wtmkuL21ujmvki6H_usKWKTtK1T2nSZ8CzxqElpswMc";
var SHEET_TEMP = "Temp"
var CATALOGO = "1C8lorTDCV7wJEb0LFoL9Xv6vHGyPBR_o_7RYbFJvuUM";

//Imagen
var URLIMAGEN = "https://lh4.googleusercontent.com/-fcBBuYEa2Us/VA2x17HypXI/AAAAAAAAAL4/xsVolaXLT2I/w589-h53-no/barra.jpg";

//Inicia APP

function doGet(e) {
  var app = UiApp.createApplication().setTitle("Cargar Documento Probatorio");
  var Panel = app.createFormPanel()
  .setWidth("810")
  .setStyleAttribute("padding","20px")
  .setStyleAttribute("fontSize", "12pt")
  .setId("Panel");
  
  var flow = app.createVerticalPanel();
  
  
  var usuario = Session.getActiveUser().getEmail();
  /*******Buscar ultimo cambio del Usuario**************/
  buscarUsuario_(usuario);
  
  varMostrar_(0);
  
  var prrText_1 = app.createTextArea().setName("prrText_1")
  .setText(txtMostrar)
  .setStyleAttribute("font-weight", "Normal")
  .setStyleAttribute("fontSize", "8pt")
  .setStyleAttribute("margin-top","10px")
  .setStyleAttribute("margin-bottom","10px")
  .setReadOnly(true)
  .setVisibleLines(10)
  .setWidth("550px");
  
  /********** Termina Busqueda ***************/
  
    
  // Mensaje Inicial
  var lblTitulo = app.createLabel()
  .setId("lblTitulo")
  .setText("Anexar Documento Probatorio");
  
  styleTitulo_(lblTitulo);
  
    
  // Mensaje Usuario
  var lblUsuario = app.createLabel()
  .setId("lblUsuario")
  .setText("Usuario: " + usuario);
  
  styleLbl_(lblUsuario);
   
  // Cargar PDF
  var uplPDF = app.createFileUpload()
  .setName('uplPDF');
  
  styleTxt_(uplPDF);
  
  var btnAnexar = app.createSubmitButton()
  .setText('Anexar Archivo');
  
  styleBtn_(btnAnexar);
  
  //Imagen
  var imgFondo = app.createImage(URLIMAGEN)
  .setStyleAttribute("margin-top","20px");
  
  flow.add(lblTitulo);
  flow.add(lblUsuario);
  flow.add(prrText_1);
  flow.add(uplPDF);
  flow.add(btnAnexar);
  flow.add(imgFondo);
   
  
  Panel.add(flow);
  app.add(Panel);
  return app;
 }

//*-*-*-*-*-*-*-*-*-*-*-*-*-* F I N A L I Z A *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*

function doPost(e) {
  var app = UiApp.getActiveApplication();
  
  var Usuario = Session.getActiveUser().getEmail();
  var sheetx  = SpreadsheetApp.openById(SPSHT_Temp).getSheetByName(SHEET_TEMP);
  var lastRowx = sheetx.getLastRow();
  var rangex = sheetx.getRange("A3:D"+lastRowx)
  var valuesx = rangex.getValues();
  var cadenax;
  var Filax;
    
  //Buscar el usuario
  for(var iix= valuesx.length - 1; iix>-1 ; iix--){
    
   //Agregamos todos los campos en los que se va buscar
    cadenax = valuesx[iix][0];
    
    //Si se cumple la funcion
    if  (cadenax.indexOf(Usuario) != -1) {

      Filax = valuesx[iix][3].toString();
      iix=-1;
    }
  }
  
  
  //Se carga el Archivo en DRIVE
  var banderita = 1
  try {
    var dropbox = "Anexos";
    var folder, folders = DriveApp.getFoldersByName(dropbox);
    // se busca la carpeta especificada si no la crea
    if (folders.hasNext()) 
    {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(dropbox);
    }
    var fileBlob = e.parameter.uplPDF; 
    var file = folder.createFile(fileBlob);//sube el archivo a la carpeta especificada
    
     var sheet  = SpreadsheetApp.openById(SPSHT_Base).getSheetByName(SHEET_NAME);
  /*var cell = sheet.getRange(celda);
  var namefile = file.getName()
  cell.setValue(namefile);*/
  
   var cell = sheet.getRange("W"+ Filax);
  var namefile = file.getId();
  cell.setValue(namefile);
    
  } catch (error) {
    //no hay error que mostrar 
    banderita = 0
  }
  
  var mensaje = app.createLabel("El archivo '" + namefile +"' fue anexado al registro 'W" + Filax + "'. Por favor cierre esta pestaña.")
  
  styleSublbl_(mensaje);
  
  mensaje.setStyleAttribute("margin-top","20px");
  
  app.add(mensaje);
  return app;
  
 }

//*-*-*-*-*-*-*-*-*-*-*-*-*-* E S T I L O S *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
function styleTitulo_(label) {
  label.setStyleAttribute("margin-top","20px");
  label.setStyleAttribute("fontSize", "14pt");
  label.setStyleAttribute("font-weight", "bold");
  label.setStyleAttribute("color","#286389");
}

function styleLbl_(label) {
  label.setStyleAttribute("margin-top","10px");
  label.setStyleAttribute("fontSize", "10pt");
  label.setStyleAttribute("font-weight", "bold");
  label.setStyleAttribute("color","#286389");
}

function styleSublbl_(label) {
  label.setStyleAttribute("margin-top","2px");
  label.setStyleAttribute("fontSize", "8pt");
  label.setStyleAttribute("font-weight", "normal");
  label.setStyleAttribute("color","gray");
}

function styleTxt_(txt) {
  txt.setStyleAttribute("margin-top","2px");
  txt.setStyleAttribute("fontSize", "10pt");
  txt.setStyleAttribute("font-weight", "normal");
  txt.setStyleAttribute("color","Black");
  txt.setWidth(TEXTBOX_WIDTH);
  //txt.setMaxLength(80);
}
  
function styleError_(label) {
  label.setStyleAttribute("margin-top","20px");
  label.setStyleAttribute("fontSize", "12pt");
  label.setStyleAttribute("font-weight", "bold");
  label.setStyleAttribute("color","#BD0E0E");
}

function styleBtn_(btn) {
  btn.setStyleAttribute("margin-top","20px");
  btn.setStyleAttribute("margin-right","5px");
  btn.setStyleAttribute("fontSize", "10pt");
  btn.setStyleAttribute("font-weight", "bold");
  btn.setStyleAttribute("color","Black");
  btn.setWidth(BUTTON_WIDTH);
}

function styleBtnFalse_(btn) {
  btn.setStyleAttribute("margin-top","20px");
  btn.setStyleAttribute("margin-right","5px");
  btn.setStyleAttribute("fontSize", "10pt");
  btn.setStyleAttribute("font-weight", "bold");
  btn.setStyleAttribute("color","gray");
  btn.setWidth(BUTTON_WIDTH);
}

//*-*-*-*-*-*-*-*-*-*-*-*-*-* F U N C I O N *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*


function buscarUsuario_(Usuario){
    
  var sheetx  = SpreadsheetApp.openById(SPSHT_Temp).getSheetByName(SHEET_TEMP);
  var lastRowx = sheetx.getLastRow();
  var rangex = sheetx.getRange("A3:D"+lastRowx)
  var valuesx = rangex.getValues();
  var cadenax;
  var Filax;
    
  //Buscar el usuario
  for(var iix= valuesx.length - 1; iix>-1 ; iix--){
    
   //Agregamos todos los campos en los que se va buscar
    cadenax = valuesx[iix][0];
    
    //Si se cumple la funcion
    if  (cadenax.indexOf(Usuario) != -1) {

      Filax = valuesx[iix][3].toString();
      iix=-1;
    }
  }
  
  //Cargar Catalogos
  var Articulo = "";
  var Marca = "";
  var Movimiento = "";
  var Condicion = "";
  var Cat_Articulos  = SpreadsheetApp.openById(CATALOGO).getSheetByName("Articulos");
  var Cat_Marcas  = SpreadsheetApp.openById(CATALOGO).getSheetByName("Marcas");
  var Cat_Mov  = SpreadsheetApp.openById(CATALOGO).getSheetByName("Movimientos");
  var Cat_Condicion  = SpreadsheetApp.openById(CATALOGO).getSheetByName("Condiciones");
  var UF_Articulos = Cat_Articulos.getLastRow();
  var UF_Marcas = Cat_Marcas.getLastRow();
  var UF_Mov = Cat_Mov.getLastRow();
  var UF_Condiciones = Cat_Condicion.getLastRow();
  var r_Articulos = Cat_Articulos.getRange("A2:B" + UF_Articulos);
  var r_Marcas = Cat_Marcas.getRange("A2:B" + UF_Marcas);
  var r_Mov = Cat_Mov.getRange("A2:B" + UF_Mov);
  var r_Condiciones = Cat_Condicion.getRange("A2:B" + UF_Condiciones);
  var v_Articulos = r_Articulos.getValues();
  var v_Marcas = r_Marcas.getValues();
  var v_Mov = r_Mov.getValues();
  var v_Condiciones = r_Condiciones.getValues();
  var i1 = 0;
  
  //datos de Busqueda
  var sheet  = SpreadsheetApp.openById(SPSHT_Base).getSheetByName(SHEET_NAME);
  var range = sheet.getRange("A"+ Filax +":W"+ Filax);
  //var range = sheet.getRange("A3128:W3128");
  var values = range.getValues();
 
  
  //Agregamos una nueva matriz
  tTemporal = new Array (22);
  
  //Articulo
  for(var i1= 0; i1 < v_Articulos.length ; i1++){
    if (v_Articulos[i1][0].toString() == values[0][2].toString()){
      Articulo = v_Articulos[i1][1].toString();
    }
  }
     
  //Marcas
  for(var i1= 0; i1 < v_Marcas.length ; i1++){
    if (v_Marcas[i1][0].toString() == values[0][3].toString()){
      Marca = v_Marcas[i1][1].toString();
    }
  }
      
  //Movimientos
  for(var i1= 0; i1 < v_Mov.length ; i1++){
    if (v_Mov[i1][0].toString() == values[0][8].toString()){
      Mov = v_Mov[i1][1].toString();
    }
  }
      
  //Condicion
  for(var i1= 0; i1 < v_Condiciones.length ; i1++){
    if (v_Condiciones[i1][0].toString() == values[0][7].toString()){
      Condicion = v_Condiciones[i1][1].toString();
    }
  }
  
  //Arreglar Fecha
  
  var Fecha = Utilities.formatDate(values[0][0],Session.getTimeZone(), "dd/MM/yyyy' 'HH:mm:ss");
  
  tTemporal[0] = Fecha;
  tTemporal[1] = values[0][1].toString();
  tTemporal[2] = Articulo;
  tTemporal[3] = Marca;
  tTemporal[4] = values[0][4].toString();
  tTemporal[5] = values[0][5].toString();
  tTemporal[6] = values[0][6].toString();
  tTemporal[7] = Condicion;
  tTemporal[8] = Mov;
  tTemporal[9] = values[0][9].toString();
  tTemporal[10] = values[0][10].toString();
  tTemporal[11] = values[0][11].toString();
  tTemporal[12] = values[0][12].toString();
  tTemporal[13] = values[0][13].toString();
  tTemporal[14] = values[0][14].toString();
  tTemporal[15] = values[0][15].toString();
  tTemporal[16] = values[0][16].toString();
  tTemporal[17] = values[0][17].toString();
  tTemporal[18] = values[0][18].toString();
  tTemporal[19] = values[0][19].toString();
  tTemporal[20] = values[0][20].toString();
  tTemporal[21] = values[0][21].toString();
  
  tResultado = tTemporal;
}

function varMostrar_(){
  var texto_Mostrar_1 = "Nº ACTIVO: " + tResultado[5].toString() + "\n" + tResultado[2].toString() + " " + tResultado[3].toString() 
        + " CON Nº SERIE: " + tResultado[4].toString() + " MODELO: " + tResultado[19].toString() + "\n" 
        + "IDENTIFICADOR: " + tResultado[6].toString() + "\n"
        + "MOVIMIENTO: " + tResultado[8].toString() + " " + tResultado[20].toString() + " FECHA: " + tResultado[0].toString();
  
  if (trim(tResultado[8].toString()) == "EN STOCK") {
    texto_Mostrar_1 = texto_Mostrar_1 + "\n" +"Almacen: " + tResultado[9].toString();
  }
  if (trim(tResultado[8].toString()) == "ENVIADO A SOPORTE") {
    texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "Asignado a: " + tResultado[10].toString();
  }
  if (trim(tResultado[8].toString()) == "ENVIADO A USUARIO FINAL") {
    texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "Usuario Final: " + tResultado[11].toString();
    texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "Unidad de Negocio: " + tResultado[12].toString();
    texto_Mostrar_1 = texto_Mostrar_1 + " " + "Ubicación: " + tResultado[13].toString();
    texto_Mostrar_1 = texto_Mostrar_1 + " " + " Folio: " + tResultado[21].toString();
  }
  if (trim(tResultado[8].toString()) == "ENVIADO A PROVEEDOR") {
    texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "Nombre del Proveedor: " + tResultado[14].toString();
    texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "Notas: " + tResultado[15].toString();
  }
  if (trim(tResultado[8].toString()) == "EN USO") {
    texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "Usuario Final: " + tResultado[11].toString();
    texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "Unidad de Negocio: " + tResultado[12].toString();
    texto_Mostrar_1 = texto_Mostrar_1 + " " + "Ubicación: " + tResultado[13].toString();
    texto_Mostrar_1 = texto_Mostrar_1 + " " + " Folio: " + tResultado[21].toString();
  }
  
  if (trim(tResultado[8].toString()) == "BAJA (DEP. ACTIVOS)") {
    texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "PERSONA QUE RECIBE: " + tResultado[11].toString();
    texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "FOLIO: " + tResultado[21].toString();
  }
  
    texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "Guia: " + tResultado[16].toString() + " Ext.: " + tResultado[17].toString();
    texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "Observaciones: " + tResultado[18].toString();
  txtMostrar = texto_Mostrar_1;
}

function trim(s)
{
  var myString = s;
  var j = 0;

  // trim front of string
  for(var i = 0; i < myString.length; i++)
  {
    if(myString.charAt(i) != " ")
    {
      break
    }
    else
    {
      j += 1;
    }
  }
  myString = myString.substring(j, myString.length);

  // trim back of string
  j = myString.length;
  for(var i = myString.length - 1; i >= 0; i--)
  {
    if(myString.charAt(i) != " ")
    {
      break
    }
    else
    {
      j -= 1;
    }
  }
  myString = myString.substring(0, j);

  return myString;
}
