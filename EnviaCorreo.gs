//*****************************************P R O D U C C I O N Ultimo cambio: 06/07/2015/12:49***********************************************************//
// Variables globales
var LABEL_WIDTH = "500px";
var TEXTBOX_WIDTH = "400px";
var BUTTON_WIDTH = "125px";
var tResultado;
var Id_Fila;
var txtMostrar;
var txtHTML;
var FilaUsuario = 0;
var correoActivos = 'josue.martinez@transpais.com.mx, adolfo.quintero@transpais.com.mx, jasanchez@transpais.com.mx';

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
  var app = UiApp.createApplication().setTitle("Enviar Correo");
  var Panel = app.createFormPanel()
  .setWidth("810px")
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
  .setVisibleLines(25)
  .setWidth("550px");
  
  /********** Termina Busqueda ***************/
  
    
  // Mensaje Inicial
  var lblTitulo = app.createLabel()
  .setId("lblTitulo")
  .setText("Enviar Correo a Usuario");
  
  styleTitulo_(lblTitulo);
  
    
  // Mensaje Usuario
  var lblUsuario = app.createLabel()
  .setId("lblUsuario")
  .setText("Usuario: " + usuario);
  
  styleLbl_(lblUsuario);
   
  // Destinatarios Usuario
  var lblDestinatarios = app.createLabel()
  .setId("lbldestinatarios")
  .setText("Se enviara correo con copia a las siguientes destinatarios: " + correoActivos + ", " + usuario);
  
  styleLbl_(lblDestinatarios);
  
  // Nuevo Destinatario
  var lblCc = app.createLabel()
  .setId("lblCc")
  .setText("Destinatario:");
  
  // 
   var lblsCc = app.createLabel()
  .setId("lblsCc")
  .setText("Introduzca la dirección del destinatario, pude agregar mas de un destinatario separando los correos por coma");
    
  // 
  var txtCc = app.createTextBox()
  .setFocus(true)
  .setName("txtCc")
  
      
  //Style
  styleLbl_(lblCc);
  styleSublbl_(lblsCc);
  styleTxt_(txtCc);
  
  
  
  var btnEnviar = app.createSubmitButton()
  .setText('Enviar Correo');
  
  styleBtn_(btnEnviar);
  
  //Imagen
  var imgFondo = app.createImage(URLIMAGEN)
  .setStyleAttribute("margin-top","20px");
  
  flow.add(lblTitulo);
  flow.add(lblDestinatarios);
  flow.add(lblCc);
  flow.add(lblsCc);
  flow.add(txtCc);
  flow.add(prrText_1);
  
  flow.add(btnEnviar);
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
  
  
  buscarUsuario_(Usuario);
  
  varMostrar_(0);
  
  var enviarA = e.parameter.txtCc
  
  MailApp.sendEmail({
     to: enviarA,
     subject: 'Se registro un movimiento con el Activo ' + tResultado[5],
     htmlBody: txtHTML,
     replyTo: 'atencionsd@transpais.com.mx',
    cc: Usuario + ", " + correoActivos
   });
   
  
  var mensaje = app.createLabel("El correo fue enviado. Por favor cierre esta pestaña.")
  
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
  
  var Fecha = Utilities.formatDate(values[0][0],Session.getScriptTimeZone(), "dd/MM/yyyy' 'HH:mm:ss");
  
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
  tTemporal[22] = values[0][22].toString();
  
  tResultado = tTemporal;
}

function varMostrar_(){
  var texto_Mostrar_1 = "El siguiente activo fue registrado con un movimiento. \n" 
        + "*************************************************************************************************************" 
        + "\n" + "Nº ACTIVO: " + tResultado[5].toString() + "\n" 
        + "ARTICULO: " + tResultado[2].toString() + "\n" + "\n"
        + "MARCA: " + tResultado[3].toString() + "\n" 
        + "MODELO: " + tResultado[19].toString() + "\n" + "\n"
        + "Nº SERIE: " + tResultado[4].toString() + "\n" 
        + "IDENTIFICADOR: " + tResultado[6].toString() + "\n" + "\n"
        + "MOVIMIENTO: " + tResultado[8].toString() + " " + tResultado[20].toString() 
        + "        FECHA: " + tResultado[0].toString();
  
  if (trim(tResultado[8].toString()) == "EN STOCK") {
    texto_Mostrar_1 = texto_Mostrar_1 + "\n" +"ALMACEN: " + tResultado[9].toString() + "\n";
  }
  if (trim(tResultado[8].toString()) == "ENVIADO A SOPORTE") {
    texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "ASIGNADO A: " + tResultado[10].toString() + "\n";
  }
  if (trim(tResultado[8].toString()) == "ENVIADO A USUARIO FINAL") {
    texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "USUARIO FINAL: " + tResultado[11].toString() + "\n";
    texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "UNIDAD DE NEGOCIO: " + tResultado[12].toString();
    texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "UBICACION: " + tResultado[13].toString() + "\n";
    texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "FOLIO: " + tResultado[21].toString();
  }
  if (trim(tResultado[8].toString()) == "ENVIADO A PROVEEDOR") {
    texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "NOMBRE DEL PROVEEDOR: " + tResultado[14].toString() + "\n";
    texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "NOTAS: " + tResultado[15].toString();
  }
  if (trim(tResultado[8].toString()) == "EN USO") {
    texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "USUARIO FINAL: " + tResultado[11].toString() + "\n";
    texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "UNIDAD DE NEGOCIO: " + tResultado[12].toString();
    texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "UBICACION: " + tResultado[13].toString()+ "\n" ;
    texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "FOLIO: " + tResultado[21].toString();
  }
  
  if (trim(tResultado[8].toString()) == "BAJA (DEP. ACTIVOS)") {
    texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "PERSONA QUE RECIBE: " + tResultado[11].toString();
    texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "FOLIO: " + tResultado[21].toString();
  }
  
    texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "GUIA: " + tResultado[16].toString() + "       EXT.: " + tResultado[17].toString();
    texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "OBSERVACIONES: " + tResultado[18].toString();
  texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "*************************************************************************************************************"
                                    + "\n" + "IMPORTANTE" + "\n"
                                    + "\n" + "ESTE CORREO DEBE SER RESPONDIDO EN UN MAXIMO DE 4 DIAS HABILES, AL NO " 
                                    + "PRESENTAR NIGUN COMENTARIO O ACALARACION SE DARA COMO ACEPTADO EL ACTIVO." + "\n"
                                    + "\n" + "Para cualquier aclaración, favor de indicar clave y nombre de resguardante, "
                                    + "asi como el centro de costo." 
                                    
  txtMostrar = texto_Mostrar_1;
  
   //*************HTML********************* Version 1.3
  var txtMostrar_2 = '<!doctype html><html><head><title>Enviar correo al Usuario</title></head><body>';
  //Tabla
  txtMostrar_2 = txtMostrar_2 + '<table border="1" cellpadding="0" cellspacing="0" dir="ltr" style="table-layout:fixed;font-size:13px;' +
    'font-family:arial,sans,sans-serif;border-collapse:collapse;border:1px solid #ccc">' +
      '<colgroup><col width="120" /><col width="200" /><col width="120" /><col width="200" /></colgroup><tbody>';
   //Titulo
  txtMostrar_2 = txtMostrar_2 + '<tr style="height:35px;"><td colspan="4" data-sheets-value="[null,2,&quot;El siguiente activo fue registrado con un movimiento&quot;]" ' +
    'rowspan="1" style="padding: 2px 3px; font-size: 120%; font-weight: bold; color: rgb(255, 255, 255); word-wrap: break-word; vertical-align: middle; text-align: center; ' +
      'background-color: rgb(11, 56, 97);">El siguiente activo fue registrado con un movimiento</td></tr>';
  //Datos del Activo
  txtMostrar_2 = txtMostrar_2 + '<tr style="height:30px;"><td data-sheets-value="[null,2,&quot;N\u00b0 Activo:&quot;]" style="padding: 2px 3px; font-weight: bold; word-wrap: ' +
    'break-word; vertical-align: middle; background-color: rgb(201, 218, 248);">N&deg; Activo:</td><td colspan="3" data-sheets-value="[null,2,&quot;texto&quot;]" rowspan="1" ' +
      'style="padding: 2px 3px; word-wrap: break-word; vertical-align: middle;">' + tResultado[5].toString() + '</td></tr>';
  txtMostrar_2 = txtMostrar_2 + '<tr style="height:30px;"><td data-sheets-value="[null,2,&quot;Articulo:&quot;]" style="padding: 2px 3px; font-weight: bold; word-wrap: ' +
    'break-word; vertical-align: middle; background-color: rgb(201, 218, 248);">Articulo:</td><td data-sheets-value="[null,2,&quot;texto&quot;]" style="padding: 2px 3px; ' +
      'word-wrap: break-word; vertical-align: middle;">' + tResultado[2].toString() + '</td><td data-sheets-value="[null,2,&quot;Marca:&quot;]" style="padding: 2px 3px; font-weight: bold; ' +
        'word-wrap: break-word; vertical-align: middle; background-color: rgb(201, 218, 248);">Marca:</td><td data-sheets-value="[null,2,&quot;texto&quot;]" style="padding: 2px 3px; ' +
          'word-wrap: break-word; vertical-align: middle;">' + tResultado[3].toString() + '</td></tr>';
   txtMostrar_2 = txtMostrar_2 + '<tr style="height:30px;"><td data-sheets-value="[null,2,&quot;Modelo:&quot;]" style="padding: 2px 3px; font-weight: bold; word-wrap: break-word; ' +
     'vertical-align: middle; background-color: rgb(201, 218, 248);">Modelo:</td><td data-sheets-value="[null,2,&quot;texto&quot;]" style="padding: 2px 3px; word-wrap: break-word; ' +
       'vertical-align: middle;">' +  tResultado[19].toString() + '</td><td data-sheets-value="[null,2,&quot;N\u00b0 Serie:&quot;]" style="padding: 2px 3px; font-weight: bold; word-wrap: ' +
         'break-word; vertical-align: middle; background-color: rgb(201, 218, 248);">N&deg; Serie:</td><td data-sheets-value="[null,2,&quot;texto&quot;]" style="padding: 2px 3px; word-wrap: ' +
           'break-word; vertical-align: middle;">'  + tResultado[4].toString() + '</td></tr>';
  txtMostrar_2 = txtMostrar_2 + '<tr style="height:30px;"><td data-sheets-value="[null,2,&quot;Identificador:&quot;]" style="padding: 2px 3px; font-weight: bold; word-wrap: break-word; ' +
    'vertical-align: middle; background-color: rgb(201, 218, 248);">Identificador:</td><td colspan="3" data-sheets-value="[null,2,&quot;texto&quot;]" rowspan="1" style="padding: 2px 3px; ' +
      'word-wrap: break-word; vertical-align: middle;">' + tResultado[6].toString() + '</td></tr>';
  
  //Datos del Movimiento
  txtMostrar_2 = txtMostrar_2 + '<tr style="height:35px;"><td colspan="4" data-sheets-value="[null,2,&quot;Datos del movimiento&quot;]" rowspan="1" style="padding: 2px 3px; font-size: 120%; ' +
    'font-weight: bold; color: rgb(255, 255, 255); word-wrap: break-word; vertical-align: middle; text-align: center; background-color: rgb(11, 56, 97);">Datos del movimiento</td></tr>';
  txtMostrar_2 = txtMostrar_2 + '<tr style="height:30px;"><td data-sheets-value="[null,2,&quot;Movimiento:&quot;]" style="padding: 2px 3px; font-weight: bold; word-wrap: break-word; ' +
    'vertical-align: middle; background-color: rgb(201, 218, 248);">Movimiento:</td><td data-sheets-value="[null,2,&quot;texto&quot;]" style="padding: 2px 3px; word-wrap: break-word; ' +
      'vertical-align: middle;">' + tResultado[8].toString() + '</td><td data-sheets-value="[null,2,&quot;Fecha:&quot;]" style="padding: 2px 3px; font-weight: bold; word-wrap: break-word; ' +
        'vertical-align: middle; background-color: rgb(201, 218, 248);">Fecha:</td><td data-sheets-value="[null,2,&quot;texto&quot;]" style="padding: 2px 3px; word-wrap: break-word; ' +
          'vertical-align: middle;">' + tResultado[0].toString(); + '</td></tr>';
  
  if (trim(tResultado[8].toString()) == "EN STOCK") {
    txtMostrar_2 = txtMostrar_2  +  '<tr style="height:30px;"><td data-sheets-value="[null,2,&quot;Almacen:&quot;]" style="padding: 2px 3px; font-weight: bold; word-wrap: break-word; ' +
      'vertical-align: middle; background-color: rgb(201, 218, 248);">Almacen:</td><td colspan="3" data-sheets-value="[null,2,&quot;texto&quot;]" rowspan="1" style="padding: 2px 3px; ' +
        'word-wrap: break-word; vertical-align: middle;">' +  tResultado[9].toString() + '</td></tr>';
  }
  
   if (trim(tResultado[8].toString()) == "ENVIADO A SOPORTE") {
    txtMostrar_2 = txtMostrar_2  +  '<tr style="height:30px;"><td data-sheets-value="[null,2,&quot;Asignado a:&quot;]" style="padding: 2px 3px; font-weight: bold; word-wrap: break-word; ' +
      'vertical-align: middle; background-color: rgb(201, 218, 248);">Asignado a:</td><td colspan="3" data-sheets-value="[null,2,&quot;texto&quot;]" rowspan="1" style="padding: 2px 3px; ' +
        'word-wrap: break-word; vertical-align: middle;">' + tResultado[10].toString() + '</td></tr>';
  }
  
  if (trim(tResultado[8].toString()) == "ENVIADO A USUARIO FINAL") {
    txtMostrar_2 = txtMostrar_2  +  '<tr style="height:30px;"><td data-sheets-value="[null,2,&quot;Usuario Final:&quot;]" style="padding: 2px 3px; font-weight: bold; word-wrap: break-word; ' +
      'vertical-align: middle; background-color: rgb(201, 218, 248);">Usuario Final:</td><td colspan="3" data-sheets-value="[null,2,&quot;texto&quot;]" rowspan="1" style="padding: 2px 3px; ' +
        'word-wrap: break-word; vertical-align: middle;">' + tResultado[11].toString() + '</td></tr>';
    txtMostrar_2 = txtMostrar_2  +  '<tr style="height:30px;"><td data-sheets-value="[null,2,&quot;Unidad de Negocio: &quot;]" style="padding: 2px 3px; font-weight: bold; word-wrap: break-word; ' +
      'vertical-align: middle; background-color: rgb(201, 218, 248);">Unidad de Negocio:</td><td data-sheets-value="[null,2,&quot;texto&quot;]" style="padding: 2px 3px; word-wrap: break-word; ' +
        'vertical-align: middle;">' + tResultado[12].toString() + '</td><td data-sheets-value="[null,2,&quot;Ubicaci\u00f3n:&quot;]" style="padding: 2px 3px; font-weight: bold; word-wrap: ' +
          'break-word; vertical-align: middle; background-color: rgb(201, 218, 248);">Ubicaci&oacute;n:</td><td data-sheets-value="[null,2,&quot;texto&quot;]" style="padding: 2px 3px; word-wrap: ' +
            'break-word; vertical-align: middle;">' + tResultado[13].toString() + '</td></tr>';
    txtMostrar_2 = txtMostrar_2  +  '<tr style="height:30px;"><td data-sheets-value="[null,2,&quot;Folio:&quot;]" style="padding: 2px 3px; font-weight: bold; word-wrap: break-word; ' +
      'vertical-align: middle; background-color: rgb(201, 218, 248);">Folio:</td><td data-sheets-value="[null,2,&quot;texto&quot;]" style="padding: 2px 3px; word-wrap: break-word; ' +
        'vertical-align: middle;">' + tResultado[21].toString() + '</td><td style="padding:2px 3px 2px 3px;vertical-align:bottom;">&nbsp;</td>' +
          '<td style="padding:2px 3px 2px 3px;vertical-align:bottom;">&nbsp;</td></tr>';
  }
  
  if (trim(tResultado[8].toString()) == "ENVIADO A PROVEEDOR") {
    txtMostrar_2 = txtMostrar_2  +  '<tr style="height:30px;"><td data-sheets-value="[null,2,&quot;Nombre del proveedor:&quot;]" style="padding: 2px 3px; font-weight: bold; word-wrap: break-word; ' +
      'vertical-align: middle; background-color: rgb(201, 218, 248);">Nombre del proveedor:</td><td colspan="3" data-sheets-value="[null,2,&quot;texto&quot;]" rowspan="1" style="padding: 2px 3px; ' +
        'word-wrap: break-word; vertical-align: middle;">' + tResultado[14].toString() + '</td></tr>';
    txtMostrar_2 = txtMostrar_2  +  '<tr style="height:30px;"><td data-sheets-value="[null,2,&quot;Notas:&quot;]" style="padding: 2px 3px; font-weight: bold; word-wrap: break-word; ' +
      'vertical-align: middle; background-color: rgb(201, 218, 248);">Notas:</td><td colspan="3" data-sheets-value="[null,2,&quot;texto&quot;]" rowspan="1" style="padding: 2px 3px; ' +
        'word-wrap: break-word; vertical-align: middle;">' + tResultado[15].toString() + '</td></tr>';
  }
  
  if (trim(tResultado[8].toString()) == "EN USO") {
    txtMostrar_2 = txtMostrar_2  +  '<tr style="height:30px;"><td data-sheets-value="[null,2,&quot;Usuario Actual:&quot;]" style="padding: 2px 3px; font-weight: bold; word-wrap: break-word; ' +
      'vertical-align: middle; background-color: rgb(201, 218, 248);">Usuario Actual:</td><td colspan="3" data-sheets-value="[null,2,&quot;texto&quot;]" rowspan="1" style="padding: 2px 3px; ' +
        'word-wrap: break-word; vertical-align: middle;">' + tResultado[11].toString() + '</td></tr>';
    txtMostrar_2 = txtMostrar_2  +  '<tr style="height:30px;"><td data-sheets-value="[null,2,&quot;Unidad de Negocio: &quot;]" style="padding: 2px 3px; font-weight: bold; word-wrap: break-word; ' +
      'vertical-align: middle; background-color: rgb(201, 218, 248);">Unidad de Negocio:</td><td data-sheets-value="[null,2,&quot;texto&quot;]" style="padding: 2px 3px; word-wrap: break-word; ' +
        'vertical-align: middle;">' + tResultado[12].toString() + '</td><td data-sheets-value="[null,2,&quot;Ubicaci\u00f3n:&quot;]" style="padding: 2px 3px; font-weight: bold; word-wrap: ' +
          'break-word; vertical-align: middle; background-color: rgb(201, 218, 248);">Ubicaci&oacute;n:</td><td data-sheets-value="[null,2,&quot;texto&quot;]" style="padding: 2px 3px; word-wrap: ' +
            'break-word; vertical-align: middle;">' + tResultado[13].toString() + '</td></tr>';
    txtMostrar_2 = txtMostrar_2  +  '<tr style="height:30px;"><td data-sheets-value="[null,2,&quot;Folio:&quot;]" style="padding: 2px 3px; font-weight: bold; word-wrap: break-word; ' +
      'vertical-align: middle; background-color: rgb(201, 218, 248);">Folio:</td><td data-sheets-value="[null,2,&quot;texto&quot;]" style="padding: 2px 3px; word-wrap: break-word; ' +
        'vertical-align: middle;">' + tResultado[21].toString() + '</td><td style="padding:2px 3px 2px 3px;vertical-align:bottom;">&nbsp;</td>' +
          '<td style="padding:2px 3px 2px 3px;vertical-align:bottom;">&nbsp;</td></tr>';
  }
  
  if (trim(tResultado[8].toString()) == "BAJA (DEP. ACTIVOS)") {
    txtMostrar_2 = txtMostrar_2  +  '<tr style="height:30px;"><td data-sheets-value="[null,2,&quot;Persona que recibe:&quot;]" style="padding: 2px 3px; font-weight: bold; word-wrap: break-word; ' +
      'vertical-align: middle; background-color: rgb(201, 218, 248);">Persona que recibe:</td><td colspan="3" data-sheets-value="[null,2,&quot;texto&quot;]" rowspan="1" style="padding: 2px 3px; ' +
        'word-wrap: break-word; vertical-align: middle;">' + tResultado[11].toString() + '</td></tr>';
    txtMostrar_2 = txtMostrar_2  +  '<tr style="height:30px;"><td data-sheets-value="[null,2,&quot;Folio:&quot;]" style="padding: 2px 3px; font-weight: bold; word-wrap: break-word; ' +
      'vertical-align: middle; background-color: rgb(201, 218, 248);">Folio:</td><td colspan="3" data-sheets-value="[null,2,&quot;texto&quot;]" rowspan="1" style="padding: 2px 3px; ' +
        'word-wrap: break-word; vertical-align: middle;">' + tResultado[21].toString() + '</td></tr>';
  }
  
   txtMostrar_2 = txtMostrar_2  +  '<tr style="height:30px;"><td data-sheets-value="[null,2,&quot;Guia:&quot;]" style="padding: 2px 3px; font-weight: bold; word-wrap: break-word; ' +
     'vertical-align: middle; background-color: rgb(201, 218, 248);">Guia:</td><td data-sheets-value="[null,2,&quot;texto&quot;]" style="padding: 2px 3px; word-wrap: break-word; ' +
       'vertical-align: middle;">' + tResultado[16].toString() + '</td><td data-sheets-value="[null,2,&quot;Ext.:&quot;]" style="padding: 2px 3px; font-weight: bold; word-wrap: break-word; ' +
         'vertical-align: middle; background-color: rgb(201, 218, 248);">Ext.:</td><td data-sheets-value="[null,2,&quot;texto&quot;]" style="padding: 2px 3px; word-wrap: break-word; ' +
           'vertical-align: middle;">' + tResultado[17].toString() + '</td></tr>';
  
  txtMostrar_2 = txtMostrar_2  +  '<tr style="height:30px;"><td data-sheets-value="[null,2,&quot;Observaciones:&quot;]" style="padding: 2px 3px; font-weight: bold; word-wrap: break-word; ' +
    'vertical-align: middle; background-color: rgb(201, 218, 248);">Observaciones:</td><td colspan="3" data-sheets-value="[null,2,&quot;texto&quot;]" rowspan="1" style="padding: 2px 3px; ' +
      'word-wrap: break-word; vertical-align: middle;">' + tResultado[18].toString() + '</td></tr></tbody></table>';
  
  txtMostrar_2 = txtMostrar_2  +  '<p style="color: rgb(34, 34, 34); font-family: arial, sans-serif; font-size: 12.8000001907349px; background-color: rgb(255, 255, 255);"><strong>IMPORTANTE' +
    '</strong></p><p style="color: rgb(34, 34, 34); font-family: arial, sans-serif; font-size: 12.8000001907349px; background-color: rgb(255, 255, 255);">' +
      '<strong>ESTE CORREO DEBE SER RESPONDIDO EN UN MAXIMO DE 4 DIAS HABILES, AL NO PRESENTAR NINGUN COMENTARIO O ACALARACION SE DARA COMO ACEPTADO EL ACTIVO.</strong></p>' +
        '<p style="color: rgb(34, 34, 34); font-family: arial, sans-serif; font-size: 12.8000001907349px; background-color: rgb(255, 255, 255);">Para cualquier aclaraci&oacute;n, ' +
          'favor de indicar clave y nombre de resguardante, asi como el centro de costo.</p><p style="color: rgb(34, 34, 34); font-family: arial, sans-serif; font-size: 12.8000001907349px; ' +
            'background-color: rgb(255, 255, 255);"><img alt="TI" src="https://lh4.googleusercontent.com/-fcBBuYEa2Us/VA2x17HypXI/AAAAAAAAAL4/xsVolaXLT2I/w589-h53-no/barra.jpg" style="width: ' +
              '589px; height: 53px;" /></p><p>&nbsp;</p></body></html>';
  
  /* Html v1.2
  
  //*************HTML*********************
  var txtMostrar_2 = '<!doctype html><html><head><title>HTML Editor - Full Version</title></head><body><h2>El siguiente activo fue registrado con un movimiento</h2><hr/>'
  txtMostrar_2 = txtMostrar_2 + '<p><strong>N&deg; Activo:</strong>' + tResultado[5].toString() +  '</p>' + 
    '<p><strong>Articulo: </strong>' + tResultado[2].toString() + '</p>' +
      '<p><strong>Marca: </strong>' + tResultado[3].toString() + '</p>' +
        '<p><strong>Modelo: </strong>' +  tResultado[19].toString() + '</p>' +
          '<p><strong>N&deg; Serie: </strong>'  + tResultado[4].toString() + '</p>' +
            '<p><strong>Identificador: </strong>' + tResultado[6].toString() + '</p>' +
              '<p><strong>Movimiento: </strong>' + tResultado[8].toString() + ' ' + tResultado[20].toString() + '&nbsp;' +
                '<strong>Fecha: </strong>' + tResultado[0].toString(); + '</p>';
  
  if (trim(tResultado[8].toString()) == "EN STOCK") {
    txtMostrar_2 = txtMostrar_2  +  '<p><strong>Almacen: </strong>' +  tResultado[9].toString() + '</p>';
  }
   if (trim(tResultado[8].toString()) == "ENVIADO A SOPORTE") {
    txtMostrar_2 = txtMostrar_2  +  '<p><strong>Asignado a: </strong>' + tResultado[10].toString() + '</p>';
  }
  if (trim(tResultado[8].toString()) == "ENVIADO A USUARIO FINAL") {
    txtMostrar_2 = txtMostrar_2  +  '<p><strong>Usuario Final: </strong>' + tResultado[11].toString() + '</p>';
    txtMostrar_2 = txtMostrar_2  +  '<p><strong>Unidad de Negocio: </strong>' + tResultado[12].toString() + '</p>';
    txtMostrar_2 = txtMostrar_2  +  '<p><strong>Ubicación: </strong>' + tResultado[13].toString() + '</p>';
    txtMostrar_2 = txtMostrar_2  +  '<p><strong>Folio: </strong>' + tResultado[21].toString() + '</p>';
  }
  if (trim(tResultado[8].toString()) == "ENVIADO A PROVEEDOR") {
    txtMostrar_2 = txtMostrar_2  +  '<p><strong>Nombre del proveedor: </strong>' + tResultado[14].toString() + '</p>';
    txtMostrar_2 = txtMostrar_2  +  '<p><strong>Notas: </strong>' + tResultado[15].toString() + '</p>';
  }
  if (trim(tResultado[8].toString()) == "EN USO") {
    txtMostrar_2 = txtMostrar_2  +  '<p><strong>Usuario Actual: </strong>' + tResultado[11].toString() + '</p>';
    txtMostrar_2 = txtMostrar_2  +  '<p><strong>Unidad de Negocio: </strong>' + tResultado[12].toString() + '</p>';
    txtMostrar_2 = txtMostrar_2  +  '<p><strong>Ubicación: </strong>' + tResultado[13].toString() + '</p>';
    txtMostrar_2 = txtMostrar_2  +  '<p><strong>Folio: </strong>' + tResultado[21].toString() + '</p>';
  }
  
  if (trim(tResultado[8].toString()) == "BAJA (DEP. ACTIVOS)") {
    txtMostrar_2 = txtMostrar_2  +  '<p><strong>Persona que recibe: </strong>' + tResultado[11].toString() + '</p>';
    txtMostrar_2 = txtMostrar_2  +  '<p><strong>Folio: </strong>' + tResultado[21].toString() + '</p>';
  }
  txtMostrar_2 = txtMostrar_2 + '<p><strong>Guia: </strong>' + tResultado[16].toString() + ' &nbsp;<strong>Ext.: </strong>' + tResultado[17].toString() + '</p>' +
    '<p><strong>Observaciones: </strong>' + tResultado[18].toString() + '</p><hr/><p><strong>IMPORTANTE</strong></p>' +
      '<p><strong>ESTE CORREO DEBE SER RESPONDIDO EN UN MAXIMO DE 4 DIAS HABILES, AL NO PRESENTAR NINGUN COMENTARIO O ACALARACION SE DARA COMO ACEPTADO EL ACTIVO.</strong></p>' +
        '<p>Para cualquier aclaraci&oacute;n, favor de indicar clave y nombre de resguardante, asi como el centro de costo.</p>' +
          '<p><img alt="barra" src="https://lh4.googleusercontent.com/-fcBBuYEa2Us/VA2x17HypXI/AAAAAAAAAL4/xsVolaXLT2I/w589-h53-no/barra.jpg" style="width: 589px; height: 53px; float: left;" /></p>' +
            '</body></html>';*/
  txtHTML = txtMostrar_2;
  
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
