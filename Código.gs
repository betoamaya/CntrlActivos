//***********************************P R O D U C C I O N / PUBLICADO:26/05/2015 12:47*****************************************************
// Variables globales
var LABEL_WIDTH = "500px";
var TEXTBOX_WIDTH = "400px";
var BUTTON_WIDTH = "125px";

// Spreadsheet de google
var SPSHT_Base = "1qpzvqfRU3fgOFb0VT0mU5VY8FNu9LkuJcelm5u5Jd64";
var SHEET_NAME = "MovRegistrados"
var SPSHT_Temp = "1wtmkuL21ujmvki6H_usKWKTtK1T2nSZ8CzxqElpswMc";
var SHEET_TEMP = "Temp"
var CATALOGO = "1C8lorTDCV7wJEb0LFoL9Xv6vHGyPBR_o_7RYbFJvuUM";
var REPORTES = "1Bkw_6vzm0BFHBHzbhCCv-yhUmUmBC05r8L1sGA3IXN8"
var SHEET_REPO1 = "Reporte1"

var DocRecporte = "1EHUnXdXJkoVr1g_GALi6MpZiU6y-45BfeYWscNAp3us"

//Imagen
var URLIMAGEN = "https://lh4.googleusercontent.com/-fcBBuYEa2Us/VA2x17HypXI/AAAAAAAAAL4/xsVolaXLT2I/w589-h53-no/barra.jpg";
var IMGPROCESO = "http://i93.photobucket.com/albums/l45/lawebera/articulos/ajax-loader.gif";
var IMGEDITAR = "http://img1.wikia.nocookie.net/__cb20101009133120/es.pokemon/images/2/2d/Icono_de_editar.png";
var IMGDOCUMENT = "http://icons.iconarchive.com/icons/hopstarter/soft-scraps/32/Adobe-PDF-Document-icon.png"

//Varibales campos
var tResultado;
var tActivo;
var resMostrar = 5;
var txtMostrar = "";

//Inicia APP
function doGet(e) {
  var app = UiApp.createApplication().setTitle("Registrar movimiento de activo");
  //Formulario de Inicio
  frmInicio_(app, "0")
  return app;
}

//*-*-*-*-*-*-*-*-*-*-*-*-*-* F O R M U L A R I O S *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
function frmInicio_(container, opc){
  
  //Crear Panel Vertical
  var vPanel = container.createVerticalPanel()
  .setWidth("810")
  .setStyleAttribute("padding","20px")
  .setStyleAttribute("fontSize", "12pt")
  .setId("panel");
  
  // Mensaje Inicial
  var lblTitulo = container.createLabel()
  .setId("lblTitulo")
  .setText("Registrar Movimientos de Activos");
    
  // Mensaje Busqueda
  var lblBusqueda = container.createLabel()
  .setId("lblBusqueda")
  .setText("Busqueda General:");
  
  // Sub-Mensaje Busqueda
   var lblsBusqueda = container.createLabel()
  .setId("lblsBusqueda")
  .setText("Introduzca Palabra a Buscar");
    
  // Text Box Busqueda
  var txtBuscar = container.createTextBox()
  .setFocus(true)
  .setName("txtBuscar")
  // Accion Boton Buscar
      
  //Style
  styleTitulo_(lblTitulo);
  styleLbl_(lblBusqueda);
  styleSublbl_(lblsBusqueda);
  styleTxt_(txtBuscar);
  
  //Imagen
  var imgFondo = container.createImage(URLIMAGEN)
  .setStyleAttribute("margin-top","20px");
    
  //Crear Panel Horizontal
  var hPanel = container.createHorizontalPanel();
  
  //Boton Buscar
  var btnBuscar = container.createButton()
  .setText("Buscar");
      
  //Boton Nuevo  
  var btnNuevo = container.createButton()
  .setText("Crear Registro");
  
  //Style
  styleBtn_(btnBuscar);
  styleBtn_(btnNuevo);  
  
  // Accion Boton Buscar
  var lnzBuscar = container.createServerHandler('clickBuscarGral_')
  .addCallbackElement(vPanel);
  btnBuscar.addClickHandler(lnzBuscar);
  
  // Accion Boton Nuevo
  var lnzNuevo = container.createServerHandler('clickNuevo_')
  .addCallbackElement(vPanel);
  btnNuevo.addClickHandler(lnzNuevo);
  
  //Agregar al Panel Horizontal
  hPanel.add(btnBuscar);
  hPanel.add(btnNuevo);
  
  // Agrega a Panel Vertical
  vPanel.add(lblTitulo);
  vPanel.add(lblBusqueda);
  vPanel.add(lblsBusqueda);
  vPanel.add(txtBuscar);
  vPanel.add(hPanel);
  vPanel.add(imgFondo);
  
  
  //Agregar a APP
  container.add(vPanel);
}

function frmResultados_(container,e, palabra_Buscar, serie){
   
  //Crear Panel Vertical
  var vPanel = container.createVerticalPanel()
  .setWidth("810px")
  .setStyleAttribute("padding","20px")
  .setStyleAttribute("fontSize", "12pt")
  .setId("panel");
    
  // Titulo
  var lblTitulo = container.createLabel()
  .setId("lblTitulo")
  .setText("Resultados:");
  styleTitulo_(lblTitulo);
  vPanel.add(lblTitulo);
  
  // Sub label
  var lblsInfo = container.createLabel()
  .setId("lblsInfo")
  .setText("Palabra buscada: " + palabra_Buscar);
  styleSublbl_(lblsInfo);
  vPanel.add(lblsInfo);
  
     
  //Desplegar Resultados
  var ii = 0;
  var op = 0;
  var ultimaFila = tResultado.length;
  var sInicio = serie - resMostrar;
  //var sFin = serie - 1;
  var sFin = sInicio + resMostrar 
  if (sFin > ultimaFila){
    sFin = ultimaFila;
  }
    
  var grdResultados = container.createGrid(5, 3)
  .setBorderWidth(0)
  .setCellSpacing(1)
  .setCellPadding(1);
  
  //Cuadro de Texto Palabra Buscar
  var txtBuscar = container.createTextBox()
  .setName("txtBuscar")
  .setVisible(false)
  .setText(palabra_Buscar);
     
  vPanel.add(txtBuscar)
  
  //Cuadro de Texto Serie
  var txtSerie_Buscar = container.createTextBox()
  .setName("txtSerie_Buscar")
  .setVisible(false)
  .setText(serie);
  
  vPanel.add(txtSerie_Buscar)
  
  
  //Cargar Resultados
  for (ii = sInicio; ii < sFin; ii++){
    //Agregar contenedores
    op = op + 1;
    //Segun op son los nombres de los contenedores
    switch(op){
      case 1:
        var txtActivo1 = container.createTextBox()
        .setText(tResultado[ii][5].toString())
        .setName("txtActivo")
        .setVisible(false);
        var txtSerie1 = container.createTextBox()
        .setName("txtSerie")
        .setText(tResultado[ii][4].toString())
        .setVisible(false);
        vPanel.add(txtActivo1);
        vPanel.add(txtSerie1);
                
        var lanzador_1 = container.createServerHandler('clickMov_')
        .addCallbackElement(txtActivo1)
        .addCallbackElement(txtSerie1)
        .addCallbackElement(txtBuscar)
        .addCallbackElement(txtSerie_Buscar);
        
        var btnAgregar_1 = container.createImage(IMGEDITAR)
        .addClickHandler(lanzador_1);
                    
        varMostrar_(ii);
        
        var prrText_1 = container.createTextArea().setName("prrText_1")
        .setText(txtMostrar)
        .setStyleAttribute("font-weight", "Normal")
        .setStyleAttribute("fontSize", "8pt")
        .setReadOnly(true)
        .setVisibleLines(6)
        .setWidth("550px");
        
        grdResultados.setWidget(op -1,0,btnAgregar_1)
        grdResultados.setWidget(op - 1,1,prrText_1)
        
        if(tResultado[ii][22] != ""){
          var ligaArchivo_1 = ligaArchivo_(tResultado[ii][22])
          
        }else{
          var ligaArchivo_1 = container.createTextBox().setVisible(false).setText(tResultado[ii][22]);
        }
                
        grdResultados.setWidget(op - 1,2,ligaArchivo_1)
                
        break;
      case 2:
        var txtActivo2 = container.createTextBox()
        .setText(tResultado[ii][5].toString())
        .setName("txtActivo")
        .setVisible(false);
        var txtSerie2 = container.createTextBox()
        .setName("txtSerie")
        .setText(tResultado[ii][4].toString())
        .setVisible(false);
        vPanel.add(txtActivo2);
        vPanel.add(txtSerie2);
        
        var lanzador_2 = container.createServerHandler('clickMov_')
        .addCallbackElement(txtActivo2)
        .addCallbackElement(txtSerie2)
        .addCallbackElement(txtBuscar)
        .addCallbackElement(txtSerie_Buscar);
        
        var btnAgregar_2 = container.createImage(IMGEDITAR)
        .addClickHandler(lanzador_2);
                    
        varMostrar_(ii);
        
        var prrText_2 = container.createTextArea().setName("prrText_2")
        .setText(txtMostrar)
        .setStyleAttribute("font-weight", "Normal")
        .setStyleAttribute("fontSize", "8pt")
        .setReadOnly(true)
        .setVisibleLines(6)
        .setWidth("550px");
                
        grdResultados.setWidget(op -1,0,btnAgregar_2)
        grdResultados.setWidget(op - 1,1,prrText_2)
        
        if(tResultado[ii][22] != ""){
          var ligaArchivo_2 = ligaArchivo_(tResultado[ii][22])
          
        }else{
          var ligaArchivo_2 = container.createTextBox().setVisible(false).setText(tResultado[ii][22]);
        }
                
        grdResultados.setWidget(op - 1,2,ligaArchivo_2)
        
        break;
      case 3:
        var txtActivo3 = container.createTextBox()
        .setText(tResultado[ii][5].toString())
        .setName("txtActivo")
        .setVisible(false);
        var txtSerie3 = container.createTextBox()
        .setName("txtSerie")
        .setText(tResultado[ii][4].toString())
        .setVisible(false);
        vPanel.add(txtActivo3);
        vPanel.add(txtSerie3);
        
        var lanzador_3 = container.createServerHandler('clickMov_')
        .addCallbackElement(txtActivo3)
        .addCallbackElement(txtSerie3)
        .addCallbackElement(txtBuscar)
        .addCallbackElement(txtSerie_Buscar);
        
        var btnAgregar_3 = container.createImage(IMGEDITAR)
        .addClickHandler(lanzador_3);
                    
        varMostrar_(ii);
        
        var prrText_3 = container.createTextArea().setName("prrText_3")
        .setText(txtMostrar)
        .setStyleAttribute("font-weight", "Normal")
        .setStyleAttribute("fontSize", "8pt")
        .setReadOnly(true)
        .setVisibleLines(6)
        .setWidth("550px");
                
        grdResultados.setWidget(op -1,0,btnAgregar_3)
        grdResultados.setWidget(op - 1,1,prrText_3)
        
        if(tResultado[ii][22] != ""){
          var ligaArchivo_3 = ligaArchivo_(tResultado[ii][22])
          
        }else{
          var ligaArchivo_3 = container.createTextBox().setVisible(false).setText(tResultado[ii][22]);
        }
                
        grdResultados.setWidget(op - 1,2,ligaArchivo_3)
        
        break;
      case 4:
        var txtActivo4 = container.createTextBox()
        .setText(tResultado[ii][5].toString())
        .setName("txtActivo")
        .setVisible(false);
        var txtSerie4 = container.createTextBox()
        .setName("txtSerie")
        .setText(tResultado[ii][4].toString())
        .setVisible(false);
        vPanel.add(txtActivo4);
        vPanel.add(txtSerie4);
        
        var lanzador_4 = container.createServerHandler('clickMov_')
        .addCallbackElement(txtActivo4)
        .addCallbackElement(txtSerie4)
        .addCallbackElement(txtBuscar)
        .addCallbackElement(txtSerie_Buscar);
        
        var btnAgregar_4 = container.createImage(IMGEDITAR)
        .addClickHandler(lanzador_4);
                    
        varMostrar_(ii);
        
        var prrText_4 = container.createTextArea().setName("prrText_4")
        .setText(txtMostrar)
        .setStyleAttribute("font-weight", "Normal")
        .setStyleAttribute("fontSize", "8pt")
        .setReadOnly(true)
        .setVisibleLines(6)
        .setWidth("550px");
                
        grdResultados.setWidget(op -1,0,btnAgregar_4)
        grdResultados.setWidget(op - 1,1,prrText_4)
        
        if(tResultado[ii][22] != ""){
          var ligaArchivo_4 = ligaArchivo_(tResultado[ii][22])
          
        }else{
          var ligaArchivo_4 = container.createTextBox().setVisible(false).setText(tResultado[ii][22]);
        }
                
        grdResultados.setWidget(op - 1,2,ligaArchivo_4)
        
        break;
        
      case 5:
        var txtActivo5 = container.createTextBox()
        .setText(tResultado[ii][5].toString())
        .setName("txtActivo")
        .setVisible(false);
        var txtSerie5 = container.createTextBox()
        .setName("txtSerie")
        .setText(tResultado[ii][4].toString())
        .setVisible(false);
        vPanel.add(txtActivo5);
        vPanel.add(txtSerie5);
        
        var lanzador_5 = container.createServerHandler('clickMov_')
        .addCallbackElement(txtActivo5)
        .addCallbackElement(txtSerie5)
        .addCallbackElement(txtBuscar)
        .addCallbackElement(txtSerie_Buscar);
        
        var btnAgregar_5 = container.createImage(IMGEDITAR)
        .addClickHandler(lanzador_5);
                    
        varMostrar_(ii);
        
        var prrText_5 = container.createTextArea().setName("prrText_5")
        .setText(txtMostrar)
        .setStyleAttribute("font-weight", "Normal")
        .setStyleAttribute("fontSize", "8pt")
        .setReadOnly(true)
        .setVisibleLines(6)
        .setWidth("550px");
                
        grdResultados.setWidget(op -1,0,btnAgregar_5)
        grdResultados.setWidget(op - 1,1,prrText_5)
        
        if(tResultado[ii][22] != ""){
          var ligaArchivo_5 = ligaArchivo_(tResultado[ii][22])
          
        }else{
          var ligaArchivo_5 = container.createTextBox().setVisible(false).setText(tResultado[ii][22]);
        }
                
        grdResultados.setWidget(op - 1,2,ligaArchivo_5)
        
        break;
    }  
  }
  vPanel.add(grdResultados);
  //etiqueta de Resultados
  var lblResultados = container.createLabel().setText("Resultados del " + (sInicio + 1)  + " al " + (sFin) + " de un total de " + ultimaFila + " registros");
  styleSublbl_(lblResultados);
  vPanel.add(lblResultados);
  
  //Crear Panel Horizontal
      var hPanel = container.createHorizontalPanel();
  
  //Boton de Inicio
  var btnRegresar = container.createButton()
  .setText("Nueva Busqueda")
  .setFocus(true);
  
  // Accion Boton nueva Busqueda
  var shandlerInicio = container.createServerHandler('clickInicio_')
  .addCallbackElement(vPanel);
  btnRegresar.addClickHandler(shandlerInicio);
  styleBtn_(btnRegresar);
  hPanel.add(btnRegresar);
  
  //Boton atras
  var btnAtras = container.createButton()
  .setText("Atras");
    
  var ssig = serie - resMostrar;
  var txtsRegresar = container.createTextBox()
  .setName("txtsRegresar")
  .setVisible(false)
  .setText(ssig);
  
  vPanel.add(txtsRegresar)
    
  // Accion Boton nueva Busqueda
  var shandlerAtras = container.createServerHandler('clickRegresarResultados_')
  .addCallbackElement(vPanel);
  btnAtras.addClickHandler(shandlerAtras);
    
  if (sInicio != 0){
    btnAtras.setEnabled(true);
    styleBtn_(btnAtras);
  } else {
    btnAtras.setEnabled(false);
    styleBtnFalse_(btnAtras);
  }
  
  hPanel.add(btnAtras)
  
  //Boton Adelante
  var btnAdelante = container.createButton()
  .setText("Adelante");
  var ssig = 0;
  ssig = serie + resMostrar;
  var txtsAvanzar = container.createTextBox()
  .setName("txtsAvanzar")
  .setVisible(false)
  .setText(ssig);
  
  // Accion Boton Adelante
  var shandlerAdelante = container.createServerHandler('clickAvanzaResultados_')
  .addCallbackElement(vPanel);
  btnAdelante.addClickHandler(shandlerAdelante);
  
  vPanel.add(txtsAvanzar);
  
  if (sFin != ultimaFila){
    btnAdelante.setEnabled(true);
    styleBtn_(btnAdelante);
  } else {
    btnAdelante.setEnabled(false);
    styleBtnFalse_(btnAdelante);
  }
  
  hPanel.add(btnAdelante);
  
  
  //Boton Reporte
  var btnReporte = container.createButton()
  .setText("Reporte");
  
  // Accion Boton Reporte
  var shandlerReporte = container.createServerHandler('clickReporte_')
  .addCallbackElement(vPanel);
  btnReporte.addClickHandler(shandlerReporte);
  
  styleBtn_(btnReporte);
  
  hPanel.add(btnReporte);
  
  vPanel.add(hPanel);
  
  //Imagen
  var imgTitulo = container.createImage(URLIMAGEN)
  .setStyleAttribute("margin-top","20px");
  
  //Mostrar Controles
  vPanel.add(imgTitulo);
  
  //Agregar a APP
  container.add(vPanel);
}

function frmError_(container, error, extras, palabra_Buscar, serie_Buscar, tTemporal){
    
  //Crear Panel Vertical
  var vPanel = container.createVerticalPanel()
  .setWidth("810px")
  .setStyleAttribute("padding","20px")
  .setStyleAttribute("fontSize", "12pt")
  .setId("panel");
  
  switch(error){
    case "1": //Error al mandar buscar_Palabra vacio
     
      // Mensaje Inicial
      var lblMensajeError = container.createLabel()
      .setId("lblMensajeError")
      .setText("Error: Debe de introducir una palabra para realizar la busqueda.");
      
      //Style
      styleError_(lblMensajeError);
        
      //Imagen
      var imgFondo = container.createImage(URLIMAGEN)
      .setStyleAttribute("margin-top","20px");
      
      //Crear Panel Horizontal
      var hPanel = container.createHorizontalPanel();
      
      //Boton Nueva Busqueda 
      var btnBuscar = container.createButton()
      .setText("Nueva Busqueda");
      
      //Boton Nuevo  
      var btnNuevo = container.createButton()
      .setText("Crear Registro");
  
      //Style
      styleBtn_(btnBuscar);
      styleBtn_(btnNuevo);  
  
      // Accion Boton Buscar
      var lnzBuscar = container.createServerHandler('clickInicio_')
      .addCallbackElement(vPanel);
      btnBuscar.addClickHandler(lnzBuscar);
  
      // Accion Boton Nuevo
      var lnzNuevo = container.createServerHandler('clickNuevo_')
      .addCallbackElement(vPanel);
      btnNuevo.addClickHandler(lnzNuevo);
      
      //Agregar al Panel Horizontal
      hPanel.add(btnBuscar);
      hPanel.add(btnNuevo);
      
      // Agrega a Panel Vertical
      vPanel.add(lblMensajeError);
      vPanel.add(hPanel);
      vPanel.add(imgFondo);
 
      break;
    case "2": //No se encontro la palabra buscada
            
      // Mensaje Inicial
      var lblMensajeError = container.createLabel()
      .setId("lblMensajeError")
      .setText("Error: No se encontro registros que contengan la palabra: " + extras);
      
      //Style
      styleError_(lblMensajeError);
        
      //Imagen
      var imgFondo = container.createImage(URLIMAGEN)
      .setStyleAttribute("margin-top","20px");
      
      //Crear Panel Horizontal
      var hPanel = container.createHorizontalPanel();
      
      //Boton Nueva Busqueda 
      var btnBuscar = container.createButton()
      .setText("Nueva Busqueda");
      
      //Boton Nuevo  
      var btnNuevo = container.createButton()
      .setText("Crear Registro");
  
      //Style
      styleBtn_(btnBuscar);
      styleBtn_(btnNuevo);  
  
      // Accion Boton Buscar
      var lnzBuscar = container.createServerHandler('clickInicio_')
      .addCallbackElement(vPanel);
      btnBuscar.addClickHandler(lnzBuscar);
  
      // Accion Boton Nuevo
      var lnzNuevo = container.createServerHandler('clickNuevo_')
      .addCallbackElement(vPanel);
      btnNuevo.addClickHandler(lnzNuevo);
      //Agregar al Panel Horizontal
      hPanel.add(btnBuscar);
      hPanel.add(btnNuevo);
      
      // Agrega a Panel Vertical
      vPanel.add(lblMensajeError);
      vPanel.add(hPanel);
      vPanel.add(imgFondo);
      
      //Agrega a FrmInicio
        
      break;
    case "3": //No se pudo crear uno nuevo con los datos proporcionados
            
      // Mensaje Inicial
      var lblMensajeError = container.createLabel()
      .setId("lblMensajeError")
      .setText("Error: No se encontro información suficiente para genera un registro de movimiento a este articulo.");
      
      //Style
      styleError_(lblMensajeError);
        
      //Imagen
      var imgFondo = container.createImage(URLIMAGEN)
      .setStyleAttribute("margin-top","20px");
      
       //Cuadro de Texto Palabra Buscar
      var txtBuscar = container.createTextBox()
      .setName("txtBuscar")
      .setVisible(false)
      .setText(palabra_Buscar);
     
      vPanel.add(txtBuscar);
  
      //Cuadro de Texto Serie
      var txtsRegresar = container.createTextBox()
      .setName("txtsRegresar")
      .setVisible(false)
      .setText(serie_Buscar);
  
      vPanel.add(txtsRegresar);
                 
      //Crear Panel Horizontal
      var hPanel = container.createHorizontalPanel();
      
      //Boton Regresar 
      var btnRegresar = container.createButton()
      .setText("Regresar");
      
      //Boton Nuevo  
      var btnNuevo = container.createButton()
      .setText("Crear Registro");
  
      //Style
      styleBtn_(btnRegresar);
      styleBtn_(btnNuevo);  
  
      // Accion Boton Buscar
      var lnzRegresar = container.createServerHandler('clickRegresarResultados_')
      .addCallbackElement(vPanel);
      btnRegresar.addClickHandler(lnzRegresar);
  
      // Accion Boton Nuevo
      var lnzNuevo = container.createServerHandler('clickNuevo_')
      .addCallbackElement(vPanel);
      btnNuevo.addClickHandler(lnzNuevo);
      //Agregar al Panel Horizontal
      hPanel.add(btnRegresar);
      hPanel.add(btnNuevo);
      
      // Agrega a Panel Vertical
      vPanel.add(lblMensajeError);
      vPanel.add(hPanel);
      vPanel.add(imgFondo);
      
      //Agrega a FrmInicio
        
      break;
    case "4": //Error no selecciono un tipo de Movimiento
            
      // Mensaje Inicial
      var lblMensajeError = container.createLabel()
      .setId("lblMensajeError")
      .setText("Error: No se selecciono un tipo de movimiento para este articulo.");
      
      //Style
      styleError_(lblMensajeError);
        
      //Imagen
      var imgFondo = container.createImage(URLIMAGEN)
      .setStyleAttribute("margin-top","20px");
      
      var txtBuscar = container.createTextBox()
      .setName("txtBuscar")
      .setVisible(false)
      .setText(palabra_Buscar);
     
      vPanel.add(txtBuscar);
      
      var txtSerie_Buscar = container.createTextBox()
      .setName("txtSerie_Buscar")
      .setVisible(false)
      .setText(serie_Buscar);
     
      vPanel.add(txtSerie_Buscar);
      
      var txtArticulo = container.createTextBox()
      .setName("txtArticulo")
      .setVisible(false)
      .setText(tTemporal[2].toString());
     
      vPanel.add(txtArticulo);
      
      var txtMarca = container.createTextBox()
      .setName("txtMarca")
      .setVisible(false)
      .setText(tTemporal[3].toString());
     
      vPanel.add(txtMarca);
      
      var txtModelo = container.createTextBox()
      .setName("txtModelo")
      .setVisible(false)
      .setText(tTemporal[4].toString());
     
      vPanel.add(txtModelo);
      
      var txtSerie = container.createTextBox()
      .setName("txtSerie")
      .setVisible(false)
      .setText(tTemporal[5].toString());
     
      vPanel.add(txtSerie);
      
      var txtActivo = container.createTextBox()
      .setName("txtActivo")
      .setVisible(false)
      .setText(tTemporal[6].toString());
     
      vPanel.add(txtActivo);
      
      var txtIdentificador = container.createTextBox()
      .setName("txtIdentificador")
      .setVisible(false)
      .setText(tTemporal[7].toString());
     
      vPanel.add(txtIdentificador);
      
      var txtCondicion = container.createTextBox()
      .setName("txtCondicion")
      .setVisible(false)
      .setText(tTemporal[8].toString());
     
      vPanel.add(txtCondicion);
      
      var txtTipoMov = container.createTextBox()
      .setName("txtTipoMov")
      .setVisible(false)
      .setText(tTemporal[9].toString());
     
      vPanel.add(txtTipoMov);
                 
      //Crear Panel Horizontal
      var hPanel = container.createHorizontalPanel();
      
      //Boton Regresar 
      var btnRegresar = container.createButton()
      .setText("Regresar");
      
      var txtOpc = container.createTextBox()
      .setName("txtOpc")
      .setVisible(false)
      .setText(tTemporal[10].toString());
     
      vPanel.add(txtOpc);
      
      //Style
      styleBtn_(btnRegresar);
      
      // Accion Boton Buscar
      var lnzRegresar = container.createServerHandler('clickRegresarPaso1_')
      .addCallbackElement(vPanel);
      btnRegresar.addClickHandler(lnzRegresar);
  
      //Agregar al Panel Horizontal
      hPanel.add(btnRegresar);
            
      // Agrega a Panel Vertical
      vPanel.add(lblMensajeError);
      vPanel.add(hPanel);
      vPanel.add(imgFondo);
      
      //Agrega a FrmInicio
        
      break;
    case "5": //Final
     
      // Mensaje Inicial
      var lblMensajeError = container.createLabel()
      .setId("lblMensajeError")
      .setText("Mensaje: El movimiento fue guardado con éxito.");
      
      //Style
       styleTitulo_(lblMensajeError);
     
      var campo_Fecha = Utilities.formatDate(tTemporal[0][0],Session.getTimeZone(), "dd/MM/yyyy' 'HH:mm:ss")
     
      
      var lblMensajeError2 = container.createLabel()
      .setId("lblMensajeError2")
      .setText("Fecha: " + campo_Fecha);
      
      //Style
       styleLbl_(lblMensajeError2);
      
      var lblMensajeError3 = container.createLabel()
      .setId("lblMensajeError3")
      .setText("Usuario: " + tTemporal[0][1].toString());
      
      //Style
       styleLbl_(lblMensajeError3);
      
      
      //Agregar Liga de documento pobatorio y la de correo
      var aDocProb = container.createAnchor("Click Aqui para agregar Documento Probatorio", "https://sites.google.com/a/transpais.com.mx/ticrlactivos/docprobatorio");
      styleLbl_(aDocProb);
      
      var aCorreo = container.createAnchor("Click Aqui para enviar Correo a Activos", "https://sites.google.com/a/transpais.com.mx/ticrlactivos/enviarcorreo");
      styleLbl_(aCorreo);
            
      if(tTemporal[0][8] == '3' || tTemporal[0][8] == '5' || tTemporal[0][8] == '6'){
        lblMensajeError3.setStyleAttribute("margin-bottom","20px");
        aDocProb.setVisible(true);
        aCorreo.setVisible(true);
      }else{
        aDocProb.setVisible(false);
        aCorreo.setVisible(false);
      }
      
      //Imagen
      var imgFondo = container.createImage(URLIMAGEN)
      .setStyleAttribute("margin-top","20px");
      
      //Crear Panel Horizontal
      var hPanel = container.createHorizontalPanel();
      
      //Boton Nueva Busqueda 
      var btnBuscar = container.createButton()
      .setText("Aceptar");
       
      //Style
      styleBtn_(btnBuscar);
      
  
      // Accion Boton Buscar
      var lnzBuscar = container.createServerHandler('clickInicio_')
      .addCallbackElement(vPanel);
      btnBuscar.addClickHandler(lnzBuscar);
        
      
      //Agregar al Panel Horizontal
      hPanel.add(btnBuscar);
      
      
      // Agrega a Panel Vertical
      vPanel.add(lblMensajeError);
      vPanel.add(lblMensajeError2);
      vPanel.add(lblMensajeError3);
      vPanel.add(aDocProb);
      vPanel.add(aCorreo);
      vPanel.add(hPanel);
      vPanel.add(imgFondo);
 
      break;
    case "6": //Ver Reporte
            
      // Mensaje Inicial
      var lblMensajeError = container.createLabel()
      .setId("lblMensajeError")
      .setText("Reporte de Control de Activos");
      
      //Style
      styleTitulo_(lblMensajeError);
      
      var anchor = container.createAnchor("Click Aqui para visualizar Reporte", "https://docs.google.com/document/d/1EHUnXdXJkoVr1g_GALi6MpZiU6y-45BfeYWscNAp3us/edit");
              
      //Imagen
      var imgFondo = container.createImage(URLIMAGEN)
      .setStyleAttribute("margin-top","20px");
      
       //Cuadro de Texto Palabra Buscar
      var txtBuscar = container.createTextBox()
      .setName("txtBuscar")
      .setVisible(false)
      .setText(palabra_Buscar);
     
      vPanel.add(txtBuscar);
  
      //Cuadro de Texto Serie
      var txtsRegresar = container.createTextBox()
      .setName("txtsRegresar")
      .setVisible(false)
      .setText(serie_Buscar);
  
      vPanel.add(txtsRegresar);
                 
      //Crear Panel Horizontal
      var hPanel = container.createHorizontalPanel();
      
      //Boton Regresar 
      var btnRegresar = container.createButton()
      .setText("Regresar");
     
      //Style
      styleBtn_(btnRegresar);
       
      // Accion Boton Buscar
      var lnzRegresar = container.createServerHandler('clickRegresarResultados_')
      .addCallbackElement(vPanel);
      btnRegresar.addClickHandler(lnzRegresar);
  
      //Agregar al Panel Horizontal
      hPanel.add(btnRegresar);
      
      // Agrega a Panel Vertical
      vPanel.add(lblMensajeError);
      vPanel.add(anchor);
      vPanel.add(hPanel);
      vPanel.add(imgFondo);
      
      //Agrega a FrmInicio
        
      break;
    
  }
  
  //Agregar a APP
  container.add(vPanel);
}

function frmRegPas1_(container, opc, palabra_Buscar, serie_Buscar, ext){
  //Crear Panel Vertical
  var vPanel = container.createVerticalPanel()
  .setWidth("810px")
  .setStyleAttribute("padding","20px")
  .setStyleAttribute("fontSize", "12pt")
  .setId("panel");
  
   // Titulo
  var lblTitulo = container.createLabel()
  .setId("lblTitulo")
  .setText("Crear Registro");
  styleTitulo_(lblTitulo);
  vPanel.add(lblTitulo);
  
  // Sub label
  var lblsInfo = container.createLabel()
  .setId("lblsInfo");
  
  if (opc == 0){
    lblsInfo.setText("Ingrese los datos para la creación de un nuevo registro");
  } else {
    lblsInfo.setText(tActivo[7].toString());
  }
  
  styleSublbl_(lblsInfo);
  vPanel.add(lblsInfo);
  
  //Cuadro de Texto Palabra Buscar
  var txtBuscar = container.createTextBox()
  .setName("txtBuscar")
  .setVisible(false)
  .setText(palabra_Buscar);
  
  vPanel.add(txtBuscar);
  
  //Cuadro de Texto Serie
  var txtsRegresar = container.createTextBox()
  .setName("txtsRegresar")
  .setVisible(false)
  .setText(serie_Buscar);
  
  vPanel.add(txtsRegresar);
  
  // Etiqueta Articulos
  var lblArticulo = container.createLabel()
  .setId("lblArticulo")
  .setText("Articulo:");
  styleLbl_(lblArticulo);  
  vPanel.add(lblArticulo);
  
  // Sub-Etiqueta Articulo
  var slblArticulo = container.createLabel()
  .setId("slblArticulo")
  .setText("Seleccione el articulo a registrar");
  styleSublbl_(slblArticulo);
  vPanel.add(slblArticulo);
  
  // List Box Articulo
  var lbxArticulo = container.createListBox()
  .setName("lbxArticulo")
  .setVisibleItemCount(1)
  .setWidth(TEXTBOX_WIDTH)
  .setFocus(true);
  
  //Cargar Articulos
  var Catalogos  = SpreadsheetApp.openById(CATALOGO).getSheetByName("Articulos");
  var UltimaFila = Catalogos.getLastRow();
  var rango = Catalogos.getRange("A2:B" + UltimaFila);
  var valores = rango.getValues();
  for(var i1= 0; i1 < valores.length ; i1++){
    lbxArticulo.addItem(valores[i1][1].toString(), valores[i1][0].toString());
  }
  
  if (opc == 0){
    lbxArticulo.setSelectedIndex(0);
    if (ext == 1){
      lbxArticulo.setSelectedIndex(tActivo[0]);
    }
  } else {
    lbxArticulo.setSelectedIndex(tActivo[0]);
  }
  
  
  
  vPanel.add(lbxArticulo);
  
   // Etiqueta Marca
  var lblMarca = container.createLabel()
    .setId("lblMarca")
  .setText("Marca:");
  styleLbl_(lblMarca);
  vPanel.add(lblMarca);
  
  // Sub-Etiqueta
  var slblMarca = container.createLabel()
  .setId("slblMarca")
  .setText("Seleccione marca del articulo");
  styleSublbl_(slblMarca);
  vPanel.add(slblMarca);
  
  // List Box Marca
  var lbxMarca = container.createListBox()
  .setName("lbxMarca")
  .setVisibleItemCount(1)
  .setWidth(TEXTBOX_WIDTH);
  
  //Cargar Marcas
  Catalogos  = SpreadsheetApp.openById(CATALOGO).getSheetByName("Marcas");
  UltimaFila = Catalogos.getLastRow();
  rango = Catalogos.getRange("A2:B" + UltimaFila);
  valores = rango.getValues();
  for(var i1= 0; i1 < valores.length ; i1++){
    lbxMarca.addItem(valores[i1][1].toString(), valores[i1][0].toString());
  }
  
  if (opc == 0){
    lbxMarca.setSelectedIndex(0);
    if (ext == 1){
      lbxMarca.setSelectedIndex(tActivo[1]);
    }
  } else {
    lbxMarca.setSelectedIndex(tActivo[1]);
  }
  
  vPanel.add(lbxMarca);
  
  // Etiqueta Modelo
  var lblModelo = container.createLabel()
  .setId("lblModelo")
  .setText("Modelo:");
  styleLbl_(lblModelo);
  vPanel.add(lblModelo);
  
  // Sub-Etiqueta
  var slblModelo = container.createLabel()
  .setId("slblModelo")
  .setText("Introduzca el modelo del articulo");
  styleSublbl_(slblModelo);
  vPanel.add(slblModelo);
  
  // Text Box Modelo
  var txtModelo = container.createTextBox()
  .setName("txtModelo");
  styleTxt_(txtModelo)
     
  if (opc == 0){
    txtModelo.setText("");
    if (ext == 1){
      txtModelo.setText(tActivo[2].toString());
    }
  } else {
    txtModelo.setText(tActivo[2].toString());
  }
      
  vPanel.add(txtModelo);
  
  // Etiqueta Serie
  var lblSerie = container.createLabel()
  .setId("lblSerie")
  .setText("No. de Serie:");
  styleLbl_(lblSerie);
  vPanel.add(lblSerie);
  
  // Sub-Etiqueta Serie
  var slblSerie = container.createLabel()
  .setId("slblSerie")
  .setText("Introduzca el no. de serie del articulo");
  styleSublbl_(slblSerie);
  vPanel.add(slblSerie);
  
  // Text Box Serie
  var txtSerie = container.createTextBox()
  .setName("txtSerie");
  styleTxt_(txtSerie);
  
  if (opc == 0){
    txtSerie.setText("");
    if (ext == 1){
      txtSerie.setText(tActivo[3].toString());
    }
  } else {
    txtSerie.setText(tActivo[3].toString());
  }
  
  vPanel.add(txtSerie);
  
  // Etiqueta Activo
  var lblActivo = container.createLabel()
  .setId("lblActivo")
  .setText("No. de Activo:");
  styleLbl_(lblActivo);
  vPanel.add(lblActivo);
  
  // Sub-Etiqueta Activo
  var slblActivo = container.createLabel()
  .setId("slblActivo")
  .setText("Introduzca el no. de activo del articulo");
  styleSublbl_(slblActivo);
  vPanel.add(slblActivo);
  
  // Text Box Serie
  var txtActivo = container.createTextBox()
  .setName("txtActivo");
  styleTxt_(txtActivo);
     
  if (opc == 0){
    txtActivo.setText("");
    if (ext == 1){
      txtActivo.setText(tActivo[4].toString());
    }
  } else {
    txtActivo.setText(tActivo[4].toString());
  }
  
  vPanel.add(txtActivo);
  
  // Etiqueta Identificador
  var lblIdentificador = container.createLabel()
  .setId("lblIdentificador")
  .setText("Identificador:");
  styleLbl_(lblIdentificador);
  vPanel.add(lblIdentificador);
  
  // Sub-Etiqueta Identificador
  var slblIdentificador = container.createLabel()
  .setId("slblIdentificador")
  .setText("Introduzca identificador Ejemplo: VICCRTI_ARMANDO");
  styleSublbl_(slblIdentificador);
  vPanel.add(slblIdentificador);
  
  // Text Box Identificador
  var txtIdentificador = container.createTextBox()
  .setName("txtIdentificador");
  styleTxt_(txtIdentificador); 
  
  if (opc == 0){
    txtIdentificador.setText("");
    if (ext == 1){
      txtIdentificador.setText(tActivo[5].toString());
    }
  } else {
    txtIdentificador.setText(tActivo[5].toString());
  }
  
  vPanel.add(txtIdentificador);
 
  // Etiqueta Condicion
  var lblCondicion = container.createLabel()
  .setId("lblCondicion")
  .setText("Condición:")
  .setVisible(true);
  styleLbl_(lblCondicion);
  vPanel.add(lblCondicion);
  
  // Sub-Etiqueta Condición
  var slblCondicion = container.createLabel()
  .setId("slblCondicion")
  .setText("Seleccione condición actual del articulo");
  styleSublbl_(slblCondicion);
  vPanel.add(slblCondicion);
  
  // List Box Condicion
  var lbxCondicion = container.createListBox()
  .setName("lbxCondicion")
  .setVisibleItemCount(1)
  .setWidth(TEXTBOX_WIDTH);
  
  //Cargar Articulos
  var Catalogos  = SpreadsheetApp.openById(CATALOGO).getSheetByName("Condiciones");
  var UltimaFila = Catalogos.getLastRow();
  var rango = Catalogos.getRange("A2:B" + UltimaFila);
  var valores = rango.getValues();
  for(var i1= 0; i1 < valores.length ; i1++){
    lbxCondicion.addItem(valores[i1][1].toString(), valores[i1][0].toString());
  }
  
  if (opc == 0){
    lbxCondicion.setSelectedIndex(0);
    if (ext == 1){
      lbxCondicion.setSelectedIndex(tActivo[6]);
    }
  } else {
    lbxCondicion.setSelectedIndex(tActivo[6]);
  }
  
  vPanel.add(lbxCondicion);
  
  // Etiqueta 
  var lblTipoMov = container.createLabel()
  .setId("lblTipoMov")
  .setText("Tipo de Movimiento:");
  styleLbl_(lblTipoMov);
  vPanel.add(lblTipoMov);
  
  // Sub-Etiqueta
  var slblTipoMov = container.createLabel()
  .setId("slblTipoMov")
  .setText("Seleccione el tipo de movimiento del articulo");
   styleSublbl_(slblTipoMov);
  vPanel.add(slblTipoMov);
  
  // List Box 
  var lbxTipoMov = container.createListBox()
  .setName("lbxTipoMov")
  .setVisibleItemCount(1)
  .setWidth(TEXTBOX_WIDTH);
  
  //Cargar Tipos de Movimientos
  var Catalogos  = SpreadsheetApp.openById(CATALOGO).getSheetByName("Movimientos");
  var UltimaFila = Catalogos.getLastRow();
  var rango = Catalogos.getRange("A2:B" + UltimaFila);
  var valores = rango.getValues();
  for(var i1= 0; i1 < valores.length ; i1++){
    lbxTipoMov.addItem(valores[i1][1].toString(), valores[i1][0].toString());
  }
  
  if (ext == 1){
    lbxTipoMov.setSelectedIndex(tActivo[8]);
  } else {
    lbxTipoMov.setSelectedIndex(0);
  }
  
  lbxTipoMov.setSelectedIndex(0);
    
  vPanel.add(lbxTipoMov);
  
  // Text Box opc
  var txtOpc = container.createTextBox()
  .setName("txtOpc")
  .setVisible(false)
  .setText(opc); 
      
  vPanel.add(txtOpc);
  
  var hPanel = container.createHorizontalPanel()
  .setId("Panel_Controles");
  
   //Boton Regresar
  var btnRegresar = container.createButton()
  .setText("Regresar");
  styleBtn_(btnRegresar);
  
  //Accion Boton Regresar  
  if (opc == 0){
    var lnzRegresar = container.createServerHandler('clickInicio_')
  .addCallbackElement(vPanel);
  btnRegresar.addClickHandler(lnzRegresar);
  } else {
    var lnzRegresar = container.createServerHandler('clickRegresarResultados_')
  .addCallbackElement(vPanel);
  btnRegresar.addClickHandler(lnzRegresar);
  }
    
  hPanel.add(btnRegresar);
      
   //Boton Cancelar
  var btnSiguiente = container.createButton()
  .setText("Siguiente");
  styleBtn_(btnSiguiente);
  
  // Accion Boton Busqueda
  var lnzSiguiente = container.createServerHandler('clickSiguiente_')
  .addCallbackElement(vPanel);
  btnSiguiente.addClickHandler(lnzSiguiente);
   
  hPanel.add(btnSiguiente);
  
  vPanel.add(hPanel);
  
  //Imagen
  var imgTitulo = container.createImage(URLIMAGEN)
  .setStyleAttribute("margin-top","20px");
  
  vPanel.add(imgTitulo);
  
  container.add(vPanel)
  
}

function frmRegPas2_(container, opc, palabra_Buscar, serie_Buscar, tTemporal, ext){
  //Crear Panel Vertical
  var vPanel = container.createVerticalPanel()
  .setWidth("810px")
  .setStyleAttribute("padding","20px")
  .setStyleAttribute("fontSize", "12pt")
  .setId("panel");
  
   // Titulo
  var lblTitulo = container.createLabel()
  .setId("lblTitulo")
  .setText("Datos del Movimiento");
  styleTitulo_(lblTitulo);
  vPanel.add(lblTitulo);
  
  // Sub label
  var lblsInfo = container.createLabel()
  .setId("lblsInfo");
  lblsInfo.setText("Ingrese los datos del movimiento");
    
  styleSublbl_(lblsInfo);
  vPanel.add(lblsInfo);
  
  var txtBuscar = container.createTextBox()
  .setName("txtBuscar")
  .setVisible(false)
  .setText(palabra_Buscar);
  
  vPanel.add(txtBuscar);
  
  var txtSerie_Buscar = container.createTextBox()
  .setName("txtSerie_Buscar")
  .setVisible(false)
  .setText(serie_Buscar);
     
  vPanel.add(txtSerie_Buscar);
  
  var txtArticulo = container.createTextBox()
  .setName("txtArticulo")
  .setVisible(false)
  .setText(tTemporal[2].toString());
  
  vPanel.add(txtArticulo);
  
  var txtMarca = container.createTextBox()
  .setName("txtMarca")
  .setVisible(false)
  .setText(tTemporal[3].toString());
  
  vPanel.add(txtMarca);
  
  var txtModelo = container.createTextBox()
  .setName("txtModelo")
  .setVisible(false)
  .setText(tTemporal[4].toString());
  
  vPanel.add(txtModelo);
  
  var txtSerie = container.createTextBox()
  .setName("txtSerie")
  .setVisible(false)
  .setText(tTemporal[5].toString());
  
  vPanel.add(txtSerie);
  
  var txtActivo = container.createTextBox()
  .setName("txtActivo")
  .setVisible(false)
  .setText(tTemporal[6].toString());
  
  vPanel.add(txtActivo);
  
  var txtIdentificador = container.createTextBox()
  .setName("txtIdentificador")
  .setVisible(false)
  .setText(tTemporal[7].toString());
  
  vPanel.add(txtIdentificador);
      
  var txtCondicion = container.createTextBox()
  .setName("txtCondicion")
  .setVisible(false)
  .setText(tTemporal[8].toString());
  
  vPanel.add(txtCondicion);
  
  var txtTipoMov = container.createTextBox()
  .setName("txtTipoMov")
  .setVisible(false)
  .setText(tTemporal[9].toString());
     
  vPanel.add(txtTipoMov);
  
  var txtOpc = container.createTextBox()
  .setName("txtOpc")
  .setVisible(false)
  .setText(tTemporal[10].toString());
     
  vPanel.add(txtOpc);
  
  //EN STOCK
  
  // Etiqueta Asignado
  var lblAsignado = container.createLabel()
  .setId("lblAsignado")
  .setText("Asignado:")
  .setVisible(false)
  styleLbl_(lblAsignado);  
  vPanel.add(lblAsignado);
  
  // Sub-Etiqueta Asignado
  var slblAsignado = container.createLabel()
  .setId("slblAsignado")
  .setText("Seleccione si el articulo ya esta asigando")
  .setVisible(false)
  styleSublbl_(slblAsignado);
  vPanel.add(slblAsignado);
  
  // List Box Asignado
  var lbxAsignado = container.createListBox()
  .setName("lbxAsignado")
  .setVisibleItemCount(1)
  .setWidth(TEXTBOX_WIDTH)
  .setVisible(false)
  
  //Cargar Asignado
  var Catalogos  = SpreadsheetApp.openById(CATALOGO).getSheetByName("Stock");
  var UltimaFila = Catalogos.getLastRow();
  var rango = Catalogos.getRange("A2:B" + UltimaFila);
  var valores = rango.getValues();
  for(var i1= 0; i1 < valores.length ; i1++){
    lbxAsignado.addItem(valores[i1][1].toString());
  }
  lbxAsignado.setSelectedIndex(0);
  vPanel.add(lbxAsignado);
  
   // Etiqueta Alamacen
  var lblAlmacen = container.createLabel()
  .setId("lblAlmacen")
  .setText("Almacen:")
  .setVisible(false)
  styleLbl_(lblAlmacen);  
  vPanel.add(lblAlmacen);
  
  // Sub-Etiqueta Almacen
  var slblAlmacen = container.createLabel()
  .setId("slblAlmacen")
  .setText("Seleccione el almacen al que sera enviado el activo")
  .setVisible(false)
  styleSublbl_(slblAlmacen);
  vPanel.add(slblAlmacen);
  
  // List Box Almacen
  var lbxAlmacen = container.createListBox()
  .setName("lbxAlmacen")
  .setVisibleItemCount(1)
  .setWidth(TEXTBOX_WIDTH)
  .setVisible(false)
  
  //Cargar Almacen
  Catalogos  = SpreadsheetApp.openById(CATALOGO).getSheetByName("Almacen");
  UltimaFila = Catalogos.getLastRow();
  rango = Catalogos.getRange("A2:B" + UltimaFila);
  valores = rango.getValues();
  for(var i1= 0; i1 < valores.length ; i1++){
    lbxAlmacen.addItem(valores[i1][1].toString());
  }
  lbxAlmacen.setSelectedIndex(0);
  vPanel.add(lbxAlmacen);
  
  //ENVIADO A SOPORTE
  
   // Etiqueta Soporte
  var lblSoporte = container.createLabel()
  .setId("lblSoporte")
  .setText("Asignado a:")
  .setVisible(false)
  styleLbl_(lblSoporte);  
  vPanel.add(lblSoporte);
  
  // Sub-Etiqueta Soporte
  var slblSoporte = container.createLabel()
  .setId("slblSoporte")
  .setText("Seleccione el Técnico asignado para soporte")
  .setVisible(false)
  styleSublbl_(slblSoporte);
  vPanel.add(slblSoporte);
  
  // List Box Soporte
  var lbxSoporte = container.createListBox()
  .setName("lbxSoporte")
  .setVisibleItemCount(1)
  .setWidth(TEXTBOX_WIDTH)
  .setVisible(false)
  
  //Cargar Soporte
  Catalogos  = SpreadsheetApp.openById(CATALOGO).getSheetByName("Soporte");
  UltimaFila = Catalogos.getLastRow();
  rango = Catalogos.getRange("A2:B" + UltimaFila);
  valores = rango.getValues();
  for(var i1= 0; i1 < valores.length ; i1++){
    lbxSoporte.addItem(valores[i1][1].toString());
  }
  lbxSoporte.setSelectedIndex(0);
  vPanel.add(lbxSoporte);
  
   //ENVIADO A USUARIO FINAL
  
  // Etiqueta Usuario
  var lblUsuario = container.createLabel()
  .setId("lblUsuario")
  .setText("Usuario:")
  .setVisible(false);
  styleLbl_(lblUsuario);
  vPanel.add(lblUsuario);
  
  // Sub-Etiqueta Usuario
  var slblUsuario = container.createLabel()
  .setId("slblUsuario")
  .setText("Introduzca el nombre de la persona a la que se envía el articulo")
  .setVisible(false);
  styleSublbl_(slblUsuario);
  vPanel.add(slblUsuario);
  
  // Text Box Usuario
  var txtUsuario = container.createTextBox()
  .setName("txtUsuario")
  .setVisible(false);
  styleTxt_(txtUsuario);
  vPanel.add(txtUsuario);
  
   // Etiqueta UEN
  var lblUEN = container.createLabel()
  .setId("lblUEN")
  .setText("Unidad de Negocio:")
  .setVisible(false)
  styleLbl_(lblUEN);  
  vPanel.add(lblUEN);
  
  // Sub-Etiqueta UEN
  var slblUEN = container.createLabel()
  .setId("slblUEN")
  .setText("Seleccione unidad de negocio")
  .setVisible(false)
  styleSublbl_(slblUEN);
  vPanel.add(slblUEN);
  
  // List Box UEN
  var lbxUEN = container.createListBox()
  .setName("lbxUEN")
  .setVisibleItemCount(1)
  .setWidth(TEXTBOX_WIDTH)
  .setVisible(false)
  
  //Cargar UEN
  Catalogos  = SpreadsheetApp.openById(CATALOGO).getSheetByName("UEN");
  UltimaFila = Catalogos.getLastRow();
  rango = Catalogos.getRange("A2:B" + UltimaFila);
  valores = rango.getValues();
  for(var i1= 0; i1 < valores.length ; i1++){
    lbxUEN.addItem(valores[i1][1].toString());
  }
  lbxUEN.setSelectedIndex(0);
  vPanel.add(lbxUEN);
  
  // Etiqueta Ubicacion
  var lblUbicacion = container.createLabel()
  .setId("lblUbicacion")
  .setText("Ubicación:")
  .setVisible(false);
  styleLbl_(lblUbicacion);
  vPanel.add(lblUbicacion);
  
  // Sub-Etiqueta Ubicacion
  var slblUbicacion = container.createLabel()
  .setId("slblUbicacion")
  .setText("Introduzca ciudad o población")
  .setVisible(false);
  styleSublbl_(slblUbicacion);
  vPanel.add(slblUbicacion);
  
  // Text Box Ubiacion
  var txtUbicacion = container.createTextBox()
  .setName("txtUbicacion")
  .setVisible(false);
  styleTxt_(txtUbicacion);
  vPanel.add(txtUbicacion);
  
  // SE AGREGO FOLIO DEL CARTAPACIO DE ACTIVOS DEL SERVICEDESK
  
  // Etiqueta Folio
  var lblFolio = container.createLabel()
  .setId("lblFolio")
  .setText("Folio:")
  .setVisible(false);
  styleLbl_(lblFolio);
  vPanel.add(lblFolio);
  
  // Sub-Etiqueta Folio
  var slblFolio = container.createLabel()
  .setId("slblFolio")
  .setText("Introduzca Folio de recibido")
  .setVisible(false);
  styleSublbl_(slblFolio);
  vPanel.add(slblFolio);
  
  // Text Box Ubiacion
  var txtFolio = container.createTextBox()
  .setName("txtFolio")
  .setVisible(false);
  styleTxt_(txtFolio);
  vPanel.add(txtFolio);
  
  //ENVIADO A PROVEEDOR
  
  // Etiqueta Proveedor
  var lblProveedor = container.createLabel()
  .setId("lblProveedor")
  .setText("Proveedor:")
  .setVisible(false);
  styleLbl_(lblProveedor);
  vPanel.add(lblProveedor);
  
  // Sub-Etiqueta Proveedor
  var slblProveedor = container.createLabel()
  .setId("slblProeveedor")
  .setText("Introduzca nombre o razón social del proveedor")
  .setVisible(false);
  styleSublbl_(slblProveedor);
  vPanel.add(slblProveedor);
  
  // Text Box Proveedor
  var txtProveedor = container.createTextBox()
  .setName("txtProveedor")
  .setVisible(false);
  styleTxt_(txtProveedor); 
  vPanel.add(txtProveedor);
  
  // Etiqueta Notas
  var lblNotas = container.createLabel()
  .setId("lblNotas")
  .setText("Notas:")
  .setVisible(false);
  styleLbl_(lblNotas);
  vPanel.add(lblNotas);
    
  // Text Box Proveedor
  var txtNotas = container.createTextBox()
  .setName("txtNotas")
  .setVisible(false);
  styleTxt_(txtNotas); 
  vPanel.add(txtNotas);
  
  //Mostrar los movimientos
  switch (tTemporal[9].toString()) {
    case "1":
      lblAsignado.setVisible(true);
      slblAsignado.setVisible(true);
      lbxAsignado.setVisible(true);
      lbxAsignado.setFocus(true);
      lblAlmacen.setVisible(true);
      slblAlmacen.setVisible(true);
      lbxAlmacen.setVisible(true);
      break;
    case "2":
      lblSoporte.setVisible(true);
      slblSoporte.setVisible(true);
      lbxSoporte.setVisible(true);
      lbxSoporte.setFocus(true);
      break;
    case "3":
      lblUsuario.setVisible(true);
      slblUsuario.setVisible(true);
      txtUsuario.setVisible(true);
      txtUsuario.setFocus(true);
      lblUEN.setVisible(true);
      slblUEN.setVisible(true);
      lbxUEN.setVisible(true);
      lblUbicacion.setVisible(true);
      slblUbicacion.setVisible(true);
      txtUbicacion.setVisible(true);
      lblFolio.setVisible(true);
      slblFolio.setVisible(true);
      txtFolio.setVisible(true);
      break;
    case "4":
      lblProveedor.setVisible(true);
      slblProveedor.setVisible(true);
      txtProveedor.setVisible(true);
      txtProveedor.setFocus(true);
      lblNotas.setVisible(true);
      txtNotas.setVisible(true);
      break;
    case "5":
      lblUsuario.setText("Persona que Recibe:")
      lblUsuario.setVisible(true);
      slblUsuario.setText("Introduzca el nombre de la persona que recibe el articulo en custodia.")
      slblUsuario.setVisible(true);
      txtUsuario.setVisible(true);
      txtUsuario.setFocus(true);
      lblFolio.setVisible(true);
      slblFolio.setVisible(true);
      txtFolio.setVisible(true);
      break;
    case "6":
     
      lblUsuario.setVisible(true);
      slblUsuario.setText("Introduzca el nombre de la persona que tiene en uso el articulo")
      slblUsuario.setVisible(true);
      txtUsuario.setVisible(true);
      txtUsuario.setFocus(true);
      lblUEN.setVisible(true);
      slblUEN.setVisible(true);
      lbxUEN.setVisible(true);
      lblUbicacion.setVisible(true);
      slblUbicacion.setVisible(true);
      txtUbicacion.setVisible(true);
      lblFolio.setVisible(true);
      slblFolio.setVisible(true);
      txtFolio.setVisible(true);
      break;
  }
  
  //GUIA
  // Etiqueta Guia
  var lblGuia = container.createLabel()
  .setId("lblGuia")
  .setText("Guia:");
  styleLbl_(lblGuia);
  vPanel.add(lblGuia);
  
  // Sub-Etiqueta Guia
  var slblGuia = container.createLabel()
  .setId("slblGuia")
  .setText("Introduzca guía (opcional)");
  styleSublbl_(slblGuia);
  vPanel.add(slblGuia);
  
  // Text Box Guia
  var txtGuia = container.createTextBox()
  .setName("txtGuia");
  styleTxt_(txtGuia); 
  vPanel.add(txtGuia);
  
  //TELEFONO
  // Etiqueta Telefono
  var lblTelefono = container.createLabel()
  .setId("lblTelefono")
  .setText("Ext. o Telefono:");
  styleLbl_(lblTelefono);
  vPanel.add(lblTelefono);
  
  // Sub-Etiqueta Telefono
  var slblTelefono = container.createLabel()
  .setId("slblTelefono")
  .setText("Introduzca Ext. o Teléfono (opcional)");
  styleSublbl_(slblTelefono);
  vPanel.add(slblTelefono);
  
  // Text Box Telefono
  var txtTelefono = container.createTextBox()
  .setName("txtTelefono");
  styleTxt_(txtTelefono); 
  vPanel.add(txtTelefono);
  
  //OBSERVACIONES
  // Etiqueta Observaciones
  var lblObservaciones = container.createLabel()
  .setId("lblObservaciones")
  .setText("Observaciones:");
  styleLbl_(lblObservaciones);
  vPanel.add(lblObservaciones);
  
  // Sub-Etiqueta Observaciones
  var slblObservaciones = container.createLabel()
  .setId("slblObservaciones")
  .setText("Introduzca observaciones");
  styleSublbl_(slblObservaciones);
  vPanel.add(slblObservaciones);
  
  // Text Box Observaciones
  var txtObservaciones = container.createTextBox()
  .setName("txtObservaciones");
  styleTxt_(txtObservaciones); 
  vPanel.add(txtObservaciones);
  
  //Crear Panel Horizontal
  var hPanel = container.createHorizontalPanel();
  
  //Boton Regresar 
  var btnRegresar = container.createButton()
  .setText("Regresar");
      
  //Style
  styleBtn_(btnRegresar);
  
  // Accion Boton Regresar
  var lnzRegresar = container.createServerHandler('clickRegresarPaso1_')
  .addCallbackElement(vPanel);
  btnRegresar.addClickHandler(lnzRegresar);
  
  //Agregar al Panel Horizontal
  hPanel.add(btnRegresar);
  
  //Boton Guardar 
  var btnGuardar = container.createButton()
  .setText("Guardar");
      
  //Style
  styleBtn_(btnGuardar);
  
  // Accion Boton Regresar
  var lnzGuardar = container.createServerHandler('clickGuardar_')
  .addCallbackElement(vPanel);
  btnGuardar.addClickHandler(lnzGuardar);
  
  //Agregar al Panel Horizontal
  hPanel.add(btnGuardar);
  
  vPanel.add(hPanel)
  
  //Imagen
  var imgFondo = container.createImage(URLIMAGEN)
  .setStyleAttribute("margin-top","20px");
  
  vPanel.add(imgFondo);
  
  container.add(vPanel);
}


//*-*-*-*-*-*-*-*-*-*-*-*-*-* C L I C K *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
function clickInicio_(e){
  var app = UiApp.getActiveApplication();
  var panel = app.getElementById("panel") ;
  panel.setVisible(false);
  frmInicio_(app, "0");
  //Actualizar APP
  return app;
}

function clickNuevo_(e){
  var app = UiApp.getActiveApplication();
  var panel = app.getElementById("panel") ;
  panel.setVisible(false);
  frmRegPas1_(app, "0");
  //Actualizar APP
  return app;
}

function clickBuscarGral_(e){
  var app = UiApp.getActiveApplication();
  var palabra_Buscar = e.parameter.txtBuscar.toUpperCase();
  var panel = app.getElementById("panel") ;
  panel.setVisible(false);
  if (palabra_Buscar == ""){
    //Si no se escribio una palabra
    frmError_(app, "1");
  } else {
    palabra_Buscar = trim(palabra_Buscar);
    var serie = resMostrar;
    //Buscar Palabra
    buscarPalabra_(palabra_Buscar);
    //Manda a mostrar los resultados de la busqueda
    if (typeof tResultado == 'undefined' && tResultado == null){
      frmError_(app, "2", palabra_Buscar);  
    } else {
      frmResultados_(app, e, palabra_Buscar, serie)
    }
  }
  
  //Actualizar APP
  return app;
}

function clickAvanzaResultados_(e){
  var app = UiApp.getActiveApplication();
  var panel = app.getElementById("panel") ;
  panel.setVisible(false);
  var palabra_Buscar = e.parameter.txtBuscar;
  buscarPalabra_(palabra_Buscar);
  var serie = 0;
  serie = parseInt(e.parameter.txtsAvanzar);
  frmResultados_(app, e, palabra_Buscar, serie);
  return app; 
}

function clickRegresarResultados_(e){
  var app = UiApp.getActiveApplication();
  var panel = app.getElementById("panel") ;
  panel.setVisible(false);
  var palabra_Buscar = e.parameter.txtBuscar;
  palabra_Buscar = trim(palabra_Buscar);
  buscarPalabra_(palabra_Buscar);
  var serie = 0;
  serie = parseInt(e.parameter.txtsRegresar);
  frmResultados_(app, e, palabra_Buscar, serie);
  return app; 
}

//Agregar Movimiento
function clickMov_(e){
  var app = UiApp.getActiveApplication();
  var panel = app.getElementById("panel");
  panel.setVisible(false);
  var Activo = e.parameter.txtActivo;
  var Serie = e.parameter.txtSerie;
  var palabra_Buscar = e.parameter.txtBuscar;
  var serie_Buscar = e.parameter.txtSerie_Buscar;
  if (Activo == "" && Serie == ""){
    frmError_(app, "3", "", palabra_Buscar, serie_Buscar);
  } else {
    if (Activo != ""){
      buscarActivo_(Activo, "1");
      frmRegPas1_(app, "1", palabra_Buscar, serie_Buscar);
    } else {
      buscarActivo_(Serie, "2");
      frmRegPas1_(app, "1", palabra_Buscar, serie_Buscar);
    }
  }
  return app;
}

//Siguiente Paso 1
function clickSiguiente_(e){
  var app = UiApp.getActiveApplication();
  var panel = app.getElementById("panel") ;
  panel.setVisible(false);
  var tTemporal = new Array(11);
  var palabra_buscar = e.parameter.txtBuscar;
  var serie_buscar = e.parameter.txtsRegresar;
     
  tTemporal[2] = e.parameter.lbxArticulo;
  tTemporal[3] = e.parameter.lbxMarca;
  tTemporal[4] = e.parameter.txtModelo.toUpperCase();
  tTemporal[5] = e.parameter.txtSerie.toUpperCase();
  tTemporal[6] = e.parameter.txtActivo.toUpperCase();
  tTemporal[7] = e.parameter.txtIdentificador.toUpperCase();
  tTemporal[8] = e.parameter.lbxCondicion;
  tTemporal[9] = e.parameter.lbxTipoMov;
  tTemporal[10] = e.parameter.txtOpc;
  
  if (e.parameter.lbxTipoMov == 0){
    frmError_(app, "4", "", palabra_buscar, serie_buscar, tTemporal);
  } else {
    frmRegPas2_(app, (e.parameter.txtOpc),palabra_buscar, serie_buscar, tTemporal)
  }
   
  
  //Actualizar APP
  return app;
}

//Regresar al paso 1
function clickRegresarPaso1_(e){
  var app = UiApp.getActiveApplication();
  var panel = app.getElementById("panel") ;
  panel.setVisible(false);
  var tTemporal = new Array(9);
  var palabra_Buscar = e.parameter.txtBuscar;
  var serie_Buscar = e.parameter.txtSerie_Buscar;
  var opc = e.parameter.txtOpc
    
  tTemporal[0] = parseInt(e.parameter.txtArticulo);
  tTemporal [1] = parseInt(e.parameter.txtMarca);
  tTemporal [2] = e.parameter.txtModelo;
  tTemporal [3] = e.parameter.txtSerie;
  tTemporal [4] = e.parameter.txtActivo;
  tTemporal [5] = e.parameter.txtIdentificador;
  tTemporal [6] = parseInt(e.parameter.txtCondicion);
  tTemporal [7] = "Ultimo movimiento registrado con el número de Activo " + e.parameter.txtActivo;
  tTemporal [8] = parseInt(e.parameter.txtTipoMov);
  
  tActivo = tTemporal
  
  frmRegPas1_(app, opc, palabra_Buscar, serie_Buscar, "1");
  
  //Actualizar APP
  return app;
}

function clickGuardar_(e){
  var app = UiApp.getActiveApplication();
  var panel = app.getElementById("panel") ;
  panel.setVisible(false);
  
  var tTemporal = new Array(1);
  tTemporal [0]= new Array(21);
  tTemporal [0][2] = parseInt(e.parameter.txtArticulo);
  tTemporal [0][3] = parseInt(e.parameter.txtMarca);
  tTemporal [0][4] = e.parameter.txtSerie;
  tTemporal [0][5] = e.parameter.txtActivo;
  tTemporal [0][6] = e.parameter.txtIdentificador;
  tTemporal [0][7] = parseInt(e.parameter.txtCondicion);
  tTemporal [0][8] = parseInt(e.parameter.txtTipoMov);
  tTemporal [0][9] = e.parameter.lbxAlmacen;
  tTemporal [0][10] = e.parameter.lbxSoporte;
  tTemporal [0][11] = e.parameter.txtUsuario.toUpperCase();
  tTemporal [0][12] = e.parameter.lbxUEN;
  tTemporal [0][13] = e.parameter.txtUbicacion.toUpperCase();
  tTemporal [0][14] = e.parameter.txtProveedor.toUpperCase();
  tTemporal [0][15] = e.parameter.txtNotas.toUpperCase();
  tTemporal [0][16] = e.parameter.txtGuia.toUpperCase();
  tTemporal [0][17] = e.parameter.txtTelefono.toUpperCase();
  tTemporal [0][18] = e.parameter.txtObservaciones.toUpperCase();
  tTemporal [0][19] = e.parameter.txtModelo;
  tTemporal [0][20] = e.parameter.lbxAsignado;
  tTemporal [0][21] = e.parameter.txtFolio.toUpperCase();
   
  Guardar_(tTemporal);
  
  frmError_(app, "5", "", "", "", tTemporal)
    
  //Actualizar APP
  return app;
}

function clickReporte_(e){
  var app = UiApp.getActiveApplication();
  var panel = app.getElementById("panel");
  panel.setVisible(false);
  var Activo = e.parameter.txtActivo;
  var Serie = e.parameter.txtSerie;
  var palabra_Buscar = e.parameter.txtBuscar;
  var serie_Buscar = e.parameter.txtSerie_Buscar;
  
  Reporte_(palabra_Buscar)
  
  frmError_(app, "6", "", palabra_Buscar, serie_Buscar);
  //Actualizar APP
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
  txt.setMaxLength(80);
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

//*-*-*-*-*-*-*-*-*-*-*-*-*-* F U N C I O N E S *-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
function buscarPalabra_(palabra_Buscar){
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
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange("A2:W"+lastRow);//OjO CAMBIAR CADA VEZ QUE SE AGREGA UN CAMPO
  var values = range.getValues();
  var rowCount = 0;
  var rowCount2 = 0;
  var Mov = 0;
  var cadena;
  var tTemporal;
  var ii = 0;
  var xx = 0;
  
  //Inicia Busqueda
  for(var ii= values.length - 1; ii>-1 ; ii--){
    
   //Agregamos todos los campos en los que se va buscar
    cadena = values[ii][4] + values[ii][5] + values[ii][6] +values[ii][11] +values[ii][12] + values[ii][13] + values[ii][14] + values[ii][15] + values[ii][18] + values[ii][19] + values[ii][21];    
    
    //Si se cumple la funcion
    if  (cadena.indexOf(palabra_Buscar) != -1) {

      rowCount = rowCount + 1;
      if (rowCount == 1){
       tTemporal = new Array(1)
      }
      //Agregamos una nueva matriz
      tTemporal[xx] = new Array (23);
      
      //Articulo
        for(var i1= 0; i1 < v_Articulos.length ; i1++){
          if (v_Articulos[i1][0].toString() == values[ii][2].toString()){
            Articulo = v_Articulos[i1][1].toString();
          }
        }
      
        //Marcas
        for(var i1= 0; i1 < v_Marcas.length ; i1++){
          if (v_Marcas[i1][0].toString() == values[ii][3].toString()){
            Marca = v_Marcas[i1][1].toString();
          }
        }
      
        //Movimientos
        for(var i1= 0; i1 < v_Mov.length ; i1++){
          if (v_Mov[i1][0].toString() == values[ii][8].toString()){
            Mov = v_Mov[i1][1].toString();
          }
        }
      
      //Condicion
        for(var i1= 0; i1 < v_Condiciones.length ; i1++){
          if (v_Condiciones[i1][0].toString() == values[ii][7].toString()){
            Condicion = v_Condiciones[i1][1].toString();
          }
        }
      
      //Arreglar Fecha
      
      var Fecha = Utilities.formatDate(values[ii][0],Session.getTimeZone(), "dd/MM/yyyy' 'HH:mm:ss");
      
      tTemporal[xx][0] = Fecha;
      tTemporal[xx][1] = values[ii][1].toString();
      tTemporal[xx][2] = Articulo;
      tTemporal[xx][3] = Marca;
      tTemporal[xx][4] = values[ii][4].toString();
      tTemporal[xx][5] = values[ii][5].toString();
      tTemporal[xx][6] = values[ii][6].toString();
      tTemporal[xx][7] = Condicion;
      tTemporal[xx][8] = Mov;
      tTemporal[xx][9] = values[ii][9].toString();
      tTemporal[xx][10] = values[ii][10].toString();
      tTemporal[xx][11] = values[ii][11].toString();
      tTemporal[xx][12] = values[ii][12].toString();
      tTemporal[xx][13] = values[ii][13].toString();
      tTemporal[xx][14] = values[ii][14].toString();
      tTemporal[xx][15] = values[ii][15].toString();
      tTemporal[xx][16] = values[ii][16].toString();
      tTemporal[xx][17] = values[ii][17].toString();
      tTemporal[xx][18] = values[ii][18].toString();
      tTemporal[xx][19] = values[ii][19].toString();
      tTemporal[xx][20] = values[ii][20].toString();
      tTemporal[xx][21] = values[ii][21].toString();
      tTemporal[xx][22] = values[ii][22].toString();
      
      xx++;
    }
  }
  
  tResultado = tTemporal;
}

function varMostrar_(fila){
  var texto_Mostrar_1 = "Nº ACTIVO: " + tResultado[fila][5].toString() + "\n" + tResultado[fila][2].toString() + " " + tResultado[fila][3].toString() 
        + " CON Nº SERIE: " + tResultado[fila][4].toString() + " MODELO: " + tResultado[fila][19].toString() + "\n" 
        + "IDENTIFICADOR: " + tResultado[fila][6].toString() + "\n"
        + "MOVIMIENTO: " + tResultado[fila][8].toString() + " " + tResultado[fila][20].toString() + " FECHA: " + tResultado[fila][0].toString();
  
  if (trim(tResultado[fila][8].toString()) == "EN STOCK") {
    texto_Mostrar_1 = texto_Mostrar_1 + "\n" +"Almacen: " + tResultado[fila][9].toString();
  }
  if (trim(tResultado[fila][8].toString()) == "ENVIADO A SOPORTE") {
    texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "Asignado a: " + tResultado[fila][10].toString();
  }
  if (trim(tResultado[fila][8].toString()) == "ENVIADO A USUARIO FINAL") {
    texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "Usuario Final: " + tResultado[fila][11].toString();
    texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "Unidad de Negocio: " + tResultado[fila][12].toString();
    texto_Mostrar_1 = texto_Mostrar_1 + " " + "Ubicación: " + tResultado[fila][13].toString();
    texto_Mostrar_1 = texto_Mostrar_1 + " " + " Folio: " + tResultado[fila][21].toString();
  }
  if (trim(tResultado[fila][8].toString()) == "ENVIADO A PROVEEDOR") {
    texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "Nombre del Proveedor: " + tResultado[fila][14].toString();
    texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "Notas: " + tResultado[fila][15].toString();
  }
  if (trim(tResultado[fila][8].toString()) == "EN USO") {
    texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "Usuario Final: " + tResultado[fila][11].toString();
    texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "Unidad de Negocio: " + tResultado[fila][12].toString();
    texto_Mostrar_1 = texto_Mostrar_1 + " " + "Ubicación: " + tResultado[fila][13].toString();
    texto_Mostrar_1 = texto_Mostrar_1 + " " + " Folio: " + tResultado[fila][21].toString();
  }
  if (trim(tResultado[fila][8].toString()) == "BAJA (DEP. ACTIVOS)") {
    texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "Persona que Recibe: " + tResultado[fila][11].toString();
    texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "Folio: " + tResultado[fila][21].toString();
  }
  
    texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "Guia: " + tResultado[fila][16].toString() + " Ext.: " + tResultado[fila][17].toString();
    texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "Observaciones: " + tResultado[fila][18].toString();
  txtMostrar = texto_Mostrar_1;
}

function buscarActivo_(Activo, opc){
  //Busqueda del Activo
  var tTemporal = new Array(8);
  var sheet  = SpreadsheetApp.openById(SPSHT_Base).getSheetByName(SHEET_NAME);
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange("A2:U"+lastRow);
  var values = range.getValues();
  var cual = 0;
  if (opc = "1"){
    cual = 5;
  } else {
    cual = 4;
  }
  for(var ii= values.length - 1; ii>-1 ; ii--)
    if (values[ii][cual] == Activo) {
      tTemporal[0] = values[ii][2];
      tTemporal [1] = values[ii][3];
      tTemporal [2] = values[ii][19].toString();
      tTemporal [3] = values[ii][4].toString();
      tTemporal [4] = values[ii][5].toString();
      tTemporal [5] = values[ii][6].toString();
      tTemporal [6] = values[ii][7];
      tTemporal [7] = "Ultimo movimiento registrado con el número de Activo " + values[ii][5].toString() + ": " + values[ii][0].toString();
      ii=0;    
    }
  tActivo = tTemporal;
}

function Guardar_(tTemporal){
  var sheet  = SpreadsheetApp.openById(SPSHT_Base).getSheetByName(SHEET_NAME);
  var lastRow = sheet.getLastRow() + 1;
  var range = sheet.getRange("A" + lastRow + ":V"+lastRow);// OjO CAMBIAR SI SE AGREGAN CAMBIOS
  var fecha = new Date();
  tTemporal[0][0] = fecha;
  tTemporal[0][1] = Session.getActiveUser().getEmail();
  var fila = lastRow
  //Guardar
  range.setValues(tTemporal); 
  if (tTemporal[0][8] == '3' || tTemporal[0][8] == '6' || tTemporal[0][8] == '5'){
    GuardarFila_(fila)
  }
}

function GuardarFila_(fila){
  var sheet  = SpreadsheetApp.openById(SPSHT_Temp).getSheetByName(SHEET_TEMP);
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange("A3:D"+lastRow)
  var values = range.getValues();
  var cadena;
  var FilaUsuario = 0;
  
    
  //Buscar el usuario
  var usuario =  Session.getActiveUser().getEmail();
  for(var ii= values.length - 1; ii>-1 ; ii--){
    
   //Agregamos todos los campos en los que se va buscar
    cadena = values[ii][0];
    
    //Si se cumple la funcion
    if  (cadena.indexOf(usuario) != -1) {

      FilaUsuario = ii + 3;
      ii=-1;
    }
  }
  
  if (FilaUsuario == 0){
    lastRow = lastRow + 1;
    var range2 = sheet.getRange("A"+ lastRow +":D"+ lastRow);
    range2.setValues([[usuario, "W", FilaUsuario, fila]]);
  }else{
    var range3 = sheet.getRange("D"+ FilaUsuario);
    range3.setValue(fila)
  }
  
}

function Reporte_(palabra_Buscar){
  buscarPalabra_(palabra_Buscar);
  var tTemporal = tResultado;
  var ultima = tTemporal.length;
  var tMostrar = new Array (1);
  var cuenta = 0;
  var primero = 0;
  var i1 = 0;
  var i2 = 0;
  var valor1 = 0;
  var valor2 = 0;
  for(i1= 0; i1 < ultima ; i1++){
    primero = 1;
    for (i2 = 0; i2 < ultima ; i2++){
      if (tTemporal[i1][5] == tTemporal[i2][5]){
        if(tTemporal[i2][0].toString() != "0"){
        tMostrar[cuenta] = new Array (21);
        if (primero == 1){
          tMostrar[cuenta][0] = "Y";
        }else{
          tMostrar[cuenta][0] = "N";
        }
        tMostrar[cuenta][1] = tTemporal[i2][0];
        tMostrar[cuenta][2] = tTemporal[i2][1];
        tMostrar[cuenta][3] = tTemporal[i2][2];
        tMostrar[cuenta][4] = tTemporal[i2][3];
        tMostrar[cuenta][5] = tTemporal[i2][4];
        tMostrar[cuenta][6] = tTemporal[i2][5];
        tMostrar[cuenta][7] = tTemporal[i2][6];
        tMostrar[cuenta][8] = tTemporal[i2][7];
        tMostrar[cuenta][9] = tTemporal[i2][8];
        tMostrar[cuenta][10] = tTemporal[i2][9];
        tMostrar[cuenta][11] = tTemporal[i2][10];
        tMostrar[cuenta][12] = tTemporal[i2][11];
        tMostrar[cuenta][13] = tTemporal[i2][12];
        tMostrar[cuenta][14] = tTemporal[i2][13];
        tMostrar[cuenta][15] = tTemporal[i2][14];
        tMostrar[cuenta][16] = tTemporal[i2][15];
        tMostrar[cuenta][17] = tTemporal[i2][16];
        tMostrar[cuenta][18] = tTemporal[i2][17];
        tMostrar[cuenta][19] = tTemporal[i2][18];
        tMostrar[cuenta][20] = tTemporal[i2][19];
        tMostrar[cuenta][21] = tTemporal[i2][20];
        tTemporal[i2][0] = "0";
        cuenta ++;
        primero = 0;
      }
      }
    }
   
  }
  /*
  var sheet  = SpreadsheetApp.openById(REPORTES).getSheetByName(SHEET_REPO1);
  var lastRow = sheet.getLastRow() + 1;
  var range = sheet.getRange("A2" + ":V"+lastRow);
  range.clearContent();
  cuenta = cuenta + 1;
  var range2 = sheet.getRange("A2" + ":V" + cuenta);
  
  //Guardar
  range2.setValues(tMostrar); 
  */
  var fecha_value = new Date();
  var Fecha = Utilities.formatDate(fecha_value,Session.getTimeZone(), "dd/MM/yyyy' 'HH:mm:ss");
  
  //Google Docs - Generar Reporte  
  var doc = DocumentApp.openById(DocRecporte);
  var body = doc.getBody();
  
  //Limpiar documento Anterior
  body.clear();
  //Insertar Titulo
  var header = body.appendParagraph("Reporte de Control de Activos");
  header.setHeading(DocumentApp.ParagraphHeading.HEADING1);

  //Insertar Subtitulo
  var section = body.appendParagraph("Resultados de la búsqueda: " + palabra_Buscar);
  section.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  
  // Insertar los datos.
  body.appendParagraph(Fecha).setAlignment(DocumentApp.HorizontalAlignment.RIGHT); 
  var texto_Mostrar_1 = ""
  var Condicion = ""
  
  var tit_Mostrar = {};
  tit_Mostrar[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.LEFT;
  tit_Mostrar[DocumentApp.Attribute.FONT_FAMILY] = DocumentApp.FontFamily.ARIAL;
  tit_Mostrar[DocumentApp.Attribute.FONT_SIZE] = 10;
  tit_Mostrar[DocumentApp.Attribute.BOLD] = true;
  
  var des_Mostrar = {};
  des_Mostrar[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.LEFT;
  des_Mostrar[DocumentApp.Attribute.FONT_FAMILY] = DocumentApp.FontFamily.ARIAL;
  des_Mostrar[DocumentApp.Attribute.FONT_SIZE] = 8;
  des_Mostrar[DocumentApp.Attribute.BOLD] = false;
  
  for(i1= 0; i1 < ultima ; i1++){
   
    Condicion = trim(tMostrar[i1][0].toString())
    texto_Mostrar_1 = "";
    if (Condicion == "Y"){
      body.appendHorizontalRule();
      texto_Mostrar_1 = "Nº ACTIVO: " + tMostrar[i1][6].toString() + "\n" + tMostrar[i1][3].toString() + " " + tMostrar[i1][4].toString() 
        + " CON Nº SERIE: " + tMostrar[i1][5].toString() + " MODELO: " + tMostrar[i1][20].toString() + "\n" 
        + "IDENTIFICADOR: " + tMostrar[i1][7].toString();
      body.appendParagraph(texto_Mostrar_1).setAttributes(tit_Mostrar);
      texto_Mostrar_1 =  "MOVIMIENTO: " + tMostrar[i1][9].toString() + " " + tMostrar[i1][21].toString() + " FECHA: " + tMostrar[i1][1] ;
      if (trim(tResultado[i1][9].toString()) == "EN STOCK") {
        texto_Mostrar_1 = texto_Mostrar_1 + "\n" +"Almacen: " + tResultado[i1][10].toString();
      }
      if (trim(tResultado[i1][9].toString()) == "ENVIADO A SOPORTE") {
        texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "Asignado a: " + tResultado[i1][11].toString();
      }
      if (trim(tResultado[i1][9].toString()) == "ENVIADO A USUARIO FINAL") {
        texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "Usuario Final: " + tResultado[i1][12].toString();
        texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "Unidad de Negocio: " + tResultado[i1][13].toString();
        texto_Mostrar_1 = texto_Mostrar_1 + " " + "Ubicación: " + tResultado[i1][14].toString();
      }
      if (trim(tResultado[i1][9].toString()) == "ENVIADO A PROVEEDOR") {
        texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "Nombre del Proveedor: " + tResultado[i1][15].toString();
        texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "Notas: " + tResultado[i1][16].toString();
      }
      texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "Guia: " + tResultado[i1][17].toString() + " Ext.: " + tResultado[i1][18].toString();
      texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "Observaciones: " + tResultado[i1][19].toString();
      body.appendParagraph(texto_Mostrar_1).setAttributes(des_Mostrar);
    }else{
      texto_Mostrar_1 =  "MOVIMIENTO: " + tMostrar[i1][9].toString() + " " + tMostrar[i1][21].toString() + " FECHA: " + tMostrar[i1][1];
      if (trim(tResultado[i1][9].toString()) == "EN STOCK") {
        texto_Mostrar_1 = texto_Mostrar_1 + "\n" +"Almacen: " + tResultado[i1][10].toString();
      }
      if (trim(tResultado[i1][9].toString()) == "ENVIADO A SOPORTE") {
        texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "Asignado a: " + tResultado[i1][11].toString();
      }
      if (trim(tResultado[i1][9].toString()) == "ENVIADO A USUARIO FINAL") {
        texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "Usuario Final: " + tResultado[i1][12].toString();
        texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "Unidad de Negocio: " + tResultado[i1][13].toString();
        texto_Mostrar_1 = texto_Mostrar_1 + " " + "Ubicación: " + tResultado[i1][14].toString();
      }
      if (trim(tResultado[i1][9].toString()) == "ENVIADO A PROVEEDOR") {
        texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "Nombre del Proveedor: " + tResultado[i1][15].toString();
        texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "Notas: " + tResultado[i1][16].toString();
      }
      texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "Guia: " + tResultado[i1][17].toString() + " Ext.: " + tResultado[i1][18].toString();
      texto_Mostrar_1 = texto_Mostrar_1 + "\n" + "Observaciones: " + tResultado[i1][19].toString();
      body.appendParagraph(texto_Mostrar_1).setAttributes(des_Mostrar);
    }
   
  }
   
}

function ligaArchivo_(archivoPDF)
{
  
  var app = UiApp.getActiveApplication();
  var file = DriveApp.getFileById(archivoPDF);
  var facturapdf = file.getUrl();
  var anchor = app.createAnchor("Documento", facturapdf);
  /*
  var folder =  DriveApp.getFileById(archivoPDF);
  var anchor;
  while (folder.hasNext()) 
  {
    var file = folder.next();
    
    //Obtiene la URL y se la asigna a una variable
    var facturapdf = file.getUrl();
    anchor = app.createAnchor("Abrir Documento", facturapdf);
    
  }*/
  return anchor;
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
