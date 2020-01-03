var svgMyDocument;
var svgMyRoot;

var ID_MENU_DEFAULT = "defaultMenu";
var ELEMENT1_WIDTH = 60 
var ELEMENT1_HEIGHT = 100

  function OnLoad(evt) {

		svgMyDocument = evt.getTarget().getOwnerDocument();
		svgMyRoot = svgMyDocument.getDocumentElement();

		configurarMenu(ID_MENU_DEFAULT);

  }

function OnClickAgregarCama() {

				 //Agrega la parte alfa numerica
				 try {
   					 var oCamasProxy =new ActiveXObject('SIGHProxies.CamaDetalleProxy');
						 } catch(e)
						{	
							return;
						}							
						oCamasProxy.Opcion = 1;
						oCamasProxy.IdTipoServicio = lIdTipoServicio;
						oCamasProxy.IdServicio = lIdServicio;
						oCamasProxy.CodigoServicio = sCodigoServicio;
						oCamasProxy.NombreServicio = sNombreServicio;
						oCamasProxy.MostrarDialogo();
						
						if (oCamasProxy.IdCama != 0) {	
							 //actualiza la parte grafica
								try {
    								var oCamas =new ActiveXObject('SIGHNegocios.ReglasHoteleria');
    								var oDOCama =new  ActiveXObject('SIGHComun.DOCama');
								} catch(e)  {
								return;
								}						
								oDOCama = oCamas.CamasSeleccionarPorId(oCamasProxy.IdCama);

							 //Agrega la parte grafica
							 DeseleccionarElemento()
							 var k = svgMyRoot.currentScale;								 
							 var xt = svgMyRoot.currentTranslate.x;
							 var yt = svgMyRoot.currentTranslate.y;
							 var Xini = ((-xt + XLast)/k - ELEMENT1_WIDTH/2);
							 var Yini = ((-yt + YLast)/k - ELEMENT1_HEIGHT/2);
							 var sNodo = "<g><use id='" + oCamasProxy.IdCama  + "' x='" + Xini  + "' y='" + Yini + "' width='"+ ELEMENT1_WIDTH + "' height='" + ELEMENT1_HEIGHT + "' xlink:href='#cama' class='estadocama1' onmousedown='ObjetoOnMouseDown(evt)'/>" 				
									 sNodo = sNodo + "<text id='t" + oCamasProxy.IdCama + "' x='" + Xini + "' y='" + Yini + "'><tspan fill='black'>" 
							   	sNodo = sNodo + "<tspan dx= '0.5em' dy='1em'>Nº : " + oDOCama.Codigo + "</tspan>"  
								   sNodo = sNodo + "</tspan></text></g>"
							

							 var oElemento = parseXML(sNodo, svgMyDocument);
							 var oElemento1 = svgMyDocument.getElementById("camas");
							 oElemento1.appendChild(oElemento)

								oDOCama.X = Xini;
								oDOCama.Y = Yini;
								
								oCamas.CamasModificar(oDOCama);
								oCamas = null
							}
							oCamasProxy = null;
}

function OnClickModificarCama()
{
        				try {
           					var oCamasProxy =new ActiveXObject('SIGHProxies.CamaDetalleProxy');
        				} catch(e)  {
    					return;
    				}
						
				try {
    				oCamasProxy.Opcion = 2;
						oCamasProxy.IdCama = oSelectedElement.getAttribute("id");
						oCamasProxy.IdTipoServicio = lIdTipoServicio;
    				oCamasProxy.MostrarDialogo();
    				oCamasProxy = null;
				}
				catch(e)
				{}
}

function OnClickEliminarCama()
{
        				try {
           					var oCamasProxy =new ActiveXObject('SIGHProxies.CamaDetalleProxy');
        				} catch(e)  {
    					return;
    				}		
				try
				{					
    				oCamasProxy.Opcion = 4;
						oCamasProxy.IdCama = oSelectedElement.getAttribute("id");
						oCamasProxy.IdTipoServicio = lIdTipoServicio;
    				oCamasProxy.MostrarDialogo();

						if (oCamasProxy.ConfirmoOperacion == 1) {
							 var oElemento1 = svgMyDocument.getElementById("camas");
							 oElemento1.removeChild(oSelectedElement);
							 }
   					 oCamasProxy = null;
				}catch(e){}
}

function configurarMenu(id)
{

 	 var newMenuRoot = parseXML(printNode(svgMyDocument.getElementById(id)), contextMenu);
 	 contextMenu.replaceChild( newMenuRoot, contextMenu.firstChild );

}

var oSelectedElement = null
var oSelectedTextElement = null
var Xini = 0;
var Yini = 0;
var Transform = "";
function ObjetoOnMouseDown(evt)
{
					DeseleccionarElemento()

					var oTracker = svgDocument.getElementById("MouseTrackerForeground")
					oTracker.setAttribute("visibility", "visible")
					oSelectedElement = evt.getTarget();
					oSelectedTextElement = svgDocument.getElementById("t"+oSelectedElement.getAttribute("id"));

					oSelectedElement.setAttribute("stroke","red");
					oSelectedElement.setAttribute("stroke-width",3);

					Transform = oSelectedElement.getAttribute("transform");
					oSelectedElement.setAttribute("transform", "")

}
function MouseTrackerOnMouseMove(evt)
{

				var width = parseFloat(oSelectedElement.getAttribute("width"));
				var height = parseFloat(oSelectedElement.getAttribute("height"));
				var k = svgMyRoot.currentScale;
				var xt = svgMyRoot.currentTranslate.x;
				var yt = svgMyRoot.currentTranslate.y;

				oSelectedElement.setAttribute("x",  (-xt + evt.getClientX())/k- width/2)																																																														
				oSelectedElement.setAttribute("y",  (-yt + evt.getClientY())/k - height/2)
																																																																																																																																																																	
				oSelectedTextElement.setAttribute("x",  (-xt + evt.getClientX())/k- width/2)																																																														
				oSelectedTextElement.setAttribute("y",  (-yt + evt.getClientY())/k - height/2)																																																																																																																																																																	

}
function MouseTrackerOnMouseUp(evt)
{
				oTracker = svgDocument.getElementById("MouseTrackerForeground")
				oTracker.setAttribute("visibility", "hidden")

				if (Transform!="") {																									
					 temp = Transform.split(",");
					 var Angle = temp[0].split("(");

						var x = parseFloat(oSelectedElement.getAttribute("x"));
						var y = parseFloat(oSelectedElement.getAttribute("y"));
						var width = parseFloat(oSelectedElement.getAttribute("width"));
						var height = parseFloat(oSelectedElement.getAttribute("height"));
						cx = x + width/2
						cy = y + height/2
						oSelectedElement.setAttribute("transform", "rotate(" +Angle[1]+ "," + cx + "," + cy + ")")
				}
				
				if (oSelectedElement.getAttribute("id") != 0){
							 //actualiza la parte grafica
								try {
    								var oCamas =new ActiveXObject('SIGHNegocios.ReglasHoteleria');
    								var oDOCama =new  ActiveXObject('SIGHComun.DOCama');
								} catch(e)  {
								return;
								}						
								
  								oDOCama = oCamas.CamasSeleccionarPorId(oSelectedElement.getAttribute("id"));
									if (oDOCama != null) { 
											oDOCama.X = oSelectedElement.getAttribute("x");
											oDOCama.Y = oSelectedElement.getAttribute("y");
  										oCamas.CamasModificar(oDOCama);
								 }

  								oCamas = null;
  								oDOCama= null;
				}
}

var XLast = 0;
var YLast = 0;		 
function MouseTrackerBackOnMouseMove(evt)
{				

	XLast = evt.getClientX();
	YLast = evt.getClientY();
}

function MouseTrackerBackOnMouseDown(evt)
{
					DeseleccionarElemento()
}

function DeseleccionarElemento(){
		if (oSelectedElement!=null) {
						oSelectedElement.setAttribute("stroke-width",0);
						oSelectedElement = null;
								    }
}
