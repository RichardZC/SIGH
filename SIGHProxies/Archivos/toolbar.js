function OnLoad(evt)
{
}

function BotonesOnMouseClick(evt)
{
		switch (evt.getTarget().getAttribute("id"))
		{
			case "btnAcercar":
					break;
			case "btnOriginal":
					break;
			case "btnAlejar":
					break;
			case "btnRotarDer":
					RotarDerecha()
					break;
			case "btnRotarIzq":
					RotarIzquierda()
					break;
			case "btnLeyenda":
					break;
			case "btnAyuda":
					MostrarAyuda();
					break;
			case "btnsalir":
					GuardarDatosGraficos();
		}
}

var Angle = 0;
function RotarDerecha()
{
	if (oSelectedElement == null) return;

      Angle = Angle - 10
      
      var x = parseFloat(oSelectedElement.getAttribute("x"));
      var y = parseFloat(oSelectedElement.getAttribute("y"));
      var width = parseFloat(oSelectedElement.getAttribute("width"));
      var height = parseFloat(oSelectedElement.getAttribute("height"));
      cx = x + width/2
      cy = y + height/2
      
      oSelectedElement.setAttribute("transform", "rotate(" +Angle+ "," + cx + "," + cy + ")")

}
function RotarIzquierda()
{
	if (oSelectedElement == null) return;		

      Angle = Angle + 10
       
      var x = parseFloat(oSelectedElement.getAttribute("x"));
      var y = parseFloat(oSelectedElement.getAttribute("y"));
      var width = parseFloat(oSelectedElement.getAttribute("width"));
      var height = parseFloat(oSelectedElement.getAttribute("height"));
      cx = x + width/2
      cy = y + height/2
       
      oSelectedElement.setAttribute("transform", "rotate(" +Angle+ "," + cx + "," + cy + ")")
}

function GuardarDatosGraficos() {

    		try {
       	 	var oCamas =new ActiveXObject('SIGHGraphics.Camas');
    		} catch(e)  {
			return;
		}
	
		var oSVGMapa = self.document.svgMapa.getSVGDocument()	
		var oNodos = oSVGMapa.getElementById("camas").getChildNodes();
		for (i=0; i<oNodos.length; i++) {
				oCamas.IdCama = oNodo.item(i).getAttribute("id")
				oCamas.SVG = printNode(oNodo)
				oCamas.ModificarSVG(); 
		}
}




