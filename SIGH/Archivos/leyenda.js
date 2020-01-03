var svgMyDocument;
var svgMyRoot;

function OnLoad(evt)
{
 				 svgMyDocument = evt.getTarget().getOwnerDocument();
    		 svgMyRoot = svgMyDocument.getDocumentElement();

    		 try {
       	 		 			var oLeyenda =new ActiveXObject('SIGHGraphics.Leyenda');
    		} catch(e)  {
										return
    								}
    oSVGLeyenda = parseXML(oLeyenda.Servicios, svgDocument)
    svgMyRoot.appendChild(oSVGLeyenda); 				 
}
