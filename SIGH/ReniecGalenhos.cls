VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReniecGalenhos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para conectarse a REINIEC
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
'
'------------------------------------------------------------------------------------
Option Explicit

' Definicion descrita en el servicio de la RENIEC
Private Const DesfRENIEC = _
    "<?xml version=""1.0"" encoding=""utf-8""?>" & _
    "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
        "<soap:Body>" & _
            "<GetReniec xmlns=""http://tempuri.org/"">" & _
                "<strDNIAuto>string</strDNIAuto>" & _
                "<strDNICon>string</strDNICon>" & _
            "</GetReniec>" & _
        "</soap:Body>" & _
    "</soap:Envelope>"

' Definicion de Ruta de Servicio Web
Dim RutaWeb As String
' Constante que representa el valor del DNI Autorizado
Dim lcDNIautorizado As String

' Aqui se guardara la estructura de datos que se recibe del servicio web
Dim XmlStructura As String
Dim ms_ApellidoMaterno As String
Dim ms_DireccionDomicilio As String
Dim ms_ApellidoPaterno As String
Dim ms_PrimerNombre As String
Dim ms_SegundoNombre As String
Dim ms_TercerNombre As String
Dim mda_FechaNacimiento As Date
Dim ml_idTipoSexo As Long
Dim ml_IdPaisDomicilio As Long
Dim ml_IdDistritoDomicilio As Long
Dim ml_IdCentroPobladoDomicilio As Long
Dim mb_SeAccesaAlaWebDesdeGalenhos As Boolean
Dim mo_ReglasComunes As New SIGHNegocios.ReglasComunes
Dim mb_lbElProcesoEsMasivoNoMuestraErrores As Boolean

Property Let ElProcesoEsMasivoNoMuestraErrores(lValue As Boolean)
   mb_lbElProcesoEsMasivoNoMuestraErrores = lValue
End Property

Property Let SeAccesaAlaWebDesdeGalenhos(lValue As Boolean)
   mb_SeAccesaAlaWebDesdeGalenhos = lValue
End Property


Property Let IdDistritoDomicilio(lValue As Long)
   ml_IdDistritoDomicilio = lValue
End Property
Property Get IdDistritoDomicilio() As Long
   IdDistritoDomicilio = ml_IdDistritoDomicilio
End Property
Property Let IdPaisDomicilio(lValue As Long)
   ml_IdPaisDomicilio = lValue
End Property
Property Get IdPaisDomicilio() As Long
   IdPaisDomicilio = ml_IdPaisDomicilio
End Property
Property Let IdCentroPobladoDomicilio(lValue As Long)
   ml_IdCentroPobladoDomicilio = lValue
End Property
Property Get IdCentroPobladoDomicilio() As Long
   IdCentroPobladoDomicilio = ml_IdCentroPobladoDomicilio
End Property
Property Let idTipoSexo(lValue As Long)
   ml_idTipoSexo = lValue
End Property
Property Get idTipoSexo() As Long
   idTipoSexo = ml_idTipoSexo
End Property
Property Let FechaNacimiento(daValue As Date)
   mda_FechaNacimiento = daValue
End Property
Property Get FechaNacimiento() As Date
   FechaNacimiento = mda_FechaNacimiento
End Property
Property Let TercerNombre(sValue As String)
   ms_TercerNombre = sValue
End Property
Property Get TercerNombre() As String
   TercerNombre = ms_TercerNombre
End Property
Property Let SegundoNombre(sValue As String)
   ms_SegundoNombre = sValue
End Property
Property Get SegundoNombre() As String
   SegundoNombre = ms_SegundoNombre
End Property
Property Let PrimerNombre(sValue As String)
   ms_PrimerNombre = sValue
End Property
Property Get PrimerNombre() As String
   PrimerNombre = ms_PrimerNombre
End Property
Property Let ApellidoPaterno(sValue As String)
   ms_ApellidoPaterno = sValue
End Property
Property Get ApellidoPaterno() As String
   ApellidoPaterno = ms_ApellidoPaterno
End Property
Property Let DireccionDomicilio(sValue As String)
   ms_DireccionDomicilio = sValue
End Property
Property Get DireccionDomicilio() As String
   DireccionDomicilio = ms_DireccionDomicilio
End Property
Property Let ApellidoMaterno(sValue As String)
   ms_ApellidoMaterno = sValue
End Property
Property Get ApellidoMaterno() As String
   ApellidoMaterno = ms_ApellidoMaterno
End Property


Public Sub ConsultarDNIenReniec(ms_NroDocumento As String)
    Dim parser As DOMDocument
    Set parser = New DOMDocument
    ms_ApellidoPaterno = ""
    ' cargar el código SOAP
    parser.LoadXml DesfRENIEC
    '
    ' Indicar el parámetro a enviar
    
    parser.selectSingleNode("/soap:Envelope/soap:Body/GetReniec/strDNIAuto").Text = lcDNIautorizado
    parser.selectSingleNode("/soap:Envelope/soap:Body/GetReniec/strDNICon").Text = ms_NroDocumento
    '
    ' Mostrar el código XML enviado al servicio Web
    ' Text1.Text = parser.xml
    '
    ' Ejecutar la operación HTTP POST
    enviarComando parser.XML, "http://tempuri.org/GetReniec"
    Set parser = Nothing
End Sub

Private Sub enviarComando(ByVal sXml As String, ByVal sSoapAction As String)
    ' Enviar el comando al servicio Web
    '
    ' usar XMLHTTPRequest para enviar la información al servicio Web
    Dim oHttReq As XMLHTTPRequest
    Set oHttReq = New XMLHTTPRequest
    '
    ' Enviar el comando de forma síncrona (se espera a que se reciba la respuesta)
    oHttReq.Open "POST", RutaWeb, False
    
    ' las cabeceras a enviar al servicio Web - (no incluir los dos puntos en el nombre de la cabecera)
    oHttReq.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    oHttReq.setRequestHeader "SOAPAction", sSoapAction
    
    ' enviar el comando
    oHttReq.send sXml
    '
    ' este será el texto recibido del servicio Web
    procesarRespuesta oHttReq.responseText
    Set oHttReq = Nothing
End Sub

Private Sub procesarRespuesta(ByVal s As String)
    ' procesar la respuesta recibida del servicio Web
    XmlStructura = s
    '
    ' Poner los datos en el analizador de XML
    Dim parser As DOMDocument
    Set parser = New DOMDocument
    parser.LoadXml s
    '
    On Error Resume Next
    'DEBB-Desagrega Datos (inicio)
    Dim lcDatosWeb As String, lcTextoBusqueda1 As String, lcTextoBusqueda2 As String
    Dim lnPosActual1 As Integer, lnPosActual2 As Integer, lnPosInicial As Integer
    Dim LcTexto1 As String, lcDniAbuscar As String, lcUbigeoDomicilio As String
    'lcDatosWeb = Me.txtDatosBasicos.Text
    lcDatosWeb = parser.selectSingleNode("/soap:Envelope/soap:Body/GetReniecResponse/GetReniecResult").XML
    If Err.Number > 0 Then
       If mb_SeAccesaAlaWebDesdeGalenhos = True Then
          MsgBox "Puede ser problema de CABLE," & Chr(13) & _
                 "No hay servicio de INTERNET," & Chr(13) & _
                 "hay problemas en la página WEB RENIEC," & Chr(13) & _
                 "el IP o Persona Responsable no está inscrita en MINSA", vbInformation, "PROBLEMAS CON LA WEB RENIEC"
       End If
       Exit Sub
    End If
    '0-Codigo de Error
    lnPosInicial = 5
    lcTextoBusqueda1 = "><string>"
    lnPosActual1 = InStr(lnPosInicial, lcDatosWeb, lcTextoBusqueda1) + Len(lcTextoBusqueda1)
    lcTextoBusqueda2 = "</string><string>"
    lnPosActual2 = InStr(lnPosActual1, lcDatosWeb, lcTextoBusqueda2)
    LcTexto1 = Mid(lcDatosWeb, lnPosActual1, lnPosActual2 - lnPosActual1)
    If Trim(LcTexto1) <> "0000" Then
       If mb_SeAccesaAlaWebDesdeGalenhos = True And mb_lbElProcesoEsMasivoNoMuestraErrores = False Then
          MsgBox "Hubo conexión a la RENIEC, pero no está devolviendo correctamente los datos", vbInformation, "Mensaje"
          Set parser = Nothing
          
       End If
       Exit Sub
    End If
    '1-Apellido paterno
    lcTextoBusqueda1 = "</string><string>"
    lnPosActual1 = lnPosActual2 + Len(lcTextoBusqueda1)
    lcTextoBusqueda2 = "</string><string>"
    lnPosActual2 = InStr(lnPosActual1, lcDatosWeb, lcTextoBusqueda2)
    ms_ApellidoPaterno = Mid(lcDatosWeb, lnPosActual1, lnPosActual2 - lnPosActual1)
    '2-Apellido materno
    lcTextoBusqueda1 = "</string><string>"
    lnPosActual1 = lnPosActual2 + Len(lcTextoBusqueda1)
    lcTextoBusqueda2 = "</string><string>"
    lnPosActual2 = InStr(lnPosActual1, lcDatosWeb, lcTextoBusqueda2)
    ms_ApellidoMaterno = Mid(lcDatosWeb, lnPosActual1, lnPosActual2 - lnPosActual1)
    '3-Primer Nombre/segundo/tercer nombre
    lcTextoBusqueda1 = "</string><string>"
    lnPosActual1 = lnPosActual2 + Len(lcTextoBusqueda1)
    lcTextoBusqueda2 = "</string><string>"
    lnPosActual2 = InStr(lnPosActual1, lcDatosWeb, lcTextoBusqueda2)
    ms_PrimerNombre = Mid(lcDatosWeb, lnPosActual1, lnPosActual2 - lnPosActual1)
    ms_SegundoNombre = ms_PrimerNombre
    ms_TercerNombre = ms_PrimerNombre
    ms_PrimerNombre = Left(RetornaPrimerNombre(ms_PrimerNombre), 20)
    ms_SegundoNombre = Left(RetornaSegundoNombre(ms_SegundoNombre), 20)
    ms_TercerNombre = Left(RetornaTercerNombre(ms_TercerNombre), 20)
    '4-Codigo Ubigeo del Continente
    lcTextoBusqueda1 = "</string><string>"
    lnPosActual1 = lnPosActual2 + Len(lcTextoBusqueda1)
    lcTextoBusqueda2 = "</string><string>"
    lnPosActual2 = InStr(lnPosActual1, lcDatosWeb, lcTextoBusqueda2)
    LcTexto1 = Mid(lcDatosWeb, lnPosActual1, lnPosActual2 - lnPosActual1)
    '5-Codigo Ubigeo del Pais domicilio
    lcTextoBusqueda1 = "</string><string>"
    lnPosActual1 = lnPosActual2 + Len(lcTextoBusqueda1)
    lcTextoBusqueda2 = "</string><string>"
    lnPosActual2 = InStr(lnPosActual1, lcDatosWeb, lcTextoBusqueda2)
    LcTexto1 = Mid(lcDatosWeb, lnPosActual1, lnPosActual2 - lnPosActual1)
    '6-Codigo Ubigeo del Departamento domicilio
    lcTextoBusqueda1 = "</string><string>"
    lnPosActual1 = lnPosActual2 + Len(lcTextoBusqueda1)
    lcTextoBusqueda2 = "</string><string>"
    lnPosActual2 = InStr(lnPosActual1, lcDatosWeb, lcTextoBusqueda2)
    LcTexto1 = Mid(lcDatosWeb, lnPosActual1, lnPosActual2 - lnPosActual1)
    lcUbigeoDomicilio = LcTexto1
    '7-Codigo Ubigeo del Provincia domicilio
    lcTextoBusqueda1 = "</string><string>"
    lnPosActual1 = lnPosActual2 + Len(lcTextoBusqueda1)
    lcTextoBusqueda2 = "</string><string>"
    lnPosActual2 = InStr(lnPosActual1, lcDatosWeb, lcTextoBusqueda2)
    LcTexto1 = Mid(lcDatosWeb, lnPosActual1, lnPosActual2 - lnPosActual1)
    lcUbigeoDomicilio = lcUbigeoDomicilio & LcTexto1
    '8-Codigo Ubigeo del Distrito domicilio
    lcTextoBusqueda1 = "</string><string>"
    lnPosActual1 = lnPosActual2 + Len(lcTextoBusqueda1)
    lcTextoBusqueda2 = "</string><string>"
    lnPosActual2 = InStr(lnPosActual1, lcDatosWeb, lcTextoBusqueda2)
    LcTexto1 = Mid(lcDatosWeb, lnPosActual1, lnPosActual2 - lnPosActual1)
    lcUbigeoDomicilio = lcUbigeoDomicilio & LcTexto1
    ml_IdDistritoDomicilio = 0
    If Val(lcUbigeoDomicilio) > 0 Then
        Dim oRsTmp1 As New Recordset
        Set oRsTmp1 = mo_ReglasComunes.DistritosSeleccionarXfiltro("IdReniec=" & lcUbigeoDomicilio)
        If oRsTmp1.RecordCount > 0 Then
           ml_IdDistritoDomicilio = oRsTmp1.Fields!IdDistrito
        End If
        oRsTmp1.Close
        Set oRsTmp1 = Nothing
    End If
    '9-Codigo Ubigeo del Centro Poblado domicilio
    lcTextoBusqueda1 = "</string><string>"
    lnPosActual1 = lnPosActual2 + Len(lcTextoBusqueda1)
    lcTextoBusqueda2 = "</string><string>"
    lnPosActual2 = InStr(lnPosActual1, lcDatosWeb, lcTextoBusqueda2)
    LcTexto1 = Mid(lcDatosWeb, lnPosActual1, lnPosActual2 - lnPosActual1)
    '10-Codigo del Continente domicilio
    lcTextoBusqueda1 = "</string><string>"
    lnPosActual1 = lnPosActual2 + Len(lcTextoBusqueda1)
    lcTextoBusqueda2 = "</string><string>"
    lnPosActual2 = InStr(lnPosActual1, lcDatosWeb, lcTextoBusqueda2)
    LcTexto1 = Mid(lcDatosWeb, lnPosActual1, lnPosActual2 - lnPosActual1)
    '11-Codigo del Pais domicilio
    lcTextoBusqueda1 = "</string><string>"
    lnPosActual1 = lnPosActual2 + Len(lcTextoBusqueda1)
    lcTextoBusqueda2 = "</string><string>"
    lnPosActual2 = InStr(lnPosActual1, lcDatosWeb, lcTextoBusqueda2)
    LcTexto1 = Mid(lcDatosWeb, lnPosActual1, lnPosActual2 - lnPosActual1)
    '12-Codigo del Departamento domicilio
    lcTextoBusqueda1 = "</string><string>"
    lnPosActual1 = lnPosActual2 + Len(lcTextoBusqueda1)
    lcTextoBusqueda2 = "</string><string>"
    lnPosActual2 = InStr(lnPosActual1, lcDatosWeb, lcTextoBusqueda2)
    LcTexto1 = Mid(lcDatosWeb, lnPosActual1, lnPosActual2 - lnPosActual1)
    '13-Codigo del Provincia domicilio
    lcTextoBusqueda1 = "</string><string>"
    lnPosActual1 = lnPosActual2 + Len(lcTextoBusqueda1)
    lcTextoBusqueda2 = "</string><string>"
    lnPosActual2 = InStr(lnPosActual1, lcDatosWeb, lcTextoBusqueda2)
    LcTexto1 = Mid(lcDatosWeb, lnPosActual1, lnPosActual2 - lnPosActual1)
    '14-Codigo del Distrito domicilio
    lcTextoBusqueda1 = "</string><string>"
    lnPosActual1 = lnPosActual2 + Len(lcTextoBusqueda1)
    lcTextoBusqueda2 = "</string><string>"
    lnPosActual2 = InStr(lnPosActual1, lcDatosWeb, lcTextoBusqueda2)
    LcTexto1 = Mid(lcDatosWeb, lnPosActual1, lnPosActual2 - lnPosActual1)
    '15-Codigo del Centro Poblado domicilio
    lcTextoBusqueda1 = "</string><string>"
    lnPosActual1 = lnPosActual2 + Len(lcTextoBusqueda1)
    lcTextoBusqueda2 = "</string><string>"
    lnPosActual2 = InStr(lnPosActual1, lcDatosWeb, lcTextoBusqueda2)
    LcTexto1 = Mid(lcDatosWeb, lnPosActual1, lnPosActual2 - lnPosActual1)
    '16-Direccion domicilio
    lcTextoBusqueda1 = "</string><string>"
    lnPosActual1 = lnPosActual2 + Len(lcTextoBusqueda1)
    lcTextoBusqueda2 = "</string><string>"
    lnPosActual2 = InStr(lnPosActual1, lcDatosWeb, lcTextoBusqueda2)
    ms_DireccionDomicilio = Mid(lcDatosWeb, lnPosActual1, lnPosActual2 - lnPosActual1)
    '17-Sexo
    lcTextoBusqueda1 = "</string><string>"
    lnPosActual1 = lnPosActual2 + Len(lcTextoBusqueda1)
    lcTextoBusqueda2 = "</string><string>"
    lnPosActual2 = InStr(lnPosActual1, lcDatosWeb, lcTextoBusqueda2)
    LcTexto1 = Mid(lcDatosWeb, lnPosActual1, lnPosActual2 - lnPosActual1)
    ml_idTipoSexo = IIf(Val(LcTexto1) = "1", 1, 2)
    '18-Fecha Nacimiento
    lcTextoBusqueda1 = "</string><string>"
    lnPosActual1 = lnPosActual2 + Len(lcTextoBusqueda1)
    lcTextoBusqueda2 = "</string><string>"
    lnPosActual2 = InStr(lnPosActual1, lcDatosWeb, lcTextoBusqueda2)
    LcTexto1 = Val(Mid(lcDatosWeb, lnPosActual1, lnPosActual2 - lnPosActual1))
    mda_FechaNacimiento = CDate(Right(LcTexto1, 2) & "/" & Mid(LcTexto1, 5, 2) & "/" & Left(LcTexto1, 4))
    '19-Fecha Emision
    lcTextoBusqueda1 = "</string><string>"
    lnPosActual1 = lnPosActual2 + Len(lcTextoBusqueda1)
    lcTextoBusqueda2 = "</string><string>"
    lnPosActual2 = InStr(lnPosActual1, lcDatosWeb, lcTextoBusqueda2)
    LcTexto1 = Mid(lcDatosWeb, lnPosActual1, lnPosActual2 - lnPosActual1)
    '20-Donacion organo
    lcTextoBusqueda1 = "</string><string>"
    lnPosActual1 = lnPosActual2 + Len(lcTextoBusqueda1)
    lcTextoBusqueda2 = "</string><string>"
    lnPosActual2 = InStr(lnPosActual1, lcDatosWeb, lcTextoBusqueda2)
    LcTexto1 = Mid(lcDatosWeb, lnPosActual1, lnPosActual2 - lnPosActual1)
    '21-Nro DNI
    lcTextoBusqueda1 = "</string><string>"
    lnPosActual1 = (lnPosActual2) + Len(lcTextoBusqueda1)
    lcTextoBusqueda2 = "</string><string>"
    lnPosActual2 = InStr(lnPosActual1, lcDatosWeb, lcTextoBusqueda2)
    LcTexto1 = Mid(lcDatosWeb, lnPosActual1, lnPosActual2 - lnPosActual1)
    lcDniAbuscar = LcTexto1

    'DEBB-Desagrega Datos (fin)
    'DEBB-Graba auditoria (inicio)
'<(Inicio)Comentado Por: WABG el: 31/03/2021-01:18:49 p.m. en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
'    If mb_SeAccesaAlaWebDesdeGalenhos = True And mb_lbElProcesoEsMasivoNoMuestraErrores = False Then
'        Dim oRsGrabaAuditoria As New Recordset
'        Dim lcBuscaParametro As New SIGHDatos.Parametros
'        Dim oReglasComunes As New ReglasComunes
'        oReglasComunes.AuditoriaInsertaDesdeReniec lcDniAbuscar, _
'                       lcDniAbuscar & " " & Left(Trim(ms_ApellidoPaterno) & " " & Trim(ms_ApellidoMaterno) & " " & ms_PrimerNombre & " " & ms_SegundoNombre, 90), _
'                       lcBuscaParametro.RetornaFechaHoraServidorSQL
'        Set oRsGrabaAuditoria = Nothing
'        Set lcBuscaParametro = Nothing
'        Set oReglasComunes = Nothing
'    End If
'</(Fin)Comentado por: WABG el: 31/03/2021-01:18:49 p.m. en el Equipo: SISGALENPLUS-PC><CAMBIO-37>
    'DEBB-Graba auditoria (fin)

End Sub

Public Sub Inicializar()
     Dim lcBuscaParametro As New SIGHDatos.Parametros
     Dim oCrypKey As New CrypKey.Util
     lcDNIautorizado = lcBuscaParametro.SeleccionaFilaParametro(297)
     If Len(lcDNIautorizado) > 0 Then
        lcDNIautorizado = Trim(oCrypKey.DecryptString(lcDNIautorizado))
     End If
     RutaWeb = lcBuscaParametro.SeleccionaFilaParametro(298)
     If Len(RutaWeb) > 0 Then
        RutaWeb = oCrypKey.DecryptString(lcBuscaParametro.SeleccionaFilaParametro(298))
     End If
     If Len(lcDNIautorizado) <> 8 Or Len(RutaWeb) = 0 Then
        If mb_SeAccesaAlaWebDesdeGalenhos = True Then
           MsgBox "No podrá accesar a la WEB RENIEC, tiene problemas con parámetros 297, 298", vbInformation, "Mensaje"
        End If
     End If
     Set lcBuscaParametro = Nothing
     Set oCrypKey = Nothing
End Sub


Function RetornaPrimerNombre(lcPrimerSegundoNombreJuntos As String) As String
    Dim ln As Integer
    RetornaPrimerNombre = ""
    ln = InStr(lcPrimerSegundoNombreJuntos, " ")
    If ln > 0 Then
       RetornaPrimerNombre = Trim(Left(lcPrimerSegundoNombreJuntos, ln))
    Else
       RetornaPrimerNombre = lcPrimerSegundoNombreJuntos
    End If
End Function

Function RetornaSegundoNombre(lcPrimerSegundoNombreJuntos As String) As String
    Dim ln As Integer
    RetornaSegundoNombre = ""
    ln = InStr(lcPrimerSegundoNombreJuntos, " ")
    If ln > 0 Then
       RetornaSegundoNombre = Trim(Mid(lcPrimerSegundoNombreJuntos, ln + 1, 100))
       ln = InStr(RetornaSegundoNombre, " ")
       If ln > 0 Then
          RetornaSegundoNombre = Trim(Left(RetornaSegundoNombre, ln))
       End If
    End If
End Function

Function RetornaTercerNombre(lcPrimerSegundoNombreJuntos As String) As String
    Dim ln As Integer, lcNombre1 As String, lcNombre2 As String, lcNombre3 As String
    RetornaTercerNombre = ""
    ln = InStr(lcPrimerSegundoNombreJuntos, " ")
    If ln > 0 Then
       lcNombre1 = Trim(Mid(lcPrimerSegundoNombreJuntos, ln + 1, 100))
       ln = InStr(lcNombre1, " ")
       If ln > 0 Then
          lcNombre2 = Trim(Left(lcNombre1, ln))
          RetornaTercerNombre = Trim(Mid(lcNombre1, ln + 1, 100))
       End If
    End If
End Function





