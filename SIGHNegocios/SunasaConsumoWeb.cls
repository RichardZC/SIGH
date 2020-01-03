VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SunasaConsumoWeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para  WEB CON EL SIS
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit

Dim ms_MensajeError As String
Dim lcBuscaParametro As New SIGHDatos.Parametros


Private Const Sunasa_ListaEstablecimientoSalud = _
        "<?xml version=""1.0"" encoding=""utf-8""?>" & _
        "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
          "<soap:Body>" & _
            "<GetListaEstablecimientoSalud xmlns=""http://tempuri.org/"">" & _
              "<Ubigeo>string</Ubigeo>" & _
              "<CodDisaRedMicroRed>string</CodDisaRedMicroRed>" & _
              "<CodCentroPoblado>string</CodCentroPoblado>" & _
              "<DescripcionEntidadSectorial>string</DescripcionEntidadSectorial>" & _
              "<CodigoSectorial>string</CodigoSectorial>" & _
              "<EstabFechaReg>string</EstabFechaReg>" & _
              "<EstabFechaMod>string</EstabFechaMod>" & _
            "</GetListaEstablecimientoSalud>" & _
          "</soap:Body>" & _
        "</soap:Envelope>"

Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property

' Aqui se guardara la estructura de datos que se recibe del servicio web
'Dim XmlStructura As String

'Dim oConexion As New Connection

'Private Function peticionServicioWeb(codigoRenaes As String, _
'                    Ubigeo As String, CodDisaRedMicroRed As String, _
'                    CodCentroPoblado As String, DescripcionEntidadSectorial As String, _
'                    CodigoSectorial As String, EstabFechaReg As String, _
'                    EstabFechaMod As String)
'
'    Dim oHttReq As XMLHTTPRequest
'    Dim parser As DOMDocument
'
'    Dim oM_eess As New m_eess
'    Dim oDom_eess As New Dom_eess
'
'    Set parser = New DOMDocument
'    parser.loadXML Sunasa_ListaEstablecimientoSalud
'
'    parser.selectSingleNode("/soap:Envelope/soap:Body/GetListaEstablecimientoSalud/Ubigeo").Text = ""
'    parser.selectSingleNode("/soap:Envelope/soap:Body/GetListaEstablecimientoSalud/CodDisaRedMicroRed").Text = ""
'    parser.selectSingleNode("/soap:Envelope/soap:Body/GetListaEstablecimientoSalud/CodCentroPoblado").Text = ""
'    parser.selectSingleNode("/soap:Envelope/soap:Body/GetListaEstablecimientoSalud/DescripcionEntidadSectorial").Text = ""
'    parser.selectSingleNode("/soap:Envelope/soap:Body/GetListaEstablecimientoSalud/CodigoSectorial").Text = ""
'    parser.selectSingleNode("/soap:Envelope/soap:Body/GetListaEstablecimientoSalud/EstabFechaReg").Text = ""
'    parser.selectSingleNode("/soap:Envelope/soap:Body/GetListaEstablecimientoSalud/EstabFechaMod").Text = ""
'
'    Set oHttReq = New XMLHTTPRequest
'    oHttReq.Open "POST", "http://app12.sunasa.gob.pe:8085/ServiceSalud.asmx", False
'    oHttReq.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
'    oHttReq.setRequestHeader "Content-Length", "length"
'    oHttReq.setRequestHeader "SOAPAction", "http://tempuri.org/GetListaEstablecimientoSalud"
'
'    On Error Resume Next
'    oHttReq.send parser.XML
'    If Err.Number > 0 Then
'         MsgBox Err.Description
'         Set parser = Nothing
'         Set oHttReq = Nothing
'         Exit Function
'    End If
'
'    Set parser = New DOMDocument
'    parser.loadXML oHttReq.responseText
'
'    Dim node As IXMLDOMNode
'    Dim node1 As IXMLDOMNode
'    Dim atributo As IXMLDOMAttribute
'
'    EstableceConexion
'
'    Dim sql As String, sqlIni As String
'    Dim contador As Integer
'
'    For Each node In parser.selectNodes("//EstablecimientoSalud")
'
'    Next node
'End Function


''Dim canal As Integer
''    canal = FreeFile
''    'Open App.Path & "/testRenaes.txt" For Output As canal
''For Each node In parser.selectNodes("//EstablecimientoSalud")
''
''        Write #canal, node.childNodes(1).nodeTypedValue & "," & node.childNodes(2).nodeTypedValue & "," & node.childNodes(5).nodeTypedValue & "," & node.childNodes(7).nodeTypedValue & "," & IIf(validarCodigoRENAES(node.childNodes(1).nodeTypedValue), 1, 0)
''
''
''        contador = 0
''        Err = 0
''        For Each node1 In node.childNodes
''            ''Debug.Print "Tipo Data : " & node1.DataType & " Nombre de nodo " & node1.nodeName & " valor : " & node1.nodeValue & " Otro:" & node1.nodeTypedValue
''
''            If contador = 0 Then
''                sqlIni = "INSERT INTO m_eesssunasa(" & node1.nodeName
''                sql = "'" & node1.nodeTypedValue & "'"
''            Else
''                sqlIni = sqlIni & "," & node1.nodeName
''                sql = sql & ",'" & Replace(Replace(node1.nodeTypedValue, vbCrLf, " "), "'", "''") & "'"
''            End If
''            contador = contador + 1
''        Next node1
''        sql = sql & ")"
''        sqlIni = sqlIni & ") VALUES ("
''        oConexion.Execute sqlIni & sql
''        If Err Then
''            contador = 0
''            Err = 0
''        End If
''    Next node
''    'Close canals

Function ConsultarServicioBuscarEESSxCodigo(ms_Codigo As String, ByRef rsTmp As Recordset) As Dom_eess
    Dim ODom_eess As Dom_eess
    Set ODom_eess = Nothing
    
    If HabilitadoParaBusquedaEnWebRenaes() = False Then
    
        Dim oParametro As New DOPArametro
        Dim oConexion As New Connection
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        oConexion.Open sighentidades.CadenaConexion
        Set lcBuscaParametro.Conexion = oConexion
        
        oParametro.IdParametro = GetIdParametroBusquedaEnWebRenaes()
        Call lcBuscaParametro.SeleccionarPorId(oParametro)
        oConexion.Close
        Set oConexion = Nothing
        MsgBox "No esta habilitado para la busqueda por web, Parametro :" & oParametro.Descripcion, vbInformation, "Parametros"
        Set ConsultarServicioBuscarEESSxCodigo = ODom_eess
        Exit Function
        
    End If
    
    Dim oHttReq As XMLHTTPRequest
    Dim parser As DOMDocument
    
    'Dim oM_eess As New m_eess
    
    Dim OReglasSis As New SisConsumoWeb
'    Dim lbSalida As Boolean

    
    Set parser = New DOMDocument
    parser.loadXML Sunasa_ListaEstablecimientoSalud

    parser.selectSingleNode("/soap:Envelope/soap:Body/GetListaEstablecimientoSalud/Ubigeo").Text = ""
    parser.selectSingleNode("/soap:Envelope/soap:Body/GetListaEstablecimientoSalud/CodDisaRedMicroRed").Text = ""
    parser.selectSingleNode("/soap:Envelope/soap:Body/GetListaEstablecimientoSalud/CodCentroPoblado").Text = ""
    parser.selectSingleNode("/soap:Envelope/soap:Body/GetListaEstablecimientoSalud/DescripcionEntidadSectorial").Text = ""
    parser.selectSingleNode("/soap:Envelope/soap:Body/GetListaEstablecimientoSalud/CodigoSectorial").Text = sighentidades.FormatoCodigoRENAES(ms_Codigo, SUNASA)
    parser.selectSingleNode("/soap:Envelope/soap:Body/GetListaEstablecimientoSalud/EstabFechaReg").Text = ""
    parser.selectSingleNode("/soap:Envelope/soap:Body/GetListaEstablecimientoSalud/EstabFechaMod").Text = ""
    
    Set oHttReq = New XMLHTTPRequest
    oHttReq.Open "POST", "http://app12.sunasa.gob.pe:8085/ServiceSalud.asmx", False
    oHttReq.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    oHttReq.setRequestHeader "Content-Length", "length"
    oHttReq.setRequestHeader "SOAPAction", "http://tempuri.org/GetListaEstablecimientoSalud"
    
    On Error Resume Next
    
    oHttReq.send parser.XML
    Dim oConsumoWebUtil As New ConsumoWebUtil
    
    ms_MensajeError = ""
    If oConsumoWebUtil.PeticionHttpEsCorrecta(oHttReq) = False Then
        ms_MensajeError = oConsumoWebUtil.ParseMensageError(oHttReq)
        Set parser = Nothing
        Set oHttReq = Nothing
        Exit Function
    End If
    
    Set parser = New DOMDocument
    parser.loadXML oHttReq.responseText
    
    Dim node As IXMLDOMNode
    
    Dim sql As String, sqlIni As String
    Dim contador As Integer
    Dim sCodigoRenaes As String
    
    For Each node In parser.selectNodes("//EstablecimientoSalud")
        Set ODom_eess = New Dom_eess
        
        sCodigoRenaes = FormatoCodigoRENAES(node.selectSingleNode("CodigoEstSalud").nodeTypedValue, SIS)
        
        ODom_eess.pre_Afilia = "S"
        ODom_eess.pre_Aisped = "N"
        ODom_eess.pre_CodEjeAdm = ""
        ODom_eess.pre_CodigoRENAES = sCodigoRenaes
        ODom_eess.pre_esmn = "N"
        
        ODom_eess.pre_idCategoriaEESS = getIdCategoriaEESS(node.selectSingleNode("CtgCodigo").nodeTypedValue) ' node.selectSingleNode("CtgCodigo").nodeTypedValue
        'puede estar nulo
        ODom_eess.pre_IdDisa = Right("000" & Val(node.selectSingleNode("EstabCodDisa").nodeTypedValue) * 10, 3)
        ODom_eess.pre_IdOdsis = Format("000", node.selectSingleNode("EstabCodODSIS").nodeTypedValue)
                
        ODom_eess.pre_IdEESS = sCodigoRenaes
        ODom_eess.pre_IdEstado = IIf(UCase(node.selectSingleNode("EstabEstado").nodeTypedValue) = UCase("ACTIVO"), "1", "0")
        
        ODom_eess.pre_idUbigeo = node.selectSingleNode("EstUbigeo").nodeTypedValue
        ODom_eess.pre_nombre = node.selectSingleNode("EstabNombre").nodeTypedValue
        ODom_eess.Entidad = node.selectSingleNode("Institucion").nodeTypedValue
        ODom_eess.pre_UCI = "N"
        ODom_eess.pre_Umbral = "N"
        ODom_eess.pre_Vrae = "N"
        Exit For
    Next node
    
    Set oHttReq = Nothing
    Set parser = Nothing
    Set oConexion = Nothing
    
    Set ConsultarServicioBuscarEESSxCodigo = ODom_eess
End Function

Function EstableceConexion() As ADODB.Connection
    Dim lcBuscaParametro As New SIGHDatos.Parametros
    Dim oConexion As New ADODB.Connection
    
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghSis)
    
    Set EstableceConexion = oConexion
End Function

Public Function EsEstablecimientoMinsa(ODom_eess As Dom_eess) As Boolean
    Dim bEsMinsa As Boolean
    Dim sEntidad As String
    
    bEsMinsa = False
    
    sEntidad = UCase(Replace(ODom_eess.Entidad, " ", ""))
    Select Case sEntidad
        Case "MINSA":
            bEsMinsa = True
        Case "GOBIERNOREGIONAL":
            bEsMinsa = True
    End Select
    EsEstablecimientoMinsa = bEsMinsa
End Function


Public Function getIdCategoriaEESS(sAbreviaturaCateg As String) As String
On Error GoTo ManejadorDeError
    Dim oDoa_categoriaeess As New DOa_categoriaeess
    Dim oa_categoriaeess As New a_categoriaeess
    Dim oConexion As New Connection
    oConexion.CommandTimeout = 300
    oConexion.CursorLocation = adUseClient
    oConexion.Open lcBuscaParametro.SeleccionaFilaParametro(sghBaseDatosExterna.sghSis)
    
    getIdCategoriaEESS = oa_categoriaeess.ObtenerValorDefectoIdCategoriaEESS
    
    
    Set oa_categoriaeess.Conexion = oConexion
    oDoa_categoriaeess.cat_Abreviatura = sAbreviaturaCateg
    If oa_categoriaeess.SeleccionarPorCatAbreviatura(oDoa_categoriaeess) Then
        getIdCategoriaEESS = oDoa_categoriaeess.cat_IdCategoriaEESS
    End If
    Set oConexion = Nothing
    Exit Function
ManejadorDeError:
    
    MsgBox Err.Description
End Function

Public Function HabilitadoParaBusquedaEnWebRenaes() As Boolean
    Dim bEstaHabilitado As Boolean
    bEstaHabilitado = False
    If lcBuscaParametro.SeleccionaFilaParametro(GetIdParametroBusquedaEnWebRenaes()) = "S" Then
        bEstaHabilitado = True
    End If
    HabilitadoParaBusquedaEnWebRenaes = bEstaHabilitado
End Function

Public Function GetIdParametroBusquedaEnWebRenaes() As Long
    GetIdParametroBusquedaEnWebRenaes = 322
End Function

