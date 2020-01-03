VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReglasHoteleria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para CAMAS
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ms_MensajeError As String
Dim mo_AdminHoteleria As New ReglasHoteleria
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Const ELEMENT1_WIDTH = 60
Const ELEMENT1_HEIGHT = 100

Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property

Function CrearArchivoSVGPorServicio(lIdServicio As Long, lTipoServicio As Long, lIdUsuario As Long, sRutaSVG As String) As String
Dim oSVGEstructura As New SVGEstructura
Dim oDOSVGEstructura As New DOSVGEstructura
Dim sHeader As String
Dim sHerramienta1 As String
Dim sHerramienta2 As String
Dim sSimbolos As String
Dim sMenu As String
Dim sMapaDefault As String
Dim sCamasPorServicio As String
Dim sMapaDeServicio As String
Dim oConexion As New Connection
Dim sCodigoServicio  As String
Dim sNombreServicio As String

    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oSVGEstructura.Conexion = oConexion

    oDOSVGEstructura.tipo = "MAPA"
    oDOSVGEstructura.Codigo = "HEADER"
    If oSVGEstructura.SeleccionarPorId(oDOSVGEstructura) Then
        sHeader = oDOSVGEstructura.SVG
    End If

    oDOSVGEstructura.tipo = "MAPA"
    oDOSVGEstructura.Codigo = "HERRAMIENTA1"
    If oSVGEstructura.SeleccionarPorId(oDOSVGEstructura) Then
        sHerramienta1 = oDOSVGEstructura.SVG
    End If

    oDOSVGEstructura.tipo = "MAPA"
    oDOSVGEstructura.Codigo = "HERRAMIENTA2"
    If oSVGEstructura.SeleccionarPorId(oDOSVGEstructura) Then
        sHerramienta2 = oDOSVGEstructura.SVG
    End If

    oDOSVGEstructura.tipo = "MAPA"
    oDOSVGEstructura.Codigo = "SIMBOLOS"
    If oSVGEstructura.SeleccionarPorId(oDOSVGEstructura) Then
        sSimbolos = oDOSVGEstructura.SVG
    End If

    oDOSVGEstructura.tipo = "MAPA"
    oDOSVGEstructura.Codigo = "MENU"
    If oSVGEstructura.SeleccionarPorId(oDOSVGEstructura) Then
        sMenu = oDOSVGEstructura.SVG
    End If

    oDOSVGEstructura.tipo = "MAPA"
    oDOSVGEstructura.Codigo = "MAP_DEFAULT"
    If oSVGEstructura.SeleccionarPorId(oDOSVGEstructura) Then
        sMapaDefault = oDOSVGEstructura.SVG
    End If
    
    Dim oDOServicio As New DOServicio
    Dim oServicio As New Servicios
    Set oServicio.Conexion = oConexion
    oDOServicio.IdServicio = lIdServicio
    If oServicio.SeleccionarPorId(oDOServicio) Then
        sCodigoServicio = oDOServicio.Codigo
        sNombreServicio = oDOServicio.nombre
        sMapaDeServicio = oDOServicio.SVG
    End If
    
    If sMapaDeServicio = "" Then
        sMapaDeServicio = sMapaDefault
    End If
    
    sCamasPorServicio = ""
    Dim rsCamas As New ADODB.Recordset
    Set rsCamas = mo_AdminHoteleria.CamasSeleccionarPorServicioUbicacionActual(lIdServicio)
    Do While Not rsCamas.EOF
        
        sCamasPorServicio = sCamasPorServicio + _
        "<g><use id='" & rsCamas!IdCama & "' x='" & rsCamas!x & "' y='" & rsCamas!Y & _
        "' width='" & ELEMENT1_WIDTH & "' height='" & ELEMENT1_HEIGHT & _
        "' xlink:href='#cama' onmousedown='ObjetoOnMouseDown(evt)' class='estadocama" & rsCamas!IdEstadoCama & "'/>" + Chr(13) + _
        "<text id='t" & rsCamas!IdCama & "' x='" & rsCamas!x & "' y='" & rsCamas!Y & "'><tspan fill='black'>" + _
        "<tspan dx= '0.5em' dy='1em'>Nº : " + rsCamas!Codigo + "</tspan>" + _
        "</tspan></text>" + _
        "</g>"
        
        rsCamas.MoveNext
    Loop
    
    oConexion.Close
    
    Open sRutaSVG For Output As #1
    Print #1, sHeader + Chr(13) + _
            "<script><![CDATA[" + Chr(13) + "var lIdTipoServicio = " & lTipoServicio & Chr(13) + _
            "var lIdUsuario = " & lIdUsuario & Chr(13) + _
            "var lIdServicio = " & lIdServicio & Chr(13) + _
            "var sCodigoServicio = '" & sCodigoServicio & "'" & Chr(13) + _
            "var sNombreServicio = '" & sNombreServicio & "'" & Chr(13) & "]]></script>" + Chr(13) + _
            sHerramienta1; Chr(13) + _
            "<g id='plano'>" + sMapaDeServicio + "</g>" + Chr(13) + _
            "<g id='camas'>" + sCamasPorServicio + Chr(13) + "</g>" + Chr(13) + _
            sHerramienta2 + Chr(13) + _
            sSimbolos + Chr(13) + _
            sMenu + "</svg>"

    Close #1
    
    
End Function
Function EstadosCamaSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New EstadosCama
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set EstadosCamaSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function TiposCamaSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposCama
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set TiposCamaSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function TiposCondicionOcupacionSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New TiposCondicionOcupacion
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set TiposCondicionOcupacionSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

'-------------------------------------------------------------------------------------
'                                       Camas
'-------------------------------------------------------------------------------------

Function CamasAgregar(oDOCama As DOCama, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcCcama As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCama As New Camas
    
    CamasAgregar = False
    
    oConexion.Open sighentidades.CadenaConexion
    Set oCama.Conexion = oConexion
    If oCama.Insertar(oDOCama) Then
        CamasAgregar = True
    Else
        ms_MensajeError = oCama.MensajeError
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOCama.IdUsuarioAuditoria, "A", oDOCama.IdCama, "Camas", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcCcama)       'ListBarItems.idListItem
    oConexion.Close
    Set oConexion = Nothing
    Set oCama = Nothing
End Function

Function CamasModificar(oDOCama As DOCama, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcCcama As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCama As New Camas
    
    CamasModificar = False
    
    oConexion.Open sighentidades.CadenaConexion
    Set oCama.Conexion = oConexion
    If oCama.Modificar(oDOCama) Then
        CamasModificar = True
    Else
        ms_MensajeError = oCama.MensajeError
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOCama.IdUsuarioAuditoria, "M", oDOCama.IdCama, "Camas", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcCcama)       'ListBarItems.idListItem
    oConexion.Close
    Set oConexion = Nothing
    Set oCama = Nothing
    
End Function

Function CamasEliminar(oDOCama As DOCama, lnIdTablaLISTBARITEMS As Long, lcNombrePc As String, lcCcama As String) As Boolean
Dim oConexion As New ADODB.Connection
Dim oCama As New Camas
    
    CamasEliminar = False
    
    oConexion.Open sighentidades.CadenaConexion
    Set oCama.Conexion = oConexion
    If oCama.Eliminar(oDOCama) Then
        CamasEliminar = True
    Else
        ms_MensajeError = oCama.MensajeError
    End If
    '
    Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOCama.IdUsuarioAuditoria, "E", oDOCama.IdCama, "Camas", oConexion, lnIdTablaLISTBARITEMS, lcNombrePc, lcCcama)      'ListBarItems.idListItem
    oConexion.Close
    Set oConexion = Nothing
    Set oCama = Nothing
    
End Function

Function CamasSeleccionarPorId(IdCama As Long, oConexion As Connection) As DOCama
Dim oCama As New Camas
Dim oDOCama As New DOCama
    
    Set CamasSeleccionarPorId = Nothing
    
    Set oCama.Conexion = oConexion
    oDOCama.IdCama = IdCama
    If oCama.SeleccionarPorId(oDOCama) Then
        Set CamasSeleccionarPorId = oDOCama
    End If

    ms_MensajeError = oCama.MensajeError
    Set oCama = Nothing
    Set oDOCama = Nothing

End Function

Function CamasSeleccionarPorServicioUbicacionActual(IdServicio As Long) As Recordset
Dim oConexion As New ADODB.Connection
Dim oCama As New Camas
Dim oDOCama As New DOCama
    
    Set CamasSeleccionarPorServicioUbicacionActual = Nothing
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oCama.Conexion = oConexion
    Set CamasSeleccionarPorServicioUbicacionActual = oCama.SeleccionarPorServicioUbicacionActual(IdServicio)
    
    oConexion.Close

    ms_MensajeError = oCama.MensajeError
    Set oConexion = Nothing
    Set oCama = Nothing
    Set oDOCama = Nothing

End Function


Function CamasSeleccionarDisponibilidadPorServicioUbicacionActual(IdServicio As Long) As Recordset
Dim oConexion As New ADODB.Connection
Dim oCama As New Camas
Dim oDOCama As New DOCama
    
    Set CamasSeleccionarDisponibilidadPorServicioUbicacionActual = Nothing
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oCama.Conexion = oConexion
    Set CamasSeleccionarDisponibilidadPorServicioUbicacionActual = oCama.SeleccionarDisponibilidadPorServicioUbicacionActual(IdServicio)
    
    oConexion.Close

    ms_MensajeError = oCama.MensajeError
    Set oConexion = Nothing
    Set oCama = Nothing
    Set oDOCama = Nothing

End Function

Function CamasSeleccionarPorCodigo(sCodigo As String) As DOCama
Dim oConexion As New ADODB.Connection
Dim oTabla As New Camas
Dim oDOCama As New DOCama
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    oDOCama.Codigo = sCodigo
    If oTabla.SeleccionarPorCodigo(oDOCama) Then
        Set CamasSeleccionarPorCodigo = oDOCama
    Else
        Set CamasSeleccionarPorCodigo = Nothing
    End If
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function CamasObtenerConElMismoCodigo(oDOCama As DOCama) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New Camas
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set CamasObtenerConElMismoCodigo = oTabla.ObtenerConElMismoCodigo(oDOCama)
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function CamasBuscarCodigoDeCama(lcCodigoCama As String, lnIdServicioPropietario As Long) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CamasBuscarCodigoDeCama"
        Set oParameter = .CreateParameter("@lcCodigoCama", adVarChar, adParamInput, 5, lcCodigoCama): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdServicioPropietario", adInteger, adParamInput, 0, lnIdServicioPropietario): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set CamasBuscarCodigoDeCama = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
        
End Function


Function CamasSeleccionarPorIdentificador(lnIdCama As Long, Optional oConexion1 As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    If oConexion1 Is Nothing Then
        oConexion.CursorLocation = adUseClient
        oConexion.CommandTimeout = 300
        oConexion.Open sighentidades.CadenaConexion
    End If
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexion1 Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexion1
        End If
        .CommandTimeout = 150
        .CommandText = "CamasSeleccionarPorIdentificador"
        Set oParameter = .CreateParameter("@lnIdCama", adInteger, adParamInput, 0, lnIdCama): .Parameters.Append oParameter
        Set oRecordset = .Execute
        If oConexion1 Is Nothing Then
           Set oRecordset.ActiveConnection = Nothing
        End If
   End With
   Set CamasSeleccionarPorIdentificador = oRecordset
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function CamasProgramadasPorAnio(lnAnio As Integer, lbSoloHospitalizacion As Boolean) As ADODB.Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CamasProgramadasPorAnio"
        Set oParameter = .CreateParameter("@lnAnio", adInteger, adParamInput, 0, lnAnio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lbSoloHospitalizacion", adBoolean, adParamInput, 0, lbSoloHospitalizacion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set CamasProgramadasPorAnio = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function




Function CamasSeleccionarXdptoServicioEspecialidad(lnMes As Integer, ml_lnAnio As Integer, ml_idDepartamento1 As Long, _
             ml_idEspecialidad1 As Long, ml_idServicio1 As Long, ml_idServicio2 As Long, ml_TipoReporte As Integer) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CamasSeleccionarXdptoServicioEspecialidad"
        Set oParameter = .CreateParameter("@ml_TipoReporte", adInteger, adParamInput, 0, ml_TipoReporte): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lnMes", adInteger, adParamInput, 0, lnMes): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@ml_lnAnio", adInteger, adParamInput, 0, ml_lnAnio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@ml_idDepartamento1", adInteger, adParamInput, 0, ml_idDepartamento1): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@ml_idEspecialidad1", adInteger, adParamInput, 0, ml_idEspecialidad1): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@ml_idServicio1", adInteger, adParamInput, 0, ml_idServicio1): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@ml_idServicio2", adInteger, adParamInput, 0, ml_idServicio2): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set CamasSeleccionarXdptoServicioEspecialidad = oRecordset
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


Function CargaCamaDisponible(ml_IdServicioConCamaDisponible As Long) As Long

On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CargaCamaDisponible"
        Set oParameter = .CreateParameter("@ml_IdServicioConCamaDisponible", adInteger, adParamInput, 0, ml_IdServicioConCamaDisponible): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   CargaCamaDisponible = 0
   If oRecordset.RecordCount > 0 Then
       CargaCamaDisponible = oRecordset.Fields!IdCama
   End If
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function cargaCamasDisponiblesXtipoServicio(ml_TipoFiltro As sghTipoFiltroAdmision, lnIdServicioActual As Long) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim lcFiltro As String
Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "cargaCamasDisponiblesXtipoServicio"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   lcFiltro = ""
   If ml_TipoFiltro = sghFiltrarHospitalizacion Then
      lcFiltro = "IdTipoServicio=3"
   Else
      lcFiltro = "IdTipoServicio<>3"
   End If
   If lnIdServicioActual > 0 Then
      lcFiltro = lcFiltro & " and IdServicioUbicacionActual=" & lnIdServicioActual
   End If
   oRecordset.Filter = lcFiltro
   Set cargaCamasDisponiblesXtipoServicio = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


'debb-14/07/2016
Function CamasSeleccionarPorIdServicio(lnIdServicio As Long, oConexion As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CamasSeleccionarPorIdServicio"
        Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, lnIdServicio): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set CamasSeleccionarPorIdServicio = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
        
End Function


Function CamasSeleccionarPorIdPaciente(lnIdPaciente As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError  As String
Dim lcFiltro As String
Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "CamasSeleccionarPorIdPaciente"
        Set oParameter = .CreateParameter("@IdPaciente", adInteger, adParamInput, 0, lnIdPaciente): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set CamasSeleccionarPorIdPaciente = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

