VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HIS_Cabecera"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla HIS_Cabecera
'        Programado por: Cachay F
'        Fecha: Noviembre 2013
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Conexion As ADODB.Connection
Dim ms_MensajeError As String

Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property
Property Set Conexion(oValue As ADODB.Connection)
   Set mo_Conexion = oValue
End Property
Property Get Conexion() As ADODB.Connection
   Set Conexion = mo_Conexion
End Property

'------------------------------------------------------------------------------------
'   Función:        Insertar
'   Descripción:    Inserta un registro a la tabla HIS_Cabecera
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOHIS_Cabecera) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "HIS_CabeceraAgregar"
   Set oParameter = .CreateParameter("@IdHisCabecera", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdHisLote", adInteger, adParamInput, 0, IIf(oTabla.IdHisLote = 0, Null, oTabla.IdHisLote))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroHojaHis", adInteger, adParamInput, 0, IIf(oTabla.NroHojaHis = 0, Null, oTabla.NroHojaHis))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroFormato", adInteger, adParamInput, 0, oTabla.NroFormato)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdTurno", adInteger, adParamInput, 0, IIf(oTabla.IdTurno = 0, Null, oTabla.IdTurno))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuario", adInteger, adParamInput, 0, IIf(oTabla.IdUsuario = 0, Null, oTabla.IdUsuario))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEstadoHis", adInteger, adParamInput, 0, IIf(oTabla.IdEstadoHis = 0, Null, oTabla.IdEstadoHis))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, IIf(oTabla.IdMedico = 0, Null, oTabla.IdMedico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEstablecimiento", adInteger, adParamInput, 0, IIf(oTabla.IdEstablecimiento = 0, Null, oTabla.IdEstablecimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, IIf(oTabla.IdServicio = 0, Null, oTabla.IdServicio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaCreacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaCreacion = 0, Null, oTabla.FechaCreacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.IdHisCabecera = .Parameters("@IdHisCabecera")
End With
 
Insertar = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        Modificar
'   Descripción:    modifica un registro a la tabla HIS_Cabecera
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOHIS_Cabecera) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "HIS_CabeceraModificar"
   Set oParameter = .CreateParameter("@IdHisCabecera", adInteger, adParamInput, 0, IIf(oTabla.IdHisCabecera = 0, Null, oTabla.IdHisCabecera))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdHisLote", adInteger, adParamInput, 0, IIf(oTabla.IdHisLote = 0, Null, oTabla.IdHisLote))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroHojaHis", adInteger, adParamInput, 0, IIf(oTabla.NroHojaHis = 0, Null, oTabla.NroHojaHis))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroFormato", adInteger, adParamInput, 0, oTabla.NroFormato)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdTurno", adInteger, adParamInput, 0, IIf(oTabla.IdTurno = 0, Null, oTabla.IdTurno))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuario", adInteger, adParamInput, 0, IIf(oTabla.IdUsuario = 0, Null, oTabla.IdUsuario))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEstadoHis", adInteger, adParamInput, 0, IIf(oTabla.IdEstadoHis = 0, Null, oTabla.IdEstadoHis))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdMedico", adInteger, adParamInput, 0, IIf(oTabla.IdMedico = 0, Null, oTabla.IdMedico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEstablecimiento", adInteger, adParamInput, 0, IIf(oTabla.IdEstablecimiento = 0, Null, oTabla.IdEstablecimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdServicio", adInteger, adParamInput, 0, IIf(oTabla.IdServicio = 0, Null, oTabla.IdServicio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaCreacion", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaCreacion = 0, Null, oTabla.FechaCreacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
End With
 
Modificar = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        Eliminar
'   Descripción:    elimia o borra un registro a la tabla HIS_Cabecera
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOHIS_Cabecera) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "HIS_CabeceraEliminar"
   Set oParameter = .CreateParameter("@IdHisCabecera", adInteger, adParamInput, 0, oTabla.IdHisCabecera)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
End With
 
Eliminar = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        SeleccionarPorId
'   Descripción:    Selecciona un unico registro de la tabla HIS_Cabecera
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOHIS_Cabecera) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "HIS_CabeceraSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdHisCabecera", adInteger, adParamInput, 0, oTabla.IdHisCabecera)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdHisCabecera = IIf(IsNull(oRecordset!IdHisCabecera), 0, oRecordset!IdHisCabecera)
       oTabla.IdHisLote = IIf(IsNull(oRecordset!IdHisLote), 0, oRecordset!IdHisLote)
       oTabla.NroHojaHis = IIf(IsNull(oRecordset!NroHojaHis), 0, oRecordset!NroHojaHis)
       oTabla.NroFormato = IIf(IsNull(oRecordset!NroFormato), 0, oRecordset!NroFormato)
       oTabla.IdTurno = IIf(IsNull(oRecordset!IdTurno), 0, oRecordset!IdTurno)
       oTabla.IdUsuario = IIf(IsNull(oRecordset!IdUsuario), 0, oRecordset!IdUsuario)
       oTabla.IdEstadoHis = IIf(IsNull(oRecordset!IdEstadoHis), 0, oRecordset!IdEstadoHis)
       oTabla.IdMedico = IIf(IsNull(oRecordset!IdMedico), 0, oRecordset!IdMedico)
       oTabla.IdEstablecimiento = IIf(IsNull(oRecordset!IdEstablecimiento), 0, oRecordset!IdEstablecimiento)
       oTabla.IdServicio = IIf(IsNull(oRecordset!IdServicio), 0, oRecordset!IdServicio)
       oTabla.FechaCreacion = IIf(IsNull(oRecordset!FechaCreacion), 0, oRecordset!FechaCreacion)
   SeleccionarPorId = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOHIS_Cabecera)
   oTabla.IdHisCabecera = 0
   oTabla.IdHisLote = 0
   oTabla.NroHojaHis = 0
   oTabla.NroFormato = 0
   oTabla.IdTurno = 0
   oTabla.IdUsuario = 0
   oTabla.IdEstadoHis = 0
   oTabla.IdMedico = 0
   oTabla.IdEstablecimiento = 0
   oTabla.IdServicio = 0
   oTabla.FechaCreacion = Now
End Sub

Function ObtenerDatosEstablecimientoPorUsuario(ml_idUsuario As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "EmpleadosObtenerDatosEstablecimientoPorUsuario"
        Set oParameter = .CreateParameter("@ml_idUsuario", adInteger, adParamInput, 0, ml_idUsuario): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ObtenerDatosEstablecimientoPorUsuario = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function ObtenerDatosDigitador(ml_idUsuario As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "EmpleadosObtenerDatosDigitador"
        Set oParameter = .CreateParameter("@ml_idUsuario", adInteger, adParamInput, 0, ml_idUsuario): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ObtenerDatosDigitador = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function ObtenerDatosNroFormatoLibre(mi_anio As Integer, ml_IdEstablecimientoActual As Long) As String
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim NroFormatoLibre As String
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "his_cabeceraObtenerDatosNroFormatoLibre"
        Set oParameter = .CreateParameter("@mi_anio", adInteger, adParamInput, 0, mi_anio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@ml_IdEstablecimientoActual", adInteger, adParamInput, 0, ml_IdEstablecimientoActual): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   NroFormatoLibre = ""
   If oRecordset.RecordCount > 0 Then
        NroFormatoLibre = CStr(oRecordset.Fields(0))
   End If
   ObtenerDatosNroFormatoLibre = NroFormatoLibre
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function ConsultarRegistroFiltroAtenciones(IdEstablecimiento As Long, NombreLote As String, Anio As String, IdMes As Integer)
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "HIS_LotesConsultarRegistroFiltroAtenciones"
        Set oParameter = .CreateParameter("@Anio", adInteger, adParamInput, 0, Val(Anio)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdMes", adInteger, adParamInput, 0, IdMes): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdEstablecimiento", adInteger, adParamInput, 0, IdEstablecimiento): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@NombreLote", adVarChar, adParamInput, 3, NombreLote): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ConsultarRegistroFiltroAtenciones = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Public Function ObtenerListaMedicosMR(IdEpecialidad As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "EmpleadosObtenerListaMedicosMR"
        Set oParameter = .CreateParameter("@IdEpecialidad", adInteger, adParamInput, 0, IdEpecialidad): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ObtenerListaMedicosMR = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

'============================== LISTA DE TODOS LOS LISTADOS DEL FORMULARIO PRINCIPAL =========================
'LISTA DE TURNOS

'LISTA DE TIPOS DE DOCUMENTO
Public Function ListaTiposDocumentos() As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "TiposDocIdentidadFiltrarTodos"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ListaTiposDocumentos = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

'LISTA DE FUENTES DE FINANCIAMIENTO
Public Function ListaFuentesFinanciamiento() As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "His_FuentesFinanciamiento"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ListaFuentesFinanciamiento = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

'LISTA DE ETNIAS
Public Function ListaEtnias() As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "his_tabetniaSeleccionarTodos"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ListaEtnias = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

''LISTA DE TIPOS DE EDAD
'Public Function ListaTiposEdad() As Recordset
'Dim oRecordset As New ADODB.Recordset
'Dim oCommand As New ADODB.Command
'Dim sSQL As String
'
'On Error GoTo ManejadorDeError
'
'sSQL = ""
'
'With oCommand
'    .CommandType = adCmdText
'    Set .ActiveConnection = mo_Conexion
'    .CommandText = sSQL
'    Set oRecordset = .Execute
'    Set oRecordset.ActiveConnection = Nothing
'End With
'
'If oRecordset.RecordCount <> 0 Then
'    Set ListaTiposEdad = oRecordset
'End If
'
'ms_MensajeError = ""
'Exit Function
'ManejadorDeError:
'ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
'End Function

'LISTA DE GENERO
Public Function ListaGeneros() As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "HIS_TiposSexo"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ListaGeneros = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

'LISTA DE ESTADO (ESTABLECIMIENTO - SERVICIO)
Public Function ListaEstadosPaciente() As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "HIS_TiposCondicionPaciente"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ListaEstadosPaciente = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Public Function ListaPaises() As Recordset
'Dim oRecordset As New ADODB.Recordset
'Dim oCommand As New ADODB.Command
'Dim sSQL As String
'
'On Error GoTo ManejadorDeError
'
'sSQL = "select * from paises"
'
'With oCommand
'    .CommandType = adCmdText
'    Set .ActiveConnection = mo_Conexion
'    .CommandText = sSQL
'    Set oRecordset = .Execute
'    Set oRecordset.ActiveConnection = Nothing
'End With
'
'If oRecordset.RecordCount <> 0 Then
'    Set ListaPaises = oRecordset
'End If
'
'ms_MensajeError = ""
'Exit Function
'ManejadorDeError:
'ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Stop
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "PaisesFiltrarTodos"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ListaPaises = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function
'Modificado por YEPE NOVIEMBRE  Inicio
Public Function ListaPaisesPorCodigoNombre(ByVal codigo As String, ByVal nombre As String) As Recordset
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open SIGHENTIDADES.CadenaConexion
On Error GoTo ManejadorDeError
Set ListaPaisesPorCodigoNombre = Nothing
With oCommand
    .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
    .CommandTimeout = 150
    .CommandText = "ListaPaisesPorCodigoNombre"
     Set oParameter = .CreateParameter("@codigo", adVarChar, adParamInput, 3, UCase(codigo)): .Parameters.Append oParameter
     Set oParameter = .CreateParameter("@nombre", adVarChar, adParamInput, 40, UCase(nombre)): .Parameters.Append oParameter
     Set oRecordset = .Execute
     Set oRecordset.ActiveConnection = Nothing
End With
   Set ListaPaisesPorCodigoNombre = oRecordset

If oRecordset.RecordCount <> 0 Then
    Set ListaPaisesPorCodigoNombre = oRecordset
End If
   Set oCommand = Nothing
   Set oParameter = Nothing
   Set oRecordset = Nothing
ms_MensajeError = ""
Exit Function
ManejadorDeError:
ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function


Function ObtenerListaCodigosProductosHisPorCodigoyNombre(ByVal codigodiagcpt As String, ByVal descripciondiagcpt As String) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "HIS_ObtenerProductosHisPorCodigoyNombre"
        Set oParameter = .CreateParameter("@codigodiagcpt", adVarChar, adParamInput, 20, CStr(codigodiagcpt)): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@descripciondiagcpt", adVarChar, adParamInput, 40, descripciondiagcpt): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ObtenerListaCodigosProductosHisPorCodigoyNombre = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function
'------------------------------------------------------------------------------------

'OBTENER LOS DATOS DE LOS PAISES DEPENDIENDO DEL ID
Public Function ObtenerDatosPais(IdPais As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "PaisesSeleccionarPorId"
        Set oParameter = .CreateParameter("@IdPais", adInteger, adParamInput, 0, IdPais): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set ObtenerDatosPais = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Public Function His_ObtenerIdPaisPorCodNac(codigo As String) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "His_ObtenerIdPaisPorCodNac"
        Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 3, codigo): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set His_ObtenerIdPaisPorCodNac = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Public Function ExisteHoja(ByVal oTabla As DOHIS_Cabecera) As Integer
'Dim oRecordset As New ADODB.Recordset
'Dim oCommand As New ADODB.Command
'Dim sSQL As String
'
'On Error GoTo ManejadorDeError
'
'sSQL = "SELECT IdHisCabecera FROM his_cabecera "
'sSQL = sSQL & " WHERE IdHisLote = " & oTabla.IdHisLote
'sSQL = sSQL & " AND NroHojaHIS = " & oTabla.NroHojaHis
'
'With oCommand
'    .CommandType = adCmdText
'    Set .ActiveConnection = mo_Conexion
'    .CommandText = sSQL
'    Set oRecordset = .Execute
'    Set oRecordset.ActiveConnection = Nothing
'End With
'
'If oRecordset.RecordCount <> 0 Then
'    ExisteHoja = CInt(oRecordset!IdHisCabecera)
'Else
'    ExisteHoja = 0
'End If
'
'ms_MensajeError = ""
'Exit Function
'ManejadorDeError:
'ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Stop
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "his_cabeceraSeleccionarXnroHojaHIS"
        Set oParameter = .CreateParameter("@IdHisLote", adInteger, adParamInput, 0, oTabla.IdHisLote): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@NroHojaHis", adInteger, adParamInput, 0, oTabla.NroHojaHis): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   If oRecordset.RecordCount <> 0 Then
        ExisteHoja = CInt(oRecordset!IdHisCabecera)
   Else
        ExisteHoja = 0
   End If
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


 

