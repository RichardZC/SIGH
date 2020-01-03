VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ImagSala"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla ImagSala
'        Programado por: Garay M
'        Fecha: Noviembre 2014
'
'------------------------------------------------------------------------------------
Option Explicit
Dim mo_Conexion As ADODB.Connection
Dim ms_MensajeError As String
Property Set Conexion(oValue As ADODB.Connection)
   Set mo_Conexion = oValue
End Property
Property Get Conexion() As ADODB.Connection
   Set Conexion = mo_Conexion
End Property
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property

'------------------------------------------------------------------------------------
'   Función:        Insertar
'   Descripción:    Inserta un registro a la tabla ImagSala
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOImagSala) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "ImagSalaAgregar"
   Set oParameter = .CreateParameter("@IdSala", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdTipoModalidadSala", adInteger, adParamInput, 0, IIf(oTabla.IdTipoModalidadSala = 0, Null, oTabla.IdTipoModalidadSala))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Sala", adVarChar, adParamInput, 50, IIf(oTabla.Sala = "", Null, oTabla.Sala))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 15, IIf(oTabla.Codigo = "", Null, oTabla.Codigo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EsActivo", adBoolean, adParamInput, 0, IIf(oTabla.EsActivo = 0, Null, oTabla.EsActivo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaCrea", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaCrea = 0, Null, oTabla.FechaCrea))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaEdita", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaEdita = 0, Null, oTabla.FechaEdita))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.IdSala = .Parameters("@IdSala")
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
'   Descripción:    modifica un registro a la tabla ImagSala
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOImagSala) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "ImagSalaModificar"
   Set oParameter = .CreateParameter("@IdSala", adInteger, adParamInput, 0, IIf(oTabla.IdSala = 0, Null, oTabla.IdSala))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdTipoModalidadSala", adInteger, adParamInput, 0, IIf(oTabla.IdTipoModalidadSala = 0, Null, oTabla.IdTipoModalidadSala))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Sala", adVarChar, adParamInput, 50, IIf(oTabla.Sala = "", Null, oTabla.Sala))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 15, IIf(oTabla.Codigo = "", Null, oTabla.Codigo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EsActivo", adBoolean, adParamInput, 0, IIf(oTabla.EsActivo = 0, Null, oTabla.EsActivo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaCrea", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaCrea = 0, Null, oTabla.FechaCrea))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaEdita", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaEdita = 0, Null, oTabla.FechaEdita))
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
'   Descripción:    elimia o borra un registro a la tabla ImagSala
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOImagSala) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "ImagSalaEliminar"
   Set oParameter = .CreateParameter("@IdSala", adInteger, adParamInput, 0, oTabla.IdSala)
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
'   Descripción:    Selecciona un unico registro de la tabla ImagSala
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOImagSala) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "ImagSalaSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdSala", adInteger, adParamInput, 0, oTabla.IdSala)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdSala = IIf(IsNull(oRecordset!IdSala), 0, oRecordset!IdSala)
       oTabla.IdTipoModalidadSala = IIf(IsNull(oRecordset!IdTipoModalidadSala), 0, oRecordset!IdTipoModalidadSala)
       oTabla.Sala = IIf(IsNull(oRecordset!Sala), "", oRecordset!Sala)
       oTabla.Codigo = IIf(IsNull(oRecordset!Codigo), "", oRecordset!Codigo)
       oTabla.EsActivo = IIf(IsNull(oRecordset!EsActivo), 0, oRecordset!EsActivo)
       oTabla.FechaCrea = IIf(IsNull(oRecordset!FechaCrea), 0, oRecordset!FechaCrea)
       oTabla.FechaEdita = IIf(IsNull(oRecordset!FechaEdita), 0, oRecordset!FechaEdita)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOImagSala)
   oTabla.IdSala = 0
   oTabla.IdTipoModalidadSala = 0
   oTabla.Sala = ""
   oTabla.Codigo = ""
   oTabla.EsActivo = False
   oTabla.FechaCrea = Now
   oTabla.FechaEdita = Now
End Sub

Function SeleccionarTodos() As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oRecordset As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "ImagSalaSeleccionarTodos"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarTodos = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function VerificarPorNombre(ByVal oTabla As DOImagSala) As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
 
    Set VerificarPorNombre = Nothing
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "ImagSalaVerificaNombre"
        Set oParameter = .CreateParameter("@IdSala", adInteger, adParamInput, 0, oTabla.IdSala)
        .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Sala", adVarChar, adParamInput, 50, IIf(oTabla.Sala = "", Null, oTabla.Sala))
        .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 15, IIf(oTabla.Codigo = "", Null, oTabla.Codigo))
        .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
    Set VerificarPorNombre = oRecordset
   Set oCommand = Nothing
   ms_MensajeError = ""
 
    Exit Function
ManejadorDeError:
    ms_MensajeError = Err.Number & " " + Err.Description
    Exit Function
End Function

Function FiltrarPorTipoModalidad(ByVal oTabla As DOImagSala) As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oRecordset As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "ImagSalaSeleccionarPorTipoModalidad"
        Set oParameter = .CreateParameter("@IdTipoModalidadSala", adVarChar, adParamInput, 50, oTabla.IdTipoModalidadSala)
        .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FiltrarPorTipoModalidad = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function FiltrarTodos(ByVal oTabla As DOImagSala) As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oRecordset As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "ImagSalaFiltrarTodos"
        Set oParameter = .CreateParameter("@Sala", adVarChar, adParamInput, 50, IIf(oTabla.Sala = "", Null, oTabla.Sala))
        .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 15, IIf(oTabla.Codigo = "", Null, oTabla.Codigo))
        .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FiltrarTodos = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function
