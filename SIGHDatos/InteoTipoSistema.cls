VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InteoTipoSistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla InteoTipoSistema
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
'   Descripción:    Inserta un registro a la tabla InteoTipoSistema
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOInteoTipoSistema) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "InteoTipoSistemaAgregar"
   Set oParameter = .CreateParameter("@IdTipoSistema", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TipoSistema", adVarChar, adParamInput, 50, IIf(oTabla.TipoSistema = "", Null, oTabla.TipoSistema))
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
   oTabla.IdTipoSistema = .Parameters("@IdTipoSistema")
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
'   Descripción:    modifica un registro a la tabla InteoTipoSistema
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOInteoTipoSistema) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "InteoTipoSistemaModificar"
   Set oParameter = .CreateParameter("@IdTipoSistema", adInteger, adParamInput, 0, IIf(oTabla.IdTipoSistema = 0, Null, oTabla.IdTipoSistema))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@TipoSistema", adVarChar, adParamInput, 50, IIf(oTabla.TipoSistema = "", Null, oTabla.TipoSistema))
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
'   Descripción:    elimia o borra un registro a la tabla InteoTipoSistema
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOInteoTipoSistema) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "InteoTipoSistemaEliminar"
   Set oParameter = .CreateParameter("@IdTipoSistema", adInteger, adParamInput, 0, oTabla.IdTipoSistema)
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
'   Descripción:    Selecciona un unico registro de la tabla InteoTipoSistema
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOInteoTipoSistema) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "InteoTipoSistemaSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdTipoSistema", adInteger, adParamInput, 0, oTabla.IdTipoSistema)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdTipoSistema = IIf(IsNull(oRecordset!IdTipoSistema), 0, oRecordset!IdTipoSistema)
       oTabla.TipoSistema = IIf(IsNull(oRecordset!TipoSistema), "", oRecordset!TipoSistema)
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
Sub SetDefaults(ByVal oTabla As DOInteoTipoSistema)
   oTabla.IdTipoSistema = 0
   oTabla.TipoSistema = ""
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
        .CommandText = "InteoTipoSistemaSeleccionarTodos"
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarTodos = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function VerificarPorNombre(ByVal oTabla As DOInteoTipoSistema) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   VerificarPorNombre = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "InteoTipoSistemaVerificaNombre"
     Set oParameter = .CreateParameter("@IdTipoSistema", adInteger, adParamInput, 0, oTabla.IdTipoSistema)
     .Parameters.Append oParameter
     Set oParameter = .CreateParameter("@TipoSistema", adVarChar, adParamInput, 50, IIf(oTabla.TipoSistema = "", Null, oTabla.TipoSistema))
    .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdTipoSistema = IIf(IsNull(oRecordset!IdTipoSistema), 0, oRecordset!IdTipoSistema)
       oTabla.TipoSistema = IIf(IsNull(oRecordset!TipoSistema), "", oRecordset!TipoSistema)
       oTabla.EsActivo = IIf(IsNull(oRecordset!EsActivo), 0, oRecordset!EsActivo)
       oTabla.FechaCrea = IIf(IsNull(oRecordset!FechaCrea), 0, oRecordset!FechaCrea)
       oTabla.FechaEdita = IIf(IsNull(oRecordset!FechaEdita), 0, oRecordset!FechaEdita)
        VerificarPorNombre = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

