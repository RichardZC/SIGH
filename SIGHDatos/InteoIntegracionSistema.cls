VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InteoIntegracionSistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla InteoIntegracionSistema
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
'   Descripción:    Inserta un registro a la tabla InteoIntegracionSistema
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOInteoIntegracionSistema) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "InteoIntegracionSistemaAgregar"
   Set oParameter = .CreateParameter("@IdIntegracionSistema", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdTipoSistema", adInteger, adParamInput, 0, IIf(oTabla.IdTipoSistema = 0, Null, oTabla.IdTipoSistema))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdProveedorSistema", adInteger, adParamInput, 0, IIf(oTabla.IdProveedorSistema = 0, Null, oTabla.IdProveedorSistema))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NombreUsuario", adVarChar, adParamInput, 35, IIf(oTabla.NombreUsuario = "", Null, oTabla.NombreUsuario))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ClaveUsuario", adVarChar, adParamInput, 50, IIf(oTabla.ClaveUsuario = "", Null, oTabla.ClaveUsuario))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EsProveedorActual", adBoolean, adParamInput, 0, oTabla.EsProveedorActual)
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
   oTabla.IdIntegracionSistema = .Parameters("@IdIntegracionSistema")
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
'   Descripción:    modifica un registro a la tabla InteoIntegracionSistema
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOInteoIntegracionSistema) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "InteoIntegracionSistemaModificar"
   Set oParameter = .CreateParameter("@IdIntegracionSistema", adInteger, adParamInput, 0, IIf(oTabla.IdIntegracionSistema = 0, Null, oTabla.IdIntegracionSistema))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdTipoSistema", adInteger, adParamInput, 0, IIf(oTabla.IdTipoSistema = 0, Null, oTabla.IdTipoSistema))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdProveedorSistema", adInteger, adParamInput, 0, IIf(oTabla.IdProveedorSistema = 0, Null, oTabla.IdProveedorSistema))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NombreUsuario", adVarChar, adParamInput, 35, IIf(oTabla.NombreUsuario = "", Null, oTabla.NombreUsuario))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ClaveUsuario", adVarChar, adParamInput, 50, IIf(oTabla.ClaveUsuario = "", Null, oTabla.ClaveUsuario))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EsProveedorActual", adBoolean, adParamInput, 0, oTabla.EsProveedorActual)
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
'   Descripción:    elimia o borra un registro a la tabla InteoIntegracionSistema
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOInteoIntegracionSistema) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "InteoIntegracionSistemaEliminar"
   Set oParameter = .CreateParameter("@IdIntegracionSistema", adInteger, adParamInput, 0, oTabla.IdIntegracionSistema)
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
'   Descripción:    Selecciona un unico registro de la tabla InteoIntegracionSistema
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOInteoIntegracionSistema) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "InteoIntegracionSistemaSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdIntegracionSistema", adInteger, adParamInput, 0, oTabla.IdIntegracionSistema)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdIntegracionSistema = IIf(IsNull(oRecordset!IdIntegracionSistema), 0, oRecordset!IdIntegracionSistema)
       oTabla.IdTipoSistema = IIf(IsNull(oRecordset!IdTipoSistema), 0, oRecordset!IdTipoSistema)
       oTabla.IdProveedorSistema = IIf(IsNull(oRecordset!IdProveedorSistema), 0, oRecordset!IdProveedorSistema)
       oTabla.NombreUsuario = IIf(IsNull(oRecordset!NombreUsuario), "", oRecordset!NombreUsuario)
       oTabla.ClaveUsuario = IIf(IsNull(oRecordset!ClaveUsuario), "", oRecordset!ClaveUsuario)
       oTabla.EsProveedorActual = IIf(IsNull(oRecordset!EsProveedorActual), 0, oRecordset!EsProveedorActual)
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
Sub SetDefaults(ByVal oTabla As DOInteoIntegracionSistema)
   oTabla.IdIntegracionSistema = 0
   oTabla.IdTipoSistema = 0
   oTabla.IdProveedorSistema = 0
   oTabla.NombreUsuario = ""
   oTabla.ClaveUsuario = ""
   oTabla.EsProveedorActual = False
   oTabla.EsActivo = False
   oTabla.FechaCrea = Now
   oTabla.FechaEdita = Now
End Sub


Function VerificarIntegracionSistema(ByVal oTabla As DOInteoIntegracionSistema) As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
 
    Set VerificarIntegracionSistema = Nothing
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "InteoIntegracionSistemaVerifica"
        Set oParameter = .CreateParameter("@IdIntegracionSistema", adInteger, adParamInput, 0, oTabla.IdIntegracionSistema)
        .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdTipoSistema", adInteger, adParamInput, 0, oTabla.IdTipoSistema)
        .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdProveedorSistema", adInteger, adParamInput, 0, oTabla.IdProveedorSistema)
        .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
    Set VerificarIntegracionSistema = oRecordset
   Set oCommand = Nothing
   ms_MensajeError = ""

    Exit Function
ManejadorDeError:
    ms_MensajeError = Err.Number & " " + Err.Description
    Exit Function
End Function


Function FiltrarTodos(ByVal oTabla As DOInteoIntegracionSistema) As ADODB.Recordset
On Error GoTo ManejadorDeError
    Dim oRecordset As New Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "InteoIntegracionSistemaFiltrarTodos"
        Set oParameter = .CreateParameter("@IdTipoSistema", adInteger, adParamInput, 0, IIf(oTabla.IdTipoSistema = 0, Null, oTabla.IdTipoSistema))
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


Function SeleccionarProveedorActual(ByVal oTabla As DOInteoIntegracionSistema) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarProveedorActual = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "InteoIntegracionSistemaProveedorActualPorTipo"
     Set oParameter = .CreateParameter("@IdTipoSistema", adInteger, adParamInput, 0, oTabla.IdTipoSistema)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdIntegracionSistema = IIf(IsNull(oRecordset!IdIntegracionSistema), 0, oRecordset!IdIntegracionSistema)
       oTabla.IdTipoSistema = IIf(IsNull(oRecordset!IdTipoSistema), 0, oRecordset!IdTipoSistema)
       oTabla.IdProveedorSistema = IIf(IsNull(oRecordset!IdProveedorSistema), 0, oRecordset!IdProveedorSistema)
       oTabla.NombreUsuario = IIf(IsNull(oRecordset!NombreUsuario), "", oRecordset!NombreUsuario)
'       oTabla.ClaveUsuario = IIf(IsNull(oRecordset!ClaveUsuario), "", oRecordset!ClaveUsuario)
       oTabla.EsProveedorActual = IIf(IsNull(oRecordset!EsProveedorActual), 0, oRecordset!EsProveedorActual)
       oTabla.EsActivo = IIf(IsNull(oRecordset!EsActivo), 0, oRecordset!EsActivo)
       oTabla.FechaCrea = IIf(IsNull(oRecordset!FechaCrea), 0, oRecordset!FechaCrea)
       oTabla.FechaEdita = IIf(IsNull(oRecordset!FechaEdita), 0, oRecordset!FechaEdita)
        SeleccionarProveedorActual = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
