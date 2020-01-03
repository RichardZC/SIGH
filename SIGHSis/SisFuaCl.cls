VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SisFuaCl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para el Mantenimiento de la tabla SisFUA
'        Programado por: Barrantes D
'        Fecha: Enero 2013
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
'   Descripción:    Inserta un registro a la tabla SisFua
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DoSisFua) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "SisFuaAgregar"
   Set oParameter = .CreateParameter("@FuaDisa", adVarChar, adParamInput, 3, IIf(oTabla.FuaDisa = "", Null, oTabla.FuaDisa))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaLote", adVarChar, adParamInput, 2, IIf(oTabla.FuaLote = "", Null, oTabla.FuaLote))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaNumeroInicial", adVarChar, adParamInput, 16, IIf(oTabla.FuaNumeroInicial = "", Null, oTabla.FuaNumeroInicial))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaNumeroFinal", adVarChar, adParamInput, 16, IIf(oTabla.FuaNumeroFinal = "", Null, oTabla.FuaNumeroFinal))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaUltimoGenerado", adVarChar, adParamInput, 16, IIf(oTabla.FuaUltimoGenerado = "", Null, oTabla.FuaUltimoGenerado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
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
'   Descripción:    modifica un registro a la tabla SisFua
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DoSisFua) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "SisFuaModificar"
   Set oParameter = .CreateParameter("@FuaDisa", adVarChar, adParamInput, 3, IIf(oTabla.FuaDisa = "", Null, oTabla.FuaDisa))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaLote", adVarChar, adParamInput, 2, IIf(oTabla.FuaLote = "", Null, oTabla.FuaLote))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaNumeroInicial", adVarChar, adParamInput, 16, IIf(oTabla.FuaNumeroInicial = "", Null, oTabla.FuaNumeroInicial))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaNumeroFinal", adVarChar, adParamInput, 16, IIf(oTabla.FuaNumeroFinal = "", Null, oTabla.FuaNumeroFinal))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FuaUltimoGenerado", adVarChar, adParamInput, 16, IIf(oTabla.FuaUltimoGenerado = "", Null, oTabla.FuaUltimoGenerado))
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
'   Descripción:    elimia o borra un registro a la tabla SisFua
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DoSisFua) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "SisFuaEliminar"
   Set oParameter = .CreateParameter("@FuaDisa", adVarChar, adParamInput, 3, oTabla.FuaDisa)
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
'   Descripción:    Selecciona un unico registro de la tabla SisFua
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DoSisFua) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "SisFuaSeleccionarPorId"
     Set oParameter = .CreateParameter("@FuaDisa", adVarChar, adParamInput, 3, oTabla.FuaDisa)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.FuaDisa = IIf(IsNull(oRecordset!FuaDisa), "", oRecordset!FuaDisa)
       oTabla.FuaLote = IIf(IsNull(oRecordset!FuaLote), "", oRecordset!FuaLote)
       oTabla.FuaNumeroInicial = IIf(IsNull(oRecordset!FuaNumeroInicial), "", oRecordset!FuaNumeroInicial)
       oTabla.FuaNumeroFinal = IIf(IsNull(oRecordset!FuaNumeroFinal), "", oRecordset!FuaNumeroFinal)
       oTabla.FuaUltimoGenerado = IIf(IsNull(oRecordset!FuaUltimoGenerado), "", oRecordset!FuaUltimoGenerado)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DoSisFua)
   oTabla.FuaDisa = ""
   oTabla.FuaLote = ""
   oTabla.FuaNumeroInicial = ""
   oTabla.FuaNumeroFinal = ""
   oTabla.FuaUltimoGenerado = ""
End Sub

Function SeleccionarTodos(ByVal oTabla As DoSisFua) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarTodos = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "SisFuaSeleccionarTodos"
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.FuaDisa = IIf(IsNull(oRecordset!FuaDisa), "", oRecordset!FuaDisa)
       oTabla.FuaLote = IIf(IsNull(oRecordset!FuaLote), "", oRecordset!FuaLote)
       oTabla.FuaNumeroInicial = IIf(IsNull(oRecordset!FuaNumeroInicial), "", oRecordset!FuaNumeroInicial)
       oTabla.FuaNumeroFinal = IIf(IsNull(oRecordset!FuaNumeroFinal), "", oRecordset!FuaNumeroFinal)
       oTabla.FuaUltimoGenerado = IIf(IsNull(oRecordset!FuaUltimoGenerado), "", oRecordset!FuaUltimoGenerado)
   SeleccionarTodos = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

