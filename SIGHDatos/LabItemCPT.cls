VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LabItemCPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla LabItemsCpt
'        Programado por: Bonilla A
'        Fecha: Mayo 2009
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
'   Descripción:    Inserta un registro a la tabla LabItemsCpt
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DoLabItemsCPT) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "LabItemsCptAgregar"
   Set oParameter = .CreateParameter("@idProductoCpt", adInteger, adParamInput, 0, IIf(oTabla.idProductoCpt = 0, Null, oTabla.idProductoCpt))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ordenXresultado", adInteger, adParamInput, 0, IIf(oTabla.ordenXresultado = 0, Null, oTabla.ordenXresultado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idGrupo", adInteger, adParamInput, 0, IIf(oTabla.IdGrupo = 0, Null, oTabla.IdGrupo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idItemGrupo", adInteger, adParamInput, 0, IIf(oTabla.idItemGrupo = 0, Null, oTabla.idItemGrupo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idItem", adInteger, adParamInput, 0, IIf(oTabla.idItem = 0, Null, oTabla.idItem))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ValorSiEsCombo", adVarChar, adParamInput, 100, IIf(oTabla.ValorSiEsCombo = "", Null, oTabla.ValorSiEsCombo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ValorReferencial", adVarChar, adParamInput, 100, IIf(oTabla.ValorReferencial = "", Null, oTabla.ValorReferencial))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Metodo", adVarChar, adParamInput, 50, IIf(oTabla.Metodo = "", Null, oTabla.Metodo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SoloNumero", adBoolean, adParamInput, 0, IIf(oTabla.SoloNumero = 0, Null, oTabla.SoloNumero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SoloTexto", adBoolean, adParamInput, 0, IIf(oTabla.SoloTexto = 0, Null, oTabla.SoloTexto))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SoloCombo", adBoolean, adParamInput, 0, IIf(oTabla.SoloCombo = 0, Null, oTabla.SoloCombo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SoloCheck", adBoolean, adParamInput, 0, IIf(oTabla.SoloCheck = 0, Null, oTabla.SoloCheck))
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
'   Descripción:    modifica un registro a la tabla LabItemsCpt
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DoLabItemsCPT) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "LabItemsCptModificar"
   Set oParameter = .CreateParameter("@idProductoCpt", adInteger, adParamInput, 0, IIf(oTabla.idProductoCpt = 0, Null, oTabla.idProductoCpt))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ordenXresultado", adInteger, adParamInput, 0, IIf(oTabla.ordenXresultado = 0, Null, oTabla.ordenXresultado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idGrupo", adInteger, adParamInput, 0, IIf(oTabla.IdGrupo = 0, Null, oTabla.IdGrupo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idItemGrupo", adInteger, adParamInput, 0, IIf(oTabla.idItemGrupo = 0, Null, oTabla.idItemGrupo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idItem", adInteger, adParamInput, 0, IIf(oTabla.idItem = 0, Null, oTabla.idItem))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ValorSiEsCombo", adVarChar, adParamInput, 100, IIf(oTabla.ValorSiEsCombo = "", Null, oTabla.ValorSiEsCombo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ValorReferencial", adVarChar, adParamInput, 100, IIf(oTabla.ValorReferencial = "", Null, oTabla.ValorReferencial))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Metodo", adVarChar, adParamInput, 50, IIf(oTabla.Metodo = "", Null, oTabla.Metodo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SoloNumero", adBoolean, adParamInput, 0, IIf(oTabla.SoloNumero = 0, Null, oTabla.SoloNumero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SoloTexto", adBoolean, adParamInput, 0, IIf(oTabla.SoloTexto = 0, Null, oTabla.SoloTexto))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SoloCombo", adBoolean, adParamInput, 0, IIf(oTabla.SoloCombo = 0, Null, oTabla.SoloCombo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SoloCheck", adBoolean, adParamInput, 0, IIf(oTabla.SoloCheck = 0, Null, oTabla.SoloCheck))
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
'   Descripción:    elimia o borra un registro a la tabla LabItemsCpt
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DoLabItemsCPT) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "LabItemsCptEliminar"
   Set oParameter = .CreateParameter("@idProductoCpt", adInteger, adParamInput, 0, oTabla.idProductoCpt)
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
'   Descripción:    Selecciona un unico registro de la tabla LabItemsCpt
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal widProductoCpt As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
  ' Set SeleccionarPorId = Null
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "LabItemsCptSeleccionarPorId"
     Set oParameter = .CreateParameter("@idProductoCpt", adInteger, adParamInput, 0, widProductoCpt)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
'   If Not (oRecordset.EOF And oRecordset.BOF) Then
'       oTabla.idProductoCpt = IIf(IsNull(oRecordset!idProductoCpt), 0, oRecordset!idProductoCpt)
'       oTabla.ordenXresultado = IIf(IsNull(oRecordset!ordenXresultado), 0, oRecordset!ordenXresultado)
'       oTabla.IdGrupo = IIf(IsNull(oRecordset!IdGrupo), 0, oRecordset!IdGrupo)
'       oTabla.idItemGrupo = IIf(IsNull(oRecordset!idItemGrupo), 0, oRecordset!idItemGrupo)
'       oTabla.idItem = IIf(IsNull(oRecordset!idItem), 0, oRecordset!idItem)
'       oTabla.ValorSiEsCombo = IIf(IsNull(oRecordset!ValorSiEsCombo), "", oRecordset!ValorSiEsCombo)
'       oTabla.ValorReferencial = IIf(IsNull(oRecordset!ValorReferencial), "", oRecordset!ValorReferencial)
'       oTabla.Metodo = IIf(IsNull(oRecordset!Metodo), "", oRecordset!Metodo)
'       oTabla.SoloNumero = IIf(IsNull(oRecordset!SoloNumero), 0, oRecordset!SoloNumero)
'       oTabla.SoloTexto = IIf(IsNull(oRecordset!SoloTexto), 0, oRecordset!SoloTexto)
'       oTabla.SoloCombo = IIf(IsNull(oRecordset!SoloCombo), 0, oRecordset!SoloCombo)
'       oTabla.SoloCheck = IIf(IsNull(oRecordset!SoloCheck), 0, oRecordset!SoloCheck)
   Set SeleccionarPorId = oRecordset
'   End If
   ms_MensajeError = ""
  ' oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DoLabItemsCPT)
   oTabla.idProductoCpt = 0
   oTabla.ordenXresultado = 0
   oTabla.IdGrupo = 0
   oTabla.idItemGrupo = 0
   oTabla.idItem = 0
   oTabla.ValorSiEsCombo = ""
   oTabla.ValorReferencial = ""
   oTabla.Metodo = ""
   oTabla.SoloNumero = False
   oTabla.SoloTexto = False
   oTabla.SoloCombo = False
   oTabla.SoloCheck = False
End Sub

 Function ModificarUno(ByVal oTabla As DoLabItemsCPT) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
ModificarUno = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "LabItemsCptModificarUno"
   Set oParameter = .CreateParameter("@idProductoCpt", adInteger, adParamInput, 0, IIf(oTabla.idProductoCpt = 0, Null, oTabla.idProductoCpt))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ordenXresultado", adInteger, adParamInput, 0, IIf(oTabla.ordenXresultado = 0, Null, oTabla.ordenXresultado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idGrupo", adInteger, adParamInput, 0, IIf(oTabla.IdGrupo = 0, Null, oTabla.IdGrupo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idItemGrupo", adInteger, adParamInput, 0, IIf(oTabla.idItemGrupo = 0, Null, oTabla.idItemGrupo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idItem", adInteger, adParamInput, 0, IIf(oTabla.idItem = 0, Null, oTabla.idItem))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ValorSiEsCombo", adVarChar, adParamInput, 100, IIf(oTabla.ValorSiEsCombo = "", Null, oTabla.ValorSiEsCombo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ValorReferencial", adVarChar, adParamInput, 100, IIf(oTabla.ValorReferencial = "", Null, oTabla.ValorReferencial))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Metodo", adVarChar, adParamInput, 50, IIf(oTabla.Metodo = "", Null, oTabla.Metodo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SoloNumero", adBoolean, adParamInput, 0, IIf(oTabla.SoloNumero = 0, Null, oTabla.SoloNumero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SoloTexto", adBoolean, adParamInput, 0, IIf(oTabla.SoloTexto = 0, Null, oTabla.SoloTexto))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SoloCombo", adBoolean, adParamInput, 0, IIf(oTabla.SoloCombo = 0, Null, oTabla.SoloCombo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SoloCheck", adBoolean, adParamInput, 0, IIf(oTabla.SoloCheck = 0, Null, oTabla.SoloCheck))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
End With
 
ModificarUno = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

