VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "labItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organizaci�n: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla LabItems
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
'   Funci�n:        Insertar
'   Descripci�n:    Inserta un registro a la tabla LabItems
'   Par�metros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOLabItems) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "LabItemsAgregar"
   Set oParameter = .CreateParameter("@idItem", adInteger, adParamInput, 0, IIf(oTabla.idItem = 0, Null, oTabla.idItem))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Item", adVarChar, adParamInput, 100, IIf(oTabla.Item = "", Null, oTabla.Item))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idProductoCpt", adInteger, adParamInput, 0, IIf(oTabla.idProductoCpt = 0, Null, oTabla.idProductoCpt))
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
'   Funci�n:        Modificar
'   Descripci�n:    modifica un registro a la tabla LabItems
'   Par�metros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOLabItems) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "LabItemsModificar"
   Set oParameter = .CreateParameter("@idItem", adInteger, adParamInput, 0, IIf(oTabla.idItem = 0, Null, oTabla.idItem))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Item", adVarChar, adParamInput, 100, IIf(oTabla.Item = "", Null, oTabla.Item))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idProductoCpt", adInteger, adParamInput, 0, IIf(oTabla.idProductoCpt = 0, Null, oTabla.idProductoCpt))
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
'   Funci�n:        Eliminar
'   Descripci�n:    elimia o borra un registro a la tabla LabItems
'   Par�metros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOLabItems) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "LabItemsEliminar"
   Set oParameter = .CreateParameter("@idItem", adInteger, adParamInput, 0, oTabla.idItem)
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
'   Funci�n:        SeleccionarPorId
'   Descripci�n:    Selecciona un unico registro de la tabla LabItems
'   Par�metros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOLabItems) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "LabItemsSeleccionarPorId"
     Set oParameter = .CreateParameter("@idItem", adInteger, adParamInput, 0, oTabla.idItem)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.idItem = IIf(IsNull(oRecordset!idItem), 0, oRecordset!idItem)
       oTabla.Item = IIf(IsNull(oRecordset!Item), "", oRecordset!Item)
       oTabla.idProductoCpt = IIf(IsNull(oRecordset!idProductoCpt), 0, oRecordset!idProductoCpt)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Funci�n:        SeleccionarTodos
'   Descripci�n:    Selecciona todos los registros de la tabla LabItems
'------------------------------------------------------------------------------------

Function SeleccionarTodos(wcriterio As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    
    If wcriterio <> "" Then
        wcriterio = " where " & wcriterio
    End If
    Set SeleccionarTodos = Nothing
    With oCommand
      .CommandType = adCmdStoredProc
      Set .ActiveConnection = mo_Conexion
      .CommandText = "LabItemsSeleccionarTodos"
      Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 200, wcriterio)
      .Parameters.Append oParameter
      Set oRecordset = .Execute
      Set oRecordset.ActiveConnection = Nothing
    End With
   
    Set SeleccionarTodos = oRecordset
 
   ms_MensajeError = ""
   Set oCommand = Nothing
   Set oRecordset = Nothing
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte t�cnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOLabItems)
   oTabla.idItem = 0
   oTabla.Item = ""
   oTabla.idProductoCpt = 0
End Sub
