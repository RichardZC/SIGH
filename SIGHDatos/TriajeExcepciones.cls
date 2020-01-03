VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TriajeExcepciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla TriajeExcepciones
'        Programado por: Garay M
'        Fecha: Agosto 2014
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
'   Descripción:    Inserta un registro a la tabla TriajeExcepciones
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOTriajeExcepciones) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "TriajeExcepcionesAgregar"
   Set oParameter = .CreateParameter("@IdTriajeExcepciones", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdTriajeVariable", adInteger, adParamInput, 0, IIf(oTabla.IdTriajeVariable = 0, Null, oTabla.IdTriajeVariable))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EdadInicialEnDia", adInteger, adParamInput, 0, IIf(oTabla.EdadInicialEnDia = 0, Null, oTabla.EdadInicialEnDia))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EdadFinalEnDia", adInteger, adParamInput, 0, IIf(oTabla.EdadFinalEnDia = 0, Null, oTabla.EdadFinalEnDia))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EsDatoObligatorio", adBoolean, adParamInput, 0, IIf(oTabla.EsDatoObligatorio = 0, Null, oTabla.EsDatoObligatorio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.IdTriajeExcepciones = .Parameters("@IdTriajeExcepciones")
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
'   Descripción:    modifica un registro a la tabla TriajeExcepciones
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOTriajeExcepciones) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "TriajeExcepcionesModificar"
   Set oParameter = .CreateParameter("@IdTriajeExcepciones", adInteger, adParamInput, 0, IIf(oTabla.IdTriajeExcepciones = 0, Null, oTabla.IdTriajeExcepciones))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdTriajeVariable", adInteger, adParamInput, 0, IIf(oTabla.IdTriajeVariable = 0, Null, oTabla.IdTriajeVariable))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EdadInicialEnDia", adInteger, adParamInput, 0, IIf(oTabla.EdadInicialEnDia = 0, Null, oTabla.EdadInicialEnDia))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EdadFinalEnDia", adInteger, adParamInput, 0, IIf(oTabla.EdadFinalEnDia = 0, Null, oTabla.EdadFinalEnDia))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EsDatoObligatorio", adBoolean, adParamInput, 0, IIf(oTabla.EsDatoObligatorio = 0, Null, oTabla.EsDatoObligatorio))
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
'   Descripción:    elimia o borra un registro a la tabla TriajeExcepciones
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOTriajeExcepciones) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "TriajeExcepcionesEliminar"
   Set oParameter = .CreateParameter("@IdTriajeExcepciones", adInteger, adParamInput, 0, oTabla.IdTriajeExcepciones)
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
'   Descripción:    Selecciona un unico registro de la tabla TriajeExcepciones
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOTriajeExcepciones) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "TriajeExcepcionesSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdTriajeExcepciones", adInteger, adParamInput, 0, oTabla.IdTriajeExcepciones)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdTriajeExcepciones = IIf(IsNull(oRecordset!IdTriajeExcepciones), 0, oRecordset!IdTriajeExcepciones)
       oTabla.IdTriajeVariable = IIf(IsNull(oRecordset!IdTriajeVariable), 0, oRecordset!IdTriajeVariable)
       oTabla.EdadInicialEnDia = IIf(IsNull(oRecordset!EdadInicialEnDia), 0, oRecordset!EdadInicialEnDia)
       oTabla.EdadFinalEnDia = IIf(IsNull(oRecordset!EdadFinalEnDia), 0, oRecordset!EdadFinalEnDia)
       oTabla.EsDatoObligatorio = IIf(IsNull(oRecordset!EsDatoObligatorio), 0, oRecordset!EsDatoObligatorio)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOTriajeExcepciones)
   oTabla.IdTriajeExcepciones = 0
   oTabla.IdTriajeVariable = 0
   oTabla.EdadInicialEnDia = 0
   oTabla.EdadFinalEnDia = 0
   oTabla.EsDatoObligatorio = False
End Sub

Function SeleccionarTodos() As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarTodos = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "TriajeExcepcionesSeleccionarTodos"
     Set oRecordset = .Execute
     Set oRecordset.ActiveConnection = Nothing
   End With
   
   Set SeleccionarTodos = oRecordset
 
   ms_MensajeError = ""
   Set oCommand = Nothing
   Set oRecordset = Nothing
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
