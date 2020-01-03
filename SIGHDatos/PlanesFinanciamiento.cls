VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlanesFinanciamiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: PlanesFinanciamiento
'        Autor: William Castro Grijalva
'        Fecha: 30/08/2004 11:37:57 a.m.
'        Empresa: Digital Works Corporation
'        Todos los derechos reservados
'        Control De Cambios:
'------------------------------------------------------------------------------------
'        Autor                      Fecha                      Cambio
'------------------------------------------------------------------------------------
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
'   Descripción:    Inserta un registro a la tabla PlanesFinanciamiento
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOPlanFinanciamiento) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "PlanesFinanciamientoAgregar"
           Set oParameter = .CreateParameter("@IdPlan", adInteger, adParamInput, 4, IIf(oTabla.IdPlan = 0, Null, oTabla.IdPlan)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdFuenteFinanciamiento", adInteger, adParamInput, 4, IIf(oTabla.IdFuenteFinanciamiento = 0, Null, oTabla.IdFuenteFinanciamiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 4, IIf(oTabla.IdTipoFinanciamiento = 0, Null, oTabla.IdTipoFinanciamiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPlanFinanciamiento", adInteger, adParamOutput, 0): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdPlanFinanciamiento = .Parameters("@IdPlanFinanciamiento")
   End With
 
   Insertar = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'------------------------------------------------------------------------------------
'   Función:        Modificar
'   Descripción:    modifica un registro a la tabla PlanesFinanciamiento
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOPlanFinanciamiento) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "PlanesFinanciamientoModificar"
           Set oParameter = .CreateParameter("@IdPlan", adInteger, adParamInput, 4, IIf(oTabla.IdPlan = 0, Null, oTabla.IdPlan)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdFuenteFinanciamiento", adInteger, adParamInput, 4, IIf(oTabla.IdFuenteFinanciamiento = 0, Null, oTabla.IdFuenteFinanciamiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 4, IIf(oTabla.IdTipoFinanciamiento = 0, Null, oTabla.IdTipoFinanciamiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPlanFinanciamiento", adInteger, adParamInput, 4, IIf(oTabla.IdPlanFinanciamiento = 0, Null, oTabla.IdPlanFinanciamiento)): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
       .Parameters.Append oParameter
       .Execute
   End With
 
   Modificar = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        Eliminar
'   Descripción:    elimia o borra un registro a la tabla PlanesFinanciamiento
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOPlanFinanciamiento) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "PlanesFinanciamientoEliminar"
           Set oParameter = .CreateParameter("@IdPlanFinanciamiento", adInteger, adParamInput, 4, IIf(oTabla.IdPlanFinanciamiento = 0, Null, oTabla.IdPlanFinanciamiento)): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
       .Parameters.Append oParameter
       .Execute
   End With
 
   Eliminar = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        SeleccionarPorId
'   Descripción:    Selecciona un unico registro de la tabla PlanesFinanciamiento
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOPlanFinanciamiento) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "PlanesFinanciamientoSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdPlanFinanciamiento", adInteger, adParamInput, 4, oTabla.IdPlanFinanciamiento): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.IdPlan = IIf(IsNull(oRecordset!IdPlan), 0, oRecordset!IdPlan)
       oTabla.IdFuenteFinanciamiento = IIf(IsNull(oRecordset!IdFuenteFinanciamiento), 0, oRecordset!IdFuenteFinanciamiento)
       oTabla.IdTipoFinanciamiento = IIf(IsNull(oRecordset!IdTipoFinanciamiento), 0, oRecordset!IdTipoFinanciamiento)
       oTabla.IdPlanFinanciamiento = IIf(IsNull(oRecordset!IdPlanFinanciamiento), 0, oRecordset!IdPlanFinanciamiento)
   Else
       SeleccionarPorId = False
   End If

   oRecordset.Close
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function EliminarPorIdPlan(ByVal lIdPlan As Long) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   EliminarPorIdPlan = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "PlanesFinanciamientoEliminarPorIdPlan"
        Set oParameter = .CreateParameter("@IdPlan", adInteger, adParamInput, 4, lIdPlan): .Parameters.Append oParameter
       .Execute
   End With
 
   EliminarPorIdPlan = True
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function SeleccionarPorIdPlan(lIdPlan As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarPorIdPlan = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "PlanesFinanciamientoSeleccionarPorIdPlan"
     Set oParameter = .CreateParameter("@IdPlan", adInteger, adParamInput, 4, lIdPlan): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
   
   Set oRecordset.ActiveConnection = Nothing
   Set SeleccionarPorIdPlan = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function SeleccionarPorTipoFinanciamiento(lIdTipoFinanciamiento As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarPorTipoFinanciamiento = Nothing
   With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandText = "PlanesFinanciamientoSeleccionarPorTipoFinanciamiento"
        Set oParameter = .CreateParameter("@lIdTipoFinanciamiento", adInteger, adParamInput, 0, IIf(lIdTipoFinanciamiento = 0, Null, lIdTipoFinanciamiento)): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   
   Set oRecordset.ActiveConnection = Nothing
   Set SeleccionarPorTipoFinanciamiento = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function SeleccionarPorTipoYFuenteFinanciamiento(lIdTipoFinanciamiento As Long, lIdFuenteFinanciamiento As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set SeleccionarPorTipoYFuenteFinanciamiento = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "PlanesFinanciamientoSeleccionarPorTipoYFuenteFinan"
     Set oParameter = .CreateParameter("@lIdTipoFinanciamiento", adInteger, adParamInput, 0, IIf(lIdTipoFinanciamiento = 0, Null, lIdTipoFinanciamiento)): .Parameters.Append oParameter
     Set oParameter = .CreateParameter("@lIdFuenteFinanciamiento", adInteger, adParamInput, 0, IIf(lIdFuenteFinanciamiento = 0, Null, lIdFuenteFinanciamiento)): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
   
   Set oRecordset.ActiveConnection = Nothing
   Set SeleccionarPorTipoYFuenteFinanciamiento = oRecordset
 
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

