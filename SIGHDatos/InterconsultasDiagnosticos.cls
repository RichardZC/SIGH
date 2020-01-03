VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InterconsultasDiagnosticos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: InterconsultasDiagnosticos
'        Autor: William Castro Grijalva
'        Fecha: 27/03/2005 01:53:22 a.m.
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
'   Descripción:    Inserta un registro a la tabla InterconsultasDiagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOInterconsultaDiagnostico) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "InterconsultasDiagnosticosAgregar"
           Set oParameter = .CreateParameter("@IdInterconsultaDiagnostico", adInteger, adParamOutput, 0): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdSubclasificacionDx", adInteger, adParamInput, 0, IIf(oTabla.IdSubclasificacionDx = 0, Null, oTabla.IdSubclasificacionDx)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdClasificacionDx", adInteger, adParamInput, 0, IIf(oTabla.IdClasificacionDx = 0, Null, oTabla.IdClasificacionDx)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.IdDiagnostico = 0, Null, oTabla.IdDiagnostico)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdInterconsulta", adInteger, adParamInput, 0, IIf(oTabla.IdInterconsulta = 0, Null, oTabla.IdInterconsulta)): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdInterconsultaDiagnostico = .Parameters("@IdInterconsultaDiagnostico")
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
'   Descripción:    modifica un registro a la tabla InterconsultasDiagnosticos
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOInterconsultaDiagnostico) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "InterconsultasDiagnosticosModificar"
           Set oParameter = .CreateParameter("@IdInterconsultaDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.IdInterconsultaDiagnostico = 0, Null, oTabla.IdInterconsultaDiagnostico)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdSubclasificacionDx", adInteger, adParamInput, 0, IIf(oTabla.IdSubclasificacionDx = 0, Null, oTabla.IdSubclasificacionDx)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdClasificacionDx", adInteger, adParamInput, 0, IIf(oTabla.IdClasificacionDx = 0, Null, oTabla.IdClasificacionDx)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.IdDiagnostico = 0, Null, oTabla.IdDiagnostico)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdInterconsulta", adInteger, adParamInput, 0, IIf(oTabla.IdInterconsulta = 0, Null, oTabla.IdInterconsulta)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla InterconsultasDiagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOInterconsultaDiagnostico) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "InterconsultasDiagnosticosEliminar"
           Set oParameter = .CreateParameter("@IdInterconsultaDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.IdInterconsultaDiagnostico = 0, Null, oTabla.IdInterconsultaDiagnostico)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla InterconsultasDiagnosticos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOInterconsultaDiagnostico) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "InterconsultasDiagnosticosSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdInterconsultaDiagnostico", adInteger, adParamInput, 0, oTabla.IdInterconsultaDiagnostico): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.IdInterconsultaDiagnostico = IIf(IsNull(oRecordset!IdInterconsultaDiagnostico), 0, oRecordset!IdInterconsultaDiagnostico)
       oTabla.IdSubclasificacionDx = IIf(IsNull(oRecordset!IdSubclasificacionDx), 0, oRecordset!IdSubclasificacionDx)
       oTabla.IdClasificacionDx = IIf(IsNull(oRecordset!IdClasificacionDx), 0, oRecordset!IdClasificacionDx)
       oTabla.IdDiagnostico = IIf(IsNull(oRecordset!IdDiagnostico), 0, oRecordset!IdDiagnostico)
       oTabla.IdInterconsulta = IIf(IsNull(oRecordset!IdInterconsulta), 0, oRecordset!IdInterconsulta)
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

