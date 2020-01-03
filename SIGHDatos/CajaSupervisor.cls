VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CajaSupervisor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'WCG Ini 02/06/2005 [Todo el archivo]

Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: CajaSupervisor
'        Autor: William Castro Grijalva
'        Fecha: 03/06/2005 10:24:53
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
'   Descripción:    Inserta un registro a la tabla CajaSupervisor
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOCajaSupervisor) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaSupervisorAgregar"
           Set oParameter = .CreateParameter("@IdEmpleado", adInteger, adParamInput, 0, IIf(oTabla.IdEmpleado = 0, Null, oTabla.IdEmpleado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@EstadoSupervisor", adChar, adParamInput, 1, IIf(oTabla.EstadoSupervisor = "", Null, oTabla.EstadoSupervisor)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdSupervisor", adInteger, adParamOutput, 0): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdSupervisor = .Parameters("@IdSupervisor")
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
'   Descripción:    modifica un registro a la tabla CajaSupervisor
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOCajaSupervisor) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaSupervisorModificar"
           Set oParameter = .CreateParameter("@IdEmpleado", adInteger, adParamInput, 0, IIf(oTabla.IdEmpleado = 0, Null, oTabla.IdEmpleado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@EstadoSupervisor", adChar, adParamInput, 1, IIf(oTabla.EstadoSupervisor = "", "0", oTabla.EstadoSupervisor)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdSupervisor", adInteger, adParamInput, 0, IIf(oTabla.IdSupervisor = 0, Null, oTabla.IdSupervisor)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla CajaSupervisor
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOCajaSupervisor) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaSupervisorEliminar"
           Set oParameter = .CreateParameter("@IdSupervisor", adInteger, adParamInput, 0, IIf(oTabla.IdSupervisor = 0, Null, oTabla.IdSupervisor)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla CajaSupervisor
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOCajaSupervisor) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdText
     Set .ActiveConnection = mo_Conexion
     .CommandText = "Select * from CajaSupervisor where IdSupervisor = " & oTabla.IdSupervisor
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.IdEmpleado = IIf(IsNull(oRecordset!IdEmpleado), 0, oRecordset!IdEmpleado)
       oTabla.EstadoSupervisor = IIf(IsNull(oRecordset!EstadoSupervisor), "", oRecordset!EstadoSupervisor)
       oTabla.IdSupervisor = IIf(IsNull(oRecordset!IdSupervisor), 0, oRecordset!IdSupervisor)
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

Function RealizarFiltro(oEmpleado As dOEmpleado) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim SQL As String

   SQL = "Select s.IdSupervisor,case s.EstadoSupervisor when '1' then 'Si' else 'No' end as EstadoSupervisor ,  s.IdEmpleado, e.ApellidoPaterno,e.ApellidoMaterno,e.Nombres from CajaSupervisor s, Empleados e where s.IdEmpleado = e.IdEmpleado"
   If oEmpleado.ApellidoPaterno <> "" Then
        SQL = SQL + " and e.ApellidoPaterno like '%" & oEmpleado.ApellidoPaterno & "%'"
   End If
   If oEmpleado.ApellidoMaterno <> "" Then
        SQL = SQL + " and e.ApellidoMaterno like '%" & oEmpleado.ApellidoMaterno & "%'"
   End If
   If oEmpleado.Nombres <> "" Then
        SQL = SQL + " and e.Nombres like '%" & oEmpleado.Nombres & "%'"
   End If
 
 
   Set RealizarFiltro = Nothing
   
   With oCommand
     .CommandType = adCmdText
     Set .ActiveConnection = mo_Conexion
     .CommandText = SQL
     Set oRecordset = .Execute
   End With
   
   Set oRecordset.ActiveConnection = Nothing
   Set RealizarFiltro = oRecordset
 
   ms_MensajeError = ""
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function SeleccionarTodos() As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim SQL As String

   SQL = "Select s.IdSupervisor,isnull(e.ApellidoPaterno,'') + ' ' +  isnull(e.ApellidoMaterno,'') + ' ' + isnull(e.Nombres,'') as NombreCompletoSupervisor from CajaCajeroSupervisor s, Empleados e where s.IdEmpleado = e.IdEmpleado"
  
   Set SeleccionarTodos = Nothing
   
   With oCommand
     .CommandType = adCmdText
     Set .ActiveConnection = mo_Conexion
     .CommandText = SQL
     Set oRecordset = .Execute
   End With
   
   Set oRecordset.ActiveConnection = Nothing
   Set SeleccionarTodos = oRecordset
 
   ms_MensajeError = ""
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

