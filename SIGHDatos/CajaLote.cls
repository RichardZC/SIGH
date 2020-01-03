VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CajaLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'WCG Ini 01/06/2005 [Todo el archivo]
'WCG02 Ini 04/07/2005
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: CajaLote
'        Autor: William Castro Grijalva
'        Fecha: 02/06/2005 11:56:29
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
'   Descripción:    Inserta un registro a la tabla CajaLote
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOCajaLote) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaLoteAgregar"
           Set oParameter = .CreateParameter("@IdCajero", adInteger, adParamInput, 0, IIf(oTabla.IdCajero = 0, Null, oTabla.IdCajero)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@SaldoInicialDolares", adCurrency, adParamInput, 0, IIf(oTabla.SaldoInicialDolares = 0, Null, oTabla.SaldoInicialDolares)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@SaldoInicialSoles", adCurrency, adParamInput, 0, IIf(oTabla.SaldoInicialSoles = 0, Null, oTabla.SaldoInicialSoles)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@EstadoLote", adChar, adParamInput, 1, IIf(oTabla.EstadoLote = "", Null, oTabla.EstadoLote)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, IIf(oTabla.Fecha = 0, Null, oTabla.Fecha)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdLote", adInteger, adParamOutput, 0): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCaja", adInteger, adParamInput, 0, IIf(oTabla.IdCaja = 0, Null, oTabla.IdCaja)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTurno", adInteger, adParamInput, 0, IIf(oTabla.IdTurno = 0, Null, oTabla.IdTurno)): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdLote = .Parameters("@IdLote")
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
'   Descripción:    modifica un registro a la tabla CajaLote
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOCajaLote) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaLoteModificar"
           Set oParameter = .CreateParameter("@IdCajero", adInteger, adParamInput, 0, IIf(oTabla.IdCajero = 0, Null, oTabla.IdCajero)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@SaldoInicialDolares", adCurrency, adParamInput, 0, IIf(oTabla.SaldoInicialDolares = 0, Null, oTabla.SaldoInicialDolares)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@SaldoInicialSoles", adCurrency, adParamInput, 0, IIf(oTabla.SaldoInicialSoles = 0, Null, oTabla.SaldoInicialSoles)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@EstadoLote", adChar, adParamInput, 1, IIf(oTabla.EstadoLote = "", Null, oTabla.EstadoLote)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Fecha", adDBTimeStamp, adParamInput, 0, IIf(oTabla.Fecha = 0, Null, oTabla.Fecha)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdLote", adInteger, adParamInput, 0, IIf(oTabla.IdLote = 0, Null, oTabla.IdLote)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCaja", adInteger, adParamInput, 0, IIf(oTabla.IdCaja = 0, Null, oTabla.IdCaja)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTurno", adInteger, adParamInput, 0, IIf(oTabla.IdTurno = 0, Null, oTabla.IdTurno)): .Parameters.Append oParameter
           
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
'   Descripción:    elimia o borra un registro a la tabla CajaLote
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOCajaLote) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaLoteEliminar"
           Set oParameter = .CreateParameter("@IdLote", adInteger, adParamInput, 0, IIf(oTabla.IdLote = 0, Null, oTabla.IdLote)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla CajaLote
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOCajaLote) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdText
     Set .ActiveConnection = mo_Conexion
     .CommandText = "Select * from CajaLote where IdLote = " & oTabla.IdLote
       Set oParameter = .CreateParameter("@IdLote", adInteger, adParamInput, 0, oTabla.IdLote): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.IdCajero = IIf(IsNull(oRecordset!IdCajero), 0, oRecordset!IdCajero)
       oTabla.SaldoInicialDolares = IIf(IsNull(oRecordset!SaldoInicialDolares), 0, oRecordset!SaldoInicialDolares)
       oTabla.SaldoInicialSoles = IIf(IsNull(oRecordset!SaldoInicialSoles), 0, oRecordset!SaldoInicialSoles)
       oTabla.EstadoLote = IIf(IsNull(oRecordset!EstadoLote), "", oRecordset!EstadoLote)
       oTabla.Fecha = IIf(IsNull(oRecordset!Fecha), 0, oRecordset!Fecha)
       oTabla.IdLote = IIf(IsNull(oRecordset!IdLote), 0, oRecordset!IdLote)
       oTabla.IdCaja = IIf(IsNull(oRecordset!IdCaja), 0, oRecordset!IdCaja)
       oTabla.IdTurno = IIf(IsNull(oRecordset!IdTurno), 0, oRecordset!IdTurno)
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

'------------------------------------------------------------------------------------
'   Función:        SeleccionarPorId
'   Descripción:    Selecciona un unico registro de la tabla CajaLote
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorFechaCajero(oTabla As DOCajaLote) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSQL As String

    sSQL = "Select * from CajaLote where idCajero = " & oTabla.IdCajero _
            & " and Fecha = '" & oTabla.Fecha & "'" _
            & " and IdTurno = " & oTabla.IdTurno
   SeleccionarPorFechaCajero = False
   With oCommand
     .CommandType = adCmdText
     Set .ActiveConnection = mo_Conexion
     .CommandText = sSQL
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorFechaCajero = True
       oTabla.IdLote = IIf(IsNull(oRecordset!IdLote), 0, oRecordset!IdLote)
       oTabla.Fecha = IIf(IsNull(oRecordset!Fecha), "", oRecordset!Fecha)
       oTabla.EstadoLote = IIf(IsNull(oRecordset!EstadoLote), "0", oRecordset!EstadoLote)
       oTabla.SaldoInicialDolares = IIf(IsNull(oRecordset!SaldoInicialDolares), 0, oRecordset!SaldoInicialDolares)
       oTabla.SaldoInicialSoles = IIf(IsNull(oRecordset!SaldoInicialSoles), 0, oRecordset!SaldoInicialSoles)
       oTabla.IdCajero = IIf(IsNull(oRecordset!IdCajero), 0, oRecordset!IdCajero)
       oTabla.IdCaja = IIf(IsNull(oRecordset!IdCaja), 0, oRecordset!IdCaja)
       oTabla.IdTurno = IIf(IsNull(oRecordset!IdTurno), 0, oRecordset!IdTurno)
   Else
       SeleccionarPorFechaCajero = False
   End If

   oRecordset.Close
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function RealizarFiltro(oLote As DOCajaLote) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim SQL As String
    
'WCG02 Ini 04/07/2005
   SQL = "Select l.IdLote,ca.Descripcion NombreCaja" _
            & " ,e.ApellidoPaterno + ' '+ e.ApellidoMaterno + ' '+ e.Nombres NombreCajero " _
            & " ,ct.Descripcion Turno " _
            & " ,l.Fecha,l.EstadoLote,l.SaldoInicialSoles,l.SaldoInicialDolares " _
            & " from CajaLote l, CajaCaja ca,CajaCajero cj, Empleados e, CajaTurno ct " _
            & " Where l.IdCaja = ca.IdCaja " _
            & " and l.IdCajero = cj.IdCajero " _
            & " and cj.IdEmpleado = e.IdEmpleado" _
            & " and l.IdTurno = ct.IdTurno"
'WCG02 Fin 04/07/2005
            
   If oLote.Fecha <> 0 Then
        SQL = SQL + " and l.Fecha = '" & oLote.Fecha & "'"
   End If
   If oLote.IdCaja <> 0 Then
        SQL = SQL + " and l.IdCaja = " & oLote.IdCaja
   End If
   If oLote.IdCajero <> 0 Then
        SQL = SQL + " and l.IdCajero = " & oLote.IdCajero
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

Function SeleccionarPendientesParaLista(IdLoteDefault As Integer) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim SQL As String

   SQL = "Select l.IdLote" _
            & " ,Convert(varchar(10),l.Fecha,103) + ': ' + ca.Descripcion + ' - '+ e.ApellidoPaterno + ' '+ e.ApellidoMaterno + ' '+ e.Nombres Descripcion " _
            & " from CajaLote l, CajaCaja ca,CajaCajero cj, Empleados e " _
            & " Where l.IdCaja = ca.IdCaja " _
            & " and l.IdCajero = cj.IdCajero " _
            & " and cj.IdEmpleado = e.IdEmpleado" _
            & " and (l.EstadoLote = 'A' or l.IdLote = " & IdLoteDefault & ")"
 
   Set SeleccionarPendientesParaLista = Nothing
   
   With oCommand
        .CommandType = adCmdText
        Set .ActiveConnection = mo_Conexion
        .CommandText = SQL
        Set oRecordset = .Execute
   End With
   
   Set oRecordset.ActiveConnection = Nothing
   Set SeleccionarPendientesParaLista = oRecordset
 
   ms_MensajeError = ""
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function ExisteAsignacionCaja(oTabla As DOCajaLote) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSQL As String

    sSQL = "Select * from CajaLote where idCajero = " & oTabla.IdCajero _
            & " and Fecha = '" & oTabla.Fecha & "'" _
            & " and IdCaja = " & oTabla.IdCaja _
            & " and IdTurno = " & oTabla.IdTurno _
            & " and IdLote <> " & oTabla.IdLote
   ExisteAsignacionCaja = False
   With oCommand
     .CommandType = adCmdText
     Set .ActiveConnection = mo_Conexion
     .CommandText = sSQL
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       ExisteAsignacionCaja = True
   Else
       ExisteAsignacionCaja = False
   End If

   oRecordset.Close
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function

End Function

Function ObtenerMontoCalculado(IdLote As Integer) As Double
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSQL As String

   sSQL = "Select isnull(sum(Total),0.0) as Total from CajaComprobantesPago where IdLote = " & IdLote
   ObtenerMontoCalculado = 0
   With oCommand
     .CommandType = adCmdText
     Set .ActiveConnection = mo_Conexion
     .CommandText = sSQL
     Set oRecordset = .Execute
   End With
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       ObtenerMontoCalculado = oRecordset.Fields("Total")
   Else
       ObtenerMontoCalculado = 0
   End If
   oRecordset.Close
   ms_MensajeError = ""
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


