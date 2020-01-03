VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CajaLoteCuadre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'WCG Ini 02/06/2005 [Todo el archivo]

Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: CajaLoteCuadre
'        Autor: William Castro Grijalva
'        Fecha: 03/06/2005 10:24:49
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
'   Descripción:    Inserta un registro a la tabla CajaLoteCuadre
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOCajaLoteCuadre) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaLoteCuadreAgregar"
           Set oParameter = .CreateParameter("@IdLote", adInteger, adParamInput, 0, IIf(oTabla.IdLote = 0, Null, oTabla.IdLote)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Diferencia", adCurrency, adParamInput, 0, IIf(oTabla.Diferencia = 0, Null, oTabla.Diferencia)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@CuadreUsuario", adCurrency, adParamInput, 0, IIf(oTabla.CuadreUsuario = 0, Null, oTabla.CuadreUsuario)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Calculado", adCurrency, adParamInput, 0, IIf(oTabla.Calculado = 0, Null, oTabla.Calculado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdLoteCuadre", adInteger, adParamOutput, 0): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoFormaPago", adInteger, adParamInput, 0, IIf(oTabla.IdTipoFormaPago = 0, Null, oTabla.IdTipoFormaPago)): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdLoteCuadre = .Parameters("@IdLoteCuadre")
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
'   Descripción:    modifica un registro a la tabla CajaLoteCuadre
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOCajaLoteCuadre) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaLoteCuadreModificar"
           Set oParameter = .CreateParameter("@IdLote", adInteger, adParamInput, 0, IIf(oTabla.IdLote = 0, Null, oTabla.IdLote)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Diferencia", adCurrency, adParamInput, 0, IIf(oTabla.Diferencia = 0, Null, oTabla.Diferencia)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@CuadreUsuario", adCurrency, adParamInput, 0, IIf(oTabla.CuadreUsuario = 0, Null, oTabla.CuadreUsuario)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Calculado", adCurrency, adParamInput, 0, IIf(oTabla.Calculado = 0, Null, oTabla.Calculado)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdLoteCuadre", adInteger, adParamInput, 0, IIf(oTabla.IdLoteCuadre = 0, Null, oTabla.IdLoteCuadre)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoFormaPago", adInteger, adParamInput, 0, IIf(oTabla.IdTipoFormaPago = 0, Null, oTabla.IdTipoFormaPago)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla CajaLoteCuadre
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOCajaLoteCuadre) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CajaLoteCuadreEliminar"
           Set oParameter = .CreateParameter("@IdLoteCuadre", adInteger, adParamInput, 0, IIf(oTabla.IdLoteCuadre = 0, Null, oTabla.IdLoteCuadre)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla CajaLoteCuadre
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOCajaLoteCuadre) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdText
     Set .ActiveConnection = mo_Conexion
     .CommandText = "Select * from CajaLoteCuadre where IdLoteCuadre =" & oTabla.IdLoteCuadre
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.IdLote = IIf(IsNull(oRecordset!IdLote), 0, oRecordset!IdLote)
       oTabla.Diferencia = IIf(IsNull(oRecordset!Diferencia), 0, oRecordset!Diferencia)
       oTabla.CuadreUsuario = IIf(IsNull(oRecordset!CuadreUsuario), 0, oRecordset!CuadreUsuario)
       oTabla.Calculado = IIf(IsNull(oRecordset!Calculado), 0, oRecordset!Calculado)
       oTabla.IdLoteCuadre = IIf(IsNull(oRecordset!IdLoteCuadre), 0, oRecordset!IdLoteCuadre)
       oTabla.IdTipoFormaPago = IIf(IsNull(oRecordset!IdTipoFormaPago), 0, oRecordset!IdTipoFormaPago)
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

Function RealizarFiltro(oLote As DOCajaLote) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim SQL As String

   SQL = "Select lc.IdLoteCuadre,ca.Descripcion NombreCaja" _
                & " ,e.ApellidoPaterno + ' '+ e.ApellidoMaterno + ' '+ e.Nombres NombreCajero " _
                & " ,l.Fecha,lc.Calculado,lc.CuadreUsuario,lc.Diferencia " _
                & " from CajaLoteCuadre lc, CajaLote l, CajaCaja ca,CajaCajero cj, Empleados e " _
                & " Where lc.IdLote = l.IdLote " _
                & " and l.IdCaja = ca.IdCaja " _
                & " and l.IdCajero = cj.IdCajero " _
                & " and cj.IdEmpleado = e.IdEmpleado"
            
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


