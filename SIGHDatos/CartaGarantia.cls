VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CartaGarantia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------------------
'        Inicio de código autogenerado para la clase: CartaGarantia
'        Autor: William Castro Grijalva
'        Fecha: 11/09/2005 01:20:47
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
'   Descripción:    Inserta un registro a la tabla CartaGarantia
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOCartaGarantia) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CartaGarantiaAgregar"
           Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaVigencia", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaVigencia = 0, Null, oTabla.FechaVigencia)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NroCarta", adChar, adParamInput, 10, IIf(oTabla.NroCarta = "", Null, oTabla.NroCarta)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Observacion", adVarChar, adParamInput, 100, IIf(oTabla.Observacion = "", Null, oTabla.Observacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@ValorCobertura", adCurrency, adParamInput, 0, IIf(oTabla.ValorCobertura = 0, Null, oTabla.ValorCobertura)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCartaGarantia", adInteger, adParamOutput, 0): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdCartaGarantia = .Parameters("@IdCartaGarantia")
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
'   Descripción:    modifica un registro a la tabla CartaGarantia
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOCartaGarantia) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CartaGarantiaModificar"
           Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FechaVigencia", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaVigencia = 0, Null, oTabla.FechaVigencia)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NroCarta", adChar, adParamInput, 10, IIf(oTabla.NroCarta = "", Null, oTabla.NroCarta)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Observacion", adVarChar, adParamInput, 100, IIf(oTabla.Observacion = "", Null, oTabla.Observacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@ValorCobertura", adCurrency, adParamInput, 0, IIf(oTabla.ValorCobertura = 0, Null, oTabla.ValorCobertura)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCartaGarantia", adInteger, adParamInput, 0, IIf(oTabla.IdCartaGarantia = 0, Null, oTabla.IdCartaGarantia)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla CartaGarantia
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOCartaGarantia) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "CartaGarantiaEliminar"
           Set oParameter = .CreateParameter("@IdCartaGarantia", adInteger, adParamInput, 0, IIf(oTabla.IdCartaGarantia = 0, Null, oTabla.IdCartaGarantia)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla CartaGarantia
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOCartaGarantia) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "CartaGarantiaSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdCartaGarantia", adInteger, adParamInput, 0, oTabla.IdCartaGarantia): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.IdCuentaAtencion = IIf(IsNull(oRecordset!IdCuentaAtencion), 0, oRecordset!IdCuentaAtencion)
       oTabla.FechaVigencia = IIf(IsNull(oRecordset!FechaVigencia), 0, oRecordset!FechaVigencia)
       oTabla.NroCarta = IIf(IsNull(oRecordset!NroCarta), "", oRecordset!NroCarta)
       oTabla.Observacion = IIf(IsNull(oRecordset!Observacion), "", oRecordset!Observacion)
       oTabla.ValorCobertura = IIf(IsNull(oRecordset!ValorCobertura), 0, oRecordset!ValorCobertura)
       oTabla.IdCartaGarantia = IIf(IsNull(oRecordset!IdCartaGarantia), 0, oRecordset!IdCartaGarantia)
   Else
       SeleccionarPorId = False
   End If

   oRecordset.Close
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function


Function Filtrar(ByVal oTabla As DOCartaGarantia) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSQL As String
Dim sWhere As String

    Set Filtrar = Nothing
     
    sWhere = ""
        
    sSQL = "Select * from CartaGarantia where 1 = 1 "
    If oTabla.IdCuentaAtencion <> 0 Then
        sSQL = sSQL + " and IdCuentaAtencion= " & oTabla.IdCuentaAtencion
    End If
    If oTabla.FechaVigencia <> 0 Then
        sSQL = sSQL + " and FechaVigencia >= '" & oTabla.FechaVigencia & "'"
    End If
    If oTabla.NroCarta <> "" Then
        sSQL = sSQL + " and NroCarta like '%" & oTabla.NroCarta & "'"
    End If
    sSQL = sSQL + " order by FechaVigencia desc"
        
    With oCommand
       .CommandType = adCmdText
       Set .ActiveConnection = mo_Conexion
       .CommandText = sSQL
       Set oRecordset = .Execute
       Set oRecordset.ActiveConnection = Nothing
    End With
    
    Set Filtrar = oRecordset
    
    ms_MensajeError = ""
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


