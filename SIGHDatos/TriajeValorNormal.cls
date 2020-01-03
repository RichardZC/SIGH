VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TriajeValorNormal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla TriajeValorNormal
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
'   Descripción:    Inserta un registro a la tabla TriajeValorNormal
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOTriajeValorNormal) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "TriajeValorNormalAgregar"
   Set oParameter = .CreateParameter("@IdTriajeValorNormal", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EdadInicialEnDia", adInteger, adParamInput, 0, IIf(oTabla.EdadInicialEnDia = 0, Null, oTabla.EdadInicialEnDia))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EdadFinalEnDia", adInteger, adParamInput, 0, IIf(oTabla.EdadFinalEnDia = 0, Null, oTabla.EdadFinalEnDia))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ValorNormalMinimo", adCurrency, adParamInput, 0, oTabla.ValorNormalMinimo)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ValorNormalMaximo", adCurrency, adParamInput, 0, oTabla.ValorNormalMaximo)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ValorCoherenteMinimo", adCurrency, adParamInput, 0, oTabla.ValorCoherenteMinimo)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ValorCoherenteMaximo", adCurrency, adParamInput, 0, oTabla.ValorCoherenteMaximo)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdTriajeVariable", adInteger, adParamInput, 0, IIf(oTabla.IdTriajeVariable = 0, Null, oTabla.IdTriajeVariable))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EstadoPaciente", adInteger, adParamInput, 0, IIf(oTabla.EstadoPaciente = 0, Null, oTabla.EstadoPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SexoPaciente", adInteger, adParamInput, 0, IIf(oTabla.SexoPaciente = 0, Null, oTabla.SexoPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaVigencia", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaVigencia = 0, Null, oTabla.FechaVigencia))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.IdTriajeValorNormal = .Parameters("@IdTriajeValorNormal")
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
'   Descripción:    modifica un registro a la tabla TriajeValorNormal
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOTriajeValorNormal) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "TriajeValorNormalModificar"
   Set oParameter = .CreateParameter("@IdTriajeValorNormal", adInteger, adParamInput, 0, IIf(oTabla.IdTriajeValorNormal = 0, Null, oTabla.IdTriajeValorNormal))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EdadInicialEnDia", adInteger, adParamInput, 0, IIf(oTabla.EdadInicialEnDia = 0, Null, oTabla.EdadInicialEnDia))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EdadFinalEnDia", adInteger, adParamInput, 0, IIf(oTabla.EdadFinalEnDia = 0, Null, oTabla.EdadFinalEnDia))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ValorNormalMinimo", adCurrency, adParamInput, 0, oTabla.ValorNormalMinimo)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ValorNormalMaximo", adCurrency, adParamInput, 0, oTabla.ValorNormalMaximo)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ValorCoherenteMinimo", adCurrency, adParamInput, 0, oTabla.ValorCoherenteMinimo)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ValorCoherenteMaximo", adCurrency, adParamInput, 0, oTabla.ValorCoherenteMaximo)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdTriajeVariable", adInteger, adParamInput, 0, IIf(oTabla.IdTriajeVariable = 0, Null, oTabla.IdTriajeVariable))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EstadoPaciente", adInteger, adParamInput, 0, IIf(oTabla.EstadoPaciente = 0, Null, oTabla.EstadoPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SexoPaciente", adInteger, adParamInput, 0, IIf(oTabla.SexoPaciente = 0, Null, oTabla.SexoPaciente))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaVigencia", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaVigencia = 0, Null, oTabla.FechaVigencia))
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
'   Descripción:    elimia o borra un registro a la tabla TriajeValorNormal
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOTriajeValorNormal) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "TriajeValorNormalEliminar"
   Set oParameter = .CreateParameter("@IdTriajeValorNormal", adInteger, adParamInput, 0, oTabla.IdTriajeValorNormal)
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
'   Descripción:    Selecciona un unico registro de la tabla TriajeValorNormal
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOTriajeValorNormal) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "TriajeValorNormalSeleccionarPorId"
     Set oParameter = .CreateParameter("@IdTriajeValorNormal", adInteger, adParamInput, 0, oTabla.IdTriajeValorNormal)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdTriajeValorNormal = IIf(IsNull(oRecordset!IdTriajeValorNormal), 0, oRecordset!IdTriajeValorNormal)
       oTabla.EdadInicialEnDia = IIf(IsNull(oRecordset!EdadInicialEnDia), 0, oRecordset!EdadInicialEnDia)
       oTabla.EdadFinalEnDia = IIf(IsNull(oRecordset!EdadFinalEnDia), 0, oRecordset!EdadFinalEnDia)
       oTabla.ValorNormalMinimo = oRecordset!ValorNormalMinimo
       oTabla.ValorNormalMaximo = oRecordset!ValorNormalMaximo
       oTabla.ValorCoherenteMinimo = oRecordset!ValorCoherenteMinimo
       oTabla.ValorCoherenteMaximo = oRecordset!ValorCoherenteMaximo
       oTabla.IdTriajeVariable = IIf(IsNull(oRecordset!IdTriajeVariable), 0, oRecordset!IdTriajeVariable)
       oTabla.EstadoPaciente = IIf(IsNull(oRecordset!EstadoPaciente), 0, oRecordset!EstadoPaciente)
       oTabla.SexoPaciente = IIf(IsNull(oRecordset!SexoPaciente), 0, oRecordset!SexoPaciente)
       oTabla.FechaVigencia = IIf(IsNull(oRecordset!FechaVigencia), 0, oRecordset!FechaVigencia)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOTriajeValorNormal)
   oTabla.IdTriajeValorNormal = 0
   oTabla.EdadInicialEnDia = 0
   oTabla.EdadFinalEnDia = 0
   oTabla.ValorNormalMinimo = 0
   oTabla.ValorNormalMaximo = 0
   oTabla.ValorCoherenteMinimo = 0
   oTabla.ValorCoherenteMaximo = 0
   oTabla.IdTriajeVariable = 0
   oTabla.EstadoPaciente = 0
   oTabla.SexoPaciente = 0
   oTabla.FechaVigencia = Now
End Sub

Function ListarValorNormalesSegunParametros(oTabla As DOTriajeValorNormal) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Set ListarValorNormalesSegunParametros = Nothing
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "TriajeValorNormalesSegunParamtros"
     
     Set oParameter = .CreateParameter("EdadInicialEnDia", adInteger, adParamInput, 0, oTabla.EdadInicialEnDia)
     .Parameters.Append oParameter
     Set oParameter = .CreateParameter("SexoPaciente", adInteger, adParamInput, 0, oTabla.SexoPaciente)
     .Parameters.Append oParameter
     Set oParameter = .CreateParameter("FechaVigencia", adDBTimeStamp, adParamInput, 0, oTabla.FechaVigencia)
     .Parameters.Append oParameter
     Set oParameter = .CreateParameter("EstadoPaciente", adInteger, adParamInput, 0, oTabla.EstadoPaciente)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
     Set oRecordset.ActiveConnection = Nothing
   End With
   
   Set ListarValorNormalesSegunParametros = oRecordset
 
   ms_MensajeError = ""
   Set oCommand = Nothing
   Set oRecordset = Nothing
    Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
End Function
