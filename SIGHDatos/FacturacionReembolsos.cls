VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FacturacionReembolsos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla FacturacionReembolsos
'        Programado por: Barrantes D
'        Fecha: Noviembre 2009
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
'   Descripción:    Inserta un registro a la tabla FacturacionReembolsos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOFacturacionReembolsos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "FacturacionReembolsosAgregar"
   Set oParameter = .CreateParameter("@idFactReembolso", adInteger, adParamInput, 0, IIf(oTabla.idFactReembolso = 0, Null, oTabla.idFactReembolso))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ConsumoPorReembolsar", adCurrency, adParamInput, 0, oTabla.ConsumoPorReembolsar)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ReembolsoPorPagar", adCurrency, adParamInput, 0, oTabla.ReembolsoPorPagar)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ReembolsoPagadoFarmacia", adCurrency, adParamInput, 0, oTabla.ReembolsoPagadoFarmacia)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ReembolsoPagadoServicio", adCurrency, adParamInput, 0, oTabla.ReembolsoPagadoServicio)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdReembolsosAnteriores", adVarChar, adParamInput, 100, IIf(oTabla.IdReembolsosAnteriores = "", Null, oTabla.IdReembolsosAnteriores))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.IdDiagnostico = 0, Null, oTabla.IdDiagnostico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroReferenciaDestino", adVarChar, adParamInput, 20, IIf(oTabla.NroReferenciaDestino = "", Null, oTabla.NroReferenciaDestino))
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
'   Descripción:    modifica un registro a la tabla FacturacionReembolsos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOFacturacionReembolsos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "FacturacionReembolsosModificar"
   Set oParameter = .CreateParameter("@idFactReembolso", adInteger, adParamInput, 0, IIf(oTabla.idFactReembolso = 0, Null, oTabla.idFactReembolso))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idCuentaAtencion", adInteger, adParamInput, 0, IIf(oTabla.IdCuentaAtencion = 0, Null, oTabla.IdCuentaAtencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ConsumoPorReembolsar", adCurrency, adParamInput, 0, oTabla.ConsumoPorReembolsar)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ReembolsoPorPagar", adCurrency, adParamInput, 0, oTabla.ReembolsoPorPagar)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ReembolsoPagadoFarmacia", adCurrency, adParamInput, 0, oTabla.ReembolsoPagadoFarmacia)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ReembolsoPagadoServicio", adCurrency, adParamInput, 0, oTabla.ReembolsoPagadoServicio)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdReembolsosAnteriores", adVarChar, adParamInput, 100, IIf(oTabla.IdReembolsosAnteriores = "", Null, oTabla.IdReembolsosAnteriores))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idDiagnostico", adInteger, adParamInput, 0, IIf(oTabla.IdDiagnostico = 0, Null, oTabla.IdDiagnostico))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NroReferenciaDestino", adVarChar, adParamInput, 20, IIf(oTabla.NroReferenciaDestino = "", Null, oTabla.NroReferenciaDestino))
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
'   Descripción:    elimia o borra un registro a la tabla FacturacionReembolsos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOFacturacionReembolsos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "FacturacionReembolsosEliminar"
   Set oParameter = .CreateParameter("@idFactReembolso", adInteger, adParamInput, 0, oTabla.idFactReembolso)
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
'   Descripción:    Selecciona un unico registro de la tabla FacturacionReembolsos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOFacturacionReembolsos) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "FacturacionReembolsosSeleccionarPorId"
     Set oParameter = .CreateParameter("@idFactReembolso", adInteger, adParamInput, 0, oTabla.idFactReembolso)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.idFactReembolso = IIf(IsNull(oRecordset!idFactReembolso), 0, oRecordset!idFactReembolso)
       oTabla.IdCuentaAtencion = IIf(IsNull(oRecordset!IdCuentaAtencion), 0, oRecordset!IdCuentaAtencion)
       oTabla.ConsumoPorReembolsar = oRecordset!ConsumoPorReembolsar
       oTabla.ReembolsoPorPagar = oRecordset!ReembolsoPorPagar
       oTabla.ReembolsoPagadoFarmacia = oRecordset!ReembolsoPagadoFarmacia
       oTabla.ReembolsoPagadoServicio = oRecordset!ReembolsoPagadoServicio
       oTabla.IdReembolsosAnteriores = IIf(IsNull(oRecordset!IdReembolsosAnteriores), "", oRecordset!IdReembolsosAnteriores)
       oTabla.IdDiagnostico = IIf(IsNull(oRecordset!IdDiagnostico), 0, oRecordset!IdDiagnostico)
       oTabla.NroReferenciaDestino = IIf(IsNull(oRecordset!NroReferenciaDestino), "", oRecordset!NroReferenciaDestino)
   SeleccionarPorId = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOFacturacionReembolsos)
   oTabla.idFactReembolso = 0
   oTabla.IdCuentaAtencion = 0
   oTabla.ConsumoPorReembolsar = 0
   oTabla.ReembolsoPorPagar = 0
   oTabla.ReembolsoPagadoFarmacia = 0
   oTabla.ReembolsoPagadoServicio = 0
   oTabla.IdReembolsosAnteriores = ""
   oTabla.IdDiagnostico = 0
   oTabla.NroReferenciaDestino = ""
End Sub

