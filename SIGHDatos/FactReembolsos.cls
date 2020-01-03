VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FactReembolsos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla FactReembolsos
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
'   Descripción:    Inserta un registro a la tabla FactReembolsos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOFactReembolsos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "FactReembolsosAgregar"
   Set oParameter = .CreateParameter("@idFactReembolso", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Anio", adInteger, adParamInput, 0, IIf(oTabla.Anio = 0, Null, oTabla.Anio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Mes", adInteger, adParamInput, 0, IIf(oTabla.Mes = 0, Null, oTabla.Mes))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idAreaTramitaSeguro", adInteger, adParamInput, 0, IIf(oTabla.idAreaTramitaSeguro = 0, Null, oTabla.idAreaTramitaSeguro))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idFuenteFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.idFuenteFinanciamiento = 0, Null, oTabla.idFuenteFinanciamiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Descripcion", adVarChar, adParamInput, 300, IIf(oTabla.Descripcion = "", Null, oTabla.Descripcion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SaldoInicial", adCurrency, adParamInput, 0, oTabla.SaldoInicial)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ConsumoPorReembolsar", adCurrency, adParamInput, 0, oTabla.ConsumoPorReembolsar)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ReembolsoPagado", adCurrency, adParamInput, 0, oTabla.ReembolsoPagado)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ReembolsoPorPagar", adCurrency, adParamInput, 0, oTabla.ReembolsoPorPagar)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SaldoFinal", adCurrency, adParamInput, 0, oTabla.SaldoFinal)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Documentos", adVarChar, adParamInput, 100, IIf(oTabla.Documentos = "", Null, oTabla.Documentos))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idEstadoReembolso", adInteger, adParamInput, 0, IIf(oTabla.idEstadoReembolso = 0, Null, oTabla.idEstadoReembolso))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTipoConsumo", adInteger, adParamInput, 0, oTabla.idTipoConsumo)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdTipoComprobante", adInteger, adParamInput, 0, IIf(oTabla.IdTipoComprobante = 0, Null, oTabla.IdTipoComprobante))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@GrabaDefinitivamente", adBoolean, adParamInput, 0, oTabla.GrabaDefinitivamente = True)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.idFactReembolso = .Parameters("@idFactReembolso")
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
'   Descripción:    modifica un registro a la tabla FactReembolsos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOFactReembolsos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "FactReembolsosModificar"
   Set oParameter = .CreateParameter("@idFactReembolso", adInteger, adParamInput, 0, IIf(oTabla.idFactReembolso = 0, Null, oTabla.idFactReembolso))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Anio", adInteger, adParamInput, 0, IIf(oTabla.Anio = 0, Null, oTabla.Anio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Mes", adInteger, adParamInput, 0, IIf(oTabla.Mes = 0, Null, oTabla.Mes))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idAreaTramitaSeguro", adInteger, adParamInput, 0, IIf(oTabla.idAreaTramitaSeguro = 0, Null, oTabla.idAreaTramitaSeguro))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idFuenteFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.idFuenteFinanciamiento = 0, Null, oTabla.idFuenteFinanciamiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Descripcion", adVarChar, adParamInput, 300, IIf(oTabla.Descripcion = "", Null, oTabla.Descripcion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SaldoInicial", adCurrency, adParamInput, 0, oTabla.SaldoInicial)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ConsumoPorReembolsar", adCurrency, adParamInput, 0, oTabla.ConsumoPorReembolsar)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ReembolsoPagado", adCurrency, adParamInput, 0, oTabla.ReembolsoPagado)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@ReembolsoPorPagar", adCurrency, adParamInput, 0, oTabla.ReembolsoPorPagar)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@SaldoFinal", adCurrency, adParamInput, 0, oTabla.SaldoFinal)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Documentos", adVarChar, adParamInput, 100, IIf(oTabla.Documentos = "", Null, oTabla.Documentos))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idEstadoReembolso", adInteger, adParamInput, 0, oTabla.idEstadoReembolso)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTipoConsumo", adInteger, adParamInput, 0, oTabla.idTipoConsumo)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdTipoComprobante", adInteger, adParamInput, 0, IIf(oTabla.IdTipoComprobante = 0, Null, oTabla.IdTipoComprobante))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@GrabaDefinitivamente", adBoolean, adParamInput, 0, oTabla.GrabaDefinitivamente = True)
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
'   Descripción:    elimia o borra un registro a la tabla FactReembolsos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOFactReembolsos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "FactReembolsosEliminar"
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
'   Descripción:    Selecciona un unico registro de la tabla FactReembolsos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOFactReembolsos) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "FactReembolsosSeleccionarPorId"
     Set oParameter = .CreateParameter("@idFactReembolso", adInteger, adParamInput, 0, oTabla.idFactReembolso)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.idFactReembolso = IIf(IsNull(oRecordset!idFactReembolso), 0, oRecordset!idFactReembolso)
       oTabla.Anio = IIf(IsNull(oRecordset!Anio), 0, oRecordset!Anio)
       oTabla.Mes = IIf(IsNull(oRecordset!Mes), 0, oRecordset!Mes)
       oTabla.idAreaTramitaSeguro = IIf(IsNull(oRecordset!idAreaTramitaSeguro), 0, oRecordset!idAreaTramitaSeguro)
       oTabla.idFuenteFinanciamiento = IIf(IsNull(oRecordset!idFuenteFinanciamiento), 0, oRecordset!idFuenteFinanciamiento)
       oTabla.Descripcion = IIf(IsNull(oRecordset!Descripcion), "", oRecordset!Descripcion)
       oTabla.SaldoInicial = oRecordset!SaldoInicial
       oTabla.ConsumoPorReembolsar = oRecordset!ConsumoPorReembolsar
       oTabla.ReembolsoPagado = oRecordset!ReembolsoPagado
       oTabla.ReembolsoPorPagar = oRecordset!ReembolsoPorPagar
       oTabla.SaldoFinal = oRecordset!SaldoFinal
       oTabla.Documentos = IIf(IsNull(oRecordset!Documentos), "", oRecordset!Documentos)
       oTabla.idEstadoReembolso = IIf(IsNull(oRecordset!idEstadoReembolso), 0, oRecordset!idEstadoReembolso)
       oTabla.idTipoConsumo = IIf(IsNull(oRecordset!idTipoConsumo), 0, oRecordset!idTipoConsumo)
       oTabla.IdTipoComprobante = IIf(IsNull(oRecordset!IdTipoComprobante), 0, oRecordset!IdTipoComprobante)
       oTabla.GrabaDefinitivamente = IIf(IsNull(oRecordset!GrabaDefinitivamente), False, oRecordset!GrabaDefinitivamente)
   SeleccionarPorId = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOFactReembolsos)
   oTabla.idFactReembolso = 0
   oTabla.Anio = 0
   oTabla.Mes = 0
   oTabla.idAreaTramitaSeguro = 0
   oTabla.idFuenteFinanciamiento = 0
   oTabla.Descripcion = ""
   oTabla.SaldoInicial = 0
   oTabla.ConsumoPorReembolsar = 0
   oTabla.ReembolsoPagado = 0
   oTabla.ReembolsoPorPagar = 0
   oTabla.SaldoFinal = 0
   oTabla.Documentos = ""
   oTabla.idEstadoReembolso = 0
   oTabla.idTipoConsumo = 0
   oTabla.IdTipoComprobante = 0
   oTabla.GrabaDefinitivamente = False
End Sub

