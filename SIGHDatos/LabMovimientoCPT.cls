VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LabMovimientoCPT"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla LabMovimientoCpt
'        Programado por: Bonilla A
'        Fecha: Mayo 2009
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
'   Descripción:    Inserta un registro a la tabla LabMovimientoCPT
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOLabMovimientoCPT) As Boolean
  On Error GoTo ManejadorDeError
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
 
  Insertar = False
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = mo_Conexion
    .CommandText = "LabMovimientoCPTAgregar"
    Set oParameter = .CreateParameter("@idMovimiento", adInteger, adParamInput, 0, IIf(oTabla.IdMovimiento = 0, Null, oTabla.IdMovimiento))
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@idProductoCPT", adInteger, adParamInput, 0, IIf(oTabla.idProductoCpt = 0, Null, oTabla.idProductoCpt))
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@cantidad", adInteger, adParamInput, 0, IIf(oTabla.Cantidad = 0, Null, oTabla.Cantidad))
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@precio", adCurrency, adParamInput, 0, oTabla.PRECIO)
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@importe", adCurrency, adParamInput, 0, oTabla.Importe)
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@ResultadoAutomatico", adInteger, adParamInput, 0, oTabla.ResultadoAutomatico)
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
'   Descripción:    modifica un registro a la tabla LabMovimientoCPT
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------
Function Modificar(ByVal oTabla As DOLabMovimientoCPT) As Boolean
  On Error GoTo ManejadorDeError
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
 
  Modificar = False
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = mo_Conexion
    .CommandText = "LabMovimientoCPTModificar"
    Set oParameter = .CreateParameter("@idMovimiento", adInteger, adParamInput, 0, IIf(oTabla.IdMovimiento = 0, Null, oTabla.IdMovimiento))
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@idProductoCPT", adInteger, adParamInput, 0, IIf(oTabla.idProductoCpt = 0, Null, oTabla.idProductoCpt))
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@cantidad", adInteger, adParamInput, 0, IIf(oTabla.Cantidad = 0, Null, oTabla.Cantidad))
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@precio", adCurrency, adParamInput, 0, oTabla.PRECIO)
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@importe", adCurrency, adParamInput, 0, oTabla.Importe)
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@ResultadoAutomatico", adInteger, adParamInput, 0, IIf(oTabla.ResultadoAutomatico = 0, Null, oTabla.ResultadoAutomatico))
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
'   Descripción:    elimia o borra un registro a la tabla LabMovimientoCPT
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------
Function Eliminar(ByVal oTabla As DOLabMovimientoCPT) As Boolean
  On Error GoTo ManejadorDeError
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
 
  Eliminar = False
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = mo_Conexion
    .CommandText = "LabMovimientoCPTEliminar"
    Set oParameter = .CreateParameter("@idMovimiento", adInteger, adParamInput, 0, oTabla.IdMovimiento)
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
'   Descripción:    Selecciona un unico registro de la tabla LabMovimientoCPT
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------
Function SeleccionarPorId(ByVal oTabla As DOLabMovimientoCPT) As Boolean
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
 
  SeleccionarPorId = False
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = mo_Conexion
    .CommandText = "LabMovimientoCPTSeleccionarPorId"
    Set oParameter = .CreateParameter("@idMovimiento", adInteger, adParamInput, 0, oTabla.IdMovimiento)
    .Parameters.Append oParameter
    Set oRecordset = .Execute
  End With
  If Not (oRecordset.EOF And oRecordset.BOF) Then
    oTabla.IdMovimiento = IIf(IsNull(oRecordset!IdMovimiento), 0, oRecordset!IdMovimiento)
    oTabla.idProductoCpt = IIf(IsNull(oRecordset!idProductoCpt), 0, oRecordset!idProductoCpt)
    oTabla.Cantidad = IIf(IsNull(oRecordset!Cantidad), 0, oRecordset!Cantidad)
    oTabla.PRECIO = oRecordset!PRECIO
    oTabla.Importe = oRecordset!Importe
    oTabla.ResultadoAutomatico = IIf(IsNull(oRecordset!ResultadoAutomatico), 0, oRecordset!ResultadoAutomatico)
    SeleccionarPorId = True
    ms_MensajeError = ""
  End If
  oRecordset.Close
  Exit Function

ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description
  Exit Function
End Function

Sub SetDefaults(ByVal oTabla As DOLabMovimientoCPT)
  oTabla.IdMovimiento = 0
  oTabla.idProductoCpt = 0
  oTabla.Cantidad = 0
  oTabla.PRECIO = 0
  oTabla.Importe = 0
End Sub

