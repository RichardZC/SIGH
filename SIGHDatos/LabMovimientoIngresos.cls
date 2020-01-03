VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LabMovimientoIngresos"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla LabMovimientoIngresos
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
'   Descripción:    Inserta un registro a la tabla LabMovimientoIngresos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOLabMovimientoIngresos) As Boolean
  On Error GoTo ManejadorDeError
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
 
  Insertar = False
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = mo_Conexion
    .CommandText = "LabMovimientoIngresosAgregar"
    Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, IIf(oTabla.IdMovimiento = 0, Null, oTabla.IdMovimiento))
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@NroDocumento", adVarChar, adParamInput, 20, IIf(oTabla.NroDocumento = "", Null, oTabla.NroDocumento))
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdPersonaRecepciona", adInteger, adParamInput, 0, IIf(oTabla.IdPersonaRecepciona = 0, Null, oTabla.IdPersonaRecepciona))
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
'   Descripción:    modifica un registro a la tabla LabMovimientoIngresos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------
Function Modificar(ByVal oTabla As DOLabMovimientoIngresos) As Boolean
  On Error GoTo ManejadorDeError
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
 
  Modificar = False
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = mo_Conexion
    .CommandText = "LabMovimientoIngresosModificar"
    Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, IIf(oTabla.IdMovimiento = 0, Null, oTabla.IdMovimiento))
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@NroDocumento", adVarChar, adParamInput, 20, IIf(oTabla.NroDocumento = "", Null, oTabla.NroDocumento))
    .Parameters.Append oParameter
    Set oParameter = .CreateParameter("@IdPersonaRecepciona", adInteger, adParamInput, 0, IIf(oTabla.IdPersonaRecepciona = 0, Null, oTabla.IdPersonaRecepciona))
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
'   Descripción:    elimia o borra un registro a la tabla LabMovimientoIngresos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------
Function Eliminar(ByVal oTabla As DOLabMovimientoIngresos) As Boolean
  On Error GoTo ManejadorDeError
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
 
  Eliminar = False
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = mo_Conexion
    .CommandText = "LabMovimientoIngresosEliminar"
    Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, oTabla.IdMovimiento)
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
'   Descripción:    Selecciona un unico registro de la tabla LabMovimientoIngresos
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------
Function SeleccionarPorId(ByVal oTabla As DOLabMovimientoIngresos) As Boolean
  On Error GoTo ManejadorDeError
  Dim oRecordset As New ADODB.Recordset
  Dim oCommand As New ADODB.Command
  Dim oParameter As ADODB.Parameter
 
  SeleccionarPorId = False
  With oCommand
    .CommandType = adCmdStoredProc
    Set .ActiveConnection = mo_Conexion
    .CommandText = "LabMovimientoIngresosSeleccionarPorId"
    Set oParameter = .CreateParameter("@IdMovimiento", adInteger, adParamInput, 0, oTabla.IdMovimiento)
    .Parameters.Append oParameter
    Set oRecordset = .Execute
  End With
  If Not (oRecordset.EOF And oRecordset.BOF) Then
    oTabla.IdMovimiento = IIf(IsNull(oRecordset!IdMovimiento), 0, oRecordset!IdMovimiento)
    oTabla.NroDocumento = IIf(IsNull(oRecordset!NroDocumento), "", oRecordset!NroDocumento)
    oTabla.IdPersonaRecepciona = IIf(IsNull(oRecordset!IdPersonaRecepciona), 0, oRecordset!IdPersonaRecepciona)
    SeleccionarPorId = True
    ms_MensajeError = ""
  End If
  oRecordset.Close
  Exit Function

ManejadorDeError:
  ms_MensajeError = Err.Number & " " + Err.Description
  Exit Function
End Function

Sub SetDefaults(ByVal oTabla As DOLabMovimientoIngresos)
  oTabla.IdMovimiento = 0
  oTabla.NroDocumento = ""
  oTabla.IdPersonaRecepciona = 0
End Sub

