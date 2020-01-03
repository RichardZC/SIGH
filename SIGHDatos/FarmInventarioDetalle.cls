VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "farmInventarioDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla farmInventarioDetalle
'        Programado por: Barrantes D
'        Fecha: Diciembre 2008
'
'------------------------------------------------------------------------------------
Option Explicit

Dim mo_Conexion As ADODB.Connection
Dim ml_IdUsuarioAuditoria As Long
Dim ms_MensajeError As String

Property Set Conexion(oValue As ADODB.Connection)
   Set mo_Conexion = oValue
End Property
Property Get Conexion() As ADODB.Connection
   Set Conexion = mo_Conexion
End Property
Property Let IdUsuarioAuditoria(lValue As Long)
   ml_IdUsuarioAuditoria = lValue
End Property
Property Get IdUsuarioAuditoria() As Long
   IdUsuarioAuditoria = ml_IdUsuarioAuditoria
End Property
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property


'------------------------------------------------------------------------------------
'   Función:        Insertar
'   Descripción:    Inserta un registro a la tabla farmInventarioDetalle
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DoFarmInventarioDetalle) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
If oTabla.idProducto = 720 And Trim(oTabla.Lote) = "111" Then
Insertar = False
End If
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "farmInventarioDetalleAgregar"
   Set oParameter = .CreateParameter("@idInventario", adInteger, adParamInput, 0, IIf(oTabla.idInventario = 0, Null, oTabla.idInventario))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Lote", adVarChar, adParamInput, 15, IIf(oTabla.Lote = "", Null, oTabla.Lote))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaVencimiento", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaVencimiento = 0, Null, oTabla.FechaVencimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Cantidad", adInteger, adParamInput, 0, oTabla.Cantidad)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Precio", adCurrency, adParamInput, 0, oTabla.Precio)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@RegistroSanitario", adVarChar, adParamInput, 50, IIf(oTabla.RegistroSanitario = "", Null, oTabla.RegistroSanitario))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTipoSalidaBienInsumo", adInteger, adParamInput, 0, IIf(oTabla.idTipoSalidaBienInsumo = 0, Null, oTabla.idTipoSalidaBienInsumo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CantidadSaldo", adInteger, adParamInput, 0, oTabla.CantidadSaldo)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CantidadFaltante", adInteger, adParamInput, 0, oTabla.CantidadFaltante)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CantidadSobrante", adInteger, adParamInput, 0, oTabla.CantidadSobrante)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EsHistoricoSaldo", adInteger, adParamInput, 0, oTabla.EsHistoricoSaldo)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
  ' oTabla.idInventario = .Parameters("@idInventario")
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
'   Descripción:    modifica un registro a la tabla farmInventarioDetalle
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DoFarmInventarioDetalle) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "farmInventarioDetalleModificar"
   Set oParameter = .CreateParameter("@idInventario", adInteger, adParamInput, 0, IIf(oTabla.idInventario = 0, Null, oTabla.idInventario))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Lote", adVarChar, adParamInput, 15, IIf(oTabla.Lote = "", Null, oTabla.Lote))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaVencimiento", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaVencimiento = 0, Null, oTabla.FechaVencimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Cantidad", adInteger, adParamInput, 0, oTabla.Cantidad)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Precio", adCurrency, adParamInput, 0, oTabla.Precio)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@RegistroSanitario", adVarChar, adParamInput, 50, IIf(oTabla.RegistroSanitario = "", Null, oTabla.RegistroSanitario))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTipoSalidaBienInsumo", adInteger, adParamInput, 0, IIf(oTabla.idTipoSalidaBienInsumo = 0, Null, oTabla.idTipoSalidaBienInsumo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CantidadSaldo", adInteger, adParamInput, 0, oTabla.CantidadSaldo)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CantidadFaltante", adInteger, adParamInput, 0, oTabla.CantidadFaltante)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CantidadSobrante", adInteger, adParamInput, 0, oTabla.CantidadSobrante)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@EsHistoricoSaldo", adInteger, adParamInput, 0, oTabla.EsHistoricoSaldo)
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
'   Descripción:    elimia o borra un registro a la tabla farmInventarioDetalle
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DoFarmInventarioDetalle) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "farmInventarioDetalleEliminar"
   Set oParameter = .CreateParameter("@idInventario", adInteger, adParamInput, 0, oTabla.idInventario)
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
'   Descripción:    Selecciona un unico registro de la tabla farmInventarioDetalle
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DoFarmInventarioDetalle) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "farmInventarioDetalleSeleccionarPorId"
     Set oParameter = .CreateParameter("@idInventario", adInteger, adParamInput, 0, oTabla.idInventario)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.idInventario = IIf(IsNull(oRecordset!idInventario), 0, oRecordset!idInventario)
       oTabla.idProducto = IIf(IsNull(oRecordset!idProducto), 0, oRecordset!idProducto)
       oTabla.Lote = IIf(IsNull(oRecordset!Lote), "", oRecordset!Lote)
       oTabla.FechaVencimiento = IIf(IsNull(oRecordset!FechaVencimiento), 0, oRecordset!FechaVencimiento)
       oTabla.Cantidad = IIf(IsNull(oRecordset!Cantidad), 0, oRecordset!Cantidad)
       oTabla.Precio = oRecordset!Precio
       oTabla.RegistroSanitario = IIf(IsNull(oRecordset!RegistroSanitario), "", oRecordset!RegistroSanitario)
       oTabla.idTipoSalidaBienInsumo = IIf(IsNull(oRecordset!idTipoSalidaBienInsumo), 1, oRecordset!idTipoSalidaBienInsumo)
       oTabla.CantidadSaldo = IIf(IsNull(oRecordset!CantidadSaldo), 0, oRecordset!CantidadSaldo)
       oTabla.CantidadFaltante = IIf(IsNull(oRecordset!CantidadFaltante), 0, oRecordset!CantidadFaltante)
       oTabla.CantidadSobrante = IIf(IsNull(oRecordset!CantidadSobrante), 0, oRecordset!CantidadSobrante)
       oTabla.EsHistoricoSaldo = IIf(IsNull(oRecordset!EsHistoricoSaldo), 0, oRecordset!EsHistoricoSaldo)
   SeleccionarPorId = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DoFarmInventarioDetalle)
   oTabla.idInventario = 0
   oTabla.idProducto = 0
   oTabla.Lote = ""
   oTabla.FechaVencimiento = Now
   oTabla.Cantidad = 0
   oTabla.Precio = 0
   oTabla.RegistroSanitario = ""
End Sub

Function DevuelveProductosLotesPorId(lnIdInventario As Long) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "farmInventarioDetalleDevuelveProductosLotesPorId"
        Set oParameter = .CreateParameter("@lnIdInventario", adInteger, adParamInput, 0, lnIdInventario): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set DevuelveProductosLotesPorId = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function



