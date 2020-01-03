VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RecetaDetalleItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla RecetaDetalleItem
'        Programado por: Barrantes D
'        Fecha: Marzo 2012
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
'   Descripción:    Inserta un registro a la tabla RecetaDetalleItem
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DoRecetaDetalleItem) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "RecetaDetalleItemAgregar"
   Set oParameter = .CreateParameter("@idReceta", adInteger, adParamInput, 0, IIf(oTabla.IdReceta = 0, Null, oTabla.IdReceta))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idItem", adInteger, adParamInput, 0, IIf(oTabla.idItem = 0, Null, oTabla.idItem))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DocumentoDespacho", adVarChar, adParamInput, 20, IIf(oTabla.DocumentoDespacho = "", Null, oTabla.DocumentoDespacho))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CantidadDespachada", adInteger, adParamInput, 0, IIf(oTabla.CantidadDespachada = 0, Null, oTabla.CantidadDespachada))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = 0, Null, oTabla.IdComprobantePago))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEstado", adInteger, adParamInput, 0, IIf(oTabla.IdEstado = 0, Null, oTabla.IdEstado))
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
'   Descripción:    modifica un registro a la tabla RecetaDetalleItem
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DoRecetaDetalleItem) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "RecetaDetalleItemModificar"
   Set oParameter = .CreateParameter("@idReceta", adInteger, adParamInput, 0, IIf(oTabla.IdReceta = 0, Null, oTabla.IdReceta))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idItem", adInteger, adParamInput, 0, IIf(oTabla.idItem = 0, Null, oTabla.idItem))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DocumentoDespacho", adVarChar, adParamInput, 20, IIf(oTabla.DocumentoDespacho = "", Null, oTabla.DocumentoDespacho))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CantidadDespachada", adInteger, adParamInput, 0, IIf(oTabla.CantidadDespachada = 0, Null, oTabla.CantidadDespachada))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = 0, Null, oTabla.IdComprobantePago))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEstado", adInteger, adParamInput, 0, IIf(oTabla.IdEstado = 0, Null, oTabla.IdEstado))
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
'   Descripción:    elimia o borra un registro a la tabla RecetaDetalleItem
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DoRecetaDetalleItem) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "RecetaDetalleItemEliminar"
   Set oParameter = .CreateParameter("@idReceta", adInteger, adParamInput, 0, oTabla.IdReceta)
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
'   Descripción:    Selecciona un unico registro de la tabla RecetaDetalleItem
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DoRecetaDetalleItem) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "RecetaDetalleItemSeleccionarPorId"
     Set oParameter = .CreateParameter("@idReceta", adInteger, adParamInput, 0, oTabla.IdReceta)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.IdReceta = IIf(IsNull(oRecordset!IdReceta), 0, oRecordset!IdReceta)
       oTabla.idItem = IIf(IsNull(oRecordset!idItem), 0, oRecordset!idItem)
       oTabla.DocumentoDespacho = IIf(IsNull(oRecordset!DocumentoDespacho), "", oRecordset!DocumentoDespacho)
       oTabla.CantidadDespachada = IIf(IsNull(oRecordset!CantidadDespachada), 0, oRecordset!CantidadDespachada)
       oTabla.IdComprobantePago = IIf(IsNull(oRecordset!IdComprobantePago), 0, oRecordset!IdComprobantePago)
       oTabla.IdEstado = IIf(IsNull(oRecordset!IdEstado), 0, oRecordset!IdEstado)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DoRecetaDetalleItem)
   oTabla.IdReceta = 0
   oTabla.idItem = 0
   oTabla.DocumentoDespacho = ""
   oTabla.CantidadDespachada = 0
End Sub
Function ModificarItem(ByVal oTabla As DoRecetaDetalleItem) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
ModificarItem = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "RecetaDetalleItemModificarItem"
   Set oParameter = .CreateParameter("@idReceta", adInteger, adParamInput, 0, IIf(oTabla.IdReceta = 0, Null, oTabla.IdReceta))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idItem", adInteger, adParamInput, 0, IIf(oTabla.idItem = 0, Null, oTabla.idItem))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DocumentoDespacho", adVarChar, adParamInput, 20, IIf(oTabla.DocumentoDespacho = "", Null, oTabla.DocumentoDespacho))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@CantidadDespachada", adInteger, adParamInput, 0, IIf(oTabla.CantidadDespachada = 0, Null, oTabla.CantidadDespachada))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdComprobantePago", adInteger, adParamInput, 0, IIf(oTabla.IdComprobantePago = 0, Null, oTabla.IdComprobantePago))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdEstado", adInteger, adParamInput, 0, IIf(oTabla.IdEstado = 0, Null, oTabla.IdEstado))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
End With
 
ModificarItem = True
ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function


