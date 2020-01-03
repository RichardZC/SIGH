VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FarmHistPrecio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla farmAlmacen
'        Programado por: Barrantes D
'        Fecha: Diciembre 2008
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
'   Descripción:    Inserta un registro a la tabla farmHistPrecio
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOfarmHistPrecio) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "farmHistPrecioAgregar"
   Set oParameter = .CreateParameter("@idHistPrecio", adInteger, adParamOutput, 0)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@fecha", adDBTimeStamp, adParamInput, 0, IIf(oTabla.fecha = 0, Null, oTabla.fecha))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PrecioCompra", adCurrency, adParamInput, 0, oTabla.PrecioCompra)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PrecioDistribucion", adCurrency, adParamInput, 0, oTabla.PrecioDistribucion)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PrecioVenta", adCurrency, adParamInput, 0, oTabla.PrecioVenta)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PrecioDonacion", adCurrency, adParamInput, 0, oTabla.PrecioDonacion)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idUsuario", adInteger, adParamInput, 0, IIf(oTabla.idUsuario = 0, Null, oTabla.idUsuario))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
   .Parameters.Append oParameter
   .Execute
   oTabla.idHistPrecio = .Parameters("@idHistPrecio")
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
'   Descripción:    modifica un registro a la tabla farmHistPrecio
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOfarmHistPrecio) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "farmHistPrecioModificar"
   Set oParameter = .CreateParameter("@idHistPrecio", adInteger, adParamInput, 0, IIf(oTabla.idHistPrecio = 0, Null, oTabla.idHistPrecio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@fecha", adDBTimeStamp, adParamInput, 0, IIf(oTabla.fecha = 0, Null, oTabla.fecha))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PrecioCompra", adCurrency, adParamInput, 0, oTabla.PrecioCompra)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PrecioDistribucion", adCurrency, adParamInput, 0, oTabla.PrecioDistribucion)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PrecioVenta", adCurrency, adParamInput, 0, oTabla.PrecioVenta)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@PrecioDonacion", adCurrency, adParamInput, 0, oTabla.PrecioDonacion)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idUsuario", adInteger, adParamInput, 0, IIf(oTabla.idUsuario = 0, Null, oTabla.idUsuario))
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
'   Descripción:    elimia o borra un registro a la tabla farmHistPrecio
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOfarmHistPrecio) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "farmHistPrecioEliminar"
   Set oParameter = .CreateParameter("@idHistPrecio", adInteger, adParamInput, 0, oTabla.idHistPrecio)
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
'   Descripción:    Selecciona un unico registro de la tabla farmHistPrecio
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOfarmHistPrecio) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "farmHistPrecioSeleccionarPorId"
     Set oParameter = .CreateParameter("@idHistPrecio", adInteger, adParamInput, 0, oTabla.idHistPrecio)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.idHistPrecio = IIf(IsNull(oRecordset!idHistPrecio), 0, oRecordset!idHistPrecio)
       oTabla.idProducto = IIf(IsNull(oRecordset!idProducto), 0, oRecordset!idProducto)
       oTabla.fecha = IIf(IsNull(oRecordset!fecha), 0, oRecordset!fecha)
       oTabla.PrecioCompra = oRecordset!PrecioCompra
       oTabla.PrecioDistribucion = oRecordset!PrecioDistribucion
       oTabla.PrecioVenta = oRecordset!PrecioVenta
       oTabla.PrecioDonacion = oRecordset!PrecioDonacion
       oTabla.idUsuario = IIf(IsNull(oRecordset!idUsuario), 0, oRecordset!idUsuario)
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOfarmHistPrecio)
   oTabla.idHistPrecio = 0
   oTabla.idProducto = 0
   oTabla.fecha = Now
   oTabla.PrecioCompra = 0
   oTabla.PrecioDistribucion = 0
   oTabla.PrecioVenta = 0
   oTabla.PrecioDonacion = 0
   oTabla.idUsuario = 0
End Sub
