VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FacturacionCatalogoPqte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla FacturacionCatalogoPaquetes
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
'   Descripción:    Inserta un registro a la tabla FacturacionCatalogoPaquetes
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DOFacturacionCatalogoPqtes) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "FacturacionCatalogoPaquetesAgregar"
   Set oParameter = .CreateParameter("@idFactPaquete", adInteger, adParamInput, 0, IIf(oTabla.idFactPaquete = 0, Null, oTabla.idFactPaquete))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idPuntoCarga", adInteger, adParamInput, 0, IIf(oTabla.idPuntoCarga = 0, Null, oTabla.idPuntoCarga))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idEspecialidadServicio", adInteger, adParamInput, 0, IIf(oTabla.idEspecialidadServicio = 0, Null, oTabla.idEspecialidadServicio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Cantidad", adInteger, adParamInput, 0, IIf(oTabla.Cantidad = 0, Null, oTabla.Cantidad))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Precio", adCurrency, adParamInput, 0, oTabla.Precio)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Importe", adCurrency, adParamInput, 0, oTabla.Importe)
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
'   Descripción:    modifica un registro a la tabla FacturacionCatalogoPaquetes
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOFacturacionCatalogoPqtes) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "FacturacionCatalogoPaquetesModificar"
   Set oParameter = .CreateParameter("@idFactPaquete", adInteger, adParamInput, 0, IIf(oTabla.idFactPaquete = 0, Null, oTabla.idFactPaquete))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idPuntoCarga", adInteger, adParamInput, 0, IIf(oTabla.idPuntoCarga = 0, Null, oTabla.idPuntoCarga))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idEspecialidadServicio", adInteger, adParamInput, 0, IIf(oTabla.idEspecialidadServicio = 0, Null, oTabla.idEspecialidadServicio))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Cantidad", adInteger, adParamInput, 0, IIf(oTabla.Cantidad = 0, Null, oTabla.Cantidad))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Precio", adCurrency, adParamInput, 0, oTabla.Precio)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Importe", adCurrency, adParamInput, 0, oTabla.Importe)
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
'   Descripción:    elimia o borra un registro a la tabla FacturacionCatalogoPaquetes
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOFacturacionCatalogoPqtes) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "FacturacionCatalogoPaquetesEliminar"
   Set oParameter = .CreateParameter("@idFactPaquete", adInteger, adParamInput, 0, oTabla.idFactPaquete)
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
'   Descripción:    Selecciona un unico registro de la tabla FacturacionCatalogoPaquetes
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOFacturacionCatalogoPqtes) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "FacturacionCatalogoPaquetesSeleccionarPorId"
     Set oParameter = .CreateParameter("@idFactPaquete", adInteger, adParamInput, 0, oTabla.idFactPaquete)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.idFactPaquete = IIf(IsNull(oRecordset!idFactPaquete), 0, oRecordset!idFactPaquete)
       oTabla.idPuntoCarga = IIf(IsNull(oRecordset!idPuntoCarga), 0, oRecordset!idPuntoCarga)
       oTabla.idEspecialidadServicio = IIf(IsNull(oRecordset!idEspecialidadServicio), 0, oRecordset!idEspecialidadServicio)
       oTabla.idProducto = IIf(IsNull(oRecordset!idProducto), 0, oRecordset!idProducto)
       oTabla.Cantidad = IIf(IsNull(oRecordset!Cantidad), 0, oRecordset!Cantidad)
       oTabla.Precio = oRecordset!Precio
       oTabla.Importe = oRecordset!Importe
   SeleccionarPorId = True
   End If
   ms_MensajeError = ""
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DOFacturacionCatalogoPqtes)
   oTabla.idFactPaquete = 0
   oTabla.idPuntoCarga = 0
   oTabla.idEspecialidadServicio = 0
   oTabla.idProducto = 0
   oTabla.Cantidad = 0
   oTabla.Precio = 0
   oTabla.Importe = 0
End Sub
