VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FinanciamientoCatalogoServ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla FinanciamientoCatalogoServicios
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
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
'   Descripción:    Inserta un registro a la tabla FinanciamientoCatalogoServicios
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOFinanciamientoCatalogoServ) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FinanciamientoCatalogoServiciosAgregar"
           Set oParameter = .CreateParameter("@Activo", adBoolean, adParamInput, 0, IIf(oTabla.Activo = 0, Null, oTabla.Activo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.idTipoFinanciamiento = 0, Null, oTabla.idTipoFinanciamiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@PrecioUnitario", adCurrency, adParamInput, 0, IIf(oTabla.PrecioUnitario = 0, Null, oTabla.PrecioUnitario)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdFinanciamientoCatalogo", adInteger, adParamOutput, 0): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
       .Execute
           oTabla.IdFinanciamientoCatalogo = .Parameters("@IdFinanciamientoCatalogo")
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
'   Descripción:    modifica un registro a la tabla FinanciamientoCatalogoServicios
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOFinanciamientoCatalogoServ) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FinanciamientoCatalogoServiciosModificar"
           Set oParameter = .CreateParameter("@Activo", adBoolean, adParamInput, 0, IIf(oTabla.Activo = 0, Null, oTabla.Activo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 0, IIf(oTabla.idTipoFinanciamiento = 0, Null, oTabla.idTipoFinanciamiento)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@PrecioUnitario", adCurrency, adParamInput, 0, IIf(oTabla.PrecioUnitario = "", Null, oTabla.PrecioUnitario)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdFinanciamientoCatalogo", adInteger, adParamInput, 0, IIf(oTabla.IdFinanciamientoCatalogo = 0, Null, oTabla.IdFinanciamientoCatalogo)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla FinanciamientoCatalogoServicios
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOFinanciamientoCatalogoServ) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FinanciamientoCatalogoServiciosEliminar"
           Set oParameter = .CreateParameter("@IdFinanciamientoCatalogo", adInteger, adParamInput, 0, IIf(oTabla.IdFinanciamientoCatalogo = 0, Null, oTabla.IdFinanciamientoCatalogo)): .Parameters.Append oParameter
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

Sub ActualizaTabla(ByRef oTabla As DOFinanciamientoCatalogoServ, oRecordset As Recordset)
       oTabla.Activo = IIf(IsNull(oRecordset!Activo), 0, oRecordset!Activo)
       oTabla.idTipoFinanciamiento = IIf(IsNull(oRecordset!idTipoFinanciamiento), 0, oRecordset!idTipoFinanciamiento)
       oTabla.idProducto = IIf(IsNull(oRecordset!idProducto), 0, oRecordset!idProducto)
       oTabla.PrecioUnitario = IIf(IsNull(oRecordset!PrecioUnitario), "", oRecordset!PrecioUnitario)
       oTabla.IdFinanciamientoCatalogo = IIf(IsNull(oRecordset!IdFinanciamientoCatalogo), 0, oRecordset!IdFinanciamientoCatalogo)
       oTabla.SeUsaSinPrecio = IIf(IsNull(oRecordset!SeUsaSinPrecio), False, oRecordset!SeUsaSinPrecio)

End Sub

'------------------------------------------------------------------------------------
'   Función:        SeleccionarPorId
'   Descripción:    Selecciona un unico registro de la tabla FinanciamientoCatalogoServicios
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOFinanciamientoCatalogoServ) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "FinanciamientoCatalogoServiciosSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdFinanciamientoCatalogo", adInteger, adParamInput, 0, oTabla.IdFinanciamientoCatalogo): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       ActualizaTabla oTabla, oRecordset
'       oTabla.Activo = IIf(IsNull(oRecordset!Activo), 0, oRecordset!Activo)
'       oTabla.IdTipoFinanciamiento = IIf(IsNull(oRecordset!IdTipoFinanciamiento), 0, oRecordset!IdTipoFinanciamiento)
'       oTabla.idProducto = IIf(IsNull(oRecordset!idProducto), 0, oRecordset!idProducto)
'       oTabla.PrecioUnitario = IIf(IsNull(oRecordset!PrecioUnitario), "", oRecordset!PrecioUnitario)
'       oTabla.IdFinanciamientoCatalogo = IIf(IsNull(oRecordset!IdFinanciamientoCatalogo), 0, oRecordset!IdFinanciamientoCatalogo)
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


Function SeleccionarPorProductoYTipoFinanciamiento(ByVal oTabla As DOFinanciamientoCatalogoServ) As Boolean

On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "FactCatalogoServiciosXidTipoFinanciamiento"
     Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, oTabla.idProducto): .Parameters.Append oParameter
     Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 0, oTabla.idTipoFinanciamiento): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorProductoYTipoFinanciamiento = True
       ActualizaTabla oTabla, oRecordset
   Else
       SeleccionarPorProductoYTipoFinanciamiento = False
   End If

   oRecordset.Close
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

