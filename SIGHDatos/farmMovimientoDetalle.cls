VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "farmMovimientoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla farmMovimientoDetalle
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
'   Descripción:    Inserta un registro a la tabla FactCatalogoServiciosN
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function usp_catalogoserviciosagregar(ByVal oTabla As DoFarmMovimientoDetalle) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
usp_catalogoserviciosagregar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "usp_catalogoserviciosagregar"
   Set oParameter = .CreateParameter("@Id", adInteger, adParamInput, 0, IIf(oTabla.Id = 0, Null, oTabla.Id))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Id_Atencion", adVarChar, adParamInput, 20, IIf(oTabla.Id_Atencion = "", Null, oTabla.Id_Atencion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Descripcion_Tipo_Item", adVarChar, adParamInput, 100, IIf(oTabla.Descripcion_Tipo_Item = "", Null, oTabla.Descripcion_Tipo_Item))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Fg_Tipo", adVarChar, adParamInput, 2, IIf(oTabla.Fg_Tipo = "", Null, oTabla.Fg_Tipo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 10, IIf(oTabla.codigo = "", Null, oTabla.codigo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@NombreProducto", adVarChar, adParamInput, 300, IIf(oTabla.NombreProducto = "", Null, oTabla.NombreProducto))
   .Parameters.Append oParameter
   
   Set oParameter = .CreateParameter("@IdSubClasificacion", adVarChar, adParamInput, 20, IIf(oTabla.IdSubClasificacion = "", Null, oTabla.IdSubClasificacion))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@LabConfHis", adVarChar, adParamInput, 10, IIf(oTabla.labConfHIS = "", Null, oTabla.labConfHIS))
   .Parameters.Append oParameter
   .Execute
End With
usp_catalogoserviciosagregar = True
ms_MensajeError = ""
 Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function


'------------------------------------------------------------------------------------
'   Función:        Insertar
'   Descripción:    Inserta un registro a la tabla farmMovimientoDetalle
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Insertar(ByVal oTabla As DoFarmMovimientoDetalle) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
If oTabla.Lote <> "" Then
   oTabla.Lote = Left(oTabla.Lote + Space(15), 15)
End If
Insertar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "farmMovimientoDetalleAgregar"
   Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, IIf(oTabla.movNumero = "", Null, oTabla.movNumero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, IIf(oTabla.movTipo = "", Null, oTabla.movTipo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Lote", adVarChar, adParamInput, 15, IIf(oTabla.Lote = "", Null, oTabla.Lote))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaVencimiento", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaVencimiento = 0, Null, oTabla.FechaVencimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Item", adInteger, adParamInput, 0, IIf(oTabla.Item = 0, Null, oTabla.Item))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Cantidad", adInteger, adParamInput, 0, oTabla.cantidad)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Precio", adCurrency, adParamInput, 0, oTabla.precio)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Total", adCurrency, adParamInput, 0, oTabla.Total)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@RegistroSanitario", adVarChar, adParamInput, 50, IIf(oTabla.RegistroSanitario = "", Null, oTabla.RegistroSanitario))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTipoSalidaBienInsumo", adInteger, adParamInput, 0, IIf(oTabla.idTipoSalidaBienInsumo = 0, Null, oTabla.idTipoSalidaBienInsumo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DocumentoNumero", adVarChar, adParamInput, 20, oTabla.DocumentoNumero)
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
'   Descripción:    modifica un registro a la tabla farmMovimientoDetalle
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DoFarmMovimientoDetalle) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
If oTabla.Lote <> "" Then
   oTabla.Lote = Left(oTabla.Lote + Space(15), 15)
End If
 
Modificar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "farmMovimientoDetalleModificar"
   Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, IIf(oTabla.movNumero = "", Null, oTabla.movNumero))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, IIf(oTabla.movTipo = "", Null, oTabla.movTipo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Lote", adVarChar, adParamInput, 15, IIf(oTabla.Lote = "", Null, oTabla.Lote))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaVencimiento", adDBTimeStamp, adParamInput, 0, IIf(oTabla.FechaVencimiento = 0, Null, oTabla.FechaVencimiento))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Item", adInteger, adParamInput, 0, IIf(oTabla.Item = 0, Null, oTabla.Item))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Cantidad", adInteger, adParamInput, 0, IIf(oTabla.cantidad = 0, Null, oTabla.cantidad))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Precio", adCurrency, adParamInput, 0, oTabla.precio)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@Total", adCurrency, adParamInput, 0, oTabla.Total)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@RegistroSanitario", adVarChar, adParamInput, 50, IIf(oTabla.RegistroSanitario = "", Null, oTabla.RegistroSanitario))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTipoSalidaBienInsumo", adInteger, adParamInput, 0, IIf(oTabla.idTipoSalidaBienInsumo = 0, Null, oTabla.idTipoSalidaBienInsumo))
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@DocumentoNumero", adVarChar, adParamInput, 20, oTabla.DocumentoNumero)
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
'   Descripción:    elimia o borra un registro a la tabla farmMovimientoDetalle
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DoFarmMovimientoDetalle) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
Eliminar = False
With oCommand
   .CommandType = adCmdStoredProc
   Set .ActiveConnection = mo_Conexion
   .CommandText = "farmMovimientoDetalleEliminar"
   Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, oTabla.movNumero)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, oTabla.movTipo)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, oTabla.idProducto)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@lote", adVarChar, adParamInput, 15, oTabla.Lote)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@FechaVencimiento", adDate, adParamInput, 10, oTabla.FechaVencimiento)
   .Parameters.Append oParameter
   Set oParameter = .CreateParameter("@idTipoSalidaBienInsumo", adInteger, adParamInput, 0, IIf(oTabla.idTipoSalidaBienInsumo = 0, Null, oTabla.idTipoSalidaBienInsumo))
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


Function usp_eliminacatalogoidat(Id As Long) As ADODB.Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim ms_MensajeError As String
    Dim oConexion As New ADODB.Connection
    oConexion.Open SIGHENTIDADES.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set usp_eliminacatalogoidat = Nothing
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "usp_eliminacatalogoidat"
        Set oParameter = .CreateParameter("@Id", adInteger, adParamInput, 0, Id): .Parameters.Append oParameter
               .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   oConexion.Close
   Set oConexion = Nothing
   Set usp_eliminacatalogoidat = oRecordset
   Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'------------------------------------------------------------------------------------
'   Función:        SeleccionarPorId
'   Descripción:    Selecciona un unico registro de la tabla farmMovimientoDetalle
'   Parámetros:     oTabla=campos de la Tabla
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DoFarmMovimientoDetalle) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "farmMovimientoDetalleSeleccionarPorId"
     Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, oTabla.movNumero)
     .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       oTabla.movNumero = IIf(IsNull(oRecordset!movNumero), "", oRecordset!movNumero)
       oTabla.movTipo = IIf(IsNull(oRecordset!movTipo), "", oRecordset!movTipo)
       oTabla.idProducto = IIf(IsNull(oRecordset!idProducto), 0, oRecordset!idProducto)
       oTabla.Lote = IIf(IsNull(oRecordset!Lote), "", oRecordset!Lote)
       oTabla.FechaVencimiento = IIf(IsNull(oRecordset!FechaVencimiento), 0, oRecordset!FechaVencimiento)
       oTabla.Item = IIf(IsNull(oRecordset!Item), 0, oRecordset!Item)
       oTabla.cantidad = IIf(IsNull(oRecordset!cantidad), 0, oRecordset!cantidad)
       oTabla.precio = oRecordset!precio
       oTabla.Total = oRecordset!Total
       oTabla.RegistroSanitario = IIf(IsNull(oRecordset!RegistroSanitario), "", oRecordset!RegistroSanitario)
       oTabla.idTipoSalidaBienInsumo = IIf(IsNull(oRecordset!idTipoSalidaBienInsumo), 1, oRecordset!idTipoSalidaBienInsumo)
       oTabla.DocumentoNumero = IIf(IsNull(oRecordset!DocumentoNumero), "", oRecordset!DocumentoNumero)    'Agrego FCV
   SeleccionarPorId = True
   ms_MensajeError = ""
   End If
   oRecordset.Close
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function
Sub SetDefaults(ByVal oTabla As DoFarmMovimientoDetalle)
   oTabla.movNumero = ""
   oTabla.movTipo = ""
   oTabla.idProducto = 0
   oTabla.Lote = ""
   oTabla.FechaVencimiento = Now
   oTabla.Item = 0
   oTabla.cantidad = 0
   oTabla.precio = 0
   oTabla.Total = 0
   oTabla.RegistroSanitario = ""
   oTabla.idTipoSalidaBienInsumo = sghTipoSalidaItemFarmacia.sghSoloVenta
   oTabla.DocumentoNumero = ""    'Agrego FCV
End Sub

Function DevuelveTodosItems(LcMovNumero As String, LcMovTipo As String) As ADODB.Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim ms_MensajeError As String
    Dim oConexion As New ADODB.Connection
    oConexion.Open SIGHENTIDADES.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set DevuelveTodosItems = Nothing
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmMovimientosDetalleDevuelveTodosItems"
        Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, LcMovNumero): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, LcMovTipo): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   oConexion.Close
   Set oConexion = Nothing
   Set DevuelveTodosItems = oRecordset
   Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'***************** LLENAR GRILLA SERVICIOS RAUL
Function usp_listarcatalogoidat(LcIdAtencion As String) As ADODB.Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim ms_MensajeError As String
    Dim oConexion As New ADODB.Connection
    oConexion.Open SIGHENTIDADES.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set usp_listarcatalogoidat = Nothing
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "usp_listarcatalogoidat"
        Set oParameter = .CreateParameter("@IdAtencion", adVarChar, adParamInput, 20, LcIdAtencion): .Parameters.Append oParameter
                Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   oConexion.Close
   Set oConexion = Nothing
   Set usp_listarcatalogoidat = oRecordset
   Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
'*******************************





Function FarmDevuelveMovimientosDeProducto(lnIdProducto As Long, ldFechaFin As Date) As ADODB.Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Dim ms_MensajeError As String
    Dim ldFechaInicio As Date
    Set FarmDevuelveMovimientosDeProducto = Nothing
    ms_MensajeError = ""
    ldFechaInicio = CDate("01/01/1900")
    oConexion.Open SIGHENTIDADES.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmDevuelveMovimientosDeProducto"
        Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FarmDevuelveMovimientosDeProducto = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function
Function FarmDevuelveMovimientosParaICIeIDI(ldFechaInicio As Date, ldFechaFin As Date, lnIdAlmacen As Long, LcMovTipo As String) As ADODB.Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Dim ms_MensajeError As String
    Set FarmDevuelveMovimientosParaICIeIDI = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 900
    oConexion.Open SIGHENTIDADES.CadenaConexion
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 900
        .CommandText = "FarmDevuelveMovimientosParaICIeIDI"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idAlmacen", adInteger, adParamInput, 0, lnIdAlmacen): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, LcMovTipo): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FarmDevuelveMovimientosParaICIeIDI = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

 '********************************   Actualiza Saldos por Producto  ***************************************
Function ActualizaSaldosPorProducto(lcEntradaOsalida As String, lnIdAlmacen As Long, lnIdProducto As Long, lcLote As String, ldFechaVencimiento As Date, lnidTipoSalidaBienInsumo As Long, lnCantidad As Long, lnPrecio As Double, ldFechaMovimiento As Date) As Boolean
    On Error GoTo ManejadorDeError
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    ActualizaSaldosPorProducto = False
    lcLote = Left(lcLote + Space(15), 15)
    With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FarmActualizaSaldosPorProducto"
       Set oParameter = .CreateParameter("@lcEntradaSalida", adVarChar, adParamInput, 1, lcEntradaOsalida): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdAlmacen", adInteger, adParamInput, 0, lnIdAlmacen): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@Lote", adVarChar, adParamInput, 15, lcLote): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@FechaVencimiento", adDate, adParamInput, 10, ldFechaVencimiento): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@Cantidad", adInteger, adParamInput, 0, lnCantidad): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@Precio", adDouble, adParamInput, 0, lnPrecio): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@idTipoSalidaBienInsumo", adInteger, adParamInput, 0, lnidTipoSalidaBienInsumo): .Parameters.Append oParameter
       .Execute
    End With
    If ActualizaSaldosMensuales(lcEntradaOsalida, lnIdAlmacen, lnIdProducto, ldFechaMovimiento, lnidTipoSalidaBienInsumo, lnCantidad, lcLote, ldFechaVencimiento, lnPrecio) = True Then
       ms_MensajeError = ""
       ActualizaSaldosPorProducto = True
    End If
    Exit Function
ManejadorDeError:
       ms_MensajeError = Err.Number & " " + Err.Description

End Function


Function FarmDevuelveSaldosConLotesSegunAlmacen(lnIdAlmacen As Long, lnOrden As Long, lcFiltro As String, oConexion As Connection) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError As String
    Set FarmDevuelveSaldosConLotesSegunAlmacen = Nothing
    ms_MensajeError = ""
'    oConexion.Open SIGHENTIDADES.CadenaConexion
'    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "farmDevuelveSaldosConLotesSegunAlmacen"
        Set oParameter = .CreateParameter("@idAlmacen", adInteger, adParamInput, 0, lnIdAlmacen): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Orden", adInteger, adParamInput, 0, lnOrden): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Filtro", adVarChar, adParamInput, 50, lcFiltro): .Parameters.Append oParameter
        Set oRecordset = .Execute
        'Set oRecordset.ActiveConnection = Nothing
   End With
'   oConexion.Close
'   Set oConexion = Nothing
   Set FarmDevuelveSaldosConLotesSegunAlmacen = oRecordset
   'Set oRecordset = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

'kike-13/05/2019
Function farmDevuelveSaldosSegunAlmacenProductoLote(lnIdAlmacen As Long, lnIdProducto As Long, lcLote As String, ldFechaVencimiento As Date, lnidTipoSalidaBienInsumo As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError As String
'Dim oConexion As New ADODB.Connection
    Set farmDevuelveSaldosSegunAlmacenProductoLote = Nothing
    ms_MensajeError = ""
'    oConexion.Open SIGHENTIDADES.CadenaConexion
'    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
'        Set .ActiveConnection = oConexion
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "farmDevuelveSaldosSegunAlmacenProductoLote"
        Set oParameter = .CreateParameter("@idAlmacen", adInteger, adParamInput, 0, lnIdAlmacen): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lote", adVarChar, adParamInput, 15, lcLote): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@fechaVencimiento", adDate, adParamInput, 10, ldFechaVencimiento): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idTipoSalidaBienInsumo", adInteger, adParamInput, 0, lnidTipoSalidaBienInsumo): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   'oConexion.Close
   'Set oConexion = Nothing
   Set farmDevuelveSaldosSegunAlmacenProductoLote = oRecordset
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function farmMovimientoDetalleDevuelveSalidasSegunAlmacenProductoLote(lnIdAlmacen As Long, lnIdProducto As Long, lcLote As String, ldFechaVencimiento As Date) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError As String
Dim oConexion As New ADODB.Connection
    Set farmMovimientoDetalleDevuelveSalidasSegunAlmacenProductoLote = Nothing
    ms_MensajeError = ""
    oConexion.Open SIGHENTIDADES.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "farmMovimientoDetalleDevuelveSalidasSegunAlmacenProductoLote"
        Set oParameter = .CreateParameter("@idAlmacen", adInteger, adParamInput, 0, lnIdAlmacen): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lote", adVarChar, adParamInput, 15, lcLote): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@fechaVencimiento", adDate, adParamInput, 10, ldFechaVencimiento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   oConexion.Close
   Set oConexion = Nothing
   Set farmMovimientoDetalleDevuelveSalidasSegunAlmacenProductoLote = oRecordset
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function ActualizaSaldosMensuales(lcEntradaOsalida As String, lnIdAlmacen As Long, lnIdProducto As Long, ldFechaMovimiento As Date, lnidTipoSalidaBienInsumo As Long, lnCantidad As Long, lcLote As String, ldFechaVencimiento As Date, lnPrecio As Double) As Boolean
    On Error GoTo ManejadorDeError
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim ldFechaMov As Date
    ActualizaSaldosMensuales = False
    ldFechaMov = SIGHENTIDADES.DevuelveFechaHoraFinalDelMesDelMovimiento(ldFechaMovimiento)
    lcLote = Left(lcLote + Space(15), 15)
    '
    With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FarmActualizaSaldosMensual"
       Set oParameter = .CreateParameter("@lcEntradaSalida", adVarChar, adParamInput, 1, lcEntradaOsalida): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdAlmacen", adInteger, adParamInput, 0, lnIdAlmacen): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@Cantidad", adInteger, adParamInput, 0, lnCantidad): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@FechaMov", adDate, adParamInput, 10, ldFechaMov): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@Lote", adVarChar, adParamInput, 15, lcLote): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@FechaVencimiento", adDate, adParamInput, 10, ldFechaVencimiento): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@Precio", adDouble, adParamInput, 0, lnPrecio): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@idTipoSalidaBienInsumo", adInteger, adParamInput, 0, lnidTipoSalidaBienInsumo): .Parameters.Append oParameter
       .Execute
    End With
    ActualizaSaldosMensuales = True
    ms_MensajeError = ""
    Exit Function
ManejadorDeError:
       ms_MensajeError = Err.Number & " " + Err.Description
End Function


Function FarmDevuelveMovimientosParaICIeIDIPorTproducto(ldFechaInicio As Date, ldFechaFin As Date, lnIdAlmacen As Long, LcMovTipo As String) As ADODB.Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Dim ms_MensajeError As String
    Set FarmDevuelveMovimientosParaICIeIDIPorTproducto = Nothing
    ms_MensajeError = ""
    oConexion.Open SIGHENTIDADES.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmDevuelveMovimientosParaICIeIDIPorTproducto"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idAlmacen", adInteger, adParamInput, 0, lnIdAlmacen): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, LcMovTipo): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set FarmDevuelveMovimientosParaICIeIDIPorTproducto = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function FarmDevuelveSaldosConLotesSegunAlmacenCliente(lnIdAlmacen As Long, lnOrden As Long, lcFiltro As String) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim ms_MensajeError As String
Dim oConexion As New ADODB.Connection
    ms_MensajeError = ""
    oConexion.Open SIGHENTIDADES.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmMovimientoDetalleDevuelveSaldosConLotesSegunAlmacenCliente"
        Set oParameter = .CreateParameter("@lnidAlmacen", adInteger, adParamInput, 0, lnIdAlmacen): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lnorden", adInteger, adParamInput, 0, lnOrden): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 100, lcFiltro): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   oConexion.Close
   Set oConexion = Nothing
   Set FarmDevuelveSaldosConLotesSegunAlmacenCliente = oRecordset
   Set oCommand = Nothing
   Set oRecordset = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function DevuelveTodosItemsSOPORTE(LcMovNumero As String, LcMovTipo As String) As ADODB.Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim ms_MensajeError As String
    Dim oConexion As New ADODB.Connection
    oConexion.CursorLocation = adUseClient
    oConexion.Open SIGHENTIDADES.CadenaConexion
    
    ms_MensajeError = ""
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FarmMovimientosDetalleDevuelveTodosItemsSOPORTE"
        Set oParameter = .CreateParameter("@MovNumero", adVarChar, adParamInput, 9, LcMovNumero): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, LcMovTipo): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   oConexion.Close
   Set oConexion = Nothing
   Set DevuelveTodosItemsSOPORTE = oRecordset
   Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function farmDevuelveMovimientosParaIDI(ldFechaInicio As Date, ldFechaFin As Date, lnIdAlmacen As Long, LcMovTipo As String) As ADODB.Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Dim ms_MensajeError As String
    Set farmDevuelveMovimientosParaIDI = Nothing
    ms_MensajeError = ""
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 900
    oConexion.Open SIGHENTIDADES.CadenaConexion
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 900
        .CommandText = "farmDevuelveMovimientosParaIDI"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idAlmacen", adInteger, adParamInput, 0, lnIdAlmacen): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, LcMovTipo): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set farmDevuelveMovimientosParaIDI = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function farmDevuelveMovimientosParaICI(ldFechaInicio As Date, ldFechaFin As Date, lnIdAlmacen As Long, LcMovTipo As String) As ADODB.Recordset
    On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Dim ms_MensajeError As String
    Set farmDevuelveMovimientosParaICI = Nothing
    ms_MensajeError = ""
    oConexion.Open SIGHENTIDADES.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "farmDevuelveMovimientosParaICI"
        Set oParameter = .CreateParameter("@FechaInicio", adDBTimeStamp, adParamInput, 0, ldFechaInicio): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@FechaFin", adDBTimeStamp, adParamInput, 0, ldFechaFin): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idAlmacen", adInteger, adParamInput, 0, lnIdAlmacen): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@MovTipo", adVarChar, adParamInput, 1, LcMovTipo): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set farmDevuelveMovimientosParaICI = oRecordset
   oConexion.Close
   Set oConexion = Nothing
   Set oRecordset = Nothing
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


