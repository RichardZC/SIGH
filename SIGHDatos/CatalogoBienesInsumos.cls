VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CatalogoBienesInsumos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla FactCatalogoBienesInsumos
'        Programado por: Barrantes D
'        Fecha: Agosto 2009
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
'   Descripción:    Inserta un registro a la tabla FactCatalogoBienesInsumos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOCatalogoBienesInsumos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FactCatalogoBienesInsumosAgregar"
           Set oParameter = .CreateParameter("@IdCentroCosto", adInteger, adParamInput, 0, IIf(oTabla.IdCentroCosto = 0, Null, oTabla.IdCentroCosto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPartida", adInteger, adParamInput, 0, IIf(oTabla.IdPartida = 0, Null, oTabla.IdPartida)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdSubGrupoFarmacologico", adInteger, adParamInput, 0, IIf(oTabla.IdSubGrupoFarmacologico = 0, Null, oTabla.IdSubGrupoFarmacologico)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdGrupoFarmacologico", adInteger, adParamInput, 0, IIf(oTabla.IdGrupoFarmacologico = 0, Null, oTabla.IdGrupoFarmacologico)): .Parameters.Append oParameter
           'Set oParameter = .CreateParameter("@IdSubGrupoFarmacologico", adInteger, adParamInput, 0, IIf(oTabla.IdSubGrupoFarmacologico = 0, Null, oTabla.IdSubGrupoFarmacologico)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NombreComercial", adVarChar, adParamInput, 50, IIf(oTabla.NombreComercial = "", Null, oTabla.NombreComercial)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Nombre", adVarChar, adParamInput, 250, IIf(oTabla.nombre = "", Null, oTabla.nombre)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 20, IIf(oTabla.codigo = "", Null, oTabla.codigo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamOutput, 0): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@PrecioCompra", adCurrency, adParamInput, 0, oTabla.PrecioCompra): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@PrecioDistribucion", adCurrency, adParamInput, 0, oTabla.PrecioDistribucion): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@PrecioDonacion", adCurrency, adParamInput, 0, oTabla.PrecioDonacion): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@PrecioUltCompra", adCurrency, adParamInput, 0, oTabla.PrecioUltCompra): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idTipoSalidaBienInsumo", adInteger, adParamInput, 0, oTabla.idTipoSalidaBienInsumo): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@StockMinimo", adInteger, adParamInput, 0, IIf(oTabla.StockMinimo = 0, Null, oTabla.StockMinimo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TipoProducto", adInteger, adParamInput, 0, oTabla.TipoProducto): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Denominacion", adVarChar, adParamInput, 100, IIf(oTabla.denominacion = "", Null, oTabla.denominacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Concentracion", adVarChar, adParamInput, 100, IIf(oTabla.Concentracion = "", Null, oTabla.Concentracion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Presentacion", adVarChar, adParamInput, 100, IIf(oTabla.Presentacion = "", Null, oTabla.Presentacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FormaFarmaceutica", adVarChar, adParamInput, 10, IIf(oTabla.FormaFarmaceutica = "", Null, oTabla.FormaFarmaceutica)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@MaterialEnvase", adVarChar, adParamInput, 100, IIf(oTabla.MaterialEnvase = "", Null, oTabla.MaterialEnvase)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@PresentacionEnvase", adVarChar, adParamInput, 100, IIf(oTabla.PresentacionEnvase = "", Null, oTabla.PresentacionEnvase)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Fabricante", adVarChar, adParamInput, 100, IIf(oTabla.Fabricante = "", Null, oTabla.Fabricante)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPaisOrigen", adInteger, adParamInput, 0, IIf(oTabla.IdPaisOrigen = 0, Null, oTabla.IdPaisOrigen)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Petitorio", adBoolean, adParamInput, 0, oTabla.Petitorio): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TipoProductoSismed", adVarChar, adParamInput, 1, oTabla.TipoProductoSismed): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@codigoSUNAT", adVarChar, adParamInput, 20, IIf(oTabla.codigoSUNAT = "", Null, oTabla.codigoSUNAT)): .Parameters.Append oParameter
           
       .Execute
           oTabla.idProducto = .Parameters("@IdProducto")
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
'   Descripción:    modifica un registro a la tabla FactCatalogoBienesInsumos
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOCatalogoBienesInsumos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FactCatalogoBienesInsumosModificar"
           Set oParameter = .CreateParameter("@IdCentroCosto", adInteger, adParamInput, 0, IIf(oTabla.IdCentroCosto = 0, Null, oTabla.IdCentroCosto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPartida", adInteger, adParamInput, 0, IIf(oTabla.IdPartida = 0, Null, oTabla.IdPartida)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdSubGrupoFarmacologico", adInteger, adParamInput, 0, IIf(oTabla.IdSubGrupoFarmacologico = 0, Null, oTabla.IdSubGrupoFarmacologico)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdGrupoFarmacologico", adInteger, adParamInput, 0, IIf(oTabla.IdGrupoFarmacologico = 0, Null, oTabla.IdGrupoFarmacologico)): .Parameters.Append oParameter
           'Set oParameter = .CreateParameter("@IdSubGrupoFarmacologico", adInteger, adParamInput, 0, IIf(oTabla.IdSubGrupoFarmacologico = 0, Null, oTabla.IdSubGrupoFarmacologico)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NombreComercial", adVarChar, adParamInput, 50, IIf(oTabla.NombreComercial = "", Null, oTabla.NombreComercial)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Nombre", adVarChar, adParamInput, 250, IIf(oTabla.nombre = "", Null, oTabla.nombre)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 20, IIf(oTabla.codigo = "", Null, oTabla.codigo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto)): .Parameters.Append oParameter
            Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria)
            .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@PrecioCompra", adCurrency, adParamInput, 0, oTabla.PrecioCompra): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@PrecioDistribucion", adCurrency, adParamInput, 0, oTabla.PrecioDistribucion): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@PrecioDonacion", adCurrency, adParamInput, 0, oTabla.PrecioDonacion): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@PrecioUltCompra", adCurrency, adParamInput, 0, oTabla.PrecioUltCompra): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idTipoSalidaBienInsumo", adInteger, adParamInput, 0, oTabla.idTipoSalidaBienInsumo): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@StockMinimo", adInteger, adParamInput, 0, IIf(oTabla.StockMinimo = 0, Null, oTabla.StockMinimo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TipoProducto", adInteger, adParamInput, 0, oTabla.TipoProducto): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Denominacion", adVarChar, adParamInput, 100, IIf(oTabla.denominacion = "", Null, oTabla.denominacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Concentracion", adVarChar, adParamInput, 100, IIf(oTabla.Concentracion = "", Null, oTabla.Concentracion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Presentacion", adVarChar, adParamInput, 100, IIf(oTabla.Presentacion = "", Null, oTabla.Presentacion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@FormaFarmaceutica", adVarChar, adParamInput, 10, IIf(oTabla.FormaFarmaceutica = "", Null, oTabla.FormaFarmaceutica)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@MaterialEnvase", adVarChar, adParamInput, 100, IIf(oTabla.MaterialEnvase = "", Null, oTabla.MaterialEnvase)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@PresentacionEnvase", adVarChar, adParamInput, 100, IIf(oTabla.PresentacionEnvase = "", Null, oTabla.PresentacionEnvase)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Fabricante", adVarChar, adParamInput, 100, IIf(oTabla.Fabricante = "", Null, oTabla.Fabricante)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPaisOrigen", adInteger, adParamInput, 0, IIf(oTabla.IdPaisOrigen = 0, Null, oTabla.IdPaisOrigen)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Petitorio", adBoolean, adParamInput, 0, oTabla.Petitorio): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@TipoProductoSismed", adVarChar, adParamInput, 1, oTabla.TipoProductoSismed): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@codigoSUNAT", adVarChar, adParamInput, 20, IIf(oTabla.codigoSUNAT = "", Null, oTabla.codigoSUNAT)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla FactCatalogoBienesInsumos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOCatalogoBienesInsumos) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FactCatalogoBienesInsumosEliminar"
           Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto)): .Parameters.Append oParameter
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
'   Descripción:    Selecciona un unico registro de la tabla FactCatalogoBienesInsumos
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function SeleccionarPorId(ByVal oTabla As DOCatalogoBienesInsumos) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "FactCatalogoBienesInsumosSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, oTabla.idProducto): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.IdCentroCosto = IIf(IsNull(oRecordset!IdCentroCosto), 0, oRecordset!IdCentroCosto)
       oTabla.IdPartida = IIf(IsNull(oRecordset!IdPartida), 0, oRecordset!IdPartida)
       oTabla.IdSubGrupoFarmacologico = IIf(IsNull(oRecordset!IdSubGrupoFarmacologico), 0, oRecordset!IdSubGrupoFarmacologico)
       oTabla.IdGrupoFarmacologico = IIf(IsNull(oRecordset!IdGrupoFarmacologico), 0, oRecordset!IdGrupoFarmacologico)
       oTabla.IdSubGrupoFarmacologico = IIf(IsNull(oRecordset!IdSubGrupoFarmacologico), 0, oRecordset!IdSubGrupoFarmacologico)
       oTabla.NombreComercial = IIf(IsNull(oRecordset!NombreComercial), "", oRecordset!NombreComercial)
       oTabla.nombre = IIf(IsNull(oRecordset!nombre), "", oRecordset!nombre)
       oTabla.codigo = IIf(IsNull(oRecordset!codigo), "", oRecordset!codigo)
       oTabla.idProducto = IIf(IsNull(oRecordset!idProducto), 0, oRecordset!idProducto)
       oTabla.PrecioCompra = IIf(IsNull(oRecordset!PrecioCompra), 0, oRecordset!PrecioCompra)
       oTabla.PrecioDistribucion = IIf(IsNull(oRecordset!PrecioDistribucion), 0, oRecordset!PrecioDistribucion)
       oTabla.PrecioDonacion = IIf(IsNull(oRecordset!PrecioDonacion), 0, oRecordset!PrecioDonacion)
       oTabla.PrecioUltCompra = IIf(IsNull(oRecordset!PrecioUltCompra), 0, oRecordset!PrecioUltCompra)
       oTabla.idTipoSalidaBienInsumo = IIf(IsNull(oRecordset!idTipoSalidaBienInsumo), 0, oRecordset!idTipoSalidaBienInsumo)
       oTabla.StockMinimo = IIf(IsNull(oRecordset!StockMinimo), 0, oRecordset!StockMinimo)
       oTabla.TipoProducto = IIf(IsNull(oRecordset!TipoProducto), 0, oRecordset!TipoProducto)
       oTabla.denominacion = IIf(IsNull(oRecordset!denominacion), "", oRecordset!denominacion)
       oTabla.Concentracion = IIf(IsNull(oRecordset!Concentracion), "", oRecordset!Concentracion)
       oTabla.Presentacion = IIf(IsNull(oRecordset!Presentacion), "", oRecordset!Presentacion)
       oTabla.FormaFarmaceutica = IIf(IsNull(oRecordset!FormaFarmaceutica), "", oRecordset!FormaFarmaceutica)
       oTabla.MaterialEnvase = IIf(IsNull(oRecordset!MaterialEnvase), "", oRecordset!MaterialEnvase)
       oTabla.PresentacionEnvase = IIf(IsNull(oRecordset!PresentacionEnvase), "", oRecordset!PresentacionEnvase)
       oTabla.Fabricante = IIf(IsNull(oRecordset!Fabricante), "", oRecordset!Fabricante)
       oTabla.IdPaisOrigen = IIf(IsNull(oRecordset!IdPaisOrigen), 0, oRecordset!IdPaisOrigen)
       oTabla.Petitorio = IIf(IsNull(oRecordset!Petitorio), 0, oRecordset!Petitorio)
       oTabla.TipoProductoSismed = IIf(IsNull(oRecordset!TipoProductoSismed), " ", oRecordset!TipoProductoSismed)
       oTabla.codigoSUNAT = IIf(IsNull(oRecordset!codigoSUNAT), " ", oRecordset!codigoSUNAT)
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




Function Filtrar(ByVal oTabla As DOCatalogoBienesInsumos) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String

    sSql = ""
    If oTabla.idProducto <> 0 Then
        sSql = sSql + " cb.IdProducto = " & oTabla.idProducto & " and"
    End If
    If oTabla.nombre <> "" Then
        sSql = sSql + " cb.Nombre like '%" & oTabla.nombre & "%'" & " and"
    End If
    If oTabla.IdSubGrupoFarmacologico <> 0 Then
        sSql = sSql + " cb.IdSubGrupoFarmacologico = " & oTabla.IdSubGrupoFarmacologico & " and"
    End If
    If sSql <> "" Then
       sSql = Left(sSql, Trim(Len(sSql)) - 4)
    End If
    sSql = sSql + " order by cb.Nombre"
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "CatalogoBienesInsumosFiltrar"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, sSql): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set Filtrar = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function



Function FiltrarCatalogoBase(ByVal oTabla As DOCatalogoBienesInsumos) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSql As String
Dim sSQL1 As String
Dim sSQL2 As String
Dim sWhere As String

    Set FiltrarCatalogoBase = Nothing
    sWhere = ""

    sSQL1 = "select IdSubGrupoFarmacologico, 0 as IdProducto, '' as Codigo, Descripcion from FactInsumosSubGrupoFarmacologico"

    sSQL2 = "select IdSubGrupoFarmacologico, IdProducto, Codigo, Nombre, 0 as PrecioUnitario, 1 as Activo from FactCatalogoBienesInsumos "

    If oTabla.codigo <> "" Then
        sWhere = sWhere + " and Codigo = '" & oTabla.codigo & "'"
    End If
    If oTabla.nombre <> "" Then
        sWhere = sWhere + " and Nombre like '%" & oTabla.nombre & "%'"
    End If

    If sWhere <> "" Then sSQL2 = sSQL2 + " where " + Mid(sWhere, 5)

    sSQL2 = sSQL2 + " order by nombre"

    sSql = "SHAPE {" + sSQL1 + "} " + _
            "APPEND ({" + sSQL2 + "} " + _
            " AS title_chap RELATE IdSubGrupoFarmacologico TO IdSubGrupoFarmacologico) "

    oRecordset.Open sSql, mo_Conexion, adOpenDynamic, adLockOptimistic
    Set oRecordset.ActiveConnection = Nothing

    Set FiltrarCatalogoBase = oRecordset

    ms_MensajeError = ""

Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
'Stop
'On Error GoTo ManejadorDeError
'Dim oRecordSet As New Recordset
'Dim oCommand As New ADODB.Command
'Dim oParameter As ADODB.Parameter
'Dim sWhere As String
'    sWhere = ""
'    If oTabla.Codigo <> "" Then
'        sWhere = sWhere + " and Codigo = '" & oTabla.Codigo & "'"
'    End If
'    If oTabla.Nombre <> "" Then
'        sWhere = sWhere + " and Nombre like '%" & oTabla.Nombre & "%'"
'    End If
'    If sWhere <> "" Then sWhere = " where " + Mid(sWhere, 5)
'    sWhere = sWhere + " order by nombre"
'    With oCommand
'        .CommandType = adCmdStoredProc
'        Set .ActiveConnection = mo_Conexion
'        .CommandTimeout = 150
'        .CommandText = "FactCatalogoBienesInsumosFiltrarCatalogoFiltrarCatalogo"
'        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, sWhere): .Parameters.Append oParameter
'        Set oRecordSet = .Execute
'        Set oRecordSet.ActiveConnection = Nothing
'   End With
'   Set FiltrarCatalogoBase = oRecordSet
'   Set oCommand = Nothing
'   Exit Function
'ManejadorDeError:
'    MsgBox Err.Description
End Function

Function FiltrarCatalogo(ByVal oTabla As DOCatalogoBienesInsumos, lTipoCatalogo As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSql As String
Dim sSQL1 As String
Dim sSQL2 As String
Dim sWhere As String

    Set FiltrarCatalogo = Nothing
    sWhere = ""

    sSQL1 = "select IdSubGrupoFarmacologico, 0 as IdProducto, '' as Codigo, Descripcion from FactInsumosSubGrupoFarmacologico"

    sSQL2 = "select FactCatalogoBienesInsumos.IdSubGrupoFarmacologico," + _
    " FactCatalogoBienesInsumos.IdProducto, " + _
    " FactCatalogoBienesInsumos.Codigo," + _
    " FactCatalogoBienesInsumos.Nombre, " + _
    " FactCatalogoBienesInsumosHosp.PrecioUnitario, " + _
    " FactCatalogoBienesInsumosHosp.Activo " + _
    " From FactCatalogoBienesInsumosHosp, FactCatalogoBienesInsumos " + _
    " Where " + _
    " FactCatalogoBienesInsumosHosp.IdTipoFinanciamiento = " & lTipoCatalogo & _
    " and FactCatalogoBienesInsumosHosp.IdProducto = FactCatalogoBienesInsumos.IdProducto"

    If oTabla.codigo <> "" Then
        sSQL2 = sSQL2 + " and FactCatalogoBienesInsumos.Codigo = '" & oTabla.codigo & "'"
    End If
    If oTabla.nombre <> "" Then
        sSQL2 = sSQL2 + " and FactCatalogoBienesInsumos.Nombre like '%" & oTabla.nombre & "%'"
    End If

    sSQL2 = sSQL2 + " order by codigo"


    sSql = "SHAPE {" + sSQL1 + "} " + _
            "APPEND ({" + sSQL2 + "} " + _
            " AS title_chap RELATE IdSubGrupoFarmacologico TO IdSubGrupoFarmacologico) "


    oRecordset.Open sSql, mo_Conexion, adOpenDynamic, adLockOptimistic
    Set oRecordset.ActiveConnection = Nothing

    Set FiltrarCatalogo = oRecordset

    ms_MensajeError = ""

Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
'Stop
'On Error GoTo ManejadorDeError
'Dim oRecordset As New Recordset
'Dim oCommand As New ADODB.Command
'Dim oParameter As ADODB.Parameter
'Dim sSQL2 As String
'    sSQL2 = ""
'
'   With oCommand
'        .CommandType = adCmdStoredProc
'        Set .ActiveConnection = mo_Conexion
'        .CommandTimeout = 150
'        .CommandText = "FactCatalogoBienesInsumosFiltrarCatalogoBase"
'        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, sSQL2): .Parameters.Append oParameter
'        Set oParameter = .CreateParameter("@lTipoCatalogo", adInteger, adParamInput, 0, lTipoCatalogo): .Parameters.Append oParameter
'        Set oRecordset = .Execute
'        Set oRecordset.ActiveConnection = Nothing
'   End With
'   Set FiltrarCatalogo = oRecordset
'   Set oCommand = Nothing
'   Exit Function
'ManejadorDeError:
'    MsgBox Err.Description
End Function

Function FiltrarCatalogoDEBB(ByVal oTabla As DOCatalogoBienesInsumos, lTipoCatalogo As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSql As String
Dim sSQL1 As String
Dim sSQL2 As String
Dim sWhere As String

    Set FiltrarCatalogoDEBB = Nothing
    sWhere = ""

    sSQL1 = "select IdSubGrupoFarmacologico, 0 as IdProducto, '' as Codigo, Descripcion from FactInsumosSubGrupoFarmacologico"

'    sSQL2 = "select FactCatalogoBienesInsumos.IdSubGrupoFarmacologico," + _
'    " FactCatalogoBienesInsumos.IdProducto, " + _
'    " FactCatalogoBienesInsumos.Codigo," + _
'    " FactCatalogoBienesInsumos.Nombre, " + _
'    " FactCatalogoBienesInsumosHosp.PrecioUnitario, " + _
'    " FactCatalogoBienesInsumosHosp.Activo, " + _
'    " FactCatalogoBienesInsumos.PrecioCompra, FactCatalogoBienesInsumos.PrecioUltCompra, " + _
'    " FactCatalogoBienesInsumos.PrecioDistribucion, FactCatalogoBienesInsumos.PrecioDonacion, " + _
'    " FactCatalogoBienesInsumosHosp.idPlanCatalogo,FactCatalogoBienesInsumos.TipoProducto, " + _
'    " FactCatalogoBienesInsumos.codigoSunat, sunat_onu.descripcion as sunat" + _
'    " From FactCatalogoBienesInsumosHosp, FactCatalogoBienesInsumos, sunat_onu " + _
'    " Where " + _
'    " FactCatalogoBienesInsumosHosp.IdTipoFinanciamiento = " & lTipoCatalogo & _
'    " and FactCatalogoBienesInsumosHosp.IdProducto = FactCatalogoBienesInsumos.IdProducto " & _
'    " and sunat_onu.codigo = FactCatalogoBienesInsumos.codigoSunat"
    
    sSQL2 = "select FactCatalogoBienesInsumos.IdSubGrupoFarmacologico," + _
    " FactCatalogoBienesInsumos.IdProducto, " + _
    " FactCatalogoBienesInsumos.Codigo," + _
    " FactCatalogoBienesInsumos.Nombre, " + _
    " FactCatalogoBienesInsumosHosp.PrecioUnitario, " + _
    " FactCatalogoBienesInsumosHosp.Activo, " + _
    " FactCatalogoBienesInsumos.PrecioCompra, FactCatalogoBienesInsumos.PrecioUltCompra, " + _
    " FactCatalogoBienesInsumos.PrecioDistribucion, FactCatalogoBienesInsumos.PrecioDonacion, " + _
    " FactCatalogoBienesInsumosHosp.idPlanCatalogo,FactCatalogoBienesInsumos.TipoProducto, " + _
    " FactCatalogoBienesInsumos.codigoSunat, sunat_onu.descripcion as sunat" + _
    " FROM  FactCatalogoBienesInsumos INNER JOIN " + _
    "       FactCatalogoBienesInsumosHosp ON " + _
    "       FactCatalogoBienesInsumos.IdProducto = FactCatalogoBienesInsumosHosp.IdProducto LEFT OUTER JOIN " + _
    "       sunat_onu ON FactCatalogoBienesInsumos.CodigoSunat  = sunat_onu.codigo " + _
    " Where  FactCatalogoBienesInsumosHosp.IdTipoFinanciamiento = " & lTipoCatalogo
    

    If oTabla.codigo <> "" Then
        sSQL2 = sSQL2 + " and FactCatalogoBienesInsumos.Codigo = '" & oTabla.codigo & "'"
    End If
    If oTabla.nombre <> "" Then
        sSQL2 = sSQL2 + " and FactCatalogoBienesInsumos.Nombre like '%" & oTabla.nombre & "%'"
    End If

    sSQL2 = sSQL2 + " order by nombre"


    sSql = "SHAPE {" + sSQL1 + "} " + _
            "APPEND ({" + sSQL2 + "} " + _
            " AS title_chap RELATE IdSubGrupoFarmacologico TO IdSubGrupoFarmacologico) "


    oRecordset.Open sSQL2, mo_Conexion, adOpenDynamic, adLockOptimistic
    Set oRecordset.ActiveConnection = Nothing

    Set FiltrarCatalogoDEBB = oRecordset

    ms_MensajeError = ""

Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
'Stop
'On Error GoTo ManejadorDeError
'Dim oRecordset As New Recordset
'Dim oCommand As New ADODB.Command
'Dim oParameter As ADODB.Parameter
'Dim sSQL2 As String
'    sSQL2 = ""
'    If oTabla.Codigo <> "" Then
'        sSQL2 = sSQL2 + " and FactCatalogoBienesInsumos.Codigo = '" & oTabla.Codigo & "'"
'    End If
'    If oTabla.Nombre <> "" Then
'        sSQL2 = sSQL2 + " and FactCatalogoBienesInsumos.Nombre like '%" & oTabla.Nombre & "%'"
'    End If
'    sSQL2 = sSQL2 + " order by nombre"
'    With oCommand
'        .CommandType = adCmdStoredProc
'        Set .ActiveConnection = mo_Conexion
'        .CommandTimeout = 150
'        .CommandText = "FactCatalogoBienesInsumosFiltrarCatalogoDEBB"
'        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, sSQL2): .Parameters.Append oParameter
'        Set oParameter = .CreateParameter("@lTipoCatalogo", adInteger, adParamInput, 0, lTipoCatalogo): .Parameters.Append oParameter
'        Set oRecordset = .Execute
'        Set oRecordset.ActiveConnection = Nothing
'   End With
'   Set FiltrarCatalogoDEBB = oRecordset
'   Set oCommand = Nothing
'   Exit Function
'ManejadorDeError:
'    MsgBox Err.Description
End Function


Function SeleccionarPorTipoCatalogo(oDoCatalogoBienInsumo As DOCatalogoBienesInsumos, lTipoCatalogo As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sOrder As String
    sSql = ""
    sOrder = ""
    If oDoCatalogoBienInsumo.codigo <> "" Then
        sSql = sSql + " and Codigo = '" + oDoCatalogoBienInsumo.codigo + "'"
        sOrder = sOrder + " ,Codigo"
    End If
    If oDoCatalogoBienInsumo.nombre <> "" Then
        sSql = sSql + " and Nombre like '%" + oDoCatalogoBienInsumo.nombre + "%'"
        sOrder = sOrder + " ,Nombre"
    End If
    If sOrder <> "" Then sSql = Mid(sSql, 5) + " order by " + Mid(sOrder, 3)
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "FactCatalogoBienesInsumosHospSeleccionarPorTipoCatalogo"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, sSql): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lTipoCatalogo", adVarChar, adParamInput, 20, Trim(Str(lTipoCatalogo))): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarPorTipoCatalogo = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


Function SeleccionarPorCuentaAtencionNoPagados(IdCuentaAtencion As Long, lTipoCatalogo As Long, IdPuntosDeCarga As Long, lEstadoFacturacion As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
    sSql = ""
    If IdPuntosDeCarga > 0 Then
        sSql = sSql + " and a.IdPuntoCarga = " & IdPuntosDeCarga
    End If
    If lEstadoFacturacion <> 0 Then
        sSql = sSql + " and a.IdEstadoFacturacion = " & lEstadoFacturacion & " "
    End If
    sSql = Mid(sSql, 5)
    sSql = sSql + " order by codigo, descripcion"
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "FactCatalogoBienesInsumosSeleccionarPorCuentaAtencionNoPagados"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, sSql): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@lTipoCatalogo", adInteger, adParamInput, 0, lTipoCatalogo): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdCuentaAtencion", adInteger, adParamInput, 0, IdCuentaAtencion): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarPorCuentaAtencionNoPagados = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function SeleccionarBienesLike(sNombre As String, lIdTipoFinanciamiento As Long, lIdPuntoCarga As Long, _
                                                    Optional lbSoloItemsDePaquete As Boolean) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String
    sSql = ""
    If lIdPuntoCarga <> 1 And lIdPuntoCarga <> 99 Then             '<> 0 Then
        sSql = sSql + " and FactPuntosCargaBienesInsumos.IdPuntoCarga = " & lIdPuntoCarga
    End If

    If lIdTipoFinanciamiento <> 0 Then
        sSql = sSql + " and FactCatalogoBienesInsumosHosp.IdTipoFinanciamiento = " & lIdTipoFinanciamiento
    End If

    If sNombre <> "" Then
        sSql = sSql + " and FactCatalogoBienesInsumos.Nombre like '" & sNombre & "%' "
    End If
    If lbSoloItemsDePaquete = True Then
       sSql = sSql + " and FactCatalogoBienesInsumos.EsPaquete=1"
    End If
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "FactCatalogoBienesInsumosSeleccionarBienesLike"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, sSql): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarBienesLike = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


Function SeleccionarPorIdYtipoFinanciamiento(idProducto As Long, idTipoFinanciamiento As Long) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "FactCatalogoBienesInsumosSeleccionarPorIdYtipoFinanciamiento"
        Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 0, idTipoFinanciamiento): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, idProducto): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarPorIdYtipoFinanciamiento = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function


Function SeleccionarPorCodigo(lcCodigo As String, Optional oConexion1 As Connection) As Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError As String
    
    ms_MensajeError = ""
    If oConexion1 Is Nothing Then
        oConexion.CommandTimeout = 300
        oConexion.CursorLocation = adUseClient
        
        oConexion.Open SIGHENTIDADES.CadenaConexion
        
    End If
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexion1 Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexion1
        End If
        .CommandTimeout = 150
        .CommandText = "FactCatalogoBienesInsumosSeleccionarXcodigo"
        Set oParameter = .CreateParameter("@lcCodigo", adVarChar, adParamInput, 7, Left(lcCodigo, 7)): .Parameters.Append oParameter
        Set oRecordset = .Execute
   End With
   Set SeleccionarPorCodigo = oRecordset
     Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Sub ActualizaPrecioSegunIdProducto(lnIdProducto As Long, lnPrecioNuevo As Double)
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FactCatalogoBienesInsumosHospActualizaPrecioSegunIdProducto"
           Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, lnIdProducto): .Parameters.Append oParameter
       Set oParameter = .CreateParameter("@PrecioNuevo", adCurrency, adParamInput, 0, lnPrecioNuevo)
       .Parameters.Append oParameter
       .Execute
   End With
   ms_MensajeError = ""
Exit Sub
ManejadorDeError:
   MsgBox Err.Description
End Sub


