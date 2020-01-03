VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CatalogoServicios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa:  Clase para capa de Mantenimiento de la tabla FactCatalogoServicios
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
'   Descripción:    Inserta un registro a la tabla FactCatalogoServicios
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------
Function Insertar(ByVal oTabla As DOCatalogoServicio) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Insertar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FactCatalogoServiciosAgregar"
           Set oParameter = .CreateParameter("@IdServicioSubGrupo", adInteger, adParamInput, 0, IIf(oTabla.IdServicioSubGrupo = 0, Null, oTabla.IdServicioSubGrupo)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPartida", adInteger, adParamInput, 0, IIf(oTabla.IdPartida = 0, Null, oTabla.IdPartida)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCentroCosto", adInteger, adParamInput, 0, IIf(oTabla.IdCentroCosto = 0, Null, oTabla.IdCentroCosto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicioSubSeccion", adInteger, adParamInput, 0, IIf(oTabla.IdServicioSubSeccion = 0, Null, oTabla.IdServicioSubSeccion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicioSeccion", adInteger, adParamInput, 0, IIf(oTabla.IdServicioSeccion = 0, Null, oTabla.IdServicioSeccion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicioGrupo", adInteger, adParamInput, 0, IIf(oTabla.IdServicioGrupo = 0, Null, oTabla.IdServicioGrupo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Nombre", adVarChar, adParamInput, 250, IIf(oTabla.nombre = "", Null, oTabla.nombre)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 20, IIf(oTabla.codigo = "", Null, oTabla.codigo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamOutput, 0): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@EsCpt", adInteger, adParamInput, 0, oTabla.EsCpt): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NombreMINSA", adVarChar, adParamInput, 250, IIf(oTabla.NombreMINSA = "", Null, oTabla.NombreMINSA)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idEstado", adInteger, adParamInput, 0, oTabla.IdEstado): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@codigoSIS", adVarChar, adParamInput, 20, IIf(oTabla.codigoSIS = "", Null, oTabla.codigoSIS)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@LabResultadoAutomatico", adInteger, adParamInput, 0, oTabla.LabResultadoAutomatico): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idProveedor", adInteger, adParamInput, 0, oTabla.idProveedor): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@equipoCodigo", adVarChar, adParamInput, 2, IIf(oTabla.EquipoCodigo = "", Null, oTabla.EquipoCodigo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@codigoSUNAT", adVarChar, adParamInput, 20, IIf(oTabla.codigoSUNAT = "", Null, oTabla.codigoSUNAT)): .Parameters.Append oParameter

       Set oParameter = .CreateParameter("@IdUsuarioAuditoria", adInteger, adParamInput, 0, oTabla.IdUsuarioAuditoria): .Parameters.Append oParameter
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
'   Descripción:    modifica un registro a la tabla FactCatalogoServicios
'   Parámetros:       Ninguno
'------------------------------------------------------------------------------------

Function Modificar(ByVal oTabla As DOCatalogoServicio) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Modificar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FactCatalogoServiciosModificar"
           Set oParameter = .CreateParameter("@IdServicioSubGrupo", adInteger, adParamInput, 0, IIf(oTabla.IdServicioSubGrupo = 0, Null, oTabla.IdServicioSubGrupo)): oParameter.Precision = 10: oParameter.NumericScale = 0: .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdPartida", adInteger, adParamInput, 0, IIf(oTabla.IdPartida = 0, Null, oTabla.IdPartida)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdCentroCosto", adInteger, adParamInput, 0, IIf(oTabla.IdCentroCosto = 0, Null, oTabla.IdCentroCosto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicioSubSeccion", adInteger, adParamInput, 0, IIf(oTabla.IdServicioSubSeccion = 0, Null, oTabla.IdServicioSubSeccion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicioSeccion", adInteger, adParamInput, 0, IIf(oTabla.IdServicioSeccion = 0, Null, oTabla.IdServicioSeccion)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdServicioGrupo", adInteger, adParamInput, 0, IIf(oTabla.IdServicioGrupo = 0, Null, oTabla.IdServicioGrupo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Nombre", adVarChar, adParamInput, 250, IIf(oTabla.nombre = "", Null, oTabla.nombre)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@Codigo", adVarChar, adParamInput, 20, IIf(oTabla.codigo = "", Null, oTabla.codigo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, IIf(oTabla.idProducto = 0, Null, oTabla.idProducto)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@EsCpt", adInteger, adParamInput, 0, oTabla.EsCpt): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@NombreMINSA", adVarChar, adParamInput, 250, IIf(oTabla.NombreMINSA = "", Null, oTabla.NombreMINSA)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idEstado", adInteger, adParamInput, 0, oTabla.IdEstado): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@codigoSIS", adVarChar, adParamInput, 20, IIf(oTabla.codigoSIS = "", Null, oTabla.codigoSIS)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@LabResultadoAutomatico", adInteger, adParamInput, 0, oTabla.LabResultadoAutomatico): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@idProveedor", adInteger, adParamInput, 0, oTabla.idProveedor): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@equipoCodigo", adVarChar, adParamInput, 2, IIf(oTabla.EquipoCodigo = "", Null, oTabla.EquipoCodigo)): .Parameters.Append oParameter
           Set oParameter = .CreateParameter("@codigoSUNAT", adVarChar, adParamInput, 20, IIf(oTabla.codigoSUNAT = "", Null, oTabla.codigoSUNAT)): .Parameters.Append oParameter
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
'   Descripción:    elimia o borra un registro a la tabla FactCatalogoServicios
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function Eliminar(ByVal oTabla As DOCatalogoServicio) As Boolean
On Error GoTo ManejadorDeError
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   Eliminar = False
   With oCommand
       .CommandType = adCmdStoredProc
       Set .ActiveConnection = mo_Conexion
       .CommandText = "FactCatalogoServiciosEliminar"
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
'   Descripción:    Selecciona un unico registro de la tabla FactCatalogoServicios
'   Parámetros:     Ninguno
'------------------------------------------------------------------------------------

Function SeleccionarPorId(ByVal oTabla As DOCatalogoServicio) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorId = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "FactCatalogoServiciosSeleccionarPorId"
       Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, oTabla.idProducto): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With
 
   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorId = True
       oTabla.IdServicioSubGrupo = IIf(IsNull(oRecordset!IdServicioSubGrupo), "", oRecordset!IdServicioSubGrupo)
       oTabla.IdPartida = IIf(IsNull(oRecordset!IdPartida), 0, oRecordset!IdPartida)
       oTabla.IdCentroCosto = IIf(IsNull(oRecordset!IdCentroCosto), 0, oRecordset!IdCentroCosto)
       oTabla.IdServicioSubSeccion = IIf(IsNull(oRecordset!IdServicioSubSeccion), 0, oRecordset!IdServicioSubSeccion)
       oTabla.IdServicioSeccion = IIf(IsNull(oRecordset!IdServicioSeccion), 0, oRecordset!IdServicioSeccion)
       oTabla.IdServicioGrupo = IIf(IsNull(oRecordset!IdServicioGrupo), 0, oRecordset!IdServicioGrupo)
       oTabla.nombre = IIf(IsNull(oRecordset!nombre), "", oRecordset!nombre)
       oTabla.codigo = IIf(IsNull(oRecordset!codigo), "", oRecordset!codigo)
       oTabla.idProducto = IIf(IsNull(oRecordset!idProducto), 0, oRecordset!idProducto)
       oTabla.EsCpt = IIf(IsNull(oRecordset!EsCpt), 0, oRecordset!EsCpt)
       oTabla.NombreMINSA = IIf(IsNull(oRecordset!NombreMINSA), "", oRecordset!NombreMINSA)
       oTabla.IdEstado = IIf(IsNull(oRecordset!IdEstado), 0, oRecordset!IdEstado)
       oTabla.codigoSIS = IIf(IsNull(oRecordset!codigoSIS), "", oRecordset!codigoSIS)
       oTabla.LabResultadoAutomatico = IIf(IsNull(oRecordset!LabResultadoAutomatico), 0, oRecordset!LabResultadoAutomatico)
       oTabla.idProveedor = IIf(IsNull(oRecordset!idProveedor), 0, oRecordset!idProveedor)
       oTabla.EquipoCodigo = IIf(IsNull(oRecordset!EquipoCodigo), "", oRecordset!EquipoCodigo)
       oTabla.codigoSUNAT = IIf(IsNull(oRecordset!codigoSUNAT), "", oRecordset!codigoSUNAT)
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




Function FiltrarCatalogoBase(ByVal oTabla As DOCatalogoServicio) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSql As String
Dim sSQL1 As String
Dim sSQL2 As String
Dim sWhere As String

    Set FiltrarCatalogoBase = Nothing
    sWhere = ""

    sSQL1 = "select IdServicioSubGrupo, 0 as IdProducto, '' as Codigo, Descripcion from FactCatalogoServiciosSubGrupo"

    sSQL2 = "select IdServicioSubGrupo, IdProducto, Codigo, Nombre, 0 as PrecioUnitario, 1 as Activo from FactCatalogoServicios "

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
            " AS title_chap RELATE IdServicioSubGrupo TO IdServicioSubGrupo) "

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
'Dim sSQL2 As String, sWhere As String
'    sSQL2 = ""
'    sWhere = ""
'    If oTabla.Codigo <> "" Then
'        sWhere = sWhere + " and Codigo = '" & oTabla.Codigo & "'"
'    End If
'    If oTabla.Nombre <> "" Then
'        sWhere = sWhere + " and Nombre like '%" & oTabla.Nombre & "%'"
'    End If
'
'    If sWhere <> "" Then sSQL2 = sSQL2 + " where " + Mid(sWhere, 5)
'
'    sSQL2 = sSQL2 + " order by nombre"
'
'    With oCommand
'        .CommandType = adCmdStoredProc
'        Set .ActiveConnection = mo_Conexion
'        .CommandTimeout = 150
'        .CommandText = "FactCatalogoServiciosFiltrarCatalogoBase"
'        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, sSQL2): .Parameters.Append oParameter
'        Set oRecordSet = .Execute
'        Set oRecordSet.ActiveConnection = Nothing
'   End With
'   Set FiltrarCatalogoBase = oRecordSet
'   Set oCommand = Nothing
'   Exit Function
'ManejadorDeError:
'    MsgBox Err.Description
End Function
Function FiltrarCatalogo(ByVal oTabla As DOCatalogoServicio, lTipoCatalogo As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSql As String
Dim sSQL1 As String
Dim sSQL2 As String
Dim sWhere As String

    Set FiltrarCatalogo = Nothing
    sWhere = ""

    sSQL1 = "select IdServicioSubGrupo, 0 as IdProducto, '' as Codigo, Descripcion from FactCatalogoServiciosSubGrupo"

    sSQL2 = "select FactCatalogoServicios.IdServicioSubGrupo," + _
    " FactCatalogoServicios.IdProducto, " + _
    " FactCatalogoServicios.Codigo," + _
    " FactCatalogoServicios.Nombre, " + _
    " FactCatalogoServiciosHosp.PrecioUnitario, " + _
    " FactCatalogoServiciosHosp.Activo " + _
    " From FactCatalogoServiciosHosp, FactCatalogoServicios " + _
    " Where " + _
    " FactCatalogoServiciosHosp.IdTipoFinanciamiento = " & lTipoCatalogo & _
    " and FactCatalogoServiciosHosp.IdProducto = FactCatalogoServicios.IdProducto"

    If oTabla.codigo <> "" Then
        sSQL2 = sSQL2 + " and FactCatalogoServicios.Codigo = '" & oTabla.codigo & "'"
    End If
    If oTabla.nombre <> "" Then
        sSQL2 = sSQL2 + " and FactCatalogoServicios.Nombre like '%" & oTabla.nombre & "%'"
    End If

    sSQL2 = sSQL2 + " order by codigo"


    sSql = "SHAPE {" + sSQL1 + "} " + _
            "APPEND ({" + sSQL2 + "} " + _
            " AS title_chap RELATE IdServicioSubGrupo TO IdServicioSubGrupo) "


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
'Dim oRecordSet As New Recordset
'Dim oCommand As New ADODB.Command
'Dim oParameter As ADODB.Parameter
'Dim sSQL2 As String, sWhere As String
'    sSQL2 = ""
'    sWhere = ""
'    If oTabla.Codigo <> "" Then
'        sSQL2 = sSQL2 + " and FactCatalogoServicios.Codigo = '" & oTabla.Codigo & "'"
'    End If
'    If oTabla.Nombre <> "" Then
'        sSQL2 = sSQL2 + " and FactCatalogoServicios.Nombre like '%" & oTabla.Nombre & "%'"
'    End If
'
'    sSQL2 = Mid(sSQL2, 5) + " order by codigo"
'
'    With oCommand
'        .CommandType = adCmdStoredProc
'        Set .ActiveConnection = mo_Conexion
'        .CommandTimeout = 150
'        .CommandText = "FactCatalogoServiciosFiltrarCatalogo"
'        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, sSQL2): .Parameters.Append oParameter
'        Set oRecordSet = .Execute
'        Set oRecordSet.ActiveConnection = Nothing
'   End With
'   Set FiltrarCatalogo = oRecordSet
'   Set oCommand = Nothing
'   Exit Function
'ManejadorDeError:
'    MsgBox Err.Description
End Function


Function FiltrarCatalogoDEBB(ByVal oTabla As DOCatalogoServicio, lTipoCatalogo As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSql As String
Dim sSQL1 As String
Dim sSQL2 As String
Dim sWhere As String

    Set FiltrarCatalogoDEBB = Nothing
    sWhere = ""

    sSQL1 = "select IdServicioSubGrupo, 0 as IdProducto, '' as Codigo, Descripcion from FactCatalogoServiciosSubGrupo"

'    sSQL2 = "select FactCatalogoServicios.IdServicioSubGrupo," + _
'    " FactCatalogoServicios.IdProducto, " + _
'    " FactCatalogoServicios.Codigo," + _
'    " FactCatalogoServicios.Nombre, " + _
'    " FactCatalogoServiciosHosp.PrecioUnitario, " + _
'    " FactCatalogoServiciosHosp.Activo,FactCatalogoServicios.codigoSunat, sunat_onu.descripcion as Sunat" + _
'    " From FactCatalogoServiciosHosp, FactCatalogoServicios, sunat_onu " + _
'    " Where " + _
'    " FactCatalogoServiciosHosp.IdTipoFinanciamiento = " & lTipoCatalogo & _
'    " and FactCatalogoServiciosHosp.IdProducto = FactCatalogoServicios.IdProducto" & _
'    " and sunat_onu.codigo = FactCatalogoServicios.codigoSunat "
    sSQL2 = "select FactCatalogoServicios.IdServicioSubGrupo," + _
    " FactCatalogoServicios.IdProducto, " + _
    " FactCatalogoServicios.Codigo," + _
    " FactCatalogoServicios.Nombre, " + _
    " FactCatalogoServiciosHosp.PrecioUnitario, " + _
    " FactCatalogoServiciosHosp.Activo,FactCatalogoServicios.codigoSunat, sunat_onu.descripcion as Sunat" + _
    " FROM FactCatalogoServicios INNER JOIN" + _
    "      FactCatalogoServiciosHosp ON FactCatalogoServicios.IdProducto = FactCatalogoServiciosHosp.IdProducto LEFT OUTER JOIN" + _
    "      sunat_onu ON FactCatalogoServicios.codigoSUNAT = sunat_onu.codigo" + _
    " WHERE FactCatalogoServiciosHosp.IdTipoFinanciamiento = " & lTipoCatalogo

    If oTabla.codigo <> "" Then
        sSQL2 = sSQL2 + " and FactCatalogoServicios.Codigo = '" & oTabla.codigo & "'"
    End If
    If oTabla.nombre <> "" Then
        sSQL2 = sSQL2 + " and FactCatalogoServicios.Nombre like '%" & oTabla.nombre & "%'"
    End If

    sSQL2 = sSQL2 + " order by nombre"


    sSql = "SHAPE {" + sSQL1 + "} " + _
            "APPEND ({" + sSQL2 + "} " + _
            " AS title_chap RELATE IdServicioSubGrupo TO IdServicioSubGrupo) "


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
'Dim oRecordSet As New Recordset
'Dim oCommand As New ADODB.Command
'Dim oParameter As ADODB.Parameter
'Dim sSQL2 As String, sWhere As String
'    sSQL2 = ""
'    sWhere = ""
'    If oTabla.Codigo <> "" Then
'        sSQL2 = sSQL2 + " and FactCatalogoServicios.Codigo = '" & oTabla.Codigo & "'"
'    End If
'    If oTabla.Nombre <> "" Then
'        sSQL2 = sSQL2 + " and FactCatalogoServicios.Nombre like '%" & oTabla.Nombre & "%'"
'    End If
'    sSQL2 = Mid(sSQL2, 5) + " order by nombre"
'    With oCommand
'        .CommandType = adCmdStoredProc
'        Set .ActiveConnection = mo_Conexion
'        .CommandTimeout = 150
'        .CommandText = "FactCatalogoServiciosFiltrarCatalogoDEBB"
'        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, sSQL2): .Parameters.Append oParameter
'        Set oParameter = .CreateParameter("@lTipoCatalogo", adInteger, adParamInput, 0, lTipoCatalogo): .Parameters.Append oParameter
'        Set oRecordSet = .Execute
'        Set oRecordSet.ActiveConnection = Nothing
'   End With
'   Set FiltrarCatalogoDEBB = oRecordSet
'   Set oCommand = Nothing
'   Exit Function
'ManejadorDeError:
'    MsgBox Err.Description
End Function


Function SeleccionarCatalogoBase() As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSql As String
Dim sSQL1 As String
Dim sSQL2 As String

     Set SeleccionarCatalogoBase = Nothing

    sSQL1 = "select IdServicioSubGrupo, 0 as IdProducto, '' as Codigo, Descripcion from FactCatalogoServiciosSubGrupo"

    sSQL2 = "select IdServicioSubGrupo, IdProducto, Codigo, Nombre, 0 as PrecioUnitario, 1 as Activo from FactCatalogoServicios "

    sSql = "SHAPE {" + sSQL1 + "} " + _
            "APPEND ({" + sSQL2 + "} " + _
            " AS title_chap RELATE IdServicioSubGrupo TO IdServicioSubGrupo) "


    oRecordset.Open sSql, mo_Conexion, adOpenDynamic, adLockOptimistic
    Set oRecordset.ActiveConnection = Nothing

    Set SeleccionarCatalogoBase = oRecordset

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
'Dim sSQL2 As String, sWhere As String
'    With oCommand
'        .CommandType = adCmdStoredProc
'        Set .ActiveConnection = mo_Conexion
'        .CommandTimeout = 150
'        .CommandText = "FactCatalogoServiciosSeleccionarCatalogoBase"
'        Set oRecordSet = .Execute
'        Set oRecordSet.ActiveConnection = Nothing
'   End With
'   Set SeleccionarCatalogoBase = oRecordSet
'   Set oCommand = Nothing
'   Exit Function
'ManejadorDeError:
'    MsgBox Err.Description

End Function

Function SeleccionarCatalogo(lTipoCatalogo As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim sSql As String
Dim sSQL1 As String
Dim sSQL2 As String

     Set SeleccionarCatalogo = Nothing

    sSQL1 = "select IdServicioSubGrupo, 0 as IdProducto, '' as Codigo, Descripcion from FactCatalogoServiciosSubGrupo"

    sSQL2 = "select FactCatalogoServicios.IdServicioSubGrupo," + _
    " FactCatalogoServicios.IdProducto, " + _
    " FactCatalogoServicios.Codigo," + _
    " FactCatalogoServicios.Nombre, " + _
    " FactCatalogoServiciosHosp.PrecioUnitario, " + _
    " FactCatalogoServiciosHosp.Activo " + _
    " From FactCatalogoServiciosHosp, FactCatalogoServicios " + _
    " Where " + _
    " FactCatalogoServiciosHosp.IdTipoFinanciamiento = " & lTipoCatalogo & _
    " and FactCatalogoServiciosHosp.IdProducto = FactCatalogoServicios.IdProducto"


    sSql = "SHAPE {" + sSQL1 + "} " + _
            "APPEND ({" + sSQL2 + "} " + _
            " AS title_chap RELATE IdServicioSubGrupo TO IdServicioSubGrupo) "


    oRecordset.Open sSql, mo_Conexion, adOpenDynamic, adLockOptimistic
    Set oRecordset.ActiveConnection = Nothing

    Set SeleccionarCatalogo = oRecordset

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
'Dim sSQL2 As String, sWhere As String
'    With oCommand
'        .CommandType = adCmdStoredProc
'        Set .ActiveConnection = mo_Conexion
'        .CommandTimeout = 150
'        .CommandText = "FactCatalogoServiciosSeleccionarCatalogo"
'        Set oParameter = .CreateParameter("@lTipoCatalogo", adInteger, adParamInput, 0, lTipoCatalogo): .Parameters.Append oParameter
'        Set oRecordSet = .Execute
'        Set oRecordSet.ActiveConnection = Nothing
'   End With
'   Set SeleccionarCatalogo = oRecordSet
'   Set oCommand = Nothing
'   Exit Function
'ManejadorDeError:
'    MsgBox Err.Description
End Function

Function SeleccionarPorTipoCatalogo(oDoCatalogoServicio As DOCatalogoServicio, lTipoCatalogo As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sOrder As String
    sSql = ""
    sOrder = ""
    If oDoCatalogoServicio.codigo <> "" Then
        sSql = sSql + " and Codigo = '" + oDoCatalogoServicio.codigo + "'"
        sOrder = sOrder + " ,Codigo"
    End If
    If oDoCatalogoServicio.nombre <> "" Then
        sSql = sSql + " and Nombre like '%" + oDoCatalogoServicio.nombre + "%'"
        sOrder = sOrder + " ,Nombre"
    End If

    If sOrder <> "" Then sSql = sSql + " order by " + Mid(sOrder, 3)
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "FactCatalogoServiciosSeleccionarPorTipoCatalogo"
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
Dim sSql As String, sOrder As String
    sSql = ""
    sOrder = ""
    
    If IdPuntosDeCarga > 0 Then
        sSql = sSql + " and a.IdPuntoCarga = " & IdPuntosDeCarga
    End If
    If lEstadoFacturacion <> 0 Then
        sSql = sSql + " and a.IdEstadoFacturacion = " & lEstadoFacturacion & " "
    End If

    sSql = sSql + " order by codigo, descripcion"
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "FactCatalogoServiciosSeleccionarPorTipoCatalogo"
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


Function SeleccionarServiciosLike(sNombre As String, lIdTipoFinanciamiento As Long, lIdPuntoCarga As Long) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim sSql As String, sOrder As String
    sSql = ""
    sOrder = ""
    If lIdPuntoCarga <> 1 And lIdPuntoCarga <> 99 Then           '<> 0 Then
        sSql = sSql + " and FactPuntosCargaServicio.IdPuntoCarga = " & lIdPuntoCarga
    End If

    If lIdTipoFinanciamiento <> 0 Then
        sSql = sSql + " and FactCatalogoServiciosHosp.IdTipoFinanciamiento = " & lIdTipoFinanciamiento
    End If

    If sNombre <> "" Then
        sSql = sSql + " and FactCatalogoServicios.Nombre like '%" & sNombre & "%' "
    End If
    sSql = sSql + " Order by FactCatalogoServicios.Nombre"
    
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = mo_Conexion
        .CommandTimeout = 150
        .CommandText = "FactCatalogoServiciosSeleccionarServiciosLike"
        Set oParameter = .CreateParameter("@lcFiltro", adVarChar, adParamInput, 1000, sSql): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarServiciosLike = oRecordset
   Set oCommand = Nothing
   Exit Function
ManejadorDeError:
    MsgBox Err.Description
End Function

Function SeleccionarServiciosPorNombreOCodigoSegunTipofinanciamientoYpuntoCarga(LnTipoFiltro As sghTipoOrden, lcFiltro As String, LnIdTipoFinanciamiento As Long, LnIdPuntoCarga As Long, lnTipoServicioAfiltrar As sghTipoServicioOfrecidos) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    Set SeleccionarServiciosPorNombreOCodigoSegunTipofinanciamientoYpuntoCarga = Nothing
    ms_MensajeError = ""
    oConexion.Open SIGHENTIDADES.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactCatalogoServiciosHospFiltraPorPuntoCargaTipoFinanciamiento"
        Set oParameter = .CreateParameter("@idPuntoCarga", adInteger, adParamInput, 0, LnIdPuntoCarga): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idTipoFinanciamiento", adInteger, adParamInput, 0, LnIdTipoFinanciamiento): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idFiltroTipo", adInteger, adParamInput, 0, LnTipoFiltro): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Filtro ", adVarChar, adParamInput, 100, lcFiltro): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@TipoServicioOfrecido", adInteger, adParamInput, 0, lnTipoServicioAfiltrar): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarServiciosPorNombreOCodigoSegunTipofinanciamientoYpuntoCarga = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function FactCatalogoServiciosSeleccionarPorCodigoOnombre(lcCodigo As String, lcNombre As String, Optional oConexion1 As Connection) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    Set FactCatalogoServiciosSeleccionarPorCodigoOnombre = Nothing
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
        .CommandText = "FactCatalogoServiciosSeleccionarPorCodigoOnombre"
        Set oParameter = .CreateParameter("@codigo", adVarChar, adParamInput, 20, lcCodigo): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@nombre", adVarChar, adParamInput, 50, lcNombre): .Parameters.Append oParameter
        Set oRecordset = .Execute
        If oConexion1 Is Nothing Then
           Set oRecordset.ActiveConnection = Nothing
        End If
   End With
   Set FactCatalogoServiciosSeleccionarPorCodigoOnombre = oRecordset
   If oConexion1 Is Nothing Then
        oConexion.Close
        Set oRecordset = Nothing
   End If
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function SeleccionarServiciosPorNombreOCodigoSegunPuntoCarga(LnTipoFiltro As sghTipoOrden, lcFiltro As String, LnIdPuntoCarga As Long, lnTipoServicioAfiltrar As sghTipoServicioOfrecidos) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    Set SeleccionarServiciosPorNombreOCodigoSegunPuntoCarga = Nothing
    ms_MensajeError = ""
    oConexion.Open SIGHENTIDADES.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactCatalogoServiciosHospFiltraPorPuntoCarga"
        Set oParameter = .CreateParameter("@idPuntoCarga", adInteger, adParamInput, 0, LnIdPuntoCarga): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@idFiltroTipo", adInteger, adParamInput, 0, LnTipoFiltro): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@Filtro ", adVarChar, adParamInput, 100, lcFiltro): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@TipoServicioOfrecido", adInteger, adParamInput, 0, lnTipoServicioAfiltrar): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarServiciosPorNombreOCodigoSegunPuntoCarga = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function SeleccionarPorIdYtipoFinanciamiento(idProducto As Long, idTipoFinanciamiento As Long) As Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    ms_MensajeError = ""
    oConexion.Open SIGHENTIDADES.CadenaConexion
    oConexion.CursorLocation = adUseClient
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandTimeout = 150
        .CommandText = "FactCatalogoServiciosSeleccionarPorIdYtipoFinanciamiento"
        Set oParameter = .CreateParameter("@idProducto", adInteger, adParamInput, 0, idProducto): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@IdTipoFinanciamiento", adInteger, adParamInput, 0, idTipoFinanciamiento): .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
   End With
   Set SeleccionarPorIdYtipoFinanciamiento = oRecordset
   oConexion.Close
   Set oRecordset = Nothing
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
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
        oConexion.Open SIGHENTIDADES.CadenaConexion
        oConexion.CursorLocation = adUseClient
    End If
    With oCommand
        .CommandType = adCmdStoredProc
        If oConexion1 Is Nothing Then
           Set .ActiveConnection = oConexion
        Else
           Set .ActiveConnection = oConexion1
        End If
        .CommandTimeout = 150
        .CommandText = "FactCatalogoServiciosXcodigo"
        Set oParameter = .CreateParameter("@lcCodigo", adVarChar, adParamInput, 20, lcCodigo): .Parameters.Append oParameter
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

'mgaray201410f 1 solo cpt, 0 no cpt y 2 ambos
Function FactCatalogoServiciosSeleccionarPorCodigoOnombreTipo(lcCodigo As String, lcNombre As String, _
                    EsCpt As Long, Optional oConexion1 As Connection) As ADODB.Recordset
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection
Dim ms_MensajeError  As String
    Set FactCatalogoServiciosSeleccionarPorCodigoOnombreTipo = Nothing
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
        .CommandText = "FactCatalogoServiciosSeleccionarPorCodigoOnombreTipoCatalogo"
        Set oParameter = .CreateParameter("@codigo", adVarChar, adParamInput, 20, lcCodigo): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@nombre", adVarChar, adParamInput, 50, lcNombre): .Parameters.Append oParameter
        Set oParameter = .CreateParameter("@EsCpt", adVarChar, adParamInput, 50, EsCpt): .Parameters.Append oParameter
        Set oRecordset = .Execute
        If oConexion1 Is Nothing Then
           Set oRecordset.ActiveConnection = Nothing
        End If
   End With
   Set FactCatalogoServiciosSeleccionarPorCodigoOnombreTipo = oRecordset
   If oConexion1 Is Nothing Then
        oConexion.Close
        Set oRecordset = Nothing
   End If
   Set oConexion = Nothing
   Set oCommand = Nothing
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function SeleccionarPorIdEsParaReceta(ByVal oTabla As DOCatalogoServicio) As Boolean
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
 
   SeleccionarPorIdEsParaReceta = False
   With oCommand
     .CommandType = adCmdStoredProc
     Set .ActiveConnection = mo_Conexion
     .CommandText = "ProductoEsParaOrdenesMedicas"
       Set oParameter = .CreateParameter("@IdProducto", adInteger, adParamInput, 0, oTabla.idProducto): .Parameters.Append oParameter
     Set oRecordset = .Execute
   End With

   If Not (oRecordset.EOF And oRecordset.BOF) Then
       SeleccionarPorIdEsParaReceta = True
   Else
       SeleccionarPorIdEsParaReceta = False
   End If

   oRecordset.Close
   ms_MensajeError = ""
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description
Exit Function
End Function

