VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReglasConfiguarcionReslab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------
'        Organización: Usaid - Politicas en Salud
'        Aplicativo: SisGalenPlus v.3
'        Programa: Clase para Reglas de Resultados de Laboratorio
'        Programado por: Barrantes D
'        Fecha: Julio 2009
'
'------------------------------------------------------------------------------------
Option Explicit
Dim ms_MensajeError As String
Dim mo_ReglasSeguridad As New SIGHNegocios.ReglasDeSeguridad
Dim lcSql As String
Property Let MensajeError(sValue As String)
   ms_MensajeError = sValue
End Property
Property Get MensajeError() As String
   MensajeError = ms_MensajeError
End Property

Function LabGruposSeleccionarTodos() As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New LabGrupos
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set LabGruposSeleccionarTodos = oTabla.SeleccionarTodos
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

Function LabItemsGruposSeleccionarTodos(criterio As String) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New LabItemsGrupos
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set LabItemsGruposSeleccionarTodos = oTabla.SeleccionarTodos(criterio)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

Function LabItemsSeleccionarTodos(criterio As String) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New LabItems
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    Set LabItemsSeleccionarTodos = oTabla.SeleccionarTodos(criterio)
    oConexion.Close
    
    ms_MensajeError = oTabla.MensajeError
    Set oConexion = Nothing
    Set oTabla = Nothing
End Function

Function CatalogoServiciosFiltrar(oDOCatalogoServicio As DOCatalogoServicio, ml_TipoCatalogo As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oTabla As New CatalogoServicios
    
    oConexion.Open sighentidades.CadenaConexionShape
    oConexion.CursorLocation = adUseClient
    
    Set oTabla.Conexion = oConexion
    
    Select Case ml_TipoCatalogo
    Case 0
        Set CatalogoServiciosFiltrar = oTabla.FiltrarCatalogoBase(oDOCatalogoServicio)
    Case Else
        Set CatalogoServiciosFiltrar = oTabla.FiltrarCatalogodebb(oDOCatalogoServicio, ml_TipoCatalogo)
    End Select
    
    oConexion.Close
    ms_MensajeError = oTabla.MensajeError

    Set oConexion = Nothing
    Set oTabla = Nothing

End Function

Function LabItemsCPTSeleccionarPorIdRecordset(idProductoCPT As Long) As ADODB.Recordset
Dim oConexion As New ADODB.Connection
Dim oLabItemCPT As New LabItemCPT
    
    Set LabItemsCPTSeleccionarPorIdRecordset = Nothing
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    Set oLabItemCPT.Conexion = oConexion
    
    Set LabItemsCPTSeleccionarPorIdRecordset = oLabItemCPT.SeleccionarPorId(idProductoCPT)
'    oConexion.Close
    ms_MensajeError = oLabItemCPT.MensajeError
    Set oConexion = Nothing
'    Set oLabItemCPT = Nothing

End Function

Function FiltrarCatalogoSC(ByVal oTabla As DOCatalogoServicio) As ADODB.Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oConexion As New Connection
Dim sSql As String
Dim sSQL1 As String
Dim sSQL2 As String

    oConexion.Open sighentidades.CadenaConexionShape
    oConexion.CursorLocation = adUseClient
    Set FiltrarCatalogoSC = Nothing

    sSQL1 = "select IdServicioSubGrupo, 0 as IdProducto, '' as Codigo, Descripcion from FactCatalogoServiciosSubGrupo"

    sSQL2 = " SELECT " & _
            " FactCatalogoServicios.IdServicioSubGrupo , " & _
            " FactCatalogoServicios.idProducto, " & _
            " FactCatalogoServicios.Codigo, " & _
            " FactCatalogoServicios.nombre, " & _
            " FactCatalogoServicios.nombreMinsa, " & _
            " FactCatalogoServiciosGrupo.Descripcion " & _
            " FROM FactCatalogoServiciosGrupo INNER JOIN FactCatalogoServicios " & _
            " ON FactCatalogoServiciosGrupo.IdServicioGrupo = FactCatalogoServicios.IdServicioGrupo " & _
            " WHERE NOT (FactCatalogoServicios.IdProducto " & _
            " IN (SELECT LabItemsCPT.idProductoCpt FROM LabItemsCPT))"

    If oTabla.Codigo <> "" Then
        sSQL2 = sSQL2 + " and FactCatalogoServicios.Codigo  = '" & oTabla.Codigo & "'"
    End If
    If oTabla.nombre <> "" Then
        sSQL2 = sSQL2 + " and FactCatalogoServicios.Nombre like '%" & oTabla.nombre & "%'"
    End If

    sSQL2 = sSQL2 + " order by nombre"


    sSql = "SHAPE {" + sSQL1 + "} " + _
            "APPEND ({" + sSQL2 + "} " + _
            " AS title_chap RELATE IdServicioSubGrupo TO IdServicioSubGrupo) "


    oRecordset.Open sSql, oConexion, adOpenDynamic, adLockOptimistic
    Set oRecordset.ActiveConnection = Nothing

    Set FiltrarCatalogoSC = oRecordset

    ms_MensajeError = ""

Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function

Function FiltrarCatalogoCC(ByVal oTabla As DOCatalogoServicio) As ADODB.Recordset

On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oConexion As New Connection
Dim sSql As String
Dim sSQL1 As String
Dim sSQL2 As String
Dim sWhere As String

    oConexion.Open sighentidades.CadenaConexionShape
    oConexion.CursorLocation = adUseClient
    Set FiltrarCatalogoCC = Nothing
    sWhere = ""

    sSQL1 = "select IdServicioSubGrupo, 0 as IdProducto, '' as Codigo, Descripcion from FactCatalogoServiciosSubGrupo "
    sSQL2 = "SELECT  FactCatalogoServicios.IdservicioSubGrupo, " & _
            "FactCatalogoServicios.IdProducto, " & _
            "FactCatalogoServicios.Codigo, " & _
            "FactCatalogoServicios.Nombre, " & _
            "FactCatalogoServicios.NombreMINSA, " & _
            "FactCatalogoServiciosSubGrupo.Descripcion " & _
            "From FactCatalogoServicios " & _
            "INNER JOIN FactCatalogoServiciosSubGrupo ON FactCatalogoServicios.IdServicioSubGrupo = FactCatalogoServiciosSubGrupo.IdServicioSubGrupo " & _
            "WHERE (FactCatalogoServicios.IdProducto IN (SELECT labItemsCPT.idProductoCpt FROM LabItemsCpt)) "

    If oTabla.Codigo <> "" Then
        sSQL2 = sSQL2 + " and FactCatalogoServicios.codigo = '" & oTabla.Codigo & "'"
    End If
    If oTabla.nombre <> "" Then
        sSQL2 = sSQL2 + " and FactCatalogoServicios.Nombre like '%" & oTabla.nombre & "%'"
    End If

    sSQL2 = sSQL2 + " order by nombre"


    sSql = "SHAPE {" + sSQL1 + "} " + _
            "APPEND ({" + sSQL2 + "} " + _
            " AS title_chap RELATE IdServicioSubGrupo TO IdServicioSubGrupo) "


    oRecordset.Open sSql, oConexion, adOpenDynamic, adLockOptimistic
    Set oRecordset.ActiveConnection = Nothing

    Set FiltrarCatalogoCC = oRecordset

    ms_MensajeError = ""

Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
    Exit Function
End Function

Function CatalogoServicioSeleccionarPorId(idProducto As Long) As DOCatalogoServicio
Dim oConexion As New ADODB.Connection
Dim oCatalogoServicio As New SIGHDatos.CatalogoServicios
Dim oDOCatalogoServicio As New DOCatalogoServicio
    
    Set CatalogoServicioSeleccionarPorId = Nothing
    
    oConexion.Open sighentidades.CadenaConexion
    oConexion.CursorLocation = adUseClient
    
    Set oCatalogoServicio.Conexion = oConexion
    oDOCatalogoServicio.idProducto = idProducto
    
    If oCatalogoServicio.SeleccionarPorId(oDOCatalogoServicio) Then
        Set CatalogoServicioSeleccionarPorId = oDOCatalogoServicio
    End If
    oConexion.Close
    ms_MensajeError = oCatalogoServicio.MensajeError
    Set oConexion = Nothing
    Set oCatalogoServicio = Nothing
End Function

Function LabItemCPTAgregar(Registro As Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcNombreServicio As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oDOLabItemsCPT As New DoLabItemsCPT
    Dim oLabItemCPT As New LabItemCPT
    Dim ob_Flag As Boolean
    Dim x As Integer

    ob_Flag = False
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    
    Set oLabItemCPT.Conexion = oConexion
    Registro.MoveFirst
    For x = 1 To Registro.RecordCount
        oConexion.BeginTrans
        oDOLabItemsCPT.IdGrupo = Registro.Fields!IdGrupo
        oDOLabItemsCPT.idItem = Registro.Fields!idItem
        oDOLabItemsCPT.idItemGrupo = Registro.Fields!idItemGrupo
        oDOLabItemsCPT.idProductoCPT = Registro.Fields!idProductoCPT
        oDOLabItemsCPT.IdUsuarioAuditoria = Registro.Fields!IdUsuarioAuditoria
        oDOLabItemsCPT.Metodo = Registro.Fields!Metodo
        oDOLabItemsCPT.ordenXresultado = Registro.Fields!ordenXresultado
        oDOLabItemsCPT.SoloCheck = Registro.Fields!SoloCheck
        oDOLabItemsCPT.SoloCombo = Registro.Fields!SoloCombo
        oDOLabItemsCPT.SoloNumero = Registro.Fields!SoloNumero
        oDOLabItemsCPT.SoloTexto = Registro.Fields!SoloTexto
        oDOLabItemsCPT.ValorReferencial = Registro.Fields!ValorReferencial
        oDOLabItemsCPT.ValorSiEsCombo = Registro.Fields!ValorSiEsCombo
        If oLabItemCPT.Insertar(oDOLabItemsCPT) Then
            Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOLabItemsCPT.IdUsuarioAuditoria, "A", oDOLabItemsCPT.idProductoCPT, "LabItemsCPT", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcNombreServicio)           'ListBarItems.idListItem
            ob_Flag = True
            oConexion.CommitTrans
        Else
            ob_Flag = False
            ms_MensajeError = oLabItemCPT.MensajeError
            oConexion.RollbackTrans
        End If
        Registro.MoveNext
    Next x
    
    LabItemCPTAgregar = ob_Flag
    oConexion.Close
    Set oConexion = Nothing
    Set oLabItemCPT = Nothing
End Function


Function LabItemCPTModificar(Registro As Recordset, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcNombreServicio As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oDOLabItemsCPT As New DoLabItemsCPT
    Dim oLabItemCPT As New LabItemCPT
    Dim ob_Flag As Boolean
    Dim x As Integer

    ob_Flag = False
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    
    Set oLabItemCPT.Conexion = oConexion
    Registro.MoveFirst
    For x = 1 To Registro.RecordCount
        oConexion.BeginTrans
        oDOLabItemsCPT.IdGrupo = Registro.Fields!IdGrupo
        oDOLabItemsCPT.idItem = Registro.Fields!idItem
        oDOLabItemsCPT.idItemGrupo = Registro.Fields!idItemGrupo
        oDOLabItemsCPT.idProductoCPT = Registro.Fields!idProductoCPT
        oDOLabItemsCPT.IdUsuarioAuditoria = Registro.Fields!IdUsuarioAuditoria
        oDOLabItemsCPT.Metodo = Registro.Fields!Metodo
        oDOLabItemsCPT.ordenXresultado = Registro.Fields!ordenXresultado
        oDOLabItemsCPT.SoloCheck = Registro.Fields!SoloCheck
        oDOLabItemsCPT.SoloCombo = Registro.Fields!SoloCombo
        oDOLabItemsCPT.SoloNumero = Registro.Fields!SoloNumero
        oDOLabItemsCPT.SoloTexto = Registro.Fields!SoloTexto
        oDOLabItemsCPT.ValorReferencial = Registro.Fields!ValorReferencial
        oDOLabItemsCPT.ValorSiEsCombo = Registro.Fields!ValorSiEsCombo
        If Registro.Fields!Estado Then
            If oLabItemCPT.modificar(oDOLabItemsCPT) Then
                Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOLabItemsCPT.IdUsuarioAuditoria, "M", oDOLabItemsCPT.idProductoCPT, "LabItemsCPT", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcNombreServicio)           'ListBarItems.idListItem
                ob_Flag = True
                oConexion.CommitTrans
            Else
                ob_Flag = False
                ms_MensajeError = oLabItemCPT.MensajeError
                oConexion.RollbackTrans
            End If
        Else
            If oLabItemCPT.Insertar(oDOLabItemsCPT) Then
                Call mo_ReglasSeguridad.AuditoriaAgregarV(oDOLabItemsCPT.IdUsuarioAuditoria, "A", oDOLabItemsCPT.idProductoCPT, "LabItemsCPT", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcNombreServicio)           'ListBarItems.idListItem
                ob_Flag = True
                oConexion.CommitTrans
            Else
                ob_Flag = False
                ms_MensajeError = oLabItemCPT.MensajeError
                oConexion.RollbackTrans
            End If
        End If
        Registro.MoveNext
    Next x
    
    LabItemCPTModificar = ob_Flag
    oConexion.Close
    Set oConexion = Nothing
    Set oLabItemCPT = Nothing
End Function


Function LabMayor(wTabla As String, wCampo As String) As Integer
On Error GoTo ManejadorDeError
Dim oRecordset As New ADODB.Recordset
Dim oCommand As New ADODB.Command
Dim oParameter As ADODB.Parameter
Dim oConexion As New ADODB.Connection

    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
      .CommandType = adCmdStoredProc
      Set .ActiveConnection = oConexion
      .CommandText = "LabMayor"
      Set oParameter = .CreateParameter("@Tabla1", adVarChar, adParamInput, 200, wTabla)
      .Parameters.Append oParameter
      Set oParameter = .CreateParameter("@Campo", adVarChar, adParamInput, 200, wCampo)
      .Parameters.Append oParameter
      Set oRecordset = .Execute
      Set oRecordset.ActiveConnection = Nothing
    End With
   
    LabMayor = oRecordset.Fields!mayor + 1
 
   ms_MensajeError = ""
   Set oCommand = Nothing
   Set oRecordset = Nothing
 
Exit Function
ManejadorDeError:
   ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
Exit Function
End Function


Function LabItemGrupoAgregar(Registro As DoLabItemsGrupos, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcNombreServicio As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oDoLabItemsGrupos As New DoLabItemsGrupos
    Dim oLabItemsGrupos As New LabItemsGrupos
    
    LabItemGrupoAgregar = False
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oLabItemsGrupos.Conexion = oConexion
    oConexion.BeginTrans
    oDoLabItemsGrupos.idItemGrupo = Registro.idItemGrupo
    oDoLabItemsGrupos.Grupo = Registro.Grupo
    
    If oLabItemsGrupos.Insertar(oDoLabItemsGrupos) Then
        Call mo_ReglasSeguridad.AuditoriaAgregarV(Registro.IdUsuarioAuditoria, "A", oDoLabItemsGrupos.idItemGrupo, "LabItemsGrupos", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcNombreServicio)
        oConexion.CommitTrans
        LabItemGrupoAgregar = True
    Else
        ms_MensajeError = oLabItemsGrupos.MensajeError
        oConexion.RollbackTrans
    End If
    
    oConexion.Close
    Set oConexion = Nothing
    Set oLabItemsGrupos = Nothing
End Function

Function LabItemGrupoEliminar(Registro As DoLabItemsGrupos, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcNombreServicio As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oLabItemsGrupos As New LabItemsGrupos
    
    LabItemGrupoEliminar = False
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oLabItemsGrupos.Conexion = oConexion
    oConexion.BeginTrans

    If oLabItemsGrupos.Eliminar(Registro) Then
        Call mo_ReglasSeguridad.AuditoriaAgregarV(Registro.IdUsuarioAuditoria, "E", Registro.idItemGrupo, "LabItemsGrupos", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcNombreServicio)
        oConexion.CommitTrans
        LabItemGrupoEliminar = True
    Else
        ms_MensajeError = oLabItemsGrupos.MensajeError
        oConexion.RollbackTrans
    End If
    
    oConexion.Close
    Set oConexion = Nothing
    Set oLabItemsGrupos = Nothing
End Function

Function LabItemEliminar(Registro As DoLabItems, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcNombreServicio As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oLabItems As New LabItems
    
    LabItemEliminar = False
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oLabItems.Conexion = oConexion
    oConexion.BeginTrans

    If oLabItems.Eliminar(Registro) Then
        Call mo_ReglasSeguridad.AuditoriaAgregarV(Registro.IdUsuarioAuditoria, "E", Registro.idItem, "LabItems", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcNombreServicio)
        oConexion.CommitTrans
        LabItemEliminar = True
    Else
        ms_MensajeError = oLabItems.MensajeError
        oConexion.RollbackTrans
    End If
    
    oConexion.Close
    Set oConexion = Nothing
    Set oLabItems = Nothing
End Function
Function LabItemAgregar(Registro As DoLabItems, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcNombreServicio As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oDoLabItems As New DoLabItems
    Dim oLabItems As New LabItems
    
    LabItemAgregar = False
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oLabItems.Conexion = oConexion
    oConexion.BeginTrans
    oDoLabItems.idItem = Registro.idItem
    oDoLabItems.Item = Registro.Item
    oDoLabItems.idProductoCPT = 0
    
    If oLabItems.Insertar(oDoLabItems) Then
        Call mo_ReglasSeguridad.AuditoriaAgregarV(Registro.IdUsuarioAuditoria, "A", oDoLabItems.idItem, "LabItems", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcNombreServicio)
        oConexion.CommitTrans
        LabItemAgregar = True
    Else
        ms_MensajeError = oLabItems.MensajeError
        oConexion.RollbackTrans
    End If
    
    oConexion.Close
    Set oConexion = Nothing
    Set oLabItems = Nothing
End Function

Function ValidarReglasGrupoItem(cod As String) As Boolean 'modificado a sp 05/08 samuel
On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Dim res As Boolean

    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandText = "ValidarReglasGrupoItem"
        Set oParameter = .CreateParameter("@cod", adVarChar, adParamInput, 200, cod)
        .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
    End With
   
    If oRecordset.RecordCount > 0 Then res = False Else res = True
    ValidarReglasGrupoItem = res
    ms_MensajeError = ""
    Set oCommand = Nothing
    Set oRecordset = Nothing
    Set oConexion = Nothing
    Exit Function
ManejadorDeError:
    ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
    Exit Function
End Function

Function ValidarReglasItem(cod As String) As Boolean 'modificado a sp 05/08 samuel
On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Dim res As Boolean

    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandText = "ValidarReglasItem"
        Set oParameter = .CreateParameter("@cod", adVarChar, adParamInput, 200, cod)
        .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
    End With
   
    If oRecordset.RecordCount > 0 Then res = False Else res = True
    ValidarReglasItem = res
    ms_MensajeError = ""
    Set oCommand = Nothing
    Set oRecordset = Nothing
    Set oConexion = Nothing
    Exit Function
ManejadorDeError:
    ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
    Exit Function
End Function

Function VerificaExamen(oCod As Long) As Boolean 'modificado a sp 05/08 samuel
On Error GoTo ManejadorDeError
    Dim oRecordset As New ADODB.Recordset
    Dim oCommand As New ADODB.Command
    Dim oParameter As ADODB.Parameter
    Dim oConexion As New ADODB.Connection
    Dim res As Boolean

    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    With oCommand
        .CommandType = adCmdStoredProc
        Set .ActiveConnection = oConexion
        .CommandText = "ValidarReglasExamen"
        Set oParameter = .CreateParameter("@cod", adInteger, adParamInput, 0, oCod)
        .Parameters.Append oParameter
        Set oRecordset = .Execute
        Set oRecordset.ActiveConnection = Nothing
    End With
   
    If oRecordset.RecordCount > 0 Then res = False Else res = True
    VerificaExamen = res
    ms_MensajeError = ""
    Set oCommand = Nothing
    Set oRecordset = Nothing
    Set oConexion = Nothing
    Exit Function
ManejadorDeError:
    ms_MensajeError = Err.Number & " " + Err.Description: MsgBox ms_MensajeError + Chr(13) + "Por favor contacte al personal de soporte técnico", vbInformation, "Error en la interface de acceso a datos"
    Exit Function
End Function

Function LabItemsCPTEliminar(Registro As DoLabItemsCPT, mo_lnIdTablaLISTBARITEMS As Long, mo_lcNombrePc As String, lcNombreServicio As String) As Boolean
    Dim oConexion As New ADODB.Connection
    Dim oLabItemsCPT As New LabItemCPT
    
    LabItemsCPTEliminar = False
    oConexion.CursorLocation = adUseClient
    oConexion.CommandTimeout = 300
    oConexion.Open sighentidades.CadenaConexion
    
    Set oLabItemsCPT.Conexion = oConexion
    oConexion.BeginTrans

    If oLabItemsCPT.Eliminar(Registro) Then
        Call mo_ReglasSeguridad.AuditoriaAgregarV(Registro.IdUsuarioAuditoria, "E", Registro.idProductoCPT, "LabItemsCPT", oConexion, mo_lnIdTablaLISTBARITEMS, mo_lcNombrePc, lcNombreServicio)
        oConexion.CommitTrans
        LabItemsCPTEliminar = True
    Else
        ms_MensajeError = oLabItemsCPT.MensajeError
        oConexion.RollbackTrans
    End If
    
    oConexion.Close
    Set oConexion = Nothing
    Set oLabItemsCPT = Nothing
End Function
